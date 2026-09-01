* FoxRemote - Multi-database abstraction for VFP
* Supported engines: MSSQL, MySQL, MariaDB, Firebird, SQLite, PostgreSQL
* Usage: SET PROCEDURE TO foxremote.prg ADDITIVE
*        loDB = CreateObject("MySQL") ... loDB.Connect()
* Use Sample.prg for standalone usage examples.
RETURN
* ======================================================================== *
* Class DBEngine
* ======================================================================== *
Define Class DBEngine As Custom
	#DEFINE CRLF CHR(13)+CHR(10)
	cDriver		= ""
	cServer		= ""
	cUser		= ""
	cPassword	= ""
	cDatabase 	= ""
	nPort		= 0
	cVersion	= "0.0.1"
	bUseCA		= .T.
	bUseSymbolDelimiter = .F.
	cPKName		= "TID"	
	nMaxLength  = 0 && Every engine should fill this value.
	bCanGenerateGUID = .T.
	
	Dimension aCustomArray[1]
	Hidden nCounter
	nCounter = 0
	bExecuteIndexScriptSeparately = .f.
	bExecuteFkScriptSeparately = .f.
	cLeft = ''
	cRight = ''
	bShowErrors = .T.
	Dimension aLastError[1]
	cLastError = ''

	Hidden oRegEx, oViews, oGroupViews, nHandle
	
	Procedure Init
		With This
			.oViews = Createobject("Collection")
			.oGroupViews = Createobject("Collection")
			.oRegEx = Createobject("VBScript.RegExp")
			.oRegEx.IgnoreCase = .T.
			.oRegEx.Global = .T.
			.nHandle = 0
		Endwith
	Endproc


*!*		Function connectFromKvp(toKvp, tcConfigFile)
*!*			If type('toKvp') == 'O' and Lower(toKvp.name) == 'kvp'
*!*				toKvp.parse(tcConfigFile)
*!*				This.cDriver 	= Strconv(toKvp.Get('driver'),14)
*!*				This.cServer 	= Strconv(toKvp.Get('server'),14)
*!*				This.cUser		= Strconv(toKvp.Get('user'),14)
*!*				This.cPassword	= Strconv(toKvp.Get('password'),14)
*!*				This.nPort		= Strconv(toKvp.Get('port'),14)
*!*				This.cDatabase	= Strconv(toKvp.Get('database'),14)
*!*			EndIf	
*!*		EndFunc

	function Connect(tbAddDatabase)
		If This.nHandle > 0
			If This.reconnect()
				Return
			Endif
		Endif

		Local lcConStr
		lcConStr = This.getConnectionString(tbAddDatabase)
		** POLICIA
		*MESSAGEBOX(lcConStr)
		** POLICIA
		This.nHandle = Sqlstringconnect(lcConStr, .T.)

		If This.nHandle <= 0
			This.sqlError()
			Return .f.
		Endif
		This.applyConnectionSettings()
		
		If tbAddDatabase
			Return && No es necesario crear la base de datos.
		EndIf
		
		this.newDataBase(this.cDatabase)
		this.selectDatabase()
		Return .t.
	EndFunc
	
	Function newDataBase(tcDataBase)
		If InList(Lower(this.Name), "firebird", "sqlite")
			Return .t.
		EndIf
		Local lcScript, lcCursor, lcDBName
		lcCursor = Sys(2015)
		lcScript = this.getDataBaseExistsScript(tcDataBase)
		If !This.SQLExec(lcScript, lcCursor)
			Return .F.
		EndIf
		
		Select (lcCursor)
		lcDBName = &lcCursor..dbName
		Use in (lcCursor)
		
		If !Empty(lcDBName)
			this.changeDB(tcDataBase)
			Return .t.
		EndIf
		
		lcScript = this.getCreateDataBaseScript(tcDataBase)
		If !This.SQLExec(lcScript)
			Return .F.
		EndIf
		this.changeDB(tcDataBase)
		
		Return .t.		
	EndFunc

	function migrate(tcTableOrPath, toScripts)
		Local lbCloseTable, laTables[1], i, j, k, lcTableName, lcTablePath, lcPathAct, ;
			lcFieldsScript, lcValuesScript, lcLeft, lcRight, lcDateAct, laDateFields[1], ;
			lcMarkAct, lcCenturyAct, loEnv, lcScript, lbMigrateDBC, loTables, lcTableDescription, ;
			loComposedIndexes, lbResult

		lcPathAct = Set("Default")
		lcTableDescription = ""
		If Directory(tcTableOrPath)
			Set Default To (Addbs(tcTableOrPath))
			=Adir(laDBFList, "*.dbf")
			j = 0
			For i = 1 to Alen(laDBFList, 1)
				j = j + 1
				Dimension laTables[j]
				laTables[j] = laDBFList[i, 1]
			EndFor
			Store 0 to i, j
			Release laDBFList
		Else
			If !InList(Upper(JustExt(tcTableOrPath)), "DBC", "DBF")
				Text to this.cLastError noshow pretext 7 textmerge
				    Error - Tipo de Archivo Inválido:
				    Solo se permiten migraciones de ficheros DBF o DBC.

				    Por favor, asegúrate de que estás intentando migrar un archivo con una extensión válida (DBF, DBC o TMG).
				EndText
				If this.bShowErrors
					MessageBox(this.cLastError, 16)
				EndIf
				Return .f.
			EndIf

			Do case
			case Upper(JustExt(tcTableOrPath)) == "DBC"
				Open Database (tcTableOrPath) Shared
				=ADBObjects(laTables, "TABLE")
				lbMigrateDBC = .T.
			Case Upper(JustExt(tcTableOrPath)) == "DBF"
				laTables[1]  = tcTableOrPath
			EndCase
		Endif

		lcLeft = This.cLeft
		lcRight = This.cRight
		
		loEnv = this.setEnvironment()
		For i = 1 To Alen(laTables,1)
			lcTablePath = laTables[i]
			Try
				If Type('loTables') != 'O'
					lcTableName = Juststem(lcTablePath)
					If !Used(lcTableName)
						lbCloseTable = .T.
						Use (lcTablePath) In 0
					EndIf
					=Afields(laFields, lcTableName)
				Else
					Local laFields[1]
					lcTableName 		= loTables(i).cTableName
					lcTableDescription 	= loTables(i).cTableDescription
					loComposedIndexes 	= loTables(i).oComposedIndexes
					Acopy(loTables(i).aTableFields, laFields)
				EndIf				

				laDateFields = this.getDateTimeFields(@laFields)

				If This.tableExists(lcTableName)
					If !this.sqlExec(this.dropTable(lcTableName))
						Return
					EndIf
				Endif
				lbResult = This.createTable(lcTableName, lcTableDescription, loComposedIndexes, @laFields, toScripts)
				
				If Type('loTables') != 'O'
					* Iterate fields
					lcFieldsScript = Space(1)
					lcValuesScript = Space(1)

					For j=1 To Alen(laFields, 1)
						lcFieldsScript = lcFieldsScript + lcLeft + laFields[j, 1] + lcRight + ','
						lcValuesScript = lcValuesScript + '?loRow.' + laFields[j, 1] + ','
					EndFor
					
					lcFieldsScript = Substr(lcFieldsScript, 1, Len(lcFieldsScript)-1)
					lcValuesScript = Substr(lcValuesScript, 1, Len(lcValuesScript)-1)

					* Insert values
					Select (lcTableName)
					Scan
						Scatter Memo Name loRow
						this.updateFetchedRow(@laDateFields, loRow)
						lcScript = "INSERT INTO " + lcLeft + lcTableName + lcRight + " (" + lcFieldsScript + ")"
						lcScript = lcScript + " VALUES (" + lcValuesScript + ");"
						This.SQLExec(lcScript)
					EndScan
				EndIf
			Catch To loEx
				This.printException(loEx)
			Endtry

			If lbCloseTable
				Use In (lcTableName)
			Endif
		Endfor
		this.restoreEnvironment(loEnv)
		
		If lbMigrateDBC
			Close Databases ALL
		EndIf
		Return lbResult
	endfunc

	function createTable(tcTableName, tcTableDescription, toComposedIndexes, taFields, toScripts)
		Local i, lcScript, lcType, lcName, lcSize, lcDecimal, lbAllowNull, lcLongName, ;
			lcComment, lnNextValue, lnStepValue, lcDefault, lcLeft, lcRight, loFields, ;
			lcFieldsScript, lcInternalID, lbInsertInternalID, loIdxScript, loFkScript, loOptions
		
		lcLeft  = This.cLeft
		lcRight = This.cRight

		lcFieldsScript 		= ''
		lcInternalID  		= lcLeft + this.cPKName + lcRight + Space(1) + This.getTidScript()
		lbInsertInternalID 	= .T.
		loIdxScript			= CreateObject("Collection")
		loFkScript 			= CreateObject("Collection")
		
		lcDefault = Space(1)
		loFields  = Createobject("Empty")
		=AddProperty(loFields, "name", "")
		=AddProperty(loFields, "type", "")
		=AddProperty(loFields, "size", "")
		=AddProperty(loFields, "decimal", "")
		=AddProperty(loFields, "allowNull", .F.)
		=AddProperty(loFields, "longName", "")
		=AddProperty(loFields, "comment", "")
		=AddProperty(loFields, "nextValue", 0)
		=AddProperty(loFields, "stepValue", 0)
		=AddProperty(loFields, "default", "")
		=AddProperty(loFields, "autoIncrement", .F.)
		=AddProperty(loFields, "primaryKey", .F.)
		=AddProperty(loFields, "addDefault", .T.)
		=AddProperty(loFields, "foreignKey", .null.)
		=AddProperty(loFields, "index", .null.)
		=AddProperty(loFields, "tag", "")

		For i = 1 To Alen(taFields, 1)
			If Upper(taFields[i, 2]) == 'U' && UUID
				If i > 1
					lcFieldsScript = lcFieldsScript + ', '
				EndIf
				lcFieldsScript = lcFieldsScript + ' ' + This.getUUIDScript()
				Loop
			EndIf
			loFields.Name 		= taFields[i, 1]
			loFields.Type 		= taFields[i, 2]
			loFields.Size 		= Alltrim(Str(taFields[i, 3]))
			loFields.Decimal 	= Alltrim(Str(taFields[i, 4]))
			loFields.allowNull 	= taFields[i, 5]
			loFields.longName 	= taFields[i, 12]
			loFields.Comment 	= taFields[i, 16]
			loFields.Nextvalue 	= taFields[i, 17]
			loFields.stepValue 	= taFields[i, 18]
			loFields.addDefault = .t.
			
			loFields.Default 		= "''"
			loFields.primaryKey 	= .f.
			loFields.foreignKey 	= .null.
			loFields.autoIncrement 	= .f.
			
			* Validate types with mandatory length
			If InList(Upper(loFields.Type), 'C') and Empty(Val(loFields.Size))
				Text to this.cLastError noshow pretext 7 textmerge
				    Error - Tipo de Dato CHAR sin Longitud:

				    El tipo de dato CHAR requiere que se especifique su longitud. Por favor, asegúrate de agregar la longitud después del tipo de dato CHAR en la definición de la columna.

				    Ejemplo Correcto:
				    NOMBRE CHAR(50)

				    Por favor, corrige la definición de la tabla para incluir la longitud del tipo CHAR y vuelve a intentarlo.
				EndText
				If this.bShowErrors
					MessageBox(this.cLastError, 48)
				EndIf
				loop
			EndIf
&& --> IRODG 16/01/2024
*If Type('taFields[i, 19]') != 'U'
&& <-- IRODG 16/01/2024			
			If Alen(taFields,2) > 18
				If taFields[i, 19] != "''"
					loFields.Default = "'" + taFields[i, 19] + "'"
				EndIf
				loFields.primaryKey 	= taFields[i, 20]
				loFields.foreignKey 	= taFields[i, 21]
				loFields.autoIncrement 	= taFields[i, 22]
				* Si el campo es autoincrement entonces primaryKey es .t.
				If loFields.autoIncrement
					loFields.primaryKey = .t.
					loFields.allowNull = .f.
					loFields.addDefault = .f.
					loFields.default = "''"
				EndIf
				loFields.index = taFields[i, 23]
			EndIf
			If i > 1
				lcFieldsScript = lcFieldsScript + ', '
			EndIf

			lcFieldsScript = lcFieldsScript + lcLeft + loFields.Name + lcRight + Space(1)
			lcMacro = "this.visit" + loFields.Type + "Type(loFields)"
			lcValue = &lcMacro
			
			If loFields.autoIncrement
				lcValue = this.changeTypeOnAutoIncrement(lcValue)
				lbInsertInternalID = .F.
			EndIf
			lcFieldsScript = lcFieldsScript + lcValue
			lcFieldsScript = lcFieldsScript + this.addFieldOptions(loFields)
			
			If !IsNull(loFields.foreignKey)
				loFkScript.Add(this.addForeignKey(loFields.foreignKey))
			EndIf
			
			If !IsNull(loFields.index)
				loIdxScript.Add(this.addSingleIndex(loFields.index))
			EndIf
		EndFor
		Local lcHeader
		Text to lcHeader noshow pretext 7 textmerge
-- =======================================
-- Tabla: <<tcTableName>>
-- Descripción: <<tcTableDescription>>
-- =======================================
		EndText
		
		lcScript = lcHeader + CRLF + "CREATE TABLE " + lcLeft + tcTableName + lcRight + '('
		If lbInsertInternalID
			lcScript = lcScript + lcInternalID + ','
		EndIf
		lcScript = lcScript + lcFieldsScript

		Local loComposedScripts, cValue
		loComposedScripts = CreateObject("Collection")
		If Type('toComposedIndexes') == 'O' and toComposedIndexes.count > 0
			loComposedScripts = this.addComposedIndex(toComposedIndexes)
		EndIf
		
		Local cValue
		cValue = ''
		
		If !this.bExecuteFkScriptSeparately
			* Agregamos las claves foráneas
			If loFkScript.count > 0
				For each cValue in loFkScript
					lcScript = lcScript + ',' + cValue
				EndFor
			EndIf
		EndIf
		
		If !this.bExecuteIndexScriptSeparately
			cValue = ''
			* Agregamos los Índices individuales
			If loIdxScript.count > 0
				For each cValue in loIdxScript
					lcScript = lcScript + ',' + cValue
				EndFor
			EndIf

			If loComposedScripts.count > 0
				cValue = ''
				* Si tenemos Índices compuestos también los agregamos
				For each cValue in loComposedScripts
					lcScript = lcScript + ',' + cValue
				EndFor
			EndIf
		EndIf

		lcScript = lcScript + ')' + This.createTableOptions()
		
		If !Empty(tcTableDescription)
			lcScript = lcScript + this.addTableComment(tcTableDescription)
		EndIf
		
		lcScript = lcScript + ';'
		If Type('toScripts') == 'O'
			toScripts.Add(lcScript + CRLF)
		else
			If !This.SQLExec(lcScript)
				Return .f.
			EndIf
		EndIf

		If this.bExecuteFkScriptSeparately
			cValue = ''
			* Agregamos las claves foráneas
			If loFkScript.count > 0
				For each cValue in loFkScript
					If Type('toScripts') == 'O'
						toScripts.Add(cValue + CRLF)
					else
						This.SQLExec(cValue)
					EndIf
				EndFor
			EndIf
		EndIf
		
		If this.bExecuteIndexScriptSeparately
			cValue = ''
			If loIdxScript.count > 0
				* Ejecutamos los Índices individuales
				For each cValue in loIdxScript
					If Type('toScripts') == 'O'
						toScripts.Add(cValue + CRLF)
					else
						This.SQLExec(cValue)
					EndIf
				EndFor
			EndIf
			
			If loComposedScripts.count > 0
				cValue = ''
				* Ejecutamos los Índices compuestos
				For each cValue in loComposedScripts
					If Type('toScripts') == 'O'
						toScripts.Add(cValue + CRLF)
					else
						This.SQLExec(cValue)
					EndIf
				EndFor
			EndIf
		EndIf
		Return .t.
	endfunc

	Function use(tcTable, tcFields, tcCriteria, tcGroup, tbReadOnly, tbNodata)
		Local lcSqlTableName, lcAlias
		this.getTableAndAlias(tcTable, @lcSqlTableName, @lcAlias)
		
		If Used(lcAlias)
			Return .f.
		EndIf

		If !This.tableExists(lcSqlTableName)
			Text to this.cLastError noshow pretext 7 textmerge
			    Error - Tabla Inexistente:
			    La tabla con el nombre '<<lcSqlTableName>>' no existe en la base de datos.

			    Por favor, asegúrate de que el nombre de la tabla está escrito correctamente y que la tabla haya sido creada previamente en la base de datos.
			EndText
			If this.bShowErrors
				MessageBox(this.cLastError, 16)
			EndIf
			Return .f.
		Endif

		Local lcPrimaryKey, lcSelectCMD, loView
		
		lcPrimaryKey = this.getKeyField(lcSqlTableName)
		lcSelectCMD = this.getSelectCommand(lcSqlTableName, tcFields, tcCriteria)

		If this.bUseCA
			Local i, lcUpdaTableFieldList, lcUpdateNameList, lcField, laFields[1], lcSchemaList
			loView = Createobject('CursorAdapter')
			=AddProperty(loView, "Database", this.cDatabase)
			loView.DataSourceType = 'ODBC'
			loView.Datasource = This.nHandle
			loView.Alias = lcAlias
			loView.SelectCmd = lcSelectCMD
			loView.Tables = lcSqlTableName
			loView.KeyFieldList = lcPrimaryKey
			loView.SendUpdates = !tbReadOnly

			* Traer solo estructura para extraer información de las columnas.
			loView.Nodata = .T.
			If !loView.CursorFill()
				this.sqlError()
				Return .f.
			EndIf
			
			Select (lcAlias)
			Store '' To lcUpdaTableFieldList, lcUpdateNameList, lcSchemaList

			For i=1 To Afields(laFields)
				lcField = laFields[i,1]
				lcUpdaTableFieldList = lcUpdaTableFieldList + lcField + ','
				lcUpdateNameList = lcUpdateNameList + lcField + Space(1) + lcSqlTableName + '.' + lcField + ','

				* Schema
				If !Empty(lcSchemaList)
					lcSchemaList = lcSchemaList + ','
				EndIf
				lcSchemaList = lcSchemaList + lcField + ' ' + laFields[i,2] + ' '
				If laFields[i,3] > 0
					lcSchemaList = lcSchemaList + '(' + Alltrim(Str(laFields[i,3]))
					If laFields[i,4] > 0
						lcSchemaList = lcSchemaList + ',' + Alltrim(Str(laFields[i,3]))
					EndIf
					lcSchemaList = lcSchemaList + ')'
				EndIf
			EndFor
			lcUpdaTableFieldList = Substr(lcUpdaTableFieldList, 1, Len(lcUpdaTableFieldList)-1) && Remove trailing comma
			lcUpdateNameList = Substr(lcUpdateNameList, 1, Len(lcUpdateNameList)-1) 			&& Remove trailing comma
						
			loView.UpdatableFieldList = lcUpdaTableFieldList
			loView.UpdateNameList = lcUpdateNameList
			loView.Nodata = tbNodata
			*loView.CursorSchema = lcSchemaList

			If !loView.CursorFill()
				this.sqlError()
				Return .f.
			EndIf

			=CursorSetProp("FetchSize", -1, lcAlias)
			* Esperar hasta completar todos los registros para eviar error 'Connection is Busy'
			Do While SQLGetprop(This.nHandle, "ConnectBusy")
				If this.bShowErrors
					Wait Window "Recuperando información de la tabla actual, espere..."  Nowait				
				EndIf
				=Inkey(0.3, "H")
				Doevents		
			EndDo			
			If this.bShowErrors
				Wait Clear
			EndIf
			Go top in (lcAlias)
		Else
			If !This.SQLExec(lcSelectCMD, lcAlias)
				Return .F.
			Endif

			loView = Createobject('RemoteCursor')
			With loView
				.Database = This.cDatabase
				.Alias = lcAlias
				.SelectCmd = lcSelectCMD
				.Tables = lcSqlTableName
				.KeyFieldList = lcPrimaryKey
				.SendUpdates = !tbReadOnly
				.Nodata = tbNodata
			EndWith
		EndIf
		If !tbReadOnly
			=CursorSetProp("Buffering", 5, lcAlias)
		EndIf

		This.oViews.Add(loView, Lower(lcAlias))

		If !Empty(tcGroup)
			This.addViewToGroup(Lower(tcGroup), Lower(lcAlias))
		EndIf
		Select (lcAlias)
		Return .t.
	Endproc

	function changeDB(tcNewDatabase)
		* Abstract
	endfunc

	Procedure requery(tcAlias)
		If Empty(tcAlias)
			tcAlias = Alias()
		EndIf

		Select (tcAlias)		
		Local lnIndex, loView, lcCursor
		lnIndex = This.oViews.GetKey(Lower(tcAlias))
		If Empty(lnIndex)
			Return .F.
		Endif			
		loView = This.oViews.Item(lnIndex)		
		If this.bUseCA
			If loView.sendUpdates
				=Requery(tcAlias)
			EndIf
		Else			
			If loView.sendUpdates				
				=TableRevert(.t.)
				Delete from (tcAlias)

				lcCursor = Sys(2015)
				this.sqlExec(loView.SelectCMD, lcCursor)

				Select (tcAlias)
				Append From Dbf(lcCursor)
				Use in (lcCursor)
			EndIf
		EndIf
	EndProc

	Procedure discard(tcAlias)
		If Empty(tcAlias)
			tcAlias = Alias()
		EndIf

		Local lnIndex, loView, lcCursor
		lnIndex = This.oViews.GetKey(Lower(tcAlias))
		If Empty(lnIndex)
			Return .F.
		Endif			
		loView = This.oViews.Item(lnIndex)		
		If loView.sendUpdates
			=TableRevert(.t.)
		EndIf
	endproc

	Function saveAndClose(tcAlias)
		If This.Save(tcAlias)
			This.Close(tcAlias)
		EndIf
	EndFunc

	function Save(tcAlias)
		If Empty(tcAlias)
			tcAlias = Alias()
		Endif
		Local lnOldTransactionSeting, lnIndex, lbOk, loView, loEnv

		lnOldTransactionSeting = SQLGetprop(This.nHandle, "Transactions")
		=SQLSetprop(This.nHandle, "Transactions", 2) && Change to manual transactions

		lnIndex = This.oViews.GetKey(Lower(tcAlias))
		If Empty(lnIndex)
			Return .F.
		Endif
		loEnv = this.setEnvironment()
		loView = This.oViews.Item(lnIndex)
		If !loView.SendUpdates
			Return .F.
		EndIf

		Try
			Begin Transaction
			This.beginTransaction()
			Select (tcAlias)
			If this.bUseCA
				lbOk = TableUpdate(.T., .T.)
				If !lbOk
					this.sqlError()
				EndIf
			Else
				lbOk = .T.
				Select (tcAlias)
				Local nNextRec, lcTypeOpe, lcFldState, lcCommand, lcLeft, lcRight, lcSQLTable, lcScript, laFields[1], laDateFields[1]
				
				lcLeft  = This.cLeft
				lcRight = This.cRight
				lcSQLTable  = loView.Tables
				lcKeyField  = loView.KeyFieldList			
				lcScript	= Space(1)
				nNextRec 	= Getnextmodified(0, tcAlias)
				AFields(laFields)
				laDateFields = this.getDateTimeFields(@laFields)

				Do While nNextRec <> 0
					Go nNextRec in (tcAlias)
					lvKeyValue  = Evaluate(tcAlias + '.' + lcKeyField)
					lcScript	= Space(1)
					lcFldState  = GetFldState(-1)
					
					If nNextRec > 0 && UPDATE			
						Do case
						case Left(lcFldState, 1) == '1' && UPDATE
							lcScript = this.updateRowScript(lcLeft, lcRight, lcSQLTable, lcKeyField, lcFldState)
							Scatter memo name loRow
							this.updateFetchedRow(@laDateFields, loRow)
						Case Left(lcFldState, 1) == '2' && DELETE
							lcScript = "DELETE FROM " + lcLeft + lcSQLTable + lcRight + " WHERE " + lcLeft + lcKeyField + lcRight + "=?lvKeyValue"
						EndCase
					Else && INSERT
						If this.rowExists(lcLeft, lcRight, lcSQLTable, lcKeyField, lvKeyValue)
							lcScript = this.updateRowScript(lcLeft, lcRight, lcSQLTable, lcKeyField, lcFldState)
							Scatter memo name loRow
							this.updateFetchedRow(@laDateFields, loRow)
						Else
							Local j, laInsFields[1], lcFieldsScript, lcValuesScript
						
							laInsFields = this.getAffectedFields(lcFldState, @laFields)
							* Iterate fields
							lcFieldsScript = Space(1)
							lcValuesScript = Space(1)

							For j=1 To Alen(laInsFields, 1)
								If Upper(laInsFields[j]) == Upper(this.cPKName)
									Loop
								EndIf
								lcFieldsScript = lcFieldsScript + lcLeft + laInsFields[j] + lcRight + ','
								lcValuesScript = lcValuesScript + '?loRow.' + laInsFields[j] + ','
							EndFor
							
							lcFieldsScript = Substr(lcFieldsScript, 1, Len(lcFieldsScript)-1)
							lcValuesScript = Substr(lcValuesScript, 1, Len(lcValuesScript)-1)

							Scatter Memo Name loRow
							this.updateFetchedRow(@laDateFields, loRow)
							lcScript = "INSERT INTO " + lcLeft + lcSQLTable + lcRight + " (" + lcFieldsScript + ")"
							lcScript = lcScript + " VALUES (" + lcValuesScript + ");"					
						EndIf															
					EndIf

					If !Empty(lcScript)
						If !this.sqlExec(lcScript)
							lbOk = .F.
							Exit
						EndIf
					EndIf
					nNextRec = Getnextmodified(nNextRec, tcAlias)
				Enddo
			EndIf

			If lbOk
				This.endTransaction()
				End Transaction
			Else
				This.cancelTransaction()
				Rollback
			EndIf
		Catch
			RollBack
		EndTry

		=SQLSetprop(This.nHandle, "Transactions", lnOldTransactionSeting)
		this.restoreEnvironment(loEnv)
		Return lbOk
	endfunc

	function saveGroup(tcGroup)
		If Empty(tcGroup)
			Return .F.
		Endif
		Local lnIndex, loViews, i, lbOk, loView, lcScript, lcAlias, lnOldTransactionSeting
		lnIndex = This.oGroupViews.GetKey(Lower(tcGroup))
		If Empty(lnIndex)
			Return .F.
		Endif

		lcScript = 'set datasession to ' + Alltrim(Str(Set("Datasession"))) + CRLF
		loViews = This.oGroupViews.Item(lnIndex)
		If Empty(loViews.Count)
			Return .F.
		Endif

		lnOldTransactionSeting = SQLGetprop(This.nHandle, "Transactions")
		SQLSetprop(This.nHandle, "Transactions", 2) && Change to manual transactions

		Begin Transaction
		This.beginTransaction()

		For i=1 To loViews.Count
			lcAlias = loViews.Item(i)
			loView = This.oViews.Item(lcAlias)
			If !loView.SendUpdates
				Loop && Ignore cursor
			Endif
			lcScript = lcScript + "select " + loView.Alias + CRLF
			lcScript = lcScript + "=TableRevert(.T.) " + CRLF
			Select (loView.Alias)
			lbOk = Tableupdate(.T., .T.)
			If !lbOk
				this.sqlError()
				Exit
			Endif
		Endfor

		If lbOk
			This.endTransaction()
			End Transaction
		Else
			This.sqlError()
			This.cancelTransaction()
			Rollback
			=Execscript(lcScript)
		Endif
		SQLSetprop(This.nHandle, "Transactions", lnOldTransactionSeting)

		Return lbOk
	EndFunc

	Procedure Close(tcAlias)

		If Empty(tcAlias)
			tcAlias = Alias()
		Endif

		If !Used(tcAlias)
			Return .F.
		Endif

		Local lnIndex, lcAlias, loView

		* Intentamos buscar como Vista
		lnIndex = This.oViews.GetKey(Lower(tcAlias))
		If Empty(lnIndex)
			Return .F.
		Endif
		loView = This.oViews.Item(lnIndex)
		Select (tcAlias)

		If loView.SendUpdates
			=Tablerevert(.T.) && just in case there's pending changes.
		Endif
		Use

		* Release the cursorAdapter allocated in global scope.
		This.oViews.Remove(lnIndex)
		Release loView

		Return .T.
	Endproc

	Procedure closeAll
		Try
			Local i, loView, lcAlias
			lcAlias = Alias()
			For i = 1 To This.oViews.Count
				loView = This.oViews.Item(i)
				If Used(loView.Alias)
					Select (loView.Alias)
					If loView.SendUpdates
						Tablerevert(.T.) && Revert pending changes
					Endif
					Use
				Endif
				Release loView
			Endfor
			If !Empty(lcAlias) And Used(lcAlias)
				Select (lcAlias)
			Endif
			This.oViews = Createobject('Collection')		&& Reset all created views.
			This.oGroupViews = Createobject('Collection')	&& Reset all created groups.
		Catch
		Endtry
	Endproc

	Procedure closeGroup(tcGroup)

		If Empty(tcGroup)
			Return .F.
		Endif
		Local lnIndex, loViews, i, loView, lcAlias
		lnIndex = This.oGroupViews.GetKey(Lower(tcGroup))
		If Empty(lnIndex)
			Return .F.
		Endif

		loViews = This.oGroupViews.Item(lnIndex)
		If Empty(loViews.Count)
			Return .F.
		Endif

		For i=1 To loViews.Count
			lcAlias = loViews.Item(i)
			loView = This.oViews.Item(lcAlias)
			Select (loView.Alias)
			If loView.SendUpdates
				Tablerevert(.T.)
			Endif
			Use
			Release loView
		Endfor

		Return .T.
	Endproc

	Function SQLExec(tcSQLCommand, tcCursorName)

		If Empty(tcCursorName)
			tcCursorName = Sys(2015)
		Endif

		If SQLExec(This.nHandle, tcSQLCommand, tcCursorName) <= 0
			this.sqlError()
			Return .f.
		Endif

		Return .t.
	Endfunc

	Procedure sqlError
		Try
			Release laError
			AError(laError)
			If Type('laError',1) == 'A'
				Acopy(laError, this.aLastError)
				Text to this.cLastError noshow pretext 7 textmerge
				    ERROR - Error en la Consulta SQL:
				    Código de Error: <<laError[1]>>
				    Mensaje de Error: <<Transform(laError[2]) + Transform(laError[3])>>
				    Por favor, revisa la consulta SQL y asegúrate de que está correctamente escrita. Verifica que los nombres de tablas, campos y condiciones sean válidos y vuelve a intentarlo.
				EndText				
				If this.bShowErrors
					MessageBox(this.cLastError, 16)
				EndIf
			EndIf
		Catch to loEx
			this.cLastError = loEx.Message
			If this.bShowErrors
				MessageBox(this.cLastError, 16)
			EndIf
		endtry
	Endproc

	Hidden function updateRowScript(tcOpenChar, tcCloseChar, tcSQLTable, tcKeyField, tcFldState)
		Local laFields[1], i, laUpdFields[1], laDateFields[1], lcScript
		=AFields(laFields)
		
		laUpdFields  = this.getAffectedFields(tcFldState, @laFields)
		laDateFields = this.getDateTimeFields(@laFields)						
		
		lcScript = "UPDATE " + tcOpenChar + tcSQLTable + tcCloseChar + " SET "

		* Iterate fields
		For i=1 to Alen(laUpdFields)
			lcScript = lcScript + tcOpenChar + laUpdFields[i] + tcCloseChar + "=?loRow." + laUpdFields[i] + ','
		EndFor
		lcScript = Substr(lcScript, 1, Len(lcScript)-1)
		
		lcScript = lcScript + " WHERE " + tcOpenChar + tcKeyField + tcCloseChar + "=?lvKeyValue"	
		Return lcScript
	EndFunc
	
	Hidden Function rowExists(tcOpenChar, tcCloseChar, tcSQLTable, tcKeyField, tvkeyValue)
		If Empty(tvkeyValue)
			Return .f.
		EndIf
		Local lcCommand, lcCursor, lnTotal, lcAlias, lvValue
		lcAlias  = Alias()
		lcCursor = Sys(2015)
		lvValue  = tvkeyValue
		lcCommand = "SELECT Count(*) as total FROM " + tcOpenChar + tcSQLTable + tcCloseChar + " WHERE " + tcOpenChar + tcKeyField + tcCloseChar + "=?lvValue"
		If !this.SQLExec(lcCommand, lcCursor)
			Release lvValue
			Return .f.
		EndIf
		Release lvValue
		
		lnTotal = &lcCursor..total
		Use in (lcCursor)
		
		If !Empty(lcAlias) and Used(lcAlias)
			Select (lcAlias)
		EndIf
		
		Return lnTotal > 0
	EndFunc

	Hidden Procedure getTableAndAlias(tcTable, tcSqlTableName, tcAlias)
		Local loResult
		This.oRegEx.Pattern = "^(\w+)\s+[asAS]+\s+(\w+)"
		loResult = This.oRegEx.Execute(tcTable)
		If Type('loResult') == 'O' And loResult.Count > 0
			tcSqlTableName = loResult.Item(0).SubMatches(0)
			tcAlias = loResult.Item(0).SubMatches(1)
		Else
			tcSqlTableName = tcTable
			tcAlias = tcTable
		Endif
	EndProc
	
	Hidden function getKeyField(tcTable)
		If !this.fieldExists(tcTable, this.cPKName)
			Return this.getPrimaryKey(tcTable)
		EndIf
		Return this.cPKName
	EndFunc 

	Hidden function getSelectCommand(tcTable, tcFields, tcCriteria)
		Local lcLeft, lcRight, lcCommand
		lcLeft = this.cLeft
		lcRight = this.cRight

		If Empty(tcFields)
			tcFields = "*"
		EndIf

		lcCommand = "SELECT " + tcFields + " FROM " + lcLeft + tcTable + lcRight
		If !Empty(tcCriteria)
			lcCommand = lcCommand + " WHERE " + tcCriteria
		EndIf
		Return lcCommand
	endfunc

	Hidden Procedure addViewToGroup(tcGroup, tcAlias)

		Local lnIndex, loViews As Collection
		lnIndex = This.oGroupViews.GetKey(Lower(tcGroup))
		If Empty(lnIndex)
			loViews = Createobject('Collection')
		Else
			loViews = This.oGroupViews.Item(lnIndex)
			This.oGroupViews.Remove(lnIndex)
		Endif
		Try
			loViews.Add(tcAlias)
			This.oGroupViews.Add(loViews, tcGroup)
		Catch
			* View already saved.
		Endtry
	Endproc

	Hidden Function getDateTimeFields(taFields as Variant)
		Local i
		This.resetArray()
		External Array taFields
		For i = 1 To alen(taFields, 1)
			If Inlist(taFields[i, 2], 'D', 'T')
				this.pushArray(taFields[i, 1])
			Endif
		Endfor
		Return @this.aCustomArray
	EndFunc
	
	Hidden function getAffectedFields(tcFldState, taFields)
		Local i, j
		This.resetArray()
		For i = 2 to Len(tcFldState)
			If InList(Val(Substr(tcFldState, i, 1)), 2, 4)
				this.pushArray(taFields[i-1, 1])
			EndIf
		EndFor
		Return @this.aCustomArray
	EndFunc	
	
	Hidden procedure updateFetchedRow(taDateFields, toRow)
		External array taDateFields
		If Type('taDateFields[1]') == 'C'
			Local i, lcMacro
			For i=1 To Alen(taDateFields, 1)
				lcMacro = "toRow." + taDateFields[i] + " = this.formatDateOrDateTime(toRow." + taDateFields[i] + ")"
				&lcMacro
			Endfor
		EndIf
	EndProc

	Hidden Procedure applyConnectionSettings
		Set Multilocks On
		SQLSetprop(This.nHandle, 'DisconnectRollback', .T.)
		SQLSetprop(This.nHandle, 'DispWarnings', .F.)
		SQLSetprop(This.nHandle, 'Asynchronous', .F.)
		SQLSetprop(This.nHandle, 'BatchMode', .T.)
		SQLSetprop(This.nHandle, 'IdleTimeout', 0)
		SQLSetprop(This.nHandle, 'QueryTimeOut', 0)
		SQLSetprop(This.nHandle, 'WaitTime', 100)
		This.sendConfigurationQuerys()
	Endproc

	Hidden Function reconnect
		Local lcQuery, lcCursor, lbAlive
		lcQuery  = This.getDummyQuery()
		lcCursor = Sys(2015)
		lbAlive  = SQLExec(This.nHandle, lcQuery, lcCursor) > 0
		If Used(lcCursor)
			Use In (lcCursor)
		EndIf
		If !lbAlive
			This.nHandle = 0
		EndIf
		Return lbAlive
	EndFunc

	Hidden Procedure printException(toError)
		Text to this.cLastError noshow pretext 7 textmerge
		    ERROR - Excepción Controlada:

		    Código de Error: <<toError.ErrorNo>>
		    Línea No.: <<toError.Lineno>>
		    Mensaje: <<toError.Message>>
		    Procedimiento: <<toError.Procedure>>
		    Detalles: <<toError.Details>>
		    Nivel de Pila: <<toError.StackLevel>>
		    Contenido de la Línea: <<toError.LineContents>>
		    Valor de Usuario: <<toError.UserValue>>

		    Por favor, toma nota de la información proporcionada y contacta al equipo de soporte para obtener asistencia adicional en la resolución de este problema.
		EndText
		If this.bShowErrors
			Messagebox(this.cLastError, 16)
		EndIf
	Endproc

	Procedure disconnect
		Try
			If This.nHandle > 0
				SQLDisconnect(This.nHandle)
				This.nHandle = 0
			Endif
		Catch
		Endtry
	EndProc

	Procedure resetArray
		Dimension This.aCustomArray[1]
		This.aCustomArray[1] = .F.
		This.nCounter = 0
	Endproc

	Procedure pushArray(tvValue)
		This.nCounter = This.nCounter + 1
		Dimension This.aCustomArray[this.nCounter]
		This.aCustomArray[this.nCounter] = tvValue
	Endproc

	Function bUseSymbolDelimiter_Assign(tvNewVal)
		this.bUseSymbolDelimiter = tvNewVal
	EndFunc

	Procedure Destroy
		This.disconnect()
	Endproc

	* ================================================================================ *
	* Abstracts methods
	* ================================================================================ *
	Function getDummyQuery
		* Abstract
	Endfunc

	Function getVersion
		* Abstract
	Endfunc

	Procedure getConnectionString
		* Abstract
	Endproc

	Procedure beginTransaction
		* Abstract
	Endproc

	Procedure endTransaction
		* Abstract
	Endproc

	Procedure cancelTransaction
		* Abstract
	Endproc

	Function tableExists(tcTableName)
		Local lcQuery, lcSchema, lcCursor
		lcCursor = Sys(2015)
		This.selectDatabase()
		lcQuery = this.getTableExistsScript(tcTableName)

		If !This.SQLExec(lcQuery, lcCursor)
			Return .F.
		Endif

		lcSchema = Alltrim(Strtran(&lcCursor..TableName, Chr(0)))

		Use In (lcCursor)

		Return !Empty(lcSchema)
	Endfunc

	Procedure selectDatabase
		* Abstract
	Endproc

	Function getTidScript
		* Abstract
	EndFunc
	
	Function getUUIDScript
		* Abstract
	EndFunc
	
	Function fieldExists(tcTable, tcField)
		Local lcQuery, lcCursor, lbResult
		This.selectDatabase()
		lcCursor = Sys(2015)

		lcQuery = this.getFieldExistsScript(tcTable, tcField)
		This.SQLExec(lcQuery, lcCursor)

		lbResult = !Empty(&lcCursor..fieldName)
		Use In (lcCursor)

		Return lbResult
	Endfunc

	Function getServerDate
		Local lcCursor, ldDate
		lcCursor = Sys(2015)
		This.SQLExec(this.getServerDateScript(), lcCursor)

		ldDate = &lcCursor..sertime

		Use In (lcCursor)

		Return ldDate
	Endfunc

	Function getNewGuid
		Local lcCursor, lcGuid
		lcCursor = Sys(2015)
		This.SQLExec(this.getNewGuidScript(), lcCursor)

		lcGuid = &lcCursor..guid
		Use In (lcCursor)

		Return lcGuid
	Endfunc

	Function getTables
		This.resetArray()

		Local lcQuery, lcCursor
		This.selectDatabase()
		lcCursor = Sys(2015)

		Dimension laTables[1]

		lcQuery = this.getTablesScript()
		This.SQLExec(lcQuery, lcCursor)

		Select table_name From (lcCursor) Into Array laTables

		=Acopy(laTables, This.aCustomArray)

		Use In (lcCursor)

		Return @This.aCustomArray
	Endfunc

	Function getTableFields(tcTable)
		This.resetArray()

		Local lcQuery, lcCursor
		This.selectDatabase()
		lcCursor = Sys(2015)

		Dimension laFields[1]

		lcQuery = this.getTableFieldsScript(tcTable)
		This.SQLExec(lcQuery, lcCursor)

		Select column_name From (lcCursor) Into Array laFields

		=Acopy(laFields, This.aCustomArray)

		Use In (lcCursor)

		Return @This.aCustomArray
	EndFunc
	
*!*		Function getFields
*!*			
*!*		EndFunc

	Procedure getPrimaryKey(tcTable)
		Local lcScript, lcCursor, lcField
		lcCursor = Sys(2015)
		lcScript = this.getPrimaryKeyScript(tcTable)

		This.SQLExec(lcScript, lcCursor)
		lcField = Alltrim(Strtran(&lcCursor..column_name, Chr(0)))
		Use In (lcCursor)

		Return lcField
	Endproc

	Function createTableOptions
		* Abstract
	Endfunc

	Procedure sendConfigurationQuerys
		* Abstract
	Endproc

	function setEnvironment
		* Abstract
	EndFunc
	
	Procedure restoreEnvironment(toEnv)
		* Abstract
	endproc

	Function formatDateOrDateTime(tdValue)
		* Abstract
	Endfunc

	Function getLastID
		Local lcScript, lcCursor, lnID
		lcCursor = Sys(2015)
		lcScript = this.getLastIDScript()
		This.SQLExec(lcScript, lcCursor)
		lnID = &lcCursor..last_id
		Use In (lcCursor)

		If IsNull(lnID)
			Return 0
		EndIf
		Return lnID

	EndFunc

	Function getCreateDatabaseScript(tcDatabase)
		* Abstract
	EndFunc

	Function getDataBaseExistsScript(tcDatabase)
		* Abstract
	EndFunc
	
	Function addFieldComment(tcComment)
		* Abstract
	EndFunc
	
	Function addTableComment(tcComment)
		* Abstract
	EndFunc
	
	Function addForeignKey(toFkData)
		* Abstract
	EndFunc
	
	Function addSingleIndex(toIndex)
		* Abstract
	EndFunc
	
	Function addComposedIndex(toIndex)
		* Abstract
	EndFunc

	Function getForeignKeyValue(tcValue)
		Do case
		Case Upper(tcValue) == 'NULL'
			Return 'SET NULL'
		Case Upper(tcValue) == 'DEFAULT'
			Return 'SET DEFAULT'
		Case Upper(tcValue) == 'RESTRICT'
			Return 'NO ACTION'
		EndCase
		Return tcValue
	EndFunc

	Function Ping
		Return "Pong"
	EndFunc

	Function addAutoIncrement
		* Abstract
	EndFunc
	
	Function addPrimaryKey
		* Abstract
	EndFunc
	
	function dropTable(tcTable)
		* Abstract
	EndFunc
	
	Function addFieldOptions(toField)
		* Abstract
	EndFunc

	Function changeTypeOnAutoIncrement(tcType)
		* Abstract
	EndFunc

	Function visitCType(toFields)
		* Abstract
	Endfunc

	Function visitYType(toFields)
		* Abstract
	Endfunc

	Function visitDType(toFields)
		* Abstract
	Endfunc

	Function visitTType(toFields)
		* Abstract
	Endfunc

	Function visitBType(toFields)
		* Abstract
	Endfunc

	Function visitFType(toFields)
		* Abstract
	Endfunc

	Function visitGType(toFields)
		* Abstract
	Endfunc

	Function visitIType(toFields)
		* Abstract
	Endfunc

	Function visitLType(toFields)
		* Abstract
	Endfunc

	Function visitMType(toFields)
		* Abstract
	Endfunc

	Function visitNType(toFields)
		* Abstract
	Endfunc

	Function visitQType(toFields)
		* Abstract
	Endfunc

	Function visitVType(toFields)
		* Abstract
	Endfunc

	Function visitWType(toFields)
		* Abstract
	EndFunc
Enddefine

* ==================================================== *
* MICROSOFT SQL SERVER
* ==================================================== *
Define Class MSSQL As DBEngine

	Procedure init
		DoDefault()
		this.nMaxLength = 128
		this.cLeft = '['
		this.cRight = ']'
		this.bExecuteFkScriptSeparately = .f.
		this.bExecuteIndexScriptSeparately = .f.
	endproc

	Function getDummyQuery
		Return "SELECT @@VERSION"
	Endfunc

	Function getVersion
		Local lcCursor, lcVersion
		lcCursor = Sys(2015)
		This.SQLExec("SELECT @@VERSION AS 'VER'", lcCursor)
		lcVersion = &lcCursor..VER
		Use In (lcCursor)

		Return lcVersion
	Endfunc

	Function getConnectionString(tbAddDatabase)
		Local lcConStr, lcDriver

		lcConStr = "DRIVER=" + This.cDriver + ";SERVER=" + This.cServer + ";UID=" + This.cUser + ";PWD=" + This.cPassword
		If This.nPort > 0
			lcConStr = lcConStr + ";PORT=" + Alltrim(Str(This.nPort))
		Endif
		If tbAddDatabase
			lcConStr = lcConStr + ";DATABASE=" + Alltrim(This.cDatabase)
		Endif
		lcConStr = lcConStr + ";APP=" + this.name + " by FoxRemote"
		Return lcConStr
	Endfunc

	Procedure beginTransaction
		This.selectDatabase()
		This.SQLExec("BEGIN TRANSACTION")
	Endproc

	Procedure endTransaction
		This.selectDatabase()
		This.SQLExec("IF @@TRANCOUNT > 0 COMMIT")
	Endproc

	Procedure cancelTransaction
		This.selectDatabase()
		This.SQLExec("IF @@TRANCOUNT > 0 ROLLBACK")
	Endproc

	Function getTableExistsScript(tcTableName)
		Local lcQuery

		TEXT to lcQuery noshow pretext 7 textmerge
			SELECT TABLE_SCHEMA AS TableName FROM INFORMATION_SCHEMA.TABLES
			 WHERE TABLE_CATALOG = '<<Alltrim(This.cDatabase)>>'
			 AND  TABLE_NAME = '<<tcTableName>>'
		ENDTEXT

		Return lcQuery
	Endfunc

	Procedure selectDatabase
		If Empty(this.cDatabase)
			Text to this.cLastError noshow pretext 7 textmerge
			    ERROR - Base de Datos No Especificada:

			    Antes de realizar esta petición, asegúrate de haber seleccionado una base de datos para trabajar.

			    Por favor, selecciona una base de datos válida y vuelve a intentar la operación.
			EndText
			If this.bShowErrors
				Messagebox(this.cLastError, 16, "Error: Base de Datos No Especificada")
			EndIf
			Return
		EndIf
		This.SQLExec("use " + This.cDatabase)
	Endproc

	Function getTidScript
		Return "INT IDENTITY(1,1) PRIMARY KEY"
	Endfunc

	Function getUUIDScript
		Return "UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID()"
	EndFunc
	
	Function getFieldExistsScript(tcTable, tcField)
		Local lcQuery

		TEXT to lcQuery noshow pretext 7 textmerge
			SELECT COLUMN_NAME AS FieldName
			FROM INFORMATION_SCHEMA.COLUMNS
			WHERE TABLE_NAME = '<<tcTable>>' AND COLUMN_NAME = '<<tcField>>';
		ENDTEXT

		Return lcQuery
	Endfunc

	Function getServerDateScript
		Return "SELECT GETDATE() AS SERTIME;"
	Endfunc

	Function getNewGuidScript
		Return "SELECT NEWID() AS GUID;"
	Endfunc

	Function getTablesScript
		Local lcQuery

		TEXT TO lcQuery NOSHOW PRETEXT 7 TEXTMERGE
			SELECT TABLE_NAME
			FROM INFORMATION_SCHEMA.TABLES
			WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG = '<<this.cDatabase>>';
		ENDTEXT

		Return lcQuery
	Endfunc

	Function getTableFieldsScript(tcTable)
		Local lcQuery

		TEXT to lcQuery noshow pretext 7 textmerge
			SELECT COLUMN_NAME
			FROM INFORMATION_SCHEMA.COLUMNS
			WHERE TABLE_NAME = '<<tcTable>>';
		ENDTEXT

		Return lcQuery
	Endfunc

	Procedure getPrimaryKeyScript(tcTable)
		Local lcScript
		TEXT to lcScript noshow pretext 7 textmerge
			SELECT COLUMN_NAME
			FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE
			WHERE CONSTRAINT_NAME = (
			    SELECT name
			    FROM sys.key_constraints
			    WHERE type = 'PK'
			        AND OBJECT_NAME(parent_object_id) = '<<this.cLeft + tcTable + this.cRight>>'
			);
		ENDTEXT

		Return lcScript
	Endproc

	Function createTableOptions
		Return " "
	Endfunc

	Procedure sendConfigurationQuerys
		* Abstract
	Endproc

	function setEnvironment
		Local loEnv
		loEnv = CreateObject("Collection")		
		loEnv.Add(Set("Date"), 'date')
		loEnv.Add(Set("Century"), 'century')
		
		Set Date To Dmy
		Set Century On
		This.SQLExec("SET DATEFORMAT dmy")
		
		Return loEnv
	EndFunc

	procedure restoreEnvironment(toEnv)
		Local lcDate, lcCentury
		lcDate = toEnv.Item(1)
		lcCentury = toEnv.Item(2)

		Set Date (lcDate)
		Set Century &lcCentury
	endproc

	Function formatDateOrDateTime(tdValue)
		If Empty(tdValue)
			If Type('tdValue') == 'D'
				Return Date(1753, 01, 01)
			Endif
			Return Datetime(1753,01,01,00,00,00)
		Endif
		Return tdValue
	Endfunc	

	Function getLastIDScript
		Return "SELECT SCOPE_IDENTITY() AS LAST_ID"
	EndFunc

	Function getCreateDatabaseScript(tcDatabase)			
		Return "CREATE DATABASE " + this.cLeft + tcDatabase + this.cRight + ";"
	EndFunc
	
	Function getDataBaseExistsScript(tcDatabase)
		Return "select NAME AS dbName from sys.databases where name = '" + tcDatabase + "'"
	EndFunc

	Function addFieldComment(tcComment)
		Return " "
	EndFunc

	Function addTableComment(tcComment)
		Return " "
	EndFunc

	Function addForeignKey(toFkData)
		Local lcScript, lcOnUpdate, lcOnDelete, lcLeft, lcRight
		lcOnUpdate = this.getForeignKeyValue(toFkData.cOnUpdate)
		lcOnDelete = this.getForeignKeyValue(toFkData.cOnDelete)
		Store "" to lcLeft, lcRight
		
		lcLeft = this.cLeft
		lcRight = this.cRight
				
		Text to lcScript noshow pretext 7 textmerge
		FOREIGN KEY (<<lcLeft+toFkData.cCurrentField+lcRight>>) REFERENCES <<lcLeft+toFkData.cTable+lcRight>>(<<lcLeft+toFkData.cField+lcRight>>)
		ON UPDATE <<lcOnUpdate>>
		ON DELETE <<lcOnDelete>>
		endtext
		Return lcScript
	EndFunc

	Function addSingleIndex(toIndex)
		Local lcScript, lcLeft, lcRight
		Store "" to lcLeft, lcRight
		
		lcLeft = this.cLeft
		lcRight = this.cRight

		Text to lcScript noshow pretext 7 textmerge
			INDEX <<lcLeft+toIndex.cName+lcRight>> <<Iif(toIndex.bUnique, "UNIQUE", "")>> (<<lcLeft+toIndex.cField+lcRight>> <<toIndex.cSort>>)
		EndText

		Return lcScript
	EndFunc

	Function addComposedIndex(toIndex)
		Local lcScript, lcLeft, lcRight, lcColumns, loResult
		Store "" to lcLeft, lcRight, lcColumns

		lcLeft = this.cLeft
		lcRight = this.cRight
		loResult = CreateObject("Collection")

		For each loComposed in toIndex			
			For each loColumn in loComposed.oColumns
				lcColumns = ''
				For each loField in loColumn
					If !Empty(lcColumns)
						lcColumns = lcColumns + ','
					EndIf
					lcColumns = lcColumns + ' ' + lcLeft + loField.cName + lcRight + ' ' + loField.cSort
				EndFor

				Text to lcScript noshow pretext 7 textmerge
					INDEX <<lcLeft + loComposed.cName + lcRight>> <<Iif(loComposed.bUnique, "UNIQUE", "")>> (<<lcColumns>>)
				EndText
				loResult.Add(lcScript)
			EndFor
		EndFor
		Return loResult
	EndFunc

	Function addAutoIncrement
		Return "IDENTITY(1,1)"
	EndFunc
	
	Function addPrimaryKey
		Return "PRIMARY KEY"
	EndFunc	
	
	function dropTable(tcTable)			
		Return "DROP TABLE IF EXISTS " + this.cLeft + tcTable + this.cRight + ';'
	endfunc

	Function addFieldOptions(toField)
		Local lcScript
		lcScript = ''
		If toField.autoIncrement
			lcScript = lcScript + ' ' + this.addAutoIncrement()
		EndIf

		If toField.addDefault AND !toField.autoIncrement
			lcScript = lcScript + " DEFAULT " + toField.Default
		EndIf
		
		If !toField.allowNull
			lcScript = lcScript + " NOT NULL "				
		EndIf
		
		If toField.primaryKey
			lcScript = lcScript + ' ' + this.addPrimaryKey()
		EndIf

		Return lcScript
	EndFunc

	function changeDB(tcNewDatabase)
		If Empty(tcNewDatabase)
		    Text to this.cLastError noshow pretext 7 textmerge
		        ERROR - Base de Datos No Especificada:

		        No has especificado el nombre de la base de datos que deseas utilizar. Por favor, asegúrate de proporcionar el nombre de la base de datos y vuelve a intentarlo.
		    EndText
		    If this.bShowErrors
			    MessageBox(this.cLastError, 16)
			EndIf
		    Return .f.
		Endif
		This.cDatabase = tcNewDatabase
		This.selectDatabase()
		Return .t.
	EndFunc

	Function changeTypeOnAutoIncrement(tcType)
		Return tcType
	EndFunc

	* C = Character
	Function visitCType(toFields)
		Return "CHAR(" + toFields.Size + ") COLLATE Latin1_General_CI_AI"
	Endfunc

	* Y = Currency
	Function visitYType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0"
		endif
		Return "MONEY"
	Endfunc

	* D = Date
	Function visitDType(toFields)
		if toFields.Default == "''"
			toFields.Default = "'1753-01-01'"
		endif
		Return "DATE"
	Endfunc

	* T = DateTime
	Function visitTType(toFields)
		if toFields.Default == "''"
			toFields.Default = "'1753-01-01 00:00:00.000'"
		endif
		Return "DATETIME"
	Endfunc

	* B = Double
	Function visitBType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0.0"
		endif
		Return "FLOAT"
	Endfunc

	* F = Float
	Function visitFType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0.0"
		endif
		Return "FLOAT"
	Endfunc

	* G = General
	Function visitGType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0x"
		endif
		Return "IMAGE"
	Endfunc

	* I = Integer
	Function visitIType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0"
		endif
		Return "INT"
	Endfunc

	* L = Logical
	Function visitLType(toFields)
		Return "BIT"
	Endfunc

	* M = Memo
	Function visitMType(toFields)
		Return "TEXT"
	Endfunc

	* N = Numeric
	Function visitNType(toFields)
		If Val(toFields.Decimal) > 0
			if toFields.Default == "''"
				toFields.Default = "0.0"
			endif
		Else
			if toFields.Default == "''"
				toFields.Default = '0'
			endif
		Endif
		Return "NUMERIC(" + toFields.Size + "," + toFields.Decimal + ")"
	Endfunc

	* Q = VarBinary
	Function visitQType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0x"
		endif
		Return "VARBINARY(max)"
	Endfunc

	* V = Varchar
	Function visitVType(toFields)
		Return "VARCHAR(" + Iif(Empty(Val(toFields.Size)), 'max', toFields.Size) + ") COLLATE Latin1_General_CI_AI"
	Endfunc

	* W = Blob
	Function visitWType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0x"
		endif
		Return "IMAGE"
	EndFunc
Enddefine

* ==================================================== *
* MySQL
* ==================================================== *
Define Class MySQL As DBEngine

	Procedure init
		DoDefault()
		this.nMaxLength = 64
		this.cLeft = '`'
		this.cRight = '`'
	endproc

	Function getDummyQuery
		Return "SELECT Version()"
	Endfunc

	Function getVersion
		Local lcCursor, lcVersion
		lcCursor = Sys(2015)
		This.SQLExec("SELECT Version() AS 'VER'", lcCursor)
		lcVersion = &lcCursor..VER
		Use In (lcCursor)

		Return lcVersion
	Endfunc

	Function getConnectionString(tbAddDatabase)
		Local lcConStr, lcDriver

		lcConStr = "DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";USER=" + This.cUser + ";PASSWORD=" + This.cPassword
		If This.nPort > 0
			lcConStr = lcConStr + ";PORT=" + Alltrim(Str(This.nPort))
		Endif
		If tbAddDatabase
			lcConStr = lcConStr + ";DATABASE=" + Alltrim(This.cDatabase)
		EndIf
		lcConStr = lcConStr + ";COLLATION=utf8_general_ci"
		lcConStr = lcConStr + ";APP=" + this.name + " by FoxRemote"
		Return lcConStr
	Endfunc

	Procedure beginTransaction
		This.selectDatabase()
		This.SQLExec("START TRANSACTION;")
	Endproc

	Procedure endTransaction
		This.selectDatabase()
		This.SQLExec("COMMIT;")
	Endproc

	Procedure cancelTransaction
		This.selectDatabase()
		This.SQLExec("ROLLBACK;")
	Endproc

	Function getTableExistsScript(tcTableName)
		Local lcQuery

		TEXT to lcQuery noshow pretext 7 textmerge
			SELECT TABLE_NAME AS TableName FROM INFORMATION_SCHEMA.TABLES
			 WHERE TABLE_SCHEMA = '<<Alltrim(This.cDatabase)>>'
			 AND  TABLE_NAME = '<<tcTableName>>'
		ENDTEXT

		Return lcQuery
	Endfunc

	Procedure selectDatabase
		IF NOT This.SQLExec("use " + This.cDatabase)
			this.cLastError = "No se pudo seleccionar la base de datos: '" + this.cDatabase + "'"
			If this.bShowErrors
				MESSAGEBOX(this.cLastError, 16, "Error")
			EndIf
			RETURN .f.
		ENDIF
	Endproc

	Function getTidScript
		Return "INT AUTO_INCREMENT PRIMARY KEY"
	Endfunc

	Function getUUIDScript
		Return "VARCHAR(36) PRIMARY KEY NOT NULL"
	EndFunc

	Function getFieldExistsScript(tcTable, tcField)
		Local lcQuery

		TEXT to lcQuery noshow pretext 7 textmerge
			SELECT COLUMN_NAME AS FieldName
			FROM INFORMATION_SCHEMA.COLUMNS
			WHERE TABLE_NAME = '<<tcTable>>' AND TABLE_SCHEMA = '<<Alltrim(This.cDatabase)>>' AND COLUMN_NAME = '<<tcField>>';
		ENDTEXT

		Return lcQuery
	Endfunc

	Function getServerDateScript
		Return "SELECT NOW() AS SERTIME;"
	Endfunc

	Function getNewGuidScript
		Return "SELECT UUID() AS GUID;"
	Endfunc

	Function getTablesScript
		Local lcQuery

		TEXT TO lcQuery noshow pretext 7 textmerge
			SELECT TABLE_NAME
			FROM INFORMATION_SCHEMA.TABLES
			WHERE TABLE_SCHEMA = '<<this.cDatabase>>';
		ENDTEXT

		Return lcQuery
	Endfunc

	Function getTableFieldsScript(tcTable)

		Local lcQuery

		TEXT to lcQuery noshow pretext 7 textmerge
			SELECT COLUMN_NAME
			FROM INFORMATION_SCHEMA.COLUMNS
			WHERE TABLE_NAME = '<<this.cLeft+tcTable+this.cRight>>' AND TABLE_SCHEMA = '<<this.cLeft+this.cDatabase+this.cRight>>';
		ENDTEXT

		Return lcQuery
	Endfunc

	Procedure getPrimaryKeyScript(tcTable)
		Local lcScript, lcLeft, lcRight
		lcLeft = this.cLeft
		lcRight = this.cRight
		TEXT to lcScript noshow pretext 7 textmerge
			SELECT COLUMN_NAME
			FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE
			WHERE TABLE_SCHEMA = '<<lcLeft+this.cDatabase+lcRight>>'
			  AND TABLE_NAME = '<<lcLeft+tcTable+lcRight>>'
			  AND CONSTRAINT_NAME = 'PRIMARY';
		ENDTEXT

		Return lcScript
	Endproc

	Function createTableOptions
		Return " ENGINE = InnoDB AUTO_INCREMENT = 0 DEFAULT CHARSET = latin1"
	EndFunc
	
	Procedure sendConfigurationQuerys
		
	EndProc

	function setEnvironment
		Local loEnv
		loEnv = CreateObject("Collection")		
		loEnv.Add(Set("Date"), 'date')
		loEnv.Add(Set("Century"), 'century')
		loEnv.Add(Set("Mark"), 'mark')				

		Set Date To YMD
		Set Century On
		Set Mark To '-'
		
		* Deshabilitar la validación de claves foráneas
		this.sqlexec("SET FOREIGN_KEY_CHECKS=0;")

		Return loEnv
	EndFunc

	procedure restoreEnvironment(toEnv)
		Local lcDate, lcCentury, lcMark
		lcDate = toEnv.Item(1)
		lcCentury = toEnv.Item(2)
		lcMark = toEnv.Item(3)

		Set Date (lcDate)
		Set Century &lcCentury
		Set Mark to (lcMark)

		* Habilitar la validación de claves foráneas
		this.sqlexec("SET FOREIGN_KEY_CHECKS=1;")

	endproc

	Function formatDateOrDateTime(tdValue)
		If Empty(tdValue)
			If Type('tdValue') == 'D'
				Return Date(1000, 01, 01)
			Endif
			Return Datetime(1000,01,01,00,00,00)
		Endif
		Return tdValue
	Endfunc

	Function getLastIDScript
		Return "SELECT LAST_INSERT_ID() AS LAST_ID"
	EndFunc

	Function getCreateDatabaseScript(tcDatabase)
		Return "CREATE DATABASE " + this.cLeft + tcDatabase + this.cRight + " DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;"
	EndFunc
	
	Function getDataBaseExistsScript(tcDatabase)
		Return "SELECT SCHEMA_NAME AS dbName FROM information_schema.schemata WHERE schema_name = '" + tcDatabase + "'"
	EndFunc

	Function addFieldComment(tcComment)
		Return " COMMENT '" + tcComment + "'"
	EndFunc
	
	Function addTableComment(tcComment)
		Return " COMMENT '" + tcComment + "'"
	EndFunc	

	Function addForeignKey(toFkData)
		Local lcScript, lcOnUpdate, lcOnDelete, lcOPen, lcClose
		lcOnUpdate = this.getForeignKeyValue(toFkData.cOnUpdate)
		lcOnDelete = this.getForeignKeyValue(toFkData.cOnDelete)
		Store "" to lcOPen, lcClose
		lcOPen = this.cLeft
		lcClose = this.cRight
				
		Text to lcScript noshow pretext 7 textmerge
		FOREIGN KEY (<<lcOPen+toFkData.cCurrentField+lcClose>>) REFERENCES <<lcOPen+toFkData.cTable+lcClose>>(<<lcOPen+toFkData.cField+lcClose>>)
		ON UPDATE <<lcOnUpdate>>
		ON DELETE <<lcOnDelete>>
		endtext		
		Return lcScript
	EndFunc

	Function addAutoIncrement
		Return "AUTO_INCREMENT"
	EndFunc
	
	Function addPrimaryKey
		Return "PRIMARY KEY"
	EndFunc	

	function dropTable(tcTable)
		Return "DROP TABLE IF EXISTS " + this.cLeft + tcTable + this.cRight + ';'
	endfunc

	Function addSingleIndex(toIndex)
		Local lcScript, lcLeft, lcRight
		Store "" to lcLeft, lcRight
		lcLeft = this.cLeft
		lcRight = this.cRight
		Text to lcScript noshow pretext 7 textmerge
			<<Iif(toIndex.bUnique, "UNIQUE", "")>> INDEX <<lcLeft+toIndex.cName+lcRight>> (<<lcLeft+toIndex.cField+lcRight>> <<toIndex.cSort>>)
		EndText

		Return lcScript
	EndFunc

	Function addComposedIndex(toIndex)
		Local lcScript, lcLeft, lcRight, lcColumns, loResult
		Store "" to lcLeft, lcRight, lcColumns

		lcLeft = this.cLeft
		lcRight = this.cRight
		loResult = CreateObject("Collection")

		For each loComposed in toIndex			
			For each loColumn in loComposed.oColumns
				lcColumns = ''
				For each loField in loColumn
					If !Empty(lcColumns)
						lcColumns = lcColumns + ','
					EndIf
					lcColumns = lcColumns + ' ' + lcLeft + loField.cName + lcRight + ' ' + loField.cSort
				EndFor

				Text to lcScript noshow pretext 7 textmerge
					<<Iif(loComposed.bUnique, "UNIQUE", "")>> INDEX <<lcLeft + loComposed.cName + lcRight>> (<<lcColumns>>)
				EndText
				loResult.Add(lcScript)
			EndFor
		EndFor
		Return loResult
	EndFunc

	Function addFieldOptions(toField)
		Local lcScript
		lcScript = ''
		If toField.autoIncrement
			lcScript = lcScript + ' ' + this.addAutoIncrement()
		EndIf

		If toField.addDefault AND !toField.autoIncrement
			lcScript = lcScript + " DEFAULT " + toField.Default
		EndIf
		
		If !toField.allowNull
			lcScript = lcScript + " NOT NULL "				
		EndIf
		
		If toField.primaryKey
			lcScript = lcScript + ' ' + this.addPrimaryKey()
		EndIf									

		If !Empty(toField.comment)
			lcScript = lcScript + this.addFieldComment(toField.comment)
		EndIf

		Return lcScript
	EndFunc

	Function changeDB(tcNewDatabase)
		If Empty(tcNewDatabase)
		    Text to this.cLastError noshow pretext 7 textmerge
		        ERROR - Base de Datos No Especificada:

		        No has especificado el nombre de la base de datos que deseas utilizar. Por favor, asegúrate de proporcionar el nombre de la base de datos y vuelve a intentarlo.
		    EndText
		    If this.bShowErrors
			    MessageBox(this.cLastError, 16)
			EndIf
		    Return .f.
		Endif
		This.cDatabase = tcNewDatabase
		This.selectDatabase()
		Return .t.
	EndFunc
	
	Function changeTypeOnAutoIncrement(tcType)
		Return tcType
	EndFunc

	* C = Character
	Function visitCType(toFields)
		Return "CHAR(" + toFields.Size + ")"
	Endfunc

	* Y = Currency
	Function visitYType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0.0"
		endif
		Return "DECIMAL(" + toFields.Size + "," + toFields.Decimal + ")"
	Endfunc

	* D = Date
	Function visitDType(toFields)
		if toFields.Default == "''"
			toFields.Default = "'1000-01-01'"
		endif
		Return "DATE"
	Endfunc

	* T = DateTime
	Function visitTType(toFields)
		if toFields.Default == "''"
			toFields.Default = "'1000-01-01 00:00:00'"
		endif
		Return "DATETIME"
	Endfunc

	* B = Double
	Function visitBType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0.0"
		endif
		Return "DOUBLE(" + toFields.Size + "," + toFields.Decimal + ")"
	Endfunc

	* F = Float
	Function visitFType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0.0"
		endif
		Return "FLOAT(" + toFields.Size + "," + toFields.Decimal + ")"
	Endfunc

	* G = General
	Function visitGType(toFields)
		toFields.addDefault = .F.
		Return "BLOB"
	Endfunc

	* I = Integer
	Function visitIType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0"
		endif
		Return "INT"
	Endfunc

	* L = Logical
	Function visitLType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0"
		endif
		Return "BOOL"
	Endfunc

	* M = Memo
	Function visitMType(toFields)
		toFields.addDefault = .F.
		Return "TEXT"
	Endfunc

	* N = Numeric
	Function visitNType(toFields)
		If Val(toFields.Decimal) > 0
			if toFields.Default == "''"
				toFields.Default = "0.0"
			endif
		else
			if toFields.Default == "''"
				toFields.Default = '0'
			endif
		Endif
		Return "DECIMAL(" + toFields.Size + "," + toFields.Decimal + ")"
	Endfunc

	* Q = VarBinary
	Function visitQType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0x"
		endif
		Local lcSize
		lcSize = "255"
		If Val(toFields.Size) > 0
			lcSize = toFields.Size
		Endif
		Return "VARBINARY(" + lcSize + ")"
	Endfunc

	* V = Varchar (usamos NVARCHAR(N) para admitir caracteres especiales)
	Function visitVType(toFields)
		* Return "VARCHAR(" + toFields.Size + ") CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
		Return "VARCHAR(" + toFields.Size + ")"
	Endfunc

	* W = Blob
	Function visitWType(toFields)
		toFields.addDefault = .F.
		Return "BLOB"
	EndFunc
EndDefine

* ==================================================== *
* MariaDB
* ==================================================== *
Define Class MariaDB As MySQL

EndDefine

* ==================================================== *
* FireBird
* ==================================================== *
Define Class Firebird As DBEngine

	Procedure init
		DoDefault()	
		this.nMaxLength = 31
		this.cLeft = '"'
		this.cRight = '"'
		this.bExecuteFkScriptSeparately = .t.
		this.bExecuteIndexScriptSeparately = .t.
		this.bCanGenerateGUID = .f.
	EndProc

    Function getDummyQuery
        Return "SELECT 1 FROM RDB$DATABASE"
    Endfunc

    Function getVersion
        Local lcCursor, lcVersion
        lcCursor = Sys(2015)
        This.SQLExec("SELECT RDB$GET_CONTEXT('SYSTEM', 'ENGINE_VERSION') AS 'VER' FROM RDB$DATABASE", lcCursor)
        lcVersion = &lcCursor..VER
        Use In (lcCursor)

        Return lcVersion
    Endfunc

    Function getConnectionString(tbAddDatabase)
        Local lcConStr, lcPort
        lcConStr = "DRIVER=" + this.cDriver + ";DBNAME=" + This.cDatabase + ";UID=" + This.cUser + ";PWD=" + This.cPassword
		lcConStr = lcConStr + ";APP=" + this.name + " by FoxRemote"
        Return lcConStr
    Endfunc

    Procedure beginTransaction
        This.selectDatabase()
        This.SQLExec("SET TRANSACTION")
    Endproc

    Procedure endTransaction
        This.selectDatabase()
        This.SQLExec("COMMIT")
    Endproc

    Procedure cancelTransaction
        This.selectDatabase()
        This.SQLExec("ROLLBACK")
    Endproc

    Function getTableExistsScript(tcTableName)
        Local lcQuery, lcLeft, lcRight
		Store "" to lcLeft, lcRight
		
		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		Else
			tcTableName = Upper(tcTableName)
		EndIf
		
        TEXT to lcQuery noshow pretext 7 textmerge
            SELECT RDB$RELATION_NAME AS TableName
            FROM RDB$RELATIONS
            WHERE RDB$RELATION_NAME = '<<lcLeft+tcTableName+lcRight>>'
        ENDTEXT

        Return lcQuery
    Endfunc

    Procedure selectDatabase
        * No aplica
    Endproc

    Function getTidScript
        Return "INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY"
    Endfunc

	Function getUUIDScript
		Return "VARCHAR(36) NOT NULL PRIMARY KEY"
	EndFunc

    Function getFieldExistsScript(tcTable, tcField)
        Local lcQuery, lcLeft, lcRight
		Store "" to lcLeft, lcRight
		
		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		Else
			tcTable = Upper(tcTable)
			tcField = Upper(tcField)
		EndIf
		
        TEXT to lcQuery noshow pretext 7 textmerge
            SELECT RDB$FIELD_NAME AS FieldName
            FROM RDB$RELATION_FIELDS
            WHERE RDB$RELATION_NAME = '<<lcLeft+tcTable+lcRight>>' AND RDB$FIELD_NAME = '<<lcLeft+tcField+lcRight>>';
        ENDTEXT

        Return lcQuery
    Endfunc

    Function getServerDateScript
        Return "SELECT CURRENT_TIMESTAMP AS SERTIME FROM RDB$DATABASE"
    Endfunc

    Function getNewGuidScript
        * Firebird no tiene forma de generar un UUID().
    Endfunc

    Function getTablesScript
        Local lcQuery

        TEXT TO lcQuery NOSHOW PRETEXT 7 TEXTMERGE
            SELECT RDB$RELATION_NAME AS TableName
            FROM RDB$RELATIONS
            WHERE RDB$VIEW_BLR IS NULL
        ENDTEXT

        Return lcQuery
    Endfunc

    Function getTableFieldsScript(tcTable)
        Local lcQuery, lcLeft, lcRight
		Store "" to lcLeft, lcRight
		
		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		Else
			tcTable = Upper(tcTable)
		EndIf
		
        TEXT to lcQuery noshow pretext 7 textmerge
            SELECT RDB$FIELD_NAME AS FieldName
            FROM RDB$RELATION_FIELDS
            WHERE RDB$RELATION_NAME = '<<lcLeft+tcTable+lcRight>>'
        ENDTEXT

        Return lcQuery
    Endfunc

    Procedure getPrimaryKeyScript(tcTable)
		Local lcScript
		If !this.bUseSymbolDelimiter
			tcTable = Upper(tcTable)
		EndIf

        TEXT to lcScript noshow pretext 7 textmerge
            SELECT SEG.RDB$FIELD_NAME AS COLUMN_NAME
            FROM RDB$RELATION_CONSTRAINTS CON
            JOIN RDB$INDEX_SEGMENTS SEG ON CON.RDB$INDEX_NAME = SEG.RDB$INDEX_NAME
            WHERE CON.RDB$RELATION_NAME = '<<this.cLeft+tcTable+this.cRight>>' AND CON.RDB$CONSTRAINT_TYPE = 'PRIMARY KEY'
        ENDTEXT

        Return lcScript
    Endproc

    Function createTableOptions
        Return " "
    Endfunc

    Procedure sendConfigurationQuerys
        * Abstract
    Endproc

	function setEnvironment
		Local loEnv
		loEnv = CreateObject("Collection")		
		loEnv.Add(Set("Date"), 'date')
		loEnv.Add(Set("Century"), 'century')
		loEnv.Add(Set("Mark"), 'mark')				

		Set Date To YMD
		Set Century On
		Set Mark To '-'

		Return loEnv
	EndFunc

    Procedure restoreEnvironment(toEnv)
        Local lcDate, lcCentury, lcMark
        lcDate    = toEnv.Item(1)
        lcCentury = toEnv.Item(2)
        lcMark    = toEnv.Item(3)

        Set Date (lcDate)
        Set Century &lcCentury
        Set Mark To (lcMark)
    Endproc

    Function formatDateOrDateTime(tdValue)
        If Empty(tdValue)
            If Type('tdValue') == 'D'
                Return Date(1753, 01, 01)
            Endif
            Return Datetime(1753,01,01,00,00,00)
        Endif

        Return tdValue
    Endfunc

    Function getOpenTableScript(tcTable)
    	Local lcLeft, lcRight
		Store "" to lcLeft, lcRight

		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		Else
			tcTable = Upper(tcTable)
		EndIf
        Return "SELECT * FROM " + lcLeft + tcTable + lcRight
    Endfunc

    Function addAutoIncrement()
        Return "GENERATED BY DEFAULT AS IDENTITY"
    Endfunc

    Function addPrimaryKey()
        Return "PRIMARY KEY"
    Endfunc

	function dropTable(tcTable)
		Local lcScript
		If !this.bUseSymbolDelimiter
			tcTable = Upper(tcTable)
		EndIf
		Text to lcScript noshow pretext 7 textmerge
			DROP TABLE <<this.cLeft+tcTable+this.cRight>>;
		endtext
		Return lcScript
	EndFunc

	Function getCreateDatabaseScript(tcDatabase)
		* Firebird
	EndFunc
	
	Function getDataBaseExistsScript(tcDatabase)
		Local lcScript, lcLeft, lcRight
		Store "" to lcLeft, lcRight

		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		Else
			tcDatabase = Upper(tcDatabase)
		EndIf
		Text to lcScript noshow pretext 7 textmerge
			SELECT 1 FROM rdb$database WHERE LOWER(rdb$database_name) = LOWER('<<lcLeft+tcDatabase+lcRight>>')
		EndText
		Return lcScript
	EndFunc

	Function addFieldComment(tcComment)
		Return " "
	EndFunc
	
	Function addTableComment(tcComment)
		Return " "
	EndFunc	
	
    Function addForeignKey(toFkData)
        Local lcScript, lcOnUpdate, lcOnDelete, lcOPen, lcClose
        lcOnUpdate = This.getForeignKeyValue(toFkData.cOnUpdate)
        lcOnDelete = This.getForeignKeyValue(toFkData.cOnDelete)
		lcOPen = this.cLeft
		lcClose = this.cRight
		
		If !this.bUseSymbolDelimiter
			toFkData.cCurrentField = Upper(toFkData.cCurrentField)
			toFkData.cField = Upper(toFkData.cField)
			toFkData.cTable = Upper(toFkData.cTable)
		EndIf
		        
        TEXT to lcScript noshow pretext 7 textmerge
        	ALTER TABLE <<lcOPen+toFkData.cCurrentTable+lcClose>> ADD CONSTRAINT <<lcOPen+toFkData.cName+lcClose>> 
            FOREIGN KEY (<<lcOPen+toFkData.cCurrentField+lcClose>>)
            REFERENCES <<lcOPen+toFkData.cTable+lcClose>>(<<lcOPen+toFkData.cField+lcClose>>)
            ON UPDATE <<lcOnUpdate>>
            ON DELETE <<lcOnDelete>>
        endtext

        Return lcScript
    Endfunc

	Function addSingleIndex(toIndex)
		Local lcScript, lcLeft, lcRight
		Store "" to lcLeft, lcRight

		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		EndIf
		Text to lcScript noshow pretext 7 textmerge
			CREATE <<Iif(toIndex.bUnique, "UNIQUE", "")>> <<toIndex.cSort>> INDEX <<lcLeft+toIndex.cName+lcRight>> ON <<lcLeft+toIndex.cTable+lcRight>>(<<lcLeft+toIndex.cField+lcRight>>);
		EndText

		Return lcScript
	EndFunc

	Function addComposedIndex(toIndex)
		Local lcScript, lcLeft, lcRight, lcColumns, loResult
		Store "" to lcLeft, lcRight, lcColumns

		lcLeft = this.cLeft
		lcRight = this.cRight
		loResult = CreateObject("Collection")

		For each loComposed in toIndex			
			For each loColumn in loComposed.oColumns
				lcColumns = ''
				For each loField in loColumn
					If !Empty(lcColumns)
						lcColumns = lcColumns + ','
					EndIf
					lcColumns = lcColumns + ' ' + lcLeft + loField.cName + lcRight
				EndFor

				Text to lcScript noshow pretext 7 textmerge
					CREATE <<Iif(loComposed.bUnique, "UNIQUE", "")>> <<loComposed.cSort>> INDEX <<lcLeft+loComposed.cName+lcRight>> ON <<lcLeft+loComposed.cTable+lcRight>>(<<lcColumns>>)
				EndText
				loResult.Add(lcScript)
			EndFor
		EndFor
		Return loResult
	EndFunc
	
	Function addFieldOptions(toField)
		Local lcScript
		lcScript = ''
		If toField.autoIncrement
			lcScript = lcScript + ' ' + this.addAutoIncrement()
		EndIf

		If toField.addDefault AND !toField.autoIncrement
			lcScript = lcScript + " DEFAULT " + toField.Default
		EndIf
		
		If !toField.allowNull
			lcScript = lcScript + " NOT NULL "				
		EndIf
		
		If toField.primaryKey
			lcScript = lcScript + ' ' + this.addPrimaryKey()
		EndIf

		Return lcScript
	EndFunc

	function changeDB(tcNewDatabase)
		If Empty(tcNewDatabase)
		    Text to this.cLastError noshow pretext 7 textmerge
		        ERROR - Base de Datos No Especificada:

		        No has especificado el nombre de la base de datos que deseas utilizar. Por favor, asegúrate de proporcionar el nombre de la base de datos y vuelve a intentarlo.
		    EndText
		    If this.bShowErrors
			    MessageBox(this.cLastError, 16)
			EndIf
		    Return .f.
		Endif

		If !File(tcNewDatabase)
		    Text to this.cLastError noshow pretext 7 textmerge
		        ERROR - Base de Datos Inexistente:

		        La base de datos con el nombre '<<tcNewDatabase>>' no existe o no se puede encontrar en el sistema de archivos local. Por favor, verifica que el nombre de la base de datos sea correcto y que la base de datos haya sido creada previamente.
		    EndText
		    If this.bShowErrors
			    MessageBox(this.cLastError, 16)
			EndIf
		    Return
		Endif

		This.cDatabase = tcNewDatabase
		this.disconnect()
		Return this.connect(.t.)
	EndFunc

	Function changeTypeOnAutoIncrement(tcType)
		Return tcType
	EndFunc

	* C = Character
    Function visitCType(toFields)
        Return "CHAR(" + toFields.Size + ")"
    Endfunc

    Function visitYType(toFields)
    	if toFields.Default == "''"
	    	toFields.Default = "0.0"
	    endif
        Return "DECIMAL(" + toFields.Size + "," + toFields.Decimal + ")"
    Endfunc

    Function visitDType(toFields)
    	if toFields.Default == "''"
		    toFields.Default = "'1858-11-18'"
		endif
        Return "DATE"
    Endfunc

    Function visitTType(toFields)
    	if toFields.Default == "''"
		    toFields.Default = "'1858-11-18 00:00:00'"
		endif
        Return "TIMESTAMP"
    Endfunc

    Function visitBType(toFields)
    	if toFields.Default == "''"
		    toFields.Default = "0.0"
		endif
	    Return "DECIMAL(" + toFields.Size + "," + toFields.Decimal + ")"
    Endfunc

    Function visitFType(toFields)
    	if toFields.Default == "''"
		    toFields.Default = "0.0"
		endif
	    Return "DECIMAL(" + toFields.Size + "," + toFields.Decimal + ")"
    Endfunc

    Function visitGType(toFields)
        Return "BLOB SUB_TYPE 0"
    Endfunc

    Function visitIType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0"
		endif
        Return "INTEGER"
    Endfunc

    Function visitLType(toFields)
    	if toFields.Default == "''"
	    	toFields.Default = "FALSE"
	    endif
        Return "BOOLEAN"
    Endfunc

    Function visitMType(toFields)
        Return "BLOB SUB_TYPE 1"
    Endfunc

    Function visitNType(toFields)
        If Val(toFields.Decimal) > 0
            Return "NUMERIC(" + toFields.Size + "," + toFields.Decimal + ")"
        Else
            Return "INTEGER"
        Endif
    Endfunc

    Function visitQType(toFields)
        Return "BLOB SUB_TYPE 0"
    Endfunc

    Function visitVType(toFields)
        Return "VARCHAR(" + Iif(Empty(Val(toFields.Size)), '8191', toFields.Size) + ")"
    Endfunc

    Function visitWType(toFields)
        Return "BLOB SUB_TYPE 0"
    EndFunc	    
EndDefine

* ==================================================== *
* SQLite
* ==================================================== *
Define Class SQLite As DBEngine

	Procedure init
		DoDefault()
		this.nMaxLength = 128
		this.cLeft = '"'
		this.cRight = '"'
		this.bExecuteIndexScriptSeparately = .t.
	EndProc

	Function getDummyQuery
		Return "SELECT 1"
	Endfunc

	Function getVersion
		Local lcCursor, lcVersion
		lcCursor = Sys(2015)
		This.SQLExec("SELECT sqlite_version() AS 'VER'", lcCursor)
		lcVersion = &lcCursor..VER
		Use In (lcCursor)

		Return lcVersion
	Endfunc

	Function getConnectionString(tbAddDatabase)
		Local lcConStr

		Text to lcConStr noshow pretext 15 textmerge
			DRIVER=<<this.cDriver>>;
			DATABASE=<<This.cDatabase>>;
			UID=<<Iif(Empty(this.cUser),'', this.cUser)>>;
			PWD=<<Iif(Empty(this.cPassword),'', this.cPassword)>>;
			LongNames=0;
			TimeOut=1000;
			NoTXN=0;
			SyncPragma=NORMAL;
			Client=<<this.name>> by FoxRemote;
			StepAPI=0
		endtext

		Return lcConStr
	Endfunc

	Procedure beginTransaction
*!*			This.SQLExec("COMMIT;")
*!*			* Desactivar las transacciones automáticas
*!*			this.SQLExec("PRAGMA autocommit = 0;")
*!*			This.SQLExec("BEGIN;")
	Endproc

	Procedure endTransaction
*!*			This.SQLExec("COMMIT;")
*!*			* Activar las transacciones automáticas
*!*			this.SQLExec("PRAGMA autocommit = 1;")
	Endproc

	Procedure cancelTransaction
*!*			This.SQLExec("ROLLBACK;")
*!*			* Activar las transacciones automáticas
*!*			this.SQLExec("PRAGMA autocommit = 1;")
	Endproc

	Function getTableExistsScript(tcTableName)
		Local lcQuery, lcLeft, lcRight
		Store "" to lcLeft, lcRight

		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		EndIf

		TEXT to lcQuery noshow pretext 7 textmerge
			SELECT name as TableName FROM sqlite_master WHERE type='table' AND name = '<<lcLeft+tcTableName+lcRight>>';
		ENDTEXT

		Return lcQuery
	Endfunc

	Procedure selectDatabase
		* No aplica para SQLite
	Endproc

	Function getTidScript
		Return "INTEGER PRIMARY KEY AUTOINCREMENT"
	Endfunc

	Function getUUIDScript
		Return "VARCHAR(36) PRIMARY KEY NOT NULL"
	EndFunc

	Function getFieldExistsScript(tcTable, tcField)
		Local lcQuery, lcLeft, lcRight
		Store "" to lcLeft, lcRight

		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		EndIf

		TEXT to lcQuery noshow pretext 7 textmerge
			SELECT name AS FieldName
			FROM pragma_table_info('<<lcLeft+tcTable+lcRight>>')
			WHERE name = '<<lcLeft+tcField+lcRight>>';
		ENDTEXT

		Return lcQuery
	Endfunc

	Function getServerDateScript
		Return "SELECT CURRENT_TIMESTAMP AS SERTIME;"
	Endfunc

	Function getNewGuidScript
		Return "SELECT LOWER(HEX(RANDOMBLOB(16))) AS GUID;"
	Endfunc

	Function getTablesScript
		Local lcQuery

		TEXT TO lcQuery NOSHOW PRETEXT 7 TEXTMERGE
			SELECT name AS TableName
			FROM sqlite_master
			WHERE type='table';
		ENDTEXT

		Return lcQuery
	Endfunc

	Function getTableFieldsScript(tcTable)
		Local lcQuery, lcLeft, lcRight
		Store "" to lcLeft, lcRight

		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		EndIf

		TEXT to lcQuery noshow pretext 7 textmerge
			SELECT name AS FieldName
			FROM pragma_table_info('<<lcLeft+tcTable+lcRight>>');
		ENDTEXT

		Return lcQuery
	Endfunc

	Procedure getPrimaryKeyScript(tcTable)
		local lcScript
		TEXT to lcScript noshow pretext 7 textmerge
		SELECT name AS COLUMN_NAME FROM pragma_table_info('<<this.cLeft+tcTable+this.cRight>>') WHERE pk = 1
		endtext
		return lcScript
*!*			Local lcScript

*!*			TEXT to lcScript noshow pretext 7 textmerge
*!*				SELECT sql as column_name
*!*				FROM sqlite_master
*!*				WHERE type='table' AND name='<<this.cLeft+tcTable+this.cRight>>';
*!*			ENDTEXT

*!*			* Extracting the primary key from the table creation script
*!*			Local lnStart, lnEnd
*!*			lnStart = AT('CONSTRAINT', lcScript) + 11
*!*			lnEnd = AT('PRIMARY KEY', lcScript) - 3
*!*			lcScript = SUBSTR(lcScript, lnStart, lnEnd - lnStart)

*!*			Return lcScript
	Endproc

	Function createTableOptions
		Return ""
	Endfunc

	function setEnvironment
		Local loEnv
		loEnv = CreateObject("Collection")		
		loEnv.Add(Set("Date"), 'date')
		loEnv.Add(Set("Century"), 'century')
		loEnv.Add(Set("Mark"), 'mark')				

		Set Date To YMD
		Set Century On
		Set Mark To '-'

		Return loEnv
	EndFunc

	Procedure restoreEnvironment(toEnv)
		Local lcDate, lcCentury, lcMark
		lcDate    = toEnv.Item(1)
		lcCentury = toEnv.Item(2)
		lcMark    = toEnv.Item(3)
		Set Date (lcDate)
		Set Century &lcCentury
		Set Mark To (lcMark)
	Endproc

	Function formatDateOrDateTime(tdValue)
		If Empty(tdValue)
			If Type('tdValue') == 'D'
				Return CTOD('0001-01-01')
			Endif
			Return CTOT('0001-01-01 00:00:00')
		Endif
		Return tdValue
	Endfunc

	Function getOpenTableScript(tcTable)
		Return "SELECT * FROM " + tcTable
	Endfunc

	Function addAutoIncrement
		Return "AUTOINCREMENT"
	Endfunc

	Function addPrimaryKey
		Return "PRIMARY KEY"
	Endfunc

	Function dropTable(tcTable)
		Return "DROP TABLE IF EXISTS " + this.cLeft + tcTable + this.cRight + ';'
	Endfunc

	Function getCreateDatabaseScript(tcDatabase)
		Return ""
	Endfunc

	Function getDataBaseExistsScript(tcDatabase)
		Return ""
	Endfunc

	Function addFieldComment(tcComment)
		Return ""
	Endfunc

	Function addTableComment(tcComment)
		Return ""
	Endfunc

    Function addForeignKey(toFkData)
        Local lcScript, lcOnUpdate, lcOnDelete, lcOPen, lcClose
        lcOnUpdate = This.getForeignKeyValue(toFkData.cOnUpdate)
        lcOnDelete = This.getForeignKeyValue(toFkData.cOnDelete)
		Store "" to lcOPen, lcClose
		
		lcOPen  = this.cLeft
		lcClose = this.cRight
		        
        TEXT to lcScript noshow pretext 7 textmerge        	
            FOREIGN KEY (<<lcOPen+toFkData.cCurrentField+lcClose>>)
            REFERENCES <<lcOPen+toFkData.cTable+lcClose>>(<<lcOPen+toFkData.cField+lcClose>>)
            ON UPDATE <<lcOnUpdate>>
            ON DELETE <<lcOnDelete>>
        endtext

        Return lcScript
    Endfunc

	Function addSingleIndex(toIndex)
		Local lcScript, lcLeft, lcRight
		Store "" to lcLeft, lcRight

		If this.bUseSymbolDelimiter
			lcLeft = this.cLeft
			lcRight = this.cRight
		EndIf
		Text to lcScript noshow pretext 7 textmerge
			CREATE <<Iif(toIndex.bUnique, "UNIQUE", "")>> INDEX <<lcLeft+toIndex.cName+lcRight>> ON <<lcLeft+toIndex.cTable+lcRight>>(<<lcLeft+toIndex.cField+lcRight>>);
		EndText

		Return lcScript
	EndFunc

	Function addComposedIndex(toIndex)
		Local lcScript, lcLeft, lcRight, lcColumns, loResult
		Store "" to lcLeft, lcRight, lcColumns

		lcLeft = this.cLeft
		lcRight = this.cRight
		loResult = CreateObject("Collection")

		For each loComposed in toIndex			
			For each loColumn in loComposed.oColumns
				lcColumns = ''
				For each loField in loColumn
					If !Empty(lcColumns)
						lcColumns = lcColumns + ','
					EndIf
					lcColumns = lcColumns + ' ' + lcLeft + loField.cName + lcRight
				EndFor

				Text to lcScript noshow pretext 7 textmerge
					CREATE <<Iif(loComposed.bUnique, "UNIQUE", "")>> INDEX <<lcLeft+loComposed.cName+lcRight>> ON <<lcLeft+loComposed.cTable+lcRight>>(<<lcColumns>>)
				EndText
				loResult.Add(lcScript)
			EndFor
		EndFor
		Return loResult
	EndFunc

	Function addFieldOptions(toField)
		Local lcScript
		lcScript = ''
		
		If toField.primaryKey
			lcScript = lcScript + ' ' + this.addPrimaryKey()
		EndIf									

		If toField.autoIncrement
			lcScript = lcScript + ' ' + this.addAutoIncrement()
		EndIf

		If toField.addDefault AND !toField.autoIncrement
			lcScript = lcScript + " DEFAULT " + toField.Default
		endif

		If !toField.allowNull
			lcScript = lcScript + " NOT NULL "				
		EndIf

		Return lcScript
	EndFunc

	function changeDB(tcNewDatabase)
		If Empty(tcNewDatabase)
		    Text to this.cLastError noshow pretext 7 textmerge
		        ERROR - Base de Datos No Especificada:

		        No has especificado el nombre de la base de datos que deseas utilizar. Por favor, asegúrate de proporcionar el nombre de la base de datos y vuelve a intentarlo.
		    EndText
		    If this.bShowErrors
			    MessageBox(this.cLastError, 16)
			EndIf
		    Return .f.
		Endif

		If !File(tcNewDatabase)
		    Text to this.cLastError noshow pretext 7 textmerge
		        ERROR - Base de Datos Inexistente:

		        La base de datos con el nombre '<<tcNewDatabase>>' no existe o no se puede encontrar en el sistema de archivos local. Por favor, verifica que el nombre de la base de datos sea correcto y que la base de datos haya sido creada previamente.
		    EndText
		    If this.bShowErrors
			    MessageBox(this.cLastError, 16)
			EndIf
		    Return
		Endif

		This.cDatabase = tcNewDatabase
		this.disconnect()
		Return this.connect(.t.)
	EndFunc

	Function changeTypeOnAutoIncrement(tcType)
		Return tcType
	EndFunc

	* C = Character
	Function visitCType(toFields)
		If Val(toFields.Size) > 0
			Return "TEXT(" + toFields.Size + ")"
		EndIf
		Return "TEXT"
	Endfunc

	* Y = Currency
	Function visitYType(toFields)
		If Val(toFields.Decimal) > 0
			if toFields.Default == "''"
				toFields.Default = "0.0"
			endif
		Else
			if toFields.Default == "''"
				toFields.Default = '0'
			endif
		Endif
		Return "NUMERIC(" + toFields.Size + "," + toFields.Decimal + ")"
	Endfunc

	* D = Date
	Function visitDType(toFields)
		if toFields.Default == "''"
			toFields.Default = "'0001-01-01'"
		endif
		Return "DATE"
	Endfunc

	* T = DateTime
	Function visitTType(toFields)
		if toFields.Default == "''"
			toFields.Default = "'0001-01-01 00:00:00.000'"
		endif
		Return "DATETIME"
	Endfunc

	* B = Double
	Function visitBType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0.0"
		endif
		Return "DOUBLE"
	Endfunc

	* F = Float
	Function visitFType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0.0"
		endif
		Return "REAL"
	Endfunc

	* G = General
	Function visitGType(toFields)
		Return "BLOB"
	Endfunc

	* I = Integer
	Function visitIType(toFields)
		Return "INTEGER"
	Endfunc

	* L = Logical
	Function visitLType(toFields)
		Return "BOOLEAN"
	Endfunc

	* M = Memo
	Function visitMType(toFields)
		Return "TEXT"
	Endfunc

	* N = Numeric
	Function visitNType(toFields)
		If Val(toFields.Decimal) > 0
			if toFields.Default == "''"
				toFields.Default = "0.0"
			endif
		Else
			if toFields.Default == "''"
				toFields.Default = '0'
			endif
		Endif
		Return "NUMERIC(" + toFields.Size + "," + toFields.Decimal + ")"
	Endfunc

	* Q = Varbinary
	Function visitQType(toFields)
		Return "BLOB"
	Endfunc

	* V = Varchar and Varchar (Binary)
	Function visitVType(toFields)
		If Val(toFields.Size) > 0
			Return "TEXT(" + toFields.Size + ")"
		EndIf
		Return "TEXT"
	Endfunc

	* W = Blob
	Function visitWType(toFields)
		Return "BLOB"
	EndFunc
EndDefine

* ==================================================== *
* PostgreSQL
* ==================================================== *
Define Class PostgreSQL As DBEngine

    Procedure init
		DoDefault()    
        this.nMaxLength = 63
        this.cLeft = '"'
        this.cRight = '"'
        this.bExecuteFkScriptSeparately = .t.
        this.bExecuteIndexScriptSeparately = .T.
    EndProc

    Function getDummyQuery
        Return "SELECT 1"
    Endfunc

    Function getVersion
        Local lcCursor, lcVersion
        lcCursor = Sys(2015)
        This.SQLExec("SELECT version()", lcCursor)
        lcVersion = &lcCursor..version
        Use In (lcCursor)

        Return lcVersion
    Endfunc

    Function getConnectionString(tbAddDatabase)
		Local lcConStr
		Text to lcConStr noshow pretext 15 textmerge
			DRIVER={<<this.cDriver>>};
			SERVER=<<this.cServer>>;
			PORT=<<Iif(Empty(this.nPort), "5432", this.nPort)>>;
			UID=<<Iif(Empty(this.cUser),'', this.cUser)>>;
			PWD=<<Iif(Empty(this.cPassword),'', this.cPassword)>>;
			APP=<<this.name>> by FoxRemote
		endtext
		If tbAddDatabase
			lcConStr = lcConStr + "DATABASE=" + Alltrim(This.cDatabase)
		Endif
		Return lcConStr
    Endfunc

    Procedure beginTransaction
        This.selectDatabase()
        This.SQLExec("BEGIN;")
    Endproc

    Procedure endTransaction
        This.selectDatabase()
        This.SQLExec("COMMIT;")
    Endproc

    Procedure cancelTransaction
        This.selectDatabase()
        This.SQLExec("ROLLBACK;")
    Endproc

    Function getTableExistsScript(tcTableName)
        Local lcQuery, lcLeft, lcRight
        Store "" to lcLeft, lcRight
        
        If this.bUseSymbolDelimiter
            lcLeft = this.cLeft
            lcRight = this.cRight
        EndIf
        
        TEXT to lcQuery noshow pretext 7 textmerge
            SELECT table_name as TableName
            FROM information_schema.tables
            WHERE table_name = '<<lcLeft+tcTableName+lcRight>>' AND table_schema = 'public'
        ENDTEXT

        Return lcQuery
    Endfunc

	function dropTable(tcTable)			
		Return "DROP TABLE IF EXISTS " + this.cLeft + tcTable + this.cRight + ';'
	endfunc

    Procedure selectDatabase
        Return ""
    Endproc

    Function getTidScript
        Return "SERIAL PRIMARY KEY"
    Endfunc

	Function getUUIDScript
		Return "UUID DEFAULT uuid_generate_v4() PRIMARY KEY NOT NULL"
	EndFunc

    Function getFieldExistsScript(tcTable, tcField)
        Local lcQuery, lcLeft, lcRight
        Store "" to lcLeft, lcRight
        
        If this.bUseSymbolDelimiter
            lcLeft = this.cLeft
            lcRight = this.cRight
        Else
            tcTable = Upper(tcTable)
            tcField = Upper(tcField)
        EndIf
        
        TEXT to lcQuery noshow pretext 7 textmerge
            SELECT column_name AS FieldName
            FROM information_schema.columns
            WHERE table_name = '<<lcLeft+tcTable+lcRight>>' AND column_name = '<<lcLeft+tcField+lcRight>>'
        ENDTEXT

        Return lcQuery
    Endfunc

    Function getServerDateScript
        Return "SELECT CURRENT_TIMESTAMP AS SERTIME"
    Endfunc

    Function getNewGuidScript
    	If this.sqlExec('CREATE EXTENSION IF NOT EXISTS "uuid-ossp";')
	        Return "SELECT uuid_generate_v4() AS GUID"
	    EndIf
	    Text to this.cLastError noshow pretext 7
	    Hubo un problema con la extensión 'uuid-ossp'.
	    Por favor, asegúrate de que la extensión está instalada y habilitada en tu servidor de base de datos.
	    Si necesitas ayuda, consulta la documentación o contacta al administrador de la base de datos.
	    EndText
	    If this.bShowErrors
		    MessageBox(this.cLastError, 48, "DBCraft")
		EndIf
    EndFunc
    
    Function getTablesScript
        Local lcQuery

        TEXT TO lcQuery NOSHOW PRETEXT 7 TEXTMERGE
            SELECT table_name
            FROM information_schema.tables
            WHERE table_schema = 'public'
        ENDTEXT

        Return lcQuery
    Endfunc

    Function getTableFieldsScript(tcTable)
        Local lcQuery, lcLeft, lcRight
        Store "" to lcLeft, lcRight
        
        If this.bUseSymbolDelimiter
            lcLeft = this.cLeft
            lcRight = this.cRight
        Else
            tcTable = Upper(tcTable)
        EndIf
        
        TEXT to lcQuery noshow pretext 7 textmerge
            SELECT column_name
            FROM information_schema.columns
            WHERE table_name = '<<lcLeft+tcTable+lcRight>>'
        ENDTEXT

        Return lcQuery
    Endfunc

    Procedure getPrimaryKeyScript(tcTable)
        Local lcScript

        If !this.bUseSymbolDelimiter
            tcTable = Upper(tcTable)
        EndIf

        TEXT to lcScript noshow pretext 7 textmerge
            SELECT a.attname as COLUMN_NAME
            FROM pg_index i
            JOIN pg_attribute a ON a.attrelid = i.indrelid
                AND a.attnum = ANY(i.indkey)
            WHERE i.indrelid = '<<this.cLeft+tcTable+this.cRight>>'::regclass
                AND i.indisprimary;
        ENDTEXT

        Return lcScript
    Endproc

    Function createTableOptions
        Return " "
    Endfunc

    Procedure sendConfigurationQuerys
        * Abstract
    Endproc

	function setEnvironment
		Local loEnv
		loEnv = CreateObject("Collection")		
		loEnv.Add(Set("Date"), 'date')
		loEnv.Add(Set("Century"), 'century')
		loEnv.Add(Set("Mark"), 'mark')				
		
		Set Date To YMD
		Set Century On
		Set Mark To '-'

    	If !this.sqlExec('CREATE EXTENSION IF NOT EXISTS "uuid-ossp";')
    		this.cLastError = "Hubo un problema con la extensión 'uuid-ossp'. Por favor, asegúrate de que la extensión está instalada y habilitada en tu servidor de base de datos. Si necesitas ayuda, consulta la documentación o contacta al administrador de la base de datos."
    		If this.bShowErrors
		        messagebox(this.cLastError, 48)
		    EndIf
	    EndIf	    

		Return loEnv
	EndFunc

    Procedure restoreEnvironment(toEnv)
        Local lcDate, lcCentury, lcMark
        lcDate    = toEnv.Item(1)
        lcCentury = toEnv.Item(2)
        lcMark    = toEnv.Item(3)

        Set Date (lcDate)
        Set Century &lcCentury
        Set Mark To (lcMark)
    Endproc

    Function formatDateOrDateTime(tdValue)
        If Empty(tdValue)
            If Type('tdValue') == 'D'
                Return Date(1753, 01, 01)
            Endif
            Return Datetime(1753,01,01,00,00,00)
        Endif

        Return tdValue
    Endfunc

    Function getOpenTableScript(tcTable)
        Local lcLeft, lcRight
        Store "" to lcLeft, lcRight
        
        If this.bUseSymbolDelimiter
            lcLeft = this.cLeft
            lcRight = this.cRight
        Else
            tcTable = Upper(tcTable)
        EndIf    
        Return "SELECT * FROM " + lcLeft + tcTable + lcRight
    Endfunc

	Function getDataBaseExistsScript(tcDatabase)
		Local lcScript
		Text to lcScript noshow pretext 7 textmerge
			SELECT datname as dbName FROM pg_database WHERE datname = '<<tcDatabase>>';
		EndText
		Return lcScript
	Endfunc

    Function addAutoIncrement()
        Return ""
    Endfunc

    Function addPrimaryKey()
        Return "PRIMARY KEY"
    Endfunc

	Function addFieldComment(tcComment)
		Return ""
	Endfunc

	Function addTableComment(tcComment)
		Return ""
	Endfunc

    Function addForeignKey(toFkData)
        Local lcScript, lcOnUpdate, lcOnDelete, lcLeft, lcRight, lcFkName
        lcOnUpdate = This.getForeignKeyValue(toFkData.cOnUpdate)
        lcOnDelete = This.getForeignKeyValue(toFkData.cOnDelete)
        Store "" to lcLeft, lcRight        
        TEXT to lcScript noshow pretext 7 textmerge
            ALTER TABLE <<lcLeft+toFkData.cCurrentTable+lcRight>> ADD CONSTRAINT <<lcLeft+toFkData.cName+lcRight>>
            FOREIGN KEY (<<lcLeft+toFkData.cCurrentField+lcRight>>)
            REFERENCES <<lcLeft+toFkData.cTable+lcRight>>(<<lcLeft+toFkData.cField+lcRight>>)
            ON UPDATE <<lcOnUpdate>>
            ON DELETE <<lcOnDelete>>
        endtext

        Return lcScript
    Endfunc

	Function getCreateDatabaseScript(tcDatabase)
		Return "CREATE DATABASE " + this.cLeft + tcDatabase + this.cRight + ";"
	EndFunc

    Function addSingleIndex(toIndex)
        Local lcScript, lcLeft, lcRight
        Store "" to lcLeft, lcRight
        
        If this.bUseSymbolDelimiter
            lcLeft = this.cLeft
            lcRight = this.cRight
        EndIf
        Text to lcScript noshow pretext 7 textmerge
            CREATE <<Iif(toIndex.bUnique, "UNIQUE", "")>> INDEX <<lcLeft+toIndex.cName+lcRight>> ON <<lcLeft+toIndex.cTable+lcRight>>(<<lcLeft+toIndex.cField+lcRight>> <<toIndex.cSort>>)
        EndText

        Return lcScript
    Endfunc

    Function addComposedIndex(toIndex)
        Local lcScript, lcLeft, lcRight, lcColumns, loResult
        Store "" to lcLeft, lcRight, lcColumns

        lcLeft = this.cLeft
        lcRight = this.cRight
        loResult = CreateObject("Collection")

        For each loComposed in toIndex            
            For each loColumn in loComposed.oColumns
                lcColumns = ''
                For each loField in loColumn
                    If !Empty(lcColumns)
                        lcColumns = lcColumns + ','
                    EndIf
                    lcColumns = lcColumns + ' ' + lcLeft + loField.cName + lcRight + ' ' + loField.cSort
                EndFor

                Text to lcScript noshow pretext 7 textmerge
                    CREATE <<Iif(loComposed.bUnique, "UNIQUE", "")>> INDEX <<lcLeft+loComposed.cName+lcRight>> ON <<lcLeft+loComposed.cTable+lcRight>>(<<lcColumns>>)
                EndText
                loResult.Add(lcScript)
            EndFor
        EndFor
        Return loResult
    Endfunc

    Function addFieldOptions(toField)
        Local lcScript
        lcScript = ''
        If toField.autoIncrement
            lcScript = lcScript + ' ' + this.addAutoIncrement()
        EndIf

        If toField.addDefault AND !toField.autoIncrement
            lcScript = lcScript + " DEFAULT " + toField.Default
        EndIf
        
        If !toField.allowNull
            lcScript = lcScript + " NOT NULL "                
        EndIf
        
        If toField.primaryKey
            lcScript = lcScript + ' ' + this.addPrimaryKey()
        EndIf

        Return lcScript
    Endfunc

	function changeDB(tcNewDatabase)
		If Empty(tcNewDatabase)
		    Text to this.cLastError noshow pretext 7 textmerge
		        ERROR - Base de Datos No Especificada:

		        No has especificado el nombre de la base de datos que deseas utilizar. Por favor, asegúrate de proporcionar el nombre de la base de datos y vuelve a intentarlo.
		    EndText
		    If this.bShowErrors
			    MessageBox(this.cLastError, 16)
			EndIf
		    Return .f.
		Endif

		This.cDatabase = tcNewDatabase
		this.disconnect()
		Return this.connect(.t.)
	EndFunc

	Function changeTypeOnAutoIncrement(tcType)
		Return "SERIAL"
	EndFunc

    * C = Character
    Function visitCType(toFields)
    	If Val(toFields.Size) > 0
    		Return "VARCHAR(" + toFields.Size + ")"
    	EndIf
    	Return "VARCHAR"
    Endfunc

    Function visitYType(toFields)
    	if toFields.Default == "''"
		    toFields.Default = "0.0"
		endif
        Return "NUMERIC(" + toFields.Size + "," + toFields.Decimal + ")"
    Endfunc

    Function visitDType(toFields)
    	if toFields.Default == "''"
	        toFields.Default = "'1858-11-18'"
	    endif
        Return "DATE"
    Endfunc

    Function visitTType(toFields)
    	if toFields.Default == "''"
	        toFields.Default = "'1858-11-18 00:00:00'"
	    endif
        Return "TIMESTAMP"
    Endfunc

    Function visitBType(toFields)
    	if toFields.Default == "''"
	        toFields.Default = "0.0"
	    endif
        Return "NUMERIC(" + toFields.Size + "," + toFields.Decimal + ")"
    Endfunc

    Function visitFType(toFields)
    	if toFields.Default == "''"
	        toFields.Default = "0.0"
	    endif
        Return "NUMERIC(" + toFields.Size + "," + toFields.Decimal + ")"
    Endfunc

    Function visitGType(toFields)
        Return "BYTEA"
    Endfunc

    Function visitIType(toFields)
		if toFields.Default == "''"
			toFields.Default = "0"
		endif
        Return "INTEGER"
    Endfunc

    Function visitLType(toFields)
    	if toFields.Default == "''"
	        toFields.Default = "FALSE"
	    endif
        Return "BOOLEAN"
    Endfunc

    Function visitMType(toFields)
        Return "TEXT"
    Endfunc

    Function visitNType(toFields)
    	if toFields.Default == "''"
		    toFields.Default = "0"
		endif
        If Val(toFields.Decimal) > 0
        	if toFields.Default == "''"
	        	toFields.Default = "0.0"
	        endif
            Return "NUMERIC(" + toFields.Size + "," + toFields.Decimal + ")"
        Else
            Return "INTEGER"
        Endif
    Endfunc

    Function visitQType(toFields)
        Return "BYTEA"
    Endfunc

    Function visitVType(toFields)
    	If Val(toFields.Size) > 0    	
	        Return "VARCHAR(" + toFields.Size + ")"
       	EndIf
       	Return "VARCHAR"
    Endfunc

    Function visitWType(toFields)
        Return "BYTEA"
    EndFunc
EndDefine

* ======================================================================== *
* Class RemoteCursor
* ======================================================================== *
Define Class RemoteCursor As CursorAdapter
	Database = ""
	Alias = ""
	SelectCmd = ""
	Tables = ""
	KeyFieldList = ""
	SendUpdates = .f.
	Nodata = .f.
EndDefine