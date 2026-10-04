* LegacyShim.prg - the FoxRemote 1.0 API mapped onto 0.x, for the red run of the tests.
*
* The integration tests are written against the 1.0 API. To see them fail against 0.x for the
* fault each one measures (and not merely for a missing method), this shim wraps a 0.x object
* (MSSQL, SQLite, ...) and translates every 1.0 call into the 0.x one that does the same job.
* Where 0.x has no equivalent (error codes, affected rows, named parameters, a migration
* report) nothing is invented: the test that needs it fails, which is the point.
* bShowErrors is switched off so that no MESSAGEBOX waits for a human.
*
* Use: FOXREMOTE_LIB=<home>\tests\legacy\LegacyShim.prg
RETURN

DEFINE CLASS LegacyRemote AS Custom
	cLegacyClass = ""
	oImpl = .NULL.
	cDriver = ""
	cServer = ""
	nPort = 0
	cDatabase = ""
	cUser = ""
	cPassword = ""
	lTrustedConnection = .F.
	lStrict = .F.
	nErrors = 0
	cLastError = ""
	cLastErrorCode = ""
	cLastSql = ""
	nHandle = 0
	lInTransaction = .F.

	PROCEDURE Init
		LOCAL lcLegacy
		lcLegacy = FrHome() + "legacy\foxremote.prg"
		IF !FILE(lcLegacy)
			lcLegacy = FrHome() + "foxremote.prg"
		ENDIF
		IF ATC("\FOXREMOTE.", UPPER(SET("PROCEDURE"))) = 0
			SET PROCEDURE TO (lcLegacy) ADDITIVE
		ENDIF
	ENDPROC

	FUNCTION cLastError_Access
		RETURN IIF(VARTYPE(This.oImpl) == "O", This.oImpl.cLastError, This.cLastError)
	ENDFUNC

	PROTECTED PROCEDURE MakeImpl
		This.oImpl = CREATEOBJECT(This.cLegacyClass)
		This.oImpl.bShowErrors = .F.
		This.oImpl.cDriver = This.cDriver
		This.oImpl.cServer = This.cServer
		This.oImpl.nPort = This.nPort
		This.oImpl.cDatabase = This.cDatabase
		This.oImpl.cUser = This.cUser
		This.oImpl.cPassword = This.cPassword
	ENDPROC

	PROTECTED FUNCTION NewestHandle
		LOCAL laH[1], lnN, lnI, lnMax
		lnMax = 0
		lnN = ASQLHANDLES(laH)
		FOR lnI = 1 TO lnN
			lnMax = MAX(lnMax, laH[lnI])
		ENDFOR
		RETURN lnMax
	ENDFUNC

	* ---- connection ----------------------------------------------------------------
	FUNCTION Connect
		LOCAL lvOk
		This.MakeImpl()
		lvOk = This.oImpl.Connect()
		lvOk = (VARTYPE(lvOk) == "L" AND lvOk)
		This.nHandle = IIF(lvOk, This.NewestHandle(), 0)
		RETURN lvOk
	ENDFUNC

	PROCEDURE Disconnect
		IF VARTYPE(This.oImpl) == "O"
			This.oImpl.disconnect()
		ENDIF
		This.nHandle = 0
	ENDPROC

	FUNCTION CreateDatabase(tcDatabase)
		IF This.nHandle = 0 AND !This.Connect()
			RETURN .F.
		ENDIF
		RETURN This.oImpl.newDataBase(tcDatabase)
	ENDFUNC

	PROCEDURE ClearErrors
		This.nErrors = 0
		IF VARTYPE(This.oImpl) == "O"
			This.oImpl.cLastError = ""
		ENDIF
	ENDPROC

	* ---- queries (0.x has no named parameters and no row count) --------------------
	FUNCTION Query(tcSql, tcCursor, toParams)
		RETURN This.oImpl.SQLExec(tcSql, tcCursor)
	ENDFUNC

	FUNCTION Execute(tcSql, toParams)
		RETURN IIF(This.oImpl.SQLExec(tcSql), 0, -1)
	ENDFUNC

	FUNCTION Scalar(tcSql, toParams)
		LOCAL lvValue
		lvValue = .NULL.
		IF This.oImpl.SQLExec(tcSql, "c_fr_shim_scalar") AND USED("c_fr_shim_scalar")
			SELECT c_fr_shim_scalar
			IF !EOF()
				lvValue = EVALUATE(FIELD(1))
			ENDIF
			USE IN c_fr_shim_scalar
		ENDIF
		RETURN lvValue
	ENDFUNC

	FUNCTION LastId
		RETURN This.oImpl.getLastID()
	ENDFUNC

	* ---- transactions: the 0.x per-engine BEGIN/COMMIT/ROLLBACK in text ------------
	FUNCTION BeginTransaction
		This.oImpl.beginTransaction()
		RETURN .T.
	ENDFUNC

	FUNCTION Commit
		This.oImpl.endTransaction()
		RETURN .T.
	ENDFUNC

	FUNCTION Rollback
		This.oImpl.cancelTransaction()
		RETURN .T.
	ENDFUNC

	* ---- cursors ---------------------------------------------------------------------
	FUNCTION Open(tcTable, tcAlias, tcWhere, toParams, tcFields, tlReadOnly, tcGroup)
		RETURN This.oImpl.use(tcTable + IIF(EMPTY(tcAlias), "", " AS " + tcAlias), ;
			IIF(EMPTY(tcFields), "", tcFields), IIF(EMPTY(tcWhere), "", tcWhere), IIF(EMPTY(tcGroup), "", tcGroup), tlReadOnly)
	ENDFUNC

	FUNCTION Save(tcAlias)
		RETURN This.oImpl.Save(tcAlias)
	ENDFUNC

	FUNCTION SaveGroup(tcGroup)
		RETURN This.oImpl.saveGroup(tcGroup)
	ENDFUNC

	FUNCTION Discard(tcAlias)
		RETURN This.oImpl.discard(tcAlias)
	ENDFUNC

	FUNCTION Refresh(tcAlias)
		LOCAL lvOk
		lvOk = This.oImpl.requery(tcAlias)
		RETURN !(VARTYPE(lvOk) == "L" AND !lvOk)
	ENDFUNC

	FUNCTION Close(tcAlias)
		RETURN This.oImpl.Close(tcAlias)
	ENDFUNC

	PROCEDURE CloseAll
		This.oImpl.closeAll()
	ENDPROC

	FUNCTION CloseGroup(tcGroup)
		RETURN This.oImpl.closeGroup(tcGroup)
	ENDFUNC

	* ---- metadata ----------------------------------------------------------------------
	FUNCTION TableExists(tcTable)
		RETURN This.oImpl.tableExists(tcTable)
	ENDFUNC

	FUNCTION GetPrimaryKey(tcTable)
		RETURN This.oImpl.getPrimaryKey(tcTable)
	ENDFUNC

	FUNCTION ServerVersion
		RETURN This.oImpl.getVersion()
	ENDFUNC

	* 0.x returns an array from a method, which reaches the caller as its first element.
	FUNCTION GetTables(tcCursor)
		LOCAL lvFirst
		lvFirst = This.oImpl.getTables()
		CREATE CURSOR (tcCursor) (name C(128))
		IF VARTYPE(lvFirst) == "C"
			INSERT INTO (tcCursor) (name) VALUES (lvFirst)
		ENDIF
		RETURN .T.
	ENDFUNC

	FUNCTION GetColumns(tcTable, tcCursor)
		LOCAL lvFirst
		lvFirst = This.oImpl.getTableFields(tcTable)
		CREATE CURSOR (tcCursor) (name C(128), type C(20), width I, decimals I, nullable L, primary_key L)
		IF VARTYPE(lvFirst) == "C"
			INSERT INTO (tcCursor) (name) VALUES (lvFirst)
		ENDIF
		RETURN .T.
	ENDFUNC

	* ---- migration (0.x returns a logical, not a report) ---------------------------------
	FUNCTION Migrate(tcSource, tlReplace)
		LOCAL loR, lvOk
		lvOk = This.oImpl.migrate(tcSource)
		loR = CREATEOBJECT("Empty")
		ADDPROPERTY(loR, "nTables", 0)
		ADDPROPERTY(loR, "nRows", 0)
		ADDPROPERTY(loR, "nErrors", IIF(VARTYPE(lvOk) == "L" AND lvOk, 0, 1))
		RETURN loR
	ENDFUNC

	FUNCTION CreateTableFromCursor(tcAlias, tcTable)
		LOCAL laF[1]
		=AFIELDS(laF, tcAlias)
		RETURN This.oImpl.createTable(IIF(EMPTY(tcTable), tcAlias, tcTable), "", .NULL., @laF)
	ENDFUNC
ENDDEFINE

DEFINE CLASS RemoteSqlite AS LegacyRemote
	cLegacyClass = "SQLite"
ENDDEFINE

DEFINE CLASS RemoteSqlServer AS LegacyRemote
	cLegacyClass = "MSSQL"
ENDDEFINE

DEFINE CLASS RemoteMySql AS LegacyRemote
	cLegacyClass = "MySQL"
ENDDEFINE

DEFINE CLASS RemoteMariaDb AS LegacyRemote
	cLegacyClass = "MariaDB"
ENDDEFINE

DEFINE CLASS RemotePostgreSql AS LegacyRemote
	cLegacyClass = "PostgreSQL"
ENDDEFINE

DEFINE CLASS RemoteFirebird AS LegacyRemote
	cLegacyClass = "Firebird"
ENDDEFINE
