* FoxRemote 1.0 - one API for remote databases from Visual FoxPro 9, through ODBC.
* Engines: SQL Server, MySQL, MariaDB, PostgreSQL, Firebird and SQLite.
* License: MIT. Manual: README.md.
*
*   SET PROCEDURE TO foxremote.prg ADDITIVE
*   loDb = CREATEOBJECT("RemoteSqlServer")
*   loDb.cDriver = "ODBC Driver 17 for SQL Server"
*   loDb.cServer = "localhost\SQLEXPRESS"
*   loDb.cDatabase = "erp"
*   loDb.lTrustedConnection = .T.
*   IF !loDb.Connect()
*       ? loDb.cLastErrorCode, loDb.cLastError
*   ENDIF
*
* The library never shows anything (no dialogs, no waits): every failure adds one to nErrors and
* leaves cLastError, cLastErrorCode, nLastError, cLastSqlState and cLastSql; with lStrict it is
* also raised. It never creates or drops a database or a table unless a method says so by name,
* and it leaves no global SET changed (Open() sets MULTILOCKS ON in the data session of the cursor,
* which table buffering needs).
RETURN

#DEFINE FR_CRLF CHR(13) + CHR(10)

* ======================================================================== *
* RemoteDatabase: what every engine shares.
* ======================================================================== *
DEFINE CLASS RemoteDatabase AS Custom
	cVersion = "1.0.1"
	cEngine = ""

	* Connection
	cDriver = ""
	cServer = ""
	nPort = 0
	cDatabase = ""
	cUser = ""
	cPassword = ""
	lTrustedConnection = .F.
	cConnectionOptions = ""
	cConnectionString = ""
	nConnectTimeout = 15
	nQueryTimeout = 0
	lAutoReconnect = .F.
	nIdleCheckSeconds = 30
	nReconnects = 0
	nHandle = 0

	* Errors
	lStrict = .F.
	nErrors = 0
	cLastError = ""
	cLastErrorCode = ""
	nLastError = 0
	cLastSqlState = ""
	cLastSql = ""

	* Transactions
	lInTransaction = .F.
	nTransactionLevel = 0

	* Data policies
	cEmptyDateMode = "null"
	lTrimOnSave = .T.
	cAutoKeyName = ""
	lDetectConflicts = .F.

	* Engine traits (set by each subclass)
	lUtf8Data = .F.
	lReturning = .T.
	cNameCase = "lower"
	lFileDatabase = .F.
	* The row count of an UPDATE is the rows it changed, not the rows it found (MySQL, MariaDB).
	lCountsChangedRows = .F.

	PROTECTED oCursors, lAttached, cLastInsertTable, tLastActivity, oOpenError
	oCursors = .NULL.
	oOpenError = .NULL.
	lAttached = .F.
	cLastInsertTable = ""
	tLastActivity = {}

	PROCEDURE Init
		This.oCursors = CREATEOBJECT("Collection")
	ENDPROC

	PROCEDURE Destroy
		IF This.nHandle > 0 AND !This.lAttached
			This.Disconnect()
		ENDIF
	ENDPROC

	* ==================================================================== *
	* Errors
	* ==================================================================== *
	PROCEDURE ClearErrors
		This.nErrors = 0
		This.cLastError = ""
		This.cLastErrorCode = ""
		This.nLastError = 0
		This.cLastSqlState = ""
	ENDPROC

	* Records a failure; with lStrict it is also raised. Always returns .F. tvOdbc: .T. takes the
	* driver's error from AERROR(), an object from OdbcError() is one taken before.
	PROTECTED FUNCTION Fail(tcCode, tcMessage, tvOdbc)
		LOCAL loErr, lcDetail
		This.nErrors = This.nErrors + 1
		This.cLastErrorCode = tcCode
		This.nLastError = 0
		This.cLastSqlState = ""
		lcDetail = ""
		DO CASE
		CASE VARTYPE(tvOdbc) == "O"
			loErr = tvOdbc
		CASE VARTYPE(tvOdbc) == "L" AND tvOdbc
			loErr = This.OdbcError()
		OTHERWISE
			loErr = .NULL.
		ENDCASE
		IF !ISNULL(loErr)
			This.nLastError = loErr.nNative
			This.cLastSqlState = loErr.cState
			lcDetail = loErr.cText
		ENDIF
		This.cLastError = This.OneLine(tcMessage + IIF(EMPTY(lcDetail), "", ": " + lcDetail))
		IF This.lStrict
			ERROR This.cLastError
		ENDIF
		RETURN .F.
	ENDFUNC

	* The driver's last error: nNative, cState (SQLSTATE) and cText.
	PROTECTED FUNCTION OdbcError
		LOCAL laErr[1], loErr
		loErr = CREATEOBJECT("Empty")
		ADDPROPERTY(loErr, "nNative", 0)
		ADDPROPERTY(loErr, "cState", "")
		ADDPROPERTY(loErr, "cText", "")
		IF AERROR(laErr) > 0
			loErr.nNative = IIF(VARTYPE(laErr[1, 5]) == "N", laErr[1, 5], laErr[1, 1])
			loErr.cState = IIF(VARTYPE(laErr[1, 4]) == "C", ALLTRIM(laErr[1, 4]), "")
			loErr.cText = IIF(VARTYPE(laErr[1, 3]) == "C" AND !EMPTY(laErr[1, 3]), laErr[1, 3], TRANSFORM(laErr[1, 2]))
		ENDIF
		RETURN loErr
	ENDFUNC

	* Methods that clean up after a failure (roll back, put ids back) hold lStrict while they work and
	* raise at the end, so a failure never leaves their transaction open.
	PROTECTED FUNCTION HoldStrict
		LOCAL llStrict
		llStrict = This.lStrict
		This.lStrict = .F.
		RETURN llStrict
	ENDFUNC

	PROTECTED FUNCTION Raise(tlOk)
		IF !tlOk AND This.lStrict
			ERROR This.cLastError
		ENDIF
		RETURN tlOk
	ENDFUNC

	PROTECTED FUNCTION OneLine(tcText)
		LOCAL lcText
		lcText = CHRTRAN(tcText, CHR(13) + CHR(10) + CHR(9), "   ")
		DO WHILE "  " $ lcText
			lcText = STRTRAN(lcText, "  ", " ")
		ENDDO
		RETURN ALLTRIM(lcText)
	ENDFUNC

	* ==================================================================== *
	* Connection
	* ==================================================================== *
	FUNCTION Connect
		LOCAL lcConn, lnHandle, lnOldLogin, lnOldTimeout, loErr
		IF This.nHandle > 0
			IF !This.lAutoReconnect OR This.Ping()
				RETURN .T.
			ENDIF
			This.DropHandle()
		ENDIF
		IF This.lFileDatabase AND EMPTY(This.cConnectionString) AND !This.FileDatabaseExists(This.cDatabase)
			RETURN This.Fail("database_not_found", "database not found: " + This.cDatabase)
		ENDIF
		lcConn = IIF(EMPTY(This.cConnectionString), This.BuildConnectionString(This.cDatabase), This.cConnectionString)
		lnHandle = This.OpenHandle(lcConn)
		IF lnHandle < 1
			loErr = This.oOpenError
			IF !This.lFileDatabase AND !EMPTY(This.cDatabase) AND EMPTY(This.cConnectionString) AND This.ServerReachableWithoutDatabase()
				RETURN This.Fail("database_not_found", "database not found: " + This.cDatabase)
			ENDIF
			RETURN This.Fail("connection_failed", "connection failed", loErr)
		ENDIF
		This.nHandle = lnHandle
		This.lAttached = .F.
		This.SetupHandle()
		This.tLastActivity = DATETIME()
		RETURN .T.
	ENDFUNC

	* Uses a connection opened by somebody else; Disconnect() and Destroy leave it open.
	FUNCTION Attach(tnHandle)
		IF VARTYPE(tnHandle) # "N" OR tnHandle < 1
			RETURN This.Fail("connection_failed", "Attach() needs an open connection handle")
		ENDIF
		This.nHandle = tnHandle
		This.lAttached = .T.
		RETURN .T.
	ENDFUNC

	PROCEDURE Disconnect
		LOCAL llRolledBack
		IF This.nHandle < 1
			RETURN
		ENDIF
		llRolledBack = .F.
		IF This.lInTransaction
			* Rolled back explicitly: a driver may keep the attachment alive on the server when it is
			* asked to disconnect with a transaction open (measured with Firebird).
			=SQLROLLBACK(This.nHandle)
			=SQLSETPROP(This.nHandle, "Transactions", 1)
			This.lInTransaction = .F.
			This.nTransactionLevel = 0
			llRolledBack = .T.
		ENDIF
		IF !This.lAttached
			=SQLDISCONNECT(This.nHandle)
		ENDIF
		This.nHandle = 0
		This.lAttached = .F.
		IF llRolledBack
			This.Fail("rolled_back_on_disconnect", "a transaction was open: it was rolled back before disconnecting")
		ENDIF
	ENDPROC

	FUNCTION Ping
		LOCAL lnR
		IF This.nHandle < 1
			RETURN .F.
		ENDIF
		lnR = SQLEXEC(This.nHandle, This.PingSql(), "c_fr_ping")
		IF USED("c_fr_ping")
			USE IN c_fr_ping
		ENDIF
		RETURN lnR > 0
	ENDFUNC

	* After a failed statement: is the connection still there? A PostgreSQL transaction that had an
	* error refuses everything (SQLSTATE 25P02) until it is rolled back, but its connection is alive.
	PROTECTED FUNCTION Alive
		LOCAL lnR, loErr
		IF This.nHandle < 1
			RETURN .F.
		ENDIF
		lnR = SQLEXEC(This.nHandle, This.PingSql(), "c_fr_ping")
		IF USED("c_fr_ping")
			USE IN c_fr_ping
		ENDIF
		IF lnR > 0
			RETURN .T.
		ENDIF
		loErr = This.OdbcError()
		RETURN loErr.cState == "25P02"
	ENDFUNC

	* The connection is gone: forget it, and the transaction the server rolled back with it, and with
	* lAutoReconnect open a new one. .T. when there is a connection again.
	PROTECTED FUNCTION ConnectionLost
		This.DropHandle()
		IF !This.lAutoReconnect OR This.lAttached
			RETURN .F.
		ENDIF
		IF !This.Connect()
			RETURN .F.
		ENDIF
		This.nReconnects = This.nReconnects + 1
		RETURN .T.
	ENDFUNC

	* With lAutoReconnect, a connection idle for nIdleCheckSeconds is checked before it is used, and
	* opened again if the server dropped it (MySQL drops idle ones after 8 hours by default). Not inside
	* a transaction: that one went with the connection.
	PROTECTED FUNCTION CheckIdle
		IF !This.lAutoReconnect OR This.lInTransaction OR This.lAttached OR This.nHandle < 1
			RETURN .T.
		ENDIF
		IF !EMPTY(This.tLastActivity) AND DATETIME() - This.tLastActivity < This.nIdleCheckSeconds
			RETURN .T.
		ENDIF
		IF This.Alive()
			This.tLastActivity = DATETIME()
			RETURN .T.
		ENDIF
		IF This.ConnectionLost()
			RETURN .T.
		ENDIF
		RETURN This.Fail("connection_lost", "the connection to the server was lost and could not be opened again: " + This.cLastError)
	ENDFUNC

	FUNCTION DatabaseExists(tcDatabase)
		LOCAL lnH, lvCount
		IF This.lFileDatabase
			RETURN This.FileDatabaseExists(tcDatabase)
		ENDIF
		lnH = This.OpenHandle(This.BuildConnectionString(This.MaintenanceDatabase()))
		IF lnH < 1
			RETURN This.Fail("connection_failed", "connection failed", This.oOpenError)
		ENDIF
		lvCount = This.HandleScalar(lnH, This.DatabaseExistsSql(tcDatabase))
		=SQLDISCONNECT(lnH)
		RETURN This.AsNumber(lvCount) > 0
	ENDFUNC

	FUNCTION CreateDatabase(tcDatabase)
		LOCAL lnH, lnR
		IF EMPTY(tcDatabase)
			RETURN This.Fail("bad_argument", "CreateDatabase() needs a name")
		ENDIF
		IF This.DatabaseExists(tcDatabase)
			RETURN This.Fail("database_exists", "the database already exists: " + tcDatabase)
		ENDIF
		RETURN This.CreateDatabaseNow(tcDatabase)
	ENDFUNC

	* Server engines: through a short connection to the maintenance database.
	PROTECTED FUNCTION CreateDatabaseNow(tcDatabase)
		LOCAL lnH, lnR
		lnH = This.OpenHandle(This.BuildConnectionString(This.MaintenanceDatabase()))
		IF lnH < 1
			RETURN This.Fail("connection_failed", "connection failed", This.oOpenError)
		ENDIF
		This.cLastSql = This.CreateDatabaseSql(tcDatabase)
		lnR = SQLEXEC(lnH, This.cLastSql)
		IF lnR < 1
			This.Fail("query_failed", "CREATE DATABASE failed", .T.)
		ENDIF
		=SQLDISCONNECT(lnH)
		RETURN lnR > 0
	ENDFUNC

	* Switches the connection to another database of the same server.
	FUNCTION UseDatabase(tcDatabase)
		LOCAL lcOld
		lcOld = This.cDatabase
		This.Disconnect()
		This.cDatabase = tcDatabase
		IF !This.Connect()
			This.cDatabase = lcOld
			RETURN .F.
		ENDIF
		RETURN .T.
	ENDFUNC

	* ---- connection internals -----------------------------------------------
	* Opens a handle without ever letting ODBC show its login dialog.
	PROTECTED FUNCTION OpenHandle(tcConnection)
		LOCAL lnOldLogin, lnOldTimeout, lnH
		lnOldLogin = SQLGETPROP(0, "DispLogin")
		lnOldTimeout = SQLGETPROP(0, "ConnectTimeOut")
		=SQLSETPROP(0, "DispLogin", 3)
		=SQLSETPROP(0, "ConnectTimeOut", This.nConnectTimeout)
		lnH = SQLSTRINGCONNECT(tcConnection, .T.)
		* The driver's reason, taken now: the SQLSETPROP() calls below clear what AERROR() returns.
		This.oOpenError = IIF(lnH < 1, This.OdbcError(), .NULL.)
		=SQLSETPROP(0, "DispLogin", lnOldLogin)
		=SQLSETPROP(0, "ConnectTimeOut", lnOldTimeout)
		RETURN lnH
	ENDFUNC

	PROTECTED PROCEDURE SetupHandle
		=SQLSETPROP(This.nHandle, "DispLogin", 3)
		=SQLSETPROP(This.nHandle, "DispWarnings", .F.)
		=SQLSETPROP(This.nHandle, "Asynchronous", .F.)
		=SQLSETPROP(This.nHandle, "BatchMode", .T.)
		=SQLSETPROP(This.nHandle, "Transactions", 1)
		=SQLSETPROP(This.nHandle, "DisconnectRollback", .T.)
		=SQLSETPROP(This.nHandle, "IdleTimeout", 0)
		=SQLSETPROP(This.nHandle, "QueryTimeOut", This.nQueryTimeout)
	ENDPROC

	PROTECTED PROCEDURE DropHandle
		IF This.nHandle > 0 AND !This.lAttached
			=SQLDISCONNECT(This.nHandle)
		ENDIF
		This.nHandle = 0
		This.lInTransaction = .F.
		This.nTransactionLevel = 0
	ENDPROC

	* After a failed Connect(): with the same credentials, does the server answer without the database?
	PROTECTED FUNCTION ServerReachableWithoutDatabase
		LOCAL lnH, llMissing
		lnH = This.OpenHandle(This.BuildConnectionString(This.MaintenanceDatabase()))
		IF lnH < 1
			RETURN .F.
		ENDIF
		llMissing = This.AsNumber(This.HandleScalar(lnH, This.DatabaseExistsSql(This.cDatabase))) = 0
		=SQLDISCONNECT(lnH)
		RETURN llMissing
	ENDFUNC

	PROTECTED FUNCTION HandleScalar(tnHandle, tcSql)
		LOCAL lvValue
		lvValue = .NULL.
		IF SQLEXEC(tnHandle, tcSql, "c_fr_hs") > 0 AND USED("c_fr_hs")
			SELECT c_fr_hs
			IF !EOF()
				lvValue = EVALUATE(FIELD(1))
			ENDIF
			USE IN c_fr_hs
		ENDIF
		RETURN lvValue
	ENDFUNC

	PROTECTED FUNCTION FileDatabaseExists(tcDatabase)
		RETURN FILE(tcDatabase)
	ENDFUNC

	* A connection string value, wrapped in {} when it carries a character that would cut it.
	PROTECTED FUNCTION Value(tcValue)
		LOCAL lcValue
		lcValue = TRANSFORM(tcValue)
		IF ";" $ lcValue OR "{" $ lcValue OR "}" $ lcValue OR "=" $ lcValue OR !(lcValue == ALLTRIM(lcValue))
			RETURN "{" + STRTRAN(lcValue, "}", "}}") + "}"
		ENDIF
		RETURN lcValue
	ENDFUNC

	PROTECTED FUNCTION Options
		LOCAL lcOptions
		lcOptions = ALLTRIM(This.cConnectionOptions)
		IF !EMPTY(lcOptions) AND RIGHT(lcOptions, 1) # ";"
			lcOptions = lcOptions + ";"
		ENDIF
		RETURN lcOptions
	ENDFUNC

	PROTECTED FUNCTION Credentials
		RETURN "UID=" + This.Value(This.cUser) + ";PWD=" + This.Value(This.cPassword) + ";"
	ENDFUNC

	* ==================================================================== *
	* Queries
	* ==================================================================== *
	FUNCTION Query(tcSql, tcCursor, toParams)
		LOCAL lnR
		IF EMPTY(tcCursor)
			tcCursor = "c_fr_query"
		ENDIF
		lnR = This.Run(tcSql, tcCursor, toParams)
		IF lnR < 0
			RETURN .F.
		ENDIF
		IF USED(tcCursor)
			This.FromServerText(tcCursor)
		ENDIF
		RETURN .T.
	ENDFUNC

	* Rows affected, or -1 on failure.
	FUNCTION Execute(tcSql, toParams)
		LOCAL lcHead
		lcHead = UPPER(LEFT(LTRIM(CHRTRAN(tcSql, CHR(13) + CHR(10) + CHR(9), "   ")), 12))
		IF lcHead = "INSERT INTO "
			This.cLastInsertTable = This.FirstWord(SUBSTR(LTRIM(tcSql), 13))
		ENDIF
		RETURN This.Run(tcSql, "", toParams)
	ENDFUNC

	FUNCTION Scalar(tcSql, toParams)
		LOCAL lvValue, laF[1]
		lvValue = .NULL.
		IF This.Run(tcSql, "c_fr_scalar", toParams) >= 0 AND USED("c_fr_scalar")
			This.FromServerText("c_fr_scalar")
			SELECT c_fr_scalar
			IF !EOF() AND FCOUNT() > 0
				lvValue = EVALUATE(FIELD(1))
				=AFIELDS(laF)
				lvValue = This.BigIntValue(lvValue, laF[1, 2], laF[1, 3])
			ENDIF
			USE IN c_fr_scalar
		ENDIF
		RETURN lvValue
	ENDFUNC

	* The id generated by the last INSERT on this connection.
	FUNCTION LastId
		RETURN This.AsNumber(This.Scalar(This.LastIdSql()))
	ENDFUNC

	* Adds one to the counter tcColumn of the one row tcWhere picks, and returns the new value (-1 on
	* failure). The UPDATE locks that row until the transaction ends, so no other session gets the
	* same number: inside your transaction it stays locked until your Commit() or Rollback(), and
	* without one, NextNumber() runs in a transaction of its own. A NULL counter counts from 0.
	FUNCTION NextNumber(tcTable, tcColumn, tcWhere, toParams)
		LOCAL llStrict, lnValue
		llStrict = This.HoldStrict()
		TRY
			lnValue = This.NextNumberNow(tcTable, tcColumn, tcWhere, toParams)
		FINALLY
			This.lStrict = llStrict
		ENDTRY
		This.Raise(lnValue >= 0)
		RETURN lnValue
	ENDFUNC

	PROTECTED FUNCTION NextNumberNow(tcTable, tcColumn, tcWhere, toParams)
		LOCAL llOwn, lnRows, lnR, lnValue, lnErrors
		IF EMPTY(tcTable) OR EMPTY(tcColumn) OR EMPTY(tcWhere)
			This.Fail("bad_argument", "NextNumber() needs a table, a column and a WHERE")
			RETURN -1
		ENDIF
		llOwn = !This.lInTransaction
		IF llOwn AND !This.BeginTransaction()
			RETURN -1
		ENDIF
		lnValue = -1
		* Counted first, so a WHERE that picks several rows changes none of them.
		lnErrors = This.nErrors
		lnRows = This.AsNumber(This.Scalar("SELECT COUNT(*) FROM " + tcTable + " WHERE " + tcWhere, toParams))
		DO CASE
		CASE This.nErrors > lnErrors
			* the count failed: its error is already recorded
		CASE lnRows = 0
			This.Fail("row_not_found", "NextNumber(): no row of " + tcTable + " matches " + tcWhere)
		CASE lnRows > 1
			This.Fail("bad_argument", "NextNumber(): " + TRANSFORM(lnRows) + " rows of " + tcTable + " match " + tcWhere + "; it needs one")
		OTHERWISE
			lnR = This.Run("UPDATE " + tcTable + " SET " + tcColumn + " = COALESCE(" + tcColumn + ", 0) + 1 WHERE " + tcWhere, "", toParams)
			IF lnR >= 0
				lnValue = This.AsNumber(NVL(This.Scalar("SELECT " + tcColumn + " FROM " + tcTable + " WHERE " + tcWhere, toParams), -1))
			ENDIF
		ENDCASE
		IF llOwn
			IF lnValue >= 0
				IF !This.Commit()
					lnValue = -1
				ENDIF
			ELSE
				IF This.lInTransaction
					This.RollbackQuietly()
				ENDIF
			ENDIF
		ENDIF
		RETURN lnValue
	ENDFUNC

	* Runs a statement. :name placeholders take their value from the same property of toParams.
	* Returns the rows affected (0 when the driver does not say), or -1 on failure.
	PROTECTED FUNCTION Run(tcSql, tcCursor, toParams)
		LOCAL loFrP, lcSql, lnR, laCount[1], lcCursor, lnRows, lnI, loErr, llInTransaction
		IF This.nHandle < 1
			IF !This.lAutoReconnect OR !This.Connect()
				This.cLastSql = tcSql
				RETURN IIF(This.Fail("not_connected", "not connected"), -1, -1)
			ENDIF
		ENDIF
		This.cLastSql = tcSql
		IF !This.CheckIdle()
			RETURN -1
		ENDIF
		loFrP = This.PrepareParams(toParams)
		lcSql = This.TranslateParams(tcSql, "loFrP")
		lcCursor = IIF(EMPTY(tcCursor), "c_fr_noresult", tcCursor)
		lnR = SQLEXEC(This.nHandle, lcSql, lcCursor, laCount)
		This.tLastActivity = DATETIME()
		IF lnR < 1
			loErr = This.OdbcError()
			IF This.Alive()
				This.Fail("query_failed", "statement failed", loErr)
				RETURN -1
			ENDIF
			* The connection is gone. A read is run again on a new one; a write is not, because the
			* server may have run it, and a transaction is not, because the server rolled it back.
			llInTransaction = This.lInTransaction
			IF !This.ConnectionLost()
				This.Fail("connection_lost", "the connection to the server was lost" + ;
					IIF(llInTransaction, ", and the server rolled back the open transaction", ""), loErr)
				RETURN -1
			ENDIF
			IF llInTransaction OR !This.IsRead(tcSql)
				This.Fail("connection_lost", "the connection to the server was lost and was opened again; " + ;
					IIF(llInTransaction, "the server rolled back the open transaction", "the statement was not repeated, because the server may have run it"), loErr)
				RETURN -1
			ENDIF
			lnR = SQLEXEC(This.nHandle, lcSql, lcCursor, laCount)
			This.tLastActivity = DATETIME()
			IF lnR < 1
				This.Fail("query_failed", "statement failed", .T.)
				RETURN -1
			ENDIF
		ENDIF
		lnRows = 0
		IF ALEN(laCount, 2) >= 2
			FOR lnI = 1 TO ALEN(laCount, 1)
				IF VARTYPE(laCount[lnI, 2]) == "N" AND laCount[lnI, 2] > 0
					lnRows = lnRows + laCount[lnI, 2]
				ENDIF
			ENDFOR
		ENDIF
		IF EMPTY(tcCursor) AND USED("c_fr_noresult")
			USE IN c_fr_noresult
		ENDIF
		RETURN lnRows
	ENDFUNC

	* :name -> ?<object>.name, outside quotes; '::' (a PostgreSQL cast) is left alone.
	PROTECTED FUNCTION TranslateParams(tcSql, tcObject)
		LOCAL lcOut, lnI, lnLen, lcCh, lcQuote, lnJ, lcName
		IF !(":" $ tcSql)
			RETURN tcSql
		ENDIF
		lcOut = ""
		lcQuote = ""
		lnLen = LEN(tcSql)
		lnI = 1
		DO WHILE lnI <= lnLen
			lcCh = SUBSTR(tcSql, lnI, 1)
			DO CASE
			CASE !EMPTY(lcQuote)
				IF lcCh == lcQuote
					lcQuote = ""
				ENDIF
				lcOut = lcOut + lcCh
				lnI = lnI + 1
			CASE INLIST(lcCh, "'", '"')
				lcQuote = lcCh
				lcOut = lcOut + lcCh
				lnI = lnI + 1
			CASE lcCh == ":" AND SUBSTR(tcSql, lnI + 1, 1) == ":"
				lcOut = lcOut + "::"
				lnI = lnI + 2
			CASE lcCh == ":" AND (ISALPHA(SUBSTR(tcSql, lnI + 1, 1)) OR SUBSTR(tcSql, lnI + 1, 1) == "_")
				lnJ = lnI + 1
				DO WHILE lnJ <= lnLen AND (ISALPHA(SUBSTR(tcSql, lnJ, 1)) OR ISDIGIT(SUBSTR(tcSql, lnJ, 1)) OR SUBSTR(tcSql, lnJ, 1) == "_")
					lnJ = lnJ + 1
				ENDDO
				lcName = SUBSTR(tcSql, lnI + 1, lnJ - lnI - 1)
				lcOut = lcOut + "?" + tcObject + "." + lcName
				lnI = lnJ
			OTHERWISE
				lcOut = lcOut + lcCh
				lnI = lnI + 1
			ENDCASE
		ENDDO
		RETURN lcOut
	ENDFUNC

	* The values as the driver has to receive them (UTF-8 for SQLite).
	PROTECTED FUNCTION PrepareParams(toParams)
		LOCAL loCopy, laM[1], lnI, lvValue
		IF VARTYPE(toParams) # "O" OR !This.lUtf8Data
			RETURN toParams
		ENDIF
		loCopy = CREATEOBJECT("Empty")
		FOR lnI = 1 TO AMEMBERS(laM, toParams, 0)
			lvValue = GETPEM(toParams, laM[lnI])
			ADDPROPERTY(loCopy, laM[lnI], IIF(VARTYPE(lvValue) == "C", STRCONV(lvValue, 9), lvValue))
		ENDFOR
		RETURN loCopy
	ENDFUNC

	* SQLite keeps UTF-8 and its driver hands the bytes over as they are: convert them to CP1252.
	PROTECTED PROCEDURE FromServerText(tcAlias)
		LOCAL laF[1], lnI, lnN, lcFields, laText[1], lnT, lcField, lvValue, lnSelect
		IF !This.lUtf8Data OR !USED(tcAlias)
			RETURN
		ENDIF
		lnT = 0
		FOR lnI = 1 TO AFIELDS(laF, tcAlias)
			IF INLIST(laF[lnI, 2], "C", "M", "V")
				lnT = lnT + 1
				DIMENSION laText[lnT]
				laText[lnT] = laF[lnI, 1]
			ENDIF
		ENDFOR
		IF lnT = 0
			RETURN
		ENDIF
		lnSelect = SELECT()
		SELECT (tcAlias)
		SCAN
			FOR lnI = 1 TO lnT
				lcField = laText[lnI]
				lvValue = EVALUATE(lcField)
				IF !ISNULL(lvValue) AND !EMPTY(lvValue)
					REPLACE (lcField) WITH STRCONV(lvValue, 11)
				ENDIF
			ENDFOR
		ENDSCAN
		GO TOP
		SELECT (lnSelect)
	ENDPROC

	* VFP receives BIGINT as C(20): give it back as a number.
	PROTECTED FUNCTION BigIntValue(tvValue, tcType, tnWidth)
		IF INLIST(tcType, "C", "V") AND BETWEEN(tnWidth, 19, 21) AND VARTYPE(tvValue) == "C" AND !EMPTY(tvValue) ;
				AND EMPTY(CHRTRAN(ALLTRIM(tvValue), "0123456789-", ""))
			RETURN VAL(tvValue)
		ENDIF
		RETURN tvValue
	ENDFUNC

	PROTECTED FUNCTION AsNumber(tvValue)
		DO CASE
		CASE VARTYPE(tvValue) == "N"
			RETURN tvValue
		CASE VARTYPE(tvValue) == "C" AND !EMPTY(tvValue)
			RETURN VAL(tvValue)
		ENDCASE
		RETURN 0
	ENDFUNC

	PROTECTED FUNCTION IsRead(tcSql)
		RETURN UPPER(LEFT(LTRIM(CHRTRAN(tcSql, CHR(13) + CHR(10) + CHR(9) + "(", "    ")), 7)) == "SELECT "
	ENDFUNC

	PROTECTED FUNCTION FirstWord(tcText)
		LOCAL lcText, lnI
		lcText = LTRIM(tcText)
		FOR lnI = 1 TO LEN(lcText)
			IF INLIST(SUBSTR(lcText, lnI, 1), " ", "(", CHR(9), CHR(13), CHR(10))
				RETURN LEFT(lcText, lnI - 1)
			ENDIF
		ENDFOR
		RETURN lcText
	ENDFUNC

	* ==================================================================== *
	* Transactions
	* ==================================================================== *
	FUNCTION BeginTransaction
		IF This.nHandle < 1 AND !(This.lAutoReconnect AND This.Connect())
			RETURN This.Fail("not_connected", "not connected")
		ENDIF
		IF This.nTransactionLevel = 0
			IF !This.CheckIdle()
				RETURN .F.
			ENDIF
			IF SQLSETPROP(This.nHandle, "Transactions", 2) < 1
				RETURN This.Fail("transaction_failed", "could not start a transaction", .T.)
			ENDIF
			This.lInTransaction = .T.
		ENDIF
		This.nTransactionLevel = This.nTransactionLevel + 1
		RETURN .T.
	ENDFUNC

	FUNCTION Commit
		LOCAL llOk
		IF This.nTransactionLevel = 0
			RETURN This.Fail("no_transaction", "Commit() without a transaction")
		ENDIF
		This.nTransactionLevel = This.nTransactionLevel - 1
		IF This.nTransactionLevel > 0
			RETURN .T.
		ENDIF
		llOk = SQLCOMMIT(This.nHandle) > 0
		IF !llOk
			This.Fail("transaction_failed", "COMMIT failed", .T.)
			=SQLROLLBACK(This.nHandle)
		ENDIF
		=SQLSETPROP(This.nHandle, "Transactions", 1)
		This.lInTransaction = .F.
		RETURN llOk
	ENDFUNC

	FUNCTION Rollback
		LOCAL llOk
		IF This.nTransactionLevel = 0
			RETURN This.Fail("no_transaction", "Rollback() without a transaction")
		ENDIF
		llOk = SQLROLLBACK(This.nHandle) > 0
		IF !llOk
			This.Fail("transaction_failed", "ROLLBACK failed", .T.)
		ENDIF
		=SQLSETPROP(This.nHandle, "Transactions", 1)
		This.nTransactionLevel = 0
		This.lInTransaction = .F.
		RETURN llOk
	ENDFUNC

	* ==================================================================== *
	* Updatable cursors
	* ==================================================================== *
	FUNCTION Open(tcTable, tcAlias, tcWhere, toParams, tcFields, tlReadOnly, tcGroup)
		LOCAL lcAlias, lcKey, lcSql, loInfo, lcKeyName, lnErrors, lcLongSql
		IF EMPTY(tcTable)
			RETURN This.Fail("bad_argument", "Open() needs a table")
		ENDIF
		lcAlias = IIF(EMPTY(tcAlias), CHRTRAN(tcTable, ".", "_"), tcAlias)
		lcKeyName = This.CursorKey(lcAlias)
		IF This.oCursors.GetKey(lcKeyName) > 0
			IF USED(lcAlias)
				RETURN This.Fail("alias_in_use", "the alias is already open: " + lcAlias)
			ENDIF
			This.oCursors.Remove(lcKeyName)
		ENDIF
		IF USED(lcAlias)
			RETURN This.Fail("alias_in_use", "the alias is in use: " + lcAlias)
		ENDIF
		lcKey = ""
		IF !tlReadOnly
			lnErrors = This.nErrors
			lcKey = This.GetPrimaryKey(tcTable)
			IF EMPTY(lcKey)
				IF This.nErrors = lnErrors
					This.Fail("no_primary_key", "a table without a primary key can only be opened read-only: " + tcTable)
				ENDIF
				RETURN .F.
			ENDIF
		ENDIF
		lcSql = "SELECT " + IIF(EMPTY(tcFields), "*", tcFields) + " FROM " + tcTable + IIF(EMPTY(tcWhere), "", " WHERE " + tcWhere)
		IF !This.Query(lcSql, lcAlias, toParams)
			RETURN .F.
		ENDIF
		lcLongSql = This.LongTextSelect(lcAlias, tcTable, tcWhere)
		IF !EMPTY(lcLongSql)
			lcSql = lcLongSql
			IF !This.Query(lcSql, lcAlias, toParams)
				RETURN .F.
			ENDIF
		ENDIF
		loInfo = CREATEOBJECT("Empty")
		ADDPROPERTY(loInfo, "cAlias", lcAlias)
		ADDPROPERTY(loInfo, "cTable", tcTable)
		ADDPROPERTY(loInfo, "cKey", lcKey)
		ADDPROPERTY(loInfo, "cSql", lcSql)
		ADDPROPERTY(loInfo, "oParams", toParams)
		ADDPROPERTY(loInfo, "lReadOnly", tlReadOnly)
		ADDPROPERTY(loInfo, "cGroup", LOWER(IIF(EMPTY(tcGroup), "", tcGroup)))
		ADDPROPERTY(loInfo, "nSession", SET("DATASESSION"))
		This.oCursors.Add(loInfo, lcKeyName)
		This.MakeBuffered(loInfo)
		SELECT (lcAlias)
		RETURN .T.
	ENDFUNC

	FUNCTION Save(tcAlias)
		LOCAL llStrict, llOk
		llStrict = This.HoldStrict()
		TRY
			llOk = This.SaveNow(tcAlias)
		FINALLY
			This.lStrict = llStrict
		ENDTRY
		RETURN This.Raise(llOk)
	ENDFUNC

	PROTECTED FUNCTION SaveNow(tcAlias)
		LOCAL loInfo, llOwn, llOk, loUndo
		loInfo = This.CursorInfo(tcAlias)
		IF ISNULL(loInfo)
			RETURN This.Fail("not_open", "this alias was not opened by FoxRemote: " + TRANSFORM(tcAlias))
		ENDIF
		IF loInfo.lReadOnly
			RETURN This.Fail("read_only", "the cursor was opened read-only: " + loInfo.cAlias)
		ENDIF
		loUndo = CREATEOBJECT("Collection")
		llOwn = !This.lInTransaction
		IF llOwn AND !This.BeginTransaction()
			RETURN .F.
		ENDIF
		llOk = This.SaveCursor(loInfo, loUndo)
		RETURN This.FinishSave(llOk, llOwn, loUndo, loInfo)
	ENDFUNC

	FUNCTION SaveGroup(tcGroup)
		LOCAL llStrict, llOk
		llStrict = This.HoldStrict()
		TRY
			llOk = This.SaveGroupNow(tcGroup)
		FINALLY
			This.lStrict = llStrict
		ENDTRY
		RETURN This.Raise(llOk)
	ENDFUNC

	PROTECTED FUNCTION SaveGroupNow(tcGroup)
		LOCAL loInfo, llOwn, llOk, loUndo, loList
		loList = This.GroupCursors(tcGroup)
		IF loList.Count = 0
			RETURN This.Fail("not_open", "no open cursor in the group: " + TRANSFORM(tcGroup))
		ENDIF
		loUndo = CREATEOBJECT("Collection")
		llOwn = !This.lInTransaction
		IF llOwn AND !This.BeginTransaction()
			RETURN .F.
		ENDIF
		llOk = .T.
		FOR EACH loInfo IN loList
			IF llOk AND !loInfo.lReadOnly
				llOk = This.SaveCursor(loInfo, loUndo)
			ENDIF
		ENDFOR
		RETURN This.FinishSave(llOk, llOwn, loUndo, loList)
	ENDFUNC

	FUNCTION Discard(tcAlias)
		LOCAL loInfo
		loInfo = This.CursorInfo(tcAlias)
		IF ISNULL(loInfo)
			RETURN This.Fail("not_open", "this alias was not opened by FoxRemote: " + TRANSFORM(tcAlias))
		ENDIF
		IF !loInfo.lReadOnly
			=TABLEREVERT(.T., loInfo.cAlias)
		ENDIF
		RETURN .T.
	ENDFUNC

	* Reads the cursor again from the server. Refused while it has changes not saved.
	FUNCTION Refresh(tcAlias)
		LOCAL loInfo
		loInfo = This.CursorInfo(tcAlias)
		IF ISNULL(loInfo)
			RETURN This.Fail("not_open", "this alias was not opened by FoxRemote: " + TRANSFORM(tcAlias))
		ENDIF
		IF !loInfo.lReadOnly AND GETNEXTMODIFIED(0, loInfo.cAlias) # 0
			RETURN This.Fail("pending_changes", "the cursor has changes not saved: Save() or Discard() first")
		ENDIF
		IF !This.Query(loInfo.cSql, loInfo.cAlias, loInfo.oParams)
			RETURN .F.
		ENDIF
		This.MakeBuffered(loInfo)
		RETURN .T.
	ENDFUNC

	* Whether the cursor has changes not saved. Without an alias: whether any cursor this object
	* opened in the current data session has.
	FUNCTION HasChanges(tcAlias)
		LOCAL loInfo
		IF EMPTY(tcAlias)
			FOR EACH loInfo IN This.oCursors
				IF loInfo.nSession = SET("DATASESSION") AND This.InfoHasChanges(loInfo)
					RETURN .T.
				ENDIF
			ENDFOR
			RETURN .F.
		ENDIF
		loInfo = This.CursorInfo(tcAlias)
		IF ISNULL(loInfo)
			RETURN This.Fail("not_open", "this alias was not opened by FoxRemote: " + TRANSFORM(tcAlias))
		ENDIF
		RETURN This.InfoHasChanges(loInfo)
	ENDFUNC

	* Closes a cursor. One with changes not saved is kept open (pending_changes) unless tlDiscard.
	FUNCTION Close(tcAlias, tlDiscard)
		LOCAL loInfo
		loInfo = This.CursorInfo(tcAlias)
		IF ISNULL(loInfo)
			RETURN This.Fail("not_open", "this alias was not opened by FoxRemote: " + TRANSFORM(tcAlias))
		ENDIF
		IF !tlDiscard AND This.InfoHasChanges(loInfo)
			RETURN This.Fail("pending_changes", "the cursor has changes not saved: Save() it, or Close() it with tlDiscard: " + loInfo.cAlias)
		ENDIF
		This.CloseInfo(loInfo)
		RETURN .T.
	ENDFUNC

	* All the cursors of the group, or none: none when one has changes not saved, unless tlDiscard.
	FUNCTION CloseGroup(tcGroup, tlDiscard)
		LOCAL loInfo, loList
		loList = This.GroupCursors(tcGroup)
		RETURN This.CloseList(loList, tlDiscard)
	ENDFUNC

	* The cursors FoxRemote opened in the current data session, and nothing else; all or none, as
	* CloseGroup().
	FUNCTION CloseAll(tlDiscard)
		LOCAL loInfo, loList
		loList = CREATEOBJECT("Collection")
		FOR EACH loInfo IN This.oCursors
			IF loInfo.nSession = SET("DATASESSION")
				loList.Add(loInfo)
			ENDIF
		ENDFOR
		RETURN This.CloseList(loList, tlDiscard)
	ENDFUNC

	PROTECTED FUNCTION CloseList(toList, tlDiscard)
		LOCAL loInfo
		IF !tlDiscard
			FOR EACH loInfo IN toList
				IF This.InfoHasChanges(loInfo)
					RETURN This.Fail("pending_changes", "a cursor has changes not saved: Save() it, or close with tlDiscard: " + loInfo.cAlias)
				ENDIF
			ENDFOR
		ENDIF
		FOR EACH loInfo IN toList
			This.CloseInfo(loInfo)
		ENDFOR
		RETURN .T.
	ENDFUNC

	PROTECTED FUNCTION InfoHasChanges(toInfo)
		RETURN !toInfo.lReadOnly AND USED(toInfo.cAlias) AND GETNEXTMODIFIED(0, toInfo.cAlias) # 0
	ENDFUNC

	* ---- cursor internals ---------------------------------------------------
	PROTECTED FUNCTION CursorKey(tcAlias)
		RETURN LOWER(ALLTRIM(tcAlias)) + "@" + TRANSFORM(SET("DATASESSION"))
	ENDFUNC

	PROTECTED FUNCTION CursorInfo(tcAlias)
		LOCAL lnAt, loInfo
		IF EMPTY(tcAlias)
			tcAlias = ALIAS()
		ENDIF
		lnAt = This.oCursors.GetKey(This.CursorKey(tcAlias))
		IF lnAt = 0
			RETURN .NULL.
		ENDIF
		loInfo = This.oCursors.Item(lnAt)
		IF !USED(loInfo.cAlias)
			This.oCursors.Remove(lnAt)
			RETURN .NULL.
		ENDIF
		RETURN loInfo
	ENDFUNC

	PROTECTED FUNCTION GroupCursors(tcGroup)
		LOCAL loList, loInfo, lcGroup
		loList = CREATEOBJECT("Collection")
		lcGroup = LOWER(ALLTRIM(TRANSFORM(tcGroup)))
		FOR EACH loInfo IN This.oCursors
			IF loInfo.cGroup == lcGroup AND loInfo.nSession = SET("DATASESSION") AND USED(loInfo.cAlias)
				loList.Add(loInfo)
			ENDIF
		ENDFOR
		RETURN loList
	ENDFUNC

	PROTECTED PROCEDURE CloseInfo(toInfo)
		LOCAL lnAt
		IF USED(toInfo.cAlias)
			IF !toInfo.lReadOnly
				=TABLEREVERT(.T., toInfo.cAlias)
			ENDIF
			USE IN (toInfo.cAlias)
		ENDIF
		lnAt = This.oCursors.GetKey(LOWER(ALLTRIM(toInfo.cAlias)) + "@" + TRANSFORM(toInfo.nSession))
		IF lnAt > 0
			This.oCursors.Remove(lnAt)
		ENDIF
	ENDPROC

	PROTECTED PROCEDURE MakeBuffered(toInfo)
		IF !toInfo.lReadOnly
			SET MULTILOCKS ON
			=CURSORSETPROP("Buffering", 5, toInfo.cAlias)
		ENDIF
	ENDPROC

	* Sends one cursor's changes inside the transaction already open. loUndo collects the ids
	* written into new rows, so a failure can put them back.
	PROTECTED FUNCTION SaveCursor(toInfo, toUndo)
		LOCAL lcAlias, lnRec, lcState, laF[1], lnFields, llOk, lnSelect
		lcAlias = toInfo.cAlias
		lnSelect = SELECT()
		SELECT (lcAlias)
		lnFields = AFIELDS(laF, lcAlias)
		llOk = .T.
		lnRec = GETNEXTMODIFIED(0, lcAlias)
		DO WHILE llOk AND lnRec # 0
			GO (lnRec) IN (lcAlias)
			lcState = GETFLDSTATE(-1, lcAlias)
			DO CASE
			CASE lnRec < 0 AND DELETED(lcAlias)
				* added and deleted before saving: nothing to send
			CASE DELETED(lcAlias)
				llOk = This.SendDelete(toInfo)
			CASE lnRec < 0 OR INLIST(LEFT(lcState, 1), "3", "4")
				llOk = This.SendInsert(toInfo, lcState, @laF, lnFields, toUndo, lnRec)
			OTHERWISE
				llOk = This.SendUpdate(toInfo, lcState, @laF, lnFields)
			ENDCASE
			lnRec = GETNEXTMODIFIED(lnRec, lcAlias)
		ENDDO
		SELECT (lnSelect)
		RETURN llOk
	ENDFUNC

	PROTECTED FUNCTION FinishSave(tlOk, tlOwn, toUndo, tvCursors)
		LOCAL loEntry, loInfo
		IF tlOk
			IF tlOwn AND !This.Commit()
				tlOk = .F.
			ENDIF
		ENDIF
		IF tlOk
			IF VARTYPE(tvCursors) == "O" AND PEMSTATUS(tvCursors, "cAlias", 5)
				=TABLEUPDATE(.T., .T., tvCursors.cAlias)
			ELSE
				FOR EACH loInfo IN tvCursors
					IF !loInfo.lReadOnly
						=TABLEUPDATE(.T., .T., loInfo.cAlias)
					ENDIF
				ENDFOR
			ENDIF
			RETURN .T.
		ENDIF
		IF tlOwn AND This.lInTransaction
			This.RollbackQuietly()
		ENDIF
		FOR EACH loEntry IN toUndo
			IF USED(loEntry.cAlias)
				GO (loEntry.nRec) IN (loEntry.cAlias)
				REPLACE (loEntry.cField) WITH loEntry.vOld IN (loEntry.cAlias)
			ENDIF
		ENDFOR
		IF !INLIST(This.cLastErrorCode, "row_not_found", "row_changed", "connection_lost")
			This.cLastErrorCode = "save_failed"
		ENDIF
		RETURN .F.
	ENDFUNC

	PROTECTED PROCEDURE RollbackQuietly
		=SQLROLLBACK(This.nHandle)
		=SQLSETPROP(This.nHandle, "Transactions", 1)
		This.nTransactionLevel = 0
		This.lInTransaction = .F.
	ENDPROC

	PROTECTED FUNCTION SendInsert(toInfo, tcState, taF, tnFields, toUndo, tnRec)
		EXTERNAL ARRAY taF
		LOCAL loP, lcCols, lcVals, lnI, lnN, lcField, lvValue, lcSql, lcKey, llReturn, lvId, loEntry, lnR
		loP = CREATEOBJECT("Empty")
		lcCols = ""
		lcVals = ""
		lnN = 0
		lcKey = toInfo.cKey
		FOR lnI = 1 TO tnFields
			lcField = taF[lnI, 1]
			lvValue = EVALUATE(toInfo.cAlias + "." + lcField)
			IF SUBSTR(tcState, lnI + 1, 1) # "4"
				LOOP
			ENDIF
			IF This.IsKeyField(lcField, lcKey) AND EMPTY(lvValue)
				LOOP
			ENDIF
			lnN = lnN + 1
			ADDPROPERTY(loP, "p" + TRANSFORM(lnN), This.ToServerValue(lvValue, taF[lnI, 2], taF[lnI, 5]))
			lcCols = lcCols + IIF(EMPTY(lcCols), "", ", ") + This.ColumnName(lcField)
			lcVals = lcVals + IIF(EMPTY(lcVals), "", ", ") + ":p" + TRANSFORM(lnN)
		ENDFOR
		llReturn = This.lReturning AND !EMPTY(lcKey) AND !("," $ lcKey)
		lcSql = This.InsertSql(toInfo.cTable, lcCols, lcVals, IIF(llReturn, lcKey, ""))
		IF !llReturn
			lnR = This.Run(lcSql, "", loP)
			IF lnR < 0
				RETURN .F.
			ENDIF
			IF !EMPTY(lcKey) AND !("," $ lcKey) AND EMPTY(EVALUATE(toInfo.cAlias + "." + lcKey))
				This.cLastInsertTable = toInfo.cTable
				lvId = This.LastId()
				This.WriteKey(toInfo, lcKey, lvId, toUndo, tnRec)
			ENDIF
			RETURN .T.
		ENDIF
		IF This.Run(lcSql, "c_fr_ret", loP) < 0
			RETURN .F.
		ENDIF
		lvId = .NULL.
		IF USED("c_fr_ret")
			SELECT c_fr_ret
			IF !EOF() AND FCOUNT() > 0
				lvId = This.KeyValue(toInfo, lcKey, EVALUATE(FIELD(1)))
			ENDIF
			USE IN c_fr_ret
			SELECT (toInfo.cAlias)
		ENDIF
		IF !ISNULL(lvId)
			This.WriteKey(toInfo, lcKey, lvId, toUndo, tnRec)
		ENDIF
		RETURN .T.
	ENDFUNC

	* A generated key as the cursor's field holds it: a number for a numeric key (BIGINT arrives as
	* text), the value as it came for a character key such as a UUID.
	PROTECTED FUNCTION KeyValue(toInfo, tcKey, tvRaw)
		IF ISNULL(tvRaw)
			RETURN .NULL.
		ENDIF
		IF TYPE(toInfo.cAlias + "." + tcKey) == "N"
			RETURN This.AsNumber(tvRaw)
		ENDIF
		RETURN tvRaw
	ENDFUNC

	PROTECTED PROCEDURE WriteKey(toInfo, tcKey, tvId, toUndo, tnRec)
		LOCAL loEntry, lvOld
		lvOld = EVALUATE(toInfo.cAlias + "." + tcKey)
		IF lvOld == tvId
			RETURN
		ENDIF
		loEntry = CREATEOBJECT("Empty")
		ADDPROPERTY(loEntry, "cAlias", toInfo.cAlias)
		ADDPROPERTY(loEntry, "nRec", tnRec)
		ADDPROPERTY(loEntry, "cField", tcKey)
		ADDPROPERTY(loEntry, "vOld", lvOld)
		toUndo.Add(loEntry)
		GO (tnRec) IN (toInfo.cAlias)
		REPLACE (tcKey) WITH tvId IN (toInfo.cAlias)
	ENDPROC

	PROTECTED FUNCTION SendUpdate(toInfo, tcState, taF, tnFields)
		EXTERNAL ARRAY taF
		LOCAL loP, lcSet, lnI, lnN, lcField, lcSql, lnR
		IF EMPTY(toInfo.cKey)
			RETURN This.Fail("no_primary_key", "no primary key for " + toInfo.cTable)
		ENDIF
		loP = CREATEOBJECT("Empty")
		lcSet = ""
		lnN = 0
		FOR lnI = 1 TO tnFields
			IF SUBSTR(tcState, lnI + 1, 1) # "2"
				LOOP
			ENDIF
			lcField = taF[lnI, 1]
			lnN = lnN + 1
			ADDPROPERTY(loP, "p" + TRANSFORM(lnN), This.ToServerValue(EVALUATE(toInfo.cAlias + "." + lcField), taF[lnI, 2], taF[lnI, 5]))
			lcSet = lcSet + IIF(EMPTY(lcSet), "", ", ") + This.ColumnName(lcField) + " = :p" + TRANSFORM(lnN)
		ENDFOR
		IF EMPTY(lcSet)
			RETURN .T.
		ENDIF
		lcSql = "UPDATE " + toInfo.cTable + " SET " + lcSet + " WHERE " + This.RowWhere(toInfo, tcState, @taF, tnFields, loP, @lnN, This.lDetectConflicts)
		lnR = This.Run(lcSql, "", loP)
		IF lnR < 0
			RETURN .F.
		ENDIF
		IF lnR = 0
			RETURN This.NothingUpdated(toInfo, tcState, @taF, tnFields)
		ENDIF
		RETURN .T.
	ENDFUNC

	* An UPDATE changed no row: the row is gone, or another user changed what this one compares.
	* MySQL and MariaDB also count 0 for a row that already had the values sent, which is no failure.
	PROTECTED FUNCTION NothingUpdated(toInfo, tcState, taF, tnFields)
		EXTERNAL ARRAY taF
		IF !This.RowStillThere(toInfo, tcState, @taF, tnFields, .F.)
			RETURN This.Fail("row_not_found", "the row to update is no longer on the server (" + toInfo.cTable + ")")
		ENDIF
		IF This.lDetectConflicts AND !(This.lCountsChangedRows AND This.RowStillThere(toInfo, tcState, @taF, tnFields, .T.))
			RETURN This.Fail("row_changed", "another user changed the row after it was read (" + toInfo.cTable + ")")
		ENDIF
		RETURN .T.
	ENDFUNC

	PROTECTED FUNCTION RowStillThere(toInfo, tcState, taF, tnFields, tlConflicts)
		EXTERNAL ARRAY taF
		LOCAL loP, lnN
		loP = CREATEOBJECT("Empty")
		lnN = 0
		RETURN This.AsNumber(This.Scalar("SELECT COUNT(*) FROM " + toInfo.cTable + " WHERE " + ;
			This.RowWhere(toInfo, tcState, @taF, tnFields, loP, @lnN, tlConflicts), loP)) > 0
	ENDFUNC

	* WHERE for the row as it was read: its key and, with tlConflicts, the values the columns this
	* update changes had when the cursor read them. Memo, binary, floating point and datetime columns
	* are not compared: they do not come back equal (a datetime keeps no fractions of a second in VFP).
	PROTECTED FUNCTION RowWhere(toInfo, tcState, taF, tnFields, toParams, tnN, tlConflicts)
		EXTERNAL ARRAY taF
		LOCAL lcWhere, lnI, lcField, lvOld
		lcWhere = This.KeyWhere(toInfo, toParams, @tnN)
		IF !tlConflicts
			RETURN lcWhere
		ENDIF
		FOR lnI = 1 TO tnFields
			IF SUBSTR(tcState, lnI + 1, 1) # "2" OR INLIST(taF[lnI, 2], "M", "G", "W", "Q", "F", "B", "T")
				LOOP
			ENDIF
			lcField = taF[lnI, 1]
			IF This.IsKeyField(lcField, toInfo.cKey)
				LOOP
			ENDIF
			lvOld = This.ToServerValue(OLDVAL(lcField, toInfo.cAlias), taF[lnI, 2], taF[lnI, 5])
			IF ISNULL(lvOld)
				lcWhere = lcWhere + " AND " + This.ColumnName(lcField) + " IS NULL"
			ELSE
				tnN = tnN + 1
				ADDPROPERTY(toParams, "p" + TRANSFORM(tnN), lvOld)
				lcWhere = lcWhere + " AND " + This.ColumnName(lcField) + " = :p" + TRANSFORM(tnN)
			ENDIF
		ENDFOR
		RETURN lcWhere
	ENDFUNC

	PROTECTED FUNCTION SendDelete(toInfo)
		LOCAL loP, lnN, lcSql, lnR
		IF EMPTY(toInfo.cKey)
			RETURN This.Fail("no_primary_key", "no primary key for " + toInfo.cTable)
		ENDIF
		loP = CREATEOBJECT("Empty")
		lnN = 0
		lcSql = "DELETE FROM " + toInfo.cTable + " WHERE " + This.KeyWhere(toInfo, loP, @lnN)
		lnR = This.Run(lcSql, "", loP)
		IF lnR < 0
			RETURN .F.
		ENDIF
		IF lnR = 0
			RETURN This.Fail("row_not_found", "the row to delete is no longer on the server (" + toInfo.cTable + ")")
		ENDIF
		RETURN .T.
	ENDFUNC

	* WHERE on the primary key, with the values the row had when it was read.
	PROTECTED FUNCTION KeyWhere(toInfo, toParams, tnN)
		LOCAL laK[1], lnI, lcWhere, lvValue
		lcWhere = ""
		FOR lnI = 1 TO ALINES(laK, toInfo.cKey, 1, ",")
			tnN = tnN + 1
			lvValue = OLDVAL(laK[lnI], toInfo.cAlias)
			ADDPROPERTY(toParams, "p" + TRANSFORM(tnN), This.ToServerValue(lvValue, VARTYPE(lvValue), .F.))
			lcWhere = lcWhere + IIF(EMPTY(lcWhere), "", " AND ") + This.ColumnName(laK[lnI]) + " = :p" + TRANSFORM(tnN)
		ENDFOR
		RETURN lcWhere
	ENDFUNC

	PROTECTED FUNCTION IsKeyField(tcField, tcKey)
		RETURN ("," + UPPER(STRTRAN(tcKey, " ", "")) + ",") $ ("," + UPPER(tcField) + ",") OR ;
			("," + UPPER(tcField) + ",") $ ("," + UPPER(STRTRAN(tcKey, " ", "")) + ",")
	ENDFUNC

	* A value as it has to travel: blanks trimmed, empty dates as NULL (or the engine's minimum).
	PROTECTED FUNCTION ToServerValue(tvValue, tcType, tlNullable)
		DO CASE
		CASE ISNULL(tvValue)
			RETURN tvValue
		CASE INLIST(tcType, "D", "T") AND EMPTY(tvValue)
			IF This.cEmptyDateMode == "min" OR !tlNullable
				RETURN IIF(tcType == "D", This.MinDate(), DTOT(This.MinDate()))
			ENDIF
			RETURN .NULL.
		CASE INLIST(tcType, "C", "V") AND This.lTrimOnSave
			RETURN RTRIM(tvValue)
		ENDCASE
		RETURN tvValue
	ENDFUNC

	PROTECTED FUNCTION ColumnName(tcField)
		RETURN This.NaturalName(tcField)
	ENDFUNC

	* A name that comes from VFP (AFIELDS, a .dbf file name) has no case of its own: VFP gives it in
	* capitals. Where the engine keeps names as written (SQL Server), it goes in lower case.
	PROTECTED FUNCTION VfpName(tcName)
		RETURN IIF(This.cNameCase == "asis", LOWER(ALLTRIM(tcName)), This.NaturalName(tcName))
	ENDFUNC

	* D11: a name as the engine writes it when it is not quoted.
	FUNCTION NaturalName(tcName)
		DO CASE
		CASE This.cNameCase == "upper"
			RETURN UPPER(ALLTRIM(tcName))
		CASE This.cNameCase == "lower"
			RETURN LOWER(ALLTRIM(tcName))
		ENDCASE
		RETURN ALLTRIM(tcName)
	ENDFUNC

	PROTECTED FUNCTION InsertSql(tcTable, tcCols, tcVals, tcReturnKey)
		LOCAL lcSql
		IF EMPTY(tcCols)
			lcSql = "INSERT INTO " + tcTable + " DEFAULT VALUES"
		ELSE
			lcSql = "INSERT INTO " + tcTable + " (" + tcCols + ") VALUES (" + tcVals + ")"
		ENDIF
		IF !EMPTY(tcReturnKey)
			lcSql = lcSql + " RETURNING " + This.ColumnName(tcReturnKey)
		ENDIF
		RETURN lcSql
	ENDFUNC

	* ==================================================================== *
	* Metadata
	* ==================================================================== *
	FUNCTION TableExists(tcTable)
		LOCAL llFound
		IF !This.GetTables("c_fr_tables")
			RETURN .F.
		ENDIF
		SELECT c_fr_tables
		LOCATE FOR UPPER(ALLTRIM(name)) == UPPER(ALLTRIM(tcTable))
		llFound = FOUND()
		USE IN c_fr_tables
		RETURN llFound
	ENDFUNC

	FUNCTION ColumnExists(tcTable, tcColumn)
		LOCAL llFound
		IF !This.GetColumns(tcTable, "c_fr_columns")
			RETURN .F.
		ENDIF
		SELECT c_fr_columns
		LOCATE FOR UPPER(ALLTRIM(name)) == UPPER(ALLTRIM(tcColumn))
		llFound = FOUND()
		USE IN c_fr_columns
		RETURN llFound
	ENDFUNC

	* The primary key's columns, separated by commas; "" when the table has none.
	FUNCTION GetPrimaryKey(tcTable)
		LOCAL lcKey
		lcKey = ""
		IF !This.Query(This.PrimaryKeySql(This.NaturalName(tcTable)), "c_fr_pk")
			RETURN ""
		ENDIF
		SELECT c_fr_pk
		SCAN
			lcKey = lcKey + IIF(EMPTY(lcKey), "", ",") + ALLTRIM(STRTRAN(TRANSFORM(EVALUATE(FIELD(1))), CHR(0), ""))
		ENDSCAN
		USE IN c_fr_pk
		RETURN lcKey
	ENDFUNC

	* A cursor with one row per table: name.
	FUNCTION GetTables(tcCursor)
		LOCAL lcCursor, lcName
		lcCursor = IIF(EMPTY(tcCursor), "c_fr_tables", tcCursor)
		IF !This.Query(This.TablesSql(), "c_fr_tbl_raw")
			RETURN .F.
		ENDIF
		CREATE CURSOR (lcCursor) (name C(128))
		SELECT c_fr_tbl_raw
		SCAN
			lcName = ALLTRIM(STRTRAN(TRANSFORM(EVALUATE(FIELD(1))), CHR(0), ""))
			INSERT INTO (lcCursor) (name) VALUES (lcName)
		ENDSCAN
		USE IN c_fr_tbl_raw
		SELECT (lcCursor)
		GO TOP
		RETURN .T.
	ENDFUNC

	* A cursor with one row per column: name, type (VFP), width, decimals, nullable, primary_key.
	FUNCTION GetColumns(tcTable, tcCursor)
		LOCAL lcCursor, laF[1], lnI, lnN, lcKey, laNames[1], lnNames
		lcCursor = IIF(EMPTY(tcCursor), "c_fr_columns", tcCursor)
		IF !This.Query("SELECT * FROM " + tcTable + " WHERE 1 = 0", "c_fr_col_raw")
			RETURN .F.
		ENDIF
		lnN = AFIELDS(laF, "c_fr_col_raw")
		USE IN c_fr_col_raw
		lcKey = "," + UPPER(This.GetPrimaryKey(tcTable)) + ","
		lnNames = This.ServerColumnNames(tcTable, @laNames)
		CREATE CURSOR (lcCursor) (name C(128), type C(1), width I, decimals I, nullable L, primary_key L)
		FOR lnI = 1 TO lnN
			INSERT INTO (lcCursor) VALUES (IIF(lnNames = lnN, laNames[lnI], This.NaturalName(laF[lnI, 1])), ;
				laF[lnI, 2], laF[lnI, 3], laF[lnI, 4], laF[lnI, 5], ("," + UPPER(laF[lnI, 1]) + ",") $ lcKey)
		ENDFOR
		SELECT (lcCursor)
		GO TOP
		RETURN .T.
	ENDFUNC

	* The column names as the server's catalog has them, in the order of SELECT *; how many.
	* (AFIELDS() gives them in capitals, so their case is lost.)
	PROTECTED FUNCTION ServerColumnNames(tcTable, taNames)
		EXTERNAL ARRAY taNames
		LOCAL lcSql, lnN
		lcSql = This.ColumnNamesSql(This.NaturalName(tcTable))
		IF EMPTY(lcSql) OR !This.Query(lcSql, "c_fr_cn")
			RETURN 0
		ENDIF
		lnN = 0
		SELECT c_fr_cn
		SCAN
			lnN = lnN + 1
			DIMENSION taNames[lnN]
			taNames[lnN] = ALLTRIM(STRTRAN(TRANSFORM(EVALUATE(FIELD(1))), CHR(0), ""))
		ENDSCAN
		USE IN c_fr_cn
		RETURN lnN
	ENDFUNC

	FUNCTION ServerVersion
		RETURN STRTRAN(TRANSFORM(This.Scalar(This.VersionSql())), CHR(0), "")
	ENDFUNC

	FUNCTION ServerDate
		RETURN This.Scalar(This.ServerDateSql())
	ENDFUNC

	FUNCTION NewGuid
		RETURN STRTRAN(TRANSFORM(This.Scalar(This.NewGuidSql())), CHR(0), "")
	ENDFUNC

	* ==================================================================== *
	* Tables and migration
	* ==================================================================== *
	FUNCTION NewTableDef(tcName, tcDescription)
		LOCAL loT
		loT = CREATEOBJECT("RemoteTableDef")
		loT.cName = tcName
		loT.cDescription = IIF(EMPTY(tcDescription), "", tcDescription)
		RETURN loT
	ENDFUNC

	* CREATE TABLE (with its foreign keys), then CREATE INDEX for each index.
	FUNCTION CreateTable(toTableDef)
		LOCAL loIndex
		IF This.Execute(This.CreateTableSql(toTableDef)) < 0
			RETURN .F.
		ENDIF
		FOR EACH loIndex IN toTableDef.oIndexes
			IF This.Execute(This.IndexSql(toTableDef, loIndex)) < 0
				RETURN .F.
			ENDIF
		ENDFOR
		RETURN .T.
	ENDFUNC

	* The statements CreateTable() would run, separated by ";" and a line break; nothing is run.
	FUNCTION TableScript(toTableDef)
		LOCAL lcScript, loIndex
		lcScript = This.CreateTableSql(toTableDef) + ";" + FR_CRLF
		FOR EACH loIndex IN toTableDef.oIndexes
			lcScript = lcScript + This.IndexSql(toTableDef, loIndex) + ";" + FR_CRLF
		ENDFOR
		RETURN lcScript
	ENDFUNC

	FUNCTION CreateTableFromCursor(tcAlias, tcTable)
		LOCAL loT
		IF !USED(tcAlias)
			RETURN This.Fail("bad_argument", "the cursor is not open: " + TRANSFORM(tcAlias))
		ENDIF
		loT = This.TableDefFromCursor(tcAlias, IIF(EMPTY(tcTable), This.VfpName(tcAlias), tcTable))
		RETURN This.CreateTable(loT)
	ENDFUNC

	* The SQL of every table of a .dbf, a .dbc or a folder, without running it.
	FUNCTION GenerateScript(tcSource)
		LOCAL laSources[1], lnN, lnI, lcScript, lcAlias, llOpened, loT, lcOpenedDbc
		lcScript = ""
		lnN = This.SourceTables(tcSource, @laSources, @lcOpenedDbc)
		FOR lnI = 1 TO lnN
			lcAlias = This.OpenSource(laSources[lnI], @llOpened)
			IF !EMPTY(lcAlias)
				loT = This.TableDefFromCursor(lcAlias, This.VfpName(JUSTSTEM(laSources[lnI])))
				lcScript = lcScript + This.TableScript(loT)
				IF llOpened
					USE IN (lcAlias)
				ENDIF
			ENDIF
		ENDFOR
		IF !EMPTY(lcOpenedDbc)
			SET DATABASE TO (lcOpenedDbc)
			CLOSE DATABASES
		ENDIF
		RETURN lcScript
	ENDFUNC

	* Copies .dbf tables (one, a whole .dbc, or every .dbf of a folder) to the server.
	* An existing table is refused unless tlReplace. Returns a report object.
	FUNCTION Migrate(tcSource, tlReplace)
		LOCAL llStrict, loR
		llStrict = This.HoldStrict()
		TRY
			loR = This.MigrateNow(tcSource, tlReplace)
		FINALLY
			This.lStrict = llStrict
		ENDTRY
		This.Raise(loR.nErrors = 0)
		RETURN loR
	ENDFUNC

	PROTECTED FUNCTION MigrateNow(tcSource, tlReplace)
		LOCAL loR, laSources[1], lnN, lnI, lcTable, lcAlias, llOpened, loT, lnRows, lcOpenedDbc
		loR = CREATEOBJECT("Empty")
		ADDPROPERTY(loR, "nTables", 0)
		ADDPROPERTY(loR, "nRows", 0)
		ADDPROPERTY(loR, "nErrors", 0)
		ADDPROPERTY(loR, "cErrors", "")
		lnN = This.SourceTables(tcSource, @laSources, @lcOpenedDbc)
		IF lnN = 0
			This.Fail("bad_argument", "nothing to migrate in " + TRANSFORM(tcSource))
			loR.nErrors = 1
			loR.cErrors = This.cLastError
			RETURN loR
		ENDIF
		FOR lnI = 1 TO lnN
			lcTable = This.VfpName(JUSTSTEM(laSources[lnI]))
			lcAlias = This.OpenSource(laSources[lnI], @llOpened)
			DO CASE
			CASE EMPTY(lcAlias)
				This.MigrationError(loR, lcTable)
			CASE This.TableExists(lcTable) AND !tlReplace
				This.Fail("table_exists", "the table already exists on the server (Migrate with tlReplace to replace it): " + lcTable)
				This.MigrationError(loR, lcTable)
			OTHERWISE
				lnRows = This.MigrateOne(lcAlias, lcTable, tlReplace)
				IF lnRows < 0
					This.MigrationError(loR, lcTable)
				ELSE
					loR.nTables = loR.nTables + 1
					loR.nRows = loR.nRows + lnRows
				ENDIF
			ENDCASE
			IF llOpened AND USED(lcAlias)
				USE IN (lcAlias)
			ENDIF
		ENDFOR
		IF !EMPTY(lcOpenedDbc)
			SET DATABASE TO (lcOpenedDbc)
			CLOSE DATABASES
		ENDIF
		RETURN loR
	ENDFUNC

	PROTECTED PROCEDURE MigrationError(toR, tcTable)
		toR.nErrors = toR.nErrors + 1
		toR.cErrors = toR.cErrors + tcTable + ": " + This.cLastError + FR_CRLF
	ENDPROC

	* Rows copied, or -1.
	PROTECTED FUNCTION MigrateOne(tcAlias, tcTable, tlReplace)
		LOCAL loT, laF[1], lnN, lnI, lcCols, lcVals, loP, lnRows, llOk, lnSelect
		IF tlReplace AND This.TableExists(tcTable)
			IF This.Execute("DROP TABLE " + tcTable) < 0
				RETURN -1
			ENDIF
		ENDIF
		loT = This.TableDefFromCursor(tcAlias, tcTable)
		IF !This.CreateTable(loT)
			RETURN -1
		ENDIF
		lnN = AFIELDS(laF, tcAlias)
		lcCols = ""
		lcVals = ""
		FOR lnI = 1 TO lnN
			lcCols = lcCols + IIF(lnI = 1, "", ", ") + This.VfpName(laF[lnI, 1])
			lcVals = lcVals + IIF(lnI = 1, "", ", ") + ":p" + TRANSFORM(lnI)
		ENDFOR
		IF !This.BeginTransaction()
			RETURN -1
		ENDIF
		lnRows = 0
		llOk = .T.
		lnSelect = SELECT()
		SELECT (tcAlias)
		SCAN FOR !DELETED()
			loP = CREATEOBJECT("Empty")
			FOR lnI = 1 TO lnN
				ADDPROPERTY(loP, "p" + TRANSFORM(lnI), This.ToServerValue(EVALUATE(laF[lnI, 1]), laF[lnI, 2], .T.))
			ENDFOR
			IF This.Run("INSERT INTO " + tcTable + " (" + lcCols + ") VALUES (" + lcVals + ")", "", loP) < 0
				llOk = .F.
				EXIT
			ENDIF
			SELECT (tcAlias)
			lnRows = lnRows + 1
		ENDSCAN
		SELECT (lnSelect)
		IF llOk
			llOk = This.Commit()
		ELSE
			This.RollbackQuietly()
		ENDIF
		RETURN IIF(llOk, lnRows, -1)
	ENDFUNC

	* Fills taSources with the full paths of the .dbf files of a source; returns how many.
	PROTECTED FUNCTION SourceTables(tcSource, taSources, tcOpenedDbc)
		LOCAL lnN, laFiles[1], lnI, lnCount, laObj[1], lcDir, lcDbc
		EXTERNAL ARRAY taSources
		tcOpenedDbc = ""
		lnCount = 0
		DO CASE
		CASE DIRECTORY(tcSource)
			lcDir = ADDBS(tcSource)
			lnN = ADIR(laFiles, lcDir + "*.dbf")
			FOR lnI = 1 TO lnN
				lnCount = lnCount + 1
				DIMENSION taSources[lnCount]
				taSources[lnCount] = lcDir + laFiles[lnI, 1]
			ENDFOR
		CASE UPPER(JUSTEXT(tcSource)) == "DBC" AND FILE(tcSource)
			lcDbc = UPPER(JUSTSTEM(tcSource))
			IF !DBUSED(lcDbc)
				OPEN DATABASE (tcSource) SHARED NOUPDATE
				tcOpenedDbc = lcDbc
			ENDIF
			SET DATABASE TO (lcDbc)
			lnN = ADBOBJECTS(laObj, "TABLE")
			FOR lnI = 1 TO lnN
				lnCount = lnCount + 1
				DIMENSION taSources[lnCount]
				taSources[lnCount] = FULLPATH(DBGETPROP(laObj[lnI], "TABLE", "Path"), tcSource)
			ENDFOR
		CASE UPPER(JUSTEXT(tcSource)) == "DBF" AND FILE(tcSource)
			lnCount = 1
			taSources[1] = tcSource
		ENDCASE
		RETURN lnCount
	ENDFUNC

	* Opens a .dbf in its own work area (or finds it open); returns the alias, "" on failure.
	PROTECTED FUNCTION OpenSource(tcFile, tlOpened)
		LOCAL lcAlias, llOk
		tlOpened = .F.
		lcAlias = "c_fr_src_" + JUSTSTEM(tcFile)
		lcAlias = CHRTRAN(lcAlias, " -.", "___")
		llOk = .T.
		TRY
			USE (tcFile) AGAIN IN 0 ALIAS (lcAlias) SHARED NOUPDATE
		CATCH
			llOk = .F.
		ENDTRY
		IF !llOk
			This.Fail("source_not_open", "the table could not be opened: " + tcFile)
			RETURN ""
		ENDIF
		tlOpened = .T.
		RETURN lcAlias
	ENDFUNC

	PROTECTED FUNCTION TableDefFromCursor(tcAlias, tcTable)
		LOCAL loT, laF[1], lnI, loC
		loT = This.NewTableDef(This.NaturalName(tcTable))
		FOR lnI = 1 TO AFIELDS(laF, tcAlias)
			loC = loT.AddColumn(This.VfpName(laF[lnI, 1]), laF[lnI, 2], laF[lnI, 3], laF[lnI, 4])
			loC.lNullable = .T.
		ENDFOR
		RETURN loT
	ENDFUNC

	PROTECTED FUNCTION CreateTableSql(toT)
		LOCAL lcCols, loC, lcSql, lcKey, loF
		lcCols = ""
		lcKey = ""
		IF !EMPTY(This.cAutoKeyName)
			lcCols = This.NaturalName(This.cAutoKeyName) + " " + This.AutoKeyType()
		ENDIF
		FOR EACH loC IN toT.oColumns
			lcCols = lcCols + IIF(EMPTY(lcCols), "", ", ") + This.NaturalName(loC.cName) + " "
			IF loC.lAutoIncrement
				lcCols = lcCols + This.AutoKeyType()
				LOOP
			ENDIF
			lcCols = lcCols + This.ColumnType(loC.cType, loC.nWidth, loC.nDecimals)
			IF !EMPTY(loC.cDefault)
				lcCols = lcCols + " DEFAULT " + loC.cDefault
			ENDIF
			IF !loC.lNullable
				lcCols = lcCols + " NOT NULL"
			ENDIF
			IF loC.lPrimaryKey
				lcKey = lcKey + IIF(EMPTY(lcKey), "", ", ") + This.NaturalName(loC.cName)
			ENDIF
		ENDFOR
		IF !EMPTY(lcKey)
			lcCols = lcCols + ", PRIMARY KEY (" + lcKey + ")"
		ENDIF
		FOR EACH loF IN toT.oForeignKeys
			lcCols = lcCols + ", FOREIGN KEY (" + This.NameList(loF.cColumns) + ") REFERENCES " + This.NaturalName(loF.cRefTable) + ;
				" (" + This.NameList(loF.cRefColumns) + ")" + ;
				IIF(EMPTY(loF.cOnDelete), "", " ON DELETE " + This.ForeignKeyAction(loF.cOnDelete)) + ;
				IIF(EMPTY(loF.cOnUpdate), "", " ON UPDATE " + This.ForeignKeyAction(loF.cOnUpdate))
		ENDFOR
		RETURN "CREATE TABLE " + This.NaturalName(toT.cName) + " (" + lcCols + ")"
	ENDFUNC

	PROTECTED FUNCTION IndexSql(toT, toIndex)
		RETURN "CREATE " + IIF(toIndex.lUnique, "UNIQUE ", "") + "INDEX " + This.NaturalName(toIndex.cName) + ;
			" ON " + This.NaturalName(toT.cName) + " (" + This.NameList(toIndex.cColumns) + ")"
	ENDFUNC

	* "a, b" with each name as the engine writes it.
	PROTECTED FUNCTION NameList(tcNames)
		LOCAL laN[1], lnI, lcList
		lcList = ""
		FOR lnI = 1 TO ALINES(laN, tcNames, 1, ",")
			lcList = lcList + IIF(lnI = 1, "", ", ") + This.NaturalName(laN[lnI])
		ENDFOR
		RETURN lcList
	ENDFUNC

	* ON DELETE / ON UPDATE: NULL and DEFAULT are short for SET NULL and SET DEFAULT.
	PROTECTED FUNCTION ForeignKeyAction(tcAction)
		LOCAL lcAction
		lcAction = UPPER(ALLTRIM(tcAction))
		DO CASE
		CASE lcAction == "NULL"
			RETURN "SET NULL"
		CASE lcAction == "DEFAULT"
			RETURN "SET DEFAULT"
		ENDCASE
		RETURN lcAction
	ENDFUNC

	* A VFP field type as this engine writes it.
	PROTECTED FUNCTION ColumnType(tcType, tnWidth, tnDecimals)
		DO CASE
		CASE INLIST(tcType, "C", "V")
			RETURN "VARCHAR(" + TRANSFORM(MAX(tnWidth, 1)) + ")"
		CASE tcType == "M"
			RETURN This.TextType()
		CASE tcType == "N"
			RETURN "NUMERIC(" + TRANSFORM(MIN(MAX(tnWidth, 1), 38)) + ", " + TRANSFORM(tnDecimals) + ")"
		CASE INLIST(tcType, "F", "B")
			RETURN This.DoubleType()
		CASE tcType == "I"
			RETURN "INTEGER"
		CASE tcType == "Y"
			RETURN "NUMERIC(19, 4)"
		CASE tcType == "D"
			RETURN "DATE"
		CASE tcType == "T"
			RETURN This.DateTimeType()
		CASE tcType == "L"
			RETURN This.BooleanType()
		ENDCASE
		RETURN This.BlobType()
	ENDFUNC

	* ==================================================================== *
	* Engine hooks (each engine overrides what differs)
	* ==================================================================== *
	PROTECTED FUNCTION BuildConnectionString(tcDatabase)
		RETURN ""
	ENDFUNC

	* A SELECT that reads the long text columns VFP cannot type; "" when there are none.
	PROTECTED FUNCTION LongTextSelect(tcAlias, tcTable, tcWhere)
		RETURN ""
	ENDFUNC

	PROTECTED FUNCTION MaintenanceDatabase
		RETURN ""
	ENDFUNC

	PROTECTED FUNCTION DatabaseExistsSql(tcDatabase)
		RETURN ""
	ENDFUNC

	PROTECTED FUNCTION CreateDatabaseSql(tcDatabase)
		RETURN "CREATE DATABASE " + tcDatabase
	ENDFUNC

	PROTECTED FUNCTION PingSql
		RETURN "SELECT 1"
	ENDFUNC

	PROTECTED FUNCTION PrimaryKeySql(tcTable)
		RETURN ""
	ENDFUNC

	PROTECTED FUNCTION TablesSql
		RETURN ""
	ENDFUNC

	PROTECTED FUNCTION ColumnNamesSql(tcTable)
		RETURN ""
	ENDFUNC

	PROTECTED FUNCTION VersionSql
		RETURN "SELECT VERSION()"
	ENDFUNC

	PROTECTED FUNCTION ServerDateSql
		RETURN "SELECT CURRENT_TIMESTAMP"
	ENDFUNC

	PROTECTED FUNCTION NewGuidSql
		RETURN ""
	ENDFUNC

	PROTECTED FUNCTION LastIdSql
		RETURN ""
	ENDFUNC

	PROTECTED FUNCTION MinDate
		RETURN DATE(1753, 1, 1)
	ENDFUNC

	PROTECTED FUNCTION AutoKeyType
		RETURN ""
	ENDFUNC

	PROTECTED FUNCTION TextType
		RETURN "TEXT"
	ENDFUNC

	PROTECTED FUNCTION DoubleType
		RETURN "DOUBLE PRECISION"
	ENDFUNC

	PROTECTED FUNCTION DateTimeType
		RETURN "TIMESTAMP"
	ENDFUNC

	PROTECTED FUNCTION BooleanType
		RETURN "BOOLEAN"
	ENDFUNC

	PROTECTED FUNCTION BlobType
		RETURN "BLOB"
	ENDFUNC
ENDDEFINE

* ======================================================================== *
* SQL Server
* ======================================================================== *
DEFINE CLASS RemoteSqlServer AS RemoteDatabase
	cEngine = "sqlserver"
	cNameCase = "asis"

	PROTECTED FUNCTION BuildConnectionString(tcDatabase)
		LOCAL lcConn
		lcConn = "DRIVER={" + This.cDriver + "};SERVER=" + This.Value(This.cServer + IIF(This.nPort > 0, "," + TRANSFORM(This.nPort), "")) + ";"
		IF !EMPTY(tcDatabase)
			lcConn = lcConn + "DATABASE=" + This.Value(tcDatabase) + ";"
		ENDIF
		lcConn = lcConn + IIF(This.lTrustedConnection, "Trusted_Connection=yes;", This.Credentials())
		RETURN lcConn + "APP=FoxRemote;" + This.Options()
	ENDFUNC

	PROTECTED FUNCTION MaintenanceDatabase
		RETURN "master"
	ENDFUNC

	* With ODBC Driver 17/18, VFP types VARCHAR(MAX) and NVARCHAR(MAX) as C(0): whatever is written
	* there is cut to nothing (measured 04-10). Read them as TEXT, which VFP takes as a memo.
	PROTECTED FUNCTION LongTextSelect(tcAlias, tcTable, tcWhere)
		LOCAL laF[1], lnI, lnN, lcCols, llLong
		lnN = AFIELDS(laF, tcAlias)
		lcCols = ""
		llLong = .F.
		FOR lnI = 1 TO lnN
			IF laF[lnI, 2] == "C" AND laF[lnI, 3] = 0
				llLong = .T.
				lcCols = lcCols + IIF(lnI = 1, "", ", ") + "CAST(CAST(" + laF[lnI, 1] + " AS VARCHAR(MAX)) AS TEXT) AS " + laF[lnI, 1]
			ELSE
				lcCols = lcCols + IIF(lnI = 1, "", ", ") + laF[lnI, 1]
			ENDIF
		ENDFOR
		IF !llLong
			RETURN ""
		ENDIF
		RETURN "SELECT " + lcCols + " FROM " + tcTable + IIF(EMPTY(tcWhere), "", " WHERE " + tcWhere)
	ENDFUNC

	PROTECTED FUNCTION DatabaseExistsSql(tcDatabase)
		RETURN "SELECT COUNT(*) FROM sys.databases WHERE name = '" + STRTRAN(tcDatabase, "'", "''") + "'"
	ENDFUNC

	PROTECTED FUNCTION CreateDatabaseSql(tcDatabase)
		RETURN "CREATE DATABASE [" + STRTRAN(tcDatabase, "]", "]]") + "]"
	ENDFUNC

	PROTECTED FUNCTION PrimaryKeySql(tcTable)
		RETURN "SELECT c.name FROM sys.indexes i JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id " + ;
			"JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id " + ;
			"WHERE i.is_primary_key = 1 AND i.object_id = OBJECT_ID('" + STRTRAN(tcTable, "'", "''") + "') ORDER BY ic.key_ordinal"
	ENDFUNC

	PROTECTED FUNCTION TablesSql
		RETURN "SELECT CAST(TABLE_NAME AS varchar(128)) AS name FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
	ENDFUNC

	PROTECTED FUNCTION ColumnNamesSql(tcTable)
		RETURN "SELECT CAST(name AS varchar(128)) FROM sys.columns WHERE object_id = OBJECT_ID('" + STRTRAN(tcTable, "'", "''") + "') ORDER BY column_id"
	ENDFUNC

	PROTECTED FUNCTION VersionSql
		RETURN "SELECT CAST(SERVERPROPERTY('ProductVersion') AS varchar(30)) + ' ' + CAST(SERVERPROPERTY('Edition') AS varchar(60))"
	ENDFUNC

	PROTECTED FUNCTION ServerDateSql
		RETURN "SELECT GETDATE()"
	ENDFUNC

	PROTECTED FUNCTION NewGuidSql
		RETURN "SELECT CAST(NEWID() AS varchar(36))"
	ENDFUNC

	PROTECTED FUNCTION LastIdSql
		RETURN "SELECT CAST(@@IDENTITY AS int)"
	ENDFUNC

	* The new id comes from SCOPE_IDENTITY() in the same batch. OUTPUT INSERTED is refused on a table
	* with a trigger (without INTO), and @@IDENTITY is moved by a trigger that inserts elsewhere.
	PROTECTED FUNCTION InsertSql(tcTable, tcCols, tcVals, tcReturnKey)
		LOCAL lcInsert
		IF EMPTY(tcCols)
			lcInsert = "INSERT INTO " + tcTable + " DEFAULT VALUES"
		ELSE
			lcInsert = "INSERT INTO " + tcTable + " (" + tcCols + ") VALUES (" + tcVals + ")"
		ENDIF
		IF EMPTY(tcReturnKey)
			RETURN lcInsert
		ENDIF
		RETURN "SET NOCOUNT ON; " + lcInsert + "; SET NOCOUNT OFF; SELECT CAST(SCOPE_IDENTITY() AS bigint) AS " + tcReturnKey
	ENDFUNC

	PROTECTED FUNCTION AutoKeyType
		RETURN "INT IDENTITY(1,1) PRIMARY KEY"
	ENDFUNC

	* SQL Server has no RESTRICT: NO ACTION is what it does.
	PROTECTED FUNCTION ForeignKeyAction(tcAction)
		LOCAL lcAction
		lcAction = DODEFAULT(tcAction)
		RETURN IIF(lcAction == "RESTRICT", "NO ACTION", lcAction)
	ENDFUNC

	PROTECTED FUNCTION TextType
		RETURN "VARCHAR(MAX)"
	ENDFUNC

	PROTECTED FUNCTION DoubleType
		RETURN "FLOAT"
	ENDFUNC

	PROTECTED FUNCTION DateTimeType
		RETURN "DATETIME2"
	ENDFUNC

	PROTECTED FUNCTION BooleanType
		RETURN "BIT"
	ENDFUNC

	PROTECTED FUNCTION BlobType
		RETURN "VARBINARY(MAX)"
	ENDFUNC
ENDDEFINE

* ======================================================================== *
* MySQL (and MariaDB below)
* ======================================================================== *
DEFINE CLASS RemoteMySql AS RemoteDatabase
	cEngine = "mysql"
	cNameCase = "lower"
	lReturning = .F.
	lCountsChangedRows = .T.

	PROTECTED FUNCTION BuildConnectionString(tcDatabase)
		LOCAL lcConn
		lcConn = "DRIVER={" + This.cDriver + "};SERVER=" + This.Value(This.cServer) + ";"
		IF This.nPort > 0
			lcConn = lcConn + "PORT=" + TRANSFORM(INT(This.nPort)) + ";"
		ENDIF
		IF !EMPTY(tcDatabase)
			lcConn = lcConn + "DATABASE=" + This.Value(tcDatabase) + ";"
		ENDIF
		RETURN lcConn + This.Credentials() + This.Options()
	ENDFUNC

	PROTECTED FUNCTION MaintenanceDatabase
		RETURN "information_schema"
	ENDFUNC

	PROTECTED FUNCTION DatabaseExistsSql(tcDatabase)
		RETURN "SELECT COUNT(*) FROM information_schema.schemata WHERE schema_name = '" + STRTRAN(tcDatabase, "'", "''") + "'"
	ENDFUNC

	PROTECTED FUNCTION PrimaryKeySql(tcTable)
		RETURN "SELECT COLUMN_NAME FROM information_schema.KEY_COLUMN_USAGE WHERE TABLE_SCHEMA = DATABASE() " + ;
			"AND TABLE_NAME = '" + STRTRAN(tcTable, "'", "''") + "' AND CONSTRAINT_NAME = 'PRIMARY' ORDER BY ORDINAL_POSITION"
	ENDFUNC

	PROTECTED FUNCTION TablesSql
		RETURN "SELECT TABLE_NAME FROM information_schema.TABLES WHERE TABLE_SCHEMA = DATABASE() AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
	ENDFUNC

	PROTECTED FUNCTION ColumnNamesSql(tcTable)
		RETURN "SELECT COLUMN_NAME FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = '" + ;
			STRTRAN(tcTable, "'", "''") + "' ORDER BY ORDINAL_POSITION"
	ENDFUNC

	PROTECTED FUNCTION ServerDateSql
		RETURN "SELECT NOW()"
	ENDFUNC

	PROTECTED FUNCTION NewGuidSql
		RETURN "SELECT UUID()"
	ENDFUNC

	PROTECTED FUNCTION LastIdSql
		RETURN "SELECT LAST_INSERT_ID()"
	ENDFUNC

	PROTECTED FUNCTION MinDate
		RETURN DATE(1000, 1, 1)
	ENDFUNC

	PROTECTED FUNCTION AutoKeyType
		RETURN "INT AUTO_INCREMENT PRIMARY KEY"
	ENDFUNC

	PROTECTED FUNCTION TextType
		RETURN "LONGTEXT"
	ENDFUNC

	PROTECTED FUNCTION DoubleType
		RETURN "DOUBLE"
	ENDFUNC

	PROTECTED FUNCTION DateTimeType
		RETURN "DATETIME"
	ENDFUNC

	PROTECTED FUNCTION BooleanType
		RETURN "TINYINT(1)"
	ENDFUNC

	PROTECTED FUNCTION BlobType
		RETURN "LONGBLOB"
	ENDFUNC
ENDDEFINE

* MariaDB 10.5 and later: INSERT ... RETURNING.
DEFINE CLASS RemoteMariaDb AS RemoteMySql
	cEngine = "mariadb"
	lReturning = .T.
ENDDEFINE

* ======================================================================== *
* PostgreSQL
* ======================================================================== *
DEFINE CLASS RemotePostgreSql AS RemoteDatabase
	cEngine = "postgresql"
	cNameCase = "lower"

	PROTECTED FUNCTION BuildConnectionString(tcDatabase)
		LOCAL lcConn
		lcConn = "DRIVER={" + This.cDriver + "};SERVER=" + This.Value(This.cServer) + ";PORT=" + TRANSFORM(IIF(This.nPort > 0, INT(This.nPort), 5432)) + ";"
		lcConn = lcConn + "DATABASE=" + This.Value(IIF(EMPTY(tcDatabase), "postgres", tcDatabase)) + ";"
		RETURN lcConn + This.Credentials() + This.Options()
	ENDFUNC

	PROTECTED FUNCTION MaintenanceDatabase
		RETURN "postgres"
	ENDFUNC

	PROTECTED FUNCTION DatabaseExistsSql(tcDatabase)
		RETURN "SELECT COUNT(*) FROM pg_database WHERE datname = '" + STRTRAN(tcDatabase, "'", "''") + "'"
	ENDFUNC

	PROTECTED FUNCTION PrimaryKeySql(tcTable)
		RETURN "SELECT a.attname FROM pg_index i JOIN pg_attribute a ON a.attrelid = i.indrelid AND a.attnum = ANY(i.indkey) " + ;
			"WHERE i.indrelid = to_regclass('" + STRTRAN(tcTable, "'", "''") + "') AND i.indisprimary ORDER BY array_position(i.indkey, a.attnum)"
	ENDFUNC

	PROTECTED FUNCTION TablesSql
		RETURN "SELECT table_name FROM information_schema.tables WHERE table_schema = current_schema() AND table_type = 'BASE TABLE' ORDER BY table_name"
	ENDFUNC

	PROTECTED FUNCTION ColumnNamesSql(tcTable)
		RETURN "SELECT attname FROM pg_attribute WHERE attrelid = to_regclass('" + STRTRAN(tcTable, "'", "''") + "') AND attnum > 0 " + ;
			"AND NOT attisdropped ORDER BY attnum"
	ENDFUNC

	PROTECTED FUNCTION ServerDateSql
		RETURN "SELECT LOCALTIMESTAMP"
	ENDFUNC

	PROTECTED FUNCTION NewGuidSql
		RETURN "SELECT CAST(gen_random_uuid() AS varchar(36))"
	ENDFUNC

	PROTECTED FUNCTION LastIdSql
		RETURN "SELECT lastval()"
	ENDFUNC

	PROTECTED FUNCTION AutoKeyType
		RETURN "SERIAL PRIMARY KEY"
	ENDFUNC

	PROTECTED FUNCTION BlobType
		RETURN "BYTEA"
	ENDFUNC
ENDDEFINE

* ======================================================================== *
* Firebird (3 and later). cDatabase is "host/port:path" (or a path or an alias).
* ======================================================================== *
DEFINE CLASS RemoteFirebird AS RemoteDatabase
	cEngine = "firebird"
	cNameCase = "upper"
	lFileDatabase = .T.

	PROTECTED FUNCTION BuildConnectionString(tcDatabase)
		RETURN "DRIVER={" + This.cDriver + "};DBNAME=" + This.Value(tcDatabase) + ";" + This.Credentials() + "CHARSET=WIN1252;" + This.Options()
	ENDFUNC

	* Only a local path can be checked before connecting; a remote one is tried.
	PROTECTED FUNCTION FileDatabaseExists(tcDatabase)
		LOCAL lcPath
		lcPath = This.LocalPath(tcDatabase)
		RETURN EMPTY(lcPath) OR FILE(lcPath)
	ENDFUNC

	PROTECTED FUNCTION LocalPath(tcDatabase)
		LOCAL lcHost, lcPath, lnColon
		lcPath = tcDatabase
		lnColon = AT(":", tcDatabase)
		IF lnColon > 2
			lcHost = LOWER(LEFT(tcDatabase, lnColon - 1))
			IF "/" $ lcHost
				lcHost = LEFT(lcHost, AT("/", lcHost) - 1)
			ENDIF
			IF !INLIST(lcHost, "127.0.0.1", "localhost")
				RETURN ""
			ENDIF
			lcPath = SUBSTR(tcDatabase, lnColon + 1)
		ENDIF
		RETURN IIF(":" $ lcPath, lcPath, "")
	ENDFUNC

	* Firebird creates a database from a connection to another one: cDatabase has to name an
	* existing database (or alias) of the same server. A plain SQLEXEC on a short connection: with a
	* cursor and a row count the driver answers "Function sequence error" (measured 04-10).
	PROTECTED FUNCTION CreateDatabaseNow(tcDatabase)
		LOCAL lnH, lnR
		lnH = This.OpenHandle(This.BuildConnectionString(This.cDatabase))
		IF lnH < 1
			RETURN This.Fail("connection_failed", "CreateDatabase() needs cDatabase to name an existing database of the server", This.oOpenError)
		ENDIF
		This.cLastSql = "CREATE DATABASE '" + STRTRAN(tcDatabase, "'", "''") + "' USER '" + STRTRAN(This.cUser, "'", "''") + ;
			"' PASSWORD '" + STRTRAN(This.cPassword, "'", "''") + "' PAGE_SIZE 8192 DEFAULT CHARACTER SET WIN1252"
		lnR = SQLEXEC(lnH, This.cLastSql)
		IF lnR < 1
			This.Fail("query_failed", "CREATE DATABASE failed", .T.)
		ENDIF
		=SQLDISCONNECT(lnH)
		RETURN lnR > 0
	ENDFUNC

	PROTECTED FUNCTION PingSql
		RETURN "SELECT 1 FROM RDB$DATABASE"
	ENDFUNC

	PROTECTED FUNCTION PrimaryKeySql(tcTable)
		RETURN "SELECT TRIM(s.RDB$FIELD_NAME) FROM RDB$RELATION_CONSTRAINTS c JOIN RDB$INDEX_SEGMENTS s ON s.RDB$INDEX_NAME = c.RDB$INDEX_NAME " + ;
			"WHERE c.RDB$RELATION_NAME = '" + STRTRAN(tcTable, "'", "''") + "' AND c.RDB$CONSTRAINT_TYPE = 'PRIMARY KEY' ORDER BY s.RDB$FIELD_POSITION"
	ENDFUNC

	PROTECTED FUNCTION TablesSql
		RETURN "SELECT TRIM(RDB$RELATION_NAME) FROM RDB$RELATIONS WHERE COALESCE(RDB$SYSTEM_FLAG, 0) = 0 AND RDB$VIEW_BLR IS NULL ORDER BY 1"
	ENDFUNC

	PROTECTED FUNCTION ColumnNamesSql(tcTable)
		RETURN "SELECT TRIM(RDB$FIELD_NAME) FROM RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = '" + STRTRAN(tcTable, "'", "''") + ;
			"' ORDER BY RDB$FIELD_POSITION"
	ENDFUNC

	PROTECTED FUNCTION VersionSql
		RETURN "SELECT RDB$GET_CONTEXT('SYSTEM', 'ENGINE_VERSION') FROM RDB$DATABASE"
	ENDFUNC

	PROTECTED FUNCTION ServerDateSql
		RETURN "SELECT LOCALTIMESTAMP FROM RDB$DATABASE"
	ENDFUNC

	PROTECTED FUNCTION NewGuidSql
		RETURN "SELECT UUID_TO_CHAR(GEN_UUID()) FROM RDB$DATABASE"
	ENDFUNC

	* Firebird has no "last id" of the session: the identity generator of the last INSERT's table.
	PROTECTED FUNCTION LastIdSql
		LOCAL lcGen
		lcGen = STRTRAN(TRANSFORM(This.Scalar("SELECT TRIM(RDB$GENERATOR_NAME) FROM RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = '" + ;
			UPPER(STRTRAN(This.cLastInsertTable, "'", "''")) + "' AND RDB$IDENTITY_TYPE IS NOT NULL")), CHR(0), "")
		IF EMPTY(lcGen) OR lcGen == ".NULL."
			RETURN "SELECT CAST(NULL AS INTEGER) FROM RDB$DATABASE"
		ENDIF
		RETURN "SELECT GEN_ID(" + lcGen + ", 0) FROM RDB$DATABASE"
	ENDFUNC

	PROTECTED FUNCTION InsertSql(tcTable, tcCols, tcVals, tcReturnKey)
		LOCAL lcSql
		IF EMPTY(tcCols)
			lcSql = "INSERT INTO " + tcTable + " DEFAULT VALUES"
		ELSE
			lcSql = "INSERT INTO " + tcTable + " (" + tcCols + ") VALUES (" + tcVals + ")"
		ENDIF
		RETURN lcSql + IIF(EMPTY(tcReturnKey), "", " RETURNING " + This.ColumnName(tcReturnKey))
	ENDFUNC

	PROTECTED FUNCTION AutoKeyType
		RETURN "INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY"
	ENDFUNC

	PROTECTED FUNCTION TextType
		RETURN "BLOB SUB_TYPE TEXT"
	ENDFUNC
ENDDEFINE

* ======================================================================== *
* SQLite (3.35 and later). cDatabase is the path of the file.
* ======================================================================== *
DEFINE CLASS RemoteSqlite AS RemoteDatabase
	cEngine = "sqlite"
	cNameCase = "lower"
	lFileDatabase = .T.
	lUtf8Data = .T.

	* NoWCHAR=1: without it, computed columns arrive as UCS-2. FKSupport=1: SQLite enforces foreign
	* keys only when each connection asks for it.
	PROTECTED FUNCTION BuildConnectionString(tcDatabase)
		RETURN "DRIVER={" + This.cDriver + "};DATABASE=" + This.FileName(tcDatabase) + ";NoWCHAR=1;FKSupport=1;" + This.Options()
	ENDFUNC

	* This driver does not read DATABASE={...}: a path with ';' (or '{', '}', '=') goes as a
	* file: URI, percent-encoded, which it hands to SQLite as it is (measured 04-10).
	PROTECTED FUNCTION FileName(tcDatabase)
		LOCAL lcUri
		IF !(";" $ tcDatabase OR "{" $ tcDatabase OR "}" $ tcDatabase OR "=" $ tcDatabase)
			RETURN tcDatabase
		ENDIF
		lcUri = STRTRAN(tcDatabase, "%", "%25")
		lcUri = STRTRAN(lcUri, ";", "%3B")
		lcUri = STRTRAN(lcUri, "#", "%23")
		lcUri = STRTRAN(lcUri, "?", "%3F")
		lcUri = STRTRAN(lcUri, " ", "%20")
		lcUri = STRTRAN(lcUri, "{", "%7B")
		lcUri = STRTRAN(lcUri, "}", "%7D")
		lcUri = STRTRAN(lcUri, "=", "%3D")
		RETURN "file:" + STRTRAN(lcUri, "\", "/")
	ENDFUNC

	PROTECTED FUNCTION FileDatabaseExists(tcDatabase)
		RETURN tcDatabase == ":memory:" OR FILE(tcDatabase)
	ENDFUNC

	* The driver creates the file on connecting.
	PROTECTED FUNCTION CreateDatabaseNow(tcDatabase)
		LOCAL lnH
		lnH = This.OpenHandle(This.BuildConnectionString(tcDatabase))
		IF lnH < 1
			RETURN This.Fail("connection_failed", "the database could not be created", This.oOpenError)
		ENDIF
		=SQLEXEC(lnH, "PRAGMA user_version = 0")
		=SQLDISCONNECT(lnH)
		RETURN FILE(tcDatabase)
	ENDFUNC

	PROTECTED FUNCTION PrimaryKeySql(tcTable)
		RETURN "SELECT name FROM pragma_table_info('" + STRTRAN(tcTable, "'", "''") + "') WHERE pk > 0 ORDER BY pk"
	ENDFUNC

	PROTECTED FUNCTION TablesSql
		RETURN "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
	ENDFUNC

	PROTECTED FUNCTION ColumnNamesSql(tcTable)
		RETURN "SELECT name FROM pragma_table_info('" + STRTRAN(tcTable, "'", "''") + "') ORDER BY cid"
	ENDFUNC

	PROTECTED FUNCTION VersionSql
		RETURN "SELECT sqlite_version()"
	ENDFUNC

	PROTECTED FUNCTION ServerDateSql
		RETURN "SELECT datetime('now', 'localtime')"
	ENDFUNC

	PROTECTED FUNCTION NewGuidSql
		RETURN "SELECT lower(hex(randomblob(4)) || '-' || hex(randomblob(2)) || '-' || hex(randomblob(2)) || '-' || hex(randomblob(2)) || '-' || hex(randomblob(6)))"
	ENDFUNC

	PROTECTED FUNCTION LastIdSql
		RETURN "SELECT last_insert_rowid()"
	ENDFUNC

	PROTECTED FUNCTION AutoKeyType
		RETURN "INTEGER PRIMARY KEY AUTOINCREMENT"
	ENDFUNC

	PROTECTED FUNCTION DoubleType
		RETURN "REAL"
	ENDFUNC

	PROTECTED FUNCTION DateTimeType
		RETURN "DATETIME"
	ENDFUNC

	PROTECTED FUNCTION BooleanType
		RETURN "INTEGER"
	ENDFUNC
ENDDEFINE

* ======================================================================== *
* Table definitions for CreateTable()
* ======================================================================== *
DEFINE CLASS RemoteTableDef AS Custom
	cName = ""
	cDescription = ""
	oColumns = .NULL.
	oIndexes = .NULL.
	oForeignKeys = .NULL.

	PROCEDURE Init
		This.oColumns = CREATEOBJECT("Collection")
		This.oIndexes = CREATEOBJECT("Collection")
		This.oForeignKeys = CREATEOBJECT("Collection")
	ENDPROC

	* tcColumns: "a" or "a, b".
	FUNCTION AddIndex(tcName, tcColumns, tlUnique)
		LOCAL loI
		loI = CREATEOBJECT("Empty")
		ADDPROPERTY(loI, "cName", tcName)
		ADDPROPERTY(loI, "cColumns", tcColumns)
		ADDPROPERTY(loI, "lUnique", tlUnique)
		This.oIndexes.Add(loI)
		RETURN loI
	ENDFUNC

	* tcOnDelete / tcOnUpdate: CASCADE, SET NULL (or NULL), SET DEFAULT (or DEFAULT), RESTRICT, NO ACTION.
	FUNCTION AddForeignKey(tcColumns, tcRefTable, tcRefColumns, tcOnDelete, tcOnUpdate)
		LOCAL loF
		loF = CREATEOBJECT("Empty")
		ADDPROPERTY(loF, "cColumns", tcColumns)
		ADDPROPERTY(loF, "cRefTable", tcRefTable)
		ADDPROPERTY(loF, "cRefColumns", tcRefColumns)
		ADDPROPERTY(loF, "cOnDelete", IIF(EMPTY(tcOnDelete), "", tcOnDelete))
		ADDPROPERTY(loF, "cOnUpdate", IIF(EMPTY(tcOnUpdate), "", tcOnUpdate))
		This.oForeignKeys.Add(loF)
		RETURN loF
	ENDFUNC

	* Type is a VFP field type letter (C, V, M, N, I, B, F, Y, D, T, L, ...).
	FUNCTION AddColumn(tcName, tcType, tnWidth, tnDecimals)
		LOCAL loC
		loC = CREATEOBJECT("RemoteColumnDef")
		loC.cName = tcName
		loC.cType = UPPER(LEFT(tcType, 1))
		loC.nWidth = IIF(VARTYPE(tnWidth) == "N", tnWidth, 0)
		loC.nDecimals = IIF(VARTYPE(tnDecimals) == "N", tnDecimals, 0)
		This.oColumns.Add(loC)
		RETURN loC
	ENDFUNC
ENDDEFINE

DEFINE CLASS RemoteColumnDef AS Custom
	cName = ""
	cType = "C"
	nWidth = 0
	nDecimals = 0
	lNullable = .T.
	lPrimaryKey = .F.
	lAutoIncrement = .F.
	cDefault = ""
ENDDEFINE
