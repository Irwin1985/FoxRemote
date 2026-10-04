* FoxRemoteIntegrationTests.prg - FoxRemote 1.0 against real database servers.
*
* Every test runs on each engine of FOXREMOTE_TEST_ENGINES (TestSupport.prg) and checks the
* result through a raw ODBC connection of its own, not through the library. A failure names
* the engine. Each test is one measured fault of 0.x (RELEVO-FOXREMOTE-20261003.md section 10)
* or one behaviour the 1.0 promises (PROPUESTA-FOXREMOTE-1.0.md).
*
* Run (servers started, tests\setup-test-engines.ps1 done once):
*   foxproof run --prg tests\FoxRemoteIntegrationTests.prg
* Red run against 0.x: set FOXREMOTE_LIB=<home>\tests\legacy\LegacyShim.prg first.

DEFINE CLASS TestRemoteIntegration AS Custom
	cEngine = ""
	cFails = ""

	PROCEDURE SetUp
		SET PROCEDURE TO "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\tests\TestSupport.prg" ADDITIVE
		IF !EMPTY(GETENV("FOXREMOTE_HOME"))
			SET PROCEDURE TO (ADDBS(GETENV("FOXREMOTE_HOME")) + "tests\TestSupport.prg") ADDITIVE
		ENDIF
		FrLoadLibrary()
		* Harness hygiene: a failed connection must not open the ODBC login dialog in the hidden
		* VFP of the runner. The one test about dialogs sets DispLogin back to 1 on purpose.
		=SQLSETPROP(0, "DispLogin", 3)
		=SQLSETPROP(0, "ConnectTimeOut", 15)
		This.cFails = ""
	ENDPROC

	PROCEDURE TearDown
		CLOSE TABLES ALL
		SET MULTILOCKS OFF
	ENDPROC

	* ---- harness -----------------------------------------------------------------
	* Runs This.<tcMethod>(loEngine) on every engine and asserts that nothing failed.
	PROCEDURE OnEachEngine(tcMethod, tcOnly)
		LOCAL laNames[1], lnI, loE, lcName
		=ALINES(laNames, FrEngineList(), 1, ",")
		FOR lnI = 1 TO ALEN(laNames)
			lcName = ALLTRIM(laNames[lnI])
			IF EMPTY(lcName) OR (!EMPTY(tcOnly) AND !(("," + lcName + ",") $ ("," + tcOnly + ",")))
				LOOP
			ENDIF
			This.cEngine = lcName
			loE = CREATEOBJECT("FrEngine", lcName)
			TRY
				This.&tcMethod.(loE)
			CATCH TO loEx
				This.Check(.F., "raised " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message + " (" + loEx.Procedure + ", line " + TRANSFORM(loEx.LineNo) + ")")
			ENDTRY
			CLOSE TABLES ALL
		ENDFOR
		__assert.True(EMPTY(This.cFails), This.cFails)
	ENDPROC

	PROCEDURE Check(tlOk, tcWhat)
		IF !tlOk
			This.cFails = This.cFails + "[" + This.cEngine + "] " + tcWhat + CHR(13) + CHR(10)
		ENDIF
	ENDPROC

	* SQL Server's ODBC driver opens a dropped idle connection again by itself (ConnectRetryCount):
	* off here, so the reconnection tests see what the library does.
	FUNCTION ConnectedWithoutDriverRetry(toE)
		LOCAL loDb
		loDb = toE.NewDb()
		IF toE.cName == "mssql"
			loDb.cConnectionOptions = loDb.cConnectionOptions + "ConnectRetryCount=0;"
		ENDIF
		This.Check(loDb.Connect(), "Connect() to the test database failed: " + TRANSFORM(loDb.cLastError))
		RETURN loDb
	ENDFUNC

	FUNCTION Connected(toE)
		LOCAL loDb
		loDb = toE.NewDb()
		This.Check(loDb.Connect(), "Connect() to the test database failed: " + TRANSFORM(loDb.cLastError))
		RETURN loDb
	ENDFUNC

	* ======================================================================== *
	* Connection
	* ======================================================================== *
	PROCEDURE TestConnectDoesNotCreateAMissingDatabase() HELP [Fact, Trait("Category", "Connection"), Timeout(180000)]
		This.OnEachEngine("ConnectDoesNotCreateAMissingDatabase")
	ENDPROC

	PROCEDURE ConnectDoesNotCreateAMissingDatabase(toE)
		LOCAL lcMissing, loDb, llOk
		lcMissing = toE.NewDatabaseName()
		loDb = toE.NewDb(lcMissing)
		llOk = loDb.Connect()
		loDb.Disconnect()
		This.Check(!llOk, "Connect() to a missing database returned .T.")
		This.Check(loDb.cLastErrorCode == "database_not_found", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected database_not_found")
		This.Check(!toE.DatabaseExists(lcMissing), "the missing database was created")
		toE.DropDatabase(lcMissing)
	ENDPROC

	PROCEDURE TestCreateDatabaseThenConnect() HELP [Fact, Trait("Category", "Connection"), Timeout(180000)]
		This.OnEachEngine("CreateDatabaseThenConnect")
	ENDPROC

	PROCEDURE CreateDatabaseThenConnect(toE)
		LOCAL lcNew, loDb
		lcNew = toE.NewDatabaseName()
		loDb = toE.NewDb(toE.MaintenanceDatabase())
		This.Check(loDb.CreateDatabase(toE.DatabaseRef(lcNew)), "CreateDatabase() failed: " + TRANSFORM(loDb.cLastError))
		loDb.Disconnect()
		This.Check(toE.DatabaseExists(lcNew), "the database does not exist after CreateDatabase()")
		loDb = toE.NewDb(lcNew)
		This.Check(loDb.Connect(), "Connect() to the new database failed: " + TRANSFORM(loDb.cLastError))
		loDb.Disconnect()
		toE.DropDatabase(lcNew)
	ENDPROC

	PROCEDURE TestConnectLandsInTheRequestedDatabase() HELP [Fact, Trait("Category", "Connection"), Timeout(180000)]
		This.OnEachEngine("ConnectLandsInTheRequestedDatabase")
	ENDPROC

	PROCEDURE ConnectLandsInTheRequestedDatabase(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		This.Check(loDb.TableExists(lcT), "TableExists() does not see a table of the requested database")
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestConnectWithASemicolonInThePassword() HELP [Fact, Trait("Category", "Connection"), Timeout(180000)]
		This.OnEachEngine("ConnectWithASemicolonInThePassword")
	ENDPROC

	PROCEDURE ConnectWithASemicolonInThePassword(toE)
		LOCAL loDb, lcDir, lcFile, lcT
		IF toE.cName == "sqlite"
			* SQLite has no password: the ';' goes in the path of the database file.
			lcT = toE.NewName()
			toE.CreateCustomers(lcT)
			lcDir = toE.cDataDir + "semi;colon" + SUBSTR(SYS(2015), 2)
			MKDIR (lcDir)
			lcFile = ADDBS(lcDir) + "fr.db"
			COPY FILE (toE.cDatabase) TO (lcFile)
			loDb = toE.NewDb(lcFile)
			This.Check(loDb.Connect() AND loDb.TableExists(lcT), "a database path with ';' did not reach the file: " + TRANSFORM(loDb.cLastError))
			loDb.Disconnect()
			toE.DropTable(lcT)
			TRY
				ERASE (lcFile)
				RMDIR (lcDir)
			CATCH
			ENDTRY
			RETURN
		ENDIF
		loDb = toE.NewDb()
		loDb.cUser = toE.cSemicolonUser
		loDb.cPassword = toE.cSemicolonPassword
		This.Check(";" $ loDb.cPassword, "fixture: the password of " + loDb.cUser + " has no ';'")
		This.Check(loDb.Connect(), "Connect() as " + loDb.cUser + " (password with ';') failed: " + TRANSFORM(loDb.cLastError))
		loDb.Disconnect()
	ENDPROC

	PROCEDURE TestConnectWithWindowsAuthentication() HELP [Fact, Trait("Category", "Connection"), Timeout(60000)]
		This.OnEachEngine("ConnectWithWindowsAuthentication", "mssql")
	ENDPROC

	PROCEDURE ConnectWithWindowsAuthentication(toE)
		LOCAL loDb, lvWho
		loDb = toE.NewDb()
		loDb.cUser = ""
		loDb.cPassword = ""
		loDb.lTrustedConnection = .T.
		This.Check(loDb.Connect(), "Connect() with lTrustedConnection failed: " + TRANSFORM(loDb.cLastError))
		lvWho = loDb.Scalar("SELECT SUSER_SNAME()")
		This.Check(VARTYPE(lvWho) == "C" AND "\" $ lvWho, "connected as " + TRANSFORM(lvWho) + ", expected a Windows account")
		loDb.Disconnect()
	ENDPROC

	PROCEDURE TestConnectDoesNotChangeGlobalSettings() HELP [Fact, Trait("Category", "Connection"), Timeout(180000)]
		This.OnEachEngine("ConnectDoesNotChangeGlobalSettings")
	ENDPROC

	PROCEDURE ConnectDoesNotChangeGlobalSettings(toE)
		LOCAL loDb
		SET MULTILOCKS OFF
		SET DATE AMERICAN
		loDb = This.Connected(toE)
		This.Check(SET("MULTILOCKS") == "OFF", "SET MULTILOCKS changed by Connect()")
		This.Check(SET("DATE") == "AMERICAN", "SET DATE changed by Connect()")
		loDb.Disconnect()
	ENDPROC

	PROCEDURE TestConnectFailureIsReportedWithoutADialog() HELP [Fact, Trait("Category", "Connection"), Timeout(60000)]
		LOCAL lnOld
		* With DispLogin 1 (VFP's default) a failed connection asks for the login in a dialog,
		* unless the library sets DispLogin 3 on its handle. A dialog here ends in the timeout.
		lnOld = SQLGETPROP(0, "DispLogin")
		=SQLSETPROP(0, "DispLogin", 1)
		This.OnEachEngine("ConnectFailureIsReportedWithoutADialog", "mssql,mariadb,mysql,postgresql,firebird")
		=SQLSETPROP(0, "DispLogin", lnOld)
	ENDPROC

	PROCEDURE ConnectFailureIsReportedWithoutADialog(toE)
		LOCAL loDb, llOk
		loDb = toE.NewDb()
		loDb.cPassword = "wrong" + toE.cPassword
		llOk = loDb.Connect()
		This.Check(!llOk, "Connect() with a wrong password returned .T.")
		This.Check(loDb.nErrors >= 1, "nErrors = " + TRANSFORM(loDb.nErrors) + " after a failed Connect()")
		This.Check(loDb.cLastErrorCode == "connection_failed", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected connection_failed")
		* The driver's reason has to reach cLastError: a wrong password, a refused certificate...
		This.Check(LEN(loDb.cLastError) > LEN("connection failed: ") AND !EMPTY(loDb.cLastSqlState), "the failure says no reason: '" + loDb.cLastError + "', SQLSTATE '" + loDb.cLastSqlState + "'")
		loDb.Disconnect()
	ENDPROC

	* ======================================================================== *
	* Errors
	* ======================================================================== *
	PROCEDURE TestErrorsAreCountedWithCodeAndStatement() HELP [Fact, Trait("Category", "Errors"), Timeout(180000)]
		This.OnEachEngine("ErrorsAreCountedWithCodeAndStatement")
	ENDPROC

	PROCEDURE ErrorsAreCountedWithCodeAndStatement(toE)
		LOCAL loDb, lnR
		loDb = This.Connected(toE)
		loDb.ClearErrors()
		lnR = loDb.Execute("SELECT * FROM no_such_table_fr")
		This.Check(lnR = -1, "Execute() of a bad statement returned " + TRANSFORM(lnR) + ", expected -1")
		This.Check(loDb.nErrors = 1, "nErrors = " + TRANSFORM(loDb.nErrors) + ", expected 1")
		This.Check(loDb.cLastErrorCode == "query_failed", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected query_failed")
		This.Check("no_such_table_fr" $ LOWER(loDb.cLastSql), "cLastSql = '" + loDb.cLastSql + "'")
		This.Check(!EMPTY(loDb.cLastError) AND !(CHR(13) $ loDb.cLastError), "cLastError should be one non-empty line: " + loDb.cLastError)
		loDb.Disconnect()
	ENDPROC

	PROCEDURE TestStrictModeThrows() HELP [Fact, Trait("Category", "Errors"), Timeout(180000)]
		This.OnEachEngine("StrictModeThrows")
	ENDPROC

	PROCEDURE StrictModeThrows(toE)
		LOCAL loDb, llThrown
		loDb = This.Connected(toE)
		loDb.lStrict = .T.
		llThrown = .F.
		TRY
			loDb.Execute("SELECT * FROM no_such_table_fr")
		CATCH
			llThrown = .T.
		ENDTRY
		This.Check(llThrown, "with lStrict, a bad statement did not throw")
		loDb.lStrict = .F.
		loDb.Disconnect()
	ENDPROC

	* ======================================================================== *
	* Queries
	* ======================================================================== *
	PROCEDURE TestQueryWithNamedParametersFromLocals() HELP [Fact, Trait("Category", "Query"), Timeout(180000)]
		This.OnEachEngine("QueryWithNamedParametersFromLocals")
	ENDPROC

	PROCEDURE QueryWithNamedParametersFromLocals(toE)
		LOCAL lcT, loDb, loP, lcCity, llOk
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "Ana", "Madrid")
		toE.InsertCustomer(lcT, "Luis", "Lima")
		loDb = This.Connected(toE)
		lcCity = "Madrid"
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "city", lcCity)
		llOk = loDb.Query("SELECT name FROM " + lcT + " WHERE city = :city", "c_frq", loP)
		This.Check(llOk AND USED("c_frq"), "Query() with :city failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frq")
			This.Check(RECCOUNT("c_frq") = 1 AND ALLTRIM(c_frq.name) == "Ana", "Query() returned " + TRANSFORM(RECCOUNT("c_frq")) + " rows")
			USE IN c_frq
		ENDIF
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestScalarReturnsNumbers() HELP [Fact, Trait("Category", "Query"), Timeout(180000)]
		This.OnEachEngine("ScalarReturnsNumbers")
	ENDPROC

	PROCEDURE ScalarReturnsNumbers(toE)
		LOCAL lcT, loDb, lvN
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "Ana", "Madrid")
		loDb = This.Connected(toE)
		lvN = loDb.Scalar("SELECT COUNT(*) FROM " + lcT)
		This.Check(VARTYPE(lvN) == "N" AND lvN = 1, "Scalar(COUNT(*)) = " + TRANSFORM(lvN) + " (type " + VARTYPE(lvN) + "), expected the number 1")
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestExecuteReturnsAffectedRows() HELP [Fact, Trait("Category", "Query"), Timeout(180000)]
		This.OnEachEngine("ExecuteReturnsAffectedRows")
	ENDPROC

	PROCEDURE ExecuteReturnsAffectedRows(toE)
		LOCAL lcT, loDb, lnR
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "Ana", "Madrid")
		toE.InsertCustomer(lcT, "Luis", "Madrid")
		loDb = This.Connected(toE)
		lnR = loDb.Execute("UPDATE " + lcT + " SET city = 'Rome' WHERE city = 'Madrid'")
		This.Check(lnR = 2, "Execute(UPDATE of 2 rows) returned " + TRANSFORM(lnR))
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	* ======================================================================== *
	* Transactions
	* ======================================================================== *
	PROCEDURE TestTransactionCommitPersists() HELP [Fact, Trait("Category", "Transactions"), Timeout(180000)]
		This.OnEachEngine("TransactionCommitPersists")
	ENDPROC

	PROCEDURE TransactionCommitPersists(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		This.Check(loDb.BeginTransaction(), "BeginTransaction() failed")
		loDb.Execute("INSERT INTO " + lcT + " (name) VALUES ('tx_commit')")
		This.Check(loDb.Commit(), "Commit() failed: " + TRANSFORM(loDb.cLastError))
		This.Check(toE.RawCount(lcT, "name = 'tx_commit'") = 1, "another connection does not see the committed row")
		This.Check(SQLGETPROP(loDb.nHandle, "Transactions") = 1, "the handle stayed in manual transactions after Commit()")
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestTransactionRollbackUndoes() HELP [Fact, Trait("Category", "Transactions"), Timeout(180000)]
		This.OnEachEngine("TransactionRollbackUndoes")
	ENDPROC

	PROCEDURE TransactionRollbackUndoes(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		loDb.BeginTransaction()
		loDb.Execute("INSERT INTO " + lcT + " (name) VALUES ('tx_rollback')")
		This.Check(loDb.Rollback(), "Rollback() failed: " + TRANSFORM(loDb.cLastError))
		This.Check(!loDb.lInTransaction, "lInTransaction still .T. after Rollback()")
		loDb.Disconnect()
		This.Check(toE.RawCount(lcT, "name = 'tx_rollback'") = 0, "the rolled back row is on the server")
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestDisconnectWithAnOpenTransactionRollsBack() HELP [Fact, Trait("Category", "Transactions"), Timeout(180000)]
		This.OnEachEngine("DisconnectWithAnOpenTransactionRollsBack")
	ENDPROC

	PROCEDURE DisconnectWithAnOpenTransactionRollsBack(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		loDb.BeginTransaction()
		loDb.Execute("INSERT INTO " + lcT + " (name) VALUES ('tx_open')")
		loDb.Disconnect()
		This.Check(loDb.cLastErrorCode == "rolled_back_on_disconnect", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected rolled_back_on_disconnect")
		This.Check(toE.RawCount(lcT, "name = 'tx_open'") = 0, "the row of the open transaction is on the server")
		* When the library did not roll back, the connection may still be alive on the server with
		* its transaction open (0.x on Firebird, measured 04-10): a DROP TABLE would wait forever.
		* The table is left behind instead.
		IF loDb.cLastErrorCode == "rolled_back_on_disconnect"
			toE.DropTable(lcT)
		ENDIF
	ENDPROC

	* ======================================================================== *
	* Updatable cursors and Save
	* ======================================================================== *
	PROCEDURE TestSaveInsertIsCommittedAndReturnsTheId() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("SaveInsertIsCommittedAndReturnsTheId")
	ENDPROC

	PROCEDURE SaveInsertIsCommittedAndReturnsTheId(toE)
		LOCAL lcT, loDb, llOk, lnId
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "Ana", "Madrid")
		loDb = This.Connected(toE)
		This.Check(loDb.Open(lcT, "c_frs", "1 = 0"), "Open() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frs")
			SELECT c_frs
			APPEND BLANK
			REPLACE name WITH "inserted", city WITH "Rome"
			llOk = loDb.Save("c_frs")
			This.Check(llOk, "Save() of an insert failed: " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcT, "name = 'inserted'") = 1, "another connection does not see the inserted row before Disconnect()")
			lnId = toE.MaxId(lcT)
			This.Check(c_frs.id = lnId AND lnId > 1, "the cursor's id is " + TRANSFORM(c_frs.id) + ", the server gave " + TRANSFORM(lnId))
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestSaveInsertIntoATableWithATrigger() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("SaveInsertIntoATableWithATrigger", "mssql")
	ENDPROC

	* SQL Server refuses OUTPUT without INTO on a table with a trigger, and a trigger that inserts
	* into another table with an identity moves @@IDENTITY: the id has to be the row's own.
	PROCEDURE SaveInsertIntoATableWithATrigger(toE)
		LOCAL lcT, lcAudit, loDb, llOk, lnId
		lcT = toE.NewName()
		lcAudit = toE.NewName("au")
		toE.CreateCustomers(lcT)
		toE.RawRun("CREATE TABLE " + lcAudit + " (n INT IDENTITY(1000,1) PRIMARY KEY, what VARCHAR(20))")
		toE.RawRun("CREATE TRIGGER tr_" + lcT + " ON " + lcT + " AFTER INSERT AS BEGIN SET NOCOUNT ON; INSERT INTO " + lcAudit + " (what) VALUES ('insert') END")
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_frtr", "1 = 0")
		IF USED("c_frtr")
			SELECT c_frtr
			APPEND BLANK
			REPLACE name WITH "with trigger"
			llOk = loDb.Save("c_frtr")
			This.Check(llOk, "Save() into a table with a trigger failed: " + TRANSFORM(loDb.cLastError))
			lnId = toE.MaxId(lcT)
			This.Check(c_frtr.id = lnId, "the cursor's id is " + TRANSFORM(c_frtr.id) + ", the row has " + TRANSFORM(lnId))
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
		toE.DropTable(lcAudit)
	ENDPROC

	PROCEDURE TestSaveUpdatesOnlyTheChangedColumns() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("SaveUpdatesOnlyTheChangedColumns")
	ENDPROC

	PROCEDURE SaveUpdatesOnlyTheChangedColumns(toE)
		LOCAL lcT, loDb, llOk, lvRow
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "old", "Oslo")
		loDb = This.Connected(toE)
		This.Check(loDb.Open(lcT, "c_fru"), "Open() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_fru")
			* Another connection changes qty after the cursor was read: Save() must not send it back.
			toE.RawRun("UPDATE " + lcT + " SET qty = 7")
			SELECT c_fru
			REPLACE name WITH "new", city WITH "Lima"
			llOk = loDb.Save("c_fru")
			This.Check(llOk, "Save() of an update failed: " + TRANSFORM(loDb.cLastError))
			lvRow = toE.RawScalar("SELECT RTRIM(name) " + IIF(toE.cName == "mssql", "+ '/' + RTRIM(city)", "|| '/' || RTRIM(city)") + " FROM " + lcT)
			IF INLIST(toE.cName, "mysql", "mariadb")
				lvRow = toE.RawScalar("SELECT CONCAT(RTRIM(name), '/', RTRIM(city)) FROM " + lcT)
			ENDIF
			* (VFP pads a VARCHAR(n) result to n in the raw cursor: compare trimmed)
			This.Check(RTRIM(TRANSFORM(lvRow)) == "new/Lima", "on the server: " + TRANSFORM(lvRow) + ", expected new/Lima")
			This.Check(toE.RawCount(lcT, "qty = 7") = 1, "Save() overwrote a column it did not change (qty)")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestSaveOfAValueWrittenAgainSucceeds() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("SaveOfAValueWrittenAgainSucceeds")
	ENDPROC

	* A form writes every field back on saving, changed or not. MySQL and MariaDB count the rows an
	* UPDATE changed, not the rows it found: the row is there, and Save() must not say it is gone.
	PROCEDURE SaveOfAValueWrittenAgainSucceeds(toE)
		LOCAL lcT, loDb, llOk
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "same", "Oslo")
		loDb = This.Connected(toE)
		This.Check(loDb.Open(lcT, "c_frw"), "Open() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frw")
			SELECT c_frw
			REPLACE name WITH "same"
			llOk = loDb.Save("c_frw")
			This.Check(llOk, "Save() of a value written again failed: " + loDb.cLastErrorCode + " " + TRANSFORM(loDb.cLastError))
			This.Check(!loDb.lInTransaction, "Save() left a transaction open")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestSaveDeletesTheRow() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("SaveDeletesTheRow")
	ENDPROC

	PROCEDURE SaveDeletesTheRow(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "gone", "Oslo")
		toE.InsertCustomer(lcT, "stays", "Oslo")
		loDb = This.Connected(toE)
		This.Check(loDb.Open(lcT, "c_frd"), "Open() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frd")
			SELECT c_frd
			DELETE FOR ALLTRIM(name) == "gone"
			This.Check(loDb.Save("c_frd"), "Save() of a delete failed: " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcT) = 1 AND toE.RawCount(lcT, "name = 'stays'") = 1, "the server has " + TRANSFORM(toE.RawCount(lcT)) + " rows, expected only 'stays'")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestSaveFailureReturnsFalseAndKeepsTheEdits() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("SaveFailureReturnsFalseAndKeepsTheEdits")
	ENDPROC

	PROCEDURE SaveFailureReturnsFalseAndKeepsTheEdits(toE)
		LOCAL lcT, loDb, llOk
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		This.Check(loDb.Open(lcT, "c_frf", "1 = 0"), "Open() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frf")
			SELECT c_frf
			APPEND BLANK
			REPLACE name WITH "good row", qty WITH 1
			APPEND BLANK
			REPLACE name WITH "bad row", qty WITH -1
			llOk = loDb.Save("c_frf")
			This.Check(!llOk, "Save() returned .T. although the server refused a row (CHECK qty >= 0)")
			This.Check(loDb.cLastErrorCode == "save_failed", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected save_failed")
			This.Check(toE.RawCount(lcT) = 0, "the server kept " + TRANSFORM(toE.RawCount(lcT)) + " rows of a failed Save()")
			This.Check(GETNEXTMODIFIED(0, "c_frf") # 0, "the user's edits were thrown away")
			This.Check(SQLGETPROP(loDb.nHandle, "Transactions") = 1, "the handle stayed in manual transactions")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestSaveOfAnUnknownAliasLeavesNoTransactionOpen() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("SaveOfAnUnknownAliasLeavesNoTransactionOpen")
	ENDPROC

	PROCEDURE SaveOfAnUnknownAliasLeavesNoTransactionOpen(toE)
		LOCAL lcT, loDb, llOk
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		llOk = loDb.Save("not_opened_fr")
		This.Check(!llOk, "Save() of an alias it never opened returned .T.")
		This.Check(loDb.cLastErrorCode == "not_open", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected not_open")
		loDb.Execute("INSERT INTO " + lcT + " (name) VALUES ('after_bad_save')")
		loDb.Disconnect()
		This.Check(toE.RawCount(lcT, "name = 'after_bad_save'") = 1, "an INSERT run after that Save() was lost at Disconnect()")
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestSaveOfAReadOnlyCursorKeepsTheSettings() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("SaveOfAReadOnlyCursorKeepsTheSettings")
	ENDPROC

	PROCEDURE SaveOfAReadOnlyCursorKeepsTheSettings(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		SET DATE AMERICAN
		This.Check(loDb.Open(lcT, "c_frr", "", .NULL., "", .T.), "Open() read-only failed: " + TRANSFORM(loDb.cLastError))
		This.Check(!loDb.Save("c_frr"), "Save() of a read-only cursor returned .T.")
		This.Check(loDb.cLastErrorCode == "read_only", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected read_only")
		This.Check(SET("DATE") == "AMERICAN", "SET DATE is " + SET("DATE") + " after Save()")
		This.Check(SQLGETPROP(loDb.nHandle, "Transactions") = 1, "the handle stayed in manual transactions")
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestSaveGroupSendsEveryCursor() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("SaveGroupSendsEveryCursor")
	ENDPROC

	PROCEDURE SaveGroupSendsEveryCursor(toE)
		LOCAL lcT1, lcT2, loDb
		lcT1 = toE.NewName()
		lcT2 = toE.NewName()
		toE.CreateCustomers(lcT1)
		toE.CreateCustomers(lcT2)
		toE.InsertCustomer(lcT1, "one", "Bonn")
		toE.InsertCustomer(lcT2, "two", "Bonn")
		loDb = This.Connected(toE)
		This.Check(loDb.Open(lcT1, "c_frg1", "", .NULL., "", .F., "grp"), "Open() 1 failed: " + TRANSFORM(loDb.cLastError))
		This.Check(loDb.Open(lcT2, "c_frg2", "", .NULL., "", .F., "grp"), "Open() 2 failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frg1") AND USED("c_frg2")
			REPLACE city WITH "Kiel" IN c_frg1
			REPLACE city WITH "Kiel" IN c_frg2
			This.Check(loDb.SaveGroup("grp"), "SaveGroup() failed: " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcT1, "city = 'Kiel'") = 1 AND toE.RawCount(lcT2, "city = 'Kiel'") = 1, "SaveGroup() did not reach the server")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT1)
		toE.DropTable(lcT2)
	ENDPROC

	PROCEDURE TestOpenForWritingWithoutAPrimaryKeyIsRefused() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("OpenForWritingWithoutAPrimaryKeyIsRefused")
	ENDPROC

	PROCEDURE OpenForWritingWithoutAPrimaryKeyIsRefused(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.RawRun("CREATE TABLE " + lcT + " (code VARCHAR(10), descr VARCHAR(40))")
		loDb = This.Connected(toE)
		This.Check(!loDb.Open(lcT, "c_frn"), "Open() for writing a table without a primary key returned .T.")
		This.Check(loDb.cLastErrorCode == "no_primary_key", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected no_primary_key")
		This.Check(loDb.Open(lcT, "c_frn2", "", .NULL., "", .T.), "Open() read-only of a table without a primary key failed: " + TRANSFORM(loDb.cLastError))
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestCloseGroupThenReopenTheSameAlias() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("CloseGroupThenReopenTheSameAlias")
	ENDPROC

	PROCEDURE CloseGroupThenReopenTheSameAlias(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_frc", "", .NULL., "", .F., "grp")
		loDb.CloseGroup("grp")
		This.Check(!USED("c_frc"), "CloseGroup() left the cursor open")
		This.Check(loDb.Open(lcT, "c_frc", "", .NULL., "", .F., "grp"), "reopening the alias after CloseGroup() failed: " + TRANSFORM(loDb.cLastError))
		loDb.Close("c_frc")
		This.Check(loDb.Open(lcT, "c_frc"), "reopening the alias after Close() failed: " + TRANSFORM(loDb.cLastError))
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestRefreshReloadsAReadOnlyCursor() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("RefreshReloadsAReadOnlyCursor")
	ENDPROC

	PROCEDURE RefreshReloadsAReadOnlyCursor(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "first", "Graz")
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_frh", "", .NULL., "", .T.)
		toE.InsertCustomer(lcT, "second", "Graz")
		This.Check(loDb.Refresh("c_frh"), "Refresh() failed: " + TRANSFORM(loDb.cLastError))
		This.Check(USED("c_frh") AND RECCOUNT("c_frh") = 2, "after Refresh() the cursor has " + IIF(USED("c_frh"), TRANSFORM(RECCOUNT("c_frh")), "no") + " rows, expected 2")
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestRefreshRefusesWithPendingChanges() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("RefreshRefusesWithPendingChanges")
	ENDPROC

	PROCEDURE RefreshRefusesWithPendingChanges(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "row", "Graz")
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_frp")
		IF USED("c_frp")
			REPLACE city WITH "Linz" IN c_frp
			This.Check(!loDb.Refresh("c_frp"), "Refresh() with a pending change returned .T.")
			This.Check(loDb.cLastErrorCode == "pending_changes", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected pending_changes")
			This.Check(ALLTRIM(c_frp.city) == "Linz", "the pending change was lost")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestCloseAllLeavesTheUsersCursorsOpen() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("CloseAllLeavesTheUsersCursorsOpen")
	ENDPROC

	PROCEDURE CloseAllLeavesTheUsersCursorsOpen(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		CREATE CURSOR c_fr_mine (x I)
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_frall")
		loDb.CloseAll()
		This.Check(USED("c_fr_mine"), "CloseAll() closed a cursor that is not FoxRemote's")
		This.Check(!USED("c_frall"), "CloseAll() left FoxRemote's cursor open")
		USE IN SELECT("c_fr_mine")
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	* ======================================================================== *
	* Text, types and dates
	* ======================================================================== *
	PROCEDURE TestAccentedTextRoundTrip() HELP [Fact, Trait("Category", "Types"), Timeout(180000)]
		This.OnEachEngine("AccentedTextRoundTrip")
	ENDPROC

	PROCEDURE AccentedTextRoundTrip(toE)
		LOCAL lcT, loDb, lcTxt, lvRaw, lnId
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		lcTxt = "Pe" + CHR(241) + "a " + CHR(193) + "lvaro " + CHR(128) + "5"
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_fra", "1 = 0")
		IF USED("c_fra")
			SELECT c_fra
			APPEND BLANK
			REPLACE name WITH lcTxt
			This.Check(loDb.Save("c_fra"), "Save() failed: " + TRANSFORM(loDb.cLastError))
		ENDIF
		loDb.CloseAll()
		lnId = toE.MaxId(lcT)
		IF toE.cName == "sqlite"
			* SQLite stores UTF-8: any other tool has to read the same text.
			lvRaw = toE.RawScalar("SELECT hex(name) FROM " + lcT + " WHERE id = " + TRANSFORM(lnId))
			This.Check(UPPER(TRANSFORM(lvRaw)) == FrUtf8Hex(lcTxt), "SQLite holds hex " + TRANSFORM(lvRaw) + ", expected the UTF-8 " + FrUtf8Hex(lcTxt))
		ELSE
			lvRaw = toE.RawScalar("SELECT RTRIM(name) FROM " + lcT + " WHERE id = " + TRANSFORM(lnId))
			This.Check(RTRIM(TRANSFORM(lvRaw)) == lcTxt, "another connection reads hex " + STRCONV(TRANSFORM(lvRaw), 15) + ", sent " + STRCONV(lcTxt, 15))
		ENDIF
		loDb.Open(lcT, "c_fra2", "id = " + TRANSFORM(lnId))
		This.Check(USED("c_fra2") AND ALLTRIM(c_fra2.name) == lcTxt, "FoxRemote reads it back as hex " + IIF(USED("c_fra2"), STRCONV(ALLTRIM(c_fra2.name), 15), "<no cursor>"))
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestAccentedTextInTheSqlRoundTrip() HELP [Fact, Trait("Category", "Types"), Timeout(180000)]
		This.OnEachEngine("AccentedTextInTheSqlRoundTrip")
	ENDPROC

	* The accented text written in the statement itself, not in a parameter.
	PROCEDURE AccentedTextInTheSqlRoundTrip(toE)
		LOCAL lcT, loDb, lcTxt, lvCount, lvName
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		lcTxt = "Pe" + CHR(241) + "a " + CHR(193) + "lvaro"
		loDb = This.Connected(toE)
		This.Check(loDb.Execute("INSERT INTO " + lcT + " (name, city) VALUES ('" + lcTxt + "', 'x')") = 1, "INSERT failed: " + TRANSFORM(loDb.cLastError))
		lvCount = loDb.Scalar("SELECT COUNT(*) FROM " + lcT + " WHERE name = '" + lcTxt + "'")
		This.Check(lvCount = 1, "a WHERE with the same text finds " + TRANSFORM(lvCount) + " rows")
		lvName = loDb.Scalar("SELECT name FROM " + lcT + " WHERE city = 'x'")
		This.Check(RTRIM(TRANSFORM(lvName)) == lcTxt, "read back as hex " + STRCONV(RTRIM(TRANSFORM(lvName)), 15) + ", sent " + STRCONV(lcTxt, 15))
		IF toE.cName == "sqlite"
			* SQLite stores UTF-8: any other tool has to read the same text.
			lvName = toE.RawScalar("SELECT hex(name) FROM " + lcT + " WHERE city = 'x'")
			This.Check(UPPER(TRANSFORM(lvName)) == FrUtf8Hex(lcTxt), "SQLite holds hex " + TRANSFORM(lvName) + ", expected the UTF-8 " + FrUtf8Hex(lcTxt))
		ENDIF
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestSaveDoesNotPadVarchar() HELP [Fact, Trait("Category", "Types"), Timeout(180000)]
		This.OnEachEngine("SaveDoesNotPadVarchar")
	ENDPROC

	PROCEDURE SaveDoesNotPadVarchar(toE)
		LOCAL lcT, loDb, lnLen
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_frv", "1 = 0")
		IF USED("c_frv")
			SELECT c_frv
			APPEND BLANK
			REPLACE name WITH "short"
			loDb.Save("c_frv")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		lnLen = FrNum(toE.RawScalar("SELECT " + toE.LengthSql("name") + " FROM " + lcT))
		This.Check(lnLen = 5, "the server stores " + TRANSFORM(lnLen) + " characters for 'short'")
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestLongTextRoundTrip() HELP [Fact, Trait("Category", "Types"), Timeout(180000)]
		This.OnEachEngine("LongTextRoundTrip")
	ENDPROC

	PROCEDURE LongTextRoundTrip(toE)
		LOCAL lcT, loDb, lcMemo, lnLen
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		lcMemo = REPLICATE("x", 300) + "end"
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_frm", "1 = 0")
		IF USED("c_frm")
			SELECT c_frm
			APPEND BLANK
			REPLACE name WITH "memo", notes WITH lcMemo
			This.Check(loDb.Save("c_frm"), "Save() failed: " + TRANSFORM(loDb.cLastError))
		ENDIF
		loDb.CloseAll()
		lnLen = FrNum(toE.RawScalar("SELECT " + toE.LengthSql("notes") + " FROM " + lcT))
		This.Check(lnLen = 303, "the server stores " + TRANSFORM(lnLen) + " characters of a 303-character text")
		loDb.Open(lcT, "c_frm2")
		This.Check(USED("c_frm2") AND c_frm2.notes == lcMemo, "the long text did not come back whole")
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestEmptyDateIsStoredAsNull() HELP [Fact, Trait("Category", "Types"), Timeout(180000)]
		This.OnEachEngine("EmptyDateIsStoredAsNull")
	ENDPROC

	PROCEDURE EmptyDateIsStoredAsNull(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_fre", "1 = 0")
		IF USED("c_fre")
			SELECT c_fre
			APPEND BLANK
			REPLACE name WITH "no date", born WITH {}
			This.Check(loDb.Save("c_fre"), "Save() failed: " + TRANSFORM(loDb.cLastError))
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		This.Check(toE.RawCount(lcT, "born IS NULL") = 1, "an empty date did not arrive as NULL: " + TRANSFORM(toE.RawScalar("SELECT born FROM " + lcT)))
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestServerTextIsNotUcs2() HELP [Fact, Trait("Category", "Types"), Timeout(180000)]
		This.OnEachEngine("ServerTextIsNotUcs2")
	ENDPROC

	PROCEDURE ServerTextIsNotUcs2(toE)
		LOCAL loDb, lcVer
		loDb = This.Connected(toE)
		lcVer = TRANSFORM(loDb.ServerVersion())
		This.Check(!EMPTY(lcVer) AND !(CHR(0) $ lcVer), "ServerVersion() = '" + STRTRAN(lcVer, CHR(0), "<NUL>") + "'")
		loDb.Disconnect()
	ENDPROC

	* ======================================================================== *
	* Metadata
	* ======================================================================== *
	PROCEDURE TestGetPrimaryKey() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("GetPrimaryKeyOfATable")
	ENDPROC

	PROCEDURE GetPrimaryKeyOfATable(toE)
		LOCAL lcT, loDb, lcPk
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		lcPk = TRANSFORM(loDb.GetPrimaryKey(lcT))
		This.Check(LOWER(ALLTRIM(lcPk)) == "id", "GetPrimaryKey() = '" + lcPk + "', expected id")
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestGetTablesReturnsACursor() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("GetTablesReturnsACursor")
	ENDPROC

	PROCEDURE GetTablesReturnsACursor(toE)
		LOCAL lcT1, lcT2, loDb, lnFound
		lcT1 = toE.NewName()
		lcT2 = toE.NewName()
		toE.CreateCustomers(lcT1)
		toE.CreateCustomers(lcT2)
		loDb = This.Connected(toE)
		This.Check(loDb.GetTables("c_frt"), "GetTables() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frt")
			SELECT c_frt
			COUNT FOR INLIST(LOWER(ALLTRIM(name)), lcT1, lcT2) TO lnFound
			This.Check(lnFound = 2, "GetTables() lists " + TRANSFORM(lnFound) + " of the 2 new tables")
			USE IN c_frt
		ENDIF
		loDb.Disconnect()
		toE.DropTable(lcT1)
		toE.DropTable(lcT2)
	ENDPROC

	PROCEDURE TestGetColumnsReturnsACursor() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("GetColumnsReturnsACursor")
	ENDPROC

	PROCEDURE GetColumnsReturnsACursor(toE)
		LOCAL lcT, loDb, lnPk
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		This.Check(loDb.GetColumns(lcT, "c_frcol"), "GetColumns() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frcol")
			SELECT c_frcol
			This.Check(RECCOUNT() = 6, "GetColumns() lists " + TRANSFORM(RECCOUNT()) + " columns, expected 6")
			LOCATE FOR LOWER(ALLTRIM(name)) == "id"
			This.Check(FOUND() AND primary_key, "the id column is not marked as primary_key")
			LOCATE FOR LOWER(ALLTRIM(name)) == "born"
			This.Check(FOUND() AND nullable, "the born column is not marked nullable")
			USE IN c_frcol
		ENDIF
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestGetColumnsKeepsTheServersNames() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("GetColumnsKeepsTheServersNames")
	ENDPROC

	* AFIELDS() gives every name in capitals: the names have to come from the server.
	PROCEDURE GetColumnsKeepsTheServersNames(toE)
		LOCAL lcT, loDb, lcNames, lcExpected
		lcT = toE.NewName()
		toE.RawRun("CREATE TABLE " + lcT + " (Id INTEGER PRIMARY KEY, CustomerName VARCHAR(20))")
		DO CASE
		CASE toE.cName == "postgresql"
			lcExpected = "id,customername"
		CASE toE.cName == "firebird"
			lcExpected = "ID,CUSTOMERNAME"
		OTHERWISE
			lcExpected = "Id,CustomerName"
		ENDCASE
		loDb = This.Connected(toE)
		lcNames = ""
		IF loDb.GetColumns(lcT, "c_frcn")
			SELECT c_frcn
			SCAN
				lcNames = lcNames + IIF(EMPTY(lcNames), "", ",") + ALLTRIM(c_frcn.name)
			ENDSCAN
			USE IN c_frcn
		ENDIF
		This.Check(lcNames == lcExpected, "GetColumns() names: " + lcNames + ", the server has " + lcExpected)
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestLastIdAfterExecute() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("LastIdAfterExecute")
	ENDPROC

	PROCEDURE LastIdAfterExecute(toE)
		LOCAL lcT, loDb, lvId, loP
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "first", "Oslo")
		loDb = This.Connected(toE)
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "name", "second")
		loDb.Execute("INSERT INTO " + lcT + " (name) VALUES (:name)", loP)
		lvId = loDb.LastId()
		This.Check(VARTYPE(lvId) == "N" AND lvId = toE.MaxId(lcT), "LastId() = " + TRANSFORM(lvId) + " (type " + VARTYPE(lvId) + "), the server gave " + TRANSFORM(toE.MaxId(lcT)))
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	* ======================================================================== *
	* Migration
	* ======================================================================== *
	PROCEDURE TestMigrateRefusesAnExistingTable() HELP [Fact, Trait("Category", "Migration"), Timeout(180000)]
		This.OnEachEngine("MigrateRefusesAnExistingTable")
	ENDPROC

	PROCEDURE MigrateRefusesAnExistingTable(toE)
		LOCAL lcT, lcDbf, loDb, loR
		lcT = toE.NewName()
		lcDbf = toE.cDataDir + lcT + ".dbf"
		CREATE TABLE (lcDbf) FREE (name C(20))
		INSERT INTO (lcT) VALUES ("from dbf")
		USE IN (lcT)
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "keepme", "Oslo")
		loDb = This.Connected(toE)
		loR = loDb.Migrate(lcDbf)
		This.Check(loDb.cLastErrorCode == "table_exists", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected table_exists")
		This.Check(toE.RawCount(lcT, "name = 'keepme'") = 1, "Migrate() destroyed the table that was on the server")
		loDb.Disconnect()
		toE.DropTable(lcT)
		ERASE (lcDbf)
	ENDPROC

	PROCEDURE TestMigrateReplaceCopiesRowsWithUsableNames() HELP [Fact, Trait("Category", "Migration"), Timeout(180000)]
		This.OnEachEngine("MigrateReplaceCopiesRowsWithUsableNames")
	ENDPROC

	PROCEDURE MigrateReplaceCopiesRowsWithUsableNames(toE)
		LOCAL lcT, lcDbf, loDb, loR
		lcT = toE.NewName()
		lcDbf = toE.cDataDir + lcT + ".dbf"
		CREATE TABLE (lcDbf) FREE (name C(20), born D NULL, amount N(10, 2), notes M)
		INSERT INTO (lcT) VALUES ("Luis", {^2021-03-04}, 12.5, "first note")
		INSERT INTO (lcT) VALUES ("Empty date", {}, 0, "")
		USE IN (lcT)
		loDb = This.Connected(toE)
		loR = loDb.Migrate(lcDbf, .T.)
		This.Check(VARTYPE(loR) == "O", "Migrate() did not return a report object")
		IF VARTYPE(loR) == "O"
			This.Check(loR.nTables = 1 AND loR.nRows = 2 AND loR.nErrors = 0, "report: nTables " + TRANSFORM(loR.nTables) + ", nRows " + TRANSFORM(loR.nRows) + ", nErrors " + TRANSFORM(loR.nErrors))
		ENDIF
		* Plain, unquoted names have to work afterwards (no "NAME" in PostgreSQL, no "fr..." in Firebird).
		This.Check(toE.RawCount(lcT) = 2, "the server table has " + TRANSFORM(toE.RawCount(lcT)) + " rows, expected 2")
		This.Check(toE.RawCount(lcT, "amount > 0 AND name = 'Luis'") = 1, "a plain SELECT with the column names fails: " + toE.cLastError)
		This.Check(toE.RawCount(lcT, "born IS NULL") = 1, "the empty DBF date did not arrive as NULL")
		loDb.Disconnect()
		toE.DropTable(lcT)
		ERASE (lcDbf)
		ERASE (FORCEEXT(lcDbf, "fpt"))
	ENDPROC

	PROCEDURE TestMigrateOfAFolderKeepsTheDefaultDirectory() HELP [Fact, Trait("Category", "Migration"), Timeout(180000)]
		This.OnEachEngine("MigrateOfAFolderKeepsTheDefaultDirectory")
	ENDPROC

	PROCEDURE MigrateOfAFolderKeepsTheDefaultDirectory(toE)
		LOCAL lcT, lcDir, lcBefore, loDb
		lcT = toE.NewName()
		lcDir = toE.cDataDir + toE.NewName("dir")
		MKDIR (lcDir)
		CREATE TABLE (ADDBS(lcDir) + lcT) FREE (code C(5))
		USE IN (lcT)
		lcBefore = SYS(5) + SYS(2003)
		loDb = This.Connected(toE)
		loDb.Migrate(lcDir, .T.)
		This.Check(SYS(5) + SYS(2003) == lcBefore, "the default directory moved to " + SYS(5) + SYS(2003))
		CD (lcBefore)
		loDb.Disconnect()
		toE.DropTable(lcT)
		ERASE (ADDBS(lcDir) + lcT + ".dbf")
		RMDIR (lcDir)
	ENDPROC

	PROCEDURE TestCreateTableFromCursorKeepsCharacterColumns() HELP [Fact, Trait("Category", "Migration"), Timeout(180000)]
		This.OnEachEngine("CreateTableFromCursorKeepsCharacterColumns")
	ENDPROC

	PROCEDURE CreateTableFromCursorKeepsCharacterColumns(toE)
		LOCAL lcT, loDb, laF[1], lnAt, lnI
		lcT = toE.NewName()
		CREATE CURSOR c_frsrc (code C(10), descr V(40))
		loDb = This.Connected(toE)
		This.Check(loDb.CreateTableFromCursor("c_frsrc", lcT), "CreateTableFromCursor() failed: " + TRANSFORM(loDb.cLastError))
		USE IN c_frsrc
		This.Check(loDb.Open(lcT, "c_frct", "", .NULL., "", .T.), "Open() of the new table failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frct")
			lnAt = 0
			FOR lnI = 1 TO AFIELDS(laF, "c_frct")
				IF UPPER(laF[lnI, 1]) == "CODE"
					lnAt = lnI
				ENDIF
			ENDFOR
			This.Check(lnAt > 0 AND INLIST(laF[lnAt, 2], "C", "V"), "a C(10) column comes back as type " + IIF(lnAt > 0, laF[lnAt, 2], "<missing>"))
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestCreateTableWithIndexAndForeignKey() HELP [Fact, Trait("Category", "Migration"), Timeout(180000)]
		This.OnEachEngine("CreateTableWithIndexAndForeignKey")
	ENDPROC

	PROCEDURE CreateTableWithIndexAndForeignKey(toE)
		LOCAL lcParent, lcChild, loDb, loT, loC, lnParent
		lcParent = toE.NewName("pa")
		lcChild = toE.NewName("ch")
		loDb = This.Connected(toE)
		loT = loDb.NewTableDef(lcParent)
		loC = loT.AddColumn("id", "I")
		loC.lAutoIncrement = .T.
		loT.AddColumn("name", "C", 40)
		loT.AddIndex(lcParent + "_name", "name", .T.)
		This.Check(loDb.CreateTable(loT), "CreateTable() of the parent failed: " + TRANSFORM(loDb.cLastError))
		loT = loDb.NewTableDef(lcChild)
		loC = loT.AddColumn("id", "I")
		loC.lAutoIncrement = .T.
		loC = loT.AddColumn("parent_id", "I")
		loC.lNullable = .F.
		loT.AddIndex(lcChild + "_parent", "parent_id")
		loT.AddForeignKey("parent_id", lcParent, "id", "CASCADE")
		This.Check(loDb.CreateTable(loT), "CreateTable() of the child failed: " + TRANSFORM(loDb.cLastError))
		loDb.Disconnect()
		This.Check(toE.RawRun("INSERT INTO " + lcParent + " (name) VALUES ('one')"), "inserting a parent failed: " + toE.cLastError)
		This.Check(!toE.RawRun("INSERT INTO " + lcParent + " (name) VALUES ('one')"), "the unique index let a repeated name in")
		lnParent = toE.MaxId(lcParent)
		This.Check(toE.RawRun("INSERT INTO " + lcChild + " (parent_id) VALUES (" + TRANSFORM(lnParent) + ")"), "inserting a child failed: " + toE.cLastError)
		This.Check(!toE.RawRun("INSERT INTO " + lcChild + " (parent_id) VALUES (" + TRANSFORM(lnParent + 1000) + ")"), "the foreign key let in a child without a parent")
		toE.RawRun("DELETE FROM " + lcParent)
		This.Check(toE.RawCount(lcChild) = 0, "ON DELETE CASCADE did not remove the child")
		toE.DropTable(lcChild)
		toE.DropTable(lcParent)
	ENDPROC

	* ======================================================================== *
	* Headless
	* ======================================================================== *
	* ---- reconnection ------------------------------------------------------------
	PROCEDURE TestQueryAfterTheServerDroppedTheConnection() HELP [Fact, Trait("Category", "Reconnect"), Timeout(180000)]
		This.OnEachEngine("QueryAfterTheServerDroppedTheConnection", "mssql,mariadb,mysql,postgresql,firebird")
	ENDPROC

	* A read that finds the connection gone is run again on a new one: it cannot have changed anything.
	PROCEDURE QueryAfterTheServerDroppedTheConnection(toE)
		LOCAL lcT, loDb, lnSession, lvCount
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "a", "Oslo")
		loDb = This.ConnectedWithoutDriverRetry(toE)
		loDb.lAutoReconnect = .T.
		loDb.nIdleCheckSeconds = 3600
		lnSession = FrNum(loDb.Scalar(toE.SessionIdSql()))
		This.Check(toE.KillSession(lnSession), "fixture: the session could not be ended: " + toE.cLastError)
		loDb.ClearErrors()
		lvCount = loDb.Scalar("SELECT COUNT(*) FROM " + lcT)
		This.Check(FrNum(lvCount) = 1, "the query after the drop returned " + TRANSFORM(lvCount) + ": " + loDb.cLastErrorCode + " " + loDb.cLastError)
		This.Check(loDb.nErrors = 0, "a reconnection that worked counted an error: " + loDb.cLastError)
		This.Check(loDb.nReconnects = 1, "nReconnects = " + TRANSFORM(loDb.nReconnects))
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	PROCEDURE TestWriteAfterAnIdleConnectionWasDropped() HELP [Fact, Trait("Category", "Reconnect"), Timeout(180000)]
		This.OnEachEngine("WriteAfterAnIdleConnectionWasDropped", "mssql,mariadb,mysql,postgresql,firebird")
	ENDPROC

	* With nIdleCheckSeconds = 0 the connection is checked before every statement, so a write finds the
	* dropped connection before it is sent, and goes on a new one.
	PROCEDURE WriteAfterAnIdleConnectionWasDropped(toE)
		LOCAL lcT, loDb, lnSession, lnRows
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.ConnectedWithoutDriverRetry(toE)
		loDb.lAutoReconnect = .T.
		loDb.nIdleCheckSeconds = 0
		lnSession = FrNum(loDb.Scalar(toE.SessionIdSql()))
		This.Check(toE.KillSession(lnSession), "fixture: the session could not be ended: " + toE.cLastError)
		lnRows = loDb.Execute("INSERT INTO " + lcT + " (name, city) VALUES ('b', 'Lima')")
		This.Check(lnRows = 1, "the write after the drop returned " + TRANSFORM(lnRows) + ": " + loDb.cLastErrorCode + " " + loDb.cLastError)
		This.Check(toE.RawCount(lcT, "name = 'b'") = 1, "the row is not on the server")
		This.Check(loDb.nReconnects = 1, "nReconnects = " + TRANSFORM(loDb.nReconnects))
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	PROCEDURE TestWriteWhenTheConnectionDropsIsNotRepeated() HELP [Fact, Trait("Category", "Reconnect"), Timeout(180000)]
		This.OnEachEngine("WriteWhenTheConnectionDropsIsNotRepeated", "mssql,mariadb,mysql,postgresql,firebird")
	ENDPROC

	* A write that finds the connection gone is not sent again (the library cannot know whether the
	* server ran it), but the connection is open again for the next statement.
	PROCEDURE WriteWhenTheConnectionDropsIsNotRepeated(toE)
		LOCAL lcT, loDb, lnSession, lnRows
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.ConnectedWithoutDriverRetry(toE)
		loDb.lAutoReconnect = .T.
		loDb.nIdleCheckSeconds = 3600
		lnSession = FrNum(loDb.Scalar(toE.SessionIdSql()))
		This.Check(toE.KillSession(lnSession), "fixture: the session could not be ended: " + toE.cLastError)
		lnRows = loDb.Execute("INSERT INTO " + lcT + " (name, city) VALUES ('b', 'Lima')")
		This.Check(lnRows = -1 AND loDb.cLastErrorCode == "connection_lost", "the write on a dropped connection: " + TRANSFORM(lnRows) + " " + loDb.cLastErrorCode + " " + loDb.cLastError)
		This.Check(toE.RawCount(lcT, "name = 'b'") = 0, "the write was sent again")
		lnRows = loDb.Execute("INSERT INTO " + lcT + " (name, city) VALUES ('c', 'Lima')")
		This.Check(lnRows = 1, "the next write did not go: " + loDb.cLastErrorCode + " " + loDb.cLastError)
		This.Check(loDb.nReconnects = 1, "nReconnects = " + TRANSFORM(loDb.nReconnects))
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	PROCEDURE TestConnectionLostInsideATransaction() HELP [Fact, Trait("Category", "Reconnect"), Timeout(180000)]
		This.OnEachEngine("ConnectionLostInsideATransaction", "mssql,mariadb,mysql,postgresql,firebird")
	ENDPROC

	* The server rolls back the transaction of a dropped connection: the library says so, ends it on its
	* side too, and repeats nothing.
	PROCEDURE ConnectionLostInsideATransaction(toE)
		LOCAL lcT, loDb, lnSession, lnRows
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.ConnectedWithoutDriverRetry(toE)
		loDb.lAutoReconnect = .T.
		lnSession = FrNum(loDb.Scalar(toE.SessionIdSql()))
		loDb.BeginTransaction()
		This.Check(loDb.Execute("INSERT INTO " + lcT + " (name, city) VALUES ('b', 'Lima')") = 1, "fixture: the first INSERT failed: " + loDb.cLastError)
		This.Check(toE.KillSession(lnSession), "fixture: the session could not be ended: " + toE.cLastError)
		lnRows = loDb.Execute("INSERT INTO " + lcT + " (name, city) VALUES ('c', 'Lima')")
		This.Check(lnRows = -1 AND loDb.cLastErrorCode == "connection_lost", "a write in a lost transaction: " + TRANSFORM(lnRows) + " " + loDb.cLastErrorCode + " " + loDb.cLastError)
		This.Check(!loDb.lInTransaction AND loDb.nTransactionLevel = 0, "the library still thinks the transaction is open")
		This.Check(toE.RawCount(lcT, "name IN ('b', 'c')") = 0, "rows of the lost transaction are on the server")
		This.Check(loDb.Execute("INSERT INTO " + lcT + " (name, city) VALUES ('d', 'Lima')") = 1, "the next write did not go: " + loDb.cLastError)
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	PROCEDURE TestConnectionLostWithoutAutoReconnect() HELP [Fact, Trait("Category", "Reconnect"), Timeout(180000)]
		This.OnEachEngine("ConnectionLostWithoutAutoReconnect", "mssql,mariadb,mysql,postgresql,firebird")
	ENDPROC

	* Without lAutoReconnect a dropped connection is reported as such, not as a failed query.
	PROCEDURE ConnectionLostWithoutAutoReconnect(toE)
		LOCAL lcT, loDb, lnSession, lvCount
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.ConnectedWithoutDriverRetry(toE)
		lnSession = FrNum(loDb.Scalar(toE.SessionIdSql()))
		This.Check(toE.KillSession(lnSession), "fixture: the session could not be ended: " + toE.cLastError)
		lvCount = loDb.Scalar("SELECT COUNT(*) FROM " + lcT)
		This.Check(ISNULL(lvCount) AND loDb.cLastErrorCode == "connection_lost", "a query on a dropped connection: " + TRANSFORM(lvCount) + " " + loDb.cLastErrorCode + " " + loDb.cLastError)
		This.Check(loDb.nHandle = 0, "nHandle still points at the dead connection")
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	* ---- conflicts -----------------------------------------------------------------
	PROCEDURE TestSaveDetectsAConflict() HELP [Fact, Trait("Category", "Conflicts"), Timeout(180000)]
		This.OnEachEngine("SaveDetectsAConflict")
	ENDPROC

	* With lDetectConflicts, a column another user changed after the cursor was read is not overwritten.
	PROCEDURE SaveDetectsAConflict(toE)
		LOCAL lcT, loDb, llOk, lvCity
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "old", "Oslo")
		loDb = This.Connected(toE)
		loDb.lDetectConflicts = .T.
		This.Check(loDb.Open(lcT, "c_frc"), "Open() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frc")
			toE.RawRun("UPDATE " + lcT + " SET city = 'Bergen'")
			SELECT c_frc
			REPLACE city WITH "Lima"
			llOk = loDb.Save("c_frc")
			This.Check(!llOk AND loDb.cLastErrorCode == "row_changed", "Save() over another user's change: " + TRANSFORM(llOk) + " " + loDb.cLastErrorCode + " " + TRANSFORM(loDb.cLastError))
			lvCity = toE.RawScalar("SELECT RTRIM(city) FROM " + lcT)
			This.Check(RTRIM(TRANSFORM(lvCity)) == "Bergen", "the server has " + TRANSFORM(lvCity) + ", expected Bergen")
			This.Check(ALLTRIM(c_frc.city) == "Lima" AND GETNEXTMODIFIED(0, "c_frc") # 0, "the edit was not kept in the cursor")
			This.Check(!loDb.lInTransaction, "Save() left a transaction open")
		ENDIF
		loDb.CloseAll(.T.)
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	PROCEDURE TestConflictDetectionAllowsOtherChanges() HELP [Fact, Trait("Category", "Conflicts"), Timeout(180000)]
		This.OnEachEngine("ConflictDetectionAllowsOtherChanges")
	ENDPROC

	* Only the columns this cursor changed are compared: another user's change to another column, a
	* NULL that was there and a value written again are not conflicts.
	PROCEDURE ConflictDetectionAllowsOtherChanges(toE)
		LOCAL lcT, loDb, llOk
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "old", "Oslo")
		loDb = This.Connected(toE)
		loDb.lDetectConflicts = .T.
		This.Check(loDb.Open(lcT, "c_frc2"), "Open() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frc2")
			toE.RawRun("UPDATE " + lcT + " SET qty = 7")
			SELECT c_frc2
			REPLACE name WITH "new", born WITH DATE(2000, 1, 2), city WITH "Oslo"
			llOk = loDb.Save("c_frc2")
			This.Check(llOk, "Save() saw a conflict that is not one: " + loDb.cLastErrorCode + " " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcT, "qty = 7 AND name = 'new' AND born IS NOT NULL") = 1, "the server row is not the merge of both changes")
		ENDIF
		loDb.CloseAll(.T.)
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	* ---- NextNumber ----------------------------------------------------------------
	PROCEDURE TestNextNumberCountsUp() HELP [Fact, Trait("Category", "NextNumber"), Timeout(180000)]
		This.OnEachEngine("NextNumberCountsUp")
	ENDPROC

	* One more each call, per row; a NULL counter starts at 1; no row, or more than one, is refused.
	PROCEDURE NextNumberCountsUp(toE)
		LOCAL lcT, loDb, loP, lnA, lnB, lnC, lnD
		lcT = toE.NewName("ct")
		toE.RawRun("CREATE TABLE " + lcT + " (name VARCHAR(20) NOT NULL PRIMARY KEY, last_no INTEGER)")
		toE.RawRun("INSERT INTO " + lcT + " (name, last_no) VALUES ('invoice', 0)")
		toE.RawRun("INSERT INTO " + lcT + " (name, last_no) VALUES ('order', 100)")
		toE.RawRun("INSERT INTO " + lcT + " (name, last_no) VALUES ('fresh', NULL)")
		loDb = This.Connected(toE)
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "name", "invoice")
		lnA = loDb.NextNumber(lcT, "last_no", "name = :name", loP)
		lnB = loDb.NextNumber(lcT, "last_no", "name = :name", loP)
		loP.name = "order"
		lnC = loDb.NextNumber(lcT, "last_no", "name = :name", loP)
		loP.name = "fresh"
		lnD = loDb.NextNumber(lcT, "last_no", "name = :name", loP)
		This.Check(lnA = 1 AND lnB = 2 AND lnC = 101 AND lnD = 1, "NextNumber() gave " + TRANSFORM(lnA) + ", " + TRANSFORM(lnB) + ", " + TRANSFORM(lnC) + ", " + TRANSFORM(lnD) + ": " + loDb.cLastError)
		This.Check(!loDb.lInTransaction, "NextNumber() left its transaction open")
		loP.name = "nope"
		lnA = loDb.NextNumber(lcT, "last_no", "name = :name", loP)
		This.Check(lnA = -1 AND loDb.cLastErrorCode == "row_not_found", "a missing counter: " + TRANSFORM(lnA) + " " + loDb.cLastErrorCode)
		lnA = loDb.NextNumber(lcT, "last_no", "1 = 1")
		This.Check(lnA = -1 AND loDb.cLastErrorCode == "bad_argument", "a WHERE with three rows: " + TRANSFORM(lnA) + " " + loDb.cLastErrorCode)
		This.Check(toE.RawCount(lcT, "last_no = 2") = 1 AND toE.RawCount(lcT, "last_no = 101") = 1 AND toE.RawCount(lcT, "last_no = 1") = 1, "the counters on the server are not 2, 101 and 1")
		This.Check(!loDb.lInTransaction, "a refused NextNumber() left its transaction open")
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	PROCEDURE TestNextNumberHoldsTheRowUntilCommit() HELP [Fact, Trait("Category", "NextNumber"), Timeout(240000)]
		This.OnEachEngine("NextNumberHoldsTheRowUntilCommit")
	ENDPROC

	* Inside a transaction the counter's row stays locked until Commit: another session waits for it,
	* and so never gets the same number.
	PROCEDURE NextNumberHoldsTheRowUntilCommit(toE)
		LOCAL lcT, loA, loB, loP, lnA, lnB
		lcT = toE.NewName("ct")
		toE.RawRun("CREATE TABLE " + lcT + " (name VARCHAR(20) NOT NULL PRIMARY KEY, last_no INTEGER)")
		toE.RawRun("INSERT INTO " + lcT + " (name, last_no) VALUES ('invoice', 0)")
		loA = This.Connected(toE)
		loB = toE.NewDb()
		loB.cConnectionOptions = loB.cConnectionOptions + toE.LockWaitOptions()
		This.Check(loB.Connect(), "the second session did not connect: " + loB.cLastError)
		toE.ShortLockWait(loB)
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "name", "invoice")
		loA.BeginTransaction()
		lnA = loA.NextNumber(lcT, "last_no", "name = :name", loP)
		lnB = loB.NextNumber(lcT, "last_no", "name = :name", loP)
		This.Check(lnA = 1 AND lnB = -1, "while A holds the counter, A got " + TRANSFORM(lnA) + " and B got " + TRANSFORM(lnB) + " " + loB.cLastError)
		loA.Commit()
		lnB = loB.NextNumber(lcT, "last_no", "name = :name", loP)
		This.Check(lnB = 2, "after the Commit B got " + TRANSFORM(lnB) + " " + loB.cLastError)
		loB.Disconnect()
		loA.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	* ---- HasChanges and Close -------------------------------------------------------
	PROCEDURE TestHasChangesAndCloseWithChanges() HELP [Fact, Trait("Category", "Cursors"), Timeout(180000)]
		This.OnEachEngine("HasChangesAndCloseWithChanges")
	ENDPROC

	* Close() keeps a cursor with changes not saved unless told to discard them, and HasChanges() says so.
	PROCEDURE HasChangesAndCloseWithChanges(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "a", "Oslo")
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_frh")
		This.Check(!loDb.HasChanges("c_frh") AND !loDb.HasChanges(), "a cursor just opened has changes")
		IF USED("c_frh")
			SELECT c_frh
			REPLACE city WITH "Lima"
			This.Check(loDb.HasChanges("c_frh") AND loDb.HasChanges(), "HasChanges() does not see the edit")
			This.Check(!loDb.Close("c_frh") AND loDb.cLastErrorCode == "pending_changes" AND USED("c_frh"), "Close() of a cursor with changes: " + loDb.cLastErrorCode)
			This.Check(!loDb.CloseAll() AND USED("c_frh"), "CloseAll() closed a cursor with changes")
			This.Check(!loDb.CloseGroup("") AND USED("c_frh"), "CloseGroup() closed a cursor with changes")
			This.Check(loDb.Save("c_frh") AND !loDb.HasChanges("c_frh"), "after Save() there are changes: " + loDb.cLastError)
			SELECT c_frh
			REPLACE city WITH "Rome"
			This.Check(loDb.Close("c_frh", .T.) AND !USED("c_frh"), "Close(alias, .T.) did not discard and close")
			This.Check(toE.RawCount(lcT, "city = 'Lima'") = 1, "the discarded edit reached the server")
		ENDIF
		loDb.Open(lcT, "c_frh2", "", .NULL., "", .T.)
		This.Check(!loDb.HasChanges("c_frh2"), "a read-only cursor has changes")
		This.Check(loDb.CloseAll() AND !USED("c_frh2"), "CloseAll() of a cursor without changes failed")
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	* ---- lStrict ---------------------------------------------------------------------
	PROCEDURE TestStrictSaveFailureLeavesNoTransactionOpen() HELP [Fact, Trait("Category", "Errors"), Timeout(180000)]
		This.OnEachEngine("StrictSaveFailureLeavesNoTransactionOpen")
	ENDPROC

	* With lStrict the failure is raised after Save() has rolled back its own transaction, not in the
	* middle, where it would leave the transaction open and its rows locked on the server.
	PROCEDURE StrictSaveFailureLeavesNoTransactionOpen(toE)
		LOCAL lcT, loDb, llRaised, llOpen
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		loDb = This.Connected(toE)
		loDb.Open(lcT, "c_frst", "1 = 0")
		IF USED("c_frst")
			loDb.lStrict = .T.
			SELECT c_frst
			APPEND BLANK
			REPLACE name WITH "bad", qty WITH -1
			llRaised = .F.
			TRY
				loDb.Save("c_frst")
			CATCH
				llRaised = .T.
			ENDTRY
			loDb.lStrict = .F.
			llOpen = loDb.lInTransaction
			IF llOpen
				loDb.Rollback()
			ENDIF
			This.Check(llRaised, "Save() with lStrict did not raise")
			This.Check(!llOpen, "a failed Save() with lStrict left its transaction open")
			This.Check(GETNEXTMODIFIED(0, "c_frst") # 0, "the edits were thrown away")
			This.Check(toE.RawCount(lcT) = 0, "the server kept rows of a failed Save()")
		ENDIF
		loDb.CloseAll(.T.)
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC


	PROCEDURE TestLibraryHasNoUserInterface() HELP [Fact, Trait("Category", "Headless")]
		LOCAL lcSource, laLines[1], lnI, lcLine, lcFound
		lcSource = FILETOSTR(FrLibrarySource())
		lcFound = ""
		FOR lnI = 1 TO ALINES(laLines, lcSource)
			lcLine = UPPER(ALLTRIM(CHRTRAN(laLines[lnI], CHR(9), " ")))
			IF LEFT(lcLine, 1) == "*" OR LEFT(lcLine, 2) == "&" + "&"
				LOOP
			ENDIF
			IF "MESSAGEBOX(" $ lcLine OR "WAIT WINDOW" $ lcLine OR "INKEY(" $ lcLine OR "DOEVENTS" $ lcLine OR "_SCREEN" $ lcLine
				lcFound = lcFound + TRANSFORM(lnI) + ": " + LEFT(ALLTRIM(laLines[lnI]), 60) + CHR(13) + CHR(10)
			ENDIF
		ENDFOR
		__assert.True(EMPTY(lcFound), "user interface in " + JUSTFNAME(FrLibrarySource()) + ":" + CHR(13) + CHR(10) + lcFound)
	ENDPROC
ENDDEFINE
