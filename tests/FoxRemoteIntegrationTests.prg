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

	PROCEDURE TestQueryWithACustomBasedParamsObject() HELP [Fact, Trait("Category", "Query"), Timeout(180000)]
		This.OnEachEngine("QueryWithACustomBasedParamsObject")
	ENDPROC

	* A parameters object based on Custom (it carries Objects, Controls and Parent, which GETPEM()
	* refuses) gives the same row as an Empty. In SQLite the text still travels in UTF-8.
	PROCEDURE QueryWithACustomBasedParamsObject(toE)
		LOCAL lcT, loDb, loP, lcCity, lcName, lnI, lcHow, lvRaw
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		lcCity = "Logroño"
		lcName = "Begoña"
		loDb = This.Connected(toE)
		* Written through the library, so that SQLite holds the UTF-8 that FoxRemote writes.
		This.Check(loDb.Execute("INSERT INTO " + lcT + " (name, city) VALUES ('" + lcName + "', '" + lcCity + "')") = 1, "INSERT failed: " + TRANSFORM(loDb.cLastError))
		This.Check(loDb.Execute("INSERT INTO " + lcT + " (name, city) VALUES ('Luis', 'Lima')") = 1, "INSERT failed: " + TRANSFORM(loDb.cLastError))
		FOR lnI = 1 TO 2
			lcHow = IIF(lnI = 1, "Custom", "Empty")
			loP = CREATEOBJECT(lcHow)
			ADDPROPERTY(loP, "city", lcCity)
			This.Check(loDb.Query("SELECT name, city FROM " + lcT + " WHERE city = :city", "c_frq", loP), lcHow + ": Query() with :city failed: " + TRANSFORM(loDb.cLastError))
			IF USED("c_frq")
				This.Check(RECCOUNT("c_frq") = 1, lcHow + ": Query() returned " + TRANSFORM(RECCOUNT("c_frq")) + " rows, expected 1")
				This.Check(ALLTRIM(c_frq.name) == lcName AND ALLTRIM(c_frq.city) == lcCity, lcHow + ": read back name hex " + ;
					STRCONV(ALLTRIM(c_frq.name), 15) + " and city hex " + STRCONV(ALLTRIM(c_frq.city), 15) + ", sent " + STRCONV(lcName, 15) + " and " + STRCONV(lcCity, 15))
				USE IN c_frq
			ENDIF
		ENDFOR
		IF toE.cName == "sqlite"
			* SQLite stores UTF-8: any other tool has to read the same text.
			lvRaw = toE.RawScalar("SELECT hex(city) FROM " + lcT + " WHERE name <> 'Luis'")
			This.Check(UPPER(TRANSFORM(lvRaw)) == FrUtf8Hex(lcCity), "SQLite holds hex " + TRANSFORM(lvRaw) + ", expected the UTF-8 " + FrUtf8Hex(lcCity))
		ELSE
			This.Check(toE.RawCount(lcT, "city = '" + lcCity + "'") = 1, "another connection does not find the row with city = '" + lcCity + "'")
		ENDIF
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestQueryWithAMissingParameterFails() HELP [Fact, Trait("Category", "Query"), Timeout(180000)]
		This.OnEachEngine("QueryWithAMissingParameterFails")
	ENDPROC

	* A :name without its property in oParams, or without oParams, would open VFP's "View Parameter"
	* dialog: it fails with bad_argument before anything is sent. A ':' inside quotes is no parameter.
	PROCEDURE QueryWithAMissingParameterFails(toE)
		LOCAL lcT, loDb, loP, lnI, lcHow, llOk
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "Ana", "Madrid")
		loDb = This.Connected(toE)
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "town", "Madrid")
		FOR lnI = 1 TO 2
			lcHow = IIF(lnI = 1, "an oParams without city", "no oParams")
			IF lnI = 1
				llOk = loDb.Query("SELECT name FROM " + lcT + " WHERE city = :city", "c_frq", loP)
			ELSE
				llOk = loDb.Query("SELECT name FROM " + lcT + " WHERE city = :city", "c_frq")
			ENDIF
			This.Check(!llOk, lcHow + ": Query() returned .T.")
			This.Check(loDb.cLastErrorCode == "bad_argument", lcHow + ": cLastErrorCode = '" + loDb.cLastErrorCode + "', expected bad_argument")
			This.Check(":city" $ loDb.cLastError, lcHow + ": the message does not name :city: " + loDb.cLastError)
			This.Check(!USED("c_frq"), lcHow + ": Query() left the cursor c_frq")
			IF USED("c_frq")
				USE IN c_frq
			ENDIF
		ENDFOR
		This.Check(loDb.Query("SELECT name FROM " + lcT + " WHERE city <> 'a:b'", "c_frq"), "Query() with ':' inside quotes failed: " + TRANSFORM(loDb.cLastError))
		This.Check(USED("c_frq") AND RECCOUNT("c_frq") = 1, "Query() with ':' inside quotes did not return the row")
		IF USED("c_frq")
			USE IN c_frq
		ENDIF
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestQueryWithParameterPairs() HELP [Fact, Trait("Category", "Query"), Timeout(180000)]
		This.OnEachEngine("QueryWithParameterPairs")
	ENDPROC

	* Name, value pairs after the cursor give the same row as the object of parameters: a city with
	* an eñe and a date. In SQLite the eñe still travels in UTF-8.
	PROCEDURE QueryWithParameterPairs(toE)
		LOCAL lcT, loDb, loP, lcCity, lcName, lnI, lcHow, llOk, lvRaw
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		lcCity = "Logroño"
		lcName = "Begoña"
		loDb = This.Connected(toE)
		This.Check(loDb.Execute("INSERT INTO " + lcT + " (name, city, born) VALUES (:name, :city, :born)", ;
			This.Values("name", lcName, "city", lcCity, "born", {^2026-11-02})) = 1, "INSERT failed: " + TRANSFORM(loDb.cLastError))
		This.Check(loDb.Execute("INSERT INTO " + lcT + " (name, city, born) VALUES (:name, :city, :born)", ;
			This.Values("name", "Luis", "city", lcCity, "born", {^2026-11-03})) = 1, "INSERT failed: " + TRANSFORM(loDb.cLastError))
		FOR lnI = 1 TO 2
			IF lnI = 1
				lcHow = "pairs"
				llOk = loDb.Query("SELECT name, city FROM " + lcT + " WHERE city = :city AND born = :born", "c_frq", "city", lcCity, "born", {^2026-11-02})
			ELSE
				lcHow = "object"
				llOk = loDb.Query("SELECT name, city FROM " + lcT + " WHERE city = :city AND born = :born", "c_frq", This.Values("city", lcCity, "born", {^2026-11-02}))
			ENDIF
			This.Check(llOk, lcHow + ": Query() failed: " + TRANSFORM(loDb.cLastError))
			IF USED("c_frq")
				This.Check(RECCOUNT("c_frq") = 1, lcHow + ": Query() returned " + TRANSFORM(RECCOUNT("c_frq")) + " rows, expected 1")
				This.Check(ALLTRIM(c_frq.name) == lcName AND ALLTRIM(c_frq.city) == lcCity, lcHow + ": read back name hex " + ;
					STRCONV(ALLTRIM(c_frq.name), 15) + " and city hex " + STRCONV(ALLTRIM(c_frq.city), 15) + ", sent " + STRCONV(lcName, 15) + " and " + STRCONV(lcCity, 15))
				USE IN c_frq
			ENDIF
		ENDFOR
		IF toE.cName == "sqlite"
			* SQLite stores UTF-8: any other tool has to read the same text.
			lvRaw = toE.RawScalar("SELECT hex(city) FROM " + lcT + " WHERE name <> 'Luis'")
			This.Check(UPPER(TRANSFORM(lvRaw)) == FrUtf8Hex(lcCity), "SQLite holds hex " + TRANSFORM(lvRaw) + ", expected the UTF-8 " + FrUtf8Hex(lcCity))
		ENDIF
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestExecuteAndScalarWithPairs() HELP [Fact, Trait("Category", "Query"), Timeout(180000)]
		This.OnEachEngine("ExecuteAndScalarWithPairs")
	ENDPROC

	* An UPDATE with two pairs and a COUNT with one.
	PROCEDURE ExecuteAndScalarWithPairs(toE)
		LOCAL lcT, loDb, lnRows, lvN
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "Ana", "Madrid")
		toE.InsertCustomer(lcT, "Luis", "Lima")
		loDb = This.Connected(toE)
		lnRows = loDb.Execute("UPDATE " + lcT + " SET city = :city WHERE name = :name", "city", "Oslo", "name", "Ana")
		This.Check(lnRows = 1, "Execute() with two pairs = " + TRANSFORM(lnRows) + ", expected 1 row: " + TRANSFORM(loDb.cLastError))
		This.Check(toE.RawCount(lcT, "name = 'Ana' AND city = 'Oslo'") = 1, "the server does not have Ana in Oslo")
		This.Check(toE.RawCount(lcT, "name = 'Luis' AND city = 'Lima'") = 1, "the UPDATE touched Luis")
		lvN = loDb.Scalar("SELECT COUNT(*) FROM " + lcT + " WHERE city = :city", "city", "Oslo")
		This.Check(VARTYPE(lvN) == "N" AND lvN = 1, "Scalar() with one pair = " + TRANSFORM(lvN) + " (type " + VARTYPE(lvN) + "), expected the number 1: " + TRANSFORM(loDb.cLastError))
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestOpenWithParamsHelper() HELP [Fact, Trait("Category", "Query"), Timeout(180000)]
		This.OnEachEngine("OpenWithParamsHelper")
	ENDPROC

	* Params() where oParams is not the last argument: Open() with Save(), and OpenView().
	PROCEDURE OpenWithParamsHelper(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "Ana", "Madrid")
		toE.InsertCustomer(lcT, "Luis", "Lima")
		loDb = This.Connected(toE)
		This.Check(loDb.Open(lcT, "c_frop", "city = :city", loDb.Params("city", "Madrid")), "Open() with Params() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frop")
			This.Check(RECCOUNT("c_frop") = 1 AND ALLTRIM(c_frop.name) == "Ana", "Open() with Params() has " + TRANSFORM(RECCOUNT("c_frop")) + " rows")
			SELECT c_frop
			REPLACE name WITH "Anna"
			This.Check(loDb.Save("c_frop"), "Save() failed: " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcT, "name = 'Anna' AND city = 'Madrid'") = 1, "the server does not have Anna")
		ENDIF
		This.Check(loDb.OpenView("SELECT id, name FROM " + lcT + " WHERE city = :city", "c_frov", loDb.Params("city", "Lima")), ;
			"OpenView() with Params() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frov")
			This.Check(RECCOUNT("c_frov") = 1 AND ALLTRIM(c_frov.name) == "Luis", "OpenView() with Params() has " + TRANSFORM(RECCOUNT("c_frov")) + " rows")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	* An Empty with the pairs, written by hand as the object of parameters always was.
	FUNCTION Values(tc1, tv1, tc2, tv2, tc3, tv3)
		LOCAL loP, lnI
		loP = CREATEOBJECT("Empty")
		FOR lnI = 1 TO INT(PCOUNT() / 2)
			ADDPROPERTY(loP, EVALUATE("tc" + TRANSFORM(lnI)), EVALUATE("tv" + TRANSFORM(lnI)))
		ENDFOR
		RETURN loP
	ENDFUNC

	PROCEDURE TestQueryWithAQuestionMarkVariableFails() HELP [Fact, Trait("Category", "Query"), Timeout(180000)]
		This.OnEachEngine("QueryWithAQuestionMarkVariableFails")
	ENDPROC

	* A ?variable that is LOCAL of the caller fails with bad_argument and the call comes back: before
	* the guard it opened VFP's "View Parameter" dialog, and foxproof would end the test on its time
	* limit, naming the dialog. A PRIVATE of the caller is seen, as always, and gives its row.
	PROCEDURE QueryWithAQuestionMarkVariableFails(toE)
		LOCAL lcT, loDb, lcCity, llOk
		PRIVATE pcCity
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "Ana", "Madrid")
		toE.InsertCustomer(lcT, "Luis", "Lima")
		loDb = This.Connected(toE)
		lcCity = "Madrid"
		pcCity = "Lima"
		llOk = loDb.Query("SELECT name FROM " + lcT + " WHERE city = ?lcCity", "c_frq")
		This.Check(!llOk, "Query() with ?lcCity (LOCAL) returned .T.")
		This.Check(loDb.cLastErrorCode == "bad_argument", "?lcCity: cLastErrorCode = '" + loDb.cLastErrorCode + "', expected bad_argument")
		This.Check("?lcCity" $ loDb.cLastError, "?lcCity: the message does not name it: " + loDb.cLastError)
		This.Check(!USED("c_frq"), "?lcCity: Query() left the cursor c_frq")
		IF USED("c_frq")
			USE IN c_frq
		ENDIF
		llOk = loDb.Query("SELECT name FROM " + lcT + " WHERE city = ?pcCity", "c_frq")
		This.Check(llOk, "Query() with ?pcCity (PRIVATE) failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frq")
			This.Check(RECCOUNT("c_frq") = 1 AND ALLTRIM(c_frq.name) == "Luis", "?pcCity: Query() returned " + TRANSFORM(RECCOUNT("c_frq")) + " rows")
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
	* Views (OpenView)
	* ======================================================================== *
	* The SELECT of an order with its customer's name, by :id.
	FUNCTION OrderViewSql(tcOrders, tcCustomers)
		RETURN "SELECT o.id, o.customer_id, o.status, c.name AS customer FROM " + tcOrders + " o JOIN " + tcCustomers + ;
			" c ON c.id = o.customer_id WHERE o.id = :id"
	ENDFUNC

	PROCEDURE TestOpenViewWithoutATableIsReadOnly() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewWithoutATableIsReadOnly")
	ENDPROC

	* A join opens and reads; without a base table Save() refuses it, and a table without a key opens too.
	PROCEDURE OpenViewWithoutATableIsReadOnly(toE)
		LOCAL lcC, lcO, lcN, loDb, loP
		lcC = toE.NewName()
		lcO = toE.NewName()
		lcN = toE.NewName()
		toE.CreateCustomers(lcC)
		toE.CreateOrders(lcO, lcC)
		toE.RawRun("CREATE TABLE " + lcN + " (code VARCHAR(10), descr VARCHAR(40))")
		toE.RawRun("INSERT INTO " + lcN + " (code, descr) VALUES ('A1', 'first')")
		toE.InsertCustomer(lcC, "Ana", "Bonn")
		toE.InsertOrder(lcO, toE.MaxId(lcC), "open")
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "id", toE.MaxId(lcO))
		loDb = This.Connected(toE)
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frv", loP), "OpenView() of a join failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frv")
			This.Check(ALIAS() == "C_FRV", "OpenView() did not select the cursor: " + ALIAS())
			This.Check(RECCOUNT("c_frv") = 1 AND ALLTRIM(c_frv.customer) == "Ana" AND ALLTRIM(c_frv.status) == "open", ;
				"the view did not read the order with its customer's name")
			REPLACE status WITH "paid" IN c_frv
			This.Check(!loDb.Save("c_frv"), "Save() of a view without a table returned .T.")
			This.Check(loDb.cLastErrorCode == "read_only", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected read_only")
			This.Check(toE.RawCount(lcO, "status = 'open'") = 1, "Save() of a read-only view changed the server")
		ENDIF
		This.Check(loDb.OpenView("SELECT code, descr FROM " + lcN, "c_frvn"), "OpenView() of a table without a primary key failed: " + TRANSFORM(loDb.cLastError))
		This.Check(USED("c_frvn") AND RECCOUNT("c_frvn") = 1, "the view of a table without a primary key did not read its row")
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcO)
		toE.DropTable(lcC)
		toE.DropTable(lcN)
	ENDPROC

	PROCEDURE TestOpenViewRejectsANonSelect() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewRejectsANonSelect")
	ENDPROC

	* Only a SELECT is a view: anything else is refused before it reaches the server.
	PROCEDURE OpenViewRejectsANonSelect(toE)
		LOCAL lcT, loDb
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "one", "Graz")
		toE.InsertCustomer(lcT, "two", "Graz")
		loDb = This.Connected(toE)
		This.Check(!loDb.OpenView("DELETE FROM " + lcT, "c_frvd"), "OpenView() of a DELETE returned .T.")
		This.Check(loDb.cLastErrorCode == "bad_argument", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected bad_argument")
		This.Check(!USED("c_frvd"), "OpenView() of a DELETE left a cursor open")
		This.Check(toE.RawCount(lcT) = 2, "the DELETE of OpenView() reached the server")
		This.Check(!loDb.OpenView("", "c_frvd") AND loDb.cLastErrorCode == "bad_argument", "OpenView() without a SELECT: '" + loDb.cLastErrorCode + "'")
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	PROCEDURE TestOpenViewRefreshKeepsTheParameters() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewRefreshKeepsTheParameters")
	ENDPROC

	* Another user changes the order; Refresh() reads it again with the same :id, and only that order.
	PROCEDURE OpenViewRefreshKeepsTheParameters(toE)
		LOCAL lcC, lcO, loDb, loP, lnId
		lcC = toE.NewName()
		lcO = toE.NewName()
		toE.CreateCustomers(lcC)
		toE.CreateOrders(lcO, lcC)
		toE.InsertCustomer(lcC, "Ana", "Bonn")
		toE.InsertOrder(lcO, toE.MaxId(lcC), "open")
		lnId = toE.MaxId(lcO)
		toE.InsertOrder(lcO, toE.MaxId(lcC), "open")
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "id", lnId)
		loDb = This.Connected(toE)
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frvr", loP), "OpenView() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frvr")
			This.Check(toE.RawRun("UPDATE " + lcO + " SET status = 'paid' WHERE id = " + TRANSFORM(lnId)), "the raw UPDATE failed: " + toE.cLastError)
			This.Check(loDb.Refresh("c_frvr"), "Refresh() failed: " + TRANSFORM(loDb.cLastError))
			This.Check(USED("c_frvr") AND RECCOUNT("c_frvr") = 1, "after Refresh() the view has " + IIF(USED("c_frvr"), TRANSFORM(RECCOUNT("c_frvr")), "no") + " rows, expected 1")
			This.Check(USED("c_frvr") AND c_frvr.id = lnId AND ALLTRIM(c_frvr.status) == "paid", "Refresh() did not bring the change of order " + TRANSFORM(lnId))
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcO)
		toE.DropTable(lcC)
	ENDPROC

	* ---- updatable views (cTable) ------------------------------------------------------
	* Customer "Ana" with one order "open": fills tcC and tcO, returns the order's :id parameter.
	FUNCTION OneOrder(toE, tcC, tcO)
		LOCAL loP
		toE.CreateCustomers(tcC)
		toE.CreateOrders(tcO, tcC)
		toE.InsertCustomer(tcC, "Ana", "Bonn")
		toE.InsertOrder(tcO, toE.MaxId(tcC), "open")
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "id", toE.MaxId(tcO))
		RETURN loP
	ENDFUNC

	PROCEDURE TestOpenViewUpdatesOnlyTheBaseTable() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewUpdatesOnlyTheBaseTable")
	ENDPROC

	* A change to the order's status in the join reaches the orders table; the customers stay as they were.
	PROCEDURE OpenViewUpdatesOnlyTheBaseTable(toE)
		LOCAL lcC, lcO, loDb, loP
		lcC = toE.NewName()
		lcO = toE.NewName()
		loP = This.OneOrder(toE, lcC, lcO)
		loDb = This.Connected(toE)
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frvu", loP, lcO), "OpenView() with a table failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frvu")
			REPLACE status WITH "paid" IN c_frvu
			This.Check(loDb.Save("c_frvu"), "Save() of the view failed: " + TRANSFORM(loDb.cLastError))
			This.Check(ALLTRIM(TRANSFORM(toE.RawScalar("SELECT status FROM " + lcO + " WHERE id = " + TRANSFORM(loP.id)))) == "paid", ;
				"the server does not have the new status")
			This.Check(toE.RawCount(lcC, "name = 'Ana'") = 1 AND toE.RawCount(lcC) = 1, "Save() of the view changed the customers")
			This.Check(ALLTRIM(c_frvu.customer) == "Ana", "the joined column left the cursor or changed: " + TRANSFORM(c_frvu.customer))
			This.Check(!loDb.HasChanges("c_frvu"), "after Save() the view has changes")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcO)
		toE.DropTable(lcC)
	ENDPROC

	PROCEDURE TestOpenViewInsertGetsTheId() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewInsertGetsTheId")
	ENDPROC

	* A row added to the view goes to the orders table, and its new id comes back to the cursor.
	PROCEDURE OpenViewInsertGetsTheId(toE)
		LOCAL lcC, lcO, loDb, loP, lnCustomer
		lcC = toE.NewName()
		lcO = toE.NewName()
		loP = This.OneOrder(toE, lcC, lcO)
		lnCustomer = toE.MaxId(lcC)
		loDb = This.Connected(toE)
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frvi", loP, lcO), "OpenView() with a table failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frvi")
			INSERT INTO c_frvi (customer_id, status) VALUES (lnCustomer, "new")
			This.Check(loDb.Save("c_frvi"), "Save() of a new row of the view failed: " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcO, "status = 'new' AND customer_id = " + TRANSFORM(lnCustomer)) = 1, "the new row is not in the orders table")
			SELECT c_frvi
			LOCATE FOR ALLTRIM(status) == "new"
			This.Check(FOUND() AND c_frvi.id = toE.MaxId(lcO), "the new id did not come back to the cursor: " + IIF(FOUND(), TRANSFORM(c_frvi.id), "no row"))
			This.Check(toE.RawCount(lcC) = 1, "Save() of the view added a customer")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcO)
		toE.DropTable(lcC)
	ENDPROC

	PROCEDURE TestOpenViewDeletesTheRow() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewDeletesTheRow")
	ENDPROC

	* A row deleted in the view is deleted from the orders table, and its customer stays.
	PROCEDURE OpenViewDeletesTheRow(toE)
		LOCAL lcC, lcO, loDb, loP
		lcC = toE.NewName()
		lcO = toE.NewName()
		loP = This.OneOrder(toE, lcC, lcO)
		loDb = This.Connected(toE)
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frvx", loP, lcO), "OpenView() with a table failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frvx")
			DELETE IN c_frvx
			This.Check(loDb.Save("c_frvx"), "Save() of a deleted row of the view failed: " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcO, "id = " + TRANSFORM(loP.id)) = 0, "the order is still on the server")
			This.Check(toE.RawCount(lcC, "name = 'Ana'") = 1, "the customer of the deleted order is gone")
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcO)
		toE.DropTable(lcC)
	ENDPROC

	PROCEDURE TestOpenViewRefusesAKeyMissingFromTheSelect() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewRefusesAKeyMissingFromTheSelect")
	ENDPROC

	* Without the key in the SELECT Save() could not find the row: OpenView() refuses, and leaves
	* neither a cursor nor a transaction behind.
	PROCEDURE OpenViewRefusesAKeyMissingFromTheSelect(toE)
		LOCAL lcC, lcO, loDb, loP, lcSql
		lcC = toE.NewName()
		lcO = toE.NewName()
		loP = This.OneOrder(toE, lcC, lcO)
		loDb = This.Connected(toE)
		lcSql = "SELECT o.status, c.name AS customer FROM " + lcO + " o JOIN " + lcC + " c ON c.id = o.customer_id WHERE o.id = :id"
		This.Check(!loDb.OpenView(lcSql, "c_frvk", loP, lcO), "OpenView() of a SELECT without the key returned .T.")
		This.Check(loDb.cLastErrorCode == "no_primary_key", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected no_primary_key")
		This.Check(" id " $ " " + LOWER(CHRTRAN(loDb.cLastError, ",:", "  ")) + " " AND "is not in the select" $ LOWER(loDb.cLastError), ;
			"the message does not say that the key column is missing from the SELECT: " + loDb.cLastError)
		This.Check(!USED("c_frvk"), "OpenView() left the cursor open")
		This.Check(!loDb.lInTransaction, "OpenView() left a transaction open")
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frvk", loP, lcO), "the alias is not free after the refusal: " + TRANSFORM(loDb.cLastError))
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcO)
		toE.DropTable(lcC)
	ENDPROC

	PROCEDURE TestOpenViewRefusesAnEditToAJoinedColumn() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewRefusesAnEditToAJoinedColumn")
	ENDPROC

	* The customer's name is not a column of orders: a change to it makes Save() fail with read_only
	* before anything is sent, and the change stays in the buffer.
	PROCEDURE OpenViewRefusesAnEditToAJoinedColumn(toE)
		LOCAL lcC, lcO, loDb, loP
		lcC = toE.NewName()
		lcO = toE.NewName()
		loP = This.OneOrder(toE, lcC, lcO)
		loDb = This.Connected(toE)
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frvj", loP, lcO), "OpenView() with a table failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frvj")
			REPLACE customer WITH "X" IN c_frvj
			This.Check(!loDb.Save("c_frvj"), "Save() of a change to a joined column returned .T.")
			This.Check(loDb.cLastErrorCode == "read_only", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected read_only")
			This.Check("customer" $ LOWER(loDb.cLastError), "the message does not name the column: " + loDb.cLastError)
			This.Check(loDb.HasChanges("c_frvj"), "the change left the buffer")
			This.Check(!loDb.lInTransaction, "the failed Save() left a transaction open")
			This.Check(toE.RawCount(lcC, "name = 'Ana'") = 1, "the customer's name changed on the server")
			This.Check(loDb.Discard("c_frvj") AND !loDb.HasChanges("c_frvj"), "Discard() did not clear the change")
			This.Check(loDb.Close("c_frvj") AND !USED("c_frvj"), "Close() after Discard() failed: " + loDb.cLastErrorCode)
		ENDIF
		* The same for a new row that fills the joined column: the INSERT is not sent either.
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frvj", loP, lcO), "OpenView() again failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frvj")
			INSERT INTO c_frvj (customer_id, status, customer) VALUES (toE.MaxId(lcC), "new", "X")
			This.Check(!loDb.Save("c_frvj"), "Save() of a new row with a joined column returned .T.")
			This.Check(loDb.cLastErrorCode == "read_only", "new row: cLastErrorCode = '" + loDb.cLastErrorCode + "', expected read_only")
			This.Check("customer" $ LOWER(loDb.cLastError), "new row: the message does not name the column: " + loDb.cLastError)
			This.Check(toE.RawCount(lcO, "status = 'new'") = 0, "the new row reached the orders table")
			This.Check(!loDb.lInTransaction, "the failed Save() of a new row left a transaction open")
		ENDIF
		loDb.CloseAll(.T.)
		loDb.Disconnect()
		toE.DropTable(lcO)
		toE.DropTable(lcC)
	ENDPROC

	PROCEDURE TestOpenViewInAGroupSavesWithTheOthers() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewInAGroupSavesWithTheOthers")
	ENDPROC

	* A view and a cursor of Open() in one group: SaveGroup() sends both in one transaction, and when
	* the view fails, the other cursor's change does not reach the server either.
	PROCEDURE OpenViewInAGroupSavesWithTheOthers(toE)
		LOCAL lcC, lcO, loDb, loP, llOk, llOpen
		lcC = toE.NewName()
		lcO = toE.NewName()
		loP = This.OneOrder(toE, lcC, lcO)
		loDb = This.Connected(toE)
		This.Check(loDb.Open(lcC, "c_frg1", "", .NULL., "", .F., "g"), "Open() in the group failed: " + TRANSFORM(loDb.cLastError))
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frg2", loP, lcO, "", "g"), "OpenView() in the group failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frg1") AND USED("c_frg2")
			REPLACE city WITH "Kiel" IN c_frg1
			REPLACE status WITH "paid" IN c_frg2
			This.Check(loDb.SaveGroup("g"), "SaveGroup() of a cursor and a view failed: " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcC, "city = 'Kiel'") = 1 AND toE.RawCount(lcO, "status = 'paid'") = 1, "SaveGroup() did not reach the server")
		ENDIF
		* The failing pair: the view's change is to the joined column.
		This.Check(loDb.Open(lcC, "c_frg3", "", .NULL., "", .F., "g2"), "Open() in the second group failed: " + TRANSFORM(loDb.cLastError))
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frg4", loP, lcO, "", "g2"), "OpenView() in the second group failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frg3") AND USED("c_frg4")
			REPLACE city WITH "Rome" IN c_frg3
			REPLACE customer WITH "X" IN c_frg4
			llOk = loDb.SaveGroup("g2")
			llOpen = loDb.lInTransaction
			IF llOpen
				loDb.Rollback()
			ENDIF
			This.Check(!llOk AND loDb.cLastErrorCode == "read_only", "SaveGroup() with a change to a joined column: " + TRANSFORM(llOk) + " " + loDb.cLastErrorCode)
			This.Check(!llOpen, "the failed SaveGroup() left its transaction open")
			This.Check(toE.RawCount(lcC, "city = 'Rome'") = 0, "the other cursor's change reached the server")
			This.Check(loDb.HasChanges("c_frg3") AND loDb.HasChanges("c_frg4"), "a cursor of the group lost its changes")
		ENDIF
		loDb.CloseAll(.T.)
		loDb.Disconnect()
		toE.DropTable(lcO)
		toE.DropTable(lcC)
	ENDPROC

	PROCEDURE TestOpenViewDetectsAConflict() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		This.OnEachEngine("OpenViewDetectsAConflict")
	ENDPROC

	* With lDetectConflicts, another user's change to the column the view changes is a conflict; a
	* change to the joined table, or to a column of the base table the view did not change, is not.
	PROCEDURE OpenViewDetectsAConflict(toE)
		LOCAL lcC, lcO, loDb, loP, llOk, lnOther
		lcC = toE.NewName()
		lcO = toE.NewName()
		loP = This.OneOrder(toE, lcC, lcO)
		toE.InsertCustomer(lcC, "Bea", "Graz")
		lnOther = toE.MaxId(lcC)
		loDb = This.Connected(toE)
		loDb.lDetectConflicts = .T.
		This.Check(loDb.OpenView(This.OrderViewSql(lcO, lcC), "c_frvc", loP, lcO), "OpenView() with a table failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frvc")
			toE.RawRun("UPDATE " + lcO + " SET status = 'cancelled' WHERE id = " + TRANSFORM(loP.id))
			REPLACE status WITH "paid" IN c_frvc
			llOk = loDb.Save("c_frvc")
			This.Check(!llOk AND loDb.cLastErrorCode == "row_changed", "Save() over another user's change: " + TRANSFORM(llOk) + " " + loDb.cLastErrorCode + " " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcO, "status = 'cancelled'") = 1, "the other user's change was overwritten")
			This.Check(!loDb.lInTransaction, "Save() left a transaction open")
			This.Check(loDb.Discard("c_frvc") AND loDb.Refresh("c_frvc"), "Discard() and Refresh() failed: " + TRANSFORM(loDb.cLastError))
			This.Check(ALLTRIM(c_frvc.status) == "cancelled", "Refresh() did not bring the other user's status")
			REPLACE status WITH "paid" IN c_frvc
			toE.RawRun("UPDATE " + lcC + " SET name = 'Zoe' WHERE name = 'Ana'")
			toE.RawRun("UPDATE " + lcO + " SET customer_id = " + TRANSFORM(lnOther) + " WHERE id = " + TRANSFORM(loP.id))
			llOk = loDb.Save("c_frvc")
			This.Check(llOk, "Save() saw a conflict that is not one: " + loDb.cLastErrorCode + " " + TRANSFORM(loDb.cLastError))
			This.Check(toE.RawCount(lcO, "status = 'paid' AND customer_id = " + TRANSFORM(lnOther)) = 1, "the server row is not the merge of both changes")
		ENDIF
		loDb.CloseAll(.T.)
		loDb.Disconnect()
		toE.DropTable(lcO)
		toE.DropTable(lcC)
	ENDPROC

	PROCEDURE TestOpenViewRefusesASelectWithoutAResultSet() HELP [Fact, Trait("Category", "View"), Timeout(180000)]
		* SQLite and Firebird have no SELECT that returns nothing: SELECT ... INTO is not SQL there.
		This.OnEachEngine("OpenViewRefusesASelectWithoutAResultSet", "mssql,postgresql,mysql,mariadb")
	ENDPROC

	* A SELECT that leaves no cursor (it writes into a table or a variable) is refused with
	* bad_argument instead of raising a VFP error.
	PROCEDURE OpenViewRefusesASelectWithoutAResultSet(toE)
		LOCAL lcT, lcNew, loDb, lcSql
		lcT = toE.NewName()
		lcNew = toE.NewName()
		toE.CreateCustomers(lcT)
		toE.InsertCustomer(lcT, "one", "Graz")
		IF INLIST(toE.cName, "mysql", "mariadb")
			lcSql = "SELECT 1 INTO @fr_v"
		ELSE
			lcSql = "SELECT * INTO " + lcNew + " FROM " + lcT
		ENDIF
		loDb = This.Connected(toE)
		This.Check(!loDb.OpenView(lcSql, "c_frvn"), "OpenView() of " + lcSql + " returned .T.")
		This.Check(loDb.cLastErrorCode == "bad_argument", "cLastErrorCode = '" + loDb.cLastErrorCode + "', expected bad_argument: " + loDb.cLastError)
		This.Check(!USED("c_frvn"), "OpenView() of a SELECT without a result set left a cursor")
		This.Check(loDb.OpenView("SELECT name FROM " + lcT + " WHERE 1 = 0", "c_frvn") AND USED("c_frvn") AND RECCOUNT("c_frvn") = 0, ;
			"a SELECT with no rows is not a view: " + loDb.cLastError)
		loDb.CloseAll()
		loDb.Disconnect()
		IF !INLIST(toE.cName, "mysql", "mariadb")
			toE.DropTable(lcNew)
		ENDIF
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

	PROCEDURE TestGetDatabasesListsTheConnectedOne() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("GetDatabasesListsTheConnectedOne")
	ENDPROC

	* The databases the login can open, one row each (name). On a file engine (SQLite, Firebird)
	* the database is the file: one row, the one connected.
	PROCEDURE GetDatabasesListsTheConnectedOne(toE)
		LOCAL loDb, lcMine, lnFound, lnRows
		loDb = This.Connected(toE)
		lcMine = LOWER(ALLTRIM(loDb.cDatabase))
		This.Check(loDb.GetDatabases("c_frdb"), "GetDatabases() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frdb")
			SELECT c_frdb
			lnRows = RECCOUNT()
			This.Check(FIELD(1) == "NAME" AND TYPE("name") == "C" AND FSIZE("name") = 128, "GetDatabases() cursor is not name C(128)")
			COUNT FOR LOWER(ALLTRIM(name)) == lcMine TO lnFound
			This.Check(lnFound = 1, "GetDatabases() lists the connected database " + TRANSFORM(lnFound) + " times, expected 1")
			IF toE.IsServer()
				This.Check(lnRows >= 2, "GetDatabases() lists " + TRANSFORM(lnRows) + " databases on a server, expected the test one and a system one")
				COUNT FOR LOWER(ALLTRIM(name)) == LOWER(toE.MaintenanceDatabase()) TO lnFound
				This.Check(lnFound = 1, "GetDatabases() does not list " + toE.MaintenanceDatabase())
			ELSE
				This.Check(lnRows = 1, "GetDatabases() lists " + TRANSFORM(lnRows) + " rows on a file engine, expected 1")
			ENDIF
			USE IN c_frdb
		ENDIF
		loDb.Disconnect()
	ENDPROC

	PROCEDURE TestGetObjectsListsTablesAndViews() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("GetObjectsListsTablesAndViews")
	ENDPROC

	* A table and a view made outside the library: GetObjects() has both with their type, and
	* GetTables() keeps giving only the table (code counts on that).
	PROCEDURE GetObjectsListsTablesAndViews(toE)
		LOCAL lcT, lcV, loDb, lnFound, lcType
		lcT = toE.NewName()
		lcV = toE.NewName("frv")
		toE.CreateCustomers(lcT)
		This.Check(toE.RawRun("CREATE VIEW " + lcV + " AS SELECT id, name FROM " + lcT), "CREATE VIEW failed: " + toE.cLastError)
		loDb = This.Connected(toE)
		This.Check(loDb.GetObjects("c_fro"), "GetObjects() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_fro")
			SELECT c_fro
			This.Check(FIELD(1) == "NAME" AND TYPE("name") == "C" AND FSIZE("name") = 128 AND FIELD(2) == "TYPE" AND TYPE("type") == "C" AND FSIZE("type") = 5, ;
				"GetObjects() cursor is not name C(128), type C(5)")
			LOCATE FOR LOWER(ALLTRIM(name)) == lcT
			lcType = IIF(FOUND(), ALLTRIM(type), "<missing>")
			This.Check(lcType == "table", "GetObjects() has the table " + lcT + " as '" + lcType + "', expected 'table'")
			LOCATE FOR LOWER(ALLTRIM(name)) == lcV
			lcType = IIF(FOUND(), ALLTRIM(type), "<missing>")
			This.Check(lcType == "view", "GetObjects() has the view " + lcV + " as '" + lcType + "', expected 'view'")
			COUNT FOR !INLIST(ALLTRIM(type), "table", "view") TO lnFound
			This.Check(lnFound = 0, "GetObjects() has " + TRANSFORM(lnFound) + " rows whose type is neither table nor view")
			USE IN c_fro
		ENDIF
		This.Check(loDb.GetTables("c_frt"), "GetTables() failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_frt")
			SELECT c_frt
			COUNT FOR LOWER(ALLTRIM(name)) == lcV TO lnFound
			This.Check(lnFound = 0, "GetTables() lists the view " + lcV + ": it must keep giving only tables")
			COUNT FOR LOWER(ALLTRIM(name)) == lcT TO lnFound
			This.Check(lnFound = 1, "GetTables() does not list the table " + lcT)
			USE IN c_frt
		ENDIF
		loDb.Disconnect()
		toE.RawRun("DROP VIEW " + lcV)
		toE.DropTable(lcT)
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
	* A table whose name is quoted in the case the engine does not write by itself (lower case in
	* Firebird, upper case in PostgreSQL): GetColumns gives its columns and GetPrimaryKey its key, of
	* one column and of two.
	PROCEDURE TestATableQuotedInTheOtherCaseGivesItsColumnsAndKey() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("QuotedInTheOtherCase", "firebird,postgresql")
	ENDPROC

	PROCEDURE QuotedInTheOtherCase(toE)
		LOCAL lcT1, lcT2, loDb
		lcT1 = This.OtherCase(toE, toE.NewName("frq"))
		lcT2 = This.OtherCase(toE, toE.NewName("frq"))
		This.Check(toE.RawRun("CREATE TABLE " + ["] + lcT1 + ["] + " (id INTEGER NOT NULL PRIMARY KEY, nm VARCHAR(10))"), toE.cLastError)
		This.Check(toE.RawRun("CREATE TABLE " + ["] + lcT2 + ["] + " (a INTEGER NOT NULL, b INTEGER NOT NULL, v VARCHAR(10), PRIMARY KEY (a, b))"), toE.cLastError)
		loDb = This.Connected(toE)
		This.Check(loDb.GetColumns(lcT1, "c_frq") AND This.ColCount("c_frq") = 2, "GetColumns(" + lcT1 + "): " + TRANSFORM(This.ColCount("c_frq")) + " " + TRANSFORM(loDb.cLastError))
		IF USED("c_frq")
			SELECT c_frq
			LOCATE FOR LOWER(ALLTRIM(name)) == "id"
			This.Check(FOUND() AND primary_key, "GetColumns(" + lcT1 + ") does not mark id as primary_key")
		ENDIF
		This.Check(LOWER(STRTRAN(loDb.GetPrimaryKey(lcT1), " ", "")) == "id", "GetPrimaryKey(" + lcT1 + ") = '" + loDb.GetPrimaryKey(lcT1) + "'")
		This.Check(loDb.GetColumns(lcT2, "c_frq") AND This.ColCount("c_frq") = 3, "GetColumns(" + lcT2 + "): " + TRANSFORM(This.ColCount("c_frq")) + " " + TRANSFORM(loDb.cLastError))
		This.Check(LOWER(STRTRAN(loDb.GetPrimaryKey(lcT2), " ", "")) == "a,b", "GetPrimaryKey(" + lcT2 + ") = '" + loDb.GetPrimaryKey(lcT2) + "'")
		USE IN SELECT("c_frq")
		loDb.Disconnect()
		toE.DropTable(["] + lcT1 + ["])
		toE.DropTable(["] + lcT2 + ["])
	ENDPROC

	* A table created without quotes, asked for in lower case, upper case and mixed: the same as before.
	PROCEDURE TestATableWithoutQuotesIsFoundInAnyCase() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("UnquotedInAnyCase")
	ENDPROC

	PROCEDURE UnquotedInAnyCase(toE)
		LOCAL lcT, loDb, lnI, laAsk[3]
		lcT = toE.NewName()
		toE.CreateCustomers(lcT)
		laAsk[1] = LOWER(lcT)
		laAsk[2] = UPPER(lcT)
		laAsk[3] = PROPER(lcT)
		loDb = This.Connected(toE)
		FOR lnI = 1 TO 3
			This.Check(loDb.GetColumns(laAsk[lnI], "c_fru") AND This.ColCount("c_fru") = 6, "GetColumns(" + laAsk[lnI] + "): " + ;
				TRANSFORM(This.ColCount("c_fru")) + " " + TRANSFORM(loDb.cLastError))
			This.Check(LOWER(ALLTRIM(loDb.GetPrimaryKey(laAsk[lnI]))) == "id", "GetPrimaryKey(" + laAsk[lnI] + ") = '" + loDb.GetPrimaryKey(laAsk[lnI]) + "'")
		ENDFOR
		USE IN SELECT("c_fru")
		loDb.Disconnect()
		toE.DropTable(lcT)
	ENDPROC

	* Two tables whose names differ only in case (one created without quotes, one quoted in the
	* other case): each gives its own columns; a name that neither has exactly goes to the natural one.
	PROCEDURE TestTwoTablesThatDifferOnlyInCaseGiveTheirOwnColumns() HELP [Fact, Trait("Category", "Metadata"), Timeout(180000)]
		This.OnEachEngine("TwoTablesByCase", "firebird,postgresql")
	ENDPROC

	PROCEDURE TwoTablesByCase(toE)
		LOCAL lcN, lcNat, lcOther, loDb
		lcN = toE.NewName("frc")
		lcNat = IIF(toE.cName == "firebird", UPPER(lcN), LOWER(lcN))
		lcOther = This.OtherCase(toE, lcN)
		This.Check(toE.RawRun("CREATE TABLE " + lcN + " (id INTEGER NOT NULL PRIMARY KEY, x INTEGER)"), toE.cLastError)
		This.Check(toE.RawRun("CREATE TABLE " + ["] + lcOther + ["] + " (k INTEGER NOT NULL PRIMARY KEY, y INTEGER, z INTEGER)"), toE.cLastError)
		loDb = This.Connected(toE)
		This.Check(loDb.GetColumns(lcOther, "c_frc") AND This.ColCount("c_frc") = 3, "GetColumns(" + lcOther + "): " + TRANSFORM(This.ColCount("c_frc")))
		This.Check(LOWER(ALLTRIM(loDb.GetPrimaryKey(lcOther))) == "k", "GetPrimaryKey(" + lcOther + ") = '" + loDb.GetPrimaryKey(lcOther) + "'")
		This.Check(loDb.GetColumns(lcNat, "c_frc") AND This.ColCount("c_frc") = 2, "GetColumns(" + lcNat + "): " + TRANSFORM(This.ColCount("c_frc")))
		This.Check(LOWER(ALLTRIM(loDb.GetPrimaryKey(lcNat))) == "id", "GetPrimaryKey(" + lcNat + ") = '" + loDb.GetPrimaryKey(lcNat) + "'")
		This.Check(loDb.GetColumns(PROPER(lcN), "c_frc") AND This.ColCount("c_frc") = 2, "GetColumns(" + PROPER(lcN) + "): " + TRANSFORM(This.ColCount("c_frc")))
		USE IN SELECT("c_frc")
		loDb.Disconnect()
		toE.DropTable(lcN)
		toE.DropTable(["] + lcOther + ["])
	ENDPROC

	* The name in the case the engine does not write by itself.
	FUNCTION OtherCase(toE, tcName)
		RETURN IIF(toE.cName == "firebird", LOWER(tcName), UPPER(tcName))
	ENDFUNC

	FUNCTION ColCount(tcAlias)
		RETURN IIF(USED(tcAlias), RECCOUNT(tcAlias), -1)
	ENDFUNC
	* Open for writing a table whose name is passed between quotes, in the case the engine does not
	* write by itself: it finds the key, and Save() writes the change.
	PROCEDURE TestOpenForWritingATableQuotedInTheOtherCase() HELP [Fact, Trait("Category", "Save"), Timeout(180000)]
		This.OnEachEngine("OpenQuotedInTheOtherCase", "firebird,postgresql")
	ENDPROC

	PROCEDURE OpenQuotedInTheOtherCase(toE)
		LOCAL lcT, lcQ, loDb, llOk
		lcT = This.OtherCase(toE, toE.NewName("fro"))
		lcQ = ["] + lcT + ["]
		This.Check(toE.RawRun("CREATE TABLE " + lcQ + " (id INTEGER NOT NULL PRIMARY KEY, nm VARCHAR(10))"), toE.cLastError)
		This.Check(toE.RawRun("INSERT INTO " + lcQ + " (id, nm) VALUES (1, 'old')"), toE.cLastError)
		loDb = This.Connected(toE)
		This.Check(loDb.Open(lcQ, "c_fro"), "Open(" + lcQ + ") failed: " + TRANSFORM(loDb.cLastError))
		IF USED("c_fro")
			SELECT c_fro
			REPLACE nm WITH "new"
			llOk = loDb.Save("c_fro")
			This.Check(llOk, "Save() failed: " + TRANSFORM(loDb.cLastError))
			This.Check(RTRIM(TRANSFORM(toE.RawScalar("SELECT nm FROM " + lcQ + " WHERE id = 1"))) == "new", "on the server: " + ;
				TRANSFORM(toE.RawScalar("SELECT nm FROM " + lcQ + " WHERE id = 1")))
		ENDIF
		loDb.CloseAll()
		loDb.Disconnect()
		toE.DropTable(lcQ)
	ENDPROC
ENDDEFINE
