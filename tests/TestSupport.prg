* TestSupport.prg - helpers for FoxRemote's integration tests.
*
* FrEngine knows, for one engine, how to build the library object under test, how to reach
* the server WITHOUT the library (a raw ODBC connection, to check what really happened), and
* the few SQL differences the tests need. Credentials come from probe\engines.local.ini, which
* tests\setup-test-engines.ps1 fills in and git never sees.
*
* Environment:
*   FOXREMOTE_HOME          the FoxRemote folder (default: the author's checkout)
*   FOXREMOTE_LIB           the .prg under test (default: <home>\foxremote.prg); the red run
*                           against 0.x uses tests\legacy\LegacyShim.prg
*   FOXREMOTE_TEST_INI      the credentials file (default: <home>\probe\engines.local.ini)
*   FOXREMOTE_TEST_ENGINES  comma list (default: sqlite,mssql,mariadb,mysql,postgresql,firebird)
RETURN

FUNCTION FrHome
	LOCAL lcHome
	lcHome = GETENV("FOXREMOTE_HOME")
	RETURN ADDBS(IIF(EMPTY(lcHome), "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote", lcHome))
ENDFUNC

FUNCTION FrLibraryPath
	LOCAL lcLib
	lcLib = GETENV("FOXREMOTE_LIB")
	RETURN IIF(EMPTY(lcLib), FrHome() + "foxremote.prg", lcLib)
ENDFUNC

* The source of the library under test (for the red run, the 0.x behind the shim).
FUNCTION FrLibrarySource
	LOCAL lcLib
	lcLib = FrLibraryPath()
	RETURN IIF("LEGACYSHIM" $ UPPER(lcLib), FrHome() + "legacyoxremote.prg", lcLib)
ENDFUNC

PROCEDURE FrLoadLibrary
	LOCAL lcLib
	lcLib = FrLibraryPath()
	* The whole path: the stem alone ("foxremote") also matches the folder of TestSupport.prg.
	IF ATC(UPPER(FULLPATH(lcLib)), UPPER(SET("PROCEDURE"))) = 0
		SET PROCEDURE TO (lcLib) ADDITIVE
	ENDIF
ENDPROC

FUNCTION FrEngineList
	LOCAL lcList
	lcList = GETENV("FOXREMOTE_TEST_ENGINES")
	RETURN IIF(EMPTY(lcList), "sqlite,mssql,mariadb,mysql,postgresql,firebird", LOWER(lcList))
ENDFUNC

FUNCTION FrIniText
	LOCAL lcIni
	lcIni = GETENV("FOXREMOTE_TEST_INI")
	lcIni = IIF(EMPTY(lcIni), FrHome() + "probe\engines.local.ini", lcIni)
	RETURN IIF(FILE(lcIni), FILETOSTR(lcIni), "")
ENDFUNC

* A value of one section, tolerant of LF and CRLF.
FUNCTION FrIniValue(tcIni, tcSection, tcKey)
	LOCAL lcText, lnAt, lcSection
	lcText = CHR(10) + STRTRAN(tcIni, CHR(13), "") + CHR(10)
	lnAt = AT(CHR(10) + "[" + tcSection + "]" + CHR(10), lcText)
	IF lnAt = 0
		RETURN ""
	ENDIF
	lcSection = SUBSTR(lcText, lnAt + LEN(tcSection) + 3)
	IF AT(CHR(10) + "[", lcSection) > 0
		lcSection = LEFT(lcSection, AT(CHR(10) + "[", lcSection))
	ENDIF
	RETURN ALLTRIM(STREXTRACT(CHR(10) + lcSection, CHR(10) + tcKey + "=", CHR(10)))
ENDFUNC

* Hex of the UTF-8 bytes of a CP1252 string (what SQLite has to store).
FUNCTION FrUtf8Hex(tcText)
	RETURN STRCONV(STRCONV(tcText, 9), 15)
ENDFUNC

* BIGINT (COUNT(*) in MySQL/MariaDB) reaches VFP as text: compare by value.
FUNCTION FrNum(tvValue)
	DO CASE
	CASE VARTYPE(tvValue) == "N"
		RETURN tvValue
	CASE VARTYPE(tvValue) == "C" AND !EMPTY(tvValue) AND EMPTY(CHRTRAN(ALLTRIM(tvValue), "0123456789-", ""))
		RETURN VAL(tvValue)
	ENDCASE
	RETURN -1
ENDFUNC

* ======================================================================== *
DEFINE CLASS FrEngine AS Custom
	cName = ""
	cClass = ""
	cDriver = ""
	cServer = ""
	nPort = 0
	cUser = ""
	cPassword = ""
	cDatabase = ""
	cSemicolonUser = ""
	cSemicolonPassword = ""
	cDataDir = ""
	cLastError = ""

	PROCEDURE Init(tcName)
		LOCAL lcIni, lcSection
		This.cName = LOWER(tcName)
		lcIni = FrIniText()
		lcSection = This.cName
		This.cDataDir = ADDBS(FrIniValue(lcIni, "tests", "data_dir"))
		This.cDriver = FrIniValue(lcIni, lcSection, "driver")
		This.cServer = FrIniValue(lcIni, lcSection, "server")
		This.nPort = VAL(FrIniValue(lcIni, lcSection, "port"))
		This.cUser = FrIniValue(lcIni, lcSection, "test_user")
		This.cPassword = FrIniValue(lcIni, lcSection, "test_password")
		This.cDatabase = FrIniValue(lcIni, lcSection, IIF(This.cName == "sqlite", "database", "test_database"))
		This.cSemicolonUser = FrIniValue(lcIni, lcSection, "user")
		This.cSemicolonPassword = FrIniValue(lcIni, lcSection, "password")
		DO CASE
		CASE This.cName == "sqlite"
			This.cClass = "RemoteSqlite"
		CASE This.cName == "mssql"
			This.cClass = "RemoteSqlServer"
		CASE This.cName == "mariadb"
			This.cClass = "RemoteMariaDb"
		CASE This.cName == "mysql"
			This.cClass = "RemoteMySql"
		CASE This.cName == "postgresql"
			This.cClass = "RemotePostgreSql"
		CASE This.cName == "firebird"
			This.cClass = "RemoteFirebird"
		ENDCASE
	ENDPROC

	FUNCTION IsServer
		RETURN !INLIST(This.cName, "sqlite", "firebird")
	ENDFUNC

	* What cDatabase holds for a database name or file.
	FUNCTION DatabaseRef(tcDatabase)
		RETURN IIF(This.cName == "firebird", This.cServer + "/" + TRANSFORM(This.nPort) + ":" + tcDatabase, tcDatabase)
	ENDFUNC

	* The library object under test, configured for the test database and NOT connected.
	FUNCTION NewDb(tcDatabase)
		LOCAL loDb
		loDb = CREATEOBJECT(This.cClass)
		loDb.cDriver = This.cDriver
		loDb.cServer = This.cServer
		IF This.nPort > 0
			loDb.nPort = This.nPort
		ENDIF
		loDb.cUser = This.cUser
		loDb.cPassword = This.cPassword
		loDb.cDatabase = This.DatabaseRef(IIF(EMPTY(tcDatabase), This.cDatabase, tcDatabase))
		RETURN loDb
	ENDFUNC

	* ---- raw access, without the library -----------------------------------------
	FUNCTION ConnStr(tcDatabase)
		LOCAL lcDb
		lcDb = IIF(EMPTY(tcDatabase), This.cDatabase, tcDatabase)
		DO CASE
		CASE This.cName == "sqlite"
			RETURN "DRIVER={" + This.cDriver + "};DATABASE=" + lcDb + ";NoWCHAR=1;FKSupport=1;"
		CASE This.cName == "mssql"
			RETURN "DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";UID=" + This.cUser + ";PWD={" + This.cPassword + "};DATABASE=" + lcDb + ";"
		CASE INLIST(This.cName, "mysql", "mariadb")
			RETURN "DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";PORT=" + TRANSFORM(This.nPort) + ";UID=" + This.cUser + ;
				";PWD={" + This.cPassword + "};DATABASE=" + lcDb + ";"
		CASE This.cName == "postgresql"
			RETURN "DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";PORT=" + TRANSFORM(This.nPort) + ";UID=" + This.cUser + ;
				";PWD={" + This.cPassword + "};DATABASE=" + lcDb + ";"
		CASE This.cName == "firebird"
			RETURN "DRIVER={" + This.cDriver + "};DBNAME=" + This.DatabaseRef(lcDb) + ";UID=" + This.cUser + ";PWD={" + This.cPassword + "};CHARSET=WIN1252;LOCKTIMEOUT=10;"
		ENDCASE
		RETURN ""
	ENDFUNC

	* (Firebird: LOCKTIMEOUT, or a raw statement waits forever behind a transaction the library left open.)

	* The database to connect to when the test database is not the point (to create or drop one).
	FUNCTION MaintenanceDatabase
		DO CASE
		CASE This.cName == "mssql"
			RETURN "master"
		CASE INLIST(This.cName, "mysql", "mariadb")
			RETURN "information_schema"
		CASE This.cName == "postgresql"
			RETURN "postgres"
		ENDCASE
		RETURN This.cDatabase
	ENDFUNC

	FUNCTION Adapt(tcSql)
		RETURN IIF(This.cName == "firebird", STRTRAN(tcSql, "RTRIM(", "TRIM(TRAILING FROM "), tcSql)
	ENDFUNC

	FUNCTION RawRun(tcSql, tcDatabase)
		LOCAL lnH, lnR, laErr[1]
		lnH = SQLSTRINGCONNECT(This.ConnStr(tcDatabase), .T.)
		IF lnH < 1
			=AERROR(laErr)
			This.cLastError = "raw connect: " + TRANSFORM(laErr[2])
			RETURN .F.
		ENDIF
		=SQLSETPROP(lnH, "QueryTimeOut", 20)
		lnR = SQLEXEC(lnH, This.Adapt(tcSql))
		IF lnR < 1
			=AERROR(laErr)
			This.cLastError = tcSql + ": " + TRANSFORM(laErr[2])
		ENDIF
		=SQLDISCONNECT(lnH)
		RETURN lnR > 0
	ENDFUNC

	* First column of the first row; .NULL. when there is no row; "<error ...>" when the SQL failed.
	FUNCTION RawScalar(tcSql, tcDatabase)
		LOCAL lnH, lvValue, laErr[1], lcAlias
		lvValue = .NULL.
		lnH = SQLSTRINGCONNECT(This.ConnStr(tcDatabase), .T.)
		IF lnH < 1
			=AERROR(laErr)
			RETURN "<raw connect: " + TRANSFORM(laErr[2]) + ">"
		ENDIF
		lcAlias = "c_fr_raw"
		=SQLSETPROP(lnH, "QueryTimeOut", 20)
		IF SQLEXEC(lnH, This.Adapt(tcSql), lcAlias) > 0 AND USED(lcAlias)
			SELECT (lcAlias)
			IF !EOF()
				lvValue = EVALUATE(FIELD(1))
			ENDIF
			USE IN (lcAlias)
		ELSE
			=AERROR(laErr)
			lvValue = "<error " + TRANSFORM(laErr[2]) + ">"
		ENDIF
		=SQLDISCONNECT(lnH)
		RETURN lvValue
	ENDFUNC

	FUNCTION RawCount(tcTable, tcWhere)
		RETURN FrNum(This.RawScalar("SELECT COUNT(*) FROM " + tcTable + IIF(EMPTY(tcWhere), "", " WHERE " + tcWhere)))
	ENDFUNC

	* ---- names and fixtures -------------------------------------------------------
	FUNCTION NewName(tcPrefix)
		RETURN LOWER(IIF(EMPTY(tcPrefix), "fr", tcPrefix) + SUBSTR(SYS(2015), 2))
	ENDFUNC

	FUNCTION TextType
		DO CASE
		CASE This.cName == "mssql"
			RETURN "VARCHAR(MAX)"
		CASE This.cName == "firebird"
			RETURN "BLOB SUB_TYPE TEXT"
		ENDCASE
		RETURN "TEXT"
	ENDFUNC

	FUNCTION AutoIdType
		DO CASE
		CASE This.cName == "sqlite"
			RETURN "INTEGER PRIMARY KEY AUTOINCREMENT"
		CASE This.cName == "mssql"
			RETURN "INT IDENTITY(1,1) PRIMARY KEY"
		CASE INLIST(This.cName, "mysql", "mariadb")
			RETURN "INT AUTO_INCREMENT PRIMARY KEY"
		CASE This.cName == "postgresql"
			RETURN "SERIAL PRIMARY KEY"
		CASE This.cName == "firebird"
			RETURN "INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY"
		ENDCASE
		RETURN ""
	ENDFUNC

	* The table most tests use: id, name, city, born (nullable date), qty (CHECK >= 0), notes (long text).
	FUNCTION CreateCustomers(tcTable)
		RETURN This.RawRun("CREATE TABLE " + tcTable + " (id " + This.AutoIdType() + ", name VARCHAR(40), city VARCHAR(40), born DATE, " + ;
			"qty INTEGER CHECK (qty >= 0), notes " + This.TextType() + ")")
	ENDFUNC

	FUNCTION InsertCustomer(tcTable, tcName, tcCity)
		RETURN This.RawRun("INSERT INTO " + tcTable + " (name, city) VALUES ('" + tcName + "', '" + tcCity + "')")
	ENDFUNC

	FUNCTION MaxId(tcTable)
		RETURN FrNum(This.RawScalar("SELECT MAX(id) FROM " + tcTable))
	ENDFUNC

	PROCEDURE DropTable(tcTable)
		This.RawRun("DROP TABLE " + tcTable)
	ENDPROC

	* Length in characters as stored, trailing blanks included.
	FUNCTION LengthSql(tcColumn)
		DO CASE
		CASE This.cName == "mssql"
			RETURN "DATALENGTH(" + tcColumn + ")"
		CASE This.cName == "firebird"
			RETURN "CHAR_LENGTH(" + tcColumn + ")"
		ENDCASE
		RETURN "LENGTH(" + tcColumn + ")"
	ENDFUNC

	* ---- sessions and locks (reconnection and NextNumber tests) ------------------
	FUNCTION SessionIdSql
		DO CASE
		CASE This.cName == "mssql"
			RETURN "SELECT @@SPID"
		CASE INLIST(This.cName, "mysql", "mariadb")
			RETURN "SELECT CONNECTION_ID()"
		CASE This.cName == "postgresql"
			RETURN "SELECT pg_backend_pid()"
		CASE This.cName == "firebird"
			RETURN "SELECT CURRENT_CONNECTION FROM RDB$DATABASE"
		ENDCASE
		RETURN ""
	ENDFUNC

	* Ends a session from another connection, as a server that drops an idle one does, and waits
	* until the server has let it go.
	FUNCTION KillSession(tnId)
		LOCAL lcId, lcAlive, llOk, lnTry
		lcId = TRANSFORM(INT(tnId))
		DO CASE
		CASE This.cName == "mssql"
			llOk = This.AdminRun("KILL " + lcId)
			lcAlive = "SELECT COUNT(*) FROM sys.dm_exec_sessions WHERE session_id = " + lcId
		CASE INLIST(This.cName, "mysql", "mariadb")
			llOk = This.RawRun("KILL " + lcId)
			lcAlive = "SELECT COUNT(*) FROM information_schema.PROCESSLIST WHERE ID = " + lcId
		CASE This.cName == "postgresql"
			llOk = This.RawRun("SELECT pg_terminate_backend(" + lcId + ")")
			lcAlive = "SELECT COUNT(*) FROM pg_stat_activity WHERE pid = " + lcId
		CASE This.cName == "firebird"
			llOk = This.RawRun("DELETE FROM MON$ATTACHMENTS WHERE MON$ATTACHMENT_ID = " + lcId)
			lcAlive = "SELECT COUNT(*) FROM MON$ATTACHMENTS WHERE MON$ATTACHMENT_ID = " + lcId
		OTHERWISE
			RETURN .F.
		ENDCASE
		DECLARE Sleep IN kernel32 INTEGER
		FOR lnTry = 1 TO 50
			IF This.cName == "mssql"
				IF FrNum(This.AdminScalar(lcAlive)) = 0
					EXIT
				ENDIF
			ELSE
				IF FrNum(This.RawScalar(lcAlive)) = 0
					EXIT
				ENDIF
			ENDIF
			Sleep(100)
		ENDFOR
		RETURN llOk
	ENDFUNC

	* SQL Server: KILL and the session list need an administrator (Windows authentication).
	FUNCTION AdminRun(tcSql)
		LOCAL lnH, lnR, laErr[1]
		lnH = SQLSTRINGCONNECT("DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";Trusted_Connection=yes;DATABASE=master;", .T.)
		IF lnH < 1
			This.cLastError = "admin connect failed"
			RETURN .F.
		ENDIF
		lnR = SQLEXEC(lnH, tcSql)
		IF lnR < 1
			=AERROR(laErr)
			This.cLastError = tcSql + ": " + TRANSFORM(laErr[2])
		ENDIF
		=SQLDISCONNECT(lnH)
		RETURN lnR > 0
	ENDFUNC

	FUNCTION AdminScalar(tcSql)
		LOCAL lnH, lvValue
		lvValue = .NULL.
		lnH = SQLSTRINGCONNECT("DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";Trusted_Connection=yes;DATABASE=master;", .T.)
		IF lnH > 0
			IF SQLEXEC(lnH, tcSql, "c_fr_adm") > 0 AND USED("c_fr_adm")
				lvValue = EVALUATE("c_fr_adm." + FIELD(1, "c_fr_adm"))
				USE IN c_fr_adm
			ENDIF
			=SQLDISCONNECT(lnH)
		ENDIF
		RETURN lvValue
	ENDFUNC

	* Connection options that make a lock wait give up after about two seconds (set before Connect).
	FUNCTION LockWaitOptions
		DO CASE
		CASE This.cName == "firebird"
			RETURN "LOCKTIMEOUT=2;"
		CASE This.cName == "sqlite"
			RETURN "Timeout=2000;"
		ENDCASE
		RETURN ""
	ENDFUNC

	* The same, for the engines that set it on the open connection.
	PROCEDURE ShortLockWait(toDb)
		DO CASE
		CASE This.cName == "mssql"
			toDb.Execute("SET LOCK_TIMEOUT 2000")
		CASE INLIST(This.cName, "mysql", "mariadb")
			toDb.Execute("SET SESSION innodb_lock_wait_timeout = 2")
		CASE This.cName == "postgresql"
			toDb.Execute("SET lock_timeout = '2s'")
		CASE This.cName == "firebird"
			* LOCKTIMEOUT in the connection string does not reach a manual transaction (measured 04-10)
			toDb.Execute("SET STATEMENT TIMEOUT 2 SECOND")
		ENDCASE
	ENDPROC

	* ---- databases ---------------------------------------------------------------
	FUNCTION NewDatabaseName
		IF INLIST(This.cName, "sqlite", "firebird")
			RETURN This.cDataDir + This.NewName("db") + IIF(This.cName == "sqlite", ".db", ".fdb")
		ENDIF
		RETURN This.NewName("frdb")
	ENDFUNC

	FUNCTION DatabaseExists(tcDatabase)
		DO CASE
		CASE INLIST(This.cName, "sqlite", "firebird")
			RETURN FILE(tcDatabase)
		CASE This.cName == "mssql"
			RETURN FrNum(This.RawScalar("SELECT COUNT(*) FROM sys.databases WHERE name = '" + tcDatabase + "'", "master")) > 0
		CASE INLIST(This.cName, "mysql", "mariadb")
			RETURN FrNum(This.RawScalar("SELECT COUNT(*) FROM information_schema.schemata WHERE schema_name = '" + tcDatabase + "'", "information_schema")) > 0
		CASE This.cName == "postgresql"
			RETURN FrNum(This.RawScalar("SELECT COUNT(*) FROM pg_database WHERE datname = '" + tcDatabase + "'", "postgres")) > 0
		ENDCASE
		RETURN .F.
	ENDFUNC

	PROCEDURE DropDatabase(tcDatabase)
		IF !This.DatabaseExists(tcDatabase)
			RETURN
		ENDIF
		DO CASE
		CASE INLIST(This.cName, "sqlite", "firebird")
			TRY
				ERASE (tcDatabase)
			CATCH
			ENDTRY
		CASE This.cName == "mssql"
			This.RawRun("ALTER DATABASE " + tcDatabase + " SET SINGLE_USER WITH ROLLBACK IMMEDIATE", "master")
			This.RawRun("DROP DATABASE " + tcDatabase, "master")
		CASE INLIST(This.cName, "mysql", "mariadb")
			This.RawRun("DROP DATABASE IF EXISTS " + tcDatabase, "information_schema")
		CASE This.cName == "postgresql"
			This.RawRun("DROP DATABASE IF EXISTS " + tcDatabase + " WITH (FORCE)", "postgres")
		ENDCASE
	ENDPROC
ENDDEFINE
