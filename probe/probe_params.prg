* probe_params.prg - how a parameter reaches the server through FoxRemote's Query() and Execute().
* Measures, it does not assert: every line of the log says what the server did.
*   DATES     a DATE and a TIMESTAMP column, one row written with literals by a raw connection and
*             one written by Save(), and which of the two each WHERE with :d or :t finds.
*   LIMIT     MySQL and MariaDB (MariaDB Connector/ODBC): LIMIT :n.
*   TOP       SQL Server: TOP (:n), FETCH NEXT :n ROWS ONLY and the same inside CAST(:n AS INT).
*   BOOLEAN   Firebird: .NULL. as the parameter of a BOOLEAN column, and :p IS NULL.
* Run:     foxproof run --prg probe\probe_params.prg
* Log:     FR_PROBE_LOG (default probe\probe-params.log), one line per measure: engine|probe|what|result.
* Engines and credentials as the integration tests: tests\TestSupport.prg, probe\engines.local.ini,
* FOXREMOTE_TEST_ENGINES and FOXREMOTE_LIB.
DEFINE CLASS ProbeParams AS Custom
	cLog = ""
	cEngine = ""
	* Tables to drop once the library has disconnected: Firebird's DROP TABLE waits for a
	* connection that has used the table.
	cDrop = ""

	PROCEDURE TestProbe() HELP [Fact, Trait("Category", "Probe"), Timeout(300000)]
		LOCAL laNames[1], lnI, loE, loDb
		SET PROCEDURE TO "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\tests\TestSupport.prg" ADDITIVE
		IF !EMPTY(GETENV("FOXREMOTE_HOME"))
			SET PROCEDURE TO (ADDBS(GETENV("FOXREMOTE_HOME")) + "tests\TestSupport.prg") ADDITIVE
		ENDIF
		FrLoadLibrary()
		This.cLog = IIF(EMPTY(GETENV("FR_PROBE_LOG")), FrHome() + "probe\probe-params.log", GETENV("FR_PROBE_LOG"))
		This.cEngine = "-"
		This.Say("START", "library", FrLibraryPath() + " at " + TTOC(DATETIME(), 3))
		=ALINES(laNames, FrEngineList(), 1, ",")
		FOR lnI = 1 TO ALEN(laNames)
			This.cEngine = ALLTRIM(laNames[lnI])
			loE = CREATEOBJECT("FrEngine", This.cEngine)
			This.cDrop = ""
			loDb = loE.NewDb()
			IF !loDb.Connect()
				This.Say("CONNECT", "Connect()", "FAIL " + loDb.cLastErrorCode + ": " + loDb.cLastError)
				LOOP
			ENDIF
			This.Say("CONNECT", "server", TRANSFORM(loDb.ServerVersion()))
			This.Dates(loE, loDb)
			DO CASE
			CASE INLIST(This.cEngine, "mysql", "mariadb")
				This.LimitRows(loE, loDb)
			CASE This.cEngine == "mssql"
				This.TopFetch(loE, loDb)
			CASE This.cEngine == "firebird"
				This.BooleanNull(loE, loDb)
			ENDCASE
			loDb.CloseAll()
			loDb.Disconnect()
			This.DropAll(loE)
		ENDFOR
		This.cEngine = "-"
		This.Say("END", "", TTOC(DATETIME(), 3))
	ENDPROC

	* ---- DATES ------------------------------------------------------------------------
	PROCEDURE Dates(toE, toDb)
		LOCAL lcT, lcTs, loP, lnId, laF[1]
		lcT = toE.NewName()
		lcTs = IIF(INLIST(This.cEngine, "mssql", "mysql", "mariadb"), "DATETIME", "TIMESTAMP")
		IF !toE.RawRun("CREATE TABLE " + lcT + " (id " + toE.AutoIdType() + ", d DATE, ts " + lcTs + ")")
			This.Say("DATES", "CREATE TABLE", "FAIL " + toE.cLastError)
			RETURN
		ENDIF
		* Row 1: the literals another tool would write. Row 2: the same values, written by Save().
		This.Say("DATES", "row 1 by a raw INSERT of '2026-11-02' and '2026-11-02 10:30:00'", ;
			IIF(toE.RawRun("INSERT INTO " + lcT + " (d, ts) VALUES ('2026-11-02', '2026-11-02 10:30:00')"), "OK", "FAIL " + toE.cLastError))
		toDb.Open(lcT, "c_frpd", "1 = 0")
		IF USED("c_frpd")
			=AFIELDS(laF, "c_frpd")
			This.Say("DATES", "Open() types of d and ts", laF[2, 2] + " " + laF[3, 2])
			SELECT c_frpd
			APPEND BLANK
			REPLACE d WITH {^2026-11-02}, ts WITH {^2026-11-02 10:30:00}
			This.Say("DATES", "row 2 by Save() of {^2026-11-02} and {^2026-11-02 10:30:00}", ;
				IIF(toDb.Save("c_frpd"), "OK", "FAIL " + toDb.cLastErrorCode + ": " + toDb.cLastError))
			toDb.Close("c_frpd")
		ELSE
			This.Say("DATES", "Open()", "FAIL " + toDb.cLastErrorCode + ": " + toDb.cLastError)
		ENDIF
		FOR lnId = 1 TO 2
			This.Say("DATES", "row " + TRANSFORM(lnId) + " id", TRANSFORM(toE.RawScalar("SELECT id FROM " + lcT + " WHERE id = " + TRANSFORM(lnId))))
		ENDFOR
		IF This.cEngine == "sqlite"
			FOR lnId = 1 TO 2
				This.Say("DATES", "row " + TRANSFORM(lnId) + " as stored (raw: typeof and text of d, ts)", ;
					TRANSFORM(toE.RawScalar("SELECT typeof(d) || ' [' || d || '] ' || typeof(ts) || ' [' || ts || ']' FROM " + lcT + " WHERE id = " + TRANSFORM(lnId))))
			ENDFOR
		ENDIF
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "d", {^2026-11-02})
		ADDPROPERTY(loP, "d2", {^2026-11-03})
		ADDPROPERTY(loP, "t", {^2026-11-02 10:30:00})
		ADDPROPERTY(loP, "t2", {^2026-11-02 10:30:01})
		IF This.cEngine == "sqlite"
			This.Rows("DATES", "how :d and :t arrive: SELECT typeof(:d), :d, typeof(:t), :t", ;
				"SELECT typeof(:d) || ' [' || :d || '] ' || typeof(:t) || ' [' || :t || ']' AS v", toDb, loP, "v")
		ENDIF
		This.Rows("DATES", "d = '2026-11-02' (a literal, no parameter)", "SELECT id FROM " + lcT + " WHERE d = '2026-11-02' ORDER BY id", toDb, loP)
		This.Rows("DATES", "d = :d", "SELECT id FROM " + lcT + " WHERE d = :d ORDER BY id", toDb, loP)
		This.Rows("DATES", "d >= :d AND d < :d2 (the next day)", "SELECT id FROM " + lcT + " WHERE d >= :d AND d < :d2 ORDER BY id", toDb, loP)
		This.Rows("DATES", "ts = :t", "SELECT id FROM " + lcT + " WHERE ts = :t ORDER BY id", toDb, loP)
		This.Rows("DATES", "ts >= :t AND ts < :t2 (one second later)", "SELECT id FROM " + lcT + " WHERE ts >= :t AND ts < :t2 ORDER BY id", toDb, loP)
		This.Rows("DATES", "ts >= :d AND ts < :d2 (the whole day)", "SELECT id FROM " + lcT + " WHERE ts >= :d AND ts < :d2 ORDER BY id", toDb, loP)
		IF This.cEngine == "sqlite"
			This.Rows("DATES", "date(d) = :d", "SELECT id FROM " + lcT + " WHERE date(d) = :d ORDER BY id", toDb, loP)
			This.Rows("DATES", "date(d) = date(:d)", "SELECT id FROM " + lcT + " WHERE date(d) = date(:d) ORDER BY id", toDb, loP)
			This.Rows("DATES", "datetime(ts) = :t", "SELECT id FROM " + lcT + " WHERE datetime(ts) = :t ORDER BY id", toDb, loP)
			This.Rows("DATES", "datetime(ts) = datetime(:t)", "SELECT id FROM " + lcT + " WHERE datetime(ts) = datetime(:t) ORDER BY id", toDb, loP)
			This.Rows("DATES", "date(ts) = date(:d)", "SELECT id FROM " + lcT + " WHERE date(ts) = date(:d) ORDER BY id", toDb, loP)
		ENDIF
		This.cDrop = This.cDrop + "," + lcT
	ENDPROC

	* ---- LIMIT (MySQL and MariaDB) ------------------------------------------------------------
	PROCEDURE LimitRows(toE, toDb)
		LOCAL lcT, loP
		lcT = This.ThreeRows(toE)
		IF EMPTY(lcT)
			RETURN
		ENDIF
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "n", 2)
		ADDPROPERTY(loP, "c", "2")
		This.Rows("LIMIT", "LIMIT 2 (written in the SQL)", "SELECT id FROM " + lcT + " ORDER BY id LIMIT 2", toDb, loP)
		This.Rows("LIMIT", "LIMIT :n with n = 2 (N)", "SELECT id FROM " + lcT + " ORDER BY id LIMIT :n", toDb, loP)
		This.Rows("LIMIT", "LIMIT :c with c = '2' (C)", "SELECT id FROM " + lcT + " ORDER BY id LIMIT :c", toDb, loP)
		This.cDrop = This.cDrop + "," + lcT
	ENDPROC

	* ---- TOP and FETCH (SQL Server) ------------------------------------------------------------
	PROCEDURE TopFetch(toE, toDb)
		LOCAL lcT, loP
		lcT = This.ThreeRows(toE)
		IF EMPTY(lcT)
			RETURN
		ENDIF
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "n", 2)
		This.Rows("TOP", "TOP (2) (written in the SQL)", "SELECT TOP (2) id FROM " + lcT + " ORDER BY id", toDb, loP)
		This.Rows("TOP", "TOP (:n) with n = 2 (N)", "SELECT TOP (:n) id FROM " + lcT + " ORDER BY id", toDb, loP)
		This.Rows("TOP", "TOP (CAST(:n AS INT))", "SELECT TOP (CAST(:n AS INT)) id FROM " + lcT + " ORDER BY id", toDb, loP)
		This.Rows("TOP", "OFFSET 0 ROWS FETCH NEXT :n ROWS ONLY", "SELECT id FROM " + lcT + " ORDER BY id OFFSET 0 ROWS FETCH NEXT :n ROWS ONLY", toDb, loP)
		This.Rows("TOP", "OFFSET 0 ROWS FETCH NEXT CAST(:n AS INT) ROWS ONLY", "SELECT id FROM " + lcT + " ORDER BY id OFFSET 0 ROWS FETCH NEXT CAST(:n AS INT) ROWS ONLY", toDb, loP)
		This.cDrop = This.cDrop + "," + lcT
	ENDPROC

	* ---- BOOLEAN (Firebird) -----------------------------------------------------------------
	PROCEDURE BooleanNull(toE, toDb)
		LOCAL lcT, loP
		lcT = toE.NewName()
		IF !toE.RawRun("CREATE TABLE " + lcT + " (id " + toE.AutoIdType() + ", flag BOOLEAN)")
			This.Say("BOOLEAN", "CREATE TABLE", "FAIL " + toE.cLastError)
			RETURN
		ENDIF
		loP = CREATEOBJECT("Empty")
		ADDPROPERTY(loP, "f", .NULL.)
		ADDPROPERTY(loP, "t", .T.)
		This.RunSql("BOOLEAN", "INSERT VALUES (:f) with f = .NULL.", "INSERT INTO " + lcT + " (flag) VALUES (:f)", toDb, loP)
		This.RunSql("BOOLEAN", "INSERT VALUES (CAST(:f AS BOOLEAN)) with f = .NULL.", "INSERT INTO " + lcT + " (flag) VALUES (CAST(:f AS BOOLEAN))", toDb, loP)
		This.RunSql("BOOLEAN", "INSERT VALUES (NULL) (written in the SQL)", "INSERT INTO " + lcT + " (flag) VALUES (NULL)", toDb, loP)
		This.RunSql("BOOLEAN", "INSERT VALUES (:t) with t = .T.", "INSERT INTO " + lcT + " (flag) VALUES (:t)", toDb, loP)
		This.Rows("BOOLEAN", "flag IS NULL (written in the SQL)", "SELECT id FROM " + lcT + " WHERE flag IS NULL ORDER BY id", toDb, loP)
		This.Rows("BOOLEAN", "flag = :t with t = .T.", "SELECT id FROM " + lcT + " WHERE flag = :t ORDER BY id", toDb, loP)
		This.Rows("BOOLEAN", ":f IS NULL with f = .NULL.", "SELECT id FROM " + lcT + " WHERE :f IS NULL ORDER BY id", toDb, loP)
		This.Rows("BOOLEAN", "CAST(:f AS BOOLEAN) IS NULL with f = .NULL.", "SELECT id FROM " + lcT + " WHERE CAST(:f AS BOOLEAN) IS NULL ORDER BY id", toDb, loP)
		This.cDrop = This.cDrop + "," + lcT
	ENDPROC

	* ---- helpers --------------------------------------------------------------------------------
	PROCEDURE DropAll(toE)
		LOCAL laT[1], lnI
		FOR lnI = 1 TO ALINES(laT, This.cDrop, 5, ",")
			This.Say("CLEANUP", "DROP TABLE " + laT[lnI], IIF(toE.RawRun("DROP TABLE " + laT[lnI]), "OK", "FAIL " + toE.cLastError))
		ENDFOR
		This.cDrop = ""
	ENDPROC

	FUNCTION ThreeRows(toE)
		LOCAL lcT, lnI
		lcT = toE.NewName()
		IF !toE.RawRun("CREATE TABLE " + lcT + " (id " + toE.AutoIdType() + ", name VARCHAR(10))")
			This.Say("SETUP", "CREATE TABLE", "FAIL " + toE.cLastError)
			RETURN ""
		ENDIF
		FOR lnI = 1 TO 3
			toE.RawRun("INSERT INTO " + lcT + " (name) VALUES ('r" + TRANSFORM(lnI) + "')")
		ENDFOR
		RETURN lcT
	ENDFUNC

	* Query() and what came back: the ids (or the field tcField) of every row, or the failure.
	PROCEDURE Rows(tcProbe, tcWhat, tcSql, toDb, toP, tcField)
		LOCAL lcOut, loEx, llOk
		lcOut = ""
		TRY
			llOk = toDb.Query(tcSql, "c_frpq", toP)
			IF llOk AND USED("c_frpq")
				SELECT c_frpq
				SCAN
					lcOut = lcOut + IIF(EMPTY(lcOut), "", ",") + ALLTRIM(TRANSFORM(EVALUATE(IIF(EMPTY(tcField), "id", tcField))))
				ENDSCAN
				lcOut = "OK " + TRANSFORM(RECCOUNT("c_frpq")) + " rows [" + lcOut + "]"
				USE IN c_frpq
			ELSE
				lcOut = "FAIL " + toDb.cLastErrorCode + ": " + toDb.cLastError
			ENDIF
		CATCH TO loEx
			lcOut = "RAISED " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message
		ENDTRY
		This.Say(tcProbe, tcWhat, lcOut)
	ENDPROC

	* Execute() and what came back.
	PROCEDURE RunSql(tcProbe, tcWhat, tcSql, toDb, toP)
		LOCAL lcOut, loEx, lnR
		TRY
			lnR = toDb.Execute(tcSql, toP)
			lcOut = IIF(lnR >= 0, "OK " + TRANSFORM(lnR) + " rows", "FAIL " + toDb.cLastErrorCode + ": " + toDb.cLastError)
		CATCH TO loEx
			lcOut = "RAISED " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message
		ENDTRY
		This.Say(tcProbe, tcWhat, lcOut)
	ENDPROC

	PROCEDURE Say(tcProbe, tcWhat, tcResult)
		=STRTOFILE(This.cEngine + "|" + tcProbe + "|" + tcWhat + "|" + CHRTRAN(tcResult, CHR(13) + CHR(10), "  ") + CHR(13) + CHR(10), This.cLog, 1)
	ENDPROC
ENDDEFINE
