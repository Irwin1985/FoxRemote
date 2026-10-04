* fr_probe.prg - integration probe of FoxRemote 0.x against a real engine.
* Run through a wrapper: DO fr_probe WITH "sqlite"   (result in gcFoxResult)
* Each probe measures one fault of RELEVO-FOXREMOTE-20261003.md section 4.
* Verdicts: BUG (fault confirmed), OK (fault not reproduced), INFO, ERROR (probe itself failed).
LPARAMETERS tcEngine
LOCAL loProbe, lcResultFile
SET PROCEDURE TO foxremote.prg ADDITIVE
=SQLSETPROP(0, "DispLogin", 3)
=SQLSETPROP(0, "ConnectTimeOut", 15)
loProbe = CREATEOBJECT("Probe", tcEngine)
gcFoxResult = loProbe.Run()
lcResultFile = FULLPATH("result_" + LOWER(tcEngine) + ".txt")
IF FILE(lcResultFile)
	ERASE (lcResultFile)
ENDIF
=STRTOFILE(gcFoxResult, lcResultFile)
RETURN

* ======================================================================== *
DEFINE CLASS Probe AS Custom
	oT = .NULL.
	cOut = ""

	PROCEDURE Init(tcEngine)
		DO CASE
		CASE LOWER(tcEngine) == "sqlite"
			This.oT = CREATEOBJECT("SqliteTarget")
		CASE LOWER(tcEngine) == "mssql"
			This.oT = CREATEOBJECT("MssqlTarget")
		CASE LOWER(tcEngine) == "mariadb"
			This.oT = CREATEOBJECT("MariadbTarget")
		CASE LOWER(tcEngine) == "postgresql"
			This.oT = CREATEOBJECT("PostgresTarget")
		CASE LOWER(tcEngine) == "firebird"
			This.oT = CREATEOBJECT("FirebirdTarget")
		OTHERWISE
			ERROR "Unknown engine: " + tcEngine
		ENDCASE
	ENDPROC

	FUNCTION Run
		LOCAL lcDir, laProbes[1], i, lcName, lcSafety
		* The probe overwrites only its own scratch files; a live session may have
		* SAFETY ON, and its "overwrite?" dialog would stop the probe.
		lcSafety = SET("SAFETY")
		SET SAFETY OFF
		lcDir = ADDBS(FULLPATH(""))
		IF !This.oT.Setup(lcDir)
			RETURN "SETUP|ERROR|" + This.oT.cLastError
		ENDIF
		This.Say("ENGINE", "INFO", This.oT.cName + " " + This.oT.ServerVersion())
		TRY
			This.oT.ExtraChecks(This)
		CATCH TO loEx
			This.Say("EXTRA", "ERROR", "probe raised " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message + " (" + loEx.Procedure + " line " + TRANSFORM(loEx.LineNo) + ")")
		ENDTRY
		=ALINES(laProbes, "S1,N1,A1,F3,F11,F1a,F1b,F1c,F12,F2,F5,F6,F6b,F7,F8a,F8b,F9,F10,F4a,F4b", ",")
		FOR i = 1 TO ALEN(laProbes)
			lcName = laProbes[i]
			TRY
				This.&lcName.()
			CATCH TO loEx
				This.Say(lcName, "ERROR", "probe raised " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message + " (" + loEx.Procedure + " line " + TRANSFORM(loEx.LineNo) + ": " + ALLTRIM(loEx.LineContents) + ")")
			ENDTRY
			CLOSE TABLES ALL
		ENDFOR
		This.oT.Teardown()
		SET SAFETY &lcSafety
		RETURN This.cOut
	ENDFUNC

	PROCEDURE Say(tcId, tcVerdict, tcDetail)
		This.cOut = This.cOut + tcId + "|" + tcVerdict + "|" + CHRTRAN(tcDetail, CHR(13) + CHR(10), "  ") + CHR(13) + CHR(10)
	ENDPROC

	* Handle FoxRemote is using (nHandle is HIDDEN): the newest live handle.
	FUNCTION LastHandle
		LOCAL laH[1], lnN, i, lnMax
		lnMax = 0
		lnN = ASQLHANDLES(laH)
		FOR i = 1 TO lnN
			lnMax = MAX(lnMax, laH[i])
		ENDFOR
		RETURN lnMax
	ENDFUNC

	FUNCTION Short(tcText)
		RETURN LEFT(ALLTRIM(CHRTRAN(tcText, CHR(13) + CHR(10) + CHR(9), "   ")), 220)
	ENDFUNC

	* BIGINT (COUNT(*) in MySQL/MariaDB) reaches VFP as text: compare by value.
	FUNCTION Num(tvValue)
		DO CASE
		CASE VARTYPE(tvValue) == "N"
			RETURN tvValue
		CASE VARTYPE(tvValue) == "C" AND !EMPTY(tvValue) AND EMPTY(CHRTRAN(ALLTRIM(tvValue), "0123456789-", ""))
			RETURN VAL(tvValue)
		ENDCASE
		RETURN -1
	ENDFUNC

	* --- S1: computed columns through FoxRemote's own connection -----------------
	PROCEDURE S1
		LOCAL loDb, lcVer
		loDb = This.oT.NewRemote()
		loDb.Connect()
		lcVer = TRANSFORM(loDb.getVersion())
		This.Say("S1", IIF(CHR(0) $ lcVer, "BUG", "OK"), "getVersion() = '" + STRTRAN(lcVer, CHR(0), "<NUL>") + "' (" + TRANSFORM(LEN(lcVer)) + " chars, " + TRANSFORM(OCCURS(CHR(0), lcVer)) + " CHR(0))")
		loDb.disconnect()
	ENDPROC

	* --- A1: accented text (CP1252) there and back ----------------------------------------
	PROCEDURE A1
		LOCAL loDb, lcTxt, lvSave, lnId, lvRaw, lcBack
		lcTxt = "Pe" + CHR(241) + "a " + CHR(193) + "lvaro " + CHR(128) + "5"
		loDb = This.oT.NewRemote()
		loDb.Connect()
		loDb.use("customers AS ac_a1", "", "1 = 0")
		SELECT ac_a1
		APPEND BLANK
		REPLACE name WITH lcTxt
		lvSave = loDb.Save("ac_a1")
		loDb.closeAll()
		lnId = This.oT.RawScalar("SELECT MAX(id) FROM customers")
		lvRaw = This.oT.RawScalar("SELECT RTRIM(name) FROM customers WHERE id = " + TRANSFORM(lnId))
		loDb.use("customers AS ac_a1b", "", "id = " + TRANSFORM(lnId))
		SELECT ac_a1b
		lcBack = IIF(EOF(), "<no row>", ALLTRIM(name))
		loDb.closeAll()
		loDb.disconnect()
		This.Say("A1", IIF(lcBack == lcTxt AND RTRIM(TRANSFORM(lvRaw)) == lcTxt, "OK", "BUG"), "Save() = " + TRANSFORM(lvSave) + "; sent hex " + STRCONV(lcTxt, 15) + ;
			"; another connection reads hex " + STRCONV(RTRIM(TRANSFORM(lvRaw)), 15) + IIF(TRANSFORM(lvRaw) == RTRIM(TRANSFORM(lvRaw)), "", " (+ " + TRANSFORM(LEN(TRANSFORM(lvRaw)) - LEN(RTRIM(TRANSFORM(lvRaw)))) + " trailing blanks)") + "; FoxRemote's use() reads hex " + STRCONV(lcBack, 15))
	ENDPROC

	* --- F3: Connect() creates a database when it does not exist -------------
	PROCEDURE F3
		LOCAL lcMissing, loDb, lvOk
		lcMissing = This.oT.MissingDatabase()
		loDb = This.oT.NewRemote(lcMissing)
		lvOk = loDb.Connect()
		loDb.disconnect()
		IF This.oT.DatabaseExists(lcMissing)
			This.Say("F3", "BUG", "Connect() to a missing database returned " + TRANSFORM(lvOk) + " and the database now exists: " + lcMissing)
		ELSE
			This.Say("F3", "OK", "Connect() returned " + TRANSFORM(lvOk) + "; the database was not created")
		ENDIF
		This.oT.DropDatabase(lcMissing)
	ENDPROC

	* --- F11: global SET changes on Connect -----------------------------------
	PROCEDURE F11
		LOCAL loDb, lcBefore
		SET MULTILOCKS OFF
		lcBefore = SET("MULTILOCKS")
		loDb = This.oT.NewRemote()
		loDb.Connect()
		This.Say("F11", IIF(SET("MULTILOCKS") == lcBefore, "OK", "BUG"), "SET MULTILOCKS before Connect: " + lcBefore + ", after: " + SET("MULTILOCKS"))
		loDb.disconnect()
	ENDPROC

	* --- F1a: Save() of an alias it does not know leaves Transactions = 2 ------
	PROCEDURE F1a
		LOCAL loDb, lnH, lnTxBefore, lnTxAfter, lvSave, lnLost
		loDb = This.oT.NewRemote()
		loDb.Connect()
		lnH = This.LastHandle()
		lnTxBefore = SQLGETPROP(lnH, "Transactions")
		lvSave = loDb.Save("not_opened")
		lnTxAfter = SQLGETPROP(lnH, "Transactions")
		loDb.SQLExec("INSERT INTO customers (name, city) VALUES ('lost_f1a', 'X')")
		loDb.disconnect()
		lnLost = This.oT.RawScalar("SELECT COUNT(*) FROM customers WHERE RTRIM(name) = 'lost_f1a'")
		This.Say("F1a", IIF(lnTxAfter # lnTxBefore OR This.Num(lnLost) = 0, "BUG", "OK"), ;
			"Save('not_opened') = " + TRANSFORM(lvSave) + "; Transactions " + TRANSFORM(lnTxBefore) + " -> " + TRANSFORM(lnTxAfter) + ;
			"; an INSERT run afterwards and then disconnect: rows on the server = " + TRANSFORM(lnLost))
	ENDPROC

	* --- F1b: Save() of a read-only cursor leaves SET DATE changed -------------
	PROCEDURE F1b
		LOCAL loDb, lnH, lnTxBefore, lnTxAfter, lvSave, lcDateBefore
		SET DATE AMERICAN
		lcDateBefore = SET("DATE")
		loDb = This.oT.NewRemote()
		loDb.Connect()
		lnH = This.LastHandle()
		lnTxBefore = SQLGETPROP(lnH, "Transactions")
		loDb.use("customers AS ro_f1b", "", "", "", .T.)
		lvSave = loDb.Save("ro_f1b")
		lnTxAfter = SQLGETPROP(lnH, "Transactions")
		This.Say("F1b", IIF(SET("DATE") # lcDateBefore OR lnTxAfter # lnTxBefore, "BUG", "OK"), ;
			"Save(read-only) = " + TRANSFORM(lvSave) + "; SET DATE " + lcDateBefore + " -> " + SET("DATE") + ;
			"; Transactions " + TRANSFORM(lnTxBefore) + " -> " + TRANSFORM(lnTxAfter))
		SET DATE AMERICAN
		loDb.closeAll()
		loDb.disconnect()
	ENDPROC

	* --- F1c: use() + APPEND + Save() with CursorAdapter, seen from another connection
	PROCEDURE F1c
		LOCAL loDb, lvSave, lnSeenBefore, lnSeenAfter, lnH, lnTx, lnSeenCommit
		loDb = This.oT.NewRemote()
		loDb.Connect()
		lnH = This.LastHandle()
		loDb.use("customers AS ca_f1c")
		SELECT ca_f1c
		APPEND BLANK
		REPLACE name WITH "f1c_row", city WITH "Rome"
		lvSave = loDb.Save("ca_f1c")
		lnSeenBefore = This.oT.RawScalar("SELECT COUNT(*) FROM customers WHERE RTRIM(name) = 'f1c_row'")
		lnTx = SQLGETPROP(lnH, "Transactions")
		=SQLCOMMIT(lnH)
		lnSeenCommit = This.oT.RawScalar("SELECT COUNT(*) FROM customers WHERE RTRIM(name) = 'f1c_row'")
		loDb.closeAll()
		loDb.disconnect()
		lnSeenAfter = This.oT.RawScalar("SELECT COUNT(*) FROM customers WHERE RTRIM(name) = 'f1c_row'")
		This.Say("F1c", IIF(This.Num(lnSeenBefore) = 1, "OK", "BUG"), ;
			"Save() = " + TRANSFORM(lvSave) + "; another connection sees the row before disconnect: " + TRANSFORM(lnSeenBefore) + ;
			"; Transactions after Save() = " + TRANSFORM(lnTx) + "; after a SQLCOMMIT() by hand: " + TRANSFORM(lnSeenCommit) + ", after disconnect: " + TRANSFORM(lnSeenAfter) + IIF(lvSave, "", "; cLastError: " + This.Short(loDb.cLastError)))
	ENDPROC

	* --- F12: Save() with bUseCA = .F. updating two columns ----------------------
	PROCEDURE F12
		LOCAL loDb, lvSave, lvRow
		This.oT.RawRun("INSERT INTO customers (name, city) VALUES ('f12_old', 'Oslo')")
		loDb = This.oT.NewRemote()
		loDb.bUseCA = .F.
		loDb.Connect()
		loDb.use("customers AS sx_f12", "", "name = 'f12_old'")
		SELECT sx_f12
		REPLACE name WITH "f12_new", city WITH "Lima"
		lvSave = loDb.Save("sx_f12")
		=SQLCOMMIT(This.LastHandle())
		loDb.closeAll()
		loDb.disconnect()
		lvRow = This.oT.RawScalar("SELECT " + This.oT.Concat("name", "'/'", "city") + " FROM customers WHERE RTRIM(name) IN ('f12_old', 'f12_new')")
		This.Say("F12", IIF(ALLTRIM(TRANSFORM(lvRow)) == "f12_new/Lima", "OK", "BUG"), ;
			"bUseCA=.F., name and city changed, Save() = " + TRANSFORM(lvSave) + "; on the server: " + TRANSFORM(lvRow) + " (expected f12_new/Lima), read after a SQLCOMMIT() by hand" + ;
			IIF(lvSave, "", "; cLastError: " + This.Short(loDb.cLastError)))
	ENDPROC

	* --- F2: saveGroup() with bUseCA = .F. ----------------------------------------
	PROCEDURE F2
		LOCAL loDb, lvSave, lvCity
		This.oT.RawRun("INSERT INTO customers (name, city) VALUES ('f2_row', 'Bonn')")
		loDb = This.oT.NewRemote()
		loDb.bUseCA = .F.
		loDb.Connect()
		loDb.use("customers AS sx_f2", "", "name = 'f2_row'", "grp_f2")
		SELECT sx_f2
		REPLACE city WITH "Kiel"
		lvSave = loDb.saveGroup("grp_f2")
		loDb.closeAll()
		loDb.disconnect()
		lvCity = This.oT.RawScalar("SELECT RTRIM(city) FROM customers WHERE RTRIM(name) = 'f2_row'")
		This.Say("F2", IIF(ALLTRIM(TRANSFORM(lvCity)) == "Kiel", "OK", "BUG"), ;
			"saveGroup() = " + TRANSFORM(lvSave) + "; city on the server: " + TRANSFORM(lvCity) + " (expected Kiel)")
	ENDPROC

	* --- F5: getTables() / getTableFields() -----------------------------------------
	PROCEDURE F5
		LOCAL loDb, lvT, lvF, lcT, lcF
		loDb = This.oT.NewRemote()
		loDb.Connect()
		TRY
			lvT = loDb.getTables()
			lcT = "TYPE=" + TYPE("lvT") + " value=" + TRANSFORM(lvT) + " ALEN(aCustomArray)=" + TRANSFORM(ALEN(loDb.aCustomArray))
		CATCH TO loEx
			lcT = "raised " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message
		ENDTRY
		TRY
			lvF = loDb.getTableFields("customers")
			lcF = "TYPE=" + TYPE("lvF") + " value=" + TRANSFORM(lvF) + " ALEN(aCustomArray)=" + TRANSFORM(ALEN(loDb.aCustomArray))
		CATCH TO loEx
			lcF = "raised " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message
		ENDTRY
		This.Say("F5", "INFO", "getTables(): " + lcT + " || getTableFields('customers') (4 columns): " + lcF)
		loDb.disconnect()
	ENDPROC

	* --- F6: getLastID() --------------------------------------------------------------
	PROCEDURE F6
		LOCAL loDb, lvId, lnReal
		loDb = This.oT.NewRemote()
		loDb.Connect()
		loDb.SQLExec("INSERT INTO customers (name) VALUES ('f6_row')")
		lnReal = This.oT.RawScalar("SELECT MAX(id) FROM customers")
		TRY
			lvId = loDb.getLastID()
			This.Say("F6", IIF(VARTYPE(lvId) == "N" AND lvId = This.Num(lnReal), "OK", "BUG"), "getLastID() = " + TRANSFORM(lvId) + " (type " + VARTYPE(lvId) + ")" + "; the id inserted is " + TRANSFORM(lnReal))
		CATCH TO loEx
			This.Say("F6", "BUG", "getLastID() raised " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message + "; the id inserted is " + TRANSFORM(lnReal))
		ENDTRY
		loDb.disconnect()
	ENDPROC

	* --- N1: primary key detection ---------------------------------------------------------
	PROCEDURE N1
		LOCAL loDb, lcPk
		loDb = This.oT.NewRemote()
		loDb.Connect()
		lcPk = TRANSFORM(loDb.getPrimaryKey("customers"))
		This.Say("N1", IIF(LOWER(ALLTRIM(lcPk)) == "id", "OK", "BUG"), "getPrimaryKey('customers') = '" + lcPk + "' (expected id); script: " + This.Short(loDb.getPrimaryKeyScript("customers")))
		loDb.disconnect()
	ENDPROC

	* --- F6b: getLastID() after a parameterised INSERT (what Save() sends) -------------------
	PROCEDURE F6b
		LOCAL loDb, lvId, lnReal
		* FoxRemote evaluates ?param inside its own method: a LOCAL is invisible there
		* (VFP then asks for it in a modal dialog), so the parameter has to be PRIVATE.
		PRIVATE pcName
		loDb = This.oT.NewRemote()
		loDb.Connect()
		pcName = "f6b_row"
		loDb.SQLExec("INSERT INTO customers (name) VALUES (?pcName)")
		lnReal = This.oT.RawScalar("SELECT MAX(id) FROM customers")
		TRY
			lvId = loDb.getLastID()
			This.Say("F6b", IIF(VARTYPE(lvId) == "N" AND lvId = This.Num(lnReal), "OK", "BUG"), "getLastID() after INSERT ... VALUES (?param) = " + TRANSFORM(lvId) + " (type " + VARTYPE(lvId) + ")" + "; the id inserted is " + TRANSFORM(lnReal))
		CATCH TO loEx
			This.Say("F6b", "BUG", "getLastID() raised " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message + "; the id inserted is " + TRANSFORM(lnReal))
		ENDTRY
		loDb.disconnect()
	ENDPROC

	* --- F7: closeGroup() and reopen the same alias -------------------------------------
	PROCEDURE F7
		LOCAL loDb, lvClose, lvReopen, lcErr
		loDb = This.oT.NewRemote()
		loDb.Connect()
		loDb.use("customers AS cg_f7", "", "", "grp_f7")
		lvClose = loDb.closeGroup("grp_f7")
		lcErr = ""
		TRY
			lvReopen = loDb.use("customers AS cg_f7", "", "", "grp_f7")
		CATCH TO loEx
			lvReopen = .F.
			lcErr = " raised " + TRANSFORM(loEx.ErrorNo) + " " + loEx.Message
		ENDTRY
		This.Say("F7", IIF(lvReopen = .T., "OK", "BUG"), "closeGroup() = " + TRANSFORM(lvClose) + ", alias open after: " + TRANSFORM(USED("cg_f7")) + ;
			"; use() again = " + TRANSFORM(lvReopen) + lcErr)
		loDb.closeAll()
		loDb.disconnect()
	ENDPROC

	* --- F8a: requery() of a read-only cursor ---------------------------------------------
	PROCEDURE F8a
		LOCAL loDb, lnBefore, lnAfter
		loDb = This.oT.NewRemote()
		loDb.Connect()
		loDb.use("customers AS ro_f8", "", "", "", .T.)
		SELECT ro_f8
		COUNT TO lnBefore
		This.oT.RawRun("INSERT INTO customers (name) VALUES ('f8a_new')")
		loDb.requery("ro_f8")
		SELECT ro_f8
		COUNT TO lnAfter
		This.Say("F8a", IIF(lnAfter = lnBefore + 1, "OK", "BUG"), "read-only cursor, a row inserted by another connection, requery(): rows " + TRANSFORM(lnBefore) + " -> " + TRANSFORM(lnAfter))
		loDb.closeAll()
		loDb.disconnect()
	ENDPROC

	* --- F8b: requery() with a pending change (README: keeps it) ----------------------------
	PROCEDURE F8b
		LOCAL loDb, lcCity, lnRows
		This.oT.RawRun("INSERT INTO customers (name, city) VALUES ('f8b_row', 'Graz')")
		loDb = This.oT.NewRemote()
		loDb.Connect()
		loDb.use("customers AS rw_f8", "", "name = 'f8b_row'")
		SELECT rw_f8
		REPLACE city WITH "Linz"
		This.oT.RawRun("UPDATE customers SET name = 'f8b_row' WHERE RTRIM(name) = 'Ana'")
		loDb.requery("rw_f8")
		SELECT rw_f8
		COUNT TO lnRows
		LOCATE FOR city = "Linz"
		lcCity = IIF(FOUND(), "Linz kept", "Linz lost")
		This.Say("F8b", IIF(FOUND() AND lnRows = 2, "OK", "BUG"), "pending change city=Linz, another connection makes a 2nd row match, requery(): " + lcCity + ", rows in the cursor " + TRANSFORM(lnRows) + " (a real requery that keeps the change gives 2 and Linz kept)")
		loDb.closeAll()
		loDb.disconnect()
	ENDPROC

	* --- F9: use() + Save() on a table without a primary key -----------------------------------
	PROCEDURE F9
		LOCAL loDb, lvUse, lvSave, lnRows
		loDb = This.oT.NewRemote()
		loDb.Connect()
		lvUse = loDb.use("nopk AS np_f9")
		IF lvUse
			SELECT np_f9
			APPEND BLANK
			REPLACE code WITH "X1", descr WITH "f9"
			lvSave = loDb.Save("np_f9")
			=SQLCOMMIT(This.LastHandle())
		ENDIF
		loDb.closeAll()
		loDb.disconnect()
		lnRows = This.oT.RawScalar("SELECT COUNT(*) FROM nopk WHERE RTRIM(code) = 'X1'")
		This.Say("F9", IIF(This.Num(lnRows) = 1, "OK", "BUG"), "use(nopk) = " + TRANSFORM(lvUse) + ", Save() = " + TRANSFORM(lvSave) + ", rows on the server after a SQLCOMMIT() by hand: " + TRANSFORM(lnRows) + ;
			"; cLastError: " + This.Short(loDb.cLastError))
	ENDPROC

	* --- F10: a ';' inside a connection value ---------------------------------------------------
	PROCEDURE F10
		LOCAL loDb, lvOk, lcValue, lcWhere, loBad, lvBad
		lcValue = This.oT.SemicolonValue()
		IF EMPTY(lcValue)
			This.Say("F10", "INFO", "not applicable to this engine")
			RETURN
		ENDIF
		loDb = This.oT.NewRemoteWithSemicolon(lcValue)
		lvOk = loDb.Connect()
		loDb.disconnect()
		lcWhere = This.oT.SemicolonWhere(lcValue)
		loBad = This.oT.NewRemoteWithSemicolon(lcValue + "x")
		lvBad = loBad.Connect()
		loBad.disconnect()
		This.Say("F10", IIF(lvOk = .T. AND EMPTY(lcWhere), "OK", "BUG"), This.oT.SemicolonWhat() + " '" + This.oT.SemicolonShow(lcValue) + "': Connect() = " + TRANSFORM(lvOk) + "; " + IIF(EMPTY(lcWhere), "the value arrived whole", lcWhere) + "; sanity, the same with a wrong value: Connect() = " + TRANSFORM(lvBad) + "; cLastError: " + This.Short(loDb.cLastError))
	ENDPROC

	* --- F4a: migrate() of a DBF whose table already exists on the server ------------------------
	PROCEDURE F4a
		LOCAL loDb, lcDbf, lvMig, lnKeep, lvBorn, lnRows, lnPre
		lcDbf = This.oT.cDir + "mig_cust.dbf"
		CREATE TABLE (lcDbf) FREE (name C(20), born D)
		INSERT INTO mig_cust VALUES ("Luis", {^2021-03-04})
		INSERT INTO mig_cust VALUES ("Empty", {})
		USE IN mig_cust
		This.oT.RawRun("CREATE TABLE mig_cust (name " + This.oT.TextType(20) + ", keep " + This.oT.TextType(10) + ")")
		This.oT.RawRun("INSERT INTO mig_cust (name, keep) VALUES ('keepme', 'yes')")
		lnPre = This.oT.RawScalar("SELECT COUNT(*) FROM mig_cust")
		loDb = This.oT.NewRemote()
		loDb.Connect()
		lvMig = loDb.migrate(lcDbf)
		loDb.disconnect()
		lnKeep = This.oT.RawScalar("SELECT COUNT(*) FROM mig_cust WHERE RTRIM(name) = 'keepme'")
		lnRows = This.oT.RawScalar("SELECT COUNT(*) FROM mig_cust")
		lvBorn = This.oT.RawScalar("SELECT " + This.oT.DateAsText("born") + " FROM mig_cust WHERE RTRIM(name) = 'Empty'")
		This.Say("F4a", IIF(This.Num(lnKeep) = 1, "OK", "BUG"), "before migrate() the server table had " + TRANSFORM(lnPre) + " row(s); migrate() = " + TRANSFORM(lvMig) + "; the row that was on the server before: " + TRANSFORM(lnKeep) + ;
			" (the table now has " + TRANSFORM(lnRows) + " rows); an empty DBF date arrived as: " + TRANSFORM(lvBorn) + ;
			"; cLastError: " + This.Short(loDb.cLastError))
	ENDPROC

	* --- F4b: migrate() of a folder changes SET DEFAULT ------------------------------------------
	PROCEDURE F4b
		LOCAL loDb, lcSub, lcBefore, lvMig
		lcSub = This.oT.cDir + "migdir_" + SYS(2015)
		MKDIR (lcSub)
		CREATE TABLE (ADDBS(lcSub) + "mig_dir1") FREE (code C(5))
		INSERT INTO mig_dir1 VALUES ("A")
		USE IN mig_dir1
		lcBefore = SYS(5) + SYS(2003)
		loDb = This.oT.NewRemote()
		loDb.Connect()
		lvMig = loDb.migrate(lcSub)
		loDb.disconnect()
		This.Say("F4b", IIF(SYS(5) + SYS(2003) == lcBefore, "OK", "BUG"), "migrate(folder) = " + TRANSFORM(lvMig) + "; default folder before: " + lcBefore + ", after: " + SYS(5) + SYS(2003))
		CD (lcBefore)
	ENDPROC
ENDDEFINE

* ======================================================================== *
* Target: what changes from one engine to another.
* ======================================================================== *
DEFINE CLASS Target AS Custom
	cName = ""
	cDir = ""
	cLastError = ""

	FUNCTION ConnStr
	ENDFUNC

	* Engine-specific checks that do not fit the common probes.
	PROCEDURE ExtraChecks(toProbe)
	ENDPROC

	* Lets an engine adapt the probe's SQL (Firebird has no RTRIM).
	FUNCTION Adapt(tcSql)
		RETURN tcSql
	ENDFUNC

	* Runs a statement on a fresh connection of its own.
	FUNCTION RawRun(tcSql)
		LOCAL lnH, lnR, laErr[1]
		lnH = SQLSTRINGCONNECT(This.ConnStr(), .T.)
		IF lnH < 1
			=AERROR(laErr)
			This.cLastError = "connect: " + TRANSFORM(laErr[2])
			RETURN .F.
		ENDIF
		=SQLSETPROP(lnH, "QueryTimeOut", 10)
		lnR = SQLEXEC(lnH, This.Adapt(tcSql))
		IF lnR < 1
			=AERROR(laErr)
			This.cLastError = tcSql + ": " + TRANSFORM(laErr[2])
		ENDIF
		=SQLDISCONNECT(lnH)
		RETURN lnR > 0
	ENDFUNC

	* First column of the first row, read on a fresh connection; .NULL. on failure.
	FUNCTION RawScalar(tcSql)
		LOCAL lnH, lvValue, laErr[1]
		lvValue = .NULL.
		lnH = SQLSTRINGCONNECT(This.ConnStr(), .T.)
		IF lnH < 1
			RETURN lvValue
		ENDIF
		=SQLSETPROP(lnH, "QueryTimeOut", 10)
		IF SQLEXEC(lnH, This.Adapt(tcSql), "c_rawscalar") > 0 AND USED("c_rawscalar")
			SELECT c_rawscalar
			IF !EOF()
				lvValue = EVALUATE(FIELD(1))
				IF VARTYPE(lvValue) == "C"
					lvValue = STRTRAN(lvValue, CHR(0), "")
				ENDIF
			ENDIF
			USE IN c_rawscalar
		ELSE
			=AERROR(laErr)
			lvValue = "<" + TRANSFORM(laErr[2]) + ">"
		ENDIF
		=SQLDISCONNECT(lnH)
		RETURN lvValue
	ENDFUNC
ENDDEFINE

DEFINE CLASS SqliteTarget AS Target
	cName = "SQLite"
	cDriver = "SQLite3 ODBC Driver"
	cDbFile = ""

	FUNCTION Setup(tcDir)
		This.cDir = ADDBS(tcDir)
		This.cDbFile = This.cDir + "fr_" + SYS(2015) + ".db"
		RETURN This.RawRun("CREATE TABLE customers (id INTEGER PRIMARY KEY AUTOINCREMENT, name VARCHAR(40), city VARCHAR(40), born DATE)") ;
			AND This.RawRun("CREATE TABLE nopk (code VARCHAR(10), descr VARCHAR(40))") ;
			AND This.RawRun("INSERT INTO customers (name, city, born) VALUES ('Ana', 'Madrid', '2020-01-02')")
	ENDFUNC

	FUNCTION ConnStr
		RETURN "DRIVER={" + This.cDriver + "};DATABASE=" + This.cDbFile + ";"
	ENDFUNC

	FUNCTION ServerVersion
		RETURN TRANSFORM(This.RawScalar("SELECT sqlite_version()"))
	ENDFUNC

	FUNCTION NewRemote(tcDatabase)
		LOCAL loDb
		loDb = CREATEOBJECT("SQLite")
		loDb.bShowErrors = .F.
		loDb.cDriver = This.cDriver
		loDb.cDatabase = IIF(EMPTY(tcDatabase), This.cDbFile, tcDatabase)
		RETURN loDb
	ENDFUNC

	FUNCTION MissingDatabase
		RETURN This.cDir + "missing_" + SYS(2015) + ".db"
	ENDFUNC

	FUNCTION DatabaseExists(tcDatabase)
		RETURN FILE(tcDatabase)
	ENDFUNC

	PROCEDURE DropDatabase(tcDatabase)
		IF FILE(tcDatabase)
			ERASE (tcDatabase)
		ENDIF
	ENDPROC

	* SQLite has no password: the ';' goes in the path of the database file.
	FUNCTION SemicolonValue
		LOCAL lcDir
		lcDir = This.cDir + "semi;colon_" + SYS(2015)
		MKDIR (lcDir)
		RETURN ADDBS(lcDir) + "fr.db"
	ENDFUNC

	FUNCTION SemicolonWhat
		RETURN "database path with ';'"
	ENDFUNC

	FUNCTION SemicolonShow(tcValue)
		RETURN tcValue
	ENDFUNC

	* Empty when the database went where it was asked.
	FUNCTION SemicolonWhere(tcValue)
		LOCAL lcCut
		IF FILE(tcValue)
			RETURN ""
		ENDIF
		lcCut = LEFT(tcValue, AT(";", tcValue) - 1)
		RETURN "the file asked for does not exist" + IIF(FILE(lcCut), "; the driver created " + lcCut + " instead (the path was cut at the ';')", "")
	ENDFUNC

	FUNCTION NewRemoteWithSemicolon(tcValue)
		RETURN This.NewRemote(tcValue)
	ENDFUNC

	FUNCTION TextType(tnLen)
		RETURN "VARCHAR(" + TRANSFORM(tnLen) + ")"
	ENDFUNC

	FUNCTION Concat(tc1, tc2, tc3)
		RETURN "RTRIM(" + tc1 + ") || " + tc2 + " || RTRIM(" + tc3 + ")"
	ENDFUNC

	FUNCTION DateAsText(tcField)
		RETURN "CAST(" + tcField + " AS TEXT)"
	ENDFUNC

	PROCEDURE Teardown
		This.DropDatabase(This.cDbFile)
	ENDPROC
ENDDEFINE

* ======================================================================== *
* SQL Server. Credentials come from FoxRemote\probe\engines.local.ini (not versioned).
* ======================================================================== *
#DEFINE PROBE_INI "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\probe\engines.local.ini"

DEFINE CLASS MssqlTarget AS Target
	cName = "SQL Server"
	cDriver = ""
	cServer = ""
	cSaUser = ""
	cSaPassword = ""
	cDatabase = ""

	FUNCTION Setup(tcDir)
		LOCAL lcIni
		This.cDir = ADDBS(tcDir)
		lcIni = FILETOSTR(PROBE_INI)
		This.cDriver = This.IniValue(lcIni, "driver")
		This.cServer = This.IniValue(lcIni, "server")
		This.cSaUser = This.IniValue(lcIni, "user")
		This.cSaPassword = This.IniValue(lcIni, "password")
		This.cDatabase = "fr_probe" + SYS(2015)
		RETURN This.MasterRun("CREATE DATABASE [" + This.cDatabase + "]") ;
			AND This.RawRun("CREATE TABLE customers (id INT IDENTITY(1,1) PRIMARY KEY, name VARCHAR(40), city VARCHAR(40), born DATE)") ;
			AND This.RawRun("CREATE TABLE nopk (code VARCHAR(10), descr VARCHAR(40))") ;
			AND This.RawRun("INSERT INTO customers (name, city, born) VALUES ('Ana', 'Madrid', '2020-01-02')")
	ENDFUNC

	* Tolerates LF and CRLF line ends (a section once came back LF-only and the value was cut).
	FUNCTION IniValue(tcIni, tcKey)
		RETURN ALLTRIM(STREXTRACT(CHR(10) + STRTRAN(tcIni, CHR(13), "") + CHR(10), CHR(10) + tcKey + "=", CHR(10)))
	ENDFUNC

	FUNCTION ConnStr(tcDatabase)
		RETURN "DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";Trusted_Connection=yes;DATABASE=" + IIF(EMPTY(tcDatabase), This.cDatabase, tcDatabase) + ";"
	ENDFUNC

	FUNCTION MasterRun(tcSql)
		LOCAL lcDb, lvOk
		lcDb = This.cDatabase
		This.cDatabase = "master"
		lvOk = This.RawRun(tcSql)
		This.cDatabase = lcDb
		RETURN lvOk
	ENDFUNC

	FUNCTION MasterScalar(tcSql)
		LOCAL lcDb, lvValue
		lcDb = This.cDatabase
		This.cDatabase = "master"
		lvValue = This.RawScalar(tcSql)
		This.cDatabase = lcDb
		RETURN lvValue
	ENDFUNC

	FUNCTION ServerVersion
		RETURN TRANSFORM(This.RawScalar("SELECT CAST(SERVERPROPERTY('ProductVersion') AS varchar(20)) + ' ' + CAST(SERVERPROPERTY('Edition') AS varchar(40))"))
	ENDFUNC

	* FoxRemote's MSSQL always sends UID/PWD (fault F10): the probe connects with
	* Windows authentication through a subclass that only changes the connection string.
	FUNCTION NewRemote(tcDatabase)
		LOCAL loDb
		loDb = CREATEOBJECT("TrustedMSSQL")
		loDb.bShowErrors = .F.
		loDb.cDriver = This.cDriver
		loDb.cServer = This.cServer
		loDb.cDatabase = IIF(EMPTY(tcDatabase), This.cDatabase, tcDatabase)
		RETURN loDb
	ENDFUNC

	FUNCTION MissingDatabase
		RETURN "fr_missing" + SYS(2015)
	ENDFUNC

	FUNCTION DatabaseExists(tcDatabase)
		RETURN This.Num(This.MasterScalar("SELECT COUNT(*) FROM sys.databases WHERE name = '" + tcDatabase + "'")) > 0
	ENDFUNC

	* BIGINT (COUNT(*) in MySQL/MariaDB) reaches VFP as text: compare by value.
	FUNCTION Num(tvValue)
		DO CASE
		CASE VARTYPE(tvValue) == "N"
			RETURN tvValue
		CASE VARTYPE(tvValue) == "C" AND !EMPTY(tvValue) AND EMPTY(CHRTRAN(ALLTRIM(tvValue), "0123456789-", ""))
			RETURN VAL(tvValue)
		ENDCASE
		RETURN -1
	ENDFUNC

	PROCEDURE DropDatabase(tcDatabase)
		IF This.DatabaseExists(tcDatabase)
			This.MasterRun("ALTER DATABASE [" + tcDatabase + "] SET SINGLE_USER WITH ROLLBACK IMMEDIATE")
			This.MasterRun("DROP DATABASE [" + tcDatabase + "]")
		ENDIF
	ENDPROC

	FUNCTION SemicolonValue
		RETURN This.cSaPassword
	ENDFUNC

	FUNCTION SemicolonWhat
		RETURN "SQL login with ';' in the password"
	ENDFUNC

	FUNCTION SemicolonShow(tcValue)
		RETURN "<hidden, " + TRANSFORM(LEN(tcValue)) + " chars, has ';'=" + TRANSFORM(";" $ tcValue) + ">"
	ENDFUNC

	FUNCTION SemicolonWhere(tcValue)
		RETURN ""
	ENDFUNC

	* The plain MSSQL class, as a customer would use it.
	FUNCTION NewRemoteWithSemicolon(tcValue)
		LOCAL loDb
		loDb = CREATEOBJECT("MSSQL")
		loDb.bShowErrors = .F.
		loDb.cDriver = This.cDriver
		loDb.cServer = This.cServer
		loDb.cUser = This.cSaUser
		loDb.cPassword = tcValue
		loDb.cDatabase = This.cDatabase
		RETURN loDb
	ENDFUNC

	FUNCTION TextType(tnLen)
		RETURN "VARCHAR(" + TRANSFORM(tnLen) + ")"
	ENDFUNC

	FUNCTION Concat(tc1, tc2, tc3)
		RETURN "RTRIM(" + tc1 + ") + " + tc2 + " + RTRIM(" + tc3 + ")"
	ENDFUNC

	FUNCTION DateAsText(tcField)
		RETURN "CONVERT(varchar(30), " + tcField + ", 121)"
	ENDFUNC

	PROCEDURE Teardown
		This.DropDatabase(This.cDatabase)
	ENDPROC
ENDDEFINE

DEFINE CLASS TrustedMSSQL AS MSSQL
	FUNCTION getConnectionString(tbAddDatabase)
		LOCAL lcConStr
		lcConStr = "DRIVER=" + This.cDriver + ";SERVER=" + This.cServer + ";Trusted_Connection=yes"
		IF tbAddDatabase
			lcConStr = lcConStr + ";DATABASE=" + ALLTRIM(This.cDatabase)
		ENDIF
		RETURN lcConStr + ";APP=FoxRemote probe"
	ENDFUNC
ENDDEFINE

* ======================================================================== *
* MariaDB. root's password (with ';') comes from engines.local.ini [mariadb].
* ======================================================================== *
DEFINE CLASS MariadbTarget AS Target
	cName = "MariaDB"
	cDriver = ""
	cServer = ""
	nPort = 0
	cUser = ""
	cPassword = ""
	cDatabase = ""

	FUNCTION Setup(tcDir)
		LOCAL lcIni
		This.cDir = ADDBS(tcDir)
		lcIni = FILETOSTR(PROBE_INI)
		lcIni = SUBSTR(lcIni, AT("[mariadb]", lcIni))
		This.cDriver = This.IniValue(lcIni, "driver")
		This.cServer = This.IniValue(lcIni, "server")
		This.nPort = VAL(This.IniValue(lcIni, "port"))
		This.cUser = This.IniValue(lcIni, "user")
		This.cPassword = This.IniValue(lcIni, "password")
		This.cDatabase = LOWER("fr_probe" + SYS(2015))
		RETURN This.ServerRun("CREATE DATABASE `" + This.cDatabase + "` DEFAULT CHARACTER SET utf8mb4") ;
			AND This.RawRun("CREATE TABLE customers (id INT AUTO_INCREMENT PRIMARY KEY, name VARCHAR(40), city VARCHAR(40), born DATE) ENGINE=InnoDB") ;
			AND This.RawRun("CREATE TABLE nopk (code VARCHAR(10), descr VARCHAR(40)) ENGINE=InnoDB") ;
			AND This.RawRun("INSERT INTO customers (name, city, born) VALUES ('Ana', 'Madrid', '2020-01-02')")
	ENDFUNC

	* Tolerates LF and CRLF line ends (a section once came back LF-only and the value was cut).
	FUNCTION IniValue(tcIni, tcKey)
		RETURN ALLTRIM(STREXTRACT(CHR(10) + STRTRAN(tcIni, CHR(13), "") + CHR(10), CHR(10) + tcKey + "=", CHR(10)))
	ENDFUNC

	FUNCTION ConnStr
		RETURN "DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";PORT=" + TRANSFORM(This.nPort) + ";UID=" + This.cUser + ;
			";PWD={" + This.cPassword + "}" + IIF(EMPTY(This.cDatabase), "", ";DATABASE=" + This.cDatabase) + ";"
	ENDFUNC

	FUNCTION ServerRun(tcSql)
		LOCAL lcDb, lvOk
		lcDb = This.cDatabase
		This.cDatabase = ""
		lvOk = This.RawRun(tcSql)
		This.cDatabase = lcDb
		RETURN lvOk
	ENDFUNC

	FUNCTION ServerScalar(tcSql)
		LOCAL lcDb, lvValue
		lcDb = This.cDatabase
		This.cDatabase = ""
		lvValue = This.RawScalar(tcSql)
		This.cDatabase = lcDb
		RETURN lvValue
	ENDFUNC

	FUNCTION ServerVersion
		RETURN TRANSFORM(This.RawScalar("SELECT VERSION()")) + " / charset " + TRANSFORM(This.RawScalar("SELECT @@character_set_connection"))
	ENDFUNC

	* FoxRemote's MySQL class writes PASSWORD= bare (fault F10) and root's password has
	* a ';': the probe connects through a subclass that only wraps it in {}.
	FUNCTION NewRemote(tcDatabase)
		LOCAL loDb
		loDb = CREATEOBJECT("BracedMariaDB")
		This.Configure(loDb, IIF(EMPTY(tcDatabase), This.cDatabase, tcDatabase))
		RETURN loDb
	ENDFUNC

	PROCEDURE Configure(toDb, tcDatabase)
		toDb.bShowErrors = .F.
		toDb.cDriver = This.cDriver
		toDb.cServer = This.cServer
		toDb.nPort = This.nPort
		toDb.cUser = This.cUser
		toDb.cPassword = This.cPassword
		toDb.cDatabase = tcDatabase
	ENDPROC

	FUNCTION MissingDatabase
		RETURN LOWER("fr_missing" + SYS(2015))
	ENDFUNC

	FUNCTION DatabaseExists(tcDatabase)
		RETURN VAL(TRANSFORM(This.ServerScalar("SELECT COUNT(*) FROM information_schema.schemata WHERE schema_name = '" + tcDatabase + "'"))) > 0
	ENDFUNC

	PROCEDURE DropDatabase(tcDatabase)
		This.ServerRun("DROP DATABASE IF EXISTS `" + tcDatabase + "`")
	ENDPROC

	FUNCTION SemicolonValue
		RETURN This.cPassword
	ENDFUNC

	FUNCTION SemicolonWhat
		RETURN "user with ';' in the password"
	ENDFUNC

	FUNCTION SemicolonShow(tcValue)
		RETURN "<hidden, " + TRANSFORM(LEN(tcValue)) + " chars, has ';'=" + TRANSFORM(";" $ tcValue) + ">"
	ENDFUNC

	FUNCTION SemicolonWhere(tcValue)
		RETURN ""
	ENDFUNC

	* The plain MariaDB class, as a customer would use it.
	FUNCTION NewRemoteWithSemicolon(tcValue)
		LOCAL loDb
		loDb = CREATEOBJECT("MariaDB")
		This.Configure(loDb, This.cDatabase)
		loDb.cPassword = tcValue
		RETURN loDb
	ENDFUNC

	FUNCTION TextType(tnLen)
		RETURN "VARCHAR(" + TRANSFORM(tnLen) + ")"
	ENDFUNC

	FUNCTION Concat(tc1, tc2, tc3)
		RETURN "CONCAT(RTRIM(" + tc1 + "), " + tc2 + ", RTRIM(" + tc3 + "))"
	ENDFUNC

	FUNCTION DateAsText(tcField)
		RETURN "CAST(" + tcField + " AS CHAR)"
	ENDFUNC

	PROCEDURE Teardown
		This.DropDatabase(This.cDatabase)
	ENDPROC
ENDDEFINE

DEFINE CLASS BracedMariaDB AS MariaDB
	FUNCTION getConnectionString(tbAddDatabase)
		LOCAL lcConStr
		lcConStr = "DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";USER=" + This.cUser + ";PASSWORD={" + This.cPassword + "}"
		IF This.nPort > 0
			lcConStr = lcConStr + ";PORT=" + ALLTRIM(STR(This.nPort))
		ENDIF
		IF tbAddDatabase
			lcConStr = lcConStr + ";DATABASE=" + ALLTRIM(This.cDatabase)
		ENDIF
		RETURN lcConStr + ";COLLATION=utf8_general_ci;APP=FoxRemote probe"
	ENDFUNC
ENDDEFINE

* ======================================================================== *
* PostgreSQL. User fr_probe (password with ';'), from engines.local.ini [postgresql].
* ======================================================================== *
DEFINE CLASS PostgresTarget AS Target
	cName = "PostgreSQL"
	cDriver = ""
	cServer = ""
	nPort = 0
	cUser = ""
	cPassword = ""
	cSuperPassword = ""
	cDatabase = ""

	FUNCTION Setup(tcDir)
		LOCAL lcIni
		This.cDir = ADDBS(tcDir)
		lcIni = FILETOSTR(PROBE_INI)
		lcIni = SUBSTR(lcIni, AT("[postgresql]", lcIni))
		This.cDriver = This.IniValue(lcIni, "driver")
		This.cServer = This.IniValue(lcIni, "server")
		This.nPort = VAL(This.IniValue(lcIni, "port"))
		This.cUser = This.IniValue(lcIni, "user")
		This.cPassword = This.IniValue(lcIni, "password")
		This.cSuperPassword = This.IniValue(lcIni, "super_password")
		This.cDatabase = LOWER("fr_probe" + SYS(2015))
		RETURN This.ServerRun("CREATE DATABASE " + This.cDatabase) ;
			AND This.RawRun("CREATE TABLE customers (id SERIAL PRIMARY KEY, name VARCHAR(40), city VARCHAR(40), born DATE)") ;
			AND This.RawRun("CREATE TABLE nopk (code VARCHAR(10), descr VARCHAR(40))") ;
			AND This.RawRun("INSERT INTO customers (name, city, born) VALUES ('Ana', 'Madrid', '2020-01-02')")
	ENDFUNC

	FUNCTION IniValue(tcIni, tcKey)
		RETURN ALLTRIM(STREXTRACT(CHR(10) + STRTRAN(tcIni, CHR(13), "") + CHR(10), CHR(10) + tcKey + "=", CHR(10)))
	ENDFUNC

	FUNCTION ConnStr
		RETURN "DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";PORT=" + TRANSFORM(This.nPort) + ";UID=" + This.cUser + ;
			";PWD={" + This.cPassword + "};DATABASE=" + IIF(EMPTY(This.cDatabase), "postgres", This.cDatabase) + ";"
	ENDFUNC

	FUNCTION ServerRun(tcSql)
		LOCAL lcDb, lvOk
		lcDb = This.cDatabase
		This.cDatabase = ""
		lvOk = This.RawRun(tcSql)
		This.cDatabase = lcDb
		RETURN lvOk
	ENDFUNC

	FUNCTION ServerScalar(tcSql)
		LOCAL lcDb, lvValue
		lcDb = This.cDatabase
		This.cDatabase = ""
		lvValue = This.RawScalar(tcSql)
		This.cDatabase = lcDb
		RETURN lvValue
	ENDFUNC

	FUNCTION ServerVersion
		RETURN LEFT(TRANSFORM(This.RawScalar("SELECT version()")), 40) + " / client_encoding " + TRANSFORM(This.RawScalar("SELECT current_setting('client_encoding')"))
	ENDFUNC

	* The PostgreSQL class as it is: the shape of its connection string, and where
	* Connect() really lands (user postgres, whose default database exists).
	PROCEDURE ExtraChecks(toProbe)
		LOCAL loDb, lcCs, lvOk, lvDb
		loDb = CREATEOBJECT("PostgreSQL")
		This.Configure(loDb, This.cDatabase)
		loDb.cUser = "postgres"
		loDb.cPassword = "<pwd>"
		lcCs = loDb.getConnectionString(.T.)
		toProbe.Say("P0a", IIF(";DATABASE=" $ UPPER(lcCs), "OK", "BUG"), "getConnectionString(.T.) = " + lcCs)
		loDb.cPassword = This.cSuperPassword
		lvOk = loDb.Connect()
		lvDb = .NULL.
		IF VARTYPE(lvOk) == "L" AND lvOk
			IF loDb.SQLExec("SELECT current_database() AS db", "c_pgdb")
				lvDb = ALLTRIM(c_pgdb.db)
				USE IN c_pgdb
			ENDIF
		ENDIF
		loDb.disconnect()
		toProbe.Say("P0b", IIF(TRANSFORM(lvDb) == This.cDatabase, "OK", "BUG"), "plain class as postgres, cDatabase = " + This.cDatabase + ": Connect() = " + TRANSFORM(lvOk) + ;
			", current_database() = " + TRANSFORM(lvDb) + "; cLastError: " + toProbe.Short(loDb.cLastError))
		loDb = CREATEOBJECT("PostgreSQL")
		This.Configure(loDb, This.cDatabase)
		loDb.cPassword = "wrong-but-no-semicolon"
		lvOk = loDb.Connect()
		toProbe.Say("P0c", "INFO", "plain class as " + This.cUser + " (no database of that name), wrong password: Connect() = " + TRANSFORM(lvOk) + "; cLastError: " + toProbe.Short(loDb.cLastError))
		loDb.disconnect()
	ENDPROC

	* The probes connect through a subclass that fixes only the connection string
	* (password in {}, the ';' before DATABASE, and the maintenance database when there is none).
	FUNCTION NewRemote(tcDatabase)
		LOCAL loDb
		loDb = CREATEOBJECT("FixedConnPostgreSQL")
		This.Configure(loDb, IIF(EMPTY(tcDatabase), This.cDatabase, tcDatabase))
		RETURN loDb
	ENDFUNC

	PROCEDURE Configure(toDb, tcDatabase)
		toDb.bShowErrors = .F.
		toDb.cDriver = This.cDriver
		toDb.cServer = This.cServer
		toDb.nPort = This.nPort
		toDb.cUser = This.cUser
		toDb.cPassword = This.cPassword
		toDb.cDatabase = tcDatabase
	ENDPROC

	FUNCTION MissingDatabase
		RETURN LOWER("fr_missing" + SYS(2015))
	ENDFUNC

	FUNCTION DatabaseExists(tcDatabase)
		RETURN VAL(TRANSFORM(This.ServerScalar("SELECT COUNT(*) FROM pg_database WHERE datname = '" + tcDatabase + "'"))) > 0
	ENDFUNC

	PROCEDURE DropDatabase(tcDatabase)
		This.ServerRun("DROP DATABASE IF EXISTS " + tcDatabase + " WITH (FORCE)")
	ENDPROC

	FUNCTION SemicolonValue
		RETURN This.cPassword
	ENDFUNC

	FUNCTION SemicolonWhat
		RETURN "user with ';' in the password"
	ENDFUNC

	FUNCTION SemicolonShow(tcValue)
		RETURN "<hidden, " + TRANSFORM(LEN(tcValue)) + " chars, has ';'=" + TRANSFORM(";" $ tcValue) + ">"
	ENDFUNC

	FUNCTION SemicolonWhere(tcValue)
		RETURN ""
	ENDFUNC

	FUNCTION NewRemoteWithSemicolon(tcValue)
		LOCAL loDb
		loDb = CREATEOBJECT("PostgreSQL")
		This.Configure(loDb, This.cDatabase)
		loDb.cPassword = tcValue
		RETURN loDb
	ENDFUNC

	FUNCTION TextType(tnLen)
		RETURN "VARCHAR(" + TRANSFORM(tnLen) + ")"
	ENDFUNC

	FUNCTION Concat(tc1, tc2, tc3)
		RETURN "RTRIM(" + tc1 + ") || " + tc2 + " || RTRIM(" + tc3 + ")"
	ENDFUNC

	FUNCTION DateAsText(tcField)
		RETURN "CAST(" + tcField + " AS TEXT)"
	ENDFUNC

	PROCEDURE Teardown
		This.DropDatabase(This.cDatabase)
	ENDPROC
ENDDEFINE

DEFINE CLASS FixedConnPostgreSQL AS PostgreSQL
	* Also fixes the primary key lookup: the original upper-cases the name before
	* ::regclass, the query fails and getPrimaryKey() raises error 13 (measured 04-10).
	PROCEDURE getPrimaryKeyScript(tcTable)
		RETURN "SELECT a.attname AS column_name FROM pg_index i JOIN pg_attribute a ON a.attrelid = i.indrelid AND a.attnum = ANY(i.indkey)" + ;
			" WHERE i.indrelid = '" + tcTable + "'::regclass AND i.indisprimary"
	ENDPROC

	FUNCTION getConnectionString(tbAddDatabase)
		RETURN "DRIVER={" + This.cDriver + "};SERVER=" + This.cServer + ";PORT=" + TRANSFORM(This.nPort) + ;
			";UID=" + This.cUser + ";PWD={" + This.cPassword + "};DATABASE=" + IIF(tbAddDatabase, ALLTRIM(This.cDatabase), "postgres") + ;
			";APP=FoxRemote probe"
	ENDFUNC
ENDDEFINE

* ======================================================================== *
* Firebird. User FR_PROBE (password with ';'), from engines.local.ini [firebird].
* ======================================================================== *
DEFINE CLASS FirebirdTarget AS Target
	cName = "Firebird"
	cDriver = ""
	cServer = ""
	nPort = 0
	cUser = ""
	cPassword = ""
	cDbFile = ""

	FUNCTION Setup(tcDir)
		LOCAL lcIni, lnH, laErr[1]
		This.cDir = ADDBS(tcDir)
		lcIni = FILETOSTR(PROBE_INI)
		lcIni = SUBSTR(lcIni, AT("[firebird]", lcIni))
		This.cDriver = This.IniValue(lcIni, "driver")
		This.cServer = This.IniValue(lcIni, "server")
		This.nPort = VAL(This.IniValue(lcIni, "port"))
		This.cUser = This.IniValue(lcIni, "user")
		This.cPassword = This.IniValue(lcIni, "password")
		This.cDbFile = This.cDir + "fr_" + SYS(2015) + ".fdb"
		* CREATE DATABASE from a connection to the sample database.
		lnH = SQLSTRINGCONNECT("DRIVER={" + This.cDriver + "};DBNAME=" + This.DbName("employee") + ";UID=" + This.cUser + ;
			";PWD={" + This.cPassword + "};CHARSET=WIN1252;", .T.)
		IF lnH < 1
			=AERROR(laErr)
			This.cLastError = "connect to employee: " + TRANSFORM(laErr[2])
			RETURN .F.
		ENDIF
		IF SQLEXEC(lnH, "CREATE DATABASE '" + This.DbName(This.cDbFile) + "' USER '" + This.cUser + "' PASSWORD '" + This.cPassword + ;
				"' PAGE_SIZE 8192 DEFAULT CHARACTER SET WIN1252") < 1
			=AERROR(laErr)
			This.cLastError = "CREATE DATABASE: " + TRANSFORM(laErr[2])
			=SQLDISCONNECT(lnH)
			RETURN .F.
		ENDIF
		=SQLDISCONNECT(lnH)
		RETURN This.RawRun("CREATE TABLE customers (id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY, name VARCHAR(40), city VARCHAR(40), born DATE)") ;
			AND This.RawRun("CREATE TABLE nopk (code VARCHAR(10), descr VARCHAR(40))") ;
			AND This.RawRun("INSERT INTO customers (name, city, born) VALUES ('Ana', 'Madrid', '2020-01-02')")
	ENDFUNC

	FUNCTION IniValue(tcIni, tcKey)
		RETURN ALLTRIM(STREXTRACT(CHR(10) + STRTRAN(tcIni, CHR(13), "") + CHR(10), CHR(10) + tcKey + "=", CHR(10)))
	ENDFUNC

	FUNCTION DbName(tcFile)
		RETURN This.cServer + "/" + TRANSFORM(This.nPort) + ":" + tcFile
	ENDFUNC

	FUNCTION ConnStr
		RETURN "DRIVER={" + This.cDriver + "};DBNAME=" + This.DbName(This.cDbFile) + ";UID=" + This.cUser + ";PWD={" + This.cPassword + "};CHARSET=WIN1252;"
	ENDFUNC

	FUNCTION Adapt(tcSql)
		RETURN STRTRAN(tcSql, "RTRIM(", "TRIM(TRAILING FROM ")
	ENDFUNC

	FUNCTION ServerVersion
		RETURN TRANSFORM(This.RawScalar("SELECT rdb$get_context('SYSTEM', 'ENGINE_VERSION') FROM rdb$database"))
	ENDFUNC

	* The probes connect through a subclass that only wraps the password in {}.
	FUNCTION NewRemote(tcDatabase)
		LOCAL loDb
		loDb = CREATEOBJECT("BracedFirebird")
		This.Configure(loDb, IIF(EMPTY(tcDatabase), This.DbName(This.cDbFile), tcDatabase))
		RETURN loDb
	ENDFUNC

	PROCEDURE Configure(toDb, tcDatabase)
		toDb.bShowErrors = .F.
		toDb.cDriver = This.cDriver
		toDb.cUser = This.cUser
		toDb.cPassword = This.cPassword
		toDb.cDatabase = tcDatabase
	ENDPROC

	FUNCTION MissingDatabase
		RETURN This.DbName(This.cDir + "missing_" + SYS(2015) + ".fdb")
	ENDFUNC

	FUNCTION DatabaseExists(tcDatabase)
		RETURN FILE(SUBSTR(tcDatabase, AT(":", tcDatabase) + 1))
	ENDFUNC

	PROCEDURE DropDatabase(tcDatabase)
		LOCAL lcFile
		lcFile = SUBSTR(tcDatabase, AT(":", tcDatabase) + 1)
		IF FILE(lcFile)
			TRY
				ERASE (lcFile)
			CATCH
			ENDTRY
		ENDIF
	ENDPROC

	FUNCTION SemicolonValue
		RETURN This.cPassword
	ENDFUNC

	FUNCTION SemicolonWhat
		RETURN "user with ';' in the password"
	ENDFUNC

	FUNCTION SemicolonShow(tcValue)
		RETURN "<hidden, " + TRANSFORM(LEN(tcValue)) + " chars, has ';'=" + TRANSFORM(";" $ tcValue) + ">"
	ENDFUNC

	FUNCTION SemicolonWhere(tcValue)
		RETURN ""
	ENDFUNC

	FUNCTION NewRemoteWithSemicolon(tcValue)
		LOCAL loDb
		loDb = CREATEOBJECT("Firebird")
		This.Configure(loDb, This.DbName(This.cDbFile))
		loDb.cPassword = tcValue
		RETURN loDb
	ENDFUNC

	FUNCTION TextType(tnLen)
		RETURN "VARCHAR(" + TRANSFORM(tnLen) + ")"
	ENDFUNC

	FUNCTION Concat(tc1, tc2, tc3)
		RETURN "RTRIM(" + tc1 + ") || " + tc2 + " || RTRIM(" + tc3 + ")"
	ENDFUNC

	FUNCTION DateAsText(tcField)
		RETURN "CAST(" + tcField + " AS VARCHAR(30))"
	ENDFUNC

	PROCEDURE Teardown
		This.DropDatabase(This.DbName(This.cDbFile))
	ENDPROC
ENDDEFINE

DEFINE CLASS BracedFirebird AS Firebird
	* Also drops the identifier quotes: the original writes "customers" (lower case,
	* quoted) and Firebird answers "Table unknown" for a table created as CUSTOMERS (measured 04-10).
	PROCEDURE Init
		DODEFAULT()
		This.cLeft = ""
		This.cRight = ""
	ENDPROC

	FUNCTION getConnectionString(tbAddDatabase)
		RETURN "DRIVER={" + This.cDriver + "};DBNAME=" + This.cDatabase + ";UID=" + This.cUser + ";PWD={" + This.cPassword + "};APP=FoxRemote probe"
	ENDFUNC
ENDDEFINE
