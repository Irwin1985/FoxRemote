* FoxRemoteTests.prg
* Unit tests for FoxRemote multi-database abstraction library.
* Tests cover type mapping, script generation, and base helpers.
* No database connection required -- all tests exercise pure functions.
* Run: foxunit run --prg FoxRemote\FoxRemoteTests.prg --format console

* =========================================================================
* Shared helper class: simulates a field descriptor object
* =========================================================================
Define Class FieldDef As Custom
    Type          = ""
    Size          = "10"
    Decimal       = "0"
    allowNull     = .F.
    longName      = ""
    Comment       = ""
    Nextvalue     = 0
    stepValue     = 0
    addDefault    = .T.
    Default       = "''"
    primaryKey    = .F.
    foreignKey    = .Null.
    autoIncrement = .F.
    index         = .Null.
    tag           = ""
EndDefine

* =========================================================================
* MSSQL Tests
* =========================================================================
Define Class TestMSSQL As Custom

    Procedure SetUp
        Set Procedure To "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\foxremote" Additive
    Endproc

    Procedure TearDown
    Endproc

    Procedure TestPing() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        __assert.Equal("Pong", loDB.Ping())
    Endproc

    Procedure TestDelimiters() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        __assert.Equal("[", loDB.cLeft)
        __assert.Equal("]", loDB.cRight)
    Endproc

    Procedure TestMaxLength() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        __assert.Equal(128, loDB.nMaxLength)
    Endproc

    Procedure TestTidScript() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        __assert.Equal("INT IDENTITY(1,1) PRIMARY KEY", loDB.getTidScript())
    Endproc

    Procedure TestUUIDScript() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        __assert.Equal("UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID()", loDB.getUUIDScript())
    Endproc

    Procedure TestDropTable() HELP [Fact]
        Local loDB, lcScript
        loDB = CreateObject("MSSQL")
        lcScript = Alltrim(loDB.dropTable("orders"))
        __assert.Equal("DROP TABLE IF EXISTS [orders];", lcScript)
    Endproc

    Procedure TestAddAutoIncrement() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        __assert.Equal("IDENTITY(1,1)", loDB.addAutoIncrement())
    Endproc

    Procedure TestVisitCType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        loF.Size = "50"
        __assert.True(At("CHAR(50)", loDB.visitCType(loF)) > 0)
    Endproc

    Procedure TestVisitNTypeWithDecimal() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        loF.Size = "12"
        loF.Decimal = "2"
        __assert.Equal("NUMERIC(12,2)", loDB.visitNType(loF))
    Endproc

    Procedure TestVisitIType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("INT", loDB.visitIType(loF))
    Endproc

    Procedure TestVisitLType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("BIT", loDB.visitLType(loF))
    Endproc

    Procedure TestVisitMType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("TEXT", loDB.visitMType(loF))
    Endproc

    Procedure TestVisitDType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("DATE", loDB.visitDType(loF))
    Endproc

    Procedure TestVisitTType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("DATETIME", loDB.visitTType(loF))
    Endproc

    Procedure TestVisitBType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("FLOAT", loDB.visitBType(loF))
    Endproc

    Procedure TestVisitYType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("MONEY", loDB.visitYType(loF))
    Endproc

    Procedure TestVisitGType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("IMAGE", loDB.visitGType(loF))
    Endproc

    Procedure TestVisitQType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("VARBINARY(max)", loDB.visitQType(loF))
    Endproc

    Procedure TestVisitVTypeWithSize() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        loF.Size = "200"
        __assert.True(At("VARCHAR(200)", loDB.visitVType(loF)) > 0)
    Endproc

    Procedure TestVisitVTypeNoSize() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        loF.Size = "0"
        __assert.True(At("VARCHAR(max)", loDB.visitVType(loF)) > 0)
    Endproc

    Procedure TestConnectionStringNoPort() HELP [Fact]
        Local loDB, lcStr
        loDB = CreateObject("MSSQL")
        loDB.cDriver   = "SQL Server"
        loDB.cServer   = "localhost"
        loDB.cUser     = "sa"
        loDB.cPassword = "pass"
        lcStr = loDB.getConnectionString(.F.)
        __assert.True(At("DRIVER=SQL Server", lcStr) > 0)
        __assert.True(At("SERVER=localhost",   lcStr) > 0)
        __assert.True(At("UID=sa",             lcStr) > 0)
        __assert.True(At("PWD=pass",           lcStr) > 0)
        __assert.True(At("PORT=",              lcStr) = 0)
    Endproc

    Procedure TestConnectionStringWithPort() HELP [Fact]
        Local loDB, lcStr
        loDB = CreateObject("MSSQL")
        loDB.cDriver   = "SQL Server"
        loDB.cServer   = "localhost"
        loDB.cUser     = "sa"
        loDB.cPassword = "pass"
        loDB.nPort     = 1433
        lcStr = loDB.getConnectionString(.F.)
        __assert.True(At("PORT=1433", lcStr) > 0)
    Endproc

    Procedure TestConnectionStringWithDatabase() HELP [Fact]
        Local loDB, lcStr
        loDB = CreateObject("MSSQL")
        loDB.cDriver   = "SQL Server"
        loDB.cServer   = "localhost"
        loDB.cUser     = "sa"
        loDB.cPassword = "pass"
        loDB.cDatabase = "mydb"
        lcStr = loDB.getConnectionString(.T.)
        __assert.True(At("DATABASE=mydb", lcStr) > 0)
    Endproc

    Procedure TestDTypeSetsDefault() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        loDB.visitDType(loF)
        __assert.Equal("'1753-01-01'", loF.Default)
    Endproc

    Procedure TestTTypeSetsDefault() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MSSQL")
        loF = CreateObject("FieldDef")
        loDB.visitTType(loF)
        __assert.Equal("'1753-01-01 00:00:00.000'", loF.Default)
    Endproc

EndDefine

* =========================================================================
* MySQL Tests
* =========================================================================
Define Class TestMySQL As Custom

    Procedure SetUp
        Set Procedure To "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\foxremote" Additive
    Endproc

    Procedure TearDown
    Endproc

    Procedure TestPing() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.Equal("Pong", loDB.Ping())
    Endproc

    Procedure TestDelimiters() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.Equal("`", loDB.cLeft)
        __assert.Equal("`", loDB.cRight)
    Endproc

    Procedure TestMaxLength() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.Equal(64, loDB.nMaxLength)
    Endproc

    Procedure TestTidScript() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.Equal("INT AUTO_INCREMENT PRIMARY KEY", loDB.getTidScript())
    Endproc

    Procedure TestDropTable() HELP [Fact]
        Local loDB, lcScript
        loDB = CreateObject("MySQL")
        lcScript = Alltrim(loDB.dropTable("clientes"))
        __assert.Equal("DROP TABLE IF EXISTS `clientes`;", lcScript)
    Endproc

    Procedure TestVisitCType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        loF.Size = "100"
        __assert.Equal("CHAR(100)", loDB.visitCType(loF))
    Endproc

    Procedure TestVisitIType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("INT", loDB.visitIType(loF))
    Endproc

    Procedure TestVisitLType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("BOOL", loDB.visitLType(loF))
    Endproc

    Procedure TestVisitNTypeWithDecimal() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        loF.Size = "10"
        loF.Decimal = "2"
        __assert.Equal("DECIMAL(10,2)", loDB.visitNType(loF))
    Endproc

    Procedure TestVisitYType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        loF.Size = "15"
        loF.Decimal = "4"
        __assert.Equal("DECIMAL(15,4)", loDB.visitYType(loF))
    Endproc

    Procedure TestVisitDType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("DATE", loDB.visitDType(loF))
    Endproc

    Procedure TestVisitTType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("DATETIME", loDB.visitTType(loF))
    Endproc

    Procedure TestVisitMType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("TEXT", loDB.visitMType(loF))
    Endproc

    Procedure TestVisitGTypeDisablesDefault() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("BLOB", loDB.visitGType(loF))
        __assert.False(loF.addDefault)
    Endproc

    Procedure TestVisitBType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        loF.Size = "15"
        loF.Decimal = "4"
        __assert.Equal("DOUBLE(15,4)", loDB.visitBType(loF))
    Endproc

    Procedure TestVisitVType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        loF.Size = "255"
        __assert.Equal("VARCHAR(255)", loDB.visitVType(loF))
    Endproc

    Procedure TestDTypeSetsDefault() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("MySQL")
        loF = CreateObject("FieldDef")
        loDB.visitDType(loF)
        __assert.Equal("'1000-01-01'", loF.Default)
    Endproc

    Procedure TestConnectionStringHasCollation() HELP [Fact]
        Local loDB, lcStr
        loDB = CreateObject("MySQL")
        loDB.cDriver   = "MySQL ODBC 8.0 ANSI Driver"
        loDB.cServer   = "localhost"
        loDB.cUser     = "root"
        loDB.cPassword = "1234"
        lcStr = loDB.getConnectionString(.F.)
        __assert.True(At("COLLATION=utf8_general_ci", lcStr) > 0)
    Endproc

    Procedure TestTableOptionsHasInnoDB() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.True(At("InnoDB", loDB.createTableOptions()) > 0)
    Endproc

EndDefine

* =========================================================================
* MariaDB Tests
* =========================================================================
Define Class TestMariaDB As Custom

    Procedure SetUp
        Set Procedure To "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\foxremote" Additive
    Endproc

    Procedure TearDown
    Endproc

    Procedure TestInheritsFromMySQL() HELP [Fact]
        Local loDB
        loDB = CreateObject("MariaDB")
        __assert.Equal(64, loDB.nMaxLength)
        __assert.Equal("`", loDB.cLeft)
        __assert.Equal("INT AUTO_INCREMENT PRIMARY KEY", loDB.getTidScript())
    Endproc

    Procedure TestPing() HELP [Fact]
        Local loDB
        loDB = CreateObject("MariaDB")
        __assert.Equal("Pong", loDB.Ping())
    Endproc

EndDefine

* =========================================================================
* Firebird Tests
* =========================================================================
Define Class TestFirebird As Custom

    Procedure SetUp
        Set Procedure To "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\foxremote" Additive
    Endproc

    Procedure TearDown
    Endproc

    Procedure TestPing() HELP [Fact]
        Local loDB
        loDB = CreateObject("Firebird")
        __assert.Equal("Pong", loDB.Ping())
    Endproc

    Procedure TestDelimiters() HELP [Fact]
        Local loDB
        loDB = CreateObject("Firebird")
        __assert.Equal('"', loDB.cLeft)
        __assert.Equal('"', loDB.cRight)
    Endproc

    Procedure TestMaxLength() HELP [Fact]
        Local loDB
        loDB = CreateObject("Firebird")
        __assert.Equal(31, loDB.nMaxLength)
    Endproc

    Procedure TestTidScript() HELP [Fact]
        Local loDB
        loDB = CreateObject("Firebird")
        __assert.Equal("INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY", loDB.getTidScript())
    Endproc

    Procedure TestCannotGenerateGUID() HELP [Fact]
        Local loDB
        loDB = CreateObject("Firebird")
        __assert.False(loDB.bCanGenerateGUID)
    Endproc

    Procedure TestFkScriptSeparately() HELP [Fact]
        Local loDB
        loDB = CreateObject("Firebird")
        __assert.True(loDB.bExecuteFkScriptSeparately)
        __assert.True(loDB.bExecuteIndexScriptSeparately)
    Endproc

    Procedure TestDropTableUppercase() HELP [Fact]
        Local loDB, lcScript
        loDB = CreateObject("Firebird")
        lcScript = loDB.dropTable("orders")
        __assert.True(At("ORDERS", lcScript) > 0)
    Endproc

    Procedure TestDropTableWithDelimiters() HELP [Fact]
        Local loDB, lcScript
        loDB = CreateObject("Firebird")
        loDB.bUseSymbolDelimiter = .T.
        lcScript = loDB.dropTable("orders")
        __assert.True(At('"orders"', lcScript) > 0)
    Endproc

    Procedure TestVisitCType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        loF.Size = "60"
        __assert.Equal("CHAR(60)", loDB.visitCType(loF))
    Endproc

    Procedure TestVisitIType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        __assert.Equal("INTEGER", loDB.visitIType(loF))
    Endproc

    Procedure TestVisitLType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        __assert.Equal("BOOLEAN", loDB.visitLType(loF))
    Endproc

    Procedure TestVisitTType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        __assert.Equal("TIMESTAMP", loDB.visitTType(loF))
    Endproc

    Procedure TestVisitDType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        __assert.Equal("DATE", loDB.visitDType(loF))
    Endproc

    Procedure TestVisitMType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        __assert.Equal("BLOB SUB_TYPE 1", loDB.visitMType(loF))
    Endproc

    Procedure TestVisitGType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        __assert.Equal("BLOB SUB_TYPE 0", loDB.visitGType(loF))
    Endproc

    Procedure TestVisitNTypeWithDecimal() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        loF.Size = "10"
        loF.Decimal = "2"
        __assert.Equal("NUMERIC(10,2)", loDB.visitNType(loF))
    Endproc

    Procedure TestVisitNTypeInteger() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        loF.Size = "10"
        loF.Decimal = "0"
        __assert.Equal("INTEGER", loDB.visitNType(loF))
    Endproc

    Procedure TestVisitVTypeWithSize() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        loF.Size = "100"
        __assert.Equal("VARCHAR(100)", loDB.visitVType(loF))
    Endproc

    Procedure TestVisitVTypeDefaultsTo8191() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        loF.Size = "0"
        __assert.Equal("VARCHAR(8191)", loDB.visitVType(loF))
    Endproc

    Procedure TestConnectionStringHasDBNAME() HELP [Fact]
        Local loDB, lcStr
        loDB = CreateObject("Firebird")
        loDB.cDriver   = "Firebird/InterBase(r) driver"
        loDB.cDatabase = "C:\db\mydb.fdb"
        loDB.cUser     = "SYSDBA"
        loDB.cPassword = "masterkey"
        lcStr = loDB.getConnectionString(.F.)
        __assert.True(At("DBNAME=", lcStr) > 0)
        __assert.True(At("UID=SYSDBA", lcStr) > 0)
    Endproc

    Procedure TestDTypeSetsDefault() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("Firebird")
        loF = CreateObject("FieldDef")
        loDB.visitDType(loF)
        __assert.Equal("'1858-11-18'", loF.Default)
    Endproc

EndDefine

* =========================================================================
* SQLite Tests
* =========================================================================
Define Class TestSQLite As Custom

    Procedure SetUp
        Set Procedure To "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\foxremote" Additive
    Endproc

    Procedure TearDown
    Endproc

    Procedure TestPing() HELP [Fact]
        Local loDB
        loDB = CreateObject("SQLite")
        __assert.Equal("Pong", loDB.Ping())
    Endproc

    Procedure TestDelimiters() HELP [Fact]
        Local loDB
        loDB = CreateObject("SQLite")
        __assert.Equal('"', loDB.cLeft)
        __assert.Equal('"', loDB.cRight)
    Endproc

    Procedure TestMaxLength() HELP [Fact]
        Local loDB
        loDB = CreateObject("SQLite")
        __assert.Equal(128, loDB.nMaxLength)
    Endproc

    Procedure TestTidScriptHasINTEGER() HELP [Fact]
        Local loDB, lcTid
        loDB = CreateObject("SQLite")
        lcTid = loDB.getTidScript()
        __assert.True(At("INTEGER", lcTid) > 0)
        __assert.True(At("PRIMARY KEY", lcTid) > 0)
        __assert.True(At("AUTOINCREMENT", lcTid) > 0)
    Endproc

    Procedure TestDropTable() HELP [Fact]
        Local loDB, lcScript
        loDB = CreateObject("SQLite")
        lcScript = Alltrim(loDB.dropTable("products"))
        __assert.Equal('DROP TABLE IF EXISTS "products";', lcScript)
    Endproc

    Procedure TestIndexScriptSeparately() HELP [Fact]
        Local loDB
        loDB = CreateObject("SQLite")
        __assert.True(loDB.bExecuteIndexScriptSeparately)
    Endproc

    Procedure TestVisitCTypeWithSize() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        loF.Size = "50"
        __assert.Equal("TEXT(50)", loDB.visitCType(loF))
    Endproc

    Procedure TestVisitCTypeNoSize() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        loF.Size = "0"
        __assert.Equal("TEXT", loDB.visitCType(loF))
    Endproc

    Procedure TestVisitIType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        __assert.Equal("INTEGER", loDB.visitIType(loF))
    Endproc

    Procedure TestVisitLType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        __assert.Equal("BOOLEAN", loDB.visitLType(loF))
    Endproc

    Procedure TestVisitNTypeWithDecimal() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        loF.Size = "10"
        loF.Decimal = "3"
        __assert.Equal("NUMERIC(10,3)", loDB.visitNType(loF))
    Endproc

    Procedure TestVisitBType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        __assert.Equal("DOUBLE", loDB.visitBType(loF))
    Endproc

    Procedure TestVisitFType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        __assert.Equal("REAL", loDB.visitFType(loF))
    Endproc

    Procedure TestVisitMType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        __assert.Equal("TEXT", loDB.visitMType(loF))
    Endproc

    Procedure TestVisitGType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        __assert.Equal("BLOB", loDB.visitGType(loF))
    Endproc

    Procedure TestVisitQType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        __assert.Equal("BLOB", loDB.visitQType(loF))
    Endproc

    Procedure TestVisitVTypeWithSize() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        loF.Size = "255"
        __assert.Equal("TEXT(255)", loDB.visitVType(loF))
    Endproc

    Procedure TestVisitWType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        __assert.Equal("BLOB", loDB.visitWType(loF))
    Endproc

    Procedure TestDTypeSetsDefault() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("SQLite")
        loF = CreateObject("FieldDef")
        loDB.visitDType(loF)
        __assert.Equal("'0001-01-01'", loF.Default)
    Endproc

    Procedure TestConnectionStringHasDATABASE() HELP [Fact]
        Local loDB, lcStr
        loDB = CreateObject("SQLite")
        loDB.cDriver   = "SQLite3 ODBC Driver"
        loDB.cDatabase = "C:\data\mydb.db"
        lcStr = loDB.getConnectionString(.F.)
        __assert.True(At("C:\data\mydb.db", lcStr) > 0)
    Endproc

EndDefine

* =========================================================================
* PostgreSQL Tests
* =========================================================================
Define Class TestPostgreSQL As Custom

    Procedure SetUp
        Set Procedure To "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\foxremote" Additive
    Endproc

    Procedure TearDown
    Endproc

    Procedure TestPing() HELP [Fact]
        Local loDB
        loDB = CreateObject("PostgreSQL")
        __assert.Equal("Pong", loDB.Ping())
    Endproc

    Procedure TestDelimiters() HELP [Fact]
        Local loDB
        loDB = CreateObject("PostgreSQL")
        __assert.Equal('"', loDB.cLeft)
        __assert.Equal('"', loDB.cRight)
    Endproc

    Procedure TestMaxLength() HELP [Fact]
        Local loDB
        loDB = CreateObject("PostgreSQL")
        __assert.Equal(63, loDB.nMaxLength)
    Endproc

    Procedure TestTidScript() HELP [Fact]
        Local loDB
        loDB = CreateObject("PostgreSQL")
        __assert.Equal("SERIAL PRIMARY KEY", loDB.getTidScript())
    Endproc

    Procedure TestDropTable() HELP [Fact]
        Local loDB, lcScript
        loDB = CreateObject("PostgreSQL")
        lcScript = Alltrim(loDB.dropTable("invoices"))
        __assert.Equal('DROP TABLE IF EXISTS "invoices";', lcScript)
    Endproc

    Procedure TestFkAndIdxSeparately() HELP [Fact]
        Local loDB
        loDB = CreateObject("PostgreSQL")
        __assert.True(loDB.bExecuteFkScriptSeparately)
        __assert.True(loDB.bExecuteIndexScriptSeparately)
    Endproc

    Procedure TestVisitCTypeWithSize() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        loF.Size = "100"
        __assert.Equal("VARCHAR(100)", loDB.visitCType(loF))
    Endproc

    Procedure TestVisitCTypeNoSize() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        loF.Size = "0"
        __assert.Equal("VARCHAR", loDB.visitCType(loF))
    Endproc

    Procedure TestVisitIType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("INTEGER", loDB.visitIType(loF))
    Endproc

    Procedure TestVisitLType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("BOOLEAN", loDB.visitLType(loF))
    Endproc

    Procedure TestVisitNTypeWithDecimal() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        loF.Size = "15"
        loF.Decimal = "4"
        __assert.Equal("NUMERIC(15,4)", loDB.visitNType(loF))
    Endproc

    Procedure TestVisitNTypeInteger() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        loF.Size = "10"
        loF.Decimal = "0"
        __assert.Equal("INTEGER", loDB.visitNType(loF))
    Endproc

    Procedure TestVisitTType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("TIMESTAMP", loDB.visitTType(loF))
    Endproc

    Procedure TestVisitDType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("DATE", loDB.visitDType(loF))
    Endproc

    Procedure TestVisitMType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("TEXT", loDB.visitMType(loF))
    Endproc

    Procedure TestVisitGType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("BYTEA", loDB.visitGType(loF))
    Endproc

    Procedure TestVisitQType() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        __assert.Equal("BYTEA", loDB.visitQType(loF))
    Endproc

    Procedure TestVisitVTypeWithSize() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        loF.Size = "200"
        __assert.Equal("VARCHAR(200)", loDB.visitVType(loF))
    Endproc

    Procedure TestChangeTypeOnAutoIncrement() HELP [Fact]
        Local loDB
        loDB = CreateObject("PostgreSQL")
        __assert.Equal("SERIAL", loDB.changeTypeOnAutoIncrement("INTEGER"))
    Endproc

    Procedure TestConnectionStringDefaultPort() HELP [Fact]
        Local loDB, lcStr
        loDB = CreateObject("PostgreSQL")
        loDB.cDriver   = "PostgreSQL Unicode"
        loDB.cServer   = "localhost"
        loDB.cUser     = "postgres"
        loDB.cPassword = "secret"
        loDB.nPort     = 0
        lcStr = loDB.getConnectionString(.F.)
        __assert.True(At("SERVER=localhost", lcStr) > 0)
        __assert.True(At("5432", lcStr) > 0)
    Endproc

    Procedure TestDTypeSetsDefault() HELP [Fact]
        Local loDB, loF
        loDB = CreateObject("PostgreSQL")
        loF = CreateObject("FieldDef")
        loDB.visitDType(loF)
        __assert.Equal("'1858-11-18'", loF.Default)
    Endproc

EndDefine

* =========================================================================
* DBEngine Base Helper Tests
* =========================================================================
Define Class TestDBEngineHelpers As Custom

    Procedure SetUp
        Set Procedure To "C:\Desarrollo\IrwinRodriguez.dev\FoxRemote\foxremote" Additive
    Endproc

    Procedure TearDown
    Endproc

    Procedure TestGetForeignKeyValueNull() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.Equal("SET NULL", loDB.getForeignKeyValue("NULL"))
    Endproc

    Procedure TestGetForeignKeyValueDefault() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.Equal("SET DEFAULT", loDB.getForeignKeyValue("DEFAULT"))
    Endproc

    Procedure TestGetForeignKeyValueRestrict() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        __assert.Equal("NO ACTION", loDB.getForeignKeyValue("RESTRICT"))
    Endproc

    Procedure TestGetForeignKeyValueCascade() HELP [Fact]
        Local loDB
        loDB = CreateObject("PostgreSQL")
        __assert.Equal("CASCADE", loDB.getForeignKeyValue("CASCADE"))
    Endproc

    Procedure TestGetForeignKeyValueCaseInsensitive() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.Equal("SET NULL", loDB.getForeignKeyValue("null"))
        __assert.Equal("SET NULL", loDB.getForeignKeyValue("Null"))
    Endproc

    Procedure TestBUseSymbolDelimiterDoesNotClearDelimiters() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        loDB.bUseSymbolDelimiter = .F.
        __assert.Equal("[", loDB.cLeft)
        __assert.Equal("]", loDB.cRight)
    Endproc

    Procedure TestBUseSymbolDelimiterCanBeEnabled() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        loDB.bUseSymbolDelimiter = .T.
        __assert.True(loDB.bUseSymbolDelimiter)
        __assert.Equal("[", loDB.cLeft)
        __assert.Equal("]", loDB.cRight)
    Endproc

    Procedure TestDefaultPKName() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.Equal("TID", loDB.cPKName)
    Endproc

    Procedure TestBShowErrorsDefault() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.True(loDB.bShowErrors)
    Endproc

    Procedure TestAllEnginesPing() HELP [Fact]
        Local loEngines[5], i
        loEngines[1] = CreateObject("MSSQL")
        loEngines[2] = CreateObject("MySQL")
        loEngines[3] = CreateObject("Firebird")
        loEngines[4] = CreateObject("SQLite")
        loEngines[5] = CreateObject("PostgreSQL")
        For i = 1 To 5
            __assert.Equal("Pong", loEngines[i].Ping())
        EndFor
    Endproc

    Procedure TestFormatDateOrDateTimeMSSQLEmptyDate() HELP [Fact]
        Local loDB, ldResult
        loDB = CreateObject("MSSQL")
        ldResult = loDB.formatDateOrDateTime({})
        __assert.Equal(Date(1753, 01, 01), ldResult)
    Endproc

    Procedure TestFormatDateOrDateTimeMySQLEmptyDate() HELP [Fact]
        Local loDB, ldResult
        loDB = CreateObject("MySQL")
        ldResult = loDB.formatDateOrDateTime({})
        __assert.Equal(Date(1000, 01, 01), ldResult)
    Endproc

    Procedure TestFormatDateOrDateTimeNonEmpty() HELP [Fact]
        Local loDB, ldInput, ldResult
        loDB    = CreateObject("MySQL")
        ldInput = Date(2024, 6, 15)
        ldResult = loDB.formatDateOrDateTime(ldInput)
        __assert.Equal(ldInput, ldResult)
    Endproc

    Procedure TestMSSQLAddPrimaryKey() HELP [Fact]
        Local loDB
        loDB = CreateObject("MSSQL")
        __assert.Equal("PRIMARY KEY", loDB.addPrimaryKey())
    Endproc

    Procedure TestMySQLAddPrimaryKey() HELP [Fact]
        Local loDB
        loDB = CreateObject("MySQL")
        __assert.Equal("PRIMARY KEY", loDB.addPrimaryKey())
    Endproc

EndDefine
