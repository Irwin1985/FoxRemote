# FoxRemote

One API for remote databases from Visual FoxPro 9: SQL Server, MySQL, MariaDB, PostgreSQL,
Firebird and SQLite, through ODBC. Connect, query with named parameters, use transactions,
edit server tables through buffered VFP cursors, and copy DBF tables to the server. It never
shows a dialog, so it runs the same in an application, a service or a COM server. One `.prg`,
no dependencies.

## Install

```
foxpack add foxremote
```

```foxpro
SET PROCEDURE TO foxremote.prg ADDITIVE
```

VFP is a 32-bit program, so it needs the **32-bit ODBC driver** of each engine, even on 64-bit
Windows. These are the drivers it is tested with:

| Engine | Class | `cDriver` | Tested with |
|---|---|---|---|
| SQL Server | `RemoteSqlServer` | `ODBC Driver 17 for SQL Server` | SQL Server 2025 Express |
| MySQL | `RemoteMySql` | `MariaDB ODBC 3.2 Driver` | MySQL 8.4.11 |
| MariaDB | `RemoteMariaDb` | `MariaDB ODBC 3.2 Driver` | MariaDB 11.8 |
| PostgreSQL | `RemotePostgreSql` | `PostgreSQL ANSI` (psqlODBC) | PostgreSQL 18.6 |
| Firebird | `RemoteFirebird` | `Firebird ODBC Driver` 3.0 | Firebird 5.0.4 |
| SQLite | `RemoteSqlite` | `SQLite3 ODBC Driver` 0.99991 | SQLite 3.43 |

For MySQL, use MariaDB's driver: Oracle's own is broken for VFP from 8.0.35 on, and 8.0.33
crashed VFP in the tests (see [Pitfalls](#pitfalls)). Firebird also needs the 32-bit `fbclient.dll`. Minimum server versions:
SQLite 3.35, MariaDB 10.5 and Firebird 3, because the library reads new ids with `RETURNING`.

## Example

Paste it into a `.prg` and run it. It uses SQLite, which needs no server, only its driver.

```foxpro
SET PROCEDURE TO foxremote.prg ADDITIVE

LOCAL loDb, loT, loC, loP, lcFile
lcFile = ADDBS(SYS(2023)) + "foxremote_demo.db"

loDb = CREATEOBJECT("RemoteSqlite")
loDb.cDriver = "SQLite3 ODBC Driver"
loDb.cDatabase = lcFile
IF !loDb.DatabaseExists(lcFile)
	loDb.CreateDatabase(lcFile)
ENDIF
IF !loDb.Connect()
	? loDb.cLastErrorCode, loDb.cLastError
	RETURN
ENDIF

* The table, the first time
IF !loDb.TableExists("customers")
	loT = loDb.NewTableDef("customers")
	loC = loT.AddColumn("id", "I")
	loC.lAutoIncrement = .T.
	loT.AddColumn("name", "C", 40)
	loT.AddColumn("city", "C", 40)
	loT.AddColumn("born", "D")
	loDb.CreateTable(loT)
ENDIF

* Write through a buffered cursor; each new row gets the id the server gave it
loDb.Open("customers", "c_customers")
INSERT INTO c_customers (name, city) VALUES ("Ana", "Lima")
INSERT INTO c_customers (name, city, born) VALUES ("Beto", "Oslo", DATE(1990, 5, 1))
IF loDb.Save("c_customers")
	? "new id:", c_customers.id
ELSE
	? loDb.cLastErrorCode, loDb.cLastError
ENDIF
loDb.Close("c_customers")

* Read with a named parameter
loP = CREATEOBJECT("Empty")
ADDPROPERTY(loP, "city", "Oslo")
IF loDb.Query("SELECT id, name FROM customers WHERE city = :city", "c_oslo", loP)
	SELECT c_oslo
	SCAN
		? c_oslo.id, c_oslo.name
	ENDSCAN
	USE IN c_oslo
ENDIF
? "customers:", loDb.Scalar("SELECT COUNT(*) FROM customers")

* Two statements, all or nothing
IF loDb.BeginTransaction()
	IF loDb.Execute("UPDATE customers SET city = :city WHERE name = 'Ana'", loP) >= 0 ;
			AND loDb.Execute("DELETE FROM customers WHERE born IS NULL AND name <> 'Ana'") >= 0
		loDb.Commit()
	ELSE
		loDb.Rollback()
	ENDIF
ENDIF

loDb.Disconnect()
```

The same code runs on every engine. Only the class and the connection properties change:

```foxpro
loDb = CREATEOBJECT("RemoteSqlServer")
loDb.cDriver = "ODBC Driver 17 for SQL Server"
loDb.cServer = "localhost\SQLEXPRESS"
loDb.cDatabase = "erp"
loDb.lTrustedConnection = .T.        && Windows authentication; or cUser and cPassword

loDb = CREATEOBJECT("RemoteMySql")   && RemoteMariaDb for MariaDB
loDb.cDriver = "MariaDB ODBC 3.2 Driver"
loDb.cServer = "127.0.0.1"
loDb.nPort = 3306
loDb.cDatabase = "erp"
loDb.cUser = "erp_user"
loDb.cPassword = "secret"

loDb = CREATEOBJECT("RemotePostgreSql")
loDb.cDriver = "PostgreSQL ANSI"
loDb.cServer = "127.0.0.1"
loDb.cDatabase = "erp"
loDb.cUser = "erp_user"
loDb.cPassword = "secret"

loDb = CREATEOBJECT("RemoteFirebird")
loDb.cDriver = "Firebird ODBC Driver"
loDb.cDatabase = "127.0.0.1/3050:C:\Data\erp.fdb"   && host/port:path, a local path or an alias
loDb.cUser = "SYSDBA"
loDb.cPassword = "secret"
```

## API

Every engine class inherits from `RemoteDatabase`. A method that fails returns `.F.` (or `-1`,
or `.NULL.`, as each one says), adds one to `nErrors` and leaves the reason in the error
properties. Nothing is ever shown on screen.

### Connection properties

| Property | Default | What it is |
|---|---|---|
| `cDriver` | `""` | The ODBC driver name, as the ODBC administrator (32 bits) lists it |
| `cServer` | `""` | Host, or `host\instance` for SQL Server. Not used by Firebird and SQLite |
| `nPort` | `0` | TCP port; 0 is the engine's default |
| `cDatabase` | `""` | Database name. Firebird: `host/port:path`, a path or an alias. SQLite: the file |
| `cUser`, `cPassword` | `""` | Credentials. A password may hold `;`, `{`, `}` or `=` |
| `lTrustedConnection` | `.F.` | SQL Server only: Windows authentication instead of `cUser` and `cPassword` |
| `cConnectionOptions` | `""` | Extra `key=value;` pairs added to the connection string |
| `cConnectionString` | `""` | The whole connection string. When set, `Connect()` uses it instead of the properties above |
| `nConnectTimeout` | `15` | Seconds to wait for the server when connecting |
| `nQueryTimeout` | `0` | Seconds a statement may run; 0 waits forever |
| `lAutoReconnect` | `.F.` | Opens the connection again when the server dropped it. See [Reconnection](#reconnection) |
| `nIdleCheckSeconds` | `30` | With `lAutoReconnect`, a connection idle this long is checked before the next statement; 0 checks before every one |
| `nReconnects` | `0` | How many times the connection was opened again, read-only |
| `nHandle` | `0` | The ODBC handle, read-only; 0 when not connected |
| `cEngine`, `cVersion` | | `"sqlserver"`, `"mysql"`, ...; and the library version, `"1.0.0"` |

### Data properties

| Property | Default | What it does |
|---|---|---|
| `cEmptyDateMode` | `"null"` | An empty VFP date or datetime goes to the server as NULL. `"min"` sends the engine's lowest date instead (also used for a NOT NULL column) |
| `lTrimOnSave` | `.T.` | Trailing blanks of character fields are cut before they are sent |
| `cAutoKeyName` | `""` | When set, `CreateTable()` (and so `Migrate()`) adds an auto-increment primary key with this name to every table |
| `lDetectConflicts` | `.F.` | `Save()` refuses to overwrite a column another user changed after the cursor was read (`row_changed`). See [Save()](#savecalias) |

### Connect()

Opens the connection to `cDatabase`. It never creates the database: a missing one fails with
`database_not_found`. Returns `.T.` or `.F.`.

```foxpro
IF !loDb.Connect()
	? loDb.cLastErrorCode, loDb.cLastError
ENDIF
```

### Disconnect()

Closes the connection. A transaction still open is rolled back first, and that is reported as
an error, `rolled_back_on_disconnect`, so it is never lost in silence. Releasing the object
disconnects too.

### DatabaseExists(cDatabase), CreateDatabase(cDatabase)

Whether the database exists, and create it (`database_exists` if it already does). On a server
they use a short connection of their own, so they work before `Connect()`. SQLite creates the
file. Firebird creates it from a connection to another database of the same server, so
`cDatabase` has to name an existing database or alias at that moment; and Firebird's
`DatabaseExists()` can only check a path on this machine, so for a remote server it answers
`.T.` and `Connect()` is the real test.

```foxpro
IF !loDb.DatabaseExists("erp")
	loDb.CreateDatabase("erp")
ENDIF
```

### UseDatabase(cDatabase)

Reconnects to another database of the same server. If it fails, `cDatabase` keeps the old
name and the object stays disconnected.

### Attach(nHandle)

Works on a connection that something else opened with `SQLCONNECT()` or `SQLSTRINGCONNECT()`.
`Disconnect()` and releasing the object leave that connection open.

```foxpro
lnH = SQLSTRINGCONNECT(lcMyConnectionString, .T.)
loDb.Attach(lnH)
```

### Ping()

`.T.` if the server answers a trivial query on the open connection.

### Reconnection

A server drops connections: MySQL and MariaDB close the ones idle for 8 hours by default, and
a restart or a network cut drops them all. With `lAutoReconnect = .T.`:

- A connection idle for `nIdleCheckSeconds` is checked before the next statement, and before
  `BeginTransaction()`, and opened again if it is gone. The statement then runs on the new one.
- A statement that finds the connection gone while it runs is repeated on a new connection only
  when it is a `SELECT`. A write is not repeated, because the server may have run it before the
  connection went: it fails with `connection_lost`, and the connection is already open again for
  the next one.
- Inside a transaction nothing is repeated: the server rolled the transaction back with the
  connection. The statement fails with `connection_lost` and `lInTransaction` is `.F.`; start the
  transaction again.

Without `lAutoReconnect`, a dropped connection fails with `connection_lost` and `nHandle`
becomes 0. A connection made with `Attach()` is never opened again by the library.

```foxpro
loDb.lAutoReconnect = .T.
IF loDb.Execute("UPDATE stock SET qty = 0 WHERE id = 7") < 0 AND loDb.cLastErrorCode == "connection_lost"
	* it may or may not have run: look before running it again
ENDIF
```

### Query(cSql, cCursor, oParams)

Runs a statement and leaves the result in the cursor `cCursor` (`c_fr_query` if empty).
Returns `.T.` or `.F.`. The cursor is a plain VFP cursor: use `Open()` to edit a table.

`:name` in the SQL takes its value from the property `name` of `oParams`. Any object works:
`CREATEOBJECT("Empty")` with `ADDPROPERTY()`, or `SCATTER NAME`. The value travels as an ODBC
parameter, never pasted into the SQL. A `:` inside quotes, and PostgreSQL's `::` cast, are left
alone.

```foxpro
loP = CREATEOBJECT("Empty")
ADDPROPERTY(loP, "city", "Oslo")
ADDPROPERTY(loP, "since", DATE(1980, 1, 1))
loDb.Query("SELECT * FROM customers WHERE city = :city AND born >= :since", "c_list", loP)
```

### Execute(cSql, oParams)

Runs a statement that returns no rows. Returns the number of rows affected (0 when the driver
does not say), or `-1` on failure. MySQL and MariaDB count the rows an `UPDATE` changed, so one
that writes the values a row already has returns 0.

```foxpro
loP = CREATEOBJECT("Empty")
ADDPROPERTY(loP, "city", "Lima")
lnRows = loDb.Execute("DELETE FROM customers WHERE city = :city", loP)
```

### Scalar(cSql, oParams)

The first column of the first row; `.NULL.` when there is no row or the statement failed
(check `nErrors` to tell them apart). A `BIGINT`, which VFP receives as text, comes back as a
number.

```foxpro
lnCount = loDb.Scalar("SELECT COUNT(*) FROM customers")
```

### LastId()

The id generated by the last `INSERT` run on this connection. A row saved with `Save()`
already has its id in the cursor, so this is for an `Execute("INSERT ...")`. See the Firebird
pitfall.

### NextNumber(cTable, cColumn, cWhere, oParams)

Adds one to the counter `cColumn` of the one row `cWhere` picks, and returns the new value, or
`-1` on failure: `row_not_found` when no row matches, `bad_argument` when several do (none of
them is changed). A NULL counter counts from 0.

The `UPDATE` locks that row until the transaction ends, so two sessions never get the same
number. Called inside your transaction, the row stays locked until your `Commit()` or
`Rollback()`: the number and the document that uses it are saved together, and a rollback
gives the number back, so there are no gaps. Without a transaction, `NextNumber()` runs in one
of its own. Keep the transaction short, because every other session that needs that counter
waits for it.

```foxpro
loP = CREATEOBJECT("Empty")
ADDPROPERTY(loP, "series", "A")
loDb.BeginTransaction()
lnNumber = loDb.NextNumber("invoice_series", "last_no", "series = :series", loP)
IF lnNumber > 0 AND loDb.Execute("INSERT INTO invoices (series, invoice_no) VALUES ('A', " + TRANSFORM(lnNumber) + ")") = 1
	loDb.Commit()
ELSE
	loDb.Rollback()
ENDIF
```

### BeginTransaction(), Commit(), Rollback()

A transaction on the connection. They nest: only the outermost `Commit()` commits, and any
`Rollback()` undoes everything and ends the transaction. `lInTransaction` and
`nTransactionLevel` tell where you are. Outside a transaction every statement commits on its
own.

```foxpro
loDb.BeginTransaction()
IF loDb.Execute("UPDATE stock SET qty = qty - 1 WHERE id = 7") = 1
	loDb.Commit()
ELSE
	loDb.Rollback()
ENDIF
```

### Open(cTable, cAlias, cWhere, oParams, cFields, lReadOnly, cGroup)

Reads a table into a cursor you can edit, with optimistic table buffering (`Buffering` 5). Only
`cTable` is needed. `cAlias` defaults to the table name; `cWhere` and `oParams` filter the
rows, with `:name` parameters as in `Query()`; `cFields` picks the columns (the primary key
must be among them); `cGroup` puts the cursor in a group for `SaveGroup()`. Returns `.T.` and
selects the cursor.

A table without a primary key can only be opened with `lReadOnly`: the library would not know
which row to change. An alias already in use is refused.

```foxpro
loP = CREATEOBJECT("Empty")
ADDPROPERTY(loP, "id", 42)
loDb.Open("orders", "c_order", "id = :id", loP, "", .F., "order")
loDb.Open("order_lines", "c_lines", "order_id = :id", loP, "", .F., "order")
```

### Save(cAlias)

Sends the changes of the cursor to the server: inserts, updates and deletes, in one
transaction. An update sends only the columns that changed, so it does not overwrite what
another user changed in the other columns. New rows get the id the server gave them, in the
same round trip. Returns `.T.` or `.F.`

When two users change the same column, the last `Save()` wins, unless `lDetectConflicts` is
`.T.`: then the `UPDATE` also requires the columns it changes to still have the values the
cursor read, as `WhereType` 3 does in a VFP remote view, and `Save()` fails with `row_changed`
when another user changed one of them. `Refresh()` after `Discard()` shows what the other user
wrote. Memo, binary, floating point (`B`, `F`) and datetime columns are not compared, because
their values do not come back equal (VFP keeps no fractions of a second, for one).

When it fails, the transaction is rolled back so nothing of the cursor reaches the server, the
edits stay in the buffer to fix and try again, and `cLastErrorCode` is `save_failed`,
`row_not_found` when the row to update or delete is no longer on the server, `row_changed`, or
`connection_lost`. Inside a
transaction you opened, `Save()` runs in it and leaves the `Commit()` or `Rollback()` to you,
also when it fails.

```foxpro
loDb.Open("orders", "c_order")
SELECT c_order
REPLACE status WITH "paid"
IF !loDb.Save("c_order")
	? loDb.cLastErrorCode, loDb.cLastError
ENDIF
```

### SaveGroup(cGroup)

`Save()` of every editable cursor of the group, in one transaction: all of them reach the
server, or none.

### Discard(cAlias)

Throws away the changes not saved (`TABLEREVERT`).

### Refresh(cAlias)

Reads the cursor again from the server, with the same filter and parameters. It is refused
(`pending_changes`) while the cursor has changes not saved: `Save()` or `Discard()` first.

### HasChanges(cAlias)

`.T.` when the cursor has changes not saved. Without an alias, `.T.` when any cursor this
object opened in the current data session has. A read-only cursor never has.

```foxpro
loDb.Open("orders", "c_order")
REPLACE status WITH "sent" IN c_order
IF loDb.HasChanges("c_order")
	loDb.Save("c_order")
ENDIF
```

### Close(cAlias, lDiscard), CloseGroup(cGroup, lDiscard), CloseAll(lDiscard)

Closes one cursor, the cursors of a group, or every cursor this object opened in the current
data session. A cursor with changes not saved is not closed: it fails with `pending_changes`,
and a group or `CloseAll()` then closes none of them. With `lDiscard` the changes are thrown
away and the cursors closed. Cursors opened by other code are never touched. Returns `.T.` or
`.F.`

### GetTables(cCursor), GetColumns(cTable, cCursor)

A cursor with the tables of the database (one field, `name`), or with the columns of a table:
`name` (as the server's catalog has it), `type` (the VFP type letter), `width`, `decimals`,
`nullable`, `primary_key`. The cursors default to `c_fr_tables` and `c_fr_columns`. Return
`.T.` or `.F.`

```foxpro
IF loDb.GetColumns("customers", "c_cols")
	SELECT c_cols
	SCAN
		? c_cols.name, c_cols.type, c_cols.width, c_cols.primary_key
	ENDSCAN
ENDIF
```

### GetPrimaryKey(cTable), TableExists(cTable), ColumnExists(cTable, cColumn)

The primary key's columns separated by commas (`""` when there is none), and whether a table
or a column exists. Names compare without regard to case.

### ServerVersion(), ServerDate(), NewGuid()

The server's version text, its date and time, and a new GUID as text, all from the server.

### NewTableDef(cName), CreateTable(oTableDef), TableScript(oTableDef)

`NewTableDef()` returns a `RemoteTableDef` to describe a table. `CreateTable()` creates it with
its primary key, foreign keys and indexes; `TableScript()` returns those statements, separated
by `;` and a line break, without running them.

```foxpro
loT = loDb.NewTableDef("order_lines")
loC = loT.AddColumn("id", "I")
loC.lAutoIncrement = .T.
loC = loT.AddColumn("order_id", "I")
loC.lNullable = .F.
loT.AddColumn("product", "C", 30)
loC = loT.AddColumn("qty", "N", 10, 2)
loC.cDefault = "0"
loT.AddIndex("ix_lines_product", "product")
loT.AddForeignKey("order_id", "orders", "id", "CASCADE")
? loDb.TableScript(loT)
```

`RemoteTableDef` methods:

| Method | What it does |
|---|---|
| `AddColumn(cName, cType, nWidth, nDecimals)` | Adds a column of a VFP type (`C`, `V`, `M`, `N`, `I`, `B`, `F`, `Y`, `D`, `T`, `L`; anything else is a blob). Returns the `RemoteColumnDef` |
| `AddIndex(cName, cColumns, lUnique)` | An index on `"a"` or `"a, b"` |
| `AddForeignKey(cColumns, cRefTable, cRefColumns, cOnDelete, cOnUpdate)` | The actions are `CASCADE`, `SET NULL` (or `NULL`), `SET DEFAULT` (or `DEFAULT`), `RESTRICT`, `NO ACTION` |

`RemoteColumnDef` properties: `lNullable` (`.T.`), `lPrimaryKey` (several make a composite
key), `lAutoIncrement` (an auto-increment primary key; type and width are ignored), `cDefault`
(the SQL text of the default, such as `"0"` or `"'new'"`).

### CreateTableFromCursor(cAlias, cTable)

Creates a table with the structure of an open cursor or DBF (`cTable` defaults to the alias).
It copies no rows.

### Migrate(cSource, lReplace), GenerateScript(cSource)

`Migrate()` copies DBF tables to the server: one `.dbf`, every table of a `.dbc`, or every
`.dbf` of a folder. Each table is created and then filled in a transaction of its own, so a
row that fails leaves that table on the server empty and the others as they are. Deleted
records are skipped. A table that already exists on the server is refused unless `lReplace`,
which drops it first. It returns a report object with `nTables`, `nRows`, `nErrors` and `cErrors`
(one line per table that failed).

`GenerateScript()` returns the `CREATE TABLE` statements of the same sources without running
anything.

```foxpro
loR = loDb.Migrate("C:\Data\erp.dbc")
? loR.nTables, loR.nRows
IF loR.nErrors > 0
	? loR.cErrors
ENDIF
```

### NaturalName(cName)

A name the way the engine writes it when it is not quoted: lower case in PostgreSQL, MySQL,
MariaDB and SQLite, upper case in Firebird, as written in SQL Server.

## Errors

| Property | What it holds |
|---|---|
| `nErrors` | How many failures since the object was created or since `ClearErrors()` |
| `cLastError` | The last one, in one line: the library's text in English, then the driver's message, which comes in the server's language |
| `cLastErrorCode` | A stable code, from the table below |
| `nLastError` | The driver's native error number, when it came from the server |
| `cLastSqlState` | The ODBC SQLSTATE, when it came from the server |
| `cLastSql` | The last statement the library sent |
| `lStrict` | `.F.`; with `.T.` every failure also raises a VFP error with `cLastError` as message. `Save()`, `SaveGroup()`, `NextNumber()` and `Migrate()` raise it after they have rolled back their own transaction |

`ClearErrors()` sets them back to zero and empty.

```foxpro
loDb.lStrict = .T.
TRY
	loDb.Execute("UPDATE nowhere SET x = 1")
CATCH TO loEx
	? loDb.cLastErrorCode, loEx.Message
ENDTRY
```

| Code | When |
|---|---|
| `connection_failed` | The server could not be reached, or refused the credentials |
| `database_not_found` | `Connect()` to a database that does not exist |
| `database_exists` | `CreateDatabase()` of one that does |
| `not_connected` | A statement without a connection (and `lAutoReconnect` off) |
| `connection_lost` | The server dropped the connection; see [Reconnection](#reconnection) for what was repeated and what was not |
| `query_failed` | The server rejected a statement; `cLastError` has its message |
| `transaction_failed` | A transaction could not start, commit or roll back |
| `no_transaction` | `Commit()` or `Rollback()` without `BeginTransaction()` |
| `rolled_back_on_disconnect` | `Disconnect()` found a transaction open and rolled it back |
| `bad_argument` | A required argument is missing, or a source to migrate has no tables |
| `alias_in_use` | `Open()` with an alias that is already open |
| `no_primary_key` | `Open()` for writing of a table without a primary key |
| `not_open` | `Save()`, `Close()` and the rest on an alias this object did not open |
| `read_only` | `Save()` of a cursor opened read-only |
| `pending_changes` | `Refresh()` or `Close()` (without `lDiscard`) of a cursor with changes not saved |
| `save_failed` | `Save()` or `SaveGroup()` failed; the cause is in `cLastError` |
| `row_not_found` | `Save()` found that the row to update or delete is no longer on the server, or `NextNumber()` found no row |
| `row_changed` | With `lDetectConflicts`, another user changed a column `Save()` was going to write |
| `table_exists` | `Migrate()` without `lReplace` of a table that is already on the server |
| `source_not_open` | `Migrate()` could not open a DBF |

## Pitfalls

- **MySQL: Oracle's ODBC driver is broken for VFP from 8.0.35 on.** The row count is right,
  but every value arrives as the text `varbinary` or `-3`, with the ANSI and the Unicode driver
  and whatever options you set. Oracle publishes no 32-bit driver after 8.0.46, so it will not
  be fixed. Use the MariaDB driver (3.2 works with MySQL 8.4). `MySQL ODBC 8.0 ANSI Driver`
  8.0.33 reads values right, but after a statement that gave up waiting for a lock, the next
  statement on that connection crashed VFP (`C0000005`) two times out of three; the Unicode
  driver delivers memo fields in UCS-2.
- **SQL Server: `VARCHAR(MAX)` and `NVARCHAR(MAX)` arrive in `Query()` as `C(0)`, empty.** That
  is how VFP types them with ODBC Driver 17. `Open()` reads them right, but in your own SQL you
  have to cast them: `CAST(CAST(notes AS VARCHAR(MAX)) AS TEXT) AS notes` reaches VFP as a memo.
- **Firebird: `LastId()` is not safe with several users.** Firebird has no "last id of this
  session"; `LastId()` reads the current value of the identity generator of the table of your
  last `Execute("INSERT ...")`, which another session may have moved. `Save()` does not have
  this problem: it reads the id with `RETURNING`.
- **SQL Server: `LastId()` reads `@@IDENTITY`**, so after an `Execute("INSERT ...")` into a table
  whose trigger inserts into another table with an identity, it returns the trigger's id. Each
  `Execute()` is a batch of its own, where `SCOPE_IDENTITY()` is already empty. `Save()` reads
  `SCOPE_IDENTITY()` in the same batch as its `INSERT`, so it gets the row's own id, and it works
  on tables with triggers.
- **Firebird needs the 32-bit `fbclient.dll`.** The 32-bit ODBC driver loads it; the 64-bit
  server installer copies it to `SysWOW64` when you ask for it.
- **`Open()` turns on `MULTILOCKS`** in the data session of the cursor, because table buffering
  requires it. Nothing else global is changed.
- **Names are sent unquoted**, so the engine writes them its own way: `customers` in
  PostgreSQL, MySQL, MariaDB and SQLite, `CUSTOMERS` in Firebird. Tables created outside the
  library with quoted mixed-case names (`"Customers"` in PostgreSQL) cannot be reached by name.
  Tables and columns made from VFP (`Migrate()`, `CreateTableFromCursor()`) go in lower case in
  SQL Server.
- **An empty date is NULL on the server**, and a NULL comes back to VFP as `.NULL.`, not as an
  empty date. Test with `EMPTY(NVL(born, {}))`, or set `cEmptyDateMode = "min"`.
- **A `?variable` in your SQL does not see your local variables.** The statement runs inside the
  library, where a `LOCAL` of yours is out of scope. Use `:name` and `oParams`.
- **`Rollback()` after a `Save()` inside your own transaction leaves the cursor showing the saved
  values**, which are no longer on the server. `Refresh()` it, or `Close()` and `Open()` again.
- **SQL Server's ODBC Driver 17 reopens a dropped idle connection by itself**
  (`ConnectRetryCount`, 1 by default), outside a transaction. The library's own reconnection
  works on top of it; add `ConnectRetryCount=0;` to `cConnectionOptions` to leave it to the
  library alone.
- **Firebird waits for a locked row forever.** Its ODBC driver starts transactions with `WAIT`
  and no limit, so a `Save()` or a `NextNumber()` behind another session's open transaction
  hangs until that one ends. The driver's `LOCKTIMEOUT` option does not reach the transactions
  the library starts; on Firebird 4 and later, run `loDb.Execute("SET STATEMENT TIMEOUT 10
  SECOND")` after `Connect()`. The other engines have their own limit: SQL Server
  `SET LOCK_TIMEOUT`, MySQL and MariaDB `innodb_lock_wait_timeout` (50 seconds), PostgreSQL
  `lock_timeout`, SQLite the driver's `Timeout` option.
- **A reserved word cannot be a name**, because names are sent unquoted. They differ by engine:
  `last_value` is fine in MariaDB and a syntax error in MySQL 8, where it is a window function.
- **SQLite stores text as UTF-8 and its driver does not convert it.** The library converts what
  it sends and what it reads in `Query()`, `Scalar()` and `Open()`; text you read with your own
  `SQLEXEC()` on `nHandle` arrives in UTF-8.

## Changes from 0.x

The 1.0 is a new API. The 0.x is kept as it was in [legacy/foxremote.prg](legacy/foxremote.prg).

| 0.x | 1.0 |
|---|---|
| `CREATEOBJECT("MSSQL")` ... `"SQLite"` | `RemoteSqlServer`, `RemoteMySql`, `RemoteMariaDb`, `RemotePostgreSql`, `RemoteFirebird`, `RemoteSqlite` |
| `Connect()` creates the database; `Connect(.T.)` | `Connect()` never creates it; `CreateDatabase()` |
| `newDataBase`, `changeDB`, `selectDatabase` | `CreateDatabase`, `UseDatabase(cName)` |
| `use(table, fields, where, group, ro, nodata)` | `Open(table, alias, where, params, fields, ro, group)` |
| `Save`, `saveGroup`, `saveAndClose` | `Save`, `SaveGroup` |
| `requery` | `Refresh`, refused with changes not saved |
| `discard`, `Close`, `closeAll`, `closeGroup` | `Discard`; `Close`, `CloseAll`, `CloseGroup`, which keep a cursor with changes unless `lDiscard` |
| `SQLExec(sql, cursor)` | `Query`, `Execute`, `Scalar`, with `:parameters` |
| `getTables()`, `getTableFields()` (arrays) | `GetTables(cursor)`, `GetColumns(table, cursor)` |
| `getLastID()` | The id is already in the saved row; `LastId()` after an `Execute` |
| `getServerDate`, `getNewGuid`, `getVersion`, `Ping` ("Pong") | `ServerDate`, `NewGuid`, `ServerVersion`, `Ping` (logical) |
| `migrate(source, toScripts)` | `Migrate(source, lReplace)`, `GenerateScript(source)` |
| `createTable(name, desc, indexes, array, toScripts)` | `CreateTable(oTableDef)`, `CreateTableFromCursor(alias)` |
| `bUseCA`, `bShowErrors`, `bUseSymbolDelimiter`, `aCustomArray` | Removed |
| `cLastError`, a long text in Spanish | `cLastError` in one line of English, plus `cLastErrorCode` |
| | New: `BeginTransaction`, `Commit`, `Rollback`, `Attach`, `NextNumber`, `HasChanges`, `lDetectConflicts`, `lStrict`, `nErrors`, `cLastSql` |

## License

MIT. See [LICENSE](LICENSE).
