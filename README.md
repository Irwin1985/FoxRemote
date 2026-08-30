# FoxRemote

**Biblioteca de abstraccion multi-motor SQL para Visual FoxPro 9**

FoxRemote permite conectar, consultar y migrar datos entre VFP y los principales motores de bases de datos relacionales usando una API unificada.

---

## Caracteristicas

- **Multi-motor**: MSSQL, MySQL, MariaDB, Firebird, SQLite, PostgreSQL
- **Migracion de datos**: Desde DBF libres o contenedores DBC
- **Cursores actualizables**: CursorAdapter o SQL directo
- **Transacciones manuales**: Control total sobre commits y rollbacks
- **Grupos de cursores**: Actualizar multiples tablas en una sola transaccion
- **Manejo de errores**: Con `bShowErrors` y `cLastError`
- **Reconexion automatica**: Detecta conexiones caidas y reconecta

---

## Inicio rapido

### 1. Cargar la biblioteca

```foxpro
SET PROCEDURE TO foxremote.prg ADDITIVE
```

### 2. Crear instancia del motor

```foxpro
* MySQL
loDB = CREATEOBJECT("MySQL")
loDB.cDriver = "MySQL ODBC 8.0 ANSI Driver"
loDB.cServer = "localhost"
loDB.cDatabase = "mi_base"
loDB.cUser = "root"
loDB.cPassword = "secret"

* MSSQL
loDB = CREATEOBJECT("MSSQL")
loDB.cDriver = "SQL Server Native Client 11.0"
loDB.cServer = "SERVIDOR\INSTANCIA"
loDB.cDatabase = "mi_base"
loDB.cUser = "sa"
loDB.cPassword = "secret"

* PostgreSQL
loDB = CREATEOBJECT("PostgreSQL")
loDB.cDriver = "PostgreSQL ANSI"
loDB.cServer = "localhost"
loDB.cDatabase = "mi_base"
loDB.cUser = "postgres"
loDB.cPassword = "secret"
loDB.nPort = 5432

* Firebird
loDB = CREATEOBJECT("Firebird")
loDB.cDriver = "Firebird/Interbase(r) driver"
loDB.cDatabase = "C:\datos\mi_base.fdb"
loDB.cUser = "SYSDBA"
loDB.cPassword = "masterkey"

* SQLite
loDB = CREATEOBJECT("SQLite")
loDB.cDriver = "SQLite3 ODBC Driver"
loDB.cDatabase = "C:\datos\mi_base.db"
```

### 3. Conectar

```foxpro
IF !loDB.Connect()
    MESSAGEBOX("Error de conexion: " + loDB.cLastError)
    RETURN
ENDIF
```

### 4. Usar

```foxpro
* Abrir tabla como cursor actualizable
loDB.use("clientes")

* Con alias personalizado
loDB.use("clientes AS cli")

* Solo lectura
loDB.use("clientes", , , , .T.)

* Con filtro
loDB.use("clientes", , "activo = 1")

* Campos especificos
loDB.use("clientes", "id, nombre, email")
```

### 5. Trabajar con datos

```foxpro
SELECT clientes

* Agregar registro
INSERT INTO clientes (nombre, email) VALUES ("Juan", "juan@mail.com")

* Modificar registro
REPLACE nombre WITH "Juan Perez"

* Eliminar registro
DELETE

* Guardar cambios (transaccion)
IF loDB.Save("clientes")
    MESSAGEBOX("Guardado correctamente")
ELSE
    MESSAGEBOX("Error: " + loDB.cLastError)
ENDIF
```

### 6. Cerrar

```foxpro
* Cerrar cursor
loDB.Close("clientes")

* Cerrar todos los cursores
loDB.closeAll()

* Desconectar
loDB.disconnect()

* Liberar objeto
loDB = .NULL.
```

---

## Migracion de datos

### Desde DBF libre

```foxpro
loDB.migrate("clientes.dbf")
```

### Desde contenedor DBC

```foxpro
loDB.migrate("midatabase.dbc")
```

### Desde directorio completo

```foxpro
loDB.migrate("C:\datos\vfp\")
```

### Generar script SQL sin ejecutar

```foxpro
loScripts = CREATEOBJECT("Collection")
loDB.migrate("clientes.dbf", loScripts)

lcSQL = ""
FOR EACH lcScript IN loScripts
    lcSQL = lcSQL + lcScript
ENDFOR

STRTOFILE(lcSQL, "schema.sql")
```

---

## API completa

### Propiedades de conexion

| Propiedad | Tipo | Descripcion |
|---|---|---|
| `cDriver` | C | Nombre del driver ODBC |
| `cServer` | C | Servidor (host\instancia) |
| `cDatabase` | C | Nombre de la base de datos |
| `cUser` | C | Usuario |
| `cPassword` | C | Contrasena |
| `nPort` | N | Puerto (opcional) |

### Propiedades de comportamiento

| Propiedad | Tipo | Default | Descripcion |
|---|---|---|---|
| `bUseCA` | L | .T. | Usar CursorAdapter (sino, SQL directo) |
| `bShowErrors` | L | .T. | Mostrar MessageBox en errores |
| `cLastError` | C | "" | Ultimo mensaje de error |
| `cPKName` | C | "TID" | Nombre de la clave primaria por defecto |

### Metodos principales

#### `Connect([tbAddDatabase])` → L
Conecta al servidor. Si `tbAddDatabase = .T.`, no crea la base de datos automaticamente.

```foxpro
loDB.Connect()           && Crea la BD si no existe
loDB.Connect(.T.)        && Solo conecta, no crea BD
```

#### `use(tcTable, [tcFields], [tcCriteria], [tcGroup], [tbReadOnly], [tbNodata])` → L
Abre una tabla remota como cursor local actualizable.

| Parametro | Tipo | Descripcion |
|---|---|---|
| `tcTable` | C | Nombre de tabla. Soporta alias: `"clientes AS cli"` |
| `tcFields` | C | Campos a seleccionar (default: `*`) |
| `tcCriteria` | C | Clausula WHERE (sin la palabra WHERE) |
| `tcGroup` | C | Grupo para actualizacion grupal |
| `tbReadOnly` | L | Solo lectura (default: .F.) |
| `tbNodata` | L | Solo estructura, sin datos (default: .F.) |

#### `Save([tcAlias])` → L
Guarda cambios del cursor en el servidor dentro de una transaccion.

#### `saveGroup(tcGroup)` → L
Guarda todos los cursores de un grupo en una sola transaccion.

#### `Close([tcAlias])` → L
Cierra un cursor y libera recursos.

#### `closeAll()`
Cierra todos los cursores abiertos.

#### `requery([tcAlias])`
Recarga datos del servidor manteniendo cambios pendientes.

#### `discard([tcAlias])`
Descarta cambios pendientes del cursor.

#### `saveAndClose([tcAlias])` → L
Guarda y cierra en un solo paso.

#### `disconnect()`
Desconecta del servidor.

### Metodos de consulta

#### `SQLExec(tcSQL, [tcCursor])` → L
Ejecuta SQL directo.

```foxpro
loDB.SQLExec("SELECT * FROM clientes WHERE activo = 1", "curActivos")
SELECT curActivos
BROWSE
```

#### `getTables()` → Array
Retorna lista de tablas en la base de datos actual.

```foxpro
laTablas = loDB.getTables()
FOR i = 1 TO ALEN(laTablas)
    ? laTablas[i]
ENDFOR
```

#### `getTableFields(tcTable)` → Array
Retorna lista de campos de una tabla.

#### `tableExists(tcTable)` → L
Verifica si una tabla existe.

#### `fieldExists(tcTable, tcField)` → L
Verifica si un campo existe en una tabla.

#### `getPrimaryKey(tcTable)` → C
Retorna el nombre del campo primary key.

#### `getServerDate()` → D/T
Retorna la fecha/hora del servidor.

#### `getNewGuid()` → C
Retorna un GUID/UUID generado por el servidor.

#### `getLastID()` → N
Retorna el ultimo ID auto-generado (despues de un INSERT).

### Metodos de migracion

#### `migrate(tcTableOrPath, [toScripts])` → L
Migra tablas desde DBF/DBC al servidor remoto.

| Parametro | Tipo | Descripcion |
|---|---|---|
| `tcTableOrPath` | C | Ruta a archivo DBF, DBC o directorio |
| `toScripts` | Collection | Si se pasa, genera SQL sin ejecutar |

#### `createTable(tcName, tcDesc, toIndexes, taFields, [toScripts])` → L
Crea una tabla en el servidor con la estructura especificada.

### Metodos de base de datos

#### `newDataBase(tcDatabase)` → L
Crea una nueva base de datos.

#### `changeDB(tcDatabase)` → L
Cambia a otra base de datos.

#### `selectDatabase()`
Selecciona la base de datos actual (USE database).

#### `getVersion()` → C
Retorna la version del servidor.

#### `Ping()` → C
Verifica conexion. Retorna `"Pong"`.

---

## Motores soportados

### MSSQL (SQL Server)

```foxpro
loDB = CREATEOBJECT("MSSQL")
loDB.cDriver = "SQL Server Native Client 11.0"
loDB.cServer = "SERVIDOR\INSTANCIA"
```

- Delimitadores: `[tabla]`
- Primary key: `INT IDENTITY(1,1)`
- UUID: `UNIQUEIDENTIFIER DEFAULT NEWID()`
- Fecha minima: 1753-01-01

### MySQL

```foxpro
loDB = CREATEOBJECT("MySQL")
loDB.cDriver = "MySQL ODBC 8.0 ANSI Driver"
loDB.cServer = "localhost"
```

- Delimitadores: `` `tabla` ``
- Primary key: `INT AUTO_INCREMENT`
- UUID: `VARCHAR(36)`
- Fecha minima: 1000-01-01
- Engine por defecto: InnoDB

### MariaDB

```foxpro
loDB = CREATEOBJECT("MariaDB")
loDB.cDriver = "MariaDB ODBC 3.1 Driver"
```

Hereda de MySQL. Mismas caracteristicas.

### PostgreSQL

```foxpro
loDB = CREATEOBJECT("PostgreSQL")
loDB.cDriver = "PostgreSQL ANSI"
loDB.nPort = 5432
```

- Delimitadores: `"tabla"`
- Primary key: `SERIAL`
- UUID: `uuid_generate_v4()` (requiere extension `uuid-ossp`)
- Fecha minima: 1753-01-01
- FK e indices se ejecutan por separado

### Firebird

```foxpro
loDB = CREATEOBJECT("Firebird")
loDB.cDriver = "Firebird/Interbase(r) driver"
loDB.cDatabase = "C:\datos\base.fdb"
```

- Delimitadores: `"TABLA"` (mayusculas por defecto)
- Primary key: `INTEGER GENERATED BY DEFAULT AS IDENTITY`
- UUID: No soportado nativamente
- Fecha minima: 1858-11-18
- FK e indices se ejecutan por separado

### SQLite

```foxpro
loDB = CREATEOBJECT("SQLite")
loDB.cDriver = "SQLite3 ODBC Driver"
loDB.cDatabase = "C:\datos\base.db"
```

- Delimitadores: `"tabla"`
- Primary key: `INTEGER PRIMARY KEY AUTOINCREMENT`
- UUID: `HEX(RANDOMBLOB(16))`
- Fecha minima: 0001-01-01
- Indices se ejecutan por separado

---

## Mapeo de tipos VFP → SQL

| VFP | MSSQL | MySQL | PostgreSQL | Firebird | SQLite |
|---|---|---|---|---|---|
| C(n) | CHAR(n) | CHAR(n) | VARCHAR(n) | CHAR(n) | TEXT(n) |
| V(n) | VARCHAR(n) | VARCHAR(n) | VARCHAR(n) | VARCHAR(n) | TEXT(n) |
| N(p,s) | NUMERIC(p,s) | DECIMAL(p,s) | NUMERIC(p,s) | NUMERIC(p,s) | NUMERIC(p,s) |
| I | INT | INT | INTEGER | INTEGER | INTEGER |
| B | FLOAT | DOUBLE | NUMERIC | DECIMAL | DOUBLE |
| F | FLOAT | FLOAT | NUMERIC | DECIMAL | REAL |
| Y | MONEY | DECIMAL | NUMERIC | DECIMAL | NUMERIC |
| L | BIT | BOOL | BOOLEAN | BOOLEAN | BOOLEAN |
| D | DATE | DATE | DATE | DATE | DATE |
| T | DATETIME | DATETIME | TIMESTAMP | TIMESTAMP | DATETIME |
| M | TEXT | TEXT | TEXT | BLOB SUB_TYPE 1 | TEXT |
| W | IMAGE | BLOB | BYTEA | BLOB | BLOB |
| G | IMAGE | BLOB | BYTEA | BLOB SUB_TYPE 0 | BLOB |
| Q | VARBINARY(max) | VARBINARY(n) | BYTEA | BLOB | BLOB |
| U | UNIQUEIDENTIFIER | VARCHAR(36) | UUID | VARCHAR(36) | VARCHAR(36) |

---

## Grupos de cursores

Los grupos permiten actualizar multiples tablas en una sola transaccion:

```foxpro
* Abrir tablas en el mismo grupo
loDB.use("clientes", , , "ventas")
loDB.use("facturas", , , "ventas")
loDB.use("detalles", , , "ventas")

* Trabajar con los cursores
SELECT clientes
INSERT INTO clientes (nombre) VALUES ("Juan")

SELECT facturas
INSERT INTO facturas (cliente_id, fecha) VALUES (1, DATE())

* Guardar todo en una transaccion
IF loDB.saveGroup("ventas")
    MESSAGEBOX("Todo guardado")
ELSE
    MESSAGEBOX("Error: " + loDB.cLastError)
ENDIF

* Cerrar grupo
loDB.closeGroup("ventas")
```

---

## Manejo de errores

```foxpro
* Desactivar MessageBox automaticos
loDB.bShowErrors = .F.

* Verificar operaciones
IF !loDB.use("clientes")
    MESSAGEBOX("Error: " + loDB.cLastError)
ENDIF

* Verificar despues de SQLExec
IF !loDB.SQLExec("SELECT * FROM tabla_inexistente")
    MESSAGEBOX("Error SQL: " + loDB.cLastError)
ENDIF
```

---

## Cursores de solo lectura

```foxpro
* Abrir como solo lectura (mas rapido, no permite Save)
loDB.use("clientes", , , , .T.)

* Abrir solo estructura (sin datos)
loDB.use("clientes", , , , , .T.)
```

---

## Transacciones manuales

FoxRemote maneja transacciones automaticamente en `Save()`, pero puedes usar transacciones manuales:

```foxpro
loDB.SQLExec("BEGIN TRANSACTION")

IF loDB.SQLExec("INSERT INTO clientes (nombre) VALUES ('Juan')")
    IF loDB.SQLExec("INSERT INTO facturas (cliente_id) VALUES (1)")
        loDB.SQLExec("COMMIT")
    ELSE
        loDB.SQLExec("ROLLBACK")
    ENDIF
ELSE
    loDB.SQLExec("ROLLBACK")
ENDIF
```

---

## Consideraciones

### Clave primaria

FoxRemote busca primero un campo llamado `TID` (configurable en `cPKName`). Si no existe, detecta el primary key de la tabla remota.

### Reconexion automatica

Si la conexion se pierde, FoxRemote intenta reconectar automaticamente en el siguiente `Connect()`.

### CursorAdapter vs SQL directo

- `bUseCA = .T.` (default): Usa CursorAdapter de VFP (recomendado)
- `bUseCA = .F.`: Usa SQL directo con INSERT/UPDATE/DELETE manuales

### Entorno de fecha

FoxRemote ajusta automaticamente `SET DATE`, `SET CENTURY` y `SET MARK` para compatibilidad con cada motor. Al finalizar, restaura los valores originales.

---

## Requisitos

- Visual FoxPro 9 SP2
- Drivers ODBC instalados para cada motor
- VBScript.RegExp (incluido en Windows)

---

## Licencia

Desarrollado por Irwin Rodriguez.

---

## Ver tambien

- [handoff.md](handoff.md) — Notas tecnicas y propuestas futuras (DBML parser)
