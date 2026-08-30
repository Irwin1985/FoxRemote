# FoxRemote — Handoff

## Estado actual

FoxRemote es una biblioteca VFP de abstraccion multi-motor SQL (MSSQL, MySQL, MariaDB, Firebird, SQLite, PostgreSQL).
Escrito por Irwin Rodriguez.

**Archivo unico**: `foxremote.prg` (3741 lineas)
**Motores soportados**: MSSQL, MySQL, MariaDB, Firebird, SQLite, PostgreSQL
**Entradas de migracion actuales**: DBF (libre), DBC (Visual FoxPro database container)

## Historial: el DSL TMG (descartado)

### Que era

Un DSL propio (archivos `.tmg`) con sintaxis tipo YAML/indentacion para definir esquemas de tablas de forma declarativa.
Incluia un parser completo: Scanner, Parser, Token, Node (~1400 lineas en `migrafox.prg`).

### Por que se descarto

1. **Verbosidad excesiva**: 26 lineas promedio por tabla vs 5-8 en SQL directo
2. **Indentacion fragil**: el parser dependia de columnas exactas (`nCol`), un espacio mal puesto rompia todo
3. **Duplicacion**: FoxRemote ya migraba desde DBF/DBC sin necesidad de un DSL intermedio
4. **Redundancia**: `type: int` + `size: 11` (el size era irrelevante para INT)
5. **Costo de mantenimiento**: ~1400 lineas de parser para una feature que no aportaba valor sobre DBF/DBC

### Ejemplo de verbosidad TMG

```tmg
-table:
    name: usuarios
    fields:
        -name: id
         type: int
         autoIncrement: true
         primaryKey: true
        -name: nombre
         type: varchar
         size: 30
         default: 'ADMIN'
         index: true
```

17 lineas para 1 tabla con 2 campos. SQL directo: 4 lineas.

---

## Propuesta: Parser DBML (pendiente)

### Idea

Adoptar **DBML** (Database Markup Language — el formato de dbdiagram.io) como DSL declarativo para definir esquemas.
No reinventar la rueda: DBML ya resuelve exactamente el problema de verbosidad que hizo descartar TMG.

### Por que DBML

| Problema TMG | Solucion DBML |
|---|---|
| Indentacion fragil | Llaves `{}` como bloques |
| Cada atributo en su linea | Atributos inline `[pk, increment]` |
| `type: int` + `size: 11` | Solo `int` |
| FK como bloque anidado | `Ref:` inline |
| 26 lineas/tabla | 5-8 lineas/tabla |

### Ejemplo DBML

```dbml
Table Clientes {
  id int [pk, increment]
  nombre varchar(100) [not null, index]
  direccion varchar(200)
}

Table Productos {
  id int [pk, increment]
  nombre varchar(100) [not null, index, unique]
  precio_unitario decimal(8,2)
  categoria_id int [ref: > Categorias.id]
}

Table Facturas {
  id int [pk, increment]
  cliente_id int [ref: > Clientes.id]
  fecha date

  indexes {
    (cliente_id, fecha desc) [unique]
  }
}

Table DetallesFactura {
  id int [pk, increment]
  factura_id int [ref: > Facturas.id, delete: cascade]
  producto_id int [ref: > Productos.id]
  cantidad int
  precio_unitario decimal(8,2)

  indexes {
    (factura_id, producto_id) [unique]
  }
}

Ref: Productos.categoria_id > Categorias.id
Ref: Facturas.cliente_id > Clientes.id
Ref: DetallesFactura.factura_id > Facturas.id [delete: cascade]
Ref: DetallesFactura.producto_id > Productos.id
```

~35 lineas para 4 tablas con FK e indices compuestos. En TMG: ~120 lineas.

### Arquitectura propuesta

```
archivo.dbml → Parser DBML → loTables (Collection) → FoxRemote.createTable() → SQL del motor
```

El metodo `migrate()` ya acepta un objeto `loTables` con:
- `cTableName`
- `cTableDescription`
- `aTableFields[23]` (mismo array que usa AFIELDS)
- `oComposedIndexes` (Collection)

**No se necesita modificar `createTable()` ni ninguna clase de motor.** Solo agregar el parser DBML que genere esa estructura.

### Componentes a implementar

1. **`foxremote_dbml.prg`** — Archivo separado, no mezclar en el core
   - `defineConstants()` — Tokens del parser
   - `Class DBMLScanner` — Tokenizador (~150 lineas)
   - `Class DBMLParser` — Parser recursivo (~200 lineas)
   - `Function parseDBML(tcFilePath)` — API publica
   - `Function dbmlToTables(tcFilePath)` — Genera la coleccion `loTables` compatible con `migrate()`

2. **Mapeo de tipos DBML → VFP letters**
   - `int` → `I`
   - `varchar(n)` → `V` (size = n)
   - `decimal(p,s)` → `N` (size = p, decimal = s)
   - `float` / `double` → `B`
   - `bool` → `L`
   - `date` → `D`
   - `datetime` / `timestamp` → `T`
   - `text` → `M`
   - `blob` / `bytea` → `W`

3. **Mapeo de atributos DBML → laFields**
   - `[pk]` → `laFields[i, 20] = .T.`
   - `[increment]` → `laFields[i, 22] = .T.`
   - `[not null]` → `laFields[i, 5] = .F.`
   - `[default: 'val']` → `laFields[i, 19] = 'val'`
   - `[index]` → `laFields[i, 23] = loIdxMeta`
   - `[unique]` → `loIdxMeta.bUnique = .T.`
   - `[ref: > Table.field]` → `laFields[i, 21] = loFkMeta`
   - `[delete: cascade]` → `loFkMeta.cOnDelete = 'CASCADE'`
   - `[update: restrict]` → `loFkMeta.cOnUpdate = 'RESTRICT'`

4. **Integracion con `migrate()`**
   - Restaurar soporte para extension `.dbml` en `migrate()`
   - Agregar `Case Upper(JustExt(tcTableOrPath)) == "DBML"` que llame a `dbmlToTables()`
   - Corregir bug residual: mensaje de error linea 145 menciona TMG pero no se soporta

### Uso esperado

```foxpro
SET PROCEDURE TO foxremote.prg ADDITIVE
SET PROCEDURE TO foxremote_dbml.prg ADDITIVE

loDB = CREATEOBJECT("MySQL")
loDB.cDriver = "MySQL ODBC 8.0 ANSI Driver"
loDB.cServer = "localhost"
loDB.cDatabase = "mi_app"
loDB.cUser = "root"
loDB.cPassword = "secret"
loDB.Connect()

* Migrar desde DBML (nuevo)
loDB.migrate("schema.dbml")

* Migrar desde DBF (existente, sin cambios)
loDB.migrate("clientes.dbf")

* Migrar desde DBC (existente, sin cambios)
loDB.migrate("midatabase.dbc")
```

### Ventajas clave

- **DBML tiene ecosistema**: dbdiagram.io para visualizacion, plugin VS Code, validadores
- **Sintaxis minimalista**: fue disenada para resolver exactamente la verbosidad que hizo descartar TMG
- **Parser liviano**: ~350 lineas vs ~1400 del TMG original
- **Zero cambios en el core**: FoxRemote no se toca, solo se agrega un archivo nuevo
- **Retrocompatible**: DBF y DBC siguen funcionando igual

### Notas de implementacion

- El parser debe ser tolerante: comentarios `//` y `/* */`, espacios flexibles
- Usar llaves `{}` en vez de indentacion para evitar el problema fragil del TMG
- Los tipos `varchar(n)` se parsean con parentesis, no como atributo separado
- Las `Ref:` globales (fuera de tabla) generan FK en `laFields[i, 21]`
- Los `indexes {}` dentro de tabla generan `oComposedIndexes`
- Estimar esfuerzo: 2-3 sesiones de trabajo
