<#
.SYNOPSIS
  Prepares the local database servers for FoxRemote's integration tests.

.DESCRIPTION
  For every server engine configured in probe\engines.local.ini (SQL Server, MariaDB, MySQL,
  PostgreSQL, Firebird) it creates, once, a test login "fr_test" whose password has no ';'
  and a shared database "foxremote_test" that login owns. The passwords are generated here
  and written to the .ini (test_user / test_password); they are never printed.
  SQLite needs no server: the test database is a file in data_dir.

  Idempotent: an engine whose section already has test_password is left alone.
  Needs the services running and the admin credentials already in the .ini
  (Windows authentication for SQL Server, root_password, super_password, sysdba_password).
#>
[CmdletBinding()]
param(
    [string]$Ini = (Join-Path $PSScriptRoot '..\probe\engines.local.ini'),
    # Inside the repo (git-ignored): a folder made from a sandboxed shell under the user profile
    # can be invisible to a database service running as LocalSystem (measured with Firebird).
    [string]$DataDir = (Join-Path $PSScriptRoot '.data')
)
$ErrorActionPreference = 'Stop'
$Ini = (Resolve-Path $Ini).Path
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('frsetup-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $scratch | Out-Null
New-Item -ItemType Directory -Force $DataDir | Out-Null

function Read-IniLines { ([IO.File]::ReadAllText($Ini) -split "`r?`n") | Where-Object { $_ -ne '' } }
function Get-IniValue([string]$Section, [string]$Key) {
    $in = $false
    foreach ($l in Read-IniLines) {
        if ($l -match '^\[(.+)\]$') { $in = ($Matches[1] -eq $Section); continue }
        if ($in -and $l.StartsWith("$Key=")) { return $l.Substring($Key.Length + 1) }
    }
    return $null
}
function Set-IniValues([string]$Section, [hashtable]$Values) {
    $lines = [Collections.Generic.List[string]]::new()
    $lines.AddRange([string[]](Read-IniLines))
    $start = $lines.IndexOf("[$Section]")
    if ($start -lt 0) { $lines.Add("[$Section]"); $start = $lines.Count - 1 }
    $end = $start + 1
    while ($end -lt $lines.Count -and -not $lines[$end].StartsWith('[')) { $end++ }
    foreach ($k in $Values.Keys) {
        $found = $false
        for ($i = $start + 1; $i -lt $end; $i++) {
            if ($lines[$i].StartsWith("$k=")) { $lines[$i] = "$k=$($Values[$k])"; $found = $true }
        }
        if (-not $found) { $lines.Insert($end, "$k=$($Values[$k])"); $end++ }
    }
    [IO.File]::WriteAllText($Ini, (($lines -join "`r`n") + "`r`n"), [Text.Encoding]::ASCII)
}
function New-Password {
    $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789'.ToCharArray()
    $b = New-Object byte[] 16
    [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($b)
    'Ft' + (-join ($b | ForEach-Object { $chars[$_ % $chars.Length] })) + '7'
}
function Write-Secret([string]$Name, [string]$Text) {
    $p = Join-Path $scratch $Name
    [IO.File]::WriteAllText($p, $Text, [Text.Encoding]::ASCII)
    return $p
}

try {
    Set-IniValues 'tests' @{ data_dir = $DataDir }

    # --- SQLite: just the file ------------------------------------------------
    $sqliteDb = Join-Path $DataDir 'foxremote_test.db'
    if (-not (Test-Path $sqliteDb)) {
        & 'C:\Programas\SQLiteODBC\sqlite3.exe' $sqliteDb 'VACUUM;'
    }
    Set-IniValues 'sqlite' @{ driver = 'SQLite3 ODBC Driver'; database = $sqliteDb }
    Write-Host "sqlite: $sqliteDb"

    # --- SQL Server (Windows authentication as administrator) -----------------
    if (-not (Get-IniValue 'mssql' 'test_password')) {
        $pw = New-Password
        $server = Get-IniValue 'mssql' 'server'
        $sql = @"
IF SUSER_ID('fr_test') IS NULL CREATE LOGIN fr_test WITH PASSWORD = '$pw', CHECK_POLICY = OFF;
ALTER SERVER ROLE dbcreator ADD MEMBER fr_test;
IF DB_ID('foxremote_test') IS NULL CREATE DATABASE foxremote_test;
GO
USE foxremote_test;
IF USER_ID('fr_test') IS NULL CREATE USER fr_test FOR LOGIN fr_test;
ALTER ROLE db_owner ADD MEMBER fr_test;
GO
"@
        $f = Write-Secret 'mssql.sql' $sql
        & sqlcmd -S $server -E -b -i $f | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "mssql: sqlcmd exit $LASTEXITCODE" }
        Set-IniValues 'mssql' @{ test_user = 'fr_test'; test_password = $pw; test_database = 'foxremote_test' }
    }
    Write-Host 'mssql: fr_test / foxremote_test'

    # --- MariaDB (root, whose password is in the .ini) ------------------------
    if (-not (Get-IniValue 'mariadb' 'test_password')) {
        $pw = New-Password
        $root = Get-IniValue 'mariadb' 'root_password'
        $cnf = Write-Secret 'my.cnf' "[client]`r`nuser=root`r`npassword=`"$root`"`r`nhost=127.0.0.1`r`n"
        $sql = @"
CREATE USER IF NOT EXISTS 'fr_test'@'localhost' IDENTIFIED BY '$pw';
CREATE USER IF NOT EXISTS 'fr_test'@'127.0.0.1' IDENTIFIED BY '$pw';
GRANT ALL PRIVILEGES ON *.* TO 'fr_test'@'localhost';
GRANT ALL PRIVILEGES ON *.* TO 'fr_test'@'127.0.0.1';
CREATE DATABASE IF NOT EXISTS foxremote_test DEFAULT CHARACTER SET utf8mb4;
"@
        $f = Write-Secret 'mariadb.sql' $sql
        Get-Content $f -Raw | & 'C:\Programas\MariaDB\bin\mariadb.exe' "--defaults-extra-file=$cnf"
        if ($LASTEXITCODE -ne 0) { throw "mariadb: client exit $LASTEXITCODE" }
        Set-IniValues 'mariadb' @{ test_user = 'fr_test'; test_password = $pw; test_database = 'foxremote_test' }
    }
    Write-Host 'mariadb: fr_test / foxremote_test'

    # --- MySQL (root, whose password is in the .ini) --------------------------
    # Besides fr_test it creates fr_probe, whose password HAS a ';' (the F10 test).
    if ((Get-IniValue 'mysql' 'root_password') -and -not (Get-IniValue 'mysql' 'test_password')) {
        $pw = New-Password
        $semi = (New-Password) + ';' + (New-Password)
        $root = Get-IniValue 'mysql' 'root_password'
        $port = Get-IniValue 'mysql' 'port'
        $cnf = Write-Secret 'mysql.cnf' "[client]`r`nuser=root`r`npassword=`"$root`"`r`nhost=127.0.0.1`r`nport=$port`r`n"
        $sql = @"
CREATE USER IF NOT EXISTS 'fr_test'@'localhost' IDENTIFIED BY '$pw';
CREATE USER IF NOT EXISTS 'fr_test'@'127.0.0.1' IDENTIFIED BY '$pw';
GRANT ALL PRIVILEGES ON *.* TO 'fr_test'@'localhost';
GRANT ALL PRIVILEGES ON *.* TO 'fr_test'@'127.0.0.1';
CREATE USER IF NOT EXISTS 'fr_probe'@'localhost' IDENTIFIED BY '$semi';
CREATE USER IF NOT EXISTS 'fr_probe'@'127.0.0.1' IDENTIFIED BY '$semi';
GRANT ALL PRIVILEGES ON foxremote_test.* TO 'fr_probe'@'localhost';
GRANT ALL PRIVILEGES ON foxremote_test.* TO 'fr_probe'@'127.0.0.1';
CREATE DATABASE IF NOT EXISTS foxremote_test DEFAULT CHARACTER SET utf8mb4;
"@
        $f = Write-Secret 'mysql.sql' $sql
        Get-Content $f -Raw | & 'C:\Programas\MySQL\bin\mysql.exe' "--defaults-extra-file=$cnf"
        if ($LASTEXITCODE -ne 0) { throw "mysql: client exit $LASTEXITCODE" }
        Set-IniValues 'mysql' @{ test_user = 'fr_test'; test_password = $pw; test_database = 'foxremote_test'; user = 'fr_probe'; password = $semi }
    }
    if (Get-IniValue 'mysql' 'root_password') { Write-Host 'mysql: fr_test / foxremote_test' }

    # --- PostgreSQL (postgres, whose password is in the .ini) -----------------
    if (-not (Get-IniValue 'postgresql' 'test_password')) {
        $pw = New-Password
        $env:PGPASSWORD = Get-IniValue 'postgresql' 'super_password'
        $psql = 'C:\Programas\PostgreSQL\18\bin\psql.exe'
        $f = Write-Secret 'pg1.sql' "CREATE ROLE fr_test LOGIN CREATEDB PASSWORD '$pw';`r`n"
        & $psql -h 127.0.0.1 -U postgres -d postgres -v ON_ERROR_STOP=1 -q -f $f
        if ($LASTEXITCODE -ne 0) { throw "postgresql: psql exit $LASTEXITCODE" }
        $f = Write-Secret 'pg2.sql' "CREATE DATABASE foxremote_test OWNER fr_test ENCODING 'UTF8';`r`n"
        & $psql -h 127.0.0.1 -U postgres -d postgres -v ON_ERROR_STOP=1 -q -f $f
        if ($LASTEXITCODE -ne 0) { throw "postgresql: psql exit $LASTEXITCODE" }
        $env:PGPASSWORD = ''
        Set-IniValues 'postgresql' @{ test_user = 'fr_test'; test_password = $pw; test_database = 'foxremote_test' }
    }
    Write-Host 'postgresql: fr_test / foxremote_test'

    # --- Firebird (SYSDBA, whose password is in the .ini) ---------------------
    if (-not (Get-IniValue 'firebird' 'test_password')) {
        $pw = New-Password
        $isql = 'C:\Programas\Firebird\5.0\isql.exe'
        $fdb = Join-Path $DataDir 'foxremote_test.fdb'
        $env:ISC_USER = 'SYSDBA'; $env:ISC_PASSWORD = Get-IniValue 'firebird' 'sysdba_password'
        $f = Write-Secret 'fb1.sql' "CREATE OR ALTER USER FR_TEST PASSWORD '$pw';`r`nGRANT CREATE DATABASE TO USER FR_TEST;`r`nCOMMIT;`r`n"
        & $isql -q -i $f '127.0.0.1:employee'
        if ($LASTEXITCODE -ne 0) { throw "firebird: isql exit $LASTEXITCODE" }
        $env:ISC_USER = ''; $env:ISC_PASSWORD = ''
        if (-not (Test-Path $fdb)) {
            $f = Write-Secret 'fb2.sql' "CREATE DATABASE '127.0.0.1/3050:$fdb' USER 'FR_TEST' PASSWORD '$pw' PAGE_SIZE 8192 DEFAULT CHARACTER SET WIN1252;`r`nCOMMIT;`r`n"
            & $isql -q -i $f
            if ($LASTEXITCODE -ne 0) { throw "firebird: isql exit $LASTEXITCODE" }
        }
        Set-IniValues 'firebird' @{ test_user = 'FR_TEST'; test_password = $pw; test_database = $fdb }
    }
    Write-Host 'firebird: FR_TEST / foxremote_test.fdb'
}
finally {
    $env:PGPASSWORD = ''; $env:ISC_USER = ''; $env:ISC_PASSWORD = ''
    Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
}
