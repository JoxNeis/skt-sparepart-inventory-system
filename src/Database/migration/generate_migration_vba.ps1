<#
.SYNOPSIS
    Builds one shareable migration version folder, version\ver_NN, from
    the SQL files in ..\tables and ..\queries.

.DESCRIPTION
    The .sql files stay the source of truth. The folder it writes holds
    everything a user needs:

        version\ver_NN\
            migrationInit.bas    - the initiator (from templates\migrationInit.bas)
            migrationTables.bas  - one Public Sub per table file (Table_<file>),
                                   run in file order by ApplyTables
            migrationQueries.bas - one Public Sub per query file (Query_<file>),
                                   run by ApplyQueries in dependency order
            README.txt           - how to import and run it

    Every version folder holds the WHOLE schema, so a user only needs the
    newest one. Running the script again for the same version rebuilds
    that folder.

.PARAMETER Version
    Migration number (1, 2, 3...). Use a new number for every schema
    change that is shared with users.

.PARAMETER Name
    Short snake_case description, e.g. initial_schema or
    add_plant_region. Stored in the [migrations] table.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File generate_migration_vba.ps1 -Version 1 -Name initial_schema
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 99)]
    [int]$Version,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[a-z0-9_]{1,100}$')]
    [string]$Name
)

$ErrorActionPreference = 'Stop'

$here         = $PSScriptRoot
$tablesDir    = Join-Path $here '..\tables'
$queriesDir   = Join-Path $here '..\queries'
$templatePath = Join-Path $here 'templates\migrationInit.bas'
$versionTag   = 'ver_{0:D2}' -f $Version
$outDir       = Join-Path $here "version\$versionTag"
$ansi         = [System.Text.Encoding]::GetEncoding(1252)
$nl           = "`r`n"

if (-not (Test-Path $outDir)) { [void](New-Item -ItemType Directory -Path $outDir) }

# Calculated fields: Access DDL cannot create them, so they are added
# through DAO right after their table (see 05_spareparts.sql).
$CalculatedFields = @{
    'spareparts' = @(
        @{ Name = 'display'; Type = 'DB_TEXT'; Expression = "[code] & ' - ' & [name]" }
    )
}

# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

# Split a .sql file into its comment lines and its SQL body lines.
function Read-SqlFile([string]$path) {
    $comment = New-Object System.Collections.Generic.List[string]
    $body    = New-Object System.Collections.Generic.List[string]

    foreach ($line in [System.IO.File]::ReadAllLines($path)) {
        $trimmed = $line.Trim()
        if ($trimmed.StartsWith('--')) {
            $text = $trimmed.Substring(2)
            if ($text.StartsWith(' ')) { $text = $text.Substring(1) }
            if ($text -notmatch '^=+$') { $comment.Add($text.TrimEnd()) }
        }
        else {
            $body.Add($line.Replace("`t", '    ').TrimEnd())
        }
    }

    [pscustomobject]@{
        Name    = [System.IO.Path]::GetFileNameWithoutExtension($path)
        Comment = $comment.ToArray()
        Body    = $body.ToArray()
    }
}

# Remove blank lines at the start and end of a block of lines.
function Get-TrimmedLines([string[]]$lines) {
    $first = 0
    $last  = $lines.Count - 1
    while ($first -le $last -and $lines[$first] -eq '') { $first++ }
    while ($last -ge $first -and $lines[$last] -eq '')  { $last-- }
    if ($first -gt $last) { return ,@() }
    return ,$lines[$first..$last]
}

# Split a body into statements on ';' (the SQL files hold no ';' in strings).
# Returns a List of string[] (one entry per statement).
function Split-Statements([string[]]$body) {
    $statements = New-Object 'System.Collections.Generic.List[string[]]'
    $text = ($body -join "`n")
    foreach ($part in $text.Split(';')) {
        $lines = Get-TrimmedLines ($part.Split("`n"))
        if ($lines.Count -gt 0) { $statements.Add([string[]]$lines) }
    }
    return ,$statements
}

# VBA lines that build the variable 'sql' from SQL lines.
function ConvertTo-VbaSql([string[]]$lines) {
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add('    sql = ""')
    foreach ($line in $lines) {
        if ($line -eq '') {
            $out.Add('    sql = sql & vbCrLf')
        }
        else {
            $out.Add('    sql = sql & "' + $line.Replace('"', '""') + '" & vbCrLf')
        }
    }
    return ,$out.ToArray()
}

# VBA comment block (the header comment of the .sql file).
function ConvertTo-VbaComment([string]$fileName, [string[]]$comment) {
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add("'" + ('-' * 70))
    $out.Add("' $fileName")
    foreach ($c in $comment) { $out.Add(("' " + $c).TrimEnd()) }
    $out.Add("'" + ('-' * 70))
    return ,$out.ToArray()
}

function Write-Module([string]$path, [string[]]$lines) {
    [System.IO.File]::WriteAllText($path, ($lines -join $nl) + $nl, $ansi)
    Write-Host "written $path"
}

$generatedNote = @(
    "' Version: $Version - $Name"
    "' GENERATED by generate_migration_vba.ps1 from the .sql files."
    "' Do not edit by hand: change the .sql file and run the script again."
)

# ---------------------------------------------------------------------
# migrationTables.bas
# ---------------------------------------------------------------------
$tableFiles = Get-ChildItem -Path $tablesDir -Filter '*.sql' | Sort-Object Name
$tableProcs = New-Object System.Collections.Generic.List[string]
$tableCalls = New-Object System.Collections.Generic.List[string]

foreach ($file in $tableFiles) {
    $sqlFile  = Read-SqlFile $file.FullName
    $procName = 'Table_' + $sqlFile.Name

    $tableProcs.Add('')
    $tableProcs.AddRange([string[]](ConvertTo-VbaComment $file.Name $sqlFile.Comment))
    $tableProcs.Add("Public Sub $procName()")
    $tableProcs.Add('    Dim sql As String')

    $createdTables = New-Object System.Collections.Generic.List[string]

    $statements = Split-Statements $sqlFile.Body
    foreach ($statement in $statements) {
        $flat = ($statement -join ' ')
        $tableProcs.Add('')
        $tableProcs.AddRange([string[]](ConvertTo-VbaSql $statement))

        if ($flat -match '^\s*CREATE\s+TABLE\s+\[(\w+)\]') {
            $tableName = $Matches[1]
            $createdTables.Add($tableName)
            $tableProcs.Add("    EnsureTable `"$tableName`", sql")

            # Every column again, so a table that already exists (from an
            # older migration) gets the columns added since then.
            $tableProcs.Add('')
            $tableProcs.Add("    ' Columns added later are created on an existing table.")
            foreach ($line in $statement) {
                if ($line -match '^\s*CONSTRAINT\b') { break }
                if ($line -match '^\s*\[(\w+)\]\s+(.+?),?\s*$') {
                    $columnName = $Matches[1]
                    $definition = ($Matches[2] -replace '\s+', ' ').Trim()
                    if ($definition -match '^AUTOINCREMENT\b') { continue }
                    $definition = $definition.Replace('"', '""')
                    $tableProcs.Add("    EnsureColumn `"$tableName`", `"$columnName`", `"$definition`"")
                }
            }
        }
        elseif ($flat -match '^\s*CREATE\s+(UNIQUE\s+)?INDEX\s+\[(\w+)\]\s+ON\s+\[(\w+)\]') {
            $indexName = $Matches[2]
            $tableName = $Matches[3]
            $tableProcs.Add("    EnsureIndex `"$tableName`", `"$indexName`", sql")
        }
        else {
            throw "$($file.Name): unsupported statement: $flat"
        }
    }

    foreach ($tableName in $createdTables) {
        if ($CalculatedFields.ContainsKey($tableName)) {
            foreach ($cf in $CalculatedFields[$tableName]) {
                $expr = $cf.Expression.Replace('"', '""')
                $tableProcs.Add('')
                $tableProcs.Add("    ' Calculated field (Access DDL cannot create one).")
                $tableProcs.Add("    EnsureCalculatedField `"$tableName`", `"$($cf.Name)`", $($cf.Type), `"$expr`"")
            }
        }
    }

    $tableProcs.Add('End Sub')
    $tableCalls.Add("    $procName")
}

$tablesModule = @(
    'Attribute VB_Name = "migrationTables"'
    "'======================================================================"
    "' Module : migrationTables   (the table injector)"
    "' Purpose: Creates every table of the schema, one procedure per table"
    "'          file in src\Database\tables. A table, index or calculated"
    "'          field that already exists is skipped, so this is safe to run"
    "'          again after a failure. A column that is in the .sql file but"
    "'          missing from an existing table is added (ALTER TABLE ADD"
    "'          COLUMN); columns are never changed or dropped."
    "'          Started by migrationInit.RunMigration."
    $generatedNote
    "'======================================================================"
    'Option Compare Database'
    'Option Explicit'
    ''
    "' Late bound (no reference to the DAO or ADO library needed)."
    'Private Const DB_TEXT As Long = 10            '' dbText'
    'Private Const DB_REFRESH_CACHE As Long = 8    '' dbRefreshCache'
    'Private Const adSchemaColumns As Long = 4'
    'Private Const adSchemaIndexes As Long = 12'
    'Private Const adSchemaTables As Long = 20'
    ''
    "'----------------------------------------------------------------------"
    "' Create all tables, in dependency (file) order."
    "'----------------------------------------------------------------------"
    'Public Sub ApplyTables()'
    $tableCalls
    'End Sub'
    $tableProcs
    ''
    "'======================================================================"
    "' Helpers"
    "'======================================================================"
    ''
    'Private Sub EnsureTable(ByVal tableName As String, ByVal sql As String)'
    '    If TableExists(tableName) Then'
    '        migrationInit.MigrationLog "  skipped   table " & tableName & " (exists)"'
    '    Else'
    '        migrationInit.ExecuteDdl sql, "table " & tableName'
    '    End If'
    'End Sub'
    ''
    "' Adds a column to an existing table when it is missing (silent when it exists)."
    'Private Sub EnsureColumn(ByVal tableName As String, ByVal columnName As String, ByVal definition As String)'
    '    If FieldExists(tableName, columnName) Then Exit Sub'
    '    migrationInit.ExecuteDdl _'
    '        "ALTER TABLE [" & tableName & "] ADD COLUMN [" & columnName & "] " & definition, _'
    '        "column " & tableName & "." & columnName'
    'End Sub'
    ''
    'Private Sub EnsureIndex(ByVal tableName As String, ByVal indexName As String, ByVal sql As String)'
    '    If IndexExists(tableName, indexName) Then'
    '        migrationInit.MigrationLog "  skipped   index " & indexName & " (exists)"'
    '    Else'
    '        migrationInit.ExecuteDdl sql, "index " & indexName & " on " & tableName'
    '    End If'
    'End Sub'
    ''
    "' Adds a calculated field through DAO (Field2.Expression)."
    'Private Sub EnsureCalculatedField(ByVal tableName As String, ByVal fieldName As String, _'
    '                                  ByVal fieldType As Long, ByVal expression As String)'
    '    Dim db As Object       '' DAO.Database'
    '    Dim tdf As Object      '' DAO.TableDef'
    '    Dim fld As Object      '' DAO.Field2'
    ''
    '    If FieldExists(tableName, fieldName) Then'
    '        migrationInit.MigrationLog "  skipped   field " & tableName & "." & fieldName & " (exists)"'
    '        Exit Sub'
    '    End If'
    ''
    '    migrationInit.SetCurrentObject "calculated field " & tableName & "." & fieldName, expression'
    ''
    "    ' The table was created through ADO; make DAO see it."
    '    DBEngine.Idle DB_REFRESH_CACHE'
    '    Set db = CurrentDb'
    '    db.TableDefs.Refresh'
    '    Set tdf = db.TableDefs(tableName)'
    ''
    '    If fieldType = DB_TEXT Then'
    '        Set fld = tdf.CreateField(fieldName, fieldType, 255)'
    '    Else'
    '        Set fld = tdf.CreateField(fieldName, fieldType)'
    '    End If'
    '    fld.expression = expression'
    '    tdf.Fields.Append fld'
    ''
    '    migrationInit.MigrationLog "  created   calculated field " & tableName & "." & fieldName'
    'End Sub'
    ''
    'Public Function TableExists(ByVal tableName As String) As Boolean'
    '    Dim rs As Object'
    '    Set rs = migrationInit.MigrationConnection.OpenSchema( _'
    '        adSchemaTables, Array(Empty, Empty, tableName, "TABLE"))'
    '    TableExists = Not rs.EOF'
    '    rs.Close'
    'End Function'
    ''
    'Public Function IndexExists(ByVal tableName As String, ByVal indexName As String) As Boolean'
    '    Dim rs As Object'
    '    Set rs = migrationInit.MigrationConnection.OpenSchema( _'
    '        adSchemaIndexes, Array(Empty, Empty, indexName, Empty, tableName))'
    '    IndexExists = Not rs.EOF'
    '    rs.Close'
    'End Function'
    ''
    'Public Function FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean'
    '    Dim rs As Object'
    '    Set rs = migrationInit.MigrationConnection.OpenSchema( _'
    '        adSchemaColumns, Array(Empty, Empty, tableName, fieldName))'
    '    FieldExists = Not rs.EOF'
    '    rs.Close'
    'End Function'
)

Write-Module (Join-Path $outDir 'migrationTables.bas') $tablesModule

# ---------------------------------------------------------------------
# migrationQueries.bas
# ---------------------------------------------------------------------
$queryFiles = Get-ChildItem -Path $queriesDir -Filter '*.sql'
$queries    = @{}
foreach ($file in $queryFiles) {
    $sqlFile = Read-SqlFile $file.FullName
    $queries[$sqlFile.Name] = [pscustomobject]@{
        File    = $file.Name
        Name    = $sqlFile.Name
        Comment = $sqlFile.Comment
        Body    = Get-TrimmedLines $sqlFile.Body
        Deps    = @()
    }
}

# Dependencies: other saved queries a query reads from.
foreach ($q in $queries.Values) {
    $text = ($q.Body -join "`n")
    $q.Deps = @($queries.Keys | Where-Object {
        $_ -ne $q.Name -and $text -match ('(?<![\w])' + [regex]::Escape($_) + '(?![\w])')
    } | Sort-Object)
}

# Order: views (vw_*) first, then the rest; every query after its deps.
$ordered = New-Object System.Collections.Generic.List[string]
$state   = @{}
function Visit([string]$queryName, [string[]]$path) {
    if ($state[$queryName] -eq 'done') { return }
    if ($state[$queryName] -eq 'visiting') {
        throw "Circular query dependency: $(($path + $queryName) -join ' -> ')"
    }
    $state[$queryName] = 'visiting'
    foreach ($dep in $queries[$queryName].Deps) { Visit $dep ($path + $queryName) }
    $state[$queryName] = 'done'
    $ordered.Add($queryName)
}
$names = $queries.Keys | Sort-Object @{ Expression = { -not $_.StartsWith('vw_') } }, @{ Expression = { $_ } }
foreach ($queryName in $names) { Visit $queryName @() }

$queryProcs = New-Object System.Collections.Generic.List[string]
$queryCalls = New-Object System.Collections.Generic.List[string]

foreach ($queryName in $ordered) {
    $q        = $queries[$queryName]
    $procName = 'Query_' + $q.Name

    $queryProcs.Add('')
    $queryProcs.AddRange([string[]](ConvertTo-VbaComment $q.File $q.Comment))
    if ($q.Deps.Count -gt 0) {
        $queryProcs.Add("' Uses    : " + ($q.Deps -join ', '))
    }
    $queryProcs.Add("Public Sub $procName()")
    $queryProcs.Add('    Dim sql As String')
    $queryProcs.Add('')
    $queryProcs.AddRange([string[]](ConvertTo-VbaSql $q.Body))
    $queryProcs.Add('')
    $queryProcs.Add("    SaveQuery `"$($q.Name)`", sql")
    $queryProcs.Add('End Sub')

    $queryCalls.Add("    $procName")
}

$queriesModule = @(
    'Attribute VB_Name = "migrationQueries"'
    "'======================================================================"
    "' Module : migrationQueries   (the queries injector)"
    "' Purpose: Creates every saved query, one procedure per query file in"
    "'          src\Database\queries. An existing query gets its SQL"
    "'          replaced, so the database always matches the .sql files."
    "'          Started by migrationInit.RunMigration, after the tables."
    $generatedNote
    "'======================================================================"
    'Option Compare Database'
    'Option Explicit'
    ''
    "' Late bound (no reference to the DAO library needed)."
    'Private Const DB_REFRESH_CACHE As Long = 8    '' dbRefreshCache'
    ''
    "'----------------------------------------------------------------------"
    "' Create / update all saved queries. A query comes after the queries"
    "' it uses (views first)."
    "'----------------------------------------------------------------------"
    'Public Sub ApplyQueries()'
    "    ' The tables were created through ADO; make DAO see them."
    '    DBEngine.Idle DB_REFRESH_CACHE'
    '    CurrentDb.TableDefs.Refresh'
    ''
    $queryCalls
    'End Sub'
    $queryProcs
    ''
    "'======================================================================"
    "' Helpers"
    "'======================================================================"
    ''
    "' Create the saved query, or replace the SQL of an existing one."
    'Private Sub SaveQuery(ByVal queryName As String, ByVal sql As String)'
    '    Dim db As Object       '' DAO.Database'
    ''
    '    migrationInit.SetCurrentObject "query " & queryName, sql'
    ''
    '    Set db = CurrentDb'
    '    If QueryExists(db, queryName) Then'
    '        db.QueryDefs(queryName).SQL = sql'
    '        migrationInit.MigrationLog "  updated   query " & queryName'
    '    Else'
    '        db.CreateQueryDef queryName, sql'
    '        migrationInit.MigrationLog "  created   query " & queryName'
    '    End If'
    'End Sub'
    ''
    'Private Function QueryExists(ByVal db As Object, ByVal queryName As String) As Boolean'
    '    Dim qdf As Object      '' DAO.QueryDef'
    '    For Each qdf In db.QueryDefs'
    '        If StrComp(qdf.Name, queryName, vbTextCompare) = 0 Then'
    '            QueryExists = True'
    '            Exit Function'
    '        End If'
    '    Next qdf'
    'End Function'
)

Write-Module (Join-Path $outDir 'migrationQueries.bas') $queriesModule

# ---------------------------------------------------------------------
# migrationInit.bas (from the template)
# ---------------------------------------------------------------------
$init = [System.IO.File]::ReadAllText($templatePath)
if ($init -notmatch '\{\{VERSION\}\}' -or $init -notmatch '\{\{NAME\}\}') {
    throw "$templatePath must contain {{VERSION}} and {{NAME}}"
}
$init = $init.Replace('{{VERSION}}', [string]$Version).Replace('{{NAME}}', $Name)
Write-Module (Join-Path $outDir 'migrationInit.bas') (($init -replace "`r?`n", "`n").TrimEnd("`n").Split("`n"))

# ---------------------------------------------------------------------
# README.txt for the person who applies the migration
# ---------------------------------------------------------------------
$readme = @(
    "SPAREPART INVENTORY SYSTEM - DATABASE MIGRATION"
    "Version $Version - $Name"
    "Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    ""
    "Files"
    "  migrationInit.bas     starts the migration (RunMigration)"
    "  migrationTables.bas   creates the tables ($($tableFiles.Count) files)"
    "  migrationQueries.bas  creates the saved queries ($($ordered.Count) files)"
    ""
    "This folder holds the whole database structure. It works on an empty"
    ".accdb and on a database that has an older migration version."
    ""
    "How to apply"
    "  1. Make a backup copy of the .accdb file."
    "  2. Open the .accdb in Microsoft Access. If a yellow SECURITY WARNING"
    "     bar appears, click Enable Content (VBA does not run otherwise)."
    "  3. Press Alt+F11 to open the VBA editor."
    "  4. If the modules migrationInit, migrationTables and migrationQueries"
    "     are already there (from an older version), right-click each one"
    "     > Remove > No (do not export)."
    "  5. File > Import File... and import the three .bas files."
    "  6. Press Ctrl+G, type   RunMigration   and press Enter."
    "  7. A message box tells whether it succeeded. Progress and any error"
    "     are printed in the Immediate window (Ctrl+G)."
    ""
    "Running it again is safe: what already exists is kept, what is"
    "missing is added. If it fails, fix the reported problem and run"
    "RunMigration again."
    ""
    "Check the version:  ? CurrentMigrationVersion   (in the Immediate window)"
)
Write-Module (Join-Path $outDir 'README.txt') $readme
