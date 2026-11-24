Attribute VB_Name = "migrationInit"
'======================================================================
' Module : migrationInit   (the initiator)
' Version: {{VERSION}} - {{NAME}}
' Purpose: Entry point of the database migration. Open the Immediate
'          window (Ctrl+G) and type:   RunMigration
'
'          1. makes sure the [migrations] table exists
'          2. refuses a database that is NEWER than this migration
'          3. creates the tables, one by one       -> migrationTables
'             (missing columns are added to existing tables)
'          4. creates / updates the saved queries  -> migrationQueries
'          5. inserts MIGRATION_VERSION into [migrations]
'
' Notes  : Each version folder holds the WHOLE schema, so this brings an
'          empty database or any older version up to MIGRATION_VERSION.
'          Running it again at the same version repairs the schema
'          (adds what is missing, refreshes the queries).
'          The version row is written LAST, so a run that fails halfway
'          is never marked as applied: fix the error and run again.
'          DDL goes through ADO (CurrentProject.Connection) because only
'          ANSI-92 mode accepts CHECK constraints and DEFAULT values.
'          Saved queries go through DAO (CurrentDb.CreateQueryDef).
'          Progress is printed to the Immediate window.
' Source : templates\migrationInit.bas - generate_migration_vba.ps1
'          fills in the version. Edit the template, not this copy.
'======================================================================
Option Compare Database
Option Explicit

Public Const MIGRATION_VERSION As Long = {{VERSION}}
Public Const MIGRATION_NAME As String = "{{NAME}}"

Private Const adCmdText As Long = 1
Private Const adExecuteNoRecords As Long = 128

Private mCurrentObject As String
Private mCurrentSql As String

Private mCurrentStep As String
Private mLastError As String

'----------------------------------------------------------------------
' Run the whole migration and show the result in a message box.
'----------------------------------------------------------------------
Public Sub RunMigration()
    If ApplyMigration() Then
        MsgBox "Migration version " & MIGRATION_VERSION & " (" & MIGRATION_NAME & ") is applied.", _
               vbInformation, "Migration"
    Else
        MsgBox "Migration FAILED at " & mCurrentStep & vbCrLf & _
               mCurrentObject & vbCrLf & vbCrLf & _
               mLastError & vbCrLf & vbCrLf & _
               "See the Immediate window (Ctrl+G) for details. " & _
               "Fix the problem and run RunMigration again.", _
               vbCritical, "Migration"
    End If
End Sub

'----------------------------------------------------------------------
' Run the whole migration without any message box (e.g. from startup
' code or automation). True = the version is applied (now or before).
' The failure is printed to the Immediate window and LastMigrationError.
'----------------------------------------------------------------------
Public Function ApplyMigration() As Boolean
    Dim databaseVersion As Long

    On Error GoTo Fail

    mCurrentStep = ""
    mCurrentObject = ""
    mCurrentSql = ""
    mLastError = ""

    MigrationLog String(70, "=")
    MigrationLog "Migration " & MIGRATION_VERSION & " - " & MIGRATION_NAME & "   (" & Now() & ")"
    MigrationLog String(70, "=")

    BeginStep "1. migration version table"
    migrationTables.Table_00_migration

    databaseVersion = CurrentMigrationVersion()
    MigrationLog "  database is at version " & databaseVersion

    If databaseVersion > MIGRATION_VERSION Then
        mCurrentObject = "version check"
        Err.Raise vbObjectError + 513, "migrationInit", _
            "The database is at version " & databaseVersion & _
            ", newer than this migration (version " & MIGRATION_VERSION & "). " & _
            "Use the newest migration files."
    ElseIf databaseVersion = MIGRATION_VERSION Then
        MigrationLog "  already at version " & MIGRATION_VERSION & _
                     ": checking the schema (adds what is missing, refreshes queries)"
    End If

    BeginStep "2. tables"
    migrationTables.ApplyTables

    BeginStep "3. queries"
    migrationQueries.ApplyQueries

    BeginStep "4. insert migration version"
    If IsVersionApplied(MIGRATION_VERSION) Then
        MigrationLog "  skipped   migrations version " & MIGRATION_VERSION & " (recorded)"
    Else
        InsertMigrationVersion MIGRATION_VERSION, MIGRATION_NAME
    End If

    Application.RefreshDatabaseWindow

    MigrationLog "DONE - database is at version " & MIGRATION_VERSION & "."
    ApplyMigration = True
    Exit Function

Fail:
    mLastError = Err.Number & " - " & Err.Description
    MigrationLog "FAILED at step : " & mCurrentStep
    MigrationLog "       object  : " & mCurrentObject
    MigrationLog "       error   : " & mLastError
    If Len(mCurrentSql) > 0 Then
        MigrationLog "       sql     :" & vbCrLf & mCurrentSql
    End If
    ApplyMigration = False
End Function

' "<step> / <object>: <error>" of the last failed run; "" after success.
Public Function LastMigrationError() As String
    If Len(mLastError) > 0 Then
        LastMigrationError = mCurrentStep & " / " & mCurrentObject & ": " & mLastError
    End If
End Function

Private Sub BeginStep(ByVal stepName As String)
    mCurrentStep = stepName
    MigrationLog stepName
End Sub

'----------------------------------------------------------------------
' Highest version recorded in [migrations]; 0 when none.
'----------------------------------------------------------------------
Public Function CurrentMigrationVersion() As Long
    Dim rs As Object

    If Not migrationTables.TableExists("migrations") Then
        CurrentMigrationVersion = 0
        Exit Function
    End If

    Set rs = MigrationConnection.Execute("SELECT MAX([version]) FROM [migrations]")
    If IsNull(rs.Fields(0).Value) Then
        CurrentMigrationVersion = 0
    Else
        CurrentMigrationVersion = rs.Fields(0).Value
    End If
    rs.Close
End Function

Private Function IsVersionApplied(ByVal version As Long) As Boolean
    Dim rs As Object

    Set rs = MigrationConnection.Execute( _
        "SELECT COUNT(*) FROM [migrations] WHERE [version] = " & version)
    IsVersionApplied = (rs.Fields(0).Value > 0)
    rs.Close
End Function

Private Sub InsertMigrationVersion(ByVal version As Long, ByVal migrationName As String)
    Dim sql As String

    sql = "INSERT INTO [migrations] ([version], [name], [applied_by], [applied_at]) " & _
          "VALUES (" & version & ", " & _
          SqlText(migrationName) & ", " & _
          SqlText(Environ$("USERNAME")) & ", " & _
          "Now())"

    mCurrentObject = "migrations row " & version
    mCurrentSql = sql
    MigrationConnection.Execute sql, , adCmdText Or adExecuteNoRecords
    MigrationLog "  inserted  migrations version " & version & " (" & migrationName & ")"
End Sub

'======================================================================
' Shared helpers, used by migrationTables and migrationQueries
'======================================================================

' The ADO connection of this database (ANSI-92 SQL).
Public Function MigrationConnection() As Object
    Set MigrationConnection = CurrentProject.Connection
End Function

' Run one DDL statement through ADO. objectLabel names it in the log
' and in the error report.
Public Sub ExecuteDdl(ByVal sql As String, ByVal objectLabel As String)
    SetCurrentObject objectLabel, sql
    MigrationConnection.Execute sql, , adCmdText Or adExecuteNoRecords
    MigrationLog "  created   " & objectLabel
End Sub

' Remember what is being worked on, for the error report.
Public Sub SetCurrentObject(ByVal objectLabel As String, Optional ByVal sql As String = "")
    mCurrentObject = objectLabel
    mCurrentSql = sql
End Sub

Public Sub MigrationLog(ByVal message As String)
    Debug.Print message
End Sub

Private Function SqlText(ByVal value As String) As String
    SqlText = "'" & Replace(value, "'", "''") & "'"
End Function
