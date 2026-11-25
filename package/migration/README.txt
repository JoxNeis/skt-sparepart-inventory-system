SPAREPART INVENTORY SYSTEM - DATABASE MIGRATION
Version 1 - initial_schema
Generated 2026-11-25 08:21

Files
  migrationInit.bas     starts the migration (RunMigration)
  migrationTables.bas   creates the tables (16 files)
  migrationQueries.bas  creates the saved queries (64 files)

This folder holds the whole database structure. It works on an empty
.accdb and on a database that has an older migration version.

How to apply
  1. Make a backup copy of the .accdb file.
  2. Open the .accdb in Microsoft Access. If a yellow SECURITY WARNING
     bar appears, click Enable Content (VBA does not run otherwise).
  3. Press Alt+F11 to open the VBA editor.
  4. If the modules migrationInit, migrationTables and migrationQueries
     are already there (from an older version), right-click each one
     > Remove > No (do not export).
  5. File > Import File... and import the three .bas files.
  6. Press Ctrl+G, type   RunMigration   and press Enter.
  7. A message box tells whether it succeeded. Progress and any error
     are printed in the Immediate window (Ctrl+G).

Running it again is safe: what already exists is kept, what is
missing is added. If it fails, fix the reported problem and run
RunMigration again.

Check the version:  ? CurrentMigrationVersion   (in the Immediate window)
