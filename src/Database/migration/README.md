# Database migration

Builds the Access database (tables + saved queries) from the SQL files in
[`../tables`](../tables) and [`../queries`](../queries), as VBA modules that
are imported into an `.accdb` and run there.

```
migration/
├─ generate_migration_vba.ps1   builds version/ver_NN from the .sql files
├─ templates/
│  └─ migrationInit.bas         the initiator (hand-written, {{VERSION}} / {{NAME}})
└─ version/
   └─ ver_01/                   one shareable folder per version
      ├─ migrationInit.bas      RunMigration: version check, then tables, queries, version row
      ├─ migrationTables.bas    one Sub per table file   (Table_01_manufacturing_plants ...)
      ├─ migrationQueries.bas   one Sub per query file   (Query_order_list ...)
      └─ README.txt             import / run steps for the person applying it
```

## Sharing a migration

Send the newest `version/ver_NN` folder. The user follows its `README.txt`:
import the three `.bas` files, then type `RunMigration` in the Immediate window.

Every version folder holds the **whole** schema, so it brings an empty
`.accdb` or any older version up to date. A user never needs to run the
older folders first.

## Making a schema change

1. Edit or add the `.sql` files in `../tables` / `../queries`.
2. Build a new version (next number, short snake_case name):

   ```
   powershell -ExecutionPolicy Bypass -File generate_migration_vba.ps1 -Version 2 -Name add_plant_region
   ```

3. Commit the new `version/ver_02` folder together with the `.sql` changes.

Running the script again with the same `-Version` rebuilds that folder. Do that only
for a version that has not been shared yet.

Do not edit the files in `version/` by hand: change the `.sql` file, or
`templates/migrationInit.bas` for the initiator, and build again.

## What a run does

| Database is at | Result |
|---|---|
| no `migrations` table / older version | creates what is missing, refreshes queries, records the version |
| the same version | repairs: adds what is missing, refreshes queries |
| a newer version | refuses, changes nothing |

- **Tables:** a table that exists is kept. A column that is in the `.sql`
  file but missing from the table is added with `ALTER TABLE ... ADD COLUMN`.
  Columns are never changed, renamed or dropped: write that by hand when needed.
- **Indexes and the calculated field** (`spareparts.display`) are created when missing.
- **Queries** are created, or their SQL is replaced.
- **The version row** is written last, so a failed run is never recorded as
  applied. Fix the reported error and run again.

DDL runs through ADO (`CurrentProject.Connection`), because only ANSI-92 mode
accepts `CHECK` constraints and `DEFAULT`. Saved queries go through DAO
(`CurrentDb.CreateQueryDef`).
