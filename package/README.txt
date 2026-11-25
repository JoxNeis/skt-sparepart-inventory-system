SPAREPART INVENTORY SYSTEM - READY-TO-USE PACKAGE
Database migration version 1 (initial_schema)

Contents
  SIS.xlsm     the application (Excel). All code is inside.
  SIS.accdb    the database (Access). All tables and saved queries,
               no data.
  migration\   the migration that built SIS.accdb. You only need it
               to repair or upgrade an existing database (see the end
               of this file).

Keep SIS.xlsm and SIS.accdb in the SAME folder. The application
always opens the SIS.accdb that is next to it.


FIRST-TIME SETUP
----------------
1. Copy SIS.xlsm and SIS.accdb into one folder, for example
   Documents\Sparepart Inventory System.

2. Add the plants (only in Access; the application cannot add them).
   Open SIS.accdb, open the table  manufacturing_plants  and add one
   row per plant:
     code     short code, used at the start of every document number,
              e.g. SBY01 (cannot be changed once the plant has documents)
     name     plant name, e.g. PLANT SURABAYA
     address  optional
     type     optional, e.g. PLANT or WAREHOUSE
     region   optional, e.g. JAWA TIMUR
   Leave id, created_at, updated_at and deleted_at empty. Close Access.

3. Open SIS.xlsm and click Enable Content if Excel asks.
   (If Excel says macros are blocked: close it, right-click
   SIS.xlsm > Properties > tick Unblock > OK, then open it again.)

4. SETTINGS sheet: choose the plant this workbook is used for.

5. MD_USER sheet: add yourself first. USERNAME must be your
   Windows login name exactly. TIER decides approvals: an update
   request can only be approved by a user with a higher tier than
   the one who asked (for example staff 1, planner 2, supervisor 3).
   Then add the other users.

6. Fill the other master data: the MD_SPAREPART, MD_VENDOR and
   MD_POS_MOBILE sheets. After that INPUT_ORDER, INPUT_PENERIMAAN
   and INPUT_REQUEST are ready to use.

7. SETTINGS sheet: check the email TO / CC / BCC lists. KIRIM EMAIL
   (INPUT_ORDER) creates an .eml file and opens it in the default
   mail app. To use the new Outlook, set it as the default app for
   .eml files: Windows Settings > Apps > Default apps.


SHARING ONE DATABASE
--------------------
Several people can use the same SIS.accdb when SIS.xlsm and SIS.accdb
are together in a shared folder that has a normal drive path (for
example a network drive). Every user must be added in MD_USER.

OneDrive / SharePoint: when Excel reports the workbook location as a
web address (https://...), the application does not look next to the
workbook but uses this fixed path (set in coreConfig.databasePath):
  C:\Users\<Windows user>\OneDrive - Philip Morris International\
  Sparepart Inventory System\System\Sparepart Inventory System Ver 5\
  Sparepart Inventory System Ver 5.accdb


REPAIRING OR UPGRADING A DATABASE (migration folder)
----------------------------------------------------
The SIS.accdb in this package is already at version 1, so a new
installation does not need this step.

Use it for an EXISTING database that was made with an older version,
or one that is missing tables or queries. Follow migration\README.txt
(make a backup, import the three .bas files into the .accdb, run
RunMigration). It keeps all data, adds what is missing and refreshes
the saved queries.
