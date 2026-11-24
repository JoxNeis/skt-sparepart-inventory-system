Attribute VB_Name = "ErrorCodesEntities"
Option Explicit

' =========================================================
' ERROR CODES - ENTITIES (CLASS MODULES)
'
' Class modules cannot declare Public Const, so the codes
' for each entity live here as a Public Enum.
'
' Usage:
'   ErrorManager.ThrowError entitiesOrder.NegativeId, ...
'   If errInfo.Number = entitiesOrder.NegativeId Then ...
'
' Always qualify the member with the enum name, because the
' same member name (for example NegativeId) is used by
' several entities.
'
' Each entity owns a block of 50 codes inside 11000 - 11999,
' in the order of the table files (src/database/tables).
' =========================================================


' ---------------------------------------------------------
' ManufacturingPlant - manufacturing_plants (11000 - 11049)
' ---------------------------------------------------------
Public Enum entitiesManufacturingPlant
    NegativeId = 11001              ' Id <= 0
    EmptyCode = 11002
    CodeTooLong = 11003             ' > 15
    EmptyName = 11004
    NameTooLong = 11005             ' > 255
    AddressTooLong = 11006          ' > 255
    TypeTooLong = 11007             ' > 50
    RegionTooLong = 11008           ' > 100
End Enum


' ---------------------------------------------------------
' User - users (11050 - 11099)
' ---------------------------------------------------------
Public Enum entitiesUser
    NegativeId = 11051              ' Id <= 0
    EmptyManufacturingPlant = 11052
    EmptyUsername = 11053
    UsernameTooLong = 11054         ' > 64
    EmptyFirstName = 11055
    FirstNameTooLong = 11056        ' > 100
    MiddleNameTooLong = 11057       ' > 100
    LastNameTooLong = 11058         ' > 100
    JobdeskTooLong = 11059          ' > 255
    NegativeTier = 11060            ' Tier < 0
End Enum


' ---------------------------------------------------------
' ActivityLog - activity_logs (11100 - 11149)
' ---------------------------------------------------------
Public Enum entitiesActivityLog
    NegativeId = 11101              ' Id <= 0
    EmptyActivityCode = 11102
    ActivityCodeTooLong = 11103     ' > 50
    EmptyUsername = 11104
    UsernameTooLong = 11105         ' > 64
End Enum


' ---------------------------------------------------------
' Vendor - vendors (11150 - 11199)
' ---------------------------------------------------------
Public Enum entitiesVendor
    NegativeId = 11151              ' Id <= 0
    EmptyCode = 11152
    CodeTooLong = 11153             ' > 40
    EmptyName = 11154
    NameTooLong = 11155             ' > 255
    CompanyNameTooLong = 11156      ' > 255
    EmailTooLong = 11157            ' > 255
    InvalidEmail = 11158
End Enum


' ---------------------------------------------------------
' Sparepart - spareparts (11200 - 11249)
' ---------------------------------------------------------
Public Enum entitiesSparepart
    NegativeId = 11201              ' Id <= 0
    EmptyCode = 11202
    CodeTooLong = 11203             ' > 40
    EmptyName = 11204
    NameTooLong = 11205             ' > 255
    CategoryTooLong = 11206         ' > 50
    ApplicableTooLong = 11207       ' > 255
    EmptyUnitOfMeasurement = 11208
    UnitOfMeasurementTooLong = 11209 ' > 15
    EmptyAlias = 11210
    AliasTooLong = 11211            ' > 255
    StatusTooLong = 11212           ' > 50
End Enum


' ---------------------------------------------------------
' MobileMaintenance - mobile_maintenances (11250 - 11299)
' ---------------------------------------------------------
Public Enum entitiesMobileMaintenance
    NegativeId = 11251              ' Id <= 0
    EmptyManufacturingPlant = 11252
    EmptyCode = 11253
    CodeTooLong = 11254             ' > 40
    EmptyName = 11255
    NameTooLong = 11256             ' > 255
End Enum


' ---------------------------------------------------------
' ManufacturingPlantStock - manufacturing_plant_stocks (11300 - 11349)
' ---------------------------------------------------------
Public Enum entitiesManufacturingPlantStock
    NegativeId = 11301              ' Id <= 0
    EmptyManufacturingPlant = 11302
    EmptySparepart = 11303
    NegativeCurrentStock = 11304
    NegativeReorderPoint = 11305
    NegativeSafetyStock = 11306
    NegativeLeadTimeDays = 11307
    NegativeAverageDailyUsage = 11308
    NegativeDailyDemandStdDev = 11309
    NegativeServiceLevelZ = 11310
End Enum


' ---------------------------------------------------------
' Transaction - transactions (11350 - 11399)
' ---------------------------------------------------------
Public Enum entitiesTransaction
    NegativeId = 11351              ' Id <= 0
    EmptyUUID = 11352
    UUIDTooLong = 11353             ' > 36
    InvalidTransactionType = 11354  ' not ORDER / RECEIVE / REQUEST
    EmptyManufacturingPlant = 11355
    EmptyCreatedBy = 11356
    CreatedByTooLong = 11357        ' > 64
    ReversedByTooLong = 11358       ' > 64
    SyncedByTooLong = 11359         ' > 64
End Enum


' ---------------------------------------------------------
' Order - orders (11400 - 11449)
' ---------------------------------------------------------
Public Enum entitiesOrder
    NegativeId = 11401              ' Id <= 0
    EmptyTransaction = 11402
    WrongTransactionType = 11403    ' header is not ORDER
    EmptySparepart = 11404
    PoNumberTooLong = 11405         ' > 50
    EmptyOrderDate = 11406
    FutureOrderDate = 11407
    InvalidQuantity = 11408         ' Quantity <= 0
    EmptyPersonInCharge = 11409
    PersonInChargeTooLong = 11410   ' > 255
    NegativeOkQuantityTotal = 11411
    NegativeReceivedQuantityTotal = 11412
    NegativeRemainingQuantity = 11413
    RemainingQuantityExceedsQuantity = 11414
    InvalidStatus = 11415
End Enum


' ---------------------------------------------------------
' Receive - receives (11450 - 11499)
' ---------------------------------------------------------
Public Enum entitiesReceive
    NegativeId = 11451              ' Id <= 0
    EmptyTransaction = 11452
    WrongTransactionType = 11453    ' header is not RECEIVE
    EmptyOrder = 11454
    EmptyArrivalDate = 11455
    FutureArrivalDate = 11456
    ArrivalAfterVerification = 11457
    InvalidVerificationDate = 11458 ' < 0
    FutureVerificationDate = 11459
    VerificationBeforeArrival = 11460
    NegativeReceivedQuantity = 11461
    ReceivedQuantityExceedsOrder = 11462
    NegativeOkQuantity = 11463
    OkQuantityExceedsOrder = 11464
    NegativeNotOkQuantity = 11465
    OkNotOkExceedsReceived = 11466
    QuantityNotBalanced = 11467     ' received <> OK + NOT OK
    EmptyPersonInCharge = 11468
    PersonInChargeTooLong = 11469   ' > 255
    NegativeRemainingQuantity = 11470
    RemainingQuantityExceedsOrder = 11471
End Enum


' ---------------------------------------------------------
' Request - requests (11500 - 11549)
' ---------------------------------------------------------
Public Enum entitiesRequest
    NegativeId = 11501              ' Id <= 0
    EmptyTransaction = 11502
    WrongTransactionType = 11503    ' header is not REQUEST
    EmptyMobileMaintenance = 11504
    EmptySparepart = 11505
    EmptyRequestDate = 11506
    FutureRequestDate = 11507
    InvalidQuantity = 11508         ' Quantity <= 0
    EmptyPersonInCharge = 11509
    PersonInChargeTooLong = 11510   ' > 255
End Enum


' ---------------------------------------------------------
' TransactionUpdateRequest - transaction_update_requests (11550 - 11599)
' ---------------------------------------------------------
Public Enum entitiesTransactionUpdateRequest
    NegativeId = 11551              ' Id <= 0
    EmptyOriginalTransaction = 11552
    EmptyProposedTransaction = 11553
    SameTransaction = 11554         ' original = proposed
    TransactionTypeMismatch = 11555
    ManufacturingPlantMismatch = 11556
    EmptyRequestedBy = 11557
    RequestedByTooLong = 11558      ' > 64
    InvalidVerdict = 11559          ' not APPROVED / REJECTED
    ReviewerTooLong = 11560         ' > 64
    ReviewerIsRequester = 11561
    ReviewWithoutVerdict = 11562    ' reviewer set while pending
    VerdictWithoutReview = 11563    ' verdict without reviewer / date
End Enum


' ---------------------------------------------------------
' MaintenanceWork - maintenance_works (11600 - 11649)
' ---------------------------------------------------------
Public Enum entitiesMaintenanceWork
    NegativeId = 11601              ' Id <= 0
    EmptyUUID = 11602
    UUIDTooLong = 11603             ' > 36
    EmptyMaintenanceName = 11604
    MaintenanceNameTooLong = 11605  ' > 255
    EmptyMaintenanceDate = 11606
    FutureMaintenanceDate = 11607
    EmptyMaintenanceType = 11608
    MaintenanceTypeTooLong = 11609  ' > 50
    EmptyDetail = 11610
    DetailOfOtherWork = 11611
End Enum


' ---------------------------------------------------------
' MaintenanceWorkDetail - maintenance_work_details (11650 - 11699)
' ---------------------------------------------------------
Public Enum entitiesMaintenanceWorkDetail
    NegativeId = 11651              ' Id <= 0
    InvalidMaintenanceWorkId = 11652
    InvalidMaintenanceTaskId = 11653
    NegativeAmount = 11654
    SparepartWithoutAmount = 11655
    AmountWithoutSparepart = 11656
End Enum


' ---------------------------------------------------------
' Assembly - assemblies (11700 - 11749)
' ---------------------------------------------------------
Public Enum entitiesAssembly
    NegativeId = 11701              ' Id <= 0
    EmptyParent = 11702
    EmptyChild = 11703
    ParentIsChild = 11704           ' ck_assemblies_not_self
End Enum


' ---------------------------------------------------------
' StockLedgerEntry - vw_stock_ledger, read only (11750 - 11799)
' ---------------------------------------------------------
Public Enum entitiesStockLedgerEntry
    EmptyMovementDate = 11751
    NegativeBeginningStock = 11752
    NegativeIncomingQuantity = 11753
    NegativeOutgoingQuantity = 11754
    NegativeEndStock = 11755
End Enum
