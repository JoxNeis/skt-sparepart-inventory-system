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
' =========================================================


' ---------------------------------------------------------
' Catalogue (11100 - 11199)
' ---------------------------------------------------------
Public Enum entitiesCatalogue
    NegativeId = 11101              ' Id <= 0
    EmptyName = 11102
    EmptyUnitOfMeasurement = 11103
End Enum


' ---------------------------------------------------------
' MobilePost (11200 - 11299)
' ---------------------------------------------------------
Public Enum entitiesMobilePost
    NegativeId = 11201              ' Id <= 0
    EmptyPlant = 11202
    EmptyName = 11203
    EmptyCode = 11204
End Enum


' ---------------------------------------------------------
' Order (11300 - 11399)
' ---------------------------------------------------------
Public Enum entitiesOrder
    NegativeId = 11301              ' Id <= 0
    EmptyUUID = 11302
    EmptyPoId = 11303
    EmptyPlant = 11304
    EmptyOrderDate = 11305
    FutureOrderDate = 11306
    EmptyCatalogue = 11307
    InvalidAmount = 11308           ' Amount <= 0
    InvalidStatus = 11309
    EmptyVendor = 11310
    EmptyPersonInCharge = 11311
End Enum


' ---------------------------------------------------------
' Plant (11400 - 11499)
' ---------------------------------------------------------
Public Enum entitiesPlant
    NegativeId = 11401              ' Id <= 0
    EmptyName = 11402
    EmptyCode = 11403
End Enum


' ---------------------------------------------------------
' Receive (11500 - 11599)
' ---------------------------------------------------------
Public Enum entitiesReceive
    NegativeId = 11501              ' Id < 0
    EmptyUUID = 11502
    EmptyOrder = 11503
    TotalReceiveExceedsOrder = 11504
    NegativeTotalReceive = 11505
    OkNotOkExceedsTotalReceive = 11506
    NegativeOkAmount = 11507
    OkAmountExceedsOrder = 11508
    NegativeNotOkAmount = 11509
    EmptyUnitOfMeasurement = 11510
    EmptyArrivalDate = 11511
    FutureArrivalDate = 11512
    ArrivalAfterVerification = 11513
    EmptyVerificationDate = 11514
    FutureVerificationDate = 11515
    VerificationBeforeArrival = 11516
    EmptyPersonInCharge = 11517
    NegativeRemainingOrder = 11518
    RemainingOrderExceedsOrder = 11519
End Enum


' ---------------------------------------------------------
' Request (11600 - 11699)
' ---------------------------------------------------------
Public Enum entitiesRequest
    NegativeId = 11601              ' Id < 0
    EmptyUUID = 11602
    EmptyMobilePost = 11603
    EmptyOrderDate = 11604
    FutureOrderDate = 11605
    OrderDateAfterArriveDate = 11606
    EmptyCatalogue = 11607
    InvalidAmount = 11608           ' Amount <= 0
    EmptyUnitOfMeasurement = 11609
    EmptyArriveDate = 11610
    FutureArriveDate = 11611
    ArriveDateBeforeOrderDate = 11612
    EmptyPersonInCharge = 11613
End Enum


' ---------------------------------------------------------
' StockMovement (11700 - 11799)
' ---------------------------------------------------------
Public Enum entitiesStockMovement
    InvalidMovementDate = 11701     ' MovementDate <= 0
    NegativeBeginningStock = 11702
    NegativeIncomingAmount = 11703
    NegativeOutgoingAmount = 11704
    NegativeEndStock = 11705
    EmptyPersonInCharge = 11706
End Enum


' ---------------------------------------------------------
' Vendor (11800 - 11899)
' ---------------------------------------------------------
Public Enum entitiesVendor
    NegativeId = 11801              ' Id <= 0
    EmptyName = 11802
    EmptyCode = 11803
    EmptyCompanyName = 11804
    EmptyEmail = 11805
End Enum
