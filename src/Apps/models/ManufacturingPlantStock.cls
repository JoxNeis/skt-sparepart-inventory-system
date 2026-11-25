Attribute VB_Name = "ManufacturingPlantStock"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - MANUFACTURING PLANT STOCK
' Table      : manufacturing_plant_stocks
' Error codes: entitiesManufacturingPlantStock (managerErrorCodesEntities.bas)
'
' One row per (plant, sparepart): on-hand stock and the
' reorder point (ROP) planning inputs and outputs.
'   safety_stock  = service_level_z * daily_demand_std_dev
'                   * Sqr(lead_time_days)
'   reorder_point = safety_stock
'                   + average_daily_usage * lead_time_days
' =========================================================

Private Const CLASS_NAME As String = "ManufacturingPlantStock"

' Table default: 1.65 = 95 % service level.
Private Const DEFAULT_SERVICE_LEVEL_Z As Double = 1.65

Private mId As Long
Private mManufacturingPlant As ManufacturingPlant
Private mSparepart As Sparepart
Private mCurrentStock As Long
Private mReorderPoint As Long
Private mSafetyStock As Long
Private mLeadTimeDays As Long
Private mAverageDailyUsage As Double
Private mDailyDemandStdDev As Double
Private mServiceLevelZ As Double
Private mCreatedAt As Date
Private mUpdatedAt As Date


Private Sub Class_Initialize()
    mServiceLevelZ = DEFAULT_SERVICE_LEVEL_Z
End Sub


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If managerValidationNumeric.IsZeroOrNegative(value) Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' MANUFACTURING PLANT                   [manufacturing_plant_id] NOT NULL
' =========================================================
Public Property Get ManufacturingPlant() As ManufacturingPlant
    Set ManufacturingPlant = mManufacturingPlant
End Property

Public Property Set ManufacturingPlant(ByVal value As ManufacturingPlant)

    If value Is Nothing Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.EmptyManufacturingPlant, _
            CLASS_NAME & ".ManufacturingPlant", _
            CLASS_NAME & ": Plant tidak boleh kosong."
    End If

    Set mManufacturingPlant = value

End Property


' =========================================================
' SPAREPART                             [sparepart_id] NOT NULL
' =========================================================
Public Property Get Sparepart() As Sparepart
    Set Sparepart = mSparepart
End Property

Public Property Set Sparepart(ByVal value As Sparepart)

    If value Is Nothing Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.EmptySparepart, _
            CLASS_NAME & ".Sparepart", _
            CLASS_NAME & ": Sparepart tidak boleh kosong."
    End If

    Set mSparepart = value

End Property


' =========================================================
' CURRENT STOCK                         [current_stock] LONG >= 0
' Cache of the stock ledger (vw_stock_ledger).
' =========================================================
Public Property Get CurrentStock() As Long
    CurrentStock = mCurrentStock
End Property

Public Property Let CurrentStock(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.NegativeCurrentStock, _
            CLASS_NAME & ".CurrentStock", _
            CLASS_NAME & ": Stok saat ini tidak boleh negatif."
    End If

    mCurrentStock = value

End Property


' =========================================================
' REORDER POINT                         [reorder_point] LONG >= 0
' =========================================================
Public Property Get ReorderPoint() As Long
    ReorderPoint = mReorderPoint
End Property

Public Property Let ReorderPoint(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.NegativeReorderPoint, _
            CLASS_NAME & ".ReorderPoint", _
            CLASS_NAME & ": Reorder point tidak boleh negatif."
    End If

    mReorderPoint = value

End Property


' =========================================================
' SAFETY STOCK                          [safety_stock] LONG >= 0
' =========================================================
Public Property Get SafetyStock() As Long
    SafetyStock = mSafetyStock
End Property

Public Property Let SafetyStock(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.NegativeSafetyStock, _
            CLASS_NAME & ".SafetyStock", _
            CLASS_NAME & ": Safety stock tidak boleh negatif."
    End If

    mSafetyStock = value

End Property


' =========================================================
' LEAD TIME DAYS                        [lead_time_days] LONG >= 0
' =========================================================
Public Property Get LeadTimeDays() As Long
    LeadTimeDays = mLeadTimeDays
End Property

Public Property Let LeadTimeDays(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.NegativeLeadTimeDays, _
            CLASS_NAME & ".LeadTimeDays", _
            CLASS_NAME & ": Lead time tidak boleh negatif."
    End If

    mLeadTimeDays = value

End Property


' =========================================================
' AVERAGE DAILY USAGE                   [average_daily_usage] DOUBLE >= 0
' =========================================================
Public Property Get AverageDailyUsage() As Double
    AverageDailyUsage = mAverageDailyUsage
End Property

Public Property Let AverageDailyUsage(ByVal value As Double)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.NegativeAverageDailyUsage, _
            CLASS_NAME & ".AverageDailyUsage", _
            CLASS_NAME & ": Rata-rata pemakaian harian tidak boleh negatif."
    End If

    mAverageDailyUsage = value

End Property


' =========================================================
' DAILY DEMAND STD DEV                  [daily_demand_std_dev] DOUBLE >= 0
' =========================================================
Public Property Get DailyDemandStdDev() As Double
    DailyDemandStdDev = mDailyDemandStdDev
End Property

Public Property Let DailyDemandStdDev(ByVal value As Double)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.NegativeDailyDemandStdDev, _
            CLASS_NAME & ".DailyDemandStdDev", _
            CLASS_NAME & ": Standar deviasi permintaan harian tidak boleh negatif."
    End If

    mDailyDemandStdDev = value

End Property


' =========================================================
' SERVICE LEVEL Z                       [service_level_z] DOUBLE >= 0
'                                       DEFAULT 1.65
' =========================================================
Public Property Get ServiceLevelZ() As Double
    ServiceLevelZ = mServiceLevelZ
End Property

Public Property Let ServiceLevelZ(ByVal value As Double)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesManufacturingPlantStock.NegativeServiceLevelZ, _
            CLASS_NAME & ".ServiceLevelZ", _
            CLASS_NAME & ": Service level (Z) tidak boleh negatif."
    End If

    mServiceLevelZ = value

End Property


' =========================================================
' TIMESTAMPS        [created_at] [updated_at]
' Set by the database; 0 = NULL.
' =========================================================
Public Property Get CreatedAt() As Date
    CreatedAt = mCreatedAt
End Property

Public Property Let CreatedAt(ByVal value As Date)
    mCreatedAt = value
End Property

Public Property Get UpdatedAt() As Date
    UpdatedAt = mUpdatedAt
End Property

Public Property Let UpdatedAt(ByVal value As Date)
    mUpdatedAt = value
End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================

' Same as manufacturing_plant_stock_list.is_at_reorder_point.
Public Property Get IsAtReorderPoint() As Boolean
    IsAtReorderPoint = (mCurrentStock <= mReorderPoint)
End Property

' service_level_z * daily_demand_std_dev * Sqr(lead_time_days),
' rounded up to a whole unit - the same formula and rounding
' as manufacturing_plant_stock_recalculate_reorder_point.
Public Property Get CalculatedSafetyStock() As Long

    CalculatedSafetyStock = RoundUp(RawSafetyStock())

End Property

' safety stock + average_daily_usage * lead_time_days, rounded
' up to a whole unit (once, on the sum, as the query does).
Public Property Get CalculatedReorderPoint() As Long

    CalculatedReorderPoint = RoundUp( _
        RawSafetyStock() + mAverageDailyUsage * mLeadTimeDays)

End Property

' Copies the calculated values into SafetyStock and
' ReorderPoint.
Public Sub RecalculateReorderPoint()

    SafetyStock = CalculatedSafetyStock
    ReorderPoint = CalculatedReorderPoint

End Sub


' =========================================================
' INTERNAL
' =========================================================
Private Function RawSafetyStock() As Double
    RawSafetyStock = mServiceLevelZ * mDailyDemandStdDev * Sqr(mLeadTimeDays)
End Function

' -Int(-Round(x, 6)): ceiling after dropping floating point
' noise, same as the query.
Private Function RoundUp(ByVal value As Double) As Long
    RoundUp = -Int(-Round(value, 6))
End Function
