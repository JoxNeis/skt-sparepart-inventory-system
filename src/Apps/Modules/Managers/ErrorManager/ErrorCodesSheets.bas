Attribute VB_Name = "ErrorCodesSheets"
Option Explicit

' =========================================================
' ERROR CODES - WORKSHEETS
'
' Worksheet modules cannot declare Public Const, so the codes
' of each sheet live here as a Public Enum named
' sheet + <sheet code name without the leading s>.
'
' Usage:
'   ErrorManager.ThrowError sheetInputRequest.EmptyPlantName, ...
'   If errInfo.Number = sheetInputRequest.EmptyPlantName Then ...
'
' Always qualify the member with the enum name: the same
' member name (for example EmptyPlantName) is used by several
' sheets.
' =========================================================


' ---------------------------------------------------------
' sInputPenerimaan (14200 - 14299)
' ---------------------------------------------------------
Public Enum sheetInputPenerimaan
    EmptyPlantName = 14201              ' save: manufacturing plant cell (B4) is blank
    EmptyOrderUUID = 14202              ' save: ORDER UUID column is blank
    EmptyCatalogueName = 14203          ' save: Nama Barang column is blank
    InvalidArrivalDate = 14204          ' save: Tanggal Datang is not a date
    InvalidVerificationDate = 14205     ' save: Tanggal Verifikasi is not a date
    InvalidTotalReceiveAmount = 14206   ' save: Jumlah Diterima <= 0
    NegativeOkAmount = 14207            ' save: Jumlah OK < 0
    NegativeNotOkAmount = 14208         ' save: Jumlah NOT OK < 0
    InvalidOkNotOkSum = 14209           ' save: OK + NOT OK <> Jumlah Diterima
    EmptyPersonInCharge = 14210         ' save: PIC column is blank
End Enum

' ---------------------------------------------------------
' sInputRequest (14300 - 14399)
' ---------------------------------------------------------
Public Enum sheetInputRequest
    EmptyPlantName = 14301              ' save: manufacturing plant cell (B4) is blank
    InvalidRequestDate = 14302          ' save: request date cell (B6) is not a date
    InvalidAmount = 14303               ' save: Jumlah Barang is not numeric
    EmptyCatalogueName = 14304          ' save: Nama Barang column is blank
    ZeroOrNegativeAmount = 14305        ' save: Jumlah Barang <= 0
    EmptyMobilePostName = 14306         ' save: Mobile Post column is blank
    EmptyPersonInCharge = 14307         ' save: PIC column is blank
End Enum

' ---------------------------------------------------------
' sInventory (14400 - 14499)
' ---------------------------------------------------------
Public Enum sheetInventory
    EmptyPlantName = 14401                  ' search / update ROP: plant cell B4 blank
    EmptyInventoryRecordset = 14402         ' search: service returned no recordset
    EmptyInventoryTableData = 14403         ' update ROP: TableInventory has no rows
    EmptyUpdateActionCell = 14404           ' update ROP: action cell is Nothing
    EmptyCatalogueName = 14405              ' update ROP: Nama Barang blank
    EmptyUnitOfMeasurement = 14406          ' update ROP: Satuan blank
    EmptyAverageDailyUsage = 14407          ' update ROP: Average Daily Usage blank
    InvalidAverageDailyUsage = 14408        ' update ROP: Average Daily Usage not a number
    EmptySdDemand = 14409                   ' update ROP: SD Demand blank
    InvalidSdDemand = 14410                 ' update ROP: SD Demand not a number
    EmptyLeadTime = 14411                   ' update ROP: Lead Time blank
    InvalidLeadTime = 14412                 ' update ROP: Lead Time not a number
    NegativeAverageDailyUsage = 14413       ' update ROP: Average Daily Usage < 0
    NegativeSdDemand = 14414                ' update ROP: SD Demand < 0
    NegativeLeadTime = 14415                ' update ROP: Lead Time < 0
End Enum

' ---------------------------------------------------------
' sKartuStock (14500 - 14599)
' ---------------------------------------------------------
Public Enum sheetKartuStock
    InvalidYearNumber = 14501               ' load: year cell (C6) is not numeric
    EmptyPlantName = 14502                  ' load: location cell (B4) blank
    EmptyCatalogueName = 14503              ' load: catalogue name cell (B8) blank
    InvalidPeriodMonth = 14504              ' load: month outside 1 - 12
    InvalidPeriodYear = 14505               ' load: year <= 0
    InvalidMonthName = 14506                ' month cell (B6) text is not a known month name
End Enum

' ---------------------------------------------------------
' sMasterDataCatalogue (14600 - 14699)
' ---------------------------------------------------------
Public Enum sheetMasterDataCatalogue
    NullCatalogueCollection = 14601         ' serviceCatalogue.GetAllCatalogues returned Nothing
    NullCatalogueInCollection = 14602       ' item in the returned collection is Nothing
    EmptyCatalogueItem = 14603              ' Catalogue parameter of WriteCatalogueTableRow is Nothing
    EmptyCatalogueName = 14604              ' B4 is blank
    EmptyCatalogueAlias = 14605             ' B5 is blank
    EmptyCatalogueCode = 14606              ' B6 is blank
    EmptyCatalogueCategory = 14607          ' B7 is blank
    EmptyCatalogueUnitOfMeasurement = 14608 ' B8 is blank
End Enum

' ---------------------------------------------------------
' sMasterDataMobilePost (14700 - 14799)
' ---------------------------------------------------------
Public Enum sheetMasterDataMobilePost
    EmptyManufacturingPlantName = 14701     ' B3 is blank
    NullMobilePostCollection = 14702        ' service returned Nothing
    NullMobilePostInCollection = 14703      ' item in the returned collection is Nothing
    EmptyMobilePostNameInCollection = 14704 ' item in the returned collection has a blank name
End Enum

' ---------------------------------------------------------
' sMasterDataUser (14800 - 14899)
' ---------------------------------------------------------
Public Enum sheetMasterDataUser
    EmptyPlantName = 14801                  ' save: location cell (B2) blank
    EmptyUsername = 14802                   ' save: username PMI cell (B3) blank
    EmptyFirstName = 14803                  ' save: first name cell (B4) blank
    EmptyLastName = 14804                   ' save: last name cell (B6) blank
    EmptyPosition = 14805                   ' save: position cell (B7) blank
    InvalidTierErrorValue = 14806           ' save: tier cell (B8) holds an error value
    EmptyTier = 14807                       ' save: tier cell blank / Null / Empty
    InvalidTierNumber = 14808               ' save: tier is not a number
    InvalidTierWholeNumber = 14809          ' save: tier has a fractional part
    ZeroOrNegativeTier = 14810              ' save: tier <= 0
    NullUserRecordset = 14811               ' refresh: SelectUsers returned Nothing
    EmptyUserFullName = 14812               ' refresh: record has no first / middle / last name
    InvalidUserId = 14813                   ' refresh: record id <= 0
End Enum

' ---------------------------------------------------------
' sMasterDataVendor (14900 - 14999)
' ---------------------------------------------------------
Public Enum sheetMasterDataVendor
    NullVendorCollection = 14901            ' refresh: SelectVendors returned Nothing
    NullVendorInCollection = 14902          ' refresh: item in the vendor collection is Nothing
    EmptyVendorItem = 14903                 ' Vendor parameter is Nothing (validate / write row)
    InvalidVendorId = 14904                 ' refresh: record id <= 0
    EmptyRecordVendorName = 14905           ' refresh: record has a blank vendor name
    EmptyRecordVendorCompanyName = 14906    ' refresh: record has a blank company name
    EmptyRecordVendorCode = 14907           ' refresh: record has a blank vendor code
    EmptyVendorName = 14908                 ' save: vendor name cell (B3) blank
    EmptyVendorCode = 14909                 ' save: vendor code cell (B4) blank
    EmptyVendorCompanyName = 14910          ' save: company name cell (B5) blank
End Enum

' ---------------------------------------------------------
' sRecapOrder (15000 - 15099)
' ---------------------------------------------------------
Public Enum sheetRecapOrder
    EmptyPlantName = 15001                    ' search: location cell (B4) is blank (old 1101)
    InvalidFromDate = 15002                   ' search: from date cell (B7) is not a date (old 1102)
    InvalidUntilDate = 15003                  ' search: until date cell (B8) is not a date (old 1103)
    InvalidDateRange = 15004                  ' search: from date > until date (old 1104)
    EmptyOrderCollection = 15005              ' search: serviceOrder.SelectOrder returned Nothing (old 1105)
    InvalidSelectedRow = 15006                ' PO ID update: selected row is not a TableRecapOrder row (old 1201)
    EmptyOrderUUID = 15007                    ' PO ID update: UUID cell is blank (old 1202)
    EmptyPoId = 15008                         ' PO ID update: PO ID cell is blank (old 1203)
End Enum

' ---------------------------------------------------------
' sRecapPendingOrder (15100 - 15199)
' ---------------------------------------------------------
Public Enum sheetRecapPendingOrder
    EmptyPlantLocation = 15101                ' search: location cell (B4) is blank
    InvalidFromDate = 15102                   ' search: from date cell (B7) is not a date
    InvalidUntilDate = 15103                  ' search: until date cell (B8) is not a date
    FromDateExceedsUntilDate = 15104          ' search: from date > until date
    EmptyPendingOrderCollection = 15105       ' search: serviceOrder.SelectPendingOrder returned Nothing
    InvalidSelectedOrderRow = 15106           ' update proposal: selected row is not a TableRecapPendingOrder row
    EmptyOriginalOrderUUID = 15107            ' update proposal: UUID cell is blank
    EmptyOrderPoId = 15108                    ' update proposal: PO ID cell is blank
    EmptyOrderCatalogueName = 15109           ' update proposal: Nama Barang cell is blank
    InvalidOrderAmount = 15110                ' update proposal: Jumlah Order is not a number
    ZeroOrNegativeOrderAmount = 15111         ' update proposal: Jumlah Order <= 0
    EmptyOrderVendorName = 15112              ' update proposal: Vendor cell is blank
    InvalidOrderDate = 15113                  ' update proposal: Tanggal Order is not a date
    EmptyUpdatePlantLocation = 15114          ' update proposal: location cell (B4) is blank
    EmptyOrderPersonInCharge = 15115          ' update proposal: PIC cell is blank
End Enum

' ---------------------------------------------------------
' sRecapPenerimaan (15200 - 15299)
' ---------------------------------------------------------
Public Enum sheetRecapPenerimaan
    EmptyPlantName = 15201                  ' search: location cell blank
    InvalidFromDate = 15202                 ' search: from date not a date
    InvalidUntilDate = 15203                ' search: until date not a date
    InvalidDateRange = 15204                ' search: from date > until date
    EmptyReceiveCollection = 15205          ' SelectReceives returned Nothing
    EmptyReceiveTable = 15206               ' update request: tbl is Nothing
    EmptyReceiveTableRows = 15207           ' update request: table has no rows
    InvalidSelectedRow = 15208              ' update request: row index outside the table
    EmptyOriginalReceiveUUID = 15209
    EmptyOrderTransactionUUID = 15210
    EmptyUpdatePlantName = 15211            ' update request: location cell blank
    EmptyCatalogueName = 15212
    EmptyUnitOfMeasurement = 15213
    EmptyPersonInCharge = 15214
    EmptyArrivalDate = 15215
    InvalidArrivalDate = 15216
    InvalidVerificationDate = 15217
    VerificationBeforeArrival = 15218       ' verification date < arrival date
    EmptyReceiveAmount = 15219
    InvalidReceiveAmount = 15220            ' Jumlah diterima not numeric
    EmptyOkAmount = 15221
    InvalidOkAmount = 15222                 ' Jumlah OK not numeric
    EmptyNotOkAmount = 15223
    InvalidNotOkAmount = 15224              ' Jumlah NOT OK not numeric
    ZeroOrNegativeReceiveAmount = 15225     ' Jumlah diterima <= 0
    NegativeOkAmount = 15226
    NegativeNotOkAmount = 15227
    InvalidOkNotOkTotal = 15228             ' OK + NOT OK <> Jumlah diterima
    EmptyFormatReceiveTable = 15229         ' ApplyReceiveTableFormat: tbl is Nothing
    NotFoundReceiveTable = 15230            ' ValidateReceiveTableColumns: tbl is Nothing
End Enum

' ---------------------------------------------------------
' sRecapRequest (15300 - 15399)
' ---------------------------------------------------------
Public Enum sheetRecapRequest
    EmptyPlantName = 15301                  ' search: location cell blank
    InvalidFromDate = 15302                 ' search: from date not a date
    InvalidUntilDate = 15303                ' search: until date not a date
    InvalidDateRange = 15304                ' search: from date > until date
    EmptyRequestTable = 15305               ' update proposal: tbl is Nothing
    EmptyRequestTableRows = 15306           ' update proposal: table has no rows
    InvalidSelectedRow = 15307              ' update proposal: row index outside the table
    EmptyActionCell = 15308                 ' update proposal: UPDATE cell is Nothing
    EmptyOriginalRequestUUID = 15309        ' update proposal: UUID column blank
    EmptyUpdatePlantName = 15310            ' update proposal: location cell blank
    EmptyMobilePostName = 15311
    EmptyCatalogueName = 15312
    EmptyUnitOfMeasurement = 15313
    EmptyPersonInCharge = 15314
    EmptyRequestDate = 15315
    InvalidRequestDate = 15316              ' Tanggal Order not a date
    EmptyRequestAmount = 15317
    InvalidRequestAmount = 15318            ' Jumlah Barang not numeric
    ZeroOrNegativeRequestAmount = 15319     ' Jumlah Barang <= 0
    EmptyProposedRequestUUID = 15320        ' CreateRequest returned a blank UUID
    SameOriginalProposedUUID = 15321        ' original UUID = proposed UUID
End Enum

' ---------------------------------------------------------
' sUpdateOrder (15400 - 15499)
' ---------------------------------------------------------
Public Enum sheetUpdateOrder
    EmptyPlantName = 15401                    ' search filter: manufacturing plant cell (B4) is blank
    EmptyFromDate = 15402                     ' search filter: from date cell (B7) is empty
    InvalidFromDate = 15403                   ' search filter: from date cell (B7) is not a date
    EmptyUntilDate = 15404                    ' search filter: until date cell (B8) is empty
    InvalidUntilDate = 15405                  ' search filter: until date cell (B8) is not a date
    InvalidDateRange = 15406                  ' search filter: until date < from date
    EmptyOriginalOrderUUID = 15407            ' approval: old order UUID (column C) is blank
    EmptyProposedOrderUUID = 15408            ' approval: new order UUID (column L) is blank
    SameOriginalProposedOrderUUID = 15409     ' approval: old UUID = new UUID
    EmptyApproverUser = 15410                 ' approval: approving Windows user cannot be determined
End Enum

' ---------------------------------------------------------
' sUpdatePenerimaan (15500 - 15599)
' ---------------------------------------------------------
Public Enum sheetUpdatePenerimaan
    EmptyPlantName = 15501                    ' search filter: manufacturing plant cell (B4) is blank
    EmptyFromDate = 15502                     ' search filter: from date cell (B7) is empty
    InvalidFromDate = 15503                   ' search filter: from date cell (B7) is not a date
    EmptyUntilDate = 15504                    ' search filter: until date cell (B8) is empty
    InvalidUntilDate = 15505                  ' search filter: until date cell (B8) is not a date
    InvalidDateRange = 15506                  ' search filter: until date < from date
    EmptyApprovalCell = 15507                 ' approval: clicked SETUJUI cell is Nothing
    EmptyRequestUpdateTable = 15508           ' approval: TableRequestUpdateReceive has no rows
    EmptyOriginalReceiveUUID = 15509          ' approval: UUID LAMA is blank
    EmptyProposedReceiveUUID = 15510          ' approval: UUID BARU is blank
    SameOriginalProposedReceiveUUID = 15511   ' approval: UUID LAMA = UUID BARU
    EmptyOriginalCatalogueName = 15512        ' approval: NAMA BARANG LAMA is blank
    EmptyProposedCatalogueName = 15513        ' approval: NAMA BARANG BARU is blank
    NegativeOriginalOkAmount = 15514          ' approval: JUMLAH OK LAMA < 0
    NegativeProposedOkAmount = 15515          ' approval: JUMLAH OK BARU < 0
    EmptyOriginalPlantName = 15516            ' approval: original manufacturing plant (B4) is blank
    EmptyProposedPlantName = 15517            ' approval: proposed manufacturing plant (B4) is blank
    EmptyApproverUser = 15518                 ' approval: approving user cannot be determined
End Enum

' ---------------------------------------------------------
' sUpdateRequest (15600 - 15699)
' ---------------------------------------------------------
Public Enum sheetUpdateRequest
    EmptyPlantName = 15601                    ' search filter: manufacturing plant cell (B4) is blank
    EmptyFromDate = 15602                     ' search filter: from date cell (B7) is empty
    InvalidFromDate = 15603                   ' search filter: from date cell (B7) is not a date
    EmptyUntilDate = 15604                    ' search filter: until date cell (B8) is empty
    InvalidUntilDate = 15605                  ' search filter: until date cell (B8) is not a date
    InvalidDateRange = 15606                  ' search filter: until date < from date
    EmptyApprovalCell = 15607                 ' approval: clicked SETUJUI cell is Nothing
    EmptyOriginalRequestUUID = 15608          ' approval: UUID LAMA is blank
    EmptyProposedRequestUUID = 15609          ' approval: UUID BARU is blank
    SameOriginalProposedRequestUUID = 15610   ' approval: UUID LAMA = UUID BARU
    EmptyOriginalCatalogueName = 15611        ' approval: NAMA BARANG LAMA is blank
    EmptyProposedCatalogueName = 15612        ' approval: NAMA BARANG BARU is blank
    EmptyOriginalPlantName = 15613            ' approval: manufacturing plant (B4) is blank
    EmptyApproverUser = 15614                 ' approval: approving user cannot be determined
End Enum
