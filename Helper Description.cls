VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim ParamDescs As Variant
    ParamDescs = Array( _
        "Cost: Total purchase cost of the asset", _
        "Rate: WDV depreciation rate (e.g., 0.2209)", _
        "Salvage: Residual/salvage value threshold", _
        "Life: Useful life in years", _
        "PurchaseDate: Date asset was acquired", _
        "RangeStart: Start date of reporting period", _
        "RangeEnd: End date of reporting period" _
    )
    
    Application.MacroOptions _
        Macro:="WDV_Depreciation", _
        Description:="Calculates WDV depreciation for a specific date range across Apr-Mar financial years.", _
        Category:="Financial", _
        ArgumentDescriptions:=ParamDescs
End Sub

