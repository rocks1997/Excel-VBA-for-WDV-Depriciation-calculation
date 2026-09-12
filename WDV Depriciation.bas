Attribute VB_Name = "Module1"
Function WDV_Depreciation( _
    Cost As Double, _
    Rate As Double, _
    Salvage As Double, _
    Life As Long, _
    PurchaseDate As Date, _
    RangeStart As Date, _
    RangeEnd As Date _
) As Double
Attribute WDV_Depreciation.VB_Description = "Calculates WDV depreciation for a specific date range across Apr-Mar financial years."
Attribute WDV_Depreciation.VB_ProcData.VB_Invoke_Func = " \n1"

    Dim p_fy As Long
    Dim fy1_start As Date, fy1_end As Date
    Dim days_in_fy1 As Double, days_used_fy1 As Double
    Dim y1_full As Double, y1_dep As Double
    Dim overlap1 As Double, contrib1 As Double
    Dim wdv_after_y1 As Double
    
    ' Determine April-March Financial Year for Purchase Date
    If Month(PurchaseDate) >= 4 Then
        p_fy = Year(PurchaseDate)
    Else
        p_fy = Year(PurchaseDate) - 1
    End If
    
    fy1_start = DateSerial(p_fy, 4, 1)
    fy1_end = DateSerial(p_fy + 1, 3, 31)
    
    days_in_fy1 = fy1_end - fy1_start + 1
    days_used_fy1 = fy1_end - PurchaseDate + 1
    
    ' Year 1 Depreciation Calculation
    y1_full = Cost * Rate
    If (Cost - Salvage) < y1_full Then y1_full = Cost - Salvage
    
    y1_dep = y1_full * (days_used_fy1 / days_in_fy1)
    If y1_dep < 0 Then y1_dep = 0
    
    ' Year 1 Overlap & Contribution
    Dim min_end1 As Date, max_start1 As Date
    min_end1 = IIf(fy1_end < RangeEnd, fy1_end, RangeEnd)
    max_start1 = IIf(PurchaseDate > RangeStart, PurchaseDate, RangeStart)
    overlap1 = min_end1 - max_start1 + 1
    If overlap1 < 0 Then overlap1 = 0
    
    If overlap1 <= 0 Then
        contrib1 = 0
    Else
        contrib1 = y1_dep * (overlap1 / days_used_fy1)
    End If
    
    wdv_after_y1 = Cost - y1_dep
    
    ' Subsequent Years Calculation (Loop through Year 2 to Life)
    Dim total_n As Double
    total_n = 0
    
    Dim max_years As Long
    max_years = IIf(Life > 2, Life, 2)
    
    Dim n As Long
    For n = 2 To max_years
        Dim valid_n As Double
        valid_n = IIf(n <= Life, 1#, 0#)
        
        Dim fyn_start As Date, fyn_end As Date
        fyn_start = DateSerial(p_fy + n - 1, 4, 1)
        fyn_end = DateSerial(p_fy + n, 3, 31)
        
        Dim days_in_fyn As Double
        days_in_fyn = fyn_end - fyn_start + 1
        
        Dim raw_open As Double, wdv_open As Double
        Dim raw_close As Double, wdv_close As Double
        
        raw_open = wdv_after_y1 * ((1 - Rate) ^ (n - 2))
        wdv_open = IIf(raw_open > Salvage, raw_open, Salvage)
        
        raw_close = wdv_after_y1 * ((1 - Rate) ^ (n - 1))
        wdv_close = IIf(raw_close > Salvage, raw_close, Salvage)
        
        Dim dep_n As Double
        dep_n = (wdv_open - wdv_close) * valid_n
        
        Dim min_end_n As Date, max_start_n As Date
        min_end_n = IIf(fyn_end < RangeEnd, fyn_end, RangeEnd)
        max_start_n = IIf(fyn_start > RangeStart, fyn_start, RangeStart)
        
        Dim overlap_n As Double
        overlap_n = min_end_n - max_start_n + 1
        If overlap_n < 0 Then overlap_n = 0
        
        Dim contrib_n As Double
        If overlap_n <= 0 Then
            contrib_n = 0
        Else
            contrib_n = dep_n * (overlap_n / days_in_fyn)
        End If
        
        total_n = total_n + contrib_n
    Next n
    
    ' Final rounded output
    WDV_Depreciation = Round(contrib1 + total_n, 0)

End Function

