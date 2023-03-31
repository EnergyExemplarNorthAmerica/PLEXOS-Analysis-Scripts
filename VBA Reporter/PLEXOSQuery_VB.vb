Dim e As New PLEXOS_NET.Solution
Dim r As ADODB.Recordset
Dim strModelPath As String
  
Dim SimulationPhaseID As Long
Dim CollectionId As Long
Dim ParentName As String
Dim ChildName As String
Dim PeriodTypeId As Long
Dim SeriesTypeId As Long


Public Sub QueryGeneration()
    SimulationPhaseID = SimulationPhaseEnum.SimulationPhaseEnum_STSchedule
    PeriodTypeId = PeriodEnum.PeriodEnum_Interval
    SeriesTypeId = SeriesTypeEnum_Properties
    ParentName = ""
    ChildName = ""
    CollectionId = CollectionEnum_SystemGenerators
    strModelPath = ActiveWorkbook.Sheets("UI").Range("C4").Value
    
    Dim PropsArray As Variant
    ReDim PropsArray(3)
    PropsArray(0) = SystemOutGeneratorsEnum.SystemOutGeneratorsEnum_GenerationCost
    PropsArray(1) = SystemOutGeneratorsEnum.SystemOutGeneratorsEnum_Generation
    PropsArray(2) = SystemOutGeneratorsEnum.SystemOutGeneratorsEnum_FuelOfftake
    PropsArray(3) = SystemOutGeneratorsEnum.SystemOutGeneratorsEnum_HoursDown
    
    Dim PropList As String
    PropList = Join(PropsArray, ",")
    
    e.Connection (strModelPath)
    Set r = e.Query(SimulationPhaseID, CollectionId, ParentName, ChildName, PeriodTypeId, SeriesTypeId, PropList)

    
    Dim WS As Worksheet
    Set WS = ActiveWorkbook.Sheets("Data")

    For iCols = 0 To r.Fields.Count - 1
        WS.Cells(1, iCols + 1).Value = r.Fields(iCols).Name
    Next
    WS.Range(WS.Cells(1, 1), _
    WS.Cells(1, r.Fields.Count)).Font.Bold = True
    WS.Range("A2").CopyFromRecordset r

End Sub