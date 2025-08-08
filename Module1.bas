Attribute VB_Name = "Module1"

Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
If Sheet1.Range("C6").Value = True Then


    ActiveWorkbook.SlicerCaches("Slicer_City").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable1"))
        
Else

    ActiveWorkbook.SlicerCaches("Slicer_City").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable1"))
        
End If

If Sheet1.Range("G6").Value = True Then


    ActiveWorkbook.SlicerCaches("Slicer_City").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable2"))
        
Else

    ActiveWorkbook.SlicerCaches("Slicer_City").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable2"))
        
End If
If Sheet1.Range("K6").Value = True Then


    ActiveWorkbook.SlicerCaches("Slicer_City").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable5"))
        
Else

    ActiveWorkbook.SlicerCaches("Slicer_City").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable5"))
        
End If
End Sub
