Attribute VB_Name = "Module1"
Sub test01()
Attribute test01.VB_Description = "規劃求解模擬分析自動化"
Attribute test01.VB_ProcData.VB_Invoke_Func = " \n14"
'
' test01 巨集
' 規劃求解模擬分析自動化
'

'
    SolverReset
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverAdd CellRef:="$C$9", Relation:=3, FormulaText:="$C$7"
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverAdd CellRef:="$C$10", Relation:=3, FormulaText:="$C$8"
    SolverAdd CellRef:="$F$4", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$4", Relation:=4, FormulaText:="整數"
    SolverAdd CellRef:="$F$5", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$5", Relation:=4, FormulaText:="整數"
    SolverAdd CellRef:="$F$6", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$6", Relation:=4, FormulaText:="整數"
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverSolve
End Sub
