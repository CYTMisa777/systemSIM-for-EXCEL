Attribute VB_Name = "Module1"
Sub test01()
Attribute test01.VB_Description = "�W���D�Ѽ������R�۰ʤ�"
Attribute test01.VB_ProcData.VB_Invoke_Func = " \n14"
'
' test01 ����
' �W���D�Ѽ������R�۰ʤ�
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
    SolverAdd CellRef:="$F$4", Relation:=4, FormulaText:="���"
    SolverAdd CellRef:="$F$5", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$5", Relation:=4, FormulaText:="���"
    SolverAdd CellRef:="$F$6", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$6", Relation:=4, FormulaText:="���"
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverSolve
End Sub
