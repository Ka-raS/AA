Attribute VB_Name = "SetupScheduleView"
' Run once after opening Project_Schedule_v7.xml in Microsoft Project.
' Rebuilds the Entry table of the Gantt Chart as:
'   WBS | Task Name | Duration | Start | Finish | Predecessors (WBS) | Resource Names
' i.e. the ID column is removed and the WBS column (1.x, 1.x.x) is shown first.
' Predecessors use the "WBS Predecessors" field so links read as WBS codes
' (the normal Predecessors field shows ID numbers, which are hidden now).
Sub SetupScheduleView()
    ViewApply Name:="Gantt Chart"

    TableEditEx Name:="Entry", TaskTable:=True, Create:=True, OverwriteExisting:=True, _
        FieldName:="WBS", Title:="WBS", Width:=8, Align:=0, ShowInMenu:=True, _
        LockFirstColumn:=False, DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
        HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Name", Title:="Task Name", _
        Width:=60, Align:=0, LockFirstColumn:=False, DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
        HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Duration", Title:="Duration", _
        Width:=9, Align:=2, LockFirstColumn:=False, DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
        HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Start", Title:="Start", _
        Width:=13, Align:=2, LockFirstColumn:=False, DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
        HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Finish", Title:="Finish", _
        Width:=13, Align:=2, LockFirstColumn:=False, DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
        HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="WBS Predecessors", Title:="Predecessors", _
        Width:=18, Align:=0, LockFirstColumn:=False, DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
        HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Resource Names", Title:="Resource Names", _
        Width:=14, Align:=0, LockFirstColumn:=False, DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
        HeaderAutoRowHeightAdjustment:=False, WrapText:=False

    TableApply Name:="Entry"
End Sub
