Attribute VB_Name = "SetupScheduleView"
' Run once after opening 4.xml in Microsoft Project (View > Macros > Macros... > SetupScheduleView > Run).
' Rebuilds the Entry table of the Gantt Chart as:
'   WBS | Task Name | Duration | Start | Finish | Predecessors (WBS) | Resource Names
' i.e. the ID column is removed and the WBS column is shown first:
'   1.x = phases, 1.x.x = work packages, 1.7.3.x / 1.7.4.x = recurring sessions,
'   M = the "Milestones" group at the end of the schedule, M-01..M-09 = milestones (not part of the WBS).
' Predecessors use the "WBS Predecessors" field so links read as WBS codes
' (the normal Predecessors field shows ID numbers, which are hidden now).
Sub SetupScheduleView()
    Dim t As Task, n As Integer

    ' Milestone codes: if Project regenerated them on import (for example 1.8.1 instead of M-01), restore them.
    On Error GoTo WbsFailed
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary And t.Name = "Milestones" Then
                If t.WBS <> "M" Then t.WBS = "M"
            ElseIf t.Milestone And t.OutlineParent.Name = "Milestones" Then
                n = n + 1
                If t.WBS <> "M-" & Format(n, "00") Then t.WBS = "M-" & Format(n, "00")
            End If
        End If
    Next t
    On Error GoTo 0
    If n <> 9 Then MsgBox "Expected 9 milestones in the Milestones group, found " & n & ".", vbExclamation

    ViewApply Name:="Gantt Chart"

    TableEditEx Name:="Entry", TaskTable:=True, Create:=True, OverwriteExisting:=True, _
        FieldName:="WBS", Title:="WBS", Width:=8, Align:=0, ShowInMenu:=True, _
        LockFirstColumn:=False, DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
        HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Name", Title:="Task Name", _
        Width:=56, Align:=0, LockFirstColumn:=False, DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
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
    Exit Sub

WbsFailed:
    MsgBox "Could not set the milestone WBS codes (" & Err.Description & "). " & _
           "Check Project > WBS > Define Code: no code mask should be defined.", vbExclamation
End Sub
