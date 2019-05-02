Option Compare Database
Dim CompDirt As Boolean

' You will need to do the following to enable this code:
'   In Form Design View right click and bring up the "Form Properties" window
'     With "Form" selected in the dropdown list, click into the "Other" tab and set "Has Module" to "Yes"
'     Then create the appropriate "On Event" entries for each indicated Form and Control ("Event" tab of the Property sheet)
'     It should read "[Event Procedure]" under the appropriate event if this is linked to the VBA code

Private Sub ComputerID_Dirty(Cancel As Integer)
    CompDirt = True
End Sub

Private Sub DateTagged_DblClick(Cancel As Integer)
    Me.DateTagged = Nz(Me.DateTagged, Date)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Not CompDirt Then Me.ComputerID = FillCompname()
End Sub

Private Sub Form_DataChange(ByVal Reason As Long)
    On Error Resume Next
    If Not CompDirt Then Me.ComputerID = FillCompname()
    Me.Parent!TaggedFishSummary.Requery
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    CompDirt = False
End Sub

Private Sub TagID_Hex_AfterUpdate()
    newval = Me.TagID_Hex.Value
    oldval = Me.TagID_Hex.OldValue
    If newval = oldval Then Exit Sub ' everything's okay
    If Trim(newval) = "" Then
        selval = MsgBox("Sorry, you have to make a selection before trying to leave or save this record" & vbCrLf & "Once this dialogue is closed, you can hit 'Esc' to cancel modifications to the current record", vbOkayOnly Or VbMsgBoxStyle.vbExclamation Or vbSystemModal Or vbDefaultButton2 Or vbMsgBoxSetForeground, "No TagID selection")
        Me.TagID_Hex.Undo
        On Error GoTo 0
        Me.TagID_Hex.SetFocus
        Me.TagID_Hex.Dropdown
        Exit Sub
    End If
    Dim rs As Object
    Set rs = Me.Form.Recordset
    whereclause = "TagID_Hex = '" & newval & "'"
    qry = "SELECT COUNT(*) as c FROM qTagDataEntry WHERE " & whereclause
    qrs = CurrentDb.OpenRecordset(qry)
    If qrs.Count = 1 Then cnt = qrs(0).Value Else cnt = 0
    If cnt >= 1 Then
        thissel = Me.TagID_Hex.ListIndex
        thisloc = Me.Form.CurrentRecord
        If thissel <= 0 Then thissel = 1 ' may not be useful, given the error handling below
        selval = MsgBox("That tag ID has already been selected." & vbCrLf & "Ask surgeon to please wait until you can find out" & vbCrLf & "if this ID or a previously-entered one is incorrect.", vbOkayOnly Or VbMsgBoxStyle.vbExclamation Or vbSystemModal Or vbDefaultButton2 Or vbMsgBoxSetForeground, "Duplicate TagID selection encountered")
        If selval = vbYes Or selval = vbOK Or selval = vbCancel Then
            Me.TagID_Hex.Value = Me.TagID_Hex.Value & "_second"
            rs.Findfirst (whereclause)
            If rs.NoMatch Then
                Me.TagID_Hex.Undo
                On Error GoTo 0
                Me.TagID_Hex.SetFocus
                Me.TagID_Hex.Dropdown
                Exit Sub
            Else
                thatloc = Me.Form.CurrentRecord
                If thatloc = thisloc Then
                    thatloc = rs.FindNext(whereclause)
                    If rs.NoMatch Then
                        On Error GoTo 0
                        Me.TagID_Hex.SetFocus
                        Me.TagID_Hex.Dropdown
                        Exit Sub
                    End If
                End If
            End If
            DoCmd.GoToRecord , , acGoTo, thisloc
            On Error GoTo 0
            Me.TagID_Hex.SetFocus
            Me.TagID_Hex.Dropdown
            ' launch selection in the current entry again.
        End If
    End If
End Sub

Private Sub TagID_Hex_OldAfterUpdate()
    newval = Me.TagID_Hex.Value
    oldval = Me.TagID_Hex.OldValue
    If newval = oldval Then Exit Sub ' everything's okay
    Dim rs As Object
    Set rs = Me.Form.Recordset
    whereclause = "TagID_Hex = '" & newval & "'"
    qry = "SELECT COUNT(*) as c FROM qTagDataEntry WHERE " & whereclause
    qrs = CurrentDb.OpenRecordset(qry)
    If qrs.Count = 1 Then cnt = qrs(0).Value Else cnt = 0
    If cnt >= 1 Then
        thissel = Me.TagID_Hex.ListIndex
        thisloc = Me.Form.CurrentRecord
        If thissel <= 0 Then thissel = 1 ' may not be useful, given the error handling below
        
        selval = MsgBox("Would you like to keep your selection for this record?" & vbCrLf & "[yes = change previous record's TagID; no = change the TagID for this record]", vbYesNo Or VbMsgBoxStyle.vbExclamation Or vbSystemModal Or vbDefaultButton2 Or vbMsgBoxSetForeground, "Duplicate TagID selection encountered")
        If selval = vbYes Then
        '    thissel = Me.TagID_Hex.ListIndex
            Me.TagID_Hex.Value = "temp"
         '   thisloc = Me.Form.CurrentRecord
            rs.Findfirst (whereclause)
            thatloc = Me.Form.CurrentRecord
            Debug.Print (thisloc & " " & thatloc)
            Me.TagID_Hex.Value = "test"
            DoCmd.GoToRecord , , acGoTo, thisloc
            Me.TagID_Hex.Value = newval
            DoCmd.GoToRecord , , acGoTo, thatloc
            If thissel <= 0 Then thissel = 1 ' may not be useful, given the error handling below
            On Error Resume Next
            Me.TagID_Hex.ListIndex = thissel - 1
            On Error GoTo 0
            Me.TagID_Hex.SetFocus
            Me.TagID_Hex.Dropdown
            ' launch selection or set focus to the other entry, preserving selection here
        ElseIf selval = vbNo Then
            On Error Resume Next
            If Me.TagID_Hex.OldValue > "" Then
                Me.TagID_Hex.Value = Me.TagID_Hex.OldValue
            Else
                Me.TagID_Hex.Value = "   "
                thissel = 1
            End If
            DoCmd.GoToRecord , , acGoTo, 1
            DoCmd.GoToRecord , , thisloc
            Me.TagID_Hex.ListIndex = thissel - 1
'            Me.TagID_Hex.Value = Me.TagID_Hex.OldValue
            Me.TagID_Hex.Undo
            On Error GoTo 0
            Me.TagID_Hex.SetFocus
            Me.TagID_Hex.Dropdown
        End If
'    Else
'        Debug.Print ("There were no matches! Das ist gut")
    End If
End Sub

Private Sub Time_in_ana_DblClick(Cancel As Integer)
    Me.Time_in_ana = Nz(Me.Time_in_ana, Time) ' Me.Dynaset.Fields.("Time_in_ana").Value
End Sub

Private Sub Time_out_anac_DblClick(Cancel As Integer)
    Me.Time_out_anac = Nz(Me.Time_out_anac, Time)
End Sub

Private Sub Time_out_surgery_DblClick(Cancel As Integer)
    Me.Time_out_surgery = Nz(Me.Time_out_surgery, Time)
End Sub

Private Sub Time_recovered_DblClick(Cancel As Integer)
    Me.Time_recovered = Nz(Me.Time_recovered, Time)
End Sub

Private Function FillCompname()
    If Len(Me!ComputerID.Value) < 1 Then Me!ComputerID.Value = Null
    FillCompname = Nz(Me.ComputerID, Environ("Computername"))
End Function

Private Sub Form_AfterUpdate()
    Dim cr As Long
    Dim csl As Integer
    Dim cst As Integer
    On Error Resume Next
    Me.Parent!TaggedFishSummary.Requery
    CompDirt = False
    On Error GoTo 0
End Sub
