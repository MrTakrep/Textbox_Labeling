Option Explicit
Dim txt_label As MSForms.label
Dim ctrl As Control

Sub textbox_labeling()
    For Each ctrl In UserForm1.Controls
        If TypeName(ctrl) = "TextBox" And InStr(ctrl.Tag, "_lb_") Then
            With UserForm1
                Set txt_label = .Controls.Add("Forms.Label.1", "txt_label" & ctrl.Value, True)
                With txt_label
                    .Caption = ctrl.Value
                    .Font.Name = ctrl.Font.Name
                    .Font.Size = ctrl.Font.Size
                    .Height = ctrl.Height
                    .Left = ctrl.Left
                    .TextAlign = fmTextAlignLeft
                    .Top = ctrl.Top - (ctrl.Height / 1.2)
                    .Width = ctrl.Width
                    .ZOrder (1)
                End With
            End With
            ctrl.Value = ""
        End If
    Next ctrl
End Sub
