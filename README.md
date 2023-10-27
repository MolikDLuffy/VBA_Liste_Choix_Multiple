# VBA utilisation d'une liste à choix multiple

    Je vous invite à aller voir ma vidéo : https://www.youtube.com/watch?v=FdINbxon1IE&list=PL1eCpY4dpfdxzhLeUG3YwQjR6PLXx_djO&index=8

    Option Explicit
    Dim i As Long
    Dim sTemp As String
    Dim a
    Dim bTest As Boolean

    Private Sub Lb_option_Change()
    If bTest Then
        Exit Sub
    End If
    sTemp = ""
    For i = 0 To Me.Lb_option.ListCount - 1
    If Me.Lb_option.Selected(i) Then
        sTemp = sTemp & Me.Lb_option.List(i) & "-"
    End If
    Next i

    sTemp = VBA.Left(sTemp, VBA.Len(sTemp) - 1)
    ActiveCell = sTemp
    End Sub

    ----------------------------------------------------------
    ----------------------------------------------------------
    Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If ActiveCell.Column = 4 Then
        If Cells(ActiveCell.Row, 3) = "" Then
            Me.Lb_option.Visible = False
            Exit Sub
        End If

    With Me.Lb_option
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .Height = 150
        .Width = 100
        .Top = ActiveCell.Top
        .Left = ActiveCell.Offset(0, 1).Left
        .Visible = True
    End With

    On Error Resume Next

    i = Application.WorksheetFunction.Match(Cells(ActiveCell.Row, 3), Worksheets("Option").Range("Classes"), 0) - 1

    Me.Lb_option.List = Worksheets("Option").Range(Worksheets("Option").Range("A1").Offset(1, i), Worksheets("Option").Range("A1").Offset(0, i).End(xlDown)).Value

    On Error GoTo 0

    a = VBA.Split(ActiveCell, "-")
        If UBound(a) >= 0 Then
        For i = 0 To Me.Lb_option.ListCount - 1
            If Not IsError(Application.Match(Me.Lb_option.List(i), a, 0)) Then
            bTest = True
            Me.Lb_option.Selected(i) = True
            bTest = False
            End If
        Next i
        End If
    Else
    Me.Lb_option.Visible = False
    End If

    End Sub



