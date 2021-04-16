VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Formulário"
   ClientHeight    =   8112
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6156
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nr_comps As Integer
Dim nr_groups As Integer
Dim nr_students As Integer
Dim nr_extra As Integer

Dim grade1_idx As Integer
Dim grade2_idx As Integer

Dim editGroup As Integer
Dim editComp As Integer
Dim editParam As Integer

Private Sub AddComponentButton_Click()
    Dim comp As String
    Dim old_comp As String
    Dim group As String
    Dim weight As Integer
    Dim weight_txt As String
    Dim flag As Integer
    Dim group_idx As Integer
    Dim i As Integer
    Dim idx As Integer
    
    comp = ComponentText.Text
    weight_txt = ComponentWeightText.Text
    
    If Len(Trim(comp)) = 0 Or Len(Trim(weight_txt)) = 0 Then
        MsgBox "Campo vazio."
        ComponentText = ""
        ComponentWeightText = ""
        GroupCombo.ListIndex = 0
        CheckBox1.Value = False
        AddComponentButton.Caption = "Adicionar"
        ComponentList.ListIndex = -1
        editComp = 0
        Exit Sub
    End If

    weight = ComponentWeightText.Text
    group_idx = GroupCombo.ListIndex
    group = GroupCombo.List(group_idx)
    
    If Not (IsNumeric(weight)) Then
        MsgBox "Peso não numérico."
        ComponentWeightText = ""
        Exit Sub
    End If

    If CheckBox1 = True Then
        flag = 1
    Else
        flag = 0
    End If

    If editComp = 0 Then
        With ComponentList
            .AddItem
            .List(nr_comps, 0) = comp
            .List(nr_comps, 1) = weight
            .List(nr_comps, 2) = group
        End With
        nr_comps = nr_comps + 1
        Call UpdateComponent(comp, ".", weight, group, flag, True)
    Else
        With ComponentList
            idx = .ListIndex
            old_comp = .List(idx, 0)
            .List(idx, 0) = comp
            .List(idx, 1) = weight
            .List(idx, 2) = group
        End With
        Call UpdateComponent(comp, old_comp, weight, group, flag, False)
    End If
    
    With ComponentList
        ComponentCombo.Clear
        For i = 0 To .ListCount - 1
            ComponentCombo.AddItem (.List(i))
        Next i
        ComponentCombo.ListIndex = -1
        ParameterList.Clear
    End With

    ComponentText = ""
    ComponentWeightText = ""
    ParameterText = ""
    ParameterWeightText = ""
    GroupCombo.ListIndex = 0
    CheckBox1.Value = False
    AddComponentButton.Caption = "Adicionar"
    GroupList.ListIndex = -1
    ComponentList.ListIndex = -1
    editComp = 0
End Sub

Private Sub AddGroupButton_Click()
    Dim group As String
    Dim idx As Integer
    Dim i As Integer

    group = GroupText.Text

    If Len(Trim(group)) = 0 Then
        MsgBox "Campo vazio."
        GroupText = ""
        AddGroupButton.Caption = "Adicionar"
        GroupList.ListIndex = -1
        editGroup = 0
        Exit Sub
    End If

    If editGroup = 0 Then
        With GroupList
            .AddItem
            .List(nr_groups) = group
        End With
        nr_groups = nr_groups + 1
    Else
        With GroupList
            idx = .ListIndex
            .List(idx) = group
            .ListIndex = -1
        End With
    End If

    With GroupList
        GroupCombo.Clear
        GroupCombo.AddItem "Sem grupo"
        GroupCombo.ListIndex = 0
        For i = 0 To .ListCount - 1
            GroupCombo.AddItem (.List(i))
        Next i
    End With

    GroupText = ""
    AddGroupButton.Caption = "Adicionar"
    GroupList.ListIndex = -1
    ComponentList.ListIndex = -1
    editGroup = 0
End Sub

Private Sub AddParameterButton_Click()
    Dim comp As String
    Dim param As String
    Dim old_param As String
    Dim weight As Integer
    Dim weight_txt As String
    Dim params_count As Integer
    Dim idx As Integer
    Dim i As Integer

    param = ParameterText.Text
    weight_txt = ParameterWeightText.Text
    
    If Len(Trim(param)) = 0 Or Len(Trim(weight_txt)) = 0 Then
        MsgBox "Campo vazio."
        ParameterText = ""
        ParameterWeightText = ""
        ComponentCombo.ListIndex = -1
        AddParameterButton.Caption = "Adicionar"
        ParameterList.ListIndex = -1
        editParam = 0
        Exit Sub
    End If
    
    weight = ParameterWeightText.Text
    idx = ComponentCombo.ListIndex

    If idx = -1 Then
        MsgBox "Não selecionou nenhuma componente."
        Exit Sub
    End If
    
    If Not (IsNumeric(weight)) Then
        MsgBox "Peso não numérico."
        ParameterWeightText = ""
        Exit Sub
    End If

    comp = ComponentCombo.List(idx)
    params_count = CountParameters(comp)

    If editParam = 0 Then
        With ParameterList
            .AddItem
            .List(params_count, 0) = param
            .List(params_count, 1) = weight
        End With
        Call UpdateParameter(comp, param, ".", weight, True)
    Else
        With ParameterList
            i = .ListIndex
            old_param = .List(i, 0)
            .List(i, 0) = param
            .List(i, 1) = weight
        End With
        Call UpdateParameter(comp, param, old_param, weight, False)
    End If

    ParameterText = ""
    ParameterWeightText = ""
    AddParameterButton.Caption = "Adicionar"
    ParameterList.ListIndex = -1
    GroupList.ListIndex = -1
    editParam = 0
End Sub

Private Sub ComponentCombo_Change()
    Sheets("Aux").Activate

    Dim i As Integer
    Dim row As Integer
    Dim idx As Integer
    Dim params_count As Integer
    
    idx = ComponentCombo.ListIndex

    If idx = -1 Then
        Exit Sub
    End If
    
    row = GetComponentRow(ComponentCombo.List(idx))
    params_count = range("D" & row).Value
    row = row + 1

    With ParameterList
        .Clear
        For i = 0 To params_count - 1
            .AddItem
            .List(i, 0) = range("A" & row).Value
            .List(i, 1) = range("E" & row).Value
            row = row + 1
        Next i
    End With
End Sub


Private Sub ComponentList_Change()
    Dim idx As Integer

    If ComponentList.ListCount < 1 Then
        Exit Sub
    End If

    idx = ComponentList.ListIndex

    If idx = -1 Then
        ParameterList.Clear
        Exit Sub
    End If
    
    GroupText.Text = ""
    GroupList.ListIndex = -1
    
    ComponentCombo.ListIndex = ComponentList.ListIndex
End Sub


Private Sub EditComponentButton_Click()
    Dim idx As Integer
    Dim comp_row As Integer
    Dim extra As Integer

    If ComponentList.ListCount < 1 Then
        Exit Sub
    End If

    idx = ComponentList.ListIndex

    If idx = -1 Then
        MsgBox "Não selecionou nenhuma componente."
        Exit Sub
    End If

    ComponentText.Text = ComponentList.List(idx, 0)
    ComponentWeightText.Text = ComponentList.List(idx, 1)
    GroupCombo.ListIndex = GetGroupID(ComponentList.List(idx, 2))
    
    comp_row = GetComponentRow(ComponentList.List(idx, 0))
    If comp_row <> -1 Then
        extra = Sheets("Aux").range("F" & comp_row).Value
        If extra = 1 Then
            CheckBox1 = True
        End If
    End If
    
    AddComponentButton.Caption = "Atualizar"
    editComp = 1
End Sub

Private Sub EditGroupButton_Click()
    Dim idx As Integer

    If GroupList.ListCount < 1 Then
        Exit Sub
    End If

    idx = GroupList.ListIndex

    If idx = -1 Then
        MsgBox "Não selecionou nenhum grupo."
        Exit Sub
    End If

    GroupText.Text = GroupList.List(idx)

    AddGroupButton.Caption = "Atualizar"
    editGroup = 1
End Sub

Private Sub EditParameterButton_Click()
    Dim idx As Integer

    If ParameterList.ListCount < 1 Then
        Exit Sub
    End If

    idx = ParameterList.ListIndex

    If idx = -1 Then
        MsgBox "Não selecionou nenhum parâmetro."
        Exit Sub
    End If

    ParameterText.Text = ParameterList.List(idx, 0)
    ParameterWeightText.Text = ParameterList.List(idx, 1)

    AddParameterButton.Caption = "Atualizar"
    editParam = 1
End Sub

Private Sub ExecuteButton_Click()
    If nr_comps > 0 Then
        CreateGroupsSheet
        CreateComponentsSheet
        UpdateGlobalSheet
        CreateSynthesisSheet
        Sheets("Aux").range("B1").Value = 1
        Sheets("Global").Activate
        range("A1").Select
    End If
    Application.ScreenUpdating = True
    UserForm.Hide
    Unload UserForm
End Sub

Private Sub GroupList_Change()
    Dim idx As Integer

    If GroupList.ListCount < 1 Then
        Exit Sub
    End If

    idx = GroupList.ListIndex

    If idx = -1 Then
        Exit Sub
    End If
    
    GroupCombo.ListIndex = idx + 1
End Sub


' ----------------------------------------------
'               General help functions
' ----------------------------------------------

Private Sub UserForm_Initialize()
    Application.ScreenUpdating = False
    InitializeUserFormElements
    nr_comps = 0
    nr_groups = 0
    nr_extra = 0
    CountStudents
    CreateAuxSheet
End Sub

Private Sub InitializeUserFormElements()
    'Focus no group
    GroupCombo.AddItem "Sem grupo"
    GroupCombo.ListIndex = 0
    
    With GroupList
        .TextAlign = fmTextAlignCenter
    End With

    With ComponentList
        .ColumnCount = 3
        .Height = 113.05
        .Width = 117
        .ColumnWidths = "1.5cm;0.75cm;0.8cm"
        .TextAlign = fmTextAlignCenter
    End With

    With ParameterList
        .ColumnCount = 2
        .Height = 91.8
        .Width = 117
        .ColumnWidths = "2.7cm;0.75cm"
        .TextAlign = fmTextAlignLeft
    End With
End Sub

Private Sub CountStudents()
    nr_students = Evaluate("COUNTA(Global!A:A)-1")
End Sub

Private Function GetGroupID(group As String) As Integer
    Dim i As Integer
    
    For i = 0 To GroupCombo.ListCount - 1
        If StrComp(GroupCombo.List(i), group) = 0 Then
            GetGroupID = i
            Exit Function
        End If
    Next i
    GetGroupID = -1
End Function

' ----------------------------------------------
'               Aux sheet
' ----------------------------------------------

Private Sub CreateAuxSheet()
    Sheets.Add.name = "Aux"
    Worksheets("Aux").Move After:=Worksheets(Worksheets.count)
    Sheets("Aux").Activate
    
    'Update column widths
    Columns("A").ColumnWidth = 20
    Columns("B").ColumnWidth = 8
    Columns("C").ColumnWidth = 15
    Columns("D").ColumnWidth = 8
    Columns("E").ColumnWidth = 15
    
    'Create structure
    range("A1").Value = "Formulário executado"
    range("B1").Value = 0
    
    range("A2").Value = "#Alunos"
    range("B2").Value = nr_students
    
    range("A3").Value = "#Componentes"
    range("B3").Value = 0
    
    range("A4").Value = "Linha para componente"
    range("B4").Value = 7
    
    range("A6").Value = "Componente"
    range("B6").Value = "Peso"
    range("C6").Value = "Grupo"
    range("D6").Value = "#Params"
    range("E6").Value = "Peso dos params"
    range("F6").Value = "Recurso"
    
    'Format structure
    range("A1:A4").Select
    With Selection
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(237, 237, 237)
    End With
    
    range("B1:B4").Select
    With Selection
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 230, 153)
    End With
    
    range("A6:F6").Select
    With Selection
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(237, 237, 237)
    End With

    range("A1").Select
    Sheets("Global").Activate
End Sub

Private Sub UpdateComponent(ByVal new_name As String, ByVal old_name As String, ByVal weight As Integer, _
                            ByVal group As String, ByVal flag As Integer, ByVal toAdd As Boolean)
    
    Sheets("Aux").Activate

    Dim row As Integer
    Dim prev_flag As Integer
    
    If toAdd = True Then
        ' Add component
        row = NextComponentRow
        range("A" & row).Value = new_name
        range("A" & row).Offset(0, 1).Value = weight
        range("A" & row).Offset(0, 2).Value = group
        range("A" & row).Offset(0, 3).Value = 0
        range("A" & row).Offset(0, 5).Value = flag
        range("B3") = nr_comps
        range("B4").Value = row + 1
        If flag = 1 Then
            nr_extra = nr_extra + 1
        End If
    Else
        ' Edit component
        row = GetComponentRow(old_name)
        prev_flag = range("A" & row).Offset(0, 5).Value
        range("A" & row).Value = new_name
        range("A" & row).Offset(0, 1).Value = weight
        range("A" & row).Offset(0, 2).Value = group
        range("A" & row).Offset(0, 5).Value = flag
        If prev_flag = 1 And flag = 0 Then
            nr_extra = nr_extra - 1
        End If
        If prev_flag = 0 And flag = 1 Then
            nr_extra = nr_extra + 1
        End If
    End If
    
    range("A" & row & ":F" & row).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(217, 225, 242)
    End With
End Sub

Private Function NextComponentRow() As Integer
    NextComponentRow = Sheets("Aux").Evaluate("B4")
End Function

Private Function GetComponentRow(comp_name As String) As Integer
    Sheets("Aux").Activate

    Dim nr_comp As Integer
    Dim row As Integer
    Dim i As Integer
    
    row = 7
    nr_comp = range("B3").Value
    
    For i = 0 To nr_comp - 1
        If StrComp(comp_name, range("A" & row).Value) = 0 Then
            GetComponentRow = row
            Exit Function
        End If
        
        row = row + range("D" & row) + 1
    Next i
    
    GetComponentRow = -1
End Function

Private Sub UpdateParameter(ByVal comp_name As String, ByVal new_param_name As String, ByVal old_param_name As String, _
                            ByVal weight As Integer, ByVal toAdd As Boolean)
                            
    Sheets("Aux").Activate
    
    Dim comp_row As Integer
    Dim param_row As Integer
    Dim aux1 As Integer
    Dim aux2 As Integer
    
    comp_row = GetComponentRow(comp_name)
    
    If toAdd = True Then
        param_row = NextParameterRow(comp_name)
        aux1 = range("B4").Value + 1
        aux2 = range("D" & comp_row).Value + 1
           
        range("A" & param_row).EntireRow.Insert
        
        range("A" & param_row).Value = new_param_name
        range("E" & param_row).Value = weight
        
        range("A" & param_row).Interior.Color = RGB(226, 239, 218)
        range("E" & param_row).Interior.Color = RGB(226, 239, 218)
        
        range("B" & param_row & ":D" & param_row).Interior.ColorIndex = 0
        range("F" & param_row).Interior.ColorIndex = 0
        
        range("B4").Value = aux1
        range("D" & comp_row).Value = aux2
    Else
        param_row = GetParameterRow(comp_name, old_param_name)
        
        If param_row = -1 Then
            MsgBox "Erro ao atualizar parâmetro"
        End If
        
        range("A" & param_row).Value = new_param_name
        range("E" & param_row).Value = weight
    End If
End Sub

Private Function NextParameterRow(comp_name As String) As Integer
    Sheets("Aux").Activate
    
    Dim comp_row As Integer
    Dim params_count As Integer
    
    comp_row = GetComponentRow(comp_name)
    
    If comp_row = -1 Then
        NextParameterRow = -1
        Exit Function
    End If
    
    params_count = range("D" & comp_row).Value
    
    NextParameterRow = comp_row + params_count + 1
End Function

Private Function GetParameterRow(comp_name As String, param_name As String) As Integer
    Sheets("Aux").Activate
    
    Dim comp_row As Integer
    comp_row = GetComponentRow(comp_name)
    
    If comp_row = -1 Then
        'component does not exist
        GetParameterRow = -1
        Exit Function
    End If
    
    Dim row As Integer
    Dim params_count As Integer
    
    params_count = range("D" & comp_row).Value
    
    For row = comp_row + 1 To comp_row + params_count
        If StrComp(range("A" & row).Value, param_name) = 0 Then
            GetParameterRow = row
            Exit Function
        End If
    Next row
    
    GetParameterRow = -1
    
End Function

Private Function CountParameters(ByVal comp_name As String) As Integer
    Sheets("Aux").Activate
    
    Dim comp_row As Integer
    comp_row = GetComponentRow(comp_name)
    
    If comp_row = -1 Then
        'component does not exist
        CountParameters = -1
        Exit Function
    End If
    
    CountParameters = range("D" & comp_row).Value
End Function

Private Function FormWasExecuted() As Boolean
    If SheetExists("Aux") = False Then
        FormWasExecuted = False
        Exit Function
    End If

    If Sheets("Aux").range("B1").Value = 1 Then
        FormWasExecuted = True
    Else
        FormWasExecuted = False
    End If
End Function

Private Function SheetExists(name As String) As Boolean
    SheetExists = Evaluate("ISREF('" & name & "'!A1)")
End Function

' ----------------------------------------------
'               Create sheets
' ----------------------------------------------

Private Sub CreateComponentsSheet()
    Dim i As Integer
    Dim row As Integer
    Dim global_row_end As Integer
    row = 7
    global_row_end = nr_students + 6
    
    For i = 0 To nr_comps - 1
        Dim comp_name As String
        Dim comp_weight As Integer
        Dim group As String
        Dim params_count As Integer
        Dim comp_flag As Integer
        
        Sheets("Aux").Activate
        
        comp_name = range("A" & row).Value
        comp_weight = range("B" & row).Value
        group = range("C" & row).Value
        params_count = range("D" & row).Value
        comp_flag = range("F" & row).Value
        
        Sheets.Add.name = comp_name
        Worksheets(comp_name).Move After:=Worksheets(1 + i)
        Sheets(comp_name).Activate
        
        Dim j As Integer
        Dim row_end As Integer
        
        If StrComp(group, "Sem grupo") = 0 Then
            range("A2").Value = "Estudantes"
            row_end = nr_students + 2
            range("A3:A" & row_end).Value = Sheets("Global").range("A7:A" & global_row_end).Value
            Columns("A").ColumnWidth = 20
        Else
            range("A2").Value = "Grupos"
            row_end = nr_students + 2
            Columns("A").ColumnWidth = 15
        End If
        
        For j = 1 To params_count
            'Parameter names
            range("A2").Offset(0, j).Value = Sheets("Aux").range("A" & row).Offset(j, 0).Value
            'Parameter Weights
            range("A1").Offset(0, j).Value = Sheets("Aux").range("E" & row).Offset(j, 0).Value
        Next j
        
        'Sum parameters
        Cells(1, params_count + 2).FormulaR1C1 = "=SUM(R1C2:R1C" & (params_count + 1) & ")"
        range("A2").Offset(0, params_count + 1).Value = "Total"
        
        'Sum input values
        range(Cells(3, params_count + 2), Cells(row_end, params_count + 2)).Select
        With Selection
            .FormulaR1C1 = "=SUM(RC2:RC" & (params_count + 1) & ")"
            .Interior.Color = RGB(255, 230, 153)
        End With
        
        'Format parameter weights
        range(Cells(1, 2), Cells(2, params_count + 2)).Select
        With Selection
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ColumnWidth = 10
        End With
        
        'Format all other cells
        range(Cells(2, 1), Cells(2, params_count + 2)).Interior.Color = RGB(217, 225, 242)
        range(Cells(2, 1), Cells(2, params_count + 2)).WrapText = True
        range(Cells(2, 1), Cells(row_end, params_count + 2)).Select
        With Selection
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        'Conditional Formating
        range(Cells(3, 2), Cells(row_end, params_count + 1)).Select
        With Selection
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).Font.Bold = True
            .FormatConditions(1).Interior.Color = RGB(128, 128, 128)
            .FormatConditions(1).StopIfTrue = False
            
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=B$1"
            .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).Font.Bold = True
            .FormatConditions(1).Interior.Color = RGB(128, 128, 128)
            .FormatConditions(1).StopIfTrue = False
        End With
        
        row = row + Sheets("Aux").range("D" & row).Value + 1
        range("A1").Select
        
        If comp_flag = 1 Then
            Sheets(comp_name).Select
            Sheets(comp_name).Copy Before:=Sheets(Sheets.count - 1)
            Sheets(Sheets.count - 2).name = comp_name & " Recurso"
        End If
    Next i
End Sub

Private Sub CreateGroupsSheet()
    Sheets.Add.name = "Grupos"
    Worksheets("Grupos").Move After:=Worksheets(1)
    Sheets("Grupos").Activate

    Dim i As Integer
    Dim row_end As Integer
    Dim global_row_end As Integer
    Dim comp_row As Integer
    Dim group As String
    
    row_end = nr_students + 1
    global_row_end = nr_students + 6
    comp_row = 7

    range("A1").Value = "Estudantes"
    range("A2:A" & row_end).Value = Sheets("Global").range("A7:A" & global_row_end).Value
    
    For i = 0 To nr_groups - 1
        With GroupList
            range("B1").Offset(0, i).Value = .List(i)
        End With
    Next i
    
    range(Cells(1, 1), Cells(1, nr_groups + 1)).Select
    With Selection
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ColumnWidth = 15
        .Interior.Color = RGB(217, 225, 242)
        .WrapText = True
    End With
    
    range(Cells(2, 1), Cells(row_end, nr_groups + 1)).Select
    With Selection
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Columns("A").ColumnWidth = 20
    range("A1").Select
End Sub

Private Sub UpdateGlobalSheet()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim z As Integer
    Dim aux1 As Integer
    Dim aux2 As Integer
    Dim aux_row As Integer
    Dim row_end As Integer
    Dim col_end As Integer
    Dim formula_str As String
    Dim formula_str2 As String
    Dim count_extra As Integer
    Dim extra_comps() As Boolean
    ReDim extra_comps(nr_comps)

    j = 2
    k = 1
    z = 1
    aux_row = 7
    row_end = nr_students + 6
    If nr_extra > 0 Then
        col_end = nr_comps + nr_extra + 5
    Else
        col_end = nr_comps + 3
    End If
    formula_str = ""
    formula_str2 = ""
    count_extra = 1

    For i = 0 To nr_comps - 1
        Dim comp_name As String
        Dim comp_weight As Integer
        Dim group As String
        Dim params_count As Integer
        Dim comp_flag As Integer
        Dim parameter_total As Integer
        Dim vlookup_extra As String
        Dim aux_formula1 As String
        Dim aux_formula2 As String
        
        Sheets("Aux").Activate
        
        comp_name = range("A" & aux_row).Value
        comp_weight = range("B" & aux_row).Value
        group = range("C" & aux_row).Value
        params_count = range("D" & aux_row).Value
        comp_flag = range("F" & aux_row).Value
        parameter_total = Sheets(comp_name).Cells(1, params_count + 2).Value
        
        Sheets("Global").Activate
        
        Cells(5, j).Value = comp_weight
        Cells(6, j).Value = comp_name
        
        If StrComp(group, "Sem grupo") = 0 Then
            range(Cells(7, j), Cells(row_end, j)).Select
            With Selection
                .FormulaR1C1 = "=VLOOKUP(RC1," & Chr(39) & comp_name & Chr(39) & "!R3C1:R" & (nr_students + 2) & "C" _
                & (params_count + 2) & "," & (params_count + 2) & ",FALSE) / " & parameter_total & "*" & comp_weight
                .NumberFormat = "0.00"
                .Interior.Color = RGB(226, 239, 218)
            End With
            
            If comp_flag = 1 Then
                Cells(5, nr_comps + 3 + k).Value = comp_weight
                Cells(6, nr_comps + 3 + k).Value = comp_name & " Recurso"
                vlookup_extra = "VLOOKUP(RC1," & Chr(39) & comp_name & " Recurso" & Chr(39) & "!R3C1:R" & (nr_students + 2) & "C" _
                    & (params_count + 2) & "," & (params_count + 2) & ",FALSE) / " & parameter_total & "*" & comp_weight
                
                range(Cells(7, nr_comps + 3 + k), Cells(row_end, nr_comps + 3 + k)).Select
                With Selection
                    .FormulaR1C1 = "=IF(" & vlookup_extra & " > 0, " & vlookup_extra & ", " & Chr(34) & Chr(34) & ")"
                    .NumberFormat = "0.00"
                    .Interior.Color = RGB(226, 239, 218)
                End With
                
                extra_comps(z) = True
                k = k + 1
            Else
                extra_comps(z) = False
            End If
            
            z = z + 1
        Else
            aux_formula1 = "MATCH(" & Chr(34) & group & Chr(34) & "," & Chr(39) & "Grupos" & Chr(39) & "!R1C1:R1C" _
                & (nr_groups + 1) & ",0)"
                
            aux_formula2 = "VLOOKUP(RC1," & Chr(39) & "Grupos" & Chr(39) & "!R1C1:R" & (nr_students + 1) & "C" _
                & (nr_groups + 1) & "," & aux_formula1 & ",FALSE)"
            
            range(Cells(7, j), Cells(row_end, j)).Select
            With Selection
                .FormulaR1C1 = "=IFNA(VLOOKUP(" & aux_formula2 & "," & Chr(39) & comp_name & Chr(39) & "!R3C1:R" & (nr_students + 2) & "C" _
                & (params_count + 2) & "," & (params_count + 2) & ",FALSE) / " & parameter_total & "*" & comp_weight & "," & Chr(34) _
                & "Faltam grupos" & Chr(34) & ")"
                .NumberFormat = "0.00"
                .Interior.Color = RGB(226, 239, 218)
            End With
            
            If comp_flag = 1 Then
                Cells(5, nr_comps + 3 + k).Value = comp_weight
                Cells(6, nr_comps + 3 + k).Value = comp_name & " Recurso"
                
                vlookup_extra = "IFNA(VLOOKUP(" & aux_formula2 & "," & Chr(39) & comp_name & " Recurso" & Chr(39) & "!R3C1:R" & (nr_students + 2) _
                & "C" & (params_count + 2) & "," & (params_count + 2) & ",FALSE) / " & parameter_total & "*" & comp_weight & "," & Chr(34) _
                & "Faltam grupos" & Chr(34) & ")"
                
                range(Cells(7, nr_comps + 3 + k), Cells(row_end, nr_comps + 3 + k)).Select
                With Selection
                    MsgBox vlookup_extra
                    .FormulaR1C1 = "=IF(" & vlookup_extra & " > 0, " & vlookup_extra & ", " & Chr(34) & Chr(34) & ")"
                    .NumberFormat = "0.00"
                    .Interior.Color = RGB(226, 239, 218)
                End With
                
                extra_comps(z) = True
                k = k + 1
            Else
                extra_comps(z) = False
            End If
            
            z = z + 1
        End If
        
        j = j + 1
        aux_row = aux_row + params_count + 1
    Next i
    
    Sheets("Global").Activate
    
    'Create Total/Sum column
    Cells(5, nr_comps + 2).FormulaR1C1 = "=SUM(R5C2:R5C" & (nr_comps + 1) & ")"
    Cells(6, nr_comps + 2).Value = "Total"
    Cells(5, nr_comps + 3).Value = Cells(5, nr_comps + 2).Value
    Cells(6, nr_comps + 3).Value = "Nota"
    
    grade1_idx = nr_comps + 3
    grade2_idx = -1
    
    range(Cells(7, nr_comps + 2), Cells(row_end, nr_comps + 2)).Select
    With Selection
        .FormulaR1C1 = "=SUM(RC2:RC" & (nr_comps + 1) & ")"
        .NumberFormat = "0.00"
        .Interior.Color = RGB(255, 230, 153)
    End With
    
    aux1 = Cells(5, nr_comps + 2).Value
    aux2 = Round(aux1 / 2)
    
    range(Cells(7, nr_comps + 3), Cells(row_end, nr_comps + 3)).Select
    With Selection
        .FormatConditions.AddTop10
        .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).TopBottom = xlTop10Top
        .FormatConditions(1).Rank = 1
        .FormatConditions(1).Percent = False
        .FormatConditions(1).Font.Color = RGB(48, 84, 150)
        .FormatConditions(1).Font.Bold = True
        .FormatConditions(1).Interior.Color = RGB(142, 169, 219)
        .FormatConditions(1).StopIfTrue = False
        
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & aux2
        .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).Font.Color = RGB(204, 0, 0)
        .FormatConditions(1).Font.Bold = True
        .FormatConditions(1).Interior.Color = RGB(240, 128, 128)
        .FormatConditions(1).StopIfTrue = False
        
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).Font.Color = RGB(0, 0, 0)
        .FormatConditions(1).Font.Bold = True
        .FormatConditions(1).Interior.Color = RGB(128, 128, 128)
        .FormatConditions(1).StopIfTrue = False
                    
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & aux1
        .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).Font.Color = RGB(0, 0, 0)
        .FormatConditions(1).Font.Bold = True
        .FormatConditions(1).Interior.Color = RGB(128, 128, 128)
        .FormatConditions(1).StopIfTrue = False
        
        .FormatConditions.Add Type:=xlBlanksCondition
        .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).Font.Color = RGB(0, 0, 0)
        .FormatConditions(1).Interior.Color = RGB(255, 255, 255)
        .FormatConditions(1).StopIfTrue = False
    End With
    
    If nr_extra > 0 Then
        Cells(5, nr_comps + nr_extra + 4).Value = Cells(5, nr_comps + 2).Value
        Cells(6, nr_comps + nr_extra + 4).Value = "Total"
        Cells(5, nr_comps + nr_extra + 5).Value = Cells(5, nr_comps + 2).Value
        Cells(6, nr_comps + nr_extra + 5).Value = "Nota"
        
        grade2_idx = nr_comps + nr_extra + 5
        
        k = 1 + nr_comps + 2 + 1
        
        For i = 1 To nr_comps
            If extra_comps(i) = True Then
                If i = nr_comps Then
                    formula_str = formula_str & "IF(EXACT(RC" & k & "," & Chr(34) & Chr(34) & "),0,RC" & k & ")"
                Else
                    formula_str = formula_str & "IF(EXACT(RC" & k & "," & Chr(34) & Chr(34) & "),0,RC" & k & ")" & " + "
                End If
                
                If count_extra = nr_extra Then
                    formula_str2 = formula_str2 & "IF(EXACT(RC" & k & "," & Chr(34) & Chr(34) & "),0,RC" & k & ")"
                Else
                    formula_str2 = formula_str2 & "IF(EXACT(RC" & k & "," & Chr(34) & Chr(34) & "),0,RC" & k & ")" & " + "
                End If
                
                k = k + 1
                count_extra = count_extra + 1
            Else
                If i = nr_comps Then
                    formula_str = formula_str & "RC" & (i + 1)
                Else
                    formula_str = formula_str & "RC" & (i + 1) & " + "
                End If
            End If
        Next i
        
        range(Cells(7, nr_comps + nr_extra + 4), Cells(row_end, nr_comps + nr_extra + 4)).Select
        With Selection
            .FormulaR1C1 = "=IF(" & formula_str2 & " > 0, " & formula_str & ", " & Chr(34) & Chr(34) & ")"
            .NumberFormat = "0.00"
            .Interior.Color = RGB(255, 230, 153)
        End With
        
        aux1 = Cells(5, nr_comps + 2).Value
        aux2 = Round(aux1 / 2)
        
        range(Cells(7, col_end), Cells(row_end, col_end)).Select
        With Selection
            .FormatConditions.AddTop10
            .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).TopBottom = xlTop10Top
            .FormatConditions(1).Rank = 1
            .FormatConditions(1).Percent = False
            .FormatConditions(1).Font.Color = RGB(48, 84, 150)
            .FormatConditions(1).Font.Bold = True
            .FormatConditions(1).Interior.Color = RGB(142, 169, 219)
            .FormatConditions(1).StopIfTrue = False
            
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & aux2
            .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).Font.Color = RGB(204, 0, 0)
            .FormatConditions(1).Font.Bold = True
            .FormatConditions(1).Interior.Color = RGB(240, 128, 128)
            .FormatConditions(1).StopIfTrue = False
        
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).Font.Color = RGB(0, 0, 0)
            .FormatConditions(1).Font.Bold = True
            .FormatConditions(1).Interior.Color = RGB(128, 128, 128)
            .FormatConditions(1).StopIfTrue = False
                    
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & aux1
            .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).Font.Color = RGB(0, 0, 0)
            .FormatConditions(1).Font.Bold = True
            .FormatConditions(1).Interior.Color = RGB(128, 128, 128)
            .FormatConditions(1).StopIfTrue = False
            
            .FormatConditions.Add Type:=xlBlanksCondition
            .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).Font.Color = RGB(0, 0, 0)
            .FormatConditions(1).Interior.Color = RGB(255, 255, 255)
            .FormatConditions(1).StopIfTrue = False
        End With
    End If
    
    Columns("A").ColumnWidth = 20
    range(Cells(5, 2), Cells(5, col_end)).Select
    With Selection
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ColumnWidth = 12
    End With

    range(Cells(6, 1), Cells(6, col_end)).Interior.Color = RGB(217, 225, 242)
    range(Cells(6, 1), Cells(6, col_end)).WrapText = True
    
    range(Cells(6, 1), Cells(row_end, col_end)).Select
    With Selection
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub CreateSynthesisSheet()
    Sheets.Add.name = "Sintese"
    Worksheets("Sintese").Move Before:=Worksheets(1)
    Sheets("Sintese").Activate
    
    Dim global_row_start As Integer
    Dim global_row_end As Integer
    
    global_row_start = 7
    global_row_end = nr_students + 6
    
    Columns("A").ColumnWidth = 2.56
    Columns("B:C").ColumnWidth = 8.11
    Columns("D").ColumnWidth = 2.56
    Columns("E:K").ColumnWidth = 8.11
    Columns("L").ColumnWidth = 2.56
    
    Call CreateSynthesisHeader(False)
    Call CreateSynthesisFreqStats(global_row_start, global_row_end, False)
    Call CreateSynthesisSearch(global_row_start, global_row_end, False)
    Call CreateSynthesisExtraStats(global_row_start, global_row_end, False)
    Call CreateChart(False)
    
    If (nr_extra > 0) Then
    
        Columns("M").ColumnWidth = 5
        Columns("N").ColumnWidth = 2.56
        Columns("O:P").ColumnWidth = 8.11
        Columns("Q").ColumnWidth = 2.56
        Columns("R:X").ColumnWidth = 8.11
        Columns("Y").ColumnWidth = 2.56
        
        CreateMiddleSeparator
        Call CreateSynthesisHeader(True)
        Call CreateSynthesisFreqStats(global_row_start, global_row_end, True)
        Call CreateSynthesisSearch(global_row_start, global_row_end, True)
        Call CreateSynthesisExtraStats(global_row_start, global_row_end, True)
        Call CreateChart(True)
    
    End If
    
    range("A1").Select
End Sub

Private Sub CreateSynthesisHeader(ByVal extra As Boolean)
    Sheets("Sintese").Activate
    
    If extra = False Then
        range("A2").Value = "Época Normal"
        range("A2:L3").Select
        With Selection
            .Font.Color = RGB(48, 84, 150)
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(180, 198, 231)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Merge
        End With
    Else
        range("N2").Value = "Época Recurso"
        range("N2:Y3").Select
        With Selection
            .Font.Color = RGB(48, 84, 150)
            .Font.Size = 12
            .Font.Bold = True
            .Interior.Color = RGB(180, 198, 231)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Merge
        End With
    End If
End Sub

Private Sub CreateSynthesisSearch(ByVal r1 As Integer, ByVal r2 As Integer, ByVal extra As Boolean)
    Dim total_count As Integer
    total_count = Sheets("Global").Cells(5, grade1_idx).Value
    
    Sheets("Sintese").Activate
    Dim row_start As Integer
    Dim row_end As Integer
    
    row_start = 11
    row_end = row_start + total_count
    
    If extra = False Then
        range("B5:C5").Interior.Color = RGB(217, 225, 242)
        range("B5:C5").Merge
        range("B5").Value = "Pesquisa"
        
        range("B6").Value = "Nota"
        range("B6").Interior.Color = RGB(237, 237, 237)
        range("C6").Value = 0
        
        range("B7").Value = "Freq"
        range("B7").Interior.Color = RGB(237, 237, 237)
        
        range("C7").FormulaR1C1 = "=SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 & "C" & grade1_idx _
            & ":R" & r2 & "C" & grade1_idx & "=R6C3)) - SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade1_idx & ":R" & r2 & "C" & grade1_idx & "=" & Chr(34) & Chr(34) & "))"
        range("C7").Interior.Color = RGB(255, 230, 153)
        
        range("B5:C7").HorizontalAlignment = xlCenter
        range("B5:C7").VerticalAlignment = xlCenter
        range("B5:C7").Borders.LineStyle = xlContinuous
    Else
        range("O5:P5").Interior.Color = RGB(217, 225, 242)
        range("O5:P5").Merge
        range("O5").Value = "Pesquisa"
    
        range("O6").Value = "Nota"
        range("O6").Interior.Color = RGB(237, 237, 237)
        range("P6").Value = 0
        
        range("O7").Value = "Freq"
        range("O7").Interior.Color = RGB(237, 237, 237)
        
        range("P7").FormulaR1C1 = "=VLOOKUP(R6C16,R" & row_start & "C15:R" & row_end & "C16,2,FALSE)"
        range("P7").Interior.Color = RGB(255, 230, 153)
        
        range("O5:P7").HorizontalAlignment = xlCenter
        range("O5:P7").VerticalAlignment = xlCenter
        range("O5:P7").Borders.LineStyle = xlContinuous
    End If
End Sub

Private Sub CreateSynthesisExtraStats(ByVal r1 As Integer, ByVal r2 As Integer, ByVal extra As Boolean)
    Sheets("Sintese").Activate
    
    If extra = False Then
        range("E21:F21").Interior.Color = RGB(217, 225, 242)
        range("E21:F21").Merge
        range("E21").Value = "#Participantes"
        
        range("G21").FormulaR1C1 = "=COUNTA(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade1_idx & ":R" & r2 & "C" & grade1_idx & ")"
        range("G21").Interior.Color = RGB(255, 230, 153)
        
        range("E21:G21").HorizontalAlignment = xlCenter
        range("E21:G21").VerticalAlignment = xlCenter
        range("E21:G21").Borders.LineStyle = xlContinuous
        
        range("E23:F23").Interior.Color = RGB(217, 225, 242)
        range("E23:F23").Merge
        range("E23").Value = "#Aprovados"
        
        range("G23").FormulaR1C1 = "=SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade1_idx & ":R" & r2 & "C" & grade1_idx & ">=ROUND(" & Chr(39) & "Global" _
            & Chr(39) & "!R" & (r1 - 2) & "C" & grade1_idx & "/2,0)))"
        range("G23").Interior.Color = RGB(255, 230, 153)
        
        range("E23:G23").HorizontalAlignment = xlCenter
        range("E23:G23").VerticalAlignment = xlCenter
        range("E23:G23").Borders.LineStyle = xlContinuous
        
        range("E24:F24").Interior.Color = RGB(217, 225, 242)
        range("E24:F24").Merge
        range("E24").Value = "#Reprovados"
        
        range("G24").FormulaR1C1 = "=SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade1_idx & ":R" & r2 & "C" & grade1_idx & "<ROUND(" & Chr(39) & "Global" _
            & Chr(39) & "!R" & (r1 - 2) & "C" & grade1_idx & "/2,0))) - SUMPRODUCT(--(" & Chr(39) _
            & "Global" & Chr(39) & "!R" & r1 & "C" & grade1_idx & ":R" & r2 & "C" & grade1_idx & "=" & Chr(34) _
            & Chr(34) & "))"
        range("G24").Interior.Color = RGB(255, 230, 153)
        
        range("E24:G24").HorizontalAlignment = xlCenter
        range("E24:G24").VerticalAlignment = xlCenter
        range("E24:G24").Borders.LineStyle = xlContinuous
        
        range("I21:J21").Interior.Color = RGB(217, 225, 242)
        range("I21:J21").Merge
        range("I21").Value = "Média"

        range("K21").FormulaR1C1 = "=AVERAGE(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade1_idx & ":R" & r2 & "C" & grade1_idx & ")"
        range("K21").Interior.Color = RGB(255, 230, 153)

        range("I21:K21").HorizontalAlignment = xlCenter
        range("I21:K21").VerticalAlignment = xlCenter
        range("I21:K21").Borders.LineStyle = xlContinuous
        
        range("I23:J23").Interior.Color = RGB(217, 225, 242)
        range("I23:J23").Merge
        range("I23").Value = "Nota mais alta"
        
        range("K23").FormulaR1C1 = "=MAX(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade1_idx & ":R" & r2 & "C" & grade1_idx & ")"
        range("K23").Interior.Color = RGB(255, 230, 153)
        
        range("I23:K24").HorizontalAlignment = xlCenter
        range("I23:K24").VerticalAlignment = xlCenter
        range("I23:K24").Borders.LineStyle = xlContinuous
        
        range("I24:J24").Interior.Color = RGB(217, 225, 242)
        range("I24:J24").Merge
        range("I24").Value = "Nota mais baixa"
        
        range("K24").FormulaR1C1 = "=MIN(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade1_idx & ":R" & r2 & "C" & grade1_idx & ")"
        range("K24").Interior.Color = RGB(255, 230, 153)
        
        range("I24:K24").HorizontalAlignment = xlCenter
        range("I24:K24").VerticalAlignment = xlCenter
        range("I24:K24").Borders.LineStyle = xlContinuous
        
    Else
        range("R21:S21").Interior.Color = RGB(217, 225, 242)
        range("R21:S21").Merge
        range("R21").Value = "#Participantes"
        
        range("T21").FormulaR1C1 = "=COUNTA(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade2_idx & ":R" & r2 & "C" & grade2_idx & ")"
        range("T21").Interior.Color = RGB(255, 230, 153)
        
        range("R21:T21").HorizontalAlignment = xlCenter
        range("R21:T21").VerticalAlignment = xlCenter
        range("R21:T21").Borders.LineStyle = xlContinuous
        
        range("R23:S23").Interior.Color = RGB(217, 225, 242)
        range("R23:S23").Merge
        range("R23").Value = "#Aprovados"
        
        range("T23").FormulaR1C1 = "=SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade2_idx & ":R" & r2 & "C" & grade2_idx & ">=ROUND(" & Chr(39) & "Global" _
            & Chr(39) & "!R" & (r1 - 2) & "C" & grade2_idx & "/2,0)))"
        range("T23").Interior.Color = RGB(255, 230, 153)
        
        range("R23:T23").HorizontalAlignment = xlCenter
        range("R23:T23").VerticalAlignment = xlCenter
        range("R23:T23").Borders.LineStyle = xlContinuous
        
        range("R24:S24").Interior.Color = RGB(217, 225, 242)
        range("R24:S24").Merge
        range("R24").Value = "#Reprovados"
        
        range("T24").FormulaR1C1 = "=SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade2_idx & ":R" & r2 & "C" & grade2_idx & "<ROUND(" & Chr(39) & "Global" _
            & Chr(39) & "!R" & (r1 - 2) & "C" & grade2_idx & "/2,0))) - SUMPRODUCT(--(" & Chr(39) _
            & "Global" & Chr(39) & "!R" & r1 & "C" & grade2_idx & ":R" & r2 & "C" & grade2_idx _
            & "=" & Chr(34) & Chr(34) & "))"
        range("T24").Interior.Color = RGB(255, 230, 153)
        
        range("R24:T24").HorizontalAlignment = xlCenter
        range("R24:T24").VerticalAlignment = xlCenter
        range("R24:T24").Borders.LineStyle = xlContinuous

        range("V21:W21").Interior.Color = RGB(217, 225, 242)
        range("V21:W21").Merge
        range("V21").Value = "Média"

        range("X21").FormulaR1C1 = "=AVERAGE(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade2_idx & ":R" & r2 & "C" & grade2_idx & ")"
        range("X21").Interior.Color = RGB(255, 230, 153)

        range("V21:X21").HorizontalAlignment = xlCenter
        range("V21:X21").VerticalAlignment = xlCenter
        range("V21:X21").Borders.LineStyle = xlContinuous
        
        range("V23:W23").Interior.Color = RGB(217, 225, 242)
        range("V23:W23").Merge
        range("V23").Value = "Nota mais alta"
        
        range("X23").FormulaR1C1 = "=MAX(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade2_idx & ":R" & r2 & "C" & grade2_idx & ")"
        range("X23").Interior.Color = RGB(255, 230, 153)
        
        range("V23:X23").HorizontalAlignment = xlCenter
        range("V23:X23").VerticalAlignment = xlCenter
        range("V23:X23").Borders.LineStyle = xlContinuous
        
        range("V24:W24").Interior.Color = RGB(217, 225, 242)
        range("V24:W24").Merge
        range("V24").Value = "Nota mais baixa"
        
        range("X24").FormulaR1C1 = "=MIN(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 _
            & "C" & grade2_idx & ":R" & r2 & "C" & grade2_idx & ")"
        range("X24").Interior.Color = RGB(255, 230, 153)
        
        range("V24:X24").HorizontalAlignment = xlCenter
        range("V24:X24").VerticalAlignment = xlCenter
        range("V24:X24").Borders.LineStyle = xlContinuous
        
    End If
End Sub

Private Sub CreateSynthesisFreqStats(ByVal r1 As Integer, ByVal r2 As Integer, ByVal extra As Boolean)
    Dim total_count As Integer
    total_count = Sheets("Global").Cells(5, grade1_idx).Value
    
    Sheets("Sintese").Activate
    Dim row_start As Integer
    Dim row_end As Integer
    Dim i As Integer
    Dim count_i As Integer
    
    row_start = 11
    row_end = row_start + total_count
    count_i = 0
    
    If extra = False Then
        range("B9:C9").Interior.Color = RGB(217, 225, 242)
        range("B9:C9").Merge
        range("B9").Value = "Estatísticas"
        
        range("B10").Value = "Nota"
        range("C10").Value = "Freq"
        range("B10:C10").Interior.Color = RGB(237, 237, 237)
        
        For i = 0 To total_count
            range("B" & (row_start + i)).Value = count_i
            range("C" & (row_start + i)).FormulaR1C1 = "=SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 & "C" & grade1_idx _
                & ":R" & r2 & "C" & grade1_idx & "=R" & (row_start + i) & "C2))"
            count_i = count_i + 1
        Next i
        
        range("C11").FormulaR1C1 = "=SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 & "C" & grade1_idx _
                & ":R" & r2 & "C" & grade1_idx & "=R" & (row_start + i) & "C2)) - SUMPRODUCT(--(" & Chr(39) & "Global" _
                & Chr(39) & "!R" & r1 & "C" & grade1_idx & ":R" & r2 & "C" & grade1_idx & "=" & Chr(34) & Chr(34) & "))"
        
        range("B9:C" & row_end).Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
    Else
        range("O9:P9").Interior.Color = RGB(217, 225, 242)
        range("O9:P9").Merge
        range("O9").Value = "Estatísticas"
        
        range("O10").Value = "Nota"
        range("P10").Value = "Freq"
        range("O10:P10").Interior.Color = RGB(237, 237, 237)
        
        For i = 0 To total_count
            range("O" & (row_start + i)).Value = count_i
            range("P" & (row_start + i)).FormulaR1C1 = "=SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 & "C" & grade2_idx _
                & ":R" & r2 & "C" & grade2_idx & "=R" & (row_start + i) & "C15))"
            count_i = count_i + 1
        Next i
        
        range("P11").FormulaR1C1 = "=SUMPRODUCT(--(" & Chr(39) & "Global" & Chr(39) & "!R" & r1 & "C" & grade2_idx _
                & ":R" & r2 & "C" & grade2_idx & "=R" & (row_start + i) & "C15)) - SUMPRODUCT(--(" & Chr(39) & "Global" _
                & Chr(39) & "!R" & r1 & "C" & grade2_idx & ":R" & r2 & "C" & grade2_idx & "=" & Chr(34) & Chr(34) & "))"
        
        range("O9:P" & row_end).Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
    End If
    
End Sub

Private Sub CreateMiddleSeparator()
    Dim total_count As Integer
    total_count = Sheets("Global").Cells(5, grade1_idx).Value
    
    Sheets("Sintese").Activate
    Dim row_start As Integer
    Dim row_end As Integer
    
    row_start = 2
    row_end = row_start + 9 + total_count
    
    range("M" & row_start & ":M" & row_end).Merge
End Sub

Private Sub CreateChart(ByVal extra As Boolean)
    Dim total_count As Integer
    total_count = Sheets("Global").Cells(5, grade1_idx).Value
    
    Sheets("Sintese").Activate

    Dim row_start As Integer
    Dim row_end As Integer
    
    row_start = 11
    row_end = row_start + total_count
    
    If extra = False Then
        ActiveSheet.Shapes.AddChart(201, xlColumnClustered).Select
        With ActiveChart
            .SetSourceData Source:=range("Sintese!C" & row_start & ":C" & row_end)
            .SeriesCollection(1).XValues = range("Sintese!B" & row_start & ":B" & row_end)
            .SeriesCollection(1).Interior.Color = RGB(180, 198, 231)
            .Parent.Left = range("E5").Left
            .Parent.Top = range("E5").Top
            .Parent.Width = range("E5:K5").Width
            .Parent.Height = range("E5:E19").Height
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Nota"
            .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
            .Axes(xlCategory, xlPrimary).AxisTitle.Font.Bold = False
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Frequência"
            .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
            .Axes(xlValue, xlPrimary).AxisTitle.Font.Bold = False
            .HasLegend = False
        End With
    Else
        ActiveSheet.Shapes.AddChart(201, xlColumnClustered).Select
        With ActiveChart
            .SetSourceData Source:=range("Sintese!P" & row_start & ":P" & row_end)
            .SeriesCollection(1).XValues = range("Sintese!O" & row_start & ":O" & row_end)
            .SeriesCollection(1).Interior.Color = RGB(180, 198, 231)
            .Parent.Left = range("R5").Left
            .Parent.Top = range("R5").Top
            .Parent.Width = range("R5:X5").Width
            .Parent.Height = range("R5:R19").Height
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Nota"
            .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
            .Axes(xlCategory, xlPrimary).AxisTitle.Font.Bold = False
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Frequência"
            .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
            .Axes(xlValue, xlPrimary).AxisTitle.Font.Bold = False
            .HasLegend = False
        End With
    End If
End Sub
