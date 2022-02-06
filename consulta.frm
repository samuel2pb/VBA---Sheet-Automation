VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} consulta 
   Caption         =   "Atendimentos em Andamento"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11910
   OleObjectBlob   =   "consulta.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub atualizar_Click()

options.Show

End Sub

Private Sub botao_x_Click()

Unload consulta

End Sub

Private Sub MultiPage1_Change()

If MultiPage1.Pages(0).Enabled = True Then

    remoto.Value = ""
    presencial.Value = ""
    finalizados.Value = ""
    

ElseIf MultiPage1.Pages(1).Enabled = True Then
    
    remoto.Value = ""
    presencial.Value = ""
    finalizados.Value = ""
    
ElseIf MultiPage1.Pages(2).Enabled = True Then

    remoto.Value = ""
    presencial.Value = ""
    finalizados.Value = ""
    
End If


End Sub

Private Sub UserForm_Initialize()

Dim cell As Range
Dim r As Range
Dim i As Integer
Dim j As Integer
Dim k As Integer

i = 0
j = 0
k = 0


Dim tab1() As String
ReDim tab1(i, 4)

Dim tab2() As String
ReDim tab2(j, 4)

Dim tab3() As String
ReDim tab3(k, 4)


remoto.Clear
presencial.Clear
finalizados.Clear

Set r = Sheets("GERAL").Range("L3").End(xlDown)

'Sheets("receber").Unprotect Password:="123"


For Each cell In Sheets("GERAL").Range("L3", r)

    
    nserie = cell.Offset(0, -11).Value
    os = cell.Offset(0, -5).Value
    cliente = cell.Offset(0, -10).Value
    data_chamado = cell.Offset(0, -3).Value
    equip = cell.Offset(0, 3).Value
     
    If cell.Value = "EM ATENDIMENTO REMOTO" Then
    
    tab1(i, 0) = os
    tab1(i, 1) = cliente
    tab1(i, 2) = equip
    tab1(i, 3) = nserie
    tab1(i, 4) = data_chamado
    
        With remoto
            .AddItem
            .List(.ListCount - 1, 0) = (tab1(i, 0))
            .List(.ListCount - 1, 1) = (tab1(i, 1))
            .List(.ListCount - 1, 2) = (tab1(i, 2))
            .List(.ListCount - 1, 3) = (tab1(i, 3))
            .List(.ListCount - 1, 4) = (tab1(i, 4))
        End With
   
        i = i + 1
        ReDim tab1((i), (4))
    
     
    ElseIf cell.Value = "EM ATENDIMENTO PRESENCIAL" Then
    
    tab2(j, 0) = os
    tab2(j, 1) = cliente
    tab2(j, 2) = equip
    tab2(j, 3) = nserie
    tab2(j, 4) = data_chamado
        
        With presencial
            .AddItem
            .List(.ListCount - 1, 0) = (tab2(j, 0))
            .List(.ListCount - 1, 1) = (tab2(j, 1))
            .List(.ListCount - 1, 2) = (tab2(j, 2))
            .List(.ListCount - 1, 3) = (tab2(j, 3))
            .List(.ListCount - 1, 4) = (tab2(j, 4))
         
        End With
    
        j = j + 1
        ReDim tab2((j), (4))
    
    
    ElseIf cell.Value = "FINALIZADO REMOTO" Or cell.Value = "FINALIZADO PRESENCIAL" Then
    
    'And cell.Offset(0, 7) = ""
    
    tab3(k, 0) = os
    tab3(k, 1) = cliente
    tab3(k, 2) = equip
    tab3(k, 3) = nserie
    tab3(k, 4) = data_chamado
        
        With finalizados
            .AddItem
            .List(.ListCount - 1, 0) = (tab3(k, 0))
            .List(.ListCount - 1, 1) = (tab3(k, 1))
            .List(.ListCount - 1, 2) = (tab3(k, 2))
            .List(.ListCount - 1, 3) = (tab3(k, 3))
            .List(.ListCount - 1, 4) = (tab3(k, 4))
         
        End With
    
        k = k + 1
        ReDim tab3((k), (4))
   
    
    End If
    
Next cell

End Sub
