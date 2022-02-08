VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} status_atendimento 
   Caption         =   "Atulização Status de Atendimento"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10440
   OleObjectBlob   =   "status_atendimento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "status_atendimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub atualizar_Click()

aux_os = options.os.Value

If atendente_presencial.Value <> "" Then
    
    Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 5).Value = status_atual.Value
    Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 23).Value = atendente_presencial.Value
    status_prev = status_atual

ElseIf atendente_presencial.Value = "" Then

    Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 5).Value = status_atual.Value
    status_prev = status_atual
    
End If

If status_atual.Value = "FINALIZADO REMOTO" Or status_atual.Value = "FINALIZADO PRESENCIAL" Then
    If data_final = "" Then
        Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 4).Value = Date
    ElseIf data_final <> "" Then
        Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 4).Value = data_final.Value
    End If
        
End If
                
End Sub

Private Sub botao_x_Click()

Unload status_atendimento
Unload options
Unload consulta


End Sub

Private Sub status_atual_AfterUpdate()

If status_atual.Value = "EM ATENDIMENTO PRESENCIAL" Then
    
    atendente_presencial.Locked = False
    
    linha = Sheets("VALIDAÇÃO").Range("K1000000").End(xlUp).Row
    atendente_presencial.RowSource = "VALIDAÇÃO!K2:K" & linha
    
ElseIf status_atual.Value <> "EM ATENDIMENTO PRESENCIAL" Then
    atendente_presencial.Locked = True
End If


End Sub


Private Sub UserForm_Initialize()

aux_os = options.os.Value

linha = Sheets("VALIDAÇÃO").Range("A1000000").End(xlUp).Row
status_atual.RowSource = "VALIDAÇÃO!A2:A" & linha

status_prev.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 5).Value


End Sub





