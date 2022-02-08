VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} c_e_s 
   Caption         =   "Causas & Soluções"
   ClientHeight    =   8505.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15240
   OleObjectBlob   =   "c_e_s.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "c_e_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub atualizar_1_Click()

aux_os = options.os.Value

Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 9).Value = pq.Value

End Sub

Private Sub atualizar_2_Click()

aux_os = options.os.Value

Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 10).Value = c_pq.Value

End Sub

Private Sub atualizar_3_Click()

aux_os = options.os.Value

Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 11).Value = solucao.Value

End Sub

Private Sub botao_x_Click()

Unload c_e_s
Unload options
Unload consulta


End Sub

Private Sub UserForm_Initialize()

aux_os = options.os.Value

os.Value = aux_os
nome.Value = options.nome.Value

On Error Resume Next

status_prev.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 5).Value
equip.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 8).Value
data_chamado.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 2).Value
horimetro.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, -2).Value
d_falha.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, -1).Value
pq.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 9).Value
c_pq.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 10).Value
solucao.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 11).Value


linha = Sheets("VALIDAÇÃO").Range("G1000000").End(xlUp).Row
c_pq.RowSource = "VALIDAÇÃO!G2:G" & linha


End Sub
