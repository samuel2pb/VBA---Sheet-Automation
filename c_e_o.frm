VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} c_e_o 
   Caption         =   "Classificações e Origens"
   ClientHeight    =   8655.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14580
   OleObjectBlob   =   "c_e_o.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "c_e_o"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub botao_ok_Click()

aux_os = os.Value

Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 12).Value = classificacao.Value
Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 13).Value = origem.Value
'Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 8).Value = sub_origem.Value

MsgBox ("Sucesso!")


End Sub

Private Sub botao_x_Click()

Unload c_e_o
Unload options
Unload consulta

End Sub


Private Sub UserForm_Initialize()

aux_os = options.os.Value

os.Value = aux_os
nome.Value = options.nome.Value

On Error Resume Next

equip.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 8).Value
data_chamado.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 2).Value
horimetro.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, -2).Value
d_falha.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, -1).Value
pq.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 9).Value
c_pq.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 10).Value
solucao.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 11).Value


classificacao.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 12).Value
origem.Value = Sheets("GERAL").Cells.Find(aux_os, , xlValues, xlWhole).Offset(0, 13).Value

linha = Sheets("VALIDAÇÃO").Range("H1000000").End(xlUp).Row
classificacao.RowSource = "VALIDAÇÃO!H2:H" & linha

linha = Sheets("VALIDAÇÃO").Range("I1000000").End(xlUp).Row
origem.RowSource = "VALIDAÇÃO!I2:I" & linha


End Sub
