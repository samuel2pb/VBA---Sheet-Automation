Attribute VB_Name = "Módulo1"
Dim pathx As Worksheet
Dim camx As Workbook
Dim pegx_path As String
Dim gatx_path As String

Sub novo_atendimento()

atendimento.Show

End Sub

Sub consultar_atendimentos()

consulta.Show

End Sub
Sub peças_em_garantia()

peg.Show

End Sub
Sub gastos_ass_tecnica()

gat.Show

End Sub

Sub Abrir_PEG()

Set camx = Workbooks.Open(ThisWorkbook.Path & "\" & "Local Paths.xlsx", 0, False, , , , True, , , , False, , False)
Set pathx = camx.Sheets("Caminhos")
pegx_path = pathx.Cells.Find("PEG", , xlValues, xlWhole).Offset(1, 0).Value

camx.Close (False)

Workbooks.Open pegx_path, 0, False, , , , True, , , , False, , False


End Sub


Sub Abrir_GAT()

Set camx = Workbooks.Open(ThisWorkbook.Path & "\" & "Local Paths.xlsx", 0, False, , , , True, , , , False, , False)
Set pathx = camx.Sheets("Caminhos")
gatx_path = pathx.Cells.Find("GAT", , xlValues, xlWhole).Offset(1, 0).Value

camx.Close (False)

Workbooks.Open gatx_path, 0, False, , , , True, , , , False, , False

End Sub



