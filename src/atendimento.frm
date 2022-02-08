VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} atendimento 
   Caption         =   "Novo Atendimento"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   OleObjectBlob   =   "atendimento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "atendimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub botao_ok_Click()


'While nome.Value = "" Or nserie.Value = "" Or canal.Value = "" Or problema.Value = "" Or horimetro.Value = "" Or data_chamado.Value = ""

    'MsgBox ("Erro - Preencha todos os campos antes de finalizar o cadastro")

'Else

    'Sheets("clientes").Unprotect Password:="123"
    
    linha_vazia = Sheets("GERAL").Range("A1000000").End(xlUp).Row + 1
    
    'Sheets("GERAL").Cells(linha_vazia - 1, 1).EntireRow.Copy
    
    Sheets("GERAL").Cells(linha_vazia, 1).Value = nserie.Value
    Sheets("GERAL").Cells(linha_vazia, 1).Locked = True
    
    Sheets("GERAL").Cells(linha_vazia, 5).Value = horimetro.Value
    Sheets("GERAL").Cells(linha_vazia, 5).Locked = True
    
    Sheets("GERAL").Cells(linha_vazia, 7).Value = Sheets("GERAL").Cells(linha_vazia - 1, 7).Value + 1
    Sheets("GERAL").Cells(linha_vazia, 7).Locked = True
    
    If data_chamado = "" Then
    
        Sheets("GERAL").Cells(linha_vazia, 9).Value = Date
        Sheets("GERAL").Cells(linha_vazia, 9).Locked = True
    
    ElseIf data_chamado <> "" Then
    
        Sheets("GERAL").Cells(linha_vazia, 9).Value = data_chamado.Value
        Sheets("GERAL").Cells(linha_vazia, 9).Locked = True
        
    End If
    
    
    Sheets("GERAL").Cells(linha_vazia, 10).Value = Date
    Sheets("GERAL").Cells(linha_vazia, 10).Locked = True
    
    Sheets("GERAL").Cells(linha_vazia, 12).Value = "EM ATENDIMENTO REMOTO"
    Sheets("GERAL").Cells(linha_vazia, 12).Locked = True
    
    Sheets("GERAL").Cells(linha_vazia, 13).Value = canal.Value
    Sheets("GERAL").Cells(linha_vazia, 13).Locked = True
    
    Sheets("GERAL").Cells(linha_vazia, 14).Value = problema.Value
    Sheets("GERAL").Cells(linha_vazia, 14).Locked = True
    
    Sheets("GERAL").Cells(linha_vazia, 29).Value = atendente.Value
    Sheets("GERAL").Cells(linha_vazia, 29).Locked = True
    
    
    MsgBox ("Atendimento aberto com sucesso!")
    
    Unload atendimento



End Sub

Private Sub botao_x_Click()

Unload atendimento


End Sub

Private Sub nome_AfterUpdate()

Dim r As Range

nserie.Clear

Set r = Sheets("GARANTIA").Range("A2").End(xlDown)

For Each cell In Sheets("GARANTIA").Range("A2", r)

    If cell = nome.Value Then
        
        With nserie
            .AddItem (cell.Offset(0, 1))
        End With
    
    End If
    
Next cell


End Sub


Private Sub nserie_Change()

cidade = Sheets("GARANTIA").Cells.Find(nserie.Value, , xlValues, xlWhole).Offset(0, 4).Value
estado = Sheets("GARANTIA").Cells.Find(nserie.Value, , xlValues, xlWhole).Offset(0, 5).Value
equip = Sheets("GARANTIA").Cells.Find(nserie.Value, , xlValues, xlWhole).Offset(0, 2).Value
garantia = Sheets("GARANTIA").Cells.Find(nserie.Value, , xlValues, xlWhole).Offset(0, 11).Value
data_venda = Sheets("GARANTIA").Cells.Find(nserie.Value, , xlValues, xlWhole).Offset(0, 7).Value



End Sub



Private Sub UserForm_Initialize()

linha = Sheets("VALIDA플O").Range("J1000000").End(xlUp).Row
nome.RowSource = "VALIDA플O!J2:J" & linha

linha = Sheets("VALIDA플O").Range("C1000000").End(xlUp).Row
canal.RowSource = "VALIDA플O!C2:C" & linha

linha = Sheets("VALIDA플O").Range("F1000000").End(xlUp).Row
problema.RowSource = "VALIDA플O!F2:F" & linha

linha = Sheets("VALIDA플O").Range("K1000000").End(xlUp).Row
atendente.RowSource = "VALIDA플O!K2:K" & linha


End Sub

