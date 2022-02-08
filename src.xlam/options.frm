VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} options 
   Caption         =   "Opções de Atualização"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9015.001
   OleObjectBlob   =   "options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub andamento_Click()

status_atendimento.nome.Value = options.nome.Value
status_atendimento.os.Value = options.os.Value


status_atendimento.Show


End Sub

Private Sub botao_x_Click()

Unload options
Unload consulta

End Sub

Private Sub causas_Click()

c_e_s.Show

End Sub

Private Sub classes_Click()


c_e_o.nome.Value = options.nome.Value
c_e_o.nome.Locked = True
c_e_o.os.Value = options.os.Value
c_e_o.os.Locked = True

c_e_o.Show


End Sub

Private Sub custos_asstec_Click()

gat.nome.Value = options.nome.Value
gat.nome.Locked = True
gat.os.Value = options.os.Value
gat.os.Locked = True

gat.Show

End Sub

Private Sub pecas_em_garantia_Click()

peg.nome.Value = options.nome.Value
peg.nome.Locked = True
peg.os.Value = options.os.Value
peg.os.Locked = True

peg.Show


End Sub


Private Sub nome_Change()

Dim cell As Range
Dim r As Range
Dim i As Integer

os.Clear

Set r = Sheets("GERAL").Range("L3").End(xlDown)

For Each cell In Sheets("GERAL").Range("L3", r)

    If cell.Value = "EM ATENDIMENTO REMOTO" Or cell.Value = "EM ATENDIMENTO PRESENCIAL" Then
        If cell.Offset(0, -10) = nome.Value Then
        
            ordem_de_serviço = cell.Offset(0, -5)
            With os
                .AddItem (ordem_de_serviço)
            End With
        End If
        
    End If
    
Next cell
    

End Sub


Private Sub UserForm_Initialize()

linha = Sheets("VALIDAÇÃO").Range("S1000000").End(xlUp).Row
nome.RowSource = "VALIDAÇÃO!S2:S" & linha


'x = consulta.MultiPage1.SelectedItem.Name

'MsgBox (x)

If consulta.MultiPage1.SelectedItem.Name = "REMOTO" Then

    y = consulta.remoto.ListIndex
    
    os.Value = consulta.remoto.List(y, 0)
    os.Locked = True
    nome.Value = consulta.remoto.List(y, 1)
    nome.Locked = True

ElseIf consulta.MultiPage1.SelectedItem.Name = "PRESENCIAL" Then

    y = consulta.presencial.ListIndex
    
    os.Value = consulta.presencial.List(y, 0)
    os.Locked = True
    nome.Value = consulta.presencial.List(y, 1)
    nome.Locked = True

ElseIf consulta.MultiPage1.SelectedItem.Name = "FINALIZADOS" Then
    
    y = consulta.finalizados.ListIndex
    
    os.Value = consulta.finalizados.List(y, 0)
    os.Locked = True
    nome.Value = consulta.finalizados.List(y, 1)
    nome.Locked = True

End If

End Sub

