VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} datas 
   Caption         =   "Datas Personalizadas"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "datas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "datas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub botao_ok_Click()

peg.x_data1 = data1
peg.x_data2 = data2
peg.x_data3 = data3

gat.x_data1 = data1
gat.x_data2 = data2
gat.x_data3 = data3

peg.data_parc = data1
gat.data_parc = data1

If data3 <> "" Then
        peg.par1 = 3
        peg.par1.Locked = True
        peg.data_parc.Locked = True
        gat.par1 = 3
        gat.par1.Locked = True
        gat.data_parc.Locked = True
           
ElseIf data3 = "" And data2 <> "" Then
        peg.par1 = 2
        peg.par1.Locked = True
        peg.data_parc.Locked = True
        gat.par1 = 2
        gat.par1.Locked = True
        gat.data_parc.Locked = True
    
ElseIf data2 = "" And data1 <> "" Then
        peg.par1 = 1
        peg.par1.Locked = True
        peg.data_parc.Locked = True
        gat.par1 = 1
        gat.par1.Locked = True
        gat.data_parc.Locked = True
End If
    

Unload datas

End Sub

Private Sub botao_x_Click()

Unload datas

End Sub
