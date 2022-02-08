VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} gat 
   Caption         =   "Gastos com Assistência Técnica"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18765
   OleObjectBlob   =   "gat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "gat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pathx As Worksheet
Dim camx As Workbook
Dim gatx_path As String
Dim fca_path As String

Dim data_parcx(2) As Date


Private Sub pdf_click()

Dim ped1_un1 As Integer
Dim ped1_un2 As String
Dim ped1_un3 As String
Dim ped1_un4 As String

If intervalo1 = "OUTRO" Then
    
    If x_data3 <> "" Then
        data_parcx(0) = x_data1.Value
        data_parcx(1) = x_data2.Value
        data_parcx(2) = x_data3.Value
           
    ElseIf x_data3 = "" And x_data2 <> "" Then
        data_parcx(0) = x_data1.Value
        data_parcx(1) = x_data2.Value
        
    ElseIf x_data2 = "" And x_data1 <> "" Then
        data_parcx(1) = x_data2.Value

    End If
    
End If

'ped1_un1 = ""
'ped1_un2 = ""
'ped1_un3 = ""
'ped1_un4 = ""

'If ped1_cat4 <> "" Then
   
    'ped1_un4 = 1
    'ped1_un3 = 1
    'ped1_un2 = 1
    'ped1_un1 = 1
   
'ElseIf ped1_cat3 <> "" Then

    'ped1_un1 = 1
    'ped1_un2 = 1
    'ped1_un3 = 1
   
'ElseIf ped1_cat2 <> "" Then

    'ped1_un1 = 1
    'ped1_un2 = 1

'ElseIf ped1_cat1 <> "" Then
    
    ped1_un1 = 1

'End If


Dim conta_azul As Workbook
Dim formulario As Worksheet

Dim dp As Date

dp = data_parc.Value

Set conta_azul = Workbooks.Open(fca_path, 0, False, , , , True, , , , False, , False)

Set formulario = conta_azul.Sheets("LANÇAMENTO")


formulario.Cells(9, 6).Value = solicitante.Value
formulario.Cells(10, 6).Value = tec1.Value
formulario.Cells(11, 6).Value = nome.Value
formulario.Cells(12, 6).Value = ped1_cat1.Value
formulario.Cells(12, 11).Value = ped1_org1.Value
formulario.Cells(13, 6).Value = os.Value

formulario.Cells(15, 5).Value = des1.Value
formulario.Cells(15, 12).Value = ped1_un1
formulario.Cells(15, 13).Value = ped1_vun1.Value

'formulario.Cells(16, 5).Value = des2.Value
'formulario.Cells(16, 12).Value = ped1_un2
'formulario.Cells(16, 13).Value = ped1_vun2.Value

'formulario.Cells(17, 5).Value = des3.Value
'formulario.Cells(17, 12).Value = ped1_un3
'formulario.Cells(17, 13).Value = ped1_vun3.Value

'formulario.Cells(18, 5).Value = des4.Value
'formulario.Cells(18, 12).Value = ped1_un4
'formulario.Cells(18, 13).Value = ped1_vun4.Value


If cartao1.Value = True Then
    formulario.Cells(30, 6).Value = "Cartão"
ElseIf boleto1.Value = True Then
    formulario.Cells(30, 6).Value = "Boleto"
ElseIf transferencia.Value = True Then
    formulario.Cells(30, 6).Value = "Transferência Bancária/Depósito"
End If


If data_comp1.Value = "" Then
     formulario.Cells(31, 6).Value = Date
Else
    formulario.Cells(31, 6).Value = data_comp1.Value
End If


'If ped1_un4 <> "" Then

    'vtx = ped1_un1 * ped1_vun1 + ped1_un2 * ped1_vun2 + ped1_un3 * ped1_vun3 + ped1_un4 * ped1_vun4

    'If entrada <> "" Then
        
            'vt = vtx - entrada
              '  vp = vt / (par1 - 1)
        'formulario.Cells(32, 6).Value = 1 & " x " & "R$ " & entrada & vbNewLine & (par1 - 1) & " x " & "R$ " & vp
    
    'ElseIf entrada = "" Then
        
       ' vp = vtx / par1
       ' formulario.Cells(32, 6).Value = par1 & " x " & "R$ " & vp
        
   ' End If


'ElseIf ped1_un3 <> "" Then

   ' vtx = ped1_un1 * ped1_vun1 + ped1_un2 * ped1_vun2 + ped1_un3 * ped1_vun3

    'If entrada <> "" Then
        
            'vt = vtx - entrada
             '   vp = vt / (par1 - 1)
        'formulario.Cells(32, 6).Value = 1 & " x " & "R$ " & entrada & vbNewLine & (par1 - 1) & " x " & "R$ " & vp
    
    'ElseIf entrada = "" Then
        
       ' vp = vtx / par1
       ' formulario.Cells(32, 6).Value = par1 & " x " & "R$ " & vp
        
    'End If
    

'ElseIf ped1_un2 <> "" Then

    'vtx = ped1_un1 * ped1_vun1 + ped1_un2 * ped1_vun2

    'If entrada <> "" Then
        
            'vt = vtx - entrada
            '    vp = vt / (par1 - 1)
        'formulario.Cells(32, 6).Value = 1 & " x " & "R$ " & entrada & vbNewLine & (par1 - 1) & " x " & "R$ " & vp
    
    'ElseIf entrada = "" Then
        
        'vp = vtx / par1
        'formulario.Cells(32, 6).Value = par1 & " x " & "R$ " & vp
        
    'End If

        

'ElseIf ped1_un1 <> "" Then

    vtx = ped1_un1 * ped1_vun1

    If entrada <> "" Then
        
            vt = vtx - entrada
                vp = vt / (par1 - 1)
        formulario.Cells(32, 6).Value = 1 & " x " & "R$ " & entrada & vbNewLine & (par1 - 1) & " x " & "R$ " & vp
    
    ElseIf entrada = "" Then
        
        vp = vtx / par1
        formulario.Cells(32, 6).Value = par1 & " x " & "R$ " & vp
        
    End If
    
    formulario.Cells(28, 6).Value = vtx
    
         
'End If


If par1 = 3 Then
    If intervalo1 = "OUTRO" Then
        formulario.Cells(33, 6).Value = data_parcx(0) & vbNewLine & data_parcx(1) & vbNewLine & data_parcx(2)
    Else
        formulario.Cells(33, 6).Value = dp + 0 * intervalo1 & vbNewLine & dp + 1 * intervalo1 & vbNewLine & dp + 2 * intervalo1
    End If
    

ElseIf par1 = 2 Then

    If intervalo1 = "OUTRO" Then
        formulario.Cells(33, 6).Value = data_parcx(0) & vbNewLine & data_parcx(1)
    Else
        formulario.Cells(33, 6).Value = dp + 0 * intervalo1 & vbNewLine & dp + 1 * intervalo1
    End If
    

ElseIf par1 = 1 Then
    
    If intervalo1 = "OUTRO" Then
        formulario.Cells(33, 6).Value = data_parcx(0)
    Else
        formulario.Cells(33, 6).Value = dp + 0 * intervalo1
    End If

End If


formulario.Cells(34, 6).Value = dados_bancarios1

formulario.Cells(35, 6).Value = obs1




'ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=os.Value

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:="NovoGasto" & "-" & os.Value

conta_azul.Close (False)

End Sub

Private Sub registrar_GAT()

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim x As Integer
Dim data_comp As Date
Dim ppar As Date

If intervalo1 = "OUTRO" Then
    
    If x_data3 <> "" Then
        data_parcx(0) = x_data1.Value
        data_parcx(1) = x_data2.Value
        data_parcx(2) = x_data3.Value
           
    ElseIf x_data3 = "" And x_data2 <> "" Then
        data_parcx(0) = x_data1.Value
        data_parcx(1) = x_data2.Value
        
    ElseIf x_data2 = "" And x_data1 <> "" Then
        data_parcx(1) = x_data2.Value

    End If
    
End If

ppar = data_parc.Value

Dim pecas_garantia As Workbook
Dim geral As Worksheet

i = 0
j = 0
k = 0
x = 0


If data_comp1.Value = "" Then
     data_comp = Date
Else
    data_comp = data_comp1.Value
End If


'If ped1_cat4 <> "" Then
    'i = 4
'ElseIf ped1_cat3 <> "" Then
    'i = 3
'ElseIf ped1_cat2 <> "" Then
    'i = 2
'ElseIf ped1_cat1 <> "" Then
    i = 1
'End If


Set gastos_asstec = Workbooks.Open(gatx_path, 0, False, , , , True, , , , False, , False)

Set geral = gastos_asstec.Sheets("GERAL")

'If i = 4 Then

'    linha_vazia = geral.Range("A1000000").End(xlUp).Row + 1
    
'    While j <> i
    
'        geral.Cells(linha_vazia + j, 1).Value = os.Value
'        geral.Cells(linha_vazia + j, 5).Value = tec1.Value
'        geral.Cells(linha_vazia + j, 8).Value = data_comp
     
'        j = j + 1
    
'    Wend
'
'    geral.Cells(linha_vazia, 2).Value = ped1_cat1
'    geral.Cells(linha_vazia + 1, 2).Value = ped1_cat2
'    geral.Cells(linha_vazia + 2, 2).Value = ped1_cat3
'    geral.Cells(linha_vazia + 3, 2).Value = ped1_cat4
    
'    geral.Cells(linha_vazia, 3).Value = ped1_org1
'    geral.Cells(linha_vazia + 1, 3).Value = ped1_org2
'    geral.Cells(linha_vazia + 2, 3).Value = ped1_org3
'    geral.Cells(linha_vazia + 3, 3).Value = ped1_org4
    
'    geral.Cells(linha_vazia, 4).Value = des1.Value
'    geral.Cells(linha_vazia + 1, 4).Value = des2.Value
'    geral.Cells(linha_vazia + 2, 4).Value = des3.Value
'    geral.Cells(linha_vazia + 3, 4).Value = des4.Value
       
'    vt1 = CCur(ped1_vun1)
'    vt2 = CCur(ped1_vun2)
'    vt3 = CCur(ped1_vun3)
'    vt4 = CCur(ped1_vun4)
    
'    If entrada <> "" Then
        
'        vtx = vt1 + vt2 + vt3 + vt4
'        valor_total = vtx - entrada
        
'        e1 = (vt1 / vtx) * entrada
'        e2 = (vt2 / vtx) * entrada
'        e3 = (vt3 / vtx) * entrada
'        e4 = (vt4 / vtx) * entrada
        
'        vt1 = (vt1 / vtx) * valor_total
'        vt2 = (vt2 / vtx) * valor_total
'        vt3 = (vt3 / vtx) * valor_total
'        vt4 = (vt4 / vtx) * valor_total
            
'    End If
    
'    j = 0
'    k = 0
'    x = 0
        
'    While j <> par1
        
'        If entrada <> "" And j = 0 Then
        
'            geral.Cells(linha_vazia, 9).Value = e1
        
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia, 10).Value = data_parcx(0)
            
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia, 10).Value = ppar
'                End If
'           y = 1
            
'        Else
            
'            geral.Cells(linha_vazia, 9 + k).Value = vt1 / (par1 - y)
            
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia, 10 + k).Value = data_parcx(x)
            
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia, 10 + k).Value = ppar + x * intervalo1
'                End If
                
'        End If
        
'        x = x + 1
'        j = j + 1
'        k = k + 2
    
'    Wend

'        j = 0
'        k = 0
'        x = 0
      
'    While j <> par1
'
'        If entrada <> "" And j = 0 Then
'
'           geral.Cells(linha_vazia + 1, 9).Value = e2
'
'            If intervalo1 = "OUTRO" Then
'              geral.Cells(linha_vazia + 1, 10).Value = data_parcx(0)
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia + 1, 10).Value = ppar
'                End If
'            y = 1
            
'        Else
            
'            geral.Cells(linha_vazia + 1, 9 + k).Value = vt2 / (par1 - y)
            
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia + 1, 10 + k).Value = data_parcx(x)
            
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia + 1, 10 + k).Value = ppar + x * intervalo1
'                End If
                
'        End If
        
'        x = x + 1
'        j = j + 1
'        k = k + 2
    
'    Wend
        
'        j = 0
'        k = 0
'        x = 0
    
    
'   While j <> par1
        
'        If entrada <> "" And j = 0 Then
        
'            geral.Cells(linha_vazia + 2, 9).Value = e3
        
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia + 2, 10).Value = data_parcx(0)
            
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia + 2, 10).Value = ppar
'                End If
'            y = 1
            
'        Else
            
'            geral.Cells(linha_vazia + 2, 9 + k).Value = vt3 / (par1 - y)
            
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia + 2, 10 + k).Value = data_parcx(x)
            
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia + 2, 10 + k).Value = ppar + x * intervalo1
'                End If
                
'        End If
        
'        x = x + 1
'        j = j + 1
'        k = k + 2
    
'   Wend
    
'        j = 0
'        k = 0
'        x = 0
    

'        While j <> par1
        
'        If entrada <> "" And j = 0 Then
        
'            geral.Cells(linha_vazia + 3, 9).Value = e4
        
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia + 3, 10).Value = data_parcx(0)
            
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia + 3, 10).Value = ppar
'                End If
'            y = 1
            
'        Else
            
'            geral.Cells(linha_vazia + 3, 9 + k).Value = vt4 / (par1 - y)
            
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia + 3, 10 + k).Value = data_parcx(x)
            
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia + 3, 10 + k).Value = ppar + x * intervalo1
'                End If
                
'        End If
        
'        x = x + 1
'        j = j + 1
'        k = k + 2
    
'    Wend
    
               
'ElseIf i = 3 Then
    
'    linha_vazia = geral.Range("A1000000").End(xlUp).Row + 1
    
'    While j <> i
    
'        geral.Cells(linha_vazia + j, 1).Value = os.Value
'        geral.Cells(linha_vazia + j, 5).Value = tec1.Value
'        geral.Cells(linha_vazia + j, 8).Value = data_comp
     
'        j = j + 1
    
'    Wend
    

'    geral.Cells(linha_vazia, 2).Value = ped1_cat1
'    geral.Cells(linha_vazia + 1, 2).Value = ped1_cat2
'    geral.Cells(linha_vazia + 2, 2).Value = ped1_cat3

'    geral.Cells(linha_vazia, 3).Value = ped1_org1
'    geral.Cells(linha_vazia + 1, 3).Value = ped1_org2
'    geral.Cells(linha_vazia + 2, 3).Value = ped1_org3

'    geral.Cells(linha_vazia, 4).Value = des1.Value
'    geral.Cells(linha_vazia + 1, 4).Value = des2.Value
'    geral.Cells(linha_vazia + 2, 4).Value = des3.Value


           
'    vt1 = CCur(ped1_vun1)
'    vt2 = CCur(ped1_vun2)
'    vt3 = CCur(ped1_vun3)
    
'    If entrada <> "" Then
        
'        vtx = vt1 + vt2 + vt3
'        valor_total = vtx - entrada
        
'        e1 = (vt1 / vtx) * entrada
'       e2 = (vt2 / vtx) * entrada
'       e3 = (vt3 / vtx) * entrada
       
'        vt1 = (vt1 / vtx) * valor_total
'        vt2 = (vt2 / vtx) * valor_total
'        vt3 = (vt3 / vtx) * valor_total
            
'    End If
    


'        j = 0
'        k = 0
'        x = 0
    
'    While j <> par1
            
'            If entrada <> "" And j = 0 Then
            
'                geral.Cells(linha_vazia, 9).Value = e1
            
'                If intervalo1 = "OUTRO" Then
'                    geral.Cells(linha_vazia, 10).Value = data_parcx(0)
                
'                ElseIf intervalo1 <> "OUTRO" Then
'                    geral.Cells(linha_vazia, 10).Value = ppar
'                    End If
'                y = 1
                
'            Else
                
'                geral.Cells(linha_vazia, 9 + k).Value = vt1 / (par1 - y)
                
'                If intervalo1 = "OUTRO" Then
'                    geral.Cells(linha_vazia, 10 + k).Value = data_parcx(x)
                
'                ElseIf intervalo1 <> "OUTRO" Then
'                    geral.Cells(linha_vazia, 10 + k).Value = ppar + x * intervalo1
'                    End If
                    
'            End If
            
'            x = x + 1
'            j = j + 1
'            k = k + 2
        
'        Wend
    
'        j = 0
'        k = 0
'        x = 0
    
'    While j <> par1
            
'            If entrada <> "" And j = 0 Then
            
'                geral.Cells(linha_vazia + 1, 9).Value = e2
            
'                If intervalo1 = "OUTRO" Then
'                    geral.Cells(linha_vazia + 1, 10).Value = data_parcx(0)
                
'                ElseIf intervalo1 <> "OUTRO" Then
'                    geral.Cells(linha_vazia + 1, 10).Value = ppar
'                    End If
'                y = 1
                
'            Else
                
'                geral.Cells(linha_vazia + 1, 9 + k).Value = vt2 / (par1 - y)
                
'                If intervalo1 = "OUTRO" Then
'                    geral.Cells(linha_vazia + 1, 10 + k).Value = data_parcx(x)
                
'                ElseIf intervalo1 <> "OUTRO" Then
'                    geral.Cells(linha_vazia + 1, 10 + k).Value = ppar + x * intervalo1
'                    End If
                    
'            End If
            
'            x = x + 1
'            j = j + 1
'            k = k + 2
        
'       Wend
    
'        j = 0
'        k = 0
'        x = 0
    
    
'    While j <> par1
            
'            If entrada <> "" And j = 0 Then
            
'                geral.Cells(linha_vazia + 2, 9).Value = e3
            
'                If intervalo1 = "OUTRO" Then
'                    geral.Cells(linha_vazia + 2, 10).Value = data_parcx(0)
                
'                ElseIf intervalo1 <> "OUTRO" Then
'                    geral.Cells(linha_vazia + 2, 10).Value = ppar
'                    End If
'                y = 1
                
'            Else
                
'                geral.Cells(linha_vazia + 2, 9 + k).Value = vt3 / (par1 - y)
                
'                If intervalo1 = "OUTRO" Then
'                    geral.Cells(linha_vazia + 2, 10 + k).Value = data_parcx(x)
                
'               ElseIf intervalo1 <> "OUTRO" Then
'                    geral.Cells(linha_vazia + 2, 10 + k).Value = ppar + x * intervalo1
'                    End If
                    
'            End If
            
'            x = x + 1
'            j = j + 1
'            k = k + 2
        
'       Wend
    

'ElseIf i = 2 Then


'    linha_vazia = geral.Range("A1000000").End(xlUp).Row + 1
    
'    While j <> i
    
'        geral.Cells(linha_vazia + j, 1).Value = os.Value
'        geral.Cells(linha_vazia + j, 5).Value = tec1.Value
'        geral.Cells(linha_vazia + j, 8).Value = data_comp
'
'        j = j + 1
'
'    Wend
'
'

'    geral.Cells(linha_vazia, 2).Value = ped1_cat1
'    geral.Cells(linha_vazia + 1, 2).Value = ped1_cat2

'    geral.Cells(linha_vazia, 3).Value = ped1_org1
'    geral.Cells(linha_vazia + 1, 3).Value = ped1_org2

'    geral.Cells(linha_vazia, 4).Value = des1.Value
'    geral.Cells(linha_vazia + 1, 4).Value = des2.Value
        
'    vt1 = CCur(ped1_vun1)
'    vt2 = CCur(ped1_vun2)
    
'    If entrada <> "" Then
        
'        vtx = vt1 + vt2
'        valor_total = vtx - entrada
        
'        e1 = (vt1 / vtx) * entrada
'        e2 = (vt2 / vtx) * entrada

 '       vt1 = (vt1 / vtx) * valor_total
 '       vt2 = (vt2 / vtx) * valor_total
 
 '   End If

    
'        j = 0
'        k = 0
'        x = 0

'     While j <> par1
        
'        If entrada <> "" And j = 0 Then
        
'            geral.Cells(linha_vazia, 9).Value = e1
        
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia, 10).Value = data_parcx(0)
            
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia, 10).Value = ppar
'                End If
'            y = 1
'
'        Else
'
'            geral.Cells(linha_vazia, 9 + k).Value = vt1 / (par1 - y)
'
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia, 10 + k).Value = data_parcx(x)
'
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia, 10 + k).Value = ppar + x * intervalo1
'                End If
'
'        End If
'
'        x = x + 1
'        j = j + 1
'        k = k + 2
    
'    Wend
    
'        j = 0
'        k = 0
'        x = 0

'    While j <> par1
'
'        If entrada <> "" And j = 0 Then
'
'            geral.Cells(linha_vazia + 1, 9).Value = e2
'
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia + 1, 10).Value = data_parcx(0)
'
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia + 1, 10).Value = ppar
'                End If
'            y = 1
'
'        Else
            
'            geral.Cells(linha_vazia + 1, 9 + k).Value = vt2 / (par1 - y)
            
'            If intervalo1 = "OUTRO" Then
'                geral.Cells(linha_vazia + 1, 10 + k).Value = data_parcx(x)
            
'            ElseIf intervalo1 <> "OUTRO" Then
'                geral.Cells(linha_vazia + 1, 10 + k).Value = ppar + x * intervalo1
'                End If
                
'        End If
        
'        x = x + 1
'        j = j + 1
'        k = k + 2

'   Wend
    

'ElseIf i = 1 Then


    linha_vazia = geral.Range("A1000000").End(xlUp).Row + 1
        
    geral.Cells(linha_vazia, 1).Value = os.Value
    geral.Cells(linha_vazia, 5).Value = tec1.Value
    geral.Cells(linha_vazia, 8).Value = data_comp
    
    geral.Cells(linha_vazia, 2).Value = ped1_cat1
    geral.Cells(linha_vazia, 3).Value = ped1_org1
    geral.Cells(linha_vazia, 4).Value = des1.Value
             
    vt1 = CCur(ped1_vun1)
    
    If entrada <> "" Then
        
        vtx = vt1
        valor_total = vtx - entrada
        e1 = (vt1 / vtx) * entrada
        vt1 = (vt1 / vtx) * valor_total

    End If

        j = 0
        k = 0
        x = 0

    While j <> par1
        
        If entrada <> "" And j = 0 Then
        
            geral.Cells(linha_vazia, 9).Value = e1
        
            If intervalo1 = "OUTRO" Then
                geral.Cells(linha_vazia, 10).Value = data_parcx(0)
            
            ElseIf intervalo1 <> "OUTRO" Then
                geral.Cells(linha_vazia, 10).Value = ppar
                End If
            y = 1
            
        Else
            
            geral.Cells(linha_vazia, 9 + k).Value = vt1 / (par1 - y)
            
            If intervalo1 = "OUTRO" Then
                geral.Cells(linha_vazia, 10 + k).Value = data_parcx(x)
            
            ElseIf intervalo1 <> "OUTRO" Then
                geral.Cells(linha_vazia, 10 + k).Value = ppar + x * intervalo1
                End If
                
        End If
        
        x = x + 1
        j = j + 1
        k = k + 2
    
    Wend

    
'End If

gastos_asstec.Save
gastos_asstec.Close (True)


MsgBox ("Incluido nos Gastos com Ass. Técnica com Sucesso")
 
End Sub

Private Sub botao_ok_Click()

Unload gat

End Sub

Private Sub botao_x_Click()
     
On Error Resume Next

Unload gat
Unload options
Unload consulta


End Sub

Private Sub Confirma1_Click()

Call registrar_GAT

status1.Value = "Incluido"

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

Private Sub os_Change()

equip.Value = Sheets("GERAL").Cells.Find(os.Value, , xlValues, xlWhole).Offset(0, 8).Value
nserie.Value = Sheets("GERAL").Cells.Find(os.Value, , xlValues, xlWhole).Offset(0, -6).Value
data_chamado.Value = Sheets("GERAL").Cells.Find(os.Value, , xlValues, xlWhole).Offset(0, 2).Value


End Sub
Private Sub intervalo1_AfterUpdate()

If intervalo1 = "OUTRO" Then
    datas.Show
End If

End Sub

Private Sub UserForm_Initialize()

Set camx = Workbooks.Open(ThisWorkbook.Path & "\" & "Local Paths.xlsx", 0, False, , , , True, , , , False, , False)
Set pathx = camx.Sheets("Caminhos")

gatx_path = pathx.Cells.Find("GAT", , xlValues, xlWhole).Offset(1, 0).Value
fca_path = pathx.Cells.Find("GAT", , xlValues, xlWhole).Offset(1, 0).Value
camx.Close (False)


linha = Sheets("VALIDAÇÃO").Range("S1000000").End(xlUp).Row
nome.RowSource = "VALIDAÇÃO!S2:S" & linha

linha = Sheets("VALIDAÇÃO").Range("K1000000").End(xlUp).Row
solicitante.RowSource = "VALIDAÇÃO!K2:K" & linha

linha = Sheets("VALIDAÇÃO").Range("K1000000").End(xlUp).Row
tec1.RowSource = "VALIDAÇÃO!K2:K" & linha
'forn2.RowSource = "VALIDAÇÃO!K2:K" & linha
'forn3.RowSource = "VALIDAÇÃO!K2:K" & linha

linha = Sheets("VALIDAÇÃO").Range("Q1000000").End(xlUp).Row
intervalo1.RowSource = "VALIDAÇÃO!Q2:Q" & linha
'intervalo2.RowSource = "VALIDAÇÃO!Q2:Q" & linha
'intervalo3.RowSource = "VALIDAÇÃO!Q2:Q" & linha

linha = Sheets("VALIDAÇÃO").Range("R1000000").End(xlUp).Row
par1.RowSource = "VALIDAÇÃO!R2:R" & linha
'par2.RowSource = "VALIDAÇÃO!R2:R" & linha
'par3.RowSource = "VALIDAÇÃO!R2:R" & linha

linha = Sheets("VALIDAÇÃO").Range("U1000000").End(xlUp).Row
ped1_cat1.RowSource = "VALIDAÇÃO!U2:U" & linha
'ped1_cat2.RowSource = "VALIDAÇÃO!U2:U" & linha
'ped1_cat3.RowSource = "VALIDAÇÃO!U2:U" & linha
'ped1_cat4.RowSource = "VALIDAÇÃO!U2:U" & linha

linha = Sheets("VALIDAÇÃO").Range("V1000000").End(xlUp).Row
ped1_org1.RowSource = "VALIDAÇÃO!V2:V" & linha
'ped1_org2.RowSource = "VALIDAÇÃO!V2:V" & linha
'ped1_org3.RowSource = "VALIDAÇÃO!V2:V" & linha
'ped1_org4.RowSource = "VALIDAÇÃO!V2:V" & linha


End Sub

