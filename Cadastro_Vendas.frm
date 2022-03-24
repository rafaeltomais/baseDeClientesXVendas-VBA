VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cadastro_Vendas 
   Caption         =   "Vendas"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4905
   OleObjectBlob   =   "Cadastro_Vendas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadastro_vendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bt_cancelar_Click()
    Unload cadastro_vendas
End Sub


Private Sub UserForm_Initialize()

    ultima_linha1 = Sheets("OP합ES").Range("A1").End(xlDown).Row
    lista_vendedor.RowSource = "OP합ES!A2:A" & ultima_linha1
    
    ultima_linha2 = Sheets("OP합ES").Range("B1").End(xlDown).Row
    lista_transportadora.RowSource = "OP합ES!B2:B" & ultima_linha2
    
    ultima_linha3 = Sheets("BASE").Range("A1").End(xlDown).Row
    lista_clientes.RowSource = "BASE!A2:A" & ultima_linha3

End Sub

Private Sub bt_ok_Click()
        
    Dim v1 As Double, v2 As Double, v3 As Double, v4 As Double
    
    If txt_pecas = empity Then
        v1 = 0
        Else
            v1 = txt_pecas
    End If
    
    If txt_quimicos = empity Then
        v2 = 0
        Else
            v2 = txt_quimicos
    End If
    
    If txt_maodeobra = empity Then
        v3 = 0
        Else
            v3 = txt_maodeobra.Value
    End If
    
    If txt_frete = empity Then
        v4 = 0
        Else
            v4 = txt_frete.Value
    End If
    
    ultima_linha = Sheets("PEDIDOS").Range("A1048576").End(xlUp).Row + 1

    Sheets("PEDIDOS").Cells(ultima_linha, 1).Value = lista_vendedor.Value
    Sheets("PEDIDOS").Cells(ultima_linha, 2).Value = txt_npedido.Value
    Sheets("PEDIDOS").Cells(ultima_linha, 3).Value = txt_nf.Value
    Sheets("PEDIDOS").Cells(ultima_linha, 4).Value = DateValue(txt_data.Value)
    Sheets("PEDIDOS").Cells(ultima_linha, 5).Value = lista_clientes.Value
    Sheets("PEDIDOS").Cells(ultima_linha, 6).Value = v1
    Sheets("PEDIDOS").Cells(ultima_linha, 7).Value = v2
    Sheets("PEDIDOS").Cells(ultima_linha, 8).Value = v3
    Sheets("PEDIDOS").Cells(ultima_linha, 9).Value = lista_transportadora.Value
    Sheets("PEDIDOS").Cells(ultima_linha, 10).Value = v4
    
    lista_vendedor.Value = Null
    txt_npedido.Value = Null
    txt_nf.Value = Null
    txt_data.Value = Null
    lista_clientes.Value = Null
    txt_pecas.Value = Null
    txt_quimicos.Value = Null
    txt_maodeobra.Value = Null
    lista_transportadora.Value = Null
    txt_frete.Value = Null

End Sub
