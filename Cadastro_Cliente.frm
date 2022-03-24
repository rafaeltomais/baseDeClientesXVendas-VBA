VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cadastro_Cliente 
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8865.001
   OleObjectBlob   =   "Cadastro_Cliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cadastro_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bt_cadastrar_Click()
    
    linha = Sheets("BASE").Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Row
    
    txt_razao = UCase(txt_razao.Value)
    txt_fazenda = UCase(txt_fazenda.Value)
    txt_uf = UCase(txt_uf.Value)
    txt_cidade = UCase(txt_cidade.Value)
    txt_bairro = UCase(txt_bairro.Value)
    txt_logradouro = UCase(txt_logradouro.Value)
    txt_contato = UCase(txt_contato.Value)
    
    Planilha7.Cells(linha, 1).Value = Me.txt_razao.Value
    Planilha7.Cells(linha, 2).Value = Me.txt_fazenda.Value
    Planilha7.Cells(linha, 3).Value = Me.txt_cpfcnpj.Value
    Planilha7.Cells(linha, 4).Value = Me.txt_ie.Value
    Planilha7.Cells(linha, 5).Value = Me.txt_uf.Value
    Planilha7.Cells(linha, 6).Value = Me.txt_cidade.Value
    Planilha7.Cells(linha, 7).Value = Me.txt_bairro.Value
    Planilha7.Cells(linha, 8).Value = Me.txt_logradouro.Value
    Planilha7.Cells(linha, 9).Value = Me.txt_n.Value
    Planilha7.Cells(linha, 10).Value = Me.txt_cep.Value
    Planilha7.Cells(linha, 11).Value = Me.txt_contato.Value
    Planilha7.Cells(linha, 12).Value = Me.txt_tel1.Value
    Planilha7.Cells(linha, 13).Value = Me.txt_tel2.Value
    
    Me.txt_razao = Null
    Me.txt_fazenda = Null
    Me.txt_cpfcnpj.Value = Null
    Me.txt_ie.Value = Null
    Me.txt_uf.Value = Null
    Me.txt_cidade.Value = Null
    Me.txt_bairro.Value = Null
    Me.txt_logradouro.Value = Null
    Me.txt_n.Value = Null
    Me.txt_cep.Value = Null
    Me.txt_contato.Value = Null
    Me.txt_tel1.Value = Null
    Me.txt_tel2.Value = Null
    
End Sub


Private Sub bt_cancelar_Click()
    Unload Cadastro_Cliente
End Sub


