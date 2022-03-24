# Abaixo os códigos utilizados na macro:

### Para inserir um novo cliente na base de dados clicando em "OK" na janela
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

## Para encerrar a janela de inserir novos clientes.
Private Sub bt_cancelar_Click()
    Unload Cadastro_Cliente
End Sub

-----------------------------------------------------------

## Para inserir uma nova venda na base de dados clicando em "OK" na janela
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

