Attribute VB_Name = "Impressao"
Public pb_Impressao_Normal As String
Public pb_Impressao_Expandida As String
Public pb_Impressao_Condensada As String
Public pb_Impressao_Condensada_N As String
Public pb_Impressao_Normal_N As String
Public pb_Impressao_Expandida_N As String
Public pb_CancelaImpressao As Boolean
Public pb_PadraoVideo As Boolean
Public pb_Buffer As String
Public pb_LinhaBuffer As Integer
Public pb_UltimaLinha As Integer
Public pb_Tamanho As Imp_Constantes
Public pb_ImpressaoMatricial As Boolean

Public Enum Imp_Constantes
   imp_Condensado = 0
   Imp_Normal = 1
   imp_Expandido = 2
   imp_Condensado_NEGRITO = 3
   imp_Normal_Negrito = 4
   imp_Expandido_Negrito = 5
End Enum

Public pb_Impressora As Integer
Public Const cn_Epson = 0
Public pb_LigaCompactado As String
Public pb_Desligacompactado As String
Public pb_Liga12cpicompactado As String
Public pb_Liga12cpi As String
Public pb_Reset As String
Public pb_LigaNegrito As String
Public pb_DesligaNegrito As String
Public pb_LigaExpandido As String
Public pb_DesligaExpandido As String
Public pb_NumerodeLinhasporPagina As String


Sub AddColunaImpressao(ByRef pMatriz, pCampo, Optional pAlinhamento = "E", Optional pFormat = "", Optional pCampoRetorno = "", Optional pSoma As Boolean = False, Optional pTabelaConsulta As Integer = 2, Optional pNumerodeCaracteres As Integer = 0)
aadd pMatriz, Array(pCampo, pAlinhamento, pFormat, pCampoRetorno, pSoma, pTabelaConsulta, pNumerodeCaracteres)
End Sub

Sub Finaliza_Impressao()
On Error Resume Next
If Not pb_PadraoVideo Then
   If Mid(pb_Impressao_Normal, 1, 5) <> "Draft" Then
      Printer.EndDoc
      pb_ImpressaoIniciada = False
   Else
      Imprime pb_LinhaImpressaoMatricial + 0.5, 0, " ", pb_Tamanho
      Close #1
      pb_Buffer = ""
      pb_LinhaBuffer = 0
   End If
Else
   Imprime pb_LinhaBuffer + 0.5, 0, " ", imp_Condensado, , , pb_PadraoVideo
   Print #1, "}"
   pausa 1
   Close #1
   If Dir(PathWindows + "WORDPAD.EXE") = "" Then
      InformaaoUsuario "Para visualizar os relatórios no vídeo localize e copie o aplicativo: WORDPAD.EXE, para a pasta " + PathWindows
   Else
      If Confirme("Visualiza o relatório?") Then
         Shell "WORDPAD.EXE " + PathWindows + "RTF\REPORT.RTF", vbMaximizedFocus
      End If
   End If
End If
pb_LinhaBuffer = 0
End Sub



Sub Inicializa_Impressora()
If pb_Impressora = cn_Epson Then
   pb_LigaCompactado = Chr(15)
   pb_Desligacompactado = Chr(18)
   pb_Liga12cpicompactado = Chr(15) + Chr(27) + Chr(77)
   pb_Liga12cpi = Chr(27) + Chr(77)
   pb_Reset = Chr(27) + Chr(64)
   pb_LigaNegrito = Chr(27) + Chr(69)
   pb_DesligaNegrito = Chr(27) + Chr(70)
   pb_LigaExpandido = Chr(14)
   pb_DesligaExpandido = Chr(20)
   pb_NumerodeLinhasporPagina = Chr(27) + "C"
End If
End Sub

Sub RelatorioPadrao(pTabela As Recordset, pTitulo As String, pCabecalho, pReferencia, pCampos, Optional ptamanho As Imp_Constantes = Imp_Normal, Optional pTituloSoma As String = "", Optional pSegundaTabela As Recordset, Optional pTerceiraTabela As Recordset, Optional pQuartaTabela As Recordset, Optional pQuintaTabela As Recordset)
Dim i As Integer
Dim wl_Linha As Currency
Dim wl_Folha As Integer
Dim pContagem As Long
Dim wl_Soma
ReDim wl_Soma(0, UBound(pCampos))
For i = 0 To UBound(pCampos)
   wl_Soma(0, i) = 0
Next
If pTabela.RecordCount = 0 Then
   Exit Sub
End If
If Not PadraodeImpressao Then Exit Sub
pTabela.MoveFirst
Do While Not pTabela.EOF
   If wl_Linha = 0 Then
      wl_Folha = wl_Folha + 1
      Monta_Cabecalho pCabecalho, pReferencia, UBound(pCampos), wl_Linha, ptamanho, pTitulo, wl_Folha
   End If
   For i = 0 To UBound(pCampos)
      If Not IsNull(pTabela(pCampos(i, 0))) Then
         If pCampos(i, 2) = "" Then
            If pCampos(i, 3) = "" Then
               If pCampos(i, 6) = 0 Then
                  Monta_LinhadeImpressao wl_Linha, pTabela(pCampos(i, 0)), i, pCampos(i, 1), ptamanho
               Else
                  Monta_LinhadeImpressao wl_Linha, Mid(pTabela(pCampos(i, 0)), 1, pCampos(i, 6)), i, pCampos(i, 1), ptamanho
               End If
            Else
               If pCampos(i, 5) = 2 Then
                  pSegundaTabela.Seek "=", pTabela(pCampos(i, 0))
                  If Not pSegundaTabela.NoMatch Then
                     If pCampos(i, 6) = 0 Then
                        Monta_LinhadeImpressao wl_Linha, pSegundaTabela(pCampos(i, 3)), i, pCampos(i, 1), ptamanho
                     Else
                        Monta_LinhadeImpressao wl_Linha, Mid(pSegundaTabela(pCampos(i, 3)), 1, pCampos(i, 6)), i, pCampos(i, 1), ptamanho
                     End If
                  End If
               ElseIf pCampos(i, 5) = 3 Then
                  pTerceiraTabela.Seek "=", pTabela(pCampos(i, 0))
                  If Not pTerceiraTabela.NoMatch Then
                     If pCampos(i, 6) = 0 Then
                        Monta_LinhadeImpressao wl_Linha, pTerceiraTabela(pCampos(i, 3)), i, pCampos(i, 1), ptamanho
                     Else
                        Monta_LinhadeImpressao wl_Linha, Mid(pTerceiraTabela(pCampos(i, 3)), 1, pCampos(i, 6)), i, pCampos(i, 1), ptamanho
                     End If
                  End If
               ElseIf pCampos(i, 5) = 4 Then
                  pQuartaTabela.Seek "=", pTabela(pCampos(i, 0))
                  If Not pQuartaTabela.NoMatch Then
                     If pCampos(i, 6) = 0 Then
                        Monta_LinhadeImpressao wl_Linha, pQuartaTabela(pCampos(i, 3)), i, pCampos(i, 1), ptamanho
                     Else
                        Monta_LinhadeImpressao wl_Linha, Mid(pQuartaTabela(pCampos(i, 3)), 1, pCampos(i, 6)), i, pCampos(i, 1), ptamanho
                     End If
                  End If
               ElseIf pCampos(i, 5) = 5 Then
                  pQuintaTabela.Seek "=", pTabela(pCampos(i, 0))
                  If Not pQuintaTabela.NoMatch Then
                     If pCampos(i, 6) = 0 Then
                        Monta_LinhadeImpressao wl_Linha, pQuintaTabela(pCampos(i, 3)), i, pCampos(i, 1), ptamanho
                     Else
                        Monta_LinhadeImpressao wl_Linha, Mid(pQuintaTabela(pCampos(i, 3)), 1, pCampos(i, 6)), i, pCampos(i, 1), ptamanho
                     End If
                  End If
               End If
            End If
         Else
            If pCampos(i, 3) = "" Then
               Monta_LinhadeImpressao wl_Linha, Format(pTabela(pCampos(i, 0)), pCampos(i, 2)), i, pCampos(i, 1), ptamanho
            Else
               pSegundaTabela.Seek "=", pTabela(pCampos(i, 0))
               If Not pSegundaTabela.NoMatch Then
                  Monta_LinhadeImpressao wl_Linha, Format(pTabela(pCampos(i, 0)), pCampos(i, 2)) + " - " + pSegundaTabela(pCampos(i, 3)), i, pCampos(i, 1), ptamanho
               End If
            End If
            If pCampos(i, 4) Then
               wl_Soma(0, i) = wl_Soma(0, i) + pTabela(pCampos(i, 0))
            End If
         End If
      End If
   Next
   wl_Linha = wl_Linha + 0.5
   pContagem = pContagem + 1
   If wl_Linha > IIf(Mid(pb_Impressao_Normal, 1, 5) = "Draft", 29, 26) Then
      wl_Linha = 0
      Salta_Pagina
   End If
   pTabela.MoveNext
Loop
If pTituloSoma <> "" Then
   wl_Linha = wl_Linha + 0.5
   Imprime wl_Linha, 0, pTituloSoma, ptamanho
End If
For i = 0 To UBound(pCampos)
   If pCampos(i, 4) Then
      If pCampos(i, 2) = "" Then
         Monta_LinhadeImpressao wl_Linha, wl_Soma(0, i), i, pCampos(i, 1), ptamanho
      Else
         Monta_LinhadeImpressao wl_Linha, Format(wl_Soma(0, i), pCampos(i, 2)), i, pCampos(i, 1), ptamanho
      End If
   End If
Next
Imprime wl_Linha + 0.5, 0, "Impresso(s) .." + Str(pContagem), ptamanho
Finaliza_Impressao
End Sub


Function Salta_Pagina()
If Mid(pb_Impressao_Normal, 1, 5) <> "Draft" And Not pb_PadraoVideo Then
   Printer.NewPage
Else
   If Not pb_PadraoVideo Then
      Imprime pb_LinhaBuffer + 0.5, 0, " ", pb_Tamanho
   Else
      Imprime pb_LinhaBuffer + 0.5, 0, String(123, "*"), imp_Condensado, , , pb_PadraoVideo
   End If
   For i = pb_UltimaLinha To 64
       Print #1, IIf(pb_PadraoVideo, "\par", " ")
   Next
   pb_LinhaBuffer = 0
   pb_Buffer = ""
   pb_UltimaLinha = 0
End If
End Function




