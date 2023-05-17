Attribute VB_Name = "TOOLS_GERAL"

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public pb_Demonstracao As Boolean
Public pb_LinhadeImpressao
Public pb_RetornodeFuncao As Boolean
Public a_Browse001
Public pbRetornoVideo
Public pbRetornoVideo2 As String
Public pbCampodeRetorno As Variant
Public aCampo
Public matriz
Public campo
Public ShowRetorno As Boolean
Public ShowRetornoProdutos As Boolean
Public ErrodeAbertura As Integer
Public BancodeDadosdaConsulta As String
Public TabeladaConsulta As String
Public IndicedaConsulta As String
Public pb_CampodePesquisa As String
Public PathPadrao As String
Public pb_ImpressaoIniciada As Boolean
Public pb_TitulodaConsulta As String
Public wp_produto As String

Public Const DB_LANG_GENERAL = ";LANGID=0x0809;CP=1252;COUNTRY=0"

Public Enum dbCONSTANTES
   dbBoolean = 1
   dbByte = 2
   dbInteger = 3
   dbLong = 4
   dbCurrency = 5
   dbSingle = 6
   dbDouble = 7
   dbDate = 8
   dbText = 10
   dbLongBinary = 11
   dbMemo = 12
End Enum

Public Const AZUL = &HFF0000
Public Const BRANCO = &HFFFFFF
Public Const PRETO = &H0&
Public Const AMARELO = &H80FFFF
Public Const AMARELO_CLARO = &HC0FFFF
Public Const VERMELHO = &HFF&
Public Const ROXO = &H800080
Public Const VERMELHO_OFF = &H40&
Public Const VERDE = &HFF00&
Public Const ROSA = &H8080FF
Public Const PALHA = &HC0FFFF
Public Const AZUL_DISABLE = &HFFC0C0
Public Const CINZA = &HC0C0C0
Public Const CIANO = &H808000
Public Const CIANO_SOMBRA = &H404000
Public Const CIANO_BRILHO = &HFFFF00

Public passw As String
Public chave As String
Public pswmsg As String

Public Const cnEsquerda = 0
Public Const cnDireita = 1
Public Const cnCentro = 2

Public Const cn_FONTE_NORMAL = "Draft 10cpi"
Public Const cn_FONTE_CONDENSADA = "Sans Serif 17cpi"

Public Const senhapadrao = "sisjm"

Public Const MF_BITMAP = &H4


Function Atualiza_QuantidadeVenda(ByRef tbquant0 As Recordset, pCodigo As String, pQTDVENDA As Currency, pVALORVENDA As Currency, pCUSTOVENDA As Currency, pLUCROOBTIDO As Currency) As Boolean
Dim tbquant As Recordset
Dim XMES As String

Atualiza_QuantidadeVenda = False

XMES = Mes_Extenso(Month(Date))

tbquant0.Index = "iCODIGO"
tbquant0.Seek "=", pCodigo, XMES

If tbquant0.NoMatch Then
   
   Exit Function
  
  Else
   
   If Not edit_reg(tbquant0) Then
      Exit Function
    Else
      tbquant0("qtdvenda") = tbquant0("qtdvenda") + pQTDVENDA
      tbquant0("vlrvenda") = tbquant0("vlrvenda") + (pVALORVENDA * pQTDVENDA)
      tbquant0("custovenda") = tbquant0("custovenda") + (pCUSTOVENDA * pQTDVENDA)
      tbquant0("lucro") = tbquant0("lucro") + ((pVALORVENDA * pQTDVENDA) - (pCUSTOVENDA * pQTDVENDA))
   End If
   
   If update_reg(tbquant0) Then
      Atualiza_QuantidadeVenda = True
   End If

End If

End Function

Function Formata_Conta_Mantovani(conta As String) As String


Formata_Conta_Mantovani = Mid$(conta, 1, 1) + Mid$(conta, 3, 1) + "." + Mid$(conta, 5, 1) + Mid$(conta, 7, 1) + Mid$(conta, 8, 1) + "." + Mid$(conta, 10, 5)


End Function

Sub Compacta(pCompactado As String, pCompactar As String)
Shell PathWindows + "RAR.EXE A " + pCompactado + " " + pCompactar
MsgBox "ARQUIVO SENDO COMPACTADO: " + pCompactar
End Sub

Function etiqueta26x15(XLIN As Single, C1 As String, n1 As String, C2 As String, n2 As String, c3 As String, n3 As String, c4 As String, n4 As String, c5 As String, n5 As String) As String
Dim lin
    
    

    lin = XLIN

    Imprime lin, 0, C1 + Chr(27) + Chr(65) + Chr(8), imp_Condensado
    Imprime lin, 23, C2, imp_Condensado
    Imprime lin, 44, c3, imp_Condensado
    Imprime lin, 65, c4, imp_Condensado
    Imprime lin, 86, c5, imp_Condensado
    
    lin = lin + 0.5
    
    Imprime lin, 0, Chr(27) + Chr(65) + Chr(8) + Mid(n1, 1, 7), imp_Condensado
    Imprime lin, 23, Mid(n2, 1, 15), imp_Condensado
    Imprime lin, 44, Mid(n3, 1, 15), imp_Condensado
    Imprime lin, 65, Mid(n4, 1, 15), imp_Condensado
    Imprime lin, 86, Mid(n5, 1, 15), imp_Condensado
    
    lin = lin + 0.5
    
    Imprime lin, 0, "", imp_Condensado

        

End Function


Function etiqueta107x23x2(XLIN As Single, Optional c11 = "", Optional c21 = "", Optional c12 = "", Optional c22 = "", Optional c13 = "", Optional c23 = "", Optional c14 = "", Optional c24 = "", Optional c15 = "", Optional c25 = "") As String

Dim lin
    
    lin = XLIN

    Imprime lin, 0, c11, imp_Condensado
    Imprime lin, 72, c21, imp_Condensado
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c12, 1, 50), imp_Condensado
    Imprime lin, 72, Mid(c22, 1, 50), imp_Condensado
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c13, 1, 40), imp_Condensado
    Imprime lin, 72, Mid(c23, 1, 40), imp_Condensado
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c14, 1, 40), imp_Condensado
    Imprime lin, 72, Mid(c24, 1, 40), imp_Condensado
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c15, 1, 40), imp_Condensado
    Imprime lin, 72, Mid(c25, 1, 40), imp_Condensado
    

End Function

Function etiqueta102x23x2(XLIN As Single, Optional c11 = "", Optional c21 = "", Optional c12 = "", Optional c22 = "", Optional c13 = "", Optional c23 = "", Optional c14 = "", Optional c24 = "", Optional c15 = "", Optional c25 = "") As String

Dim lin
    
    lin = XLIN

    Imprime lin, 0, c11, Imp_Normal
    Imprime lin, 41, c21, Imp_Normal
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c12, 1, 40), Imp_Normal
    Imprime lin, 41, Mid(c22, 1, 40), Imp_Normal
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c13, 1, 40), Imp_Normal
    Imprime lin, 41, Mid(c23, 1, 40), Imp_Normal
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c14, 1, 40), Imp_Normal
    Imprime lin, 41, Mid(c24, 1, 40), Imp_Normal
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c15, 1, 40), Imp_Normal
    Imprime lin, 41, Mid(c25, 1, 40), Imp_Normal
    

End Function

Function etiqueta89x23x2(XLIN As Single, Optional c11 = "", Optional c21 = "", Optional c12 = "", Optional c22 = "", Optional c13 = "", Optional c23 = "", Optional c14 = "", Optional c24 = "", Optional c15 = "", Optional c25 = "") As String

Dim lin
    
    lin = XLIN

    Imprime lin, 0, Mid(c11, 1, 35), Imp_Normal
    Imprime lin, 37, Mid(c21, 1, 35), Imp_Normal
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c12, 1, 35), Imp_Normal
    Imprime lin, 37, Mid(c22, 1, 35), Imp_Normal
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c13, 1, 35), Imp_Normal
    Imprime lin, 37, Mid(c23, 1, 35), Imp_Normal
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c14, 1, 35), Imp_Normal
    Imprime lin, 37, Mid(c24, 1, 35), Imp_Normal
    
    lin = lin + 0.5
    
    Imprime lin, 0, Mid(c15, 1, 35), Imp_Normal
    Imprime lin, 37, Mid(c25, 1, 35), Imp_Normal
    

End Function


Function calc_jurosdias(pValor As Currency, pJurosDia, pQuantDias) As Currency
Dim ZZ As Long
Dim juros As Currency

juros = 0

For ZZ = 1 To pQuantDias Step 1

    juros = juros + ((pValor * pJurosDia) / 100)

Next ZZ

calc_jurosdias = juros

End Function

Function Numero_Contabil(pValor As Currency) As String

Numero_Contabil = Format(IIf(pValor < 0, pValor * -1, pValor), "###,###,##0.00") + IIf(pValor < 0, " D", " C")

End Function

Function Dif_Datas(Data1 As Date, Data2 As Date) As Long

Dif_Datas = Data2 - Data1

End Function

Sub Monta_Cabecalho(ByVal pCabecalho, ByVal pReferencia, pElementos As Integer, Optional ByRef pLinha = 0, Optional ptamanho As Imp_Constantes = 1, Optional pNomedoRelatorio As String = "", Optional pFolha As Integer = -1, Optional pData As Date)
Dim wl_Coluna As Currency
Dim i As Integer
Dim wl_LinhaInicial As Currency
Dim wl_LinhaNome As Currency
Dim wl_ImpressaoMatricial As Boolean
Dim wl_Tamanho As Integer

If pData = "00:00:00" Then
   pData = Date
End If

wl_ImpressaoMatricial = Mid(pb_Impressao_Normal, 1, 5) = "Draft"
ReDim pb_LinhadeImpressao(0)
wl_Coluna = 0
If Not wl_ImpressaoMatricial And Not pb_PadraoVideo Then
   If Not pb_PadraoVideo Then
      Printer.ScaleMode = 7
   End If
   wl_LinhaInicial = pLinha
   If pNomedoRelatorio <> "" Then
      Imprime pLinha, 0, pb_RAZAOSOCIAL, ptamanho
      pLinha = pLinha + 0.5
      Imprime pLinha, 0, pNomedoRelatorio, ptamanho
      wl_LinhaNome = pLinha
      pLinha = pLinha + 0.5
   End If
   For i = 0 To pElementos
      Imprime pLinha, wl_Coluna, pCabecalho(0, i), ptamanho
      If Printer.TextWidth(pReferencia(0, i)) >= Printer.TextWidth(pCabecalho(0, i)) Then
         If Not pb_PadraoVideo Then Call Linha(pLinha + 0.5, wl_Coluna, pLinha + 0.5, wl_Coluna + Printer.TextWidth(pReferencia(0, i)))
         aadd pb_LinhadeImpressao, Array(wl_Coluna, wl_Coluna + Printer.TextWidth(pReferencia(0, i)))
         wl_Coluna = wl_Coluna + Printer.TextWidth(pReferencia(0, i)) + IIf(Dir(PathPadrao + "ITUVEPLAST.SYS") = "", 0.5, 1)
      Else
         If Not pb_PadraoVideo Then Call Linha(pLinha + 0.5, wl_Coluna, pLinha + 0.5, wl_Coluna + Printer.TextWidth(pCabecalho(0, i)))
         aadd pb_LinhadeImpressao, Array(wl_Coluna, wl_Coluna + Printer.TextWidth(pCabecalho(0, i)))
         wl_Coluna = wl_Coluna + Printer.TextWidth(pCabecalho(0, i)) + IIf(Dir(PathPadrao + "ITUVEPLAST.SYS") = "", 0.5, 1)
      End If
   Next
   If pFolha >= 0 Then
      Imprime wl_LinhaInicial, wl_Coluna - Printer.TextWidth("Folha .." + Str(pFolha)) - IIf(Dir(PathPadrao + "ITUVEPLAST.SYS") = "", 0.5, 1), "Folha .." + Str(pFolha), ptamanho
   End If
   If pNomedoRelatorio <> "" Then
      Imprime wl_LinhaNome, wl_Coluna - Printer.TextWidth("Emissão:" + CStr(pData)) - IIf(Dir(PathPadrao + "ITUVEPLAST.SYS") = "", 0.5, 1), "Emissão: " + CStr(Date), ptamanho
      pLinha = pLinha + 0.5
   End If
   pLinha = pLinha + 1
Else
   wl_Tamanho = 0
   For i = 0 To pElementos
      If Len(pCabecalho(0, i)) > Len(pReferencia(0, i)) Then
         wl_Tamanho = wl_Tamanho + Len(pCabecalho(0, i)) + 2
      Else
         wl_Tamanho = wl_Tamanho + Len(pReferencia(0, i)) + 2
      End If
   Next
   wl_Tamanho = wl_Tamanho - 3
   If pNomedoRelatorio <> "" Then
      Imprime pLinha, 0, pb_RAZAOSOCIAL, ptamanho
      Imprime pLinha, wl_Tamanho - Len("Folha .." + Str(pFolha)) - 4, "Folha .." + Str(pFolha), ptamanho
      pLinha = pLinha + 0.5
      Imprime pLinha, 0, pNomedoRelatorio, ptamanho
      Imprime pLinha, wl_Tamanho - Len("Emissao .." + CStr(Date)) - 2, "Emissao:" + CStr(pData), ptamanho
      pLinha = pLinha + 0.5
   End If
   wl_Coluna = 0
   For i = 0 To pElementos
      If Len(pReferencia(0, i)) > Len(pCabecalho(0, i)) Then
         Imprime pLinha, wl_Coluna, pCabecalho(0, i), ptamanho, , , pb_PadraoVideo
         wl_Coluna = wl_Coluna + Len(pReferencia(0, i)) + 2
      Else
         Imprime pLinha, wl_Coluna, pCabecalho(0, i), ptamanho, , , pb_PadraoVideo
         wl_Coluna = wl_Coluna + Len(pCabecalho(0, i)) + 2
      End If
   Next
   pLinha = pLinha + 0.5
   wl_Coluna = 0
   For i = 0 To pElementos
      If Len(pReferencia(0, i)) > Len(pCabecalho(0, i)) Then
         Imprime pLinha, wl_Coluna, String(Len(pReferencia(0, i)), "-"), ptamanho, , , pb_PadraoVideo
         aadd pb_LinhadeImpressao, Array(wl_Coluna, wl_Coluna + Len(pReferencia(0, i)))
         wl_Coluna = wl_Coluna + Len(pReferencia(0, i)) + 2
      Else
         Imprime pLinha, wl_Coluna, String(Len(pCabecalho(0, i)), "-"), ptamanho, , , pb_PadraoVideo
         aadd pb_LinhadeImpressao, Array(wl_Coluna, wl_Coluna + Len(pCabecalho(0, i)))
         wl_Coluna = wl_Coluna + Len(pCabecalho(0, i)) + 2
      End If
   Next
   pLinha = pLinha + 0.5
End If
End Sub


Sub Monta_LinhadeImpressao(pLinha, pTexto, pColuna, Optional pAlinhamento = "E", Optional ptamanho As Imp_Constantes = 1, Optional pGuia As Boolean = False)
Dim wl_ImpressaoMatricial As Boolean
wl_ImpressaoMatricial = Mid(pb_Impressao_Normal, 1, 5) = "Draft"
If pAlinhamento = "D" Then
   If Not wl_ImpressaoMatricial And Not pb_PadraoVideo Then
      Imprime pLinha, pb_LinhadeImpressao(pColuna, 1) - Printer.TextWidth(pTexto), pTexto, ptamanho
   Else
      If pTexto = "" And pb_PadraoVideo Then
         pTexto = String(pb_LinhadeImpressao(pColuna, 1) - pb_LinhadeImpressao(pColuna, 0), "_")
      End If
      Imprime pLinha, pb_LinhadeImpressao(pColuna, 1) - Len(pTexto), IIf(Len(pTexto) > pb_LinhadeImpressao(pColuna, 1), Mid(pTexto, 1, Len(pb_LinhadeImpressao(pColuna, 1))), pTexto), ptamanho, , , pb_PadraoVideo
   End If
Else
   If Not wl_ImpressaoMatricial And Not pb_PadraoVideo Then
      Imprime pLinha, pb_LinhadeImpressao(pColuna, 0), pTexto, ptamanho
   Else
      If Len(pTexto) > pb_LinhadeImpressao(pColuna, 1) - pb_LinhadeImpressao(pColuna, 0) Then
   '   If Len(pTexto) > pb_LinhadeImpressao(pColuna, 1) - pb_LinhadeImpressao(pColuna, 0) Then
         Imprime pLinha, pb_LinhadeImpressao(pColuna, 0), Mid(pTexto, 1, pb_LinhadeImpressao(pColuna, 1) - pb_LinhadeImpressao(pColuna, 0)), ptamanho, , , pb_PadraoVideo
      Else
         Imprime pLinha, pb_LinhadeImpressao(pColuna, 0), pTexto, ptamanho, , , pb_PadraoVideo
      End If
   End If
End If
If pGuia And pColuna > 0 Then
   If Not pb_PadraoVideo Then
      Linha pLinha + 0.3, pb_LinhadeImpressao(pColuna - 1, 0), pLinha + 0.3, pb_LinhadeImpressao(pColuna, 1)
      Linha 1.5, pb_LinhadeImpressao(pColuna, 1), pLinha + 0.3, pb_LinhadeImpressao(pColuna, 1)
   End If
End If
End Sub

Sub Linha(pLinhaInicial, Optional pColunaInicial = 0, Optional pLinhaFinal = 0, Optional pColunaFinal = 0, Optional ptamanho As Imp_Constantes = 1)
Dim wl_ImpressaoMatricial As Boolean
If ptamanho = imp_Condensado Or ptamanho = imp_Condensado_NEGRITO Then
   If pColunaFinal = 0 Then pColunaFinal = 136
ElseIf ptamanho = imp_Expandido Or ptamanho = imp_Expandido_Negrito Then
   If pColunaFinal = 0 Then pColunaFinal = 40
ElseIf ptamanho = Imp_Normal Or ptamanho = imp_Normal_Negrito Then
   If pColunaFinal = 0 Then pColunaFinal = 80
End If
wl_ImpressaoMatricial = Mid(pb_Impressao_Normal, 1, 5) = "Draft"
If Not wl_ImpressaoMatricial And Not pb_PadraoVideo Then
   Printer.Line (pColunaInicial + 1, pLinhaInicial)-(pColunaFinal + 1, pLinhaFinal)
Else
    Imprime pLinhaInicial, pColunaInicial, String(IIf(pb_PadraoVideo, 123, pColunaFinal) - pColunaInicial, "-"), ptamanho
End If
End Sub

Sub Imprime(ByVal pLinha, pColuna, pTexto, Optional ptamanho As Imp_Constantes = 1, Optional pAlinhamento As String = "", Optional pColunaFormulario As Integer = 0, Optional pMontaCabecalhoRTF As Boolean = False)
On Error Resume Next
Dim i As Integer
Dim wl_salto As String
Dim wl_CaracterAtiva As String
Dim wl_CaracterDesativa As String
Static wl_UltimoAtiva As String
Static wl_UltimoDesativa As String
Dim wl_ImpressaoMatricial As Boolean
Static wl_BufferReal
Dim wl_correcao As Boolean
Dim wl_IndiceColuna As Currency
pTexto = CStr(pTexto)
pMontaCabecalhoRTF = pb_PadraoVideo
wl_ImpressaoMatricial = Mid(pb_Impressao_Normal, 1, 5) = "Draft"
If wl_ImpressaoMatricial Or pb_PadraoVideo Then
   pTexto = StrTran(pTexto, "á", "a")
   pTexto = StrTran(pTexto, "à", "a")
   pTexto = StrTran(pTexto, "Á", "A")
   pTexto = StrTran(pTexto, "À", "A")
   pTexto = StrTran(pTexto, "é", "e")
   pTexto = StrTran(pTexto, "è", "e")
   pTexto = StrTran(pTexto, "É", "E")
   pTexto = StrTran(pTexto, "È", "E")
   pTexto = StrTran(pTexto, "í", "i")
   pTexto = StrTran(pTexto, "ì", "i")
   pTexto = StrTran(pTexto, "Í", "I")
   pTexto = StrTran(pTexto, "Ì", "I")
   pTexto = StrTran(pTexto, "ó", "o")
   pTexto = StrTran(pTexto, "ò", "o")
   pTexto = StrTran(pTexto, "Ó", "O")
   pTexto = StrTran(pTexto, "Ò", "O")
   pTexto = StrTran(pTexto, "ú", "u")
   pTexto = StrTran(pTexto, "ù", "u")
   pTexto = StrTran(pTexto, "Ú", "U")
   pTexto = StrTran(pTexto, "Ù", "U")
   pTexto = StrTran(pTexto, "ä", "a")
   pTexto = StrTran(pTexto, "ë", "e")
   pTexto = StrTran(pTexto, "ï", "i")
   pTexto = StrTran(pTexto, "ö", "o")
   pTexto = StrTran(pTexto, "ü", "u")
   pTexto = StrTran(pTexto, "Ä", "A")
   pTexto = StrTran(pTexto, "Ë", "E")
   pTexto = StrTran(pTexto, "Ï", "I")
   pTexto = StrTran(pTexto, "Ö", "O")
   pTexto = StrTran(pTexto, "Ü", "U")
   pTexto = StrTran(pTexto, "ã", "a")
   pTexto = StrTran(pTexto, "õ", "o")
   pTexto = StrTran(pTexto, "Ã", "A")
   pTexto = StrTran(pTexto, "Õ", "O")
   pTexto = StrTran(pTexto, "Ç", "C")
   pTexto = StrTran(pTexto, "ç", "c")
   pTexto = StrTran(pTexto, "â", "a")
   pTexto = StrTran(pTexto, "ê", "e")
   pTexto = StrTran(pTexto, "î", "i")
   pTexto = StrTran(pTexto, "ô", "o")
   pTexto = StrTran(pTexto, "û", "u")
   pTexto = StrTran(pTexto, "Â", "A")
   pTexto = StrTran(pTexto, "Ê", "E")
   pTexto = StrTran(pTexto, "Î", "I")
   pTexto = StrTran(pTexto, "Ô", "O")
   pTexto = StrTran(pTexto, "Û", "U")
End If
If Not pb_PadraoVideo Then
   If Not wl_ImpressaoMatricial Then
      Printer.ScaleMode = 7
      Printer.FontBold = False
      pb_ImpressaoIniciada = True
      If pLinha < 0 Then pLinha = 0
      If pLinha >= Printer.ScaleHeight Then
         Printer.NewPage
         pLinha = pLinha - Printer.ScaleHeight
      ElseIf pLinha + Printer.TextHeight(pTexto) > Printer.ScaleHeight Then
         wl_correcao = True
      End If
   End If
End If
If ptamanho = imp_Condensado Then
   If Not pb_PadraoVideo Then
      If Not wl_ImpressaoMatricial Then
         Printer.FontName = pb_Impressao_Condensada
         Printer.FontSize = 8
      End If
      wl_CaracterAtiva = pb_LigaCompactado
      wl_CaracterDesativa = pb_Desligacompactado
   Else
      wl_CaracterAtiva = "\fs12 "
      wl_CaracterDesativa = "\fs16"
   End If
ElseIf ptamanho = Imp_Normal Then
   If Not pb_PadraoVideo Then
      If Not wl_ImpressaoMatricial Then
         Printer.FontName = pb_Impressao_Normal
         Printer.FontSize = 10
      End If
      wl_CaracterAtiva = pb_Reset
      wl_CaracterDesativa = ""
   Else
      wl_CaracterAtiva = "\fs16 "
      wl_CaracterDesativa = "\fs16 "
   End If
ElseIf ptamanho = imp_Expandido Then
   If Not pb_PadraoVideo Then
      If Not wl_ImpressaoMatricial Then
         Printer.FontName = pb_Impressao_Expandida
         Printer.FontSize = 14
      End If
      wl_CaracterAtiva = pb_LigaExpandido
      wl_CaracterDesativa = pb_DesligaExpandido
   Else
      wl_CaracterAtiva = "\fs28 "
      wl_CaracterDesativa = "\fs16 "
   End If
ElseIf ptamanho = imp_Condensado_NEGRITO Then
   If Not pb_PadraoVideo Then
      If Not wl_ImpressaoMatricial Then
         Printer.FontName = pb_Impressao_Condensada_N
         Printer.FontSize = 8
         Printer.FontBold = True
      End If
      wl_CaracterAtiva = pb_LigaCompactado + pbliganegrito
      wl_CaracterDesativa = pb_DesligaNegrito + pb_Desligacompactado
   Else
      wl_CaracterAtiva = "\b\fs12 "
      wl_CaracterDesativa = "\b0\fs16 "
   End If
ElseIf ptamanho = imp_Normal_Negrito Then
   If Not pb_PadraoVideo Then
      If Not wl_ImpressaoMatricial Then
         Printer.FontName = pb_Impressao_Normal_N
         Printer.FontSize = 10
         Printer.FontBold = True
      End If
      wl_CaracterAtiva = pb_LigaNegrito
      wl_CaracterDesativa = pb_DesligaNegrito
   Else
      wl_CaracterAtiva = "\b\fs16 "
      wl_CaracterDesativa = "\b0\fs16 "
   End If
ElseIf ptamanho = imp_Expandido_Negrito Then
   If Not pb_PadraoVideo Then
      If Not wl_ImpressaoMatricial Then
         Printer.FontName = pb_Impressao_Expandida_N
         Printer.FontSize = 14
         Printer.FontBold = True
      End If
      wl_CaracterAtiva = pb_LigaExpandido + pb_LigaNegrito
      wl_CaracterDesativa = pb_DesligaNegrito + pb_DesligaExpandido
   Else
      wl_CaracterAtiva = "\b\fs28 "
      wl_CaracterDesativa = "\b0\fs16 "
   End If
End If
If pAlinhamento <> "" Then
   If pAlinhamento = "D" Then
      If wl_ImpressaoMatricial Or pb_PadraoVideo Then
         If ptamanho = imp_Condensado Or ptamanho = imp_Condensado_NEGRITO Then
            pColuna = IIf(pColunaFormulario = 0, IIf(pb_PadraoVideo, 124, 135), pColunaFormulario) - Len(pTexto)
         ElseIf ptamanho = Imp_Normal Or ptamanho = imp_Normal_Negrito Then
            pColuna = IIf(pColunaFormulario = 0, IIf(pb_PadraoVideo, 124, 80), pColunaFormulario) - Len(pTexto)
         ElseIf ptamanho = imp_Expandido Or ptamanho = imp_Expandido_Negrito Then
            pColuna = IIf(pColunaFormulario = 0, 40, pColunaFormulario) - Len(pTexto)
         End If
      Else
         pColuna = Printer.ScaleWidth - Printer.TextWidth(pTexto)
      End If
   ElseIf pAlinhamento = "E" Then
      pColuna = 0
   ElseIf pAlinhamento = "C" Then
      If wl_ImpressaoMatricial Or pb_PadraoVideo Then
         If ptamanho = imp_Condensado Or ptamanho = imp_Condensado_NEGRITO Then
            pColuna = IIf(pb_PadraoVideo, 60, 68) - Len(pTexto) / 2
         ElseIf ptamanho = Imp_Normal Or ptamanho = imp_Normal_Negrito Then
            pColuna = IIf(pb_PadraoVideo, 60, 40) - Len(pTexto) / 2
         ElseIf ptamanho = imp_Expandido Or ptamanho = imp_Expandido_Negrito Then
            pColuna = IIf(pb_PadraoVideo, 40, 20) - Len(pTexto) / 2
         End If
      Else
         pColuna = ((Printer.ScaleWidth) / 2) - (Printer.TextWidth(pTexto) / 2)
      End If
      If pColuna < 0 Then
         pTexto = Mid(pTexto, (pColuna * -1) + 1)
         pColuna = 0
      End If
   End If
End If
If Not pb_PadraoVideo Then
   If Not wl_ImpressaoMatricial Then
      Printer.CurrentX = pColuna + 1
      Printer.CurrentY = pLinha
      If wl_correcao Then
         Printer.Print pTexto;
      Else
         Printer.Print pTexto
      End If
   Else
      pLinha = pLinha * 2
      If pLinha = pb_LinhaBuffer Then
         pb_Buffer = pb_Buffer + Space(pColuna - Len(RTrim(pb_Buffer))) + pTexto
         pb_LinhaBuffer = pLinha
         pb_Tamanho = ptamanho
         wl_UltimoAtiva = wl_CaracterAtiva
         wl_UltimoDesativa = wl_CaracterDesativa
         Exit Sub
      Else
         pb_Buffer = wl_UltimoAtiva + pb_Buffer + wl_UltimoDesativa
      End If
      If pb_LinhaBuffer < pb_UltimaLinha Then
         Salta_Pagina
      End If
      For i = pb_UltimaLinha + 1 To pb_LinhaBuffer - 1
         Print #1, ""
      Next
      Print #1, pb_Buffer
      pb_UltimaLinha = pb_LinhaBuffer
      pb_LinhaImpressaoMatricial = pLinha
      pb_ColunaImpressaoMatricial = 0
      pb_Buffer = Space(pColuna) + pTexto
      pb_LinhaBuffer = pLinha
      wl_UltimoAtiva = wl_CaracterAtiva
      wl_UltimoDesativa = wl_CaracterDesativa
   End If
Else
   pLinha = pLinha * 2
   If pLinha = pb_LinhaBuffer Then
      If pMontaCabecalhoRTF Then
         wl_IndiceColuna = pColuna - Len(RTrim(wl_BufferReal))
         wl_BufferReal = wl_BufferReal + String(wl_IndiceColuna, " ") + pTexto
         pb_Buffer = pb_Buffer + String(wl_IndiceColuna, " ") + wl_CaracterAtiva + pTexto + wl_CaracterDesativa
      Else
         pb_Buffer = pb_Buffer + String(pColuna - Len(RTrim(pb_Buffer)), " ") + wl_CaracterAtiva + pTexto + wl_CaracterDesativa
      End If
      pb_LinhaBuffer = pLinha
      Exit Sub
   End If
   For i = pb_UltimaLinha To pb_LinhaBuffer - 1
      Print #1, "\par "
   Next
   Print #1, pb_Buffer
   pb_UltimaLinha = pb_LinhaBuffer
   pb_LinhaImpressaoMatricial = pLinha
   pb_ColunaImpressaoMatricial = 0
   pb_Buffer = String(pColuna, " ") + wl_CaracterAtiva + pTexto + wl_CaracterDesativa
   wl_BufferReal = String(pColuna, " ") + pTexto
   pb_LinhaBuffer = pLinha
End If
End Sub

Function BuscaPreferencia(pSecao As String, pChave As String, pRetorno As String) As Boolean
BuscaPreferencia = RetornaConfiguracao(pSecao, pChave) = pRetorno
End Function




Function Grava_Configuracoes(pSecao As String, pChave As String, pValor As String, Optional pFile As String = "")
Dim wl_File As String
If pFile = "" Then
   wl_File = PathWindows + pb_Sistema + ".INI"
Else
   wl_File = PathWindows + pFile + ".INI"
End If
write_ini wl_File, pSecao, pChave, pValor
End Function





   
Sub AddIndex(ByRef pMatriz, pNome As String, pCampos As Variant, Optional pPrimaryKey As Boolean = False, Optional pUnique As Boolean = False, Optional pRequired As Boolean = False)
If IsArray(pCampos) Then
   aadd pMatriz, Array(pNome, pCampos, pPrimaryKey, pUnique, pRequired)
Else
   aadd pMatriz, Array(pNome, Array(pCampos), pPrimaryKey, pUnique, pRequired)
End If
End Sub
Sub AddField(ByRef pMatrizCampo, pNome, pTipo As dbCONSTANTES, Optional ptamanho As Integer = 0, Optional pRequired As Boolean = False, Optional pAllowLenghtZero As Boolean = False)
aadd pMatrizCampo, Array(pNome, pTipo, ptamanho, pRequired, pAllowLenghtZero)
End Sub

Public Function Calc_CGC(VALOR As String) As Boolean
Dim Mult1 As String
Dim Mult2 As String
Dim dig1 As Integer
Dim dig2 As Integer
Dim X As Integer
Mult1 = "543298765432"
Mult2 = "6543298765432"

For X = 1 To 12
    dig1 = dig1 + (Val(Mid$(VALOR, X, 1)) * Val(Mid$(Mult1, X, 1)))
Next

For X = 1 To 13
    dig2 = dig2 + (Val(Mid$(VALOR, X, 1)) * Val(Mid$(Mult2, X, 1)))
Next
dig1 = (dig1 * 10) Mod 11
dig2 = (dig2 * 10) Mod 11

If dig1 = 10 Then dig1 = 0
If dig2 = 10 Then dig2 = 0

Calc_CGC = True
If dig1 <> Val(Mid$(VALOR, 13, 1)) Then Calc_CGC = False
If dig2 <> Val(Mid$(VALOR, 14, 1)) Then Calc_CGC = False

End Function
Function Calc_CPF(VALOR As String) As Boolean
Dim dig1 As Integer
Dim dig2 As Integer
Dim Mult1 As Integer
Dim Mult2 As Integer
Dim X As Integer
Mult1 = 10
Mult2 = 11
   
For X = 1 To 9
    dig1 = dig1 + (Val(Mid$(VALOR, X, 1)) * Mult1)
    Mult1 = Mult1 - 1
Next
   
For X = 1 To 10
    dig2 = dig2 + (Val(Mid$(VALOR, X, 1)) * Mult2)
    Mult2 = Mult2 - 1
Next
   
dig1 = (dig1 * 10) Mod 11
dig2 = (dig2 * 10) Mod 11
If dig1 = 10 Then dig1 = 0
If dig2 = 10 Then dig2 = 0
   
Calc_CPF = True
If Val(Mid$(VALOR, 10, 1)) <> dig1 Then Calc_CPF = False
If Val(Mid$(VALOR, 11, 1)) <> dig2 Then Calc_CPF = False
End Function
Function LadoaLado(pForm As Form, pImage As Object)
Dim X
Dim Y
Do While True
pForm.PaintPicture pImage.Picture, X, Y
X = X + pImage.Width
If X > pForm.Width Then
   Y = Y + pImage.Height
   X = 0
   If Y > pForm.Height Then
      Exit Function
   End If
   DoEvents
End If
Loop
End Function

Function RetornaRegistroWindows(pSecao As String, pChave As String)
RetornaRegistroWindows = GetSetting(pb_Sistema, pSecao, pChave)
End Function

Function write_ini(Arquiv$, ByVal section$, ByVal chv$, ByVal Variavel$) As String
iRet = WritePrivateProfileString(ByVal section$, ByVal chv$, ByVal Variavel$, ByVal Arquiv$)
End Function

Function get_ini(pArquiv$, pSecao$, pChave$) As String
Dim wl_Retorno$, wl_Tamanho As Long
wl_Retorno = Space(128)
X% = GetPrivateProfileString(pSecao, pChave, "", wl_Retorno, Len(wl_Retorno), pArquiv)
get_ini = Left$(wl_Retorno$, X%)
End Function

Function aOpen(pBancodeDados As String, pTabela As String, ByRef pObjeto_Database As Database, ByRef pObjeto_Recordset As Recordset, Optional pEXCLUSIVO = False)
On Error Resume Next
Dim i As Integer
Dim wl_ExisteTabela As Boolean
Dim wl_contador As Integer
pBancodeDados = PathPadrao + pBancodeDados
If Dir(pBancodeDados) = "" Then
   Call MsgBox("Impossível encontrar " + pBancodeDados, vbCritical, "Mensagem do Sistema")
   aOpen = False
   Exit Function
End If
Err = 0
DisplayMensagem "Abrindo: " + pBancodeDados + " / Tabela: " + pTabela
Set pObjeto_Database = OpenDatabase(pBancodeDados)
wl_contador = 0
Do While Err <> 0
   wl_contador = wl_contador + 1
   If wl_contador > 10 Then
      Call MsgBox("Erro ao abrir o banco de dados", vbCritical, Str(Err.Number) + " - " + Err.Description)
      aOpen = False
      Exit Function
   End If
   Err = 0
   Set pObjeto_Database = OpenDatabase(pBancodeDados)
Loop
wl_ExisteTabela = False
For i = 0 To pObjeto_Database.TableDefs.Count - 1
   If UCase(pObjeto_Database.TableDefs(i).Name) = UCase(pTabela) Then
      wl_ExisteTabela = True
      Exit For
   End If
Next
If Not wl_ExisteTabela Then
   Call MsgBox("Impossível encontrar tabela " + pTabela + " .", vbExclamation, "Banco de Dados: " + pBancodeDados)
   aOpen = False
   Exit Function
End If
Err = 0
If Not pEXCLUSIVO Then
   Set pObjeto_Recordset = pObjeto_Database.OpenRecordset(pTabela, Table)
Else
   Set pObjeto_Recordset = pObjeto_Database.OpenRecordset(pTabela, , dbDenyWrite + dbDenyRead)
End If
wl_contador = 0
Do While Err <> 0
   wl_contador = wl_contador + 1
   If wl_contador > 10 Then
      aOpen = False
      Call MsgBox("Erro ao abrir a tabela " + pTabela, vbInformation, Str(Err.Number) + " - " + Err.Description)
      Exit Function
   End If
   Err = 0
   If Not pEXCLUSIVO Then
      Set pObjeto_Recordset = pObjeto_Database.OpenRecordset(pTabela)
   Else
      Set pObjeto_Recordset = pObjeto_Database.OpenRecordset(pTabela, , dbDenyWrite + dbDenyRead)
   End If
Loop
Err = 0
If pObjeto_Database.TableDefs(pTabela).Indexes.Count > 0 Then
   pObjeto_Recordset.Index = pObjeto_Database.TableDefs(pTabela).Indexes(0).Name
End If
If Err <> 0 Then
   Call MsgBox("Erro ao acessar o índice", vbCritical, "Mensagem do Sistema")
   aOpen = False
   Exit Function
End If
aOpen = True
End Function




Function Extenso(vlr As Currency, Optional pMoeda As Boolean = True)
Dim Formato As String
Dim i, z, wl_Espacos As Integer
Dim wl_Falta As Integer
Dim wl_Extenso As String
Dim wl_Linha1 As String
Dim wl_Linha2 As String

Formato = Format(vlr, "000,000,000,000,000.00")
Extenso = ""
wl_Extenso = Trim(ExtNivel(Mid(Formato, 1, 3)))
If wl_Extenso <> "" Then
   Extenso = Extenso + wl_Extenso + IIf(Extenso = "Hum", " Trilhao", " Trilhoes") + ", "
End If

wl_Extenso = Trim(ExtNivel(Mid(Formato, 5, 3)))
If wl_Extenso <> "" Then
   Extenso = Extenso + wl_Extenso + IIf(Extenso = "Hum", " Bilhao", " Bilhoes") + ", "
End If

wl_Extenso = Trim(ExtNivel(Mid(Formato, 9, 3)))
If wl_Extenso <> "" Then
   Extenso = Extenso + wl_Extenso + IIf(Extenso = "Hum", " Milhao", " Milhoes") + ", "
End If

wl_Extenso = Trim(ExtNivel(Mid(Formato, 13, 3)))
If wl_Extenso <> "" Then
   Extenso = Extenso + wl_Extenso + " Mil, "
End If

wl_Extenso = Trim(ExtNivel(Mid(Formato, 17, 3)))
If wl_Extenso <> "" Then
   Extenso = Extenso + wl_Extenso
End If
Extenso = Trim(Extenso)
If Mid(Extenso, Len(Extenso), 1) = "," Then
   Extenso = Mid(Extenso, 1, Len(Extenso) - 1)
End If
If Extenso <> "" And pMoeda Then
   If Extenso = "Hum" Then
      Extenso = Extenso + " Real"
   Else
      Extenso = Extenso + " Reais"
   End If
End If
wl_Extenso = Trim(ExtNivel("0" + Mid(Formato, 21, 2)))
If wl_Extenso <> "" And pMoeda Then
   Extenso = Extenso + IIf(Extenso <> "", " e ", "") + wl_Extenso + IIf(wl_Extenso = "Hum", " Centavo", " Centavos")
End If
Extenso = RTrim(Extenso)
End Function

Function ExtNivel(vlr)
Dim Formato As String
Dim wl_Extenso As String
Dim wl_Analise
Dim wl_Unidades, wl_Dez, wl_Dezenas, wl_Centenas, wl_Cento, wl_milhar, wl_milhao, wl_Bilhao, wl_Trilhao
wl_Cento = "Cem"
aadd wl_Unidades, Array("Um", "Dois", "Tres", "Quatro", "Cinco", "Seis", "Sete", "Oito", "Nove")
aadd wl_Dez, Array("Onze", "Doze", "Treze", "Quatorze", "Quinze", "Dezesseis", "Dezessete", "Dezoito", "Dezenove")
aadd wl_Dezenas, Array("Dez", "Vinte", "Trinta", "Quarenta", "Cinquenta", "Sessenta", "Setenta", "Oitenta", "Noventa")
aadd wl_Centenas, Array("Cem", "Duzentos", "Trezentos", "Quatrocentos", "Quinhentos", "Seiscentos", "Setecentos", "Oitocentos", "Novecentos")
aadd wl_milhar, Array("Mil")
aadd wl_milhao, Array("Milhao", "Milhoes")
aadd wl_Bilhao, Array("Bilhao", "Bilhoes")
aadd wl_Trilhao, Array("Trilhao", "Trilhoes")


Formato = Format(vlr, "000")
wl_Extenso = ""
wl_Analise = Mid(Formato, 1, 1)
If Mid(wl_Analise, 1, 1) > 0 Then
   If Mid(wl_Analise, 1, 1) = "1" And Mid(Formato, 2, 2) <> "00" Then
      wl_Extenso = wl_Extenso + "Cento e "
   Else
      wl_Extenso = wl_Extenso + wl_Centenas(0, Val(wl_Analise) - 1) + " e "
   End If
End If
wl_Analise = Mid(Formato, 2, 1)
If Mid(wl_Analise, 1, 1) > 0 Then
   If wl_Analise = "1" And Mid(Formato, 3, 1) <> "0" Then
      wl_Extenso = wl_Extenso + wl_Dez(0, Val(Mid(Formato, 3, 1) - 1)) + " , "
   Else
      wl_Extenso = wl_Extenso + wl_Dezenas(0, Val(wl_Analise) - 1) + " e "
   End If
End If
wl_Analise = Mid(Formato, 3, 1)
If Mid(wl_Analise, 1, 1) > 0 Then
   If Mid(Formato, 2, 1) <> "1" Then
      wl_Extenso = wl_Extenso + wl_Unidades(0, Val(wl_Analise) - 1) + " "
   End If
End If
ExtNivel = wl_Extenso
End Function

Function CarregaDriver(pPrinter As String)
Dim lPrinter As String
lPrinter = UCase(pPrinter)
For Each p In Printers
   If UCase(p.DeviceName) Like "*" + lPrinter + "*" Then
      Set Printer = p
      Exit For
   End If
Next p
End Function

Function EhNumero(key As Integer)
Dim car As String
car = "-0123456789,"
If key = 46 Then
   SendKeys ","
End If
If InStr(car, Chr(key)) = 0 And key <> 8 Then
   EhNumero = False
Else
   EhNumero = True
End If
End Function





Function FormataData(wp_Data)
FormataData = Mid(wp_Data, 1, 2) + "/" + Mid(wp_Data, 3, 2) + "/" + Mid(wp_Data, 5)
End Function


Function NomedoMes(pMes As Integer, Optional pInicial As Boolean = False)
ReDim aMes(0)
aadd aMes, Array("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro")
NomedoMes = IIf(pInicial, Mid(aMes(0, pMes - 1), 3), aMes(0, pMes - 1))
End Function


Sub Seleciona()
HomeEnd
End Sub

Function Video(titulo, matriz, BancodeDados, Tabela, Optional indice, Optional CampodeRetorno = "", Optional CampodePesquisa = "")
On Error Resume Next
Dim i As Integer
ReDim a_Browse001(0)
tamanho = 0
aviso "Aguarde. Montando Consulta ..."
pb_FormAtivo = VB.Screen.ActiveForm.Name
pb_ObjetoAtivo = VB.Screen.ActiveForm.ActiveControl.Name
For i = 0 To UBound(matriz)
   Err = 0
   aadd a_Browse001, Array(matriz(i, 0), matriz(i, 1), matriz(i, 2), matriz(i, 3), matriz(i, 4))
   If Err <> 0 Then
      aadd a_Browse001, Array(matriz(i, 0), matriz(i, 1), matriz(i, 2), matriz(i, 3), 0)
   End If
Next

BancodeDados = PathPadrao + BancodeDados

If Dir(BancodeDados) = "" Then
   Call MsgBox("Não consigo encontrar o Banco de Dados " + BancodeDados, vbCritical, "Mensagem do Sistema")
   pbRetornoVideo = ""
   aviso
   Exit Function
End If
BancodeDadosdaConsulta = BancodeDados
TabeladaConsulta = Tabela
If Not IsMissing(indice) Then
   IndicedaConsulta = indice
Else
   IndicedaConsulta = ""
   If CampodePesquisa <> "" Then pb_CampodePesquisa = CampodePesquisa
End If
pb_TitulodaConsulta = titulo
pbCampodeRetorno = CampodeRetorno
aviso
frmVideo.Show 1
Video = pbRetornoVideo
End Function


Function DIG(n)
Dim t As Integer
Dim i As Integer
Dim d As Integer
t = 6
s = 0
X = strzero(n, t, " ")
For i = 1 To t
   s = s + Val(Mid(X, i, 1)) * 1

Next
d = s Mod 11
DIG = IIf(d = 10, "0", IIf(d = 0, "1", strzero(d, 1, " ")))
End Function



Function cria(ByVal Arquivo, Tabela, campo, Optional indices) As Boolean
Dim db As Database
Dim tb As TableDef
Dim criaind As Boolean
Dim wl_CriaBancodeDados As Boolean
Dim j, X, z, m As Integer
Dim wl_CriaCampo As Boolean
Dim wl_Pasta As String
Dim BancodeDados As String
Dim wl_Field As Field
Dim wl_CriaIndice As Boolean
Dim wl_Idx As New Index
Dim wl_CriaTabela As Boolean
On Error Resume Next
BancodeDados = Arquivo
wl_Pasta = Mid(PathPadrao, 1, Len(PathPadrao) - 1)
Err = 0
If Dir(wl_Pasta, vbDirectory) = "" Then
   MkDir wl_Pasta
   If Err <> 0 Then
      Call MsgBox("Erro ao criar a pasta " + PathPadrao, vbCritical, "Mensagem do Sistema")
      MsgBox Err.Description
      cria = False
      Exit Function
   End If
End If
If Dir(PathPadrao + "COMUM", vbDirectory) = "" Then
   MkDir PathPadrao + "COMUM"
   If Err <> 0 Then
      Call MsgBox("Erro ao criar a pasta " + PathPadrao + "COMUM", vbCritical, "Mensagem do Sistema")
      MsgBox Err.Description
      cria = False
      Exit Function
   End If
End If
ReDim fd(1 To 50)
Err = 0
If Not IsMissing(indices) Then
   criaind = True
Else
   criaind = False
End If
If pb_MontaProgressao Then
   fMENU.Progressao.Value = fMENU.Progressao.Value + 1
End If
Err = 0
BancodeDados = UCase(BancodeDados)
If Dir(BancodeDados) = "" Then
   Set db = CreateDatabase(BancodeDados, DB_LANG_GENERAL)
   wl_CriaBancodeDados = True
Else
   Set db = OpenDatabase(BancodeDados)
   wl_CriaBancodeDados = False
   wl_CriaTabela = False
   For i = 0 To db.TableDefs.Count - 1
      If UCase(db.TableDefs(i).Name) = UCase(Tabela) Then
         GoSub Verifica_Campos
         If criaind Then
            GoSub Verifica_Indices
         End If
         cria = True
         Exit Function
      End If
   Next
End If
aviso
If Err <> 0 Then
   Call MsgBox("Erro ao criar/abrir o banco de dados " + BancodeDados, vbCritical, "Mensagem do Sistema")
   cria = False
   Exit Function
End If
aviso "Criando " + BancodeDados + "." + Tabela
Set tb = db.CreateTableDef(Tabela)
With tb
   For X = 0 To UBound(campo)
      If campo(X, 1) = dbText Or campo(X, 1) = dbMemo Then
         campo(X, 4) = True
      End If
      .Fields.Append .CreateField(campo(X, 0), campo(X, 1), campo(X, 2))
      .Fields(X).Required = campo(X, 3)
      .Fields(X).AllowZeroLength = campo(X, 4)
   Next X
End With
aviso
db.TableDefs.Append tb
db.Close
If Err <> 0 Then
   Call MsgBox("Erro ao criar a tabela " + Tabela, vbCritical, "Mensagem do Sistema")
   Kill BancodeDados
   cria = False
   Exit Function
End If
Set db = OpenDatabase(BancodeDados)
wl_CriaBancodeDados = False
wl_CriaTabela = False
For i = 0 To db.TableDefs.Count - 1
   If UCase(db.TableDefs(i).Name) = UCase(Tabela) Then
      GoSub Verifica_Campos
      If criaind Then
         GoSub Verifica_Indices
      End If
      cria = True
      Exit Function
   End If
Next
aviso
cria = True
Exit Function

Verifica_Campos:
For X = 0 To UBound(campo)
   wl_CriaCampo = True
   For z = 0 To db.TableDefs(i).Fields.Count - 1
      If UCase(db.TableDefs(i).Fields(z).Name) = UCase(campo(X, 0)) Then
         wl_CriaCampo = False
         Exit For
      End If
   Next
   If wl_CriaCampo Then
      GoSub AcrescentaCampo
   End If
Next
Return

AcrescentaCampo:
Set wl_Field = New Field
If campo(X, 1) = dbText Or campo(X, 1) = dbMemo Then
   campo(X, 4) = True
End If
With wl_Field
   .Name = campo(X, 0)
   .Type = campo(X, 1)
   .Size = campo(X, 2)
   .Required = campo(X, 3)
   .AllowZeroLength = campo(X, 4)
   If Err = 3219 Then
      Err = 0
   End If
End With
db.TableDefs(i).Fields.Append wl_Field
If Err <> 0 Then
   aviso
   Call MsgBox("Erro ao acrescentar campo " + campo(X, 0), vbCritical, "Mensagem do Sistema")
   cria = False
   Exit Function
End If
Return

Verifica_Indices:
For X = 0 To UBound(indices)
   wl_CriaIndice = True
   For z = 0 To db.TableDefs(i).Indexes.Count - 1
      If UCase(db.TableDefs(i).Indexes(z).Name) = UCase(indices(X, 0)) Then
         wl_CriaIndice = False
         Exit For
      End If
   Next
   If wl_CriaIndice Then
      GoSub AcrescentaIndice
   End If
Next
Return

AcrescentaIndice:
Err = 0
Set wl_Idx = db.TableDefs(i).CreateIndex(indices(X, 0))
For j = 0 To UBound(indices(X, 1))
   Set fld = wl_Idx.CreateField(indices(X, 1)(j))
   wl_Idx.Fields.Append fld
Next
wl_Idx.Name = indices(X, 0)
wl_Idx.Primary = indices(X, 2)
wl_Idx.Required = indices(X, 3)
wl_Idx.Unique = indices(X, 4)
db.TableDefs(i).Indexes.Append wl_Idx
If Err <> 0 Then
   aviso
   Call MsgBox("Erro ao acrescentar indice", vbCritical, "Indice: " + indices(X, 0) + " tabela: " + db.TableDefs(i).Name)
   cria = False
   Exit Function
End If
Return
End Function
Function RetornaConfiguracao(pSecao As String, pChave As String, Optional pArquivo As String = "") As String
Dim wl_File As String
If pChave = "PathPadrao" And pb_Demonstracao Then
   RetornaConfiguracao = "INFOSOFT_DEMO"
   Exit Function
End If
If pArquivo = "" Then
   wl_File = PathWindows + pb_Sistema + ".ini"
Else
   wl_File = PathWindows + pArquivo
End If
If Dir(wl_File) = "" Then
   RetornaConfiguracao = ""
   Exit Function
End If
RetornaConfiguracao = get_ini(wl_File, pSecao, pChave)
End Function
Function add_reg(ByRef objetotb As Recordset) As Boolean
Dim tentativa As Integer
On Error Resume Next
tentativa = 0
If pb_Demonstracao And objetotb.RecordCount > 100 Then
   Call MsgBox("Essa é uma versão de demonstração, não é possível adicionar mais registros em suas tabelas. Entre em contato com o desenvolvedor para registrar o aplicativo", vbInformation, "Registre sua cópia")
   add_reg = False
   Exit Function
End If
Do While tentativa <= 10
   tentativa = tentativa + 1
   Err = 0
   objetotb.AddNew
   If Err = 0 Then
      Exit Do
   End If
Loop
If Err <> 0 Then
   MsgBox Err.Description, vbInformation, "Informe esta mensagem do programador. Erro:" + Str(Err.Number)
   add_reg = False
Else
   add_reg = True
End If
End Function

Function pausa(Nsegundos As Integer)
Dim t As Integer
If Nsegundos > 26 Then
   Call MsgBox("Pausa máxima permitida: 30 segundos", vbCritical, "Mensagem do Sistema")
   Exit Function
End If
t = Val(Mid(Format(Time, "hh:nn:ss"), 7, 2))
If t + Nsegundos >= 60 Then
   t = (t + Nsegundos) - 60
Else
   t = t + Nsegundos
End If
Do While Mid(Format(Time, "hh:nn:ss"), 7, 2) <> strzero(t, 2)
   DoEvents
Loop
End Function


Function Refresh_reg(ByVal objetotb)
On Error Resume Next
objetotb.Edit
objetotb.Update
End Function

Function update_reg(ByRef objetotb)
Dim tentativa As Integer
On Error Resume Next
tentativa = 0
Do While tentativa <= 10
   tentativa = tentativa + 1
   Err = 0
   objetotb.Update
   If Err = 0 Then
      Exit Do
   End If
Loop
If Err <> 0 Then
   MsgBox Err.Description, vbInformation, "Informe esta mensagem ao programador. Erro: " + Str(Err.Number)
   update_reg = False
Else
   update_reg = True
End If
End Function


Function edit_reg(ByRef objetotb As Recordset) As Boolean
Dim tentativa As Integer
On Error Resume Next
tentativa = 0
Do While tentativa <= 10
   tentativa = tentativa + 1
   Err = 0
   objetotb.Edit
   If Err = 0 Then
      Exit Do
   End If
Loop
If Err <> 0 Then
   MsgBox "Erro ao tentar editar o registro", vbInformation, "Tabela : " + objetotb.Name
   edit_reg = False
Else
   edit_reg = True
End If
End Function


Function StrTran(pString, pTroca, pPor)
Dim i As Integer
Dim w As String
If InStr(pString, pTroca) = 0 Then
   StrTran = pString
   Exit Function
End If
For i = 1 To Len(pString)
   If Mid(pString, i, Len(pTroca)) = pTroca Then
      w = w + pPor
      i = i + Len(pTroca) - 1
   Else
      w = w + Mid(pString, i, 1)
   End If
Next
StrTran = w
End Function




Function CentraForm(formulario)
formulario.Left = (Screen.Width - formulario.Width) / 2
formulario.Top = (Screen.Height - formulario.Height) / 2
End Function


Function ctox(c)
Dim i As Integer
Dim char As String
Dim Retorno As String
Retorno = ""
For i = 1 To Len(c)
   char = Asc(Mid(c, i, 1)) + i * 10
   Do While char > 255
      char = char - 255
   Loop
   Retorno = Retorno + Chr$(char)
Next i
ctox = Retorno
End Function

Function ctox_old(c)
Dim i As Integer
Dim char As String
Dim Retorno As String
Retorno = ""
For i = 1 To Len(c)
   char = Asc(Mid(c, i, 1)) + i
   Retorno = Retorno + Chr$(char)
Next i
ctox_old = Retorno
End Function


Function Ultimo_DiasdoMes(pMes As String, pANO As String) As Date
Dim XDATA As Date

If Val(pMes) = 12 Then
   XDATA = CDate("01/01/" + Format(Val(pANO) + 1, "00"))
 Else
   XDATA = CDate("01/" + Format(Val(pMes) + 1, "00") + "/" + pANO)
End If

XDATA = XDATA - 1

Ultimo_DiasdoMes = XDATA

End Function

Function DiadaSemana(data)
Dim dia As Integer

dia = Weekday(data)

If dia = 1 Then
   DiadaSemana = "Dom"
ElseIf dia = 2 Then
   DiadaSemana = "Seg"
ElseIf dia = 3 Then
   DiadaSemana = "Ter"
ElseIf dia = 4 Then
   DiadaSemana = "Qua"
ElseIf dia = 5 Then
   DiadaSemana = "Qui"
ElseIf dia = 6 Then
   DiadaSemana = "Sex"
ElseIf dia = 7 Then
   DiadaSemana = "Sab"
End If


End Function


Function repl(char, qtos)
Dim X As String
X = ""
For i = 1 To qtos
   X = X + char
Next i
repl = X
End Function

Function RetornaFormatado(inteiro, finteiro)
Dim palavra As String
Dim formatando As String
Dim z As Integer
Dim i As Integer
Dim reversao As String

palavra = strzero(inteiro, Len(finteiro), "*")
formatando = ""
z = Len(palavra)
For i = Len(finteiro) To 1 Step -1
   If Mid(finteiro, i, 1) = "#" Then
      If Mid(palavra, z, 1) <> "*" Then
         formatando = formatando + Mid(palavra, z, 1)
         z = z - 1
      Else
         z = z - 1
      End If
   Else
      formatando = formatando + Mid(finteiro, i, 1)
   End If
Next i
reversao = ""
For i = Len(formatando) To 1 Step -1
   reversao = reversao + Mid(formatando, i, 1)
Next i
reversao = LTrim(reversao)
Do While Asc(Mid(reversao, 1, 1)) < 49 Or Asc(Mid(reversao, 1, 1)) > 57
   reversao = Mid(reversao, 2)
Loop
RetornaFormatado = reversao
End Function

Function AT(char, palavra) As Integer

Dim i As Integer
For i = 1 To Len(palavra)
If Mid(palavra, i, Len(char)) = char Then
   AT = i
   Exit Function
End If
Next i
AT = 0
End Function

Function cdpasta(pasta)
Attribute cdpasta.VB_Description = "funciona exatamente como o comando CD do DOS"
On Error Resume Next
Err = 0
ChDir pasta
If Err <> 0 Then
   MsgBox "Não consigo acessar a pasta " + pasta, vbExclamation, "Atenção"
   cdpasta = False
Else
   cdpasta = True
End If
End Function


Function mkpasta(pasta)
On Error Resume Next
Err = 0
MkDir pasta
If Err <> 0 Then
   Err = 0
   Exit Function
End If
End Function




Function existepasta(pasta)
On Error Resume Next
Dim atual As String
atual = CurDir
ChDir pasta
If Err <> 0 Then
   existepasta = False
Else
   existepasta = True
End If
ChDir atual
End Function








Function ctod(dt)
ctod = CVDate(dt)
End Function

Function dtoc(dt)
dtoc = strzero(Day(dt), 2) + "/" + strzero(Month(dt), 2) + "/" + strzero(Year(dt), 4)
End Function




Function senha(k)
chave = k
frm_pas.Show 1
If passw <> chave Then
   senha = False
Else
   senha = True
End If
End Function

Function strzero(vlr, tam, Optional char)
On Error Resume Next
Dim a
a = a + char
If Err <> 0 Then
   char = "0"
   Err = 0
End If
strzero = String$(tam - (Len(Trim$(Str$(vlr)))), char) & Trim$(Str$(vlr))
End Function

Function cria_indexes(dbs, tbl, indices)
On Error Resume Next
Dim db As Database
Dim tb As Recordset
Dim idx As New Index
Dim fld As Field
Dim i As Integer
Dim z As Integer
Dim nomedoindice As String
Dim chaveprimaria, requerido, unico As Boolean
Err = 0
Set db = OpenDatabase(dbs)
If Err <> 0 Then
   cria_indexes = False
   Exit Function
End If
Err = 0
For i = 0 To db.TableDefs(tbl).Indexes.Count - 1
   If Err <> 0 Then
      Exit For
   End If
   For z = 0 To UBound(indices(i, 1))
      indname = indices(z, 0)
      If UCase(db.TableDefs(tbl).Indexes(i).Name) = UCase(indname) Then
         cria_indexes = True
         Exit Function
      End If
   Next
Next
Err = 0
For i = 0 To UBound(indices)
   indname = indices(i, 0)
   chaveprimaria = indices(i, 2)
   requerido = indices(i, 3)
   unico = indices(i, 4)
   Set idx = db.TableDefs(tbl).CreateIndex(indname)
   For z = 0 To UBound(indices(i, 1))
      Set fld = idx.CreateField(indices(i, 1)(z))
      idx.Fields.Append fld
   Next
   idx.Name = indname
   idx.Primary = chaveprimaria
   idx.Required = requerido
   idx.Unique = unico
   db.TableDefs(tbl).Indexes.Append idx
   If Err <> 0 Then
      cria_indexes = False
   End If
Next i
db.Close
cria_indexes = True
End Function

Sub aadd(ByRef mt, el)
On Error Resume Next
Dim tam As Integer
Dim i As Integer
If IsArray(el) Then
   ReDim mtz(UBound(mt, 1), UBound(mt, 2))
   Err = 0
   tam = UBound(mt, 1)
   tam2 = UBound(mt, 2)
   If Err <> 0 Then
      ReDim mt(0, UBound(el))
      For i = 0 To UBound(el)
         mt(UBound(mt), i) = el(i)
      Next i
      Exit Sub
   End If
   For i = 0 To UBound(mt, 1)
      For j = 0 To UBound(mt, 2)
         mtz(i, j) = mt(i, j)
      Next j
   Next i
   tam = tam + 1
   ReDim mt(tam, tam2)
   For i = 0 To UBound(mtz, 1)
      For j = 0 To UBound(mtz, 2)
         mt(i, j) = mtz(i, j)
      Next j
   Next i
   For i = 0 To UBound(el)
      mt(UBound(mt), i) = el(i)
   Next i
Else
   ReDim Preserve mt(UBound(mt) + 1)
   If Err <> 0 Then
      mt = Array(el)
   Else
      mt(UBound(mt)) = el
   End If
End If
End Sub






Function file(Caminho)
If Dir(Caminho) = "" Then
   file = False
Else
   file = True
End If
End Function






Function VtoP(vlr)
Dim Word As String
Dim i As Integer
For i = 1 To Len(vlr)
   If Mid(vlr, i, 1) = "," Then
      Word = Word + "."
   Else
      Word = Word + Mid(vlr, i, 1)
   End If
Next
VtoP = Val(Word)
End Function


Function VtoPC(vlr)
Dim Word As String
Dim i As Integer
For i = 1 To Len(vlr)
   If Mid(vlr, i, 1) = "," Then
      Word = Word + "."
   Else
      Word = Word + Mid(vlr, i, 1)
   End If
Next
VtoPC = Word

End Function

Function xtoc(c)
Dim i As Integer
Dim char As String
Dim Retorno As String
Retorno = ""
For i = 1 To Len(c)
   char = Asc(Mid(c, i, 1)) - i * 10
   Do While char < 1
      char = char + 255
   Loop
   Retorno = Retorno + Chr$(char)
Next i
xtoc = Retorno
End Function

Function xtoc_old(c)
Dim i As Integer
Dim char As String
Dim Retorno As String
Retorno = ""
For i = 1 To Len(c)
   char = Asc(Mid(c, i, 1)) - i
   Retorno = Retorno + Chr$(char)
Next i
xtoc_old = Retorno
End Function


