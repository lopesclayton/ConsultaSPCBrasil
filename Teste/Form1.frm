VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   15855
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8LCtl.VSFlexGrid Grid 
      Height          =   7095
      Left            =   6240
      TabIndex        =   8
      Top             =   240
      Width           =   9255
      _cx             =   16325
      _cy             =   12515
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ListaProdutos"
      Height          =   735
      Left            =   3000
      TabIndex        =   7
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   3360
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Text            =   "03274289108"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "675"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consulta"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Numero CPF/CNPJ"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo operacao"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Myf        As New cFuncoes
Dim Software   As New Software

Private Sub Command1_Click()
    Dim con As New SPC_Brasil_Consulta
    'Usuario = "395793"
    'Senha = "12012016"
    'ACodigoProduto = "116"
    'ATipoConsumidor = "F"
    'ACPF = "00000000191"
    'ACPF = "00829161600"


    Dim insumos As New insumos
    ' With insumos.Add
    '   .Codigo = "5193"
    ' .Nome = "xxx"
    ' End With

    con.UsarArquivo = "C:\TempSia\respostaspc.xml"
    'con.Gravar_Resposta_Arquivo = "C:\TempSia\respostaspc.xml"
    con.ConsultarViaWebService "132044823", "Magoga$03", 675, Fisica, "03274289108", insumos, Producao

    'Call con.ListarProdutosDisponiveis("132044823", "Magoga$03", Producao)



    If con.Erro = True Then
        MsgBox con.Msg_Erro, vbCritical, "ERRO"
    Else
        MsgBox "terminou"

        MsgBox con.Produtos.Count
        'MsgBox con.Produtos.Item(1).Nome
        MsgBox con.StrResposta
    End If
End Sub

Private Sub Command2_Click()
    Dim Consulta As New SPC_Brasil_Consulta

    Consulta.UsarArquivo = "D:\Desenvolvimento\Projetos VB 6\SPC BRASIL\SPCXML.xml"
    'Consulta.ConsultarViaWebService "", "", "", Juridica, "", 2, Homologacao

    Grid.Rows = 0
    Grid.Cols = 1
    Grid.FixedCols = 0
    Grid.ColWidth(0) = Grid.Width * 1


    ''' RESUMO DAS INFORMACOES
    With Consulta
        With .Header_Resposta
            Grid.AddItem "=====================================      RESUMO      ==========================================="
            Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = &H80FF&
            Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1

            If .Total_Acoes <> "0000" Then Grid.AddItem Myf.PadL("AÇÕES ", 50) & .Total_Acoes
            If .Total_Protesto <> "0000" Then Grid.AddItem Myf.PadL("PROTESTOS ", 50) & .Total_Protesto
            If .Total_Alertas <> "0000" Then Grid.AddItem Myf.PadL("ALERTAS", 50) & .Total_Alertas
            If .Total_InfoPoderJudiciario <> "0000" Then Grid.AddItem Myf.PadL("INFO PODE JUDICIÁRIO ", 50) & .Total_InfoPoderJudiciario
            If .Total_SPC <> "0000" Then Grid.AddItem Myf.PadL("SPC ", 50) & .Total_SPC
            If .Total_Cheque_CCF <> "0000" Then Grid.AddItem Myf.PadL("CHEQUES SEM FUNDO ", 50) & .Total_Cheque_CCF
            If .Total_Cheques_Sustados_Linha21 <> "0000" Then Grid.AddItem Myf.PadL("CHEQUES SUSTADOS ", 50) & .Total_Cheques_Sustados_Linha21
            If .Total_Cheques_Sustados_Online <> "0000" Then Grid.AddItem Myf.PadL("CHEQUES SUSTADOS ONLINE ", 50) & .Total_Cheques_Sustados_Online
            If .Total_RegistroConsulta <> "0000" Then Grid.AddItem Myf.PadL("REGISTRO DE CONSULTAS ", 50) & .Total_RegistroConsulta
            If .Total_Consultas_Realizadas <> "0000" Then Grid.AddItem Myf.PadL("CONSULTAS ", 50) & .Total_Consultas_Realizadas

            If .Total_HistoricoPagamento <> "0000" Then Grid.AddItem Myf.PadL("HISTÓRICO DE PAGAMENTO ", 50) & .Total_HistoricoPagamento
            If .Total_Cheque_ContraOrdenados <> "0000" Then Grid.AddItem Myf.PadL("CHEQUE CONTRA ORDEM ", 50) & .Total_Cheque_ContraOrdenados

            If .Total_ControleSocietario <> "0000" Then Grid.AddItem Myf.PadL("CONTROLE SOCIETÁRIO ", 50) & .Total_ControleSocietario

            If .Total_PendenciaSerasa <> "0000" Then Grid.AddItem Myf.PadL("PENDÊNCIA SERASA ", 50) & .Total_PendenciaSerasa

            If .Total_ParticipacaoFalencia <> "0000" Then Grid.AddItem Myf.PadL("PARTICIPAÇÃO EM FALÊNCIA ", 50) & .Total_ParticipacaoFalencia


            Grid.AddItem "=================================================================================================="
            Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
            Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
        End With
        '''' FIM


        With .REG_095
            If Consulta.Header_Resposta.CPF_CNPJ <> "" And Consulta.Header_Resposta.Tipo_Documento = "CPF" Then
                Grid.AddItem "===================================== DADOS CADASTRAIS ==========================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem "NOME: " & Myf.PadL(.Nome, 60) & " NASCIMENTO: " & .Data_Nasc_Fundacao
                Grid.AddItem "CPF.: " & Format(Consulta.Header_Resposta.CPF_CNPJ, "@@@.@@@.@@@-@@") & "  SITUAÇÃO: " & SpcBrasil_SituacaoCPF(.Situacao_CPF_CNPJ) & "        RG: " & .RG & " UF: " & .UF_IE
                Grid.AddItem "MAE.: " & .Nome_Mae
                Grid.AddItem "TITULO ELEITORAL: " & .Titulo_Eleitor
                Grid.AddItem "ENDEREÇO: " & .Endereco & "         BAIRRO: " & .Bairro & "      CEP: " & .Cep
                Grid.AddItem "CIDADE  : " & .Cidade & "  UF: " & .UF
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            ElseIf Consulta.Header_Resposta.CPF_CNPJ <> "" And Consulta.Header_Resposta.Tipo_Documento = "CNPJ" Then
                Grid.AddItem "======================================== DADOS CADASTRAIS ==========================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem "NOME: " & .Nome & " FUNDAÇÃO: " & .Data_Nasc_Fundacao
                Grid.AddItem "CNPJ: " & Format(Consulta.Header_Resposta.CPF_CNPJ, "@@.@@@.@@@/@@@@-@@") & "   SITUAÇÃO: " & SpcBrasil_SituacaoCPF(.Situacao_CPF_CNPJ) & "       IE: " & .IE
                Grid.AddItem "FANTASIA: " & .Nome_Fantasia
                Grid.AddItem "ENDEREÇO: " & .Endereco & "         BAIRRO: " & .Bairro & "      CEP: " & .Cep
                Grid.AddItem "CIDADE  : " & .Cidade & "  UF: " & .UF
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0

            End If
        End With

        With .Reg_006
            If .Count > 0 Then
                Grid.AddItem "====================================== REGISTRO DE GRAFIAS ==========================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                For i = 1 To .Count
                    With .Item(i)
                        Grid.AddItem "NOME: " & .Razao_Social & " "
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With

        With .REG_CAPITALSOCIAL
            If Myf.MyIsNumeric(.ValorCapitalSocial) > 0 Then
                Grid.AddItem "====================== CAPITAL SOCIAL ATUALIZAÇÃO EM (" & .DataAtualizacao & ") ======================"
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem "CAPITAL SOCIAL: " & Myf.PadL(.ValorCapitalSocial, 30) & " REALIZADO: " & .ValorCapitalRealizado & "          "
                Grid.AddItem "AUTORIZADO    : " & Myf.PadL(.ValorCapitalAutorizado, 30) & " NACIONALIDADE: " & .NACIONALIDADE & "   "
                Grid.AddItem "ORIGEM        : " & Myf.PadL(.Origem, 30) & " NATUREZA: " & .Natureza & "             "

                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With


        With .REG_096
            If .Count > 0 Then
                Grid.AddItem "====================================== SÓCIOS ======================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                For i = 1 To .Count
                    With .Item(i)
                        Grid.AddItem "   CPF: " & .CPF_CNPJ & "                 NOME: " & .Nome_Socio_Proprietario & "    "
                        Grid.AddItem "   % PARTICIPAÇÃO: " & .Percentual_Participacao & "            DATA ENTRADA: " & .Data_Entrada & "     "

                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With



        With .REG_007
            If .Data_Obito <> "" Then
                Grid.AddItem "====================================== INFORMAÇÃO DE ÓBITO ==========================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem "DATA ÓBITO: " & .Data_Obito & "          CARTÓRIO: " & .Cartorio
                Grid.AddItem "CIDADE: " & .Cidade & "  UF: " & .UF
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With


        With .Reg_008
            Grid.AddItem "===================================== SCORE DE CRÉDITO ==========================================="
            Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
            If .Valor <> "" Then
                Grid.AddItem "HORIZONTE: " & .Horizonte & " Meses" & "    SCORE: " & .Valor & "   PROBABILIDADE: " & .Probabilidade
                Grid.AddItem "CLASSE DE RISCO: " & .Classe_Risco
                Grid.AddItem "MENSAGEM         : " & .Menssagem_Curta
                Grid.AddItem "MSG SCORE        : " & .Menssagem_Interpretativa_Score
                Grid.AddItem "MSG PROBABILIDADE: " & .Messagem_Interpretativa_Probabilidade
            End If

            If Consulta.REG_009.Menssagem_Score <> "" Then Grid.AddItem "*** " & Consulta.REG_009.Menssagem_Score

            Grid.AddItem "=================================================================================================="
            Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
            Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
        End With


        With .Reg_010
            If .Count > 0 Then
                Grid.AddItem "===================================== OCORRÊNCIAS DE DÉBITOS ==========================================="
                Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = 255
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem "  DATA    | VENCIMENTO|    CONTRATO     |T|    VALOR     |CIDADE                |INFORMANTE           "
                For i = 1 To .Count
                    With .Item(i)
                        TextoLinha = ""
                        TextoLinha = TextoLinha & Myf.PadL(.Data_Inclusao, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Data_Vencimento, 11) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Numero_Contrato, 17) & " "
                        'TextoLinha = TextoLinha & myf.PadL(.parcela, 3) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Comprador_Fiador, 1) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Valor_Debito, 13) & "  "
                        TextoLinha = TextoLinha & Myf.PadL(.Cidade, 22) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Nome_Associado, 22) & " "
                        Grid.AddItem TextoLinha
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With


        With .REG_020
            If .Count > 0 Then
                Grid.AddItem "===================================== OCORRÊNCIAS DE CHEQUES ==========================================="
                Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = 255
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem " MOT| EMISSAO|  INCLUSAO | BANCO| AGEN|    CONTA    |  Nº CH  | QT |   VALOR    | INFORMANTE"
                For i = 1 To .Count
                    With .Item(i)
                        TextoLinha = Myf.PadR(.Codigo_Alinea, 3) & "  "
                        TextoLinha = TextoLinha & Myf.PadR(.Data_Emissao, 9) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Data_Inclusao, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Codigo_Banco, 6) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Codigo_Agencia, 5) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Conta_Corrente, 13) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Numero_Cheque, 9) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Qtde_Ocorrencias, 3) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Valor_Documento, 13) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Entidade_Origem, 20) & " "
                        Grid.AddItem TextoLinha
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With


        With .REG_028
            If .Count > 0 Then
                Grid.AddItem "=====================================    CHEQUES SUSTADOS    ==========================================="
                Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = 255
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem " MOT| BANCO| AGEN|    CONTA    | ABERTURA |  Nº CH  | INFORMANTE"
                For i = 1 To .Count
                    With .Item(i)
                        TextoLinha = Myf.PadR(.Codigo_Alinea, 3) & "  "
                        TextoLinha = TextoLinha & Myf.PadR(.Codigo_Banco, 6) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Codigo_Agencia, 5) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Conta_Corrente, 13) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Data_Abertura_ContaCorrente, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Numero_Cheque, 9) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Origem, 20) & " "
                        Grid.AddItem TextoLinha
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With


        With .REG_092
            If .Count > 0 Then
                Grid.AddItem "================================    CHEQUES SUSTADOS CONTUMÁCIA    ==========================================="
                Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = 255
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem " MOT| EMISSAO|  INCLUSAO | BANCO| AGEN|    CONTA    |  Nº CH  |    VALOR    | INFORMANTE"
                For i = 1 To .Count
                    With .Item(i)
                        TextoLinha = Myf.PadR(.Codigo_Alinea, 3) & "  "
                        TextoLinha = TextoLinha & Myf.PadR(.Data_Ocorrencia, 9) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Data_Incluso, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Codigo_Banco, 6) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Codigo_Agencia, 5) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Conta_Corrente, 13) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Cheque_Inicial, 9) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Valor, 13) & "  "
                        TextoLinha = TextoLinha & Myf.PadL(.Entidade_Origem, 20) & " "
                        Grid.AddItem TextoLinha
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With



        With .REG_027
            If .Count > 0 Then
                Grid.AddItem "================================== INFORMAÇÃO ONLINE INSTITUIÇÃO FINANCEIRA ==========================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                For i = 1 To .Count
                    With .Item(i)
                        Grid.AddItem .Menssagem
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With

        With .REG_030
            If .Count > 0 Then
                Grid.AddItem "=====================================        PROTESTOS      ==========================================="
                Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = 255
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem "   DATA   | PROTESTO | CARTORIO  |     VALOR     | CIDADE             | INFORMANTE                   "
                For i = 1 To .Count
                    With .Item(i)
                        TextoLinha = Myf.PadR(.Data_Inclusao, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Data_Protesto, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Cartorio, 11) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Valor, 15) & "  "
                        TextoLinha = TextoLinha & Myf.PadL(.Cidade, 20) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Requerente, 30) & " "
                        Grid.AddItem TextoLinha
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With


        With .REG_040
            If .Count > 0 Then
                Grid.AddItem "=====================================        AÇÕES       ==========================================="
                Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = 255
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem " INCLUSÃO |   DATA   |TIPO AÇÃO           |VARA                |     VALOR     | CIDADE             | INFORMANTE                   "
                For i = 1 To .Count
                    With .Item(i)
                        TextoLinha = Myf.PadR(.Data_Inclusao, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Data_Acao, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Tipo_Acao, 20) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Vara, 20) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Valor, 15) & "  "
                        TextoLinha = TextoLinha & Myf.PadL(.Cidade, 20) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Entidade_Origem, 30) & " "
                        Grid.AddItem TextoLinha
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With


        With .REG_050
            If .Count > 0 Then
                Grid.AddItem "===================================== CRÉDITO CONCEDIDO  ==========================================="
                Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = vbBlue
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem " MÊS/ANO |   QTDE   |     VALOR     | NOME"
                For i = 1 To .Count
                    With .Item(i)
                        TextoLinha = Myf.PadC(.Mes_Ano, 9) & " "
                        TextoLinha = TextoLinha & Myf.PadC(.Qdte, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadR(.Valor, 15) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Nome_Consumidor, 20) & " "
                        Grid.AddItem TextoLinha
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With


        With .REG_060
            If .Count > 0 Then
                Grid.AddItem "===================================== OCORRÊNCIAS DE ALERTA ==========================================="
                Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = 255
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem "   DATA   | TIPO DOCUMENTO       | MOTIVO               | CIDADE               | INFORMANTE         "
                For i = 1 To .Count
                    With .Item(i)
                        TextoLinha = Myf.PadR(.Data_Inclusao, 10) & "  "
                        TextoLinha = TextoLinha & Myf.PadL(.Tipo_Documento, 22) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.motivo, 22) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Cidade, 22) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Entidade_Origem, 22) & " "
                        Grid.AddItem TextoLinha
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With


        With .REG_070
            If .Count > 0 Then
                Grid.AddItem "====================================== HISTÓRICO DE CONSULTAS ==========================================="
                Grid.Cell(flexcpForeColor, Grid.Rows - 1, 0) = vbBlue
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.AddItem "   DATA   | HORA | ASSOCIADO               | CIDADE                |ORIGEM                            "
                For i = 1 To .Count
                    With .Item(i)
                        TextoLinha = Myf.PadR(.Data_Consulta, 10) & " "
                        TextoLinha = TextoLinha & Myf.PadR(Format(.Hora_Consulta, "HH:NN"), 6) & "  "
                        TextoLinha = TextoLinha & Myf.PadL(.Nome_Associado, 23) & "  "
                        TextoLinha = TextoLinha & Myf.PadL(" " & .Cidade, 23) & " "
                        TextoLinha = TextoLinha & Myf.PadL(.Entidade_Origem, 30) & "  "
                        Grid.AddItem TextoLinha
                    End With
                Next i
                Grid.AddItem "=================================================================================================="
                Grid.Cell(flexcpFontBold, Grid.Rows - 1, 0) = 1
                Grid.Cell(flexcpBackColor, Grid.Rows - 1, 0) = &HE0E0E0
            End If
        End With




        With .Header_Resposta
            'Grid.AddItem MyF.PadR("Mycommerce " & App.Major & "." & App.Minor & "." & App.Revision & Space(52) & "-------------------------", 97)
            Grid.AddItem Myf.PadR(Software.VersaoSoftware & Space(51) & "-------------------------", 97)
            Grid.AddItem Myf.PadR("DATA: " & Format(Date, "DD/MM/YYYY"), 97)
            Grid.AddItem Myf.PadR(" Nº CONSULTA: " & .Numero_Protocolo, 97)
            Grid.AddItem Myf.PadR("-------------------------", 97)
        End With



    End With
End Sub



Private Function SpcBrasil_SituacaoCPF(ByVal Situacao As String) As String
    If Situacao = "C" Then
        SpcBrasil_SituacaoCPF = "CANCELADO"
    ElseIf Situacao = "S" Then
        SpcBrasil_SituacaoCPF = "SUSPENSO"
    Else
        SpcBrasil_SituacaoCPF = "ATIVO"
    End If
End Function

Private Sub Command3_Click()
    Dim con    As New SPC_Brasil_Consulta
    Dim i      As Integer

    If con.ListarProdutosDisponiveis("2166977", "30012019", Producao) = True Then
        For i = 0 To con.Produtos.Count
        Next i
    End If

End Sub

Private Sub Command4_Click()
    Dim con    As New SPC_Brasil_Consulta
    'Usuario = "395793"
    'Senha = "12012016"
    'ACodigoProduto = "116"
    'ATipoConsumidor = "F"
    'ACPF = "00000000191"
    'ACPF = "00829161600"


    con.Gravar_Resposta_Arquivo = "C:\tempsia\ListarProdutosDisponiveis.xml"
    Call con.ListarProdutosDisponiveis("132044823", "Magoga$03", Producao)



    If con.Erro = True Then
        MsgBox con.Msg_Erro, vbCritical, "ERRO"
    Else

        MsgBox "terminou"

        With Grid
            .Cols = 1
            .Rows = 1
            .FormatString = "|Codigo|Nome|Insumo|Nome Insumo"
            .ColWidth(1) = .Width * 0.09
            .ColWidth(2) = .Width * 0.25
            .ColWidth(3) = .Width * 0.09
            '.ColWidth(4) = .Width * 0.25
            .ExtendLastCol = True
        End With

        For i = 1 To con.Produtos.Count
            Grid.AddItem ""
            Grid.TextMatrix(Grid.Rows - 1, 1) = con.Produtos(i).Codigo
            Grid.TextMatrix(Grid.Rows - 1, 2) = con.Produtos(i).Nome
            Grid.TextMatrix(Grid.Rows - 1, 3) = con.Produtos(i).Tipo
            For x = 1 To con.Produtos(i).insumos.Count
                Grid.AddItem ""
                Grid.TextMatrix(Grid.Rows - 1, 3) = con.Produtos(i).insumos(x).Codigo
                Grid.TextMatrix(Grid.Rows - 1, 4) = con.Produtos(i).insumos(x).Nome
            Next x
        Next i
        MsgBox con.StrResposta
    End If

End Sub
