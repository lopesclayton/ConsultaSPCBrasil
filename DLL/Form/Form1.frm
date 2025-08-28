VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   4335
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "Form1.frx":0000
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumo"
      Height          =   4935
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
      Begin VB.Label lblProtestos 
         Caption         =   "lblAlertas"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   1800
         Width           =   935
      End
      Begin VB.Label Label7 
         Caption         =   "Protestos"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblTotal_Sustados 
         Caption         =   "lblAlertas"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   1536
         Width           =   935
      End
      Begin VB.Label Label5 
         Caption         =   "Sustados"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1536
         Width           =   2175
      End
      Begin VB.Label lblTotal_Alertas 
         Caption         =   "lblAlertas"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   1272
         Width           =   935
      End
      Begin VB.Label Label4 
         Caption         =   "Alertas"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1272
         Width           =   2175
      End
      Begin VB.Label lblTotal_Acoes 
         Caption         =   "Label2"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   1008
         Width           =   935
      End
      Begin VB.Label Label3 
         Caption         =   "Açoes"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1008
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Cheques sem Fundo"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   744
         Width           =   2175
      End
      Begin VB.Label lblTotal_CCF 
         Caption         =   "Label2"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   744
         Width           =   935
      End
      Begin VB.Label lblTotal_SPC 
         Caption         =   "Label2"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   480
         Width           =   935
      End
      Begin VB.Label Label1 
         Caption         =   "Registros SPC"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7920
      TabIndex        =   0
      Top             =   5640
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim SCPC As New Consulta_SPCBrasil.SPC_Brasil_Consulta

    Command1.Enabled = False
    DoEvents

    With SCPC

        .Codigo_Transacao = "SPC100"
        .Versao = "014"
        .Meio_de_Acesso = "20"
        .Codigo_Estacao = ""                               'INFORMAR VAZIO
        .Codigo_Operador = "390044"
        .Senha_Operador = "19072012"
        .Tipo_Documento = "1"
        .CPF_CNPJ = "12312312387"
        '.CPF_CNPJ = "12345678909"
        .RG = ""
        .Codigo_Produto = "04"                             'InputBox("Codigo do Servico")
        .Tipo_Resposta = "A"
        .Informacaoes_Obito = "N"
        .Informacoes_RG = "N"

        .CNPJ_Empresa = "11111111111"
        .IDLiberacao = "MY2012A10SO000197246E"


        If .Consulta("treina.spc.org.br", "3348") = False Then
            MsgBox .Msg_Erro, vbCritical, "Erro"
        Else
            With .Header_Resposta
                Text1.Text = .CPF_CNPJ
                Text2.Text = .Nome
                lblTotal_SPC.Caption = .Total_SPC
                lblTotal_CCF.Caption = .Total_Cheque_CCF
                lblTotal_Acoes.Caption = .Total_Acoes
                lblTotal_Alertas.Caption = .Total_Alertas
                lblTotal_Sustados.Caption = .Total_Cheques_Sustados_Linha21
                lblProtestos.Caption = .Total_Protesto
                MsgBox .Indicativo_Consulta_Parcial
            End With
            MsgBox "consultou"
        End If

    End With

    Set SCPC = Nothing

    Command1.Enabled = True
    DoEvents



End Sub

