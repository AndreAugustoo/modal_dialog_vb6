VERSION 5.00
Begin VB.Form ModalDialog 
   BorderStyle     =   0  'None
   Caption         =   "ModalDialog"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.RetroButton btn_sim 
      Height          =   480
      Left            =   2280
      TabIndex        =   0
      Top             =   3720
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Alignment       =   2
      Name            =   "btn_sim"
      Object.Visible         =   -1  'True
      Object.Tag             =   ""
      Tag2            =   ""
      Object.Left            =   2280
      Object.Top             =   3720
      Object.Width           =   1500
      Object.Height          =   480
      BackColor       =   15025743
      Caption         =   "Sim"
      DisabledPicturePath=   ""
      DownPicturePath =   ""
      DragIconPath    =   ""
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontItalic      =   0   'False
      FontName        =   "Segoe UI Semibold"
      FontSize        =   9,75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      MaskColor       =   12632256
      MouseIconPath   =   ""
      PicturePath     =   ""
      PictureAlignment=   0
      Object.ToolTipText     =   ""
      UseMaskColor    =   0   'False
   End
   Begin Project1.RetroButton btn_ok 
      Height          =   480
      Left            =   600
      TabIndex        =   1
      Top             =   3720
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Alignment       =   2
      Name            =   "btn_ok"
      Object.Visible         =   -1  'True
      Object.Tag             =   ""
      Tag2            =   ""
      Object.Left            =   600
      Object.Top             =   3720
      Object.Width           =   1500
      Object.Height          =   480
      BackColor       =   15025743
      Caption         =   "Ok"
      DisabledPicturePath=   ""
      DownPicturePath =   ""
      DragIconPath    =   ""
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontItalic      =   0   'False
      FontName        =   "Segoe UI Semibold"
      FontSize        =   9,75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      MaskColor       =   12632256
      MouseIconPath   =   ""
      PicturePath     =   ""
      PictureAlignment=   0
      Object.ToolTipText     =   ""
      UseMaskColor    =   0   'False
   End
   Begin Project1.RetroButton btn_nao 
      Height          =   480
      Left            =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Alignment       =   2
      Name            =   "btn_nao"
      Object.Visible         =   -1  'True
      Object.Tag             =   ""
      Tag2            =   ""
      Object.Left            =   1440
      Object.Top             =   3120
      Object.Width           =   1500
      Object.Height          =   480
      BackColor       =   12632256
      Caption         =   "Não"
      DisabledPicturePath=   ""
      DownPicturePath =   ""
      DragIconPath    =   ""
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontItalic      =   0   'False
      FontName        =   "Segoe UI Semibold"
      FontSize        =   9,75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   4210752
      MaskColor       =   12632256
      MouseIconPath   =   ""
      PicturePath     =   ""
      PictureAlignment=   0
      Object.ToolTipText     =   ""
      UseMaskColor    =   0   'False
   End
   Begin Project1.RetroButton btn_cancelar 
      Height          =   480
      Left            =   3960
      TabIndex        =   3
      Top             =   3720
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Alignment       =   2
      Name            =   "btn_cancelar"
      Object.Visible         =   -1  'True
      Object.Tag             =   ""
      Tag2            =   ""
      Object.Left            =   3960
      Object.Top             =   3720
      Object.Width           =   1500
      Object.Height          =   480
      BackColor       =   12632256
      Caption         =   "Cancelar"
      DisabledPicturePath=   ""
      DownPicturePath =   ""
      DragIconPath    =   ""
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontItalic      =   0   'False
      FontName        =   "Segoe UI Semibold"
      FontSize        =   9,75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   4210752
      MaskColor       =   12632256
      MouseIconPath   =   ""
      PicturePath     =   ""
      PictureAlignment=   0
      Object.ToolTipText     =   ""
      UseMaskColor    =   0   'False
   End
   Begin Project1.RetroButton btn_ignorar 
      Height          =   480
      Left            =   5640
      TabIndex        =   4
      Top             =   3720
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Alignment       =   2
      Name            =   "btn_ignorar"
      Object.Visible         =   -1  'True
      Object.Tag             =   ""
      Tag2            =   ""
      Object.Left            =   5640
      Object.Top             =   3720
      Object.Width           =   1500
      Object.Height          =   480
      BackColor       =   12632256
      Caption         =   "Ignorar"
      DisabledPicturePath=   ""
      DownPicturePath =   ""
      DragIconPath    =   ""
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontItalic      =   0   'False
      FontName        =   "Segoe UI Semibold"
      FontSize        =   9,75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   4210752
      MaskColor       =   12632256
      MouseIconPath   =   ""
      PicturePath     =   ""
      PictureAlignment=   0
      Object.ToolTipText     =   ""
      UseMaskColor    =   0   'False
   End
   Begin Project1.RetroButton btn_retornar 
      Height          =   480
      Left            =   3120
      TabIndex        =   5
      Top             =   3120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      Alignment       =   2
      Name            =   "btn_retornar"
      Object.Visible         =   -1  'True
      Object.Tag             =   ""
      Tag2            =   ""
      Object.Left            =   3120
      Object.Top             =   3120
      Object.Width           =   1500
      Object.Height          =   480
      BackColor       =   12632256
      Caption         =   "Retornar"
      DisabledPicturePath=   ""
      DownPicturePath =   ""
      DragIconPath    =   ""
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontItalic      =   0   'False
      FontName        =   "Segoe UI Semibold"
      FontSize        =   9,75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   4210752
      MaskColor       =   12632256
      MouseIconPath   =   ""
      PicturePath     =   ""
      PictureAlignment=   0
      Object.ToolTipText     =   ""
      UseMaskColor    =   0   'False
   End
   Begin Project1.RetroButton btn_abortar 
      Height          =   480
      Left            =   4800
      TabIndex        =   6
      Top             =   3120
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   847
      Alignment       =   2
      Name            =   "btn_abortar"
      Object.Visible         =   -1  'True
      Object.Tag             =   ""
      Tag2            =   ""
      Object.Left            =   4800
      Object.Top             =   3120
      Object.Width           =   1380
      Object.Height          =   480
      BackColor       =   12632256
      Caption         =   "Abortar"
      DisabledPicturePath=   ""
      DownPicturePath =   ""
      DragIconPath    =   ""
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontItalic      =   0   'False
      FontName        =   "Segoe UI Semibold"
      FontSize        =   9,75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   4210752
      MaskColor       =   12632256
      MouseIconPath   =   ""
      PicturePath     =   ""
      PictureAlignment=   0
      Object.ToolTipText     =   ""
      UseMaskColor    =   0   'False
   End
   Begin VB.Image imgIcone 
      Height          =   735
      Left            =   3600
      Top             =   360
      Width           =   900
   End
   Begin VB.Label lbl_mensagem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1335
      Left            =   720
      TabIndex        =   8
      Top             =   2160
      Width           =   6735
   End
   Begin VB.Label lbl_titulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   1680
      Width           =   6615
   End
   Begin VB.Line borda 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   7200
      X2              =   7200
      Y1              =   360
      Y2              =   840
   End
End
Attribute VB_Name = "ModalDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public resposta As Integer

Private Sub AjustarTela()
   
   Call AjustarBorda(Me)
   
   With lbl_titulo
      .Left = (Me.ScaleWidth - .Width) / 2
   End With
   
   With lbl_mensagem
      .Left = (Me.ScaleWidth - .Width) / 2
   End With
   
   With imgIcone
      .Left = (Me.ScaleWidth - .Width) / 2
   End With
   
   Call AjustarBotoes
   
End Sub

Public Sub AjustarBorda(P_Formulario As Form)
    On Error Resume Next
    
    Const Margem As Integer = 8
    
    Load P_Formulario.borda(0)
    Load P_Formulario.borda(1)
    Load P_Formulario.borda(2)
    Load P_Formulario.borda(3)
    
    With P_Formulario.borda(0)
        .X1 = 0
        .Y1 = 0
        .X2 = P_Formulario.ScaleWidth - Margem
        .Y2 = 0
        .Visible = True
        .ZOrder 0
    End With
    
    With P_Formulario.borda(1)
        .X1 = 0
        .Y1 = 0
        .X2 = 0
        .Y2 = P_Formulario.ScaleHeight - Margem
        .Visible = True
        .ZOrder 0
    End With
    
    With P_Formulario.borda(2)
        .X1 = P_Formulario.ScaleWidth - Margem
        .Y1 = 0
        .X2 = P_Formulario.ScaleWidth - Margem
        .Y2 = P_Formulario.ScaleHeight - Margem
        .Visible = True
        .ZOrder 0
    End With
    
    With P_Formulario.borda(3)
        .X1 = 0
        .Y1 = P_Formulario.ScaleHeight - Margem
        .X2 = P_Formulario.ScaleWidth - Margem
        .Y2 = P_Formulario.ScaleHeight - Margem
        .Visible = True
        .ZOrder 0
    End With
    
End Sub

Private Sub AjustarBotoes()

   On Error GoTo TrataErro
   
   Const MARGEM_LATERAL As Integer = 720
   Const MARGEM_ENTRE_BOTOES As Integer = 240

   Dim I As Integer
   Dim visiveis As Integer
   Dim posicaoX As Integer
   Dim Botoes() As RetroButton
   Dim BotaoTemp As RetroButton
   
   Dim ListaBotoes As Collection
   Set ListaBotoes = New Collection
   
   If btn_ok.Visible Then ListaBotoes.Add btn_ok
   If btn_cancelar.Visible Then Lista.Add btn_cancelar
   If btn_sim.Visible Then ListaBotoes.Add btn_sim
   If btn_nao.Visible Then ListaBotoes.Add btn_nao
   If btn_retornar.Visible Then ListaBotoes.Add btn_retornar
   If btn_ignorar.Visible Then ListaBotoes.Add btn_ignorar
   If btn_abortar.Visible Then ListaBotoes.Add btn_abortar
   
   visiveis = ListaBotoes.count
   If visiveis = 0 Then Exit Sub
   
   Dim larguraBotao As Integer
   Dim sobra As Integer
   Dim larguraDisponivel As Integer

   larguraDisponivel = Me.Width - (MARGEM_LATERAL * 2) - (MARGEM_ENTRE_BOTOES * (visiveis - 1))
   
   larguraBotao = larguraDisponivel / visiveis
   sobra = larguraDisponivel Mod visiveis
   
   posicaoX = MARGEM_LATERAL
   
   For I = 1 To visiveis
      Set BotaoTemp = ListaBotoes(I)
      
      With BotaoTemp
         .Top = btn_ok.Top
         .Left = posicaoX
         .Width = larguraBotao
      
      
      If I = visiveis Then .Width = .Width + sobra
      
         .Height = btn_ok.Height
         posicaoX = posicaoX + .Width + MARGEM_ENTRE_BOTOES
      End With
      
   Next I
   
   Exit Sub
   
TrataErro:
   MsgBox "Erro ao ajustar os botões: " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub btn_abortar_Click()
   resposta = vbAbort
   Unload Me
End Sub

Private Sub btn_cancelar_Click()
   resposta = vbCancel
   Unload Me
End Sub

Private Sub btn_ignorar_Click()
   resposta = vbIgnore
   Unload Me
End Sub

Private Sub btn_nao_Click()
   resposta = vbNo
   Unload Me
End Sub

Private Sub btn_ok_Click()
   resposta = vbOK
   Unload Me
End Sub

Private Sub btn_retornar_Click()
   resposta = vbRetry
   Unload Me
End Sub

Private Sub btn_sim_Click()
   resposta = vbYes
   Unload Me
End Sub

Private Sub Form_Activate()

   Call AjustarTela

End Sub

