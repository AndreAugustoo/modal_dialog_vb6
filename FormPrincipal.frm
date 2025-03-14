VERSION 5.00
Begin VB.Form FormPrincipal 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
   Const TextoGenerico = "Lorem ipsum, dolor sit amet consectetur adipisicing elit. Eius aliquam laudantium explicabo pariatur iste dolorem animi vitae error totam. At sapiente aliquam accusamus facere veritatis."
   
   MsgBox TextoGenerico, vbCritical, "Erro"
   RetroMsgBox TextoGenerico, vbCritical, "Erro"

   MsgBox TextoGenerico, vbInformation, "Sucesso"
    RetroMsgBox TextoGenerico, vbInformation, "Sucesso"
    
    MsgBox TextoGenerico, vbExclamation, "Alerta"
    RetroMsgBox TextoGenerico, vbExclamation, "Alerta"
    
    MsgBox TextoGenerico, vbQuestion, "Pergunta"
    RetroMsgBox TextoGenerico, vbQuestion, "Pergunta"
    
    MsgBox TextoGenerico, vbYesNo + vbQuestion, "Sim ou não"
    RetroMsgBox TextoGenerico, vbYesNo + vbQuestion, "Sim ou não"
    
    MsgBox TextoGenerico, vbCritical + vbYesNo + vbDefaultButton2, "Sim ou não"
    RetroMsgBox TextoGenerico, vbCritical + vbYesNo + vbDefaultButton2, "Sim ou não"

End Sub
