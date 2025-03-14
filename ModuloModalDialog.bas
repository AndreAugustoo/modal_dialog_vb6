Attribute VB_Name = "ModuloModalDialog"
Public Function RetroMsgBox(ByVal Mensagem As String, ByVal Tipo As Integer, Optional ByVal Titulo As String = "Titulo") As Integer

    ModalDialog.lbl_mensagem.Caption = Mensagem
    ModalDialog.lbl_titulo.Caption = Titulo

    ModalDialog.btn_ok.Visible = False
    ModalDialog.btn_cancelar.Visible = False
    ModalDialog.btn_sim.Visible = False
    ModalDialog.btn_nao.Visible = False
    ModalDialog.btn_retornar.Visible = False
    ModalDialog.btn_ignorar.Visible = False
    ModalDialog.btn_abortar.Visible = False

    Dim botaoTipo As Integer
    botaoTipo = Tipo And &H7

    If botaoTipo = vbOKOnly Then
        ModalDialog.btn_ok.Visible = True
    ElseIf botaoTipo = vbOKCancel Then
        ModalDialog.btn_ok.Visible = True
        ModalDialog.btn_cancelar.Visible = True
    ElseIf botaoTipo = vbAbortRetryIgnore Then
        ModalDialog.btn_abortar.Visible = True
        ModalDialog.btn_retornar.Visible = True
        ModalDialog.btn_ignorar.Visible = True
    ElseIf botaoTipo = vbYesNoCancel Then
        ModalDialog.btn_sim.Visible = True
        ModalDialog.btn_nao.Visible = True
        ModalDialog.btn_cancelar.Visible = True
    ElseIf botaoTipo = vbYesNo Then
        ModalDialog.btn_sim.Visible = True
        ModalDialog.btn_nao.Visible = True
    ElseIf botaoTipo = vbRetryCancel Then
        ModalDialog.btn_retornar.Visible = True
        ModalDialog.btn_cancelar.Visible = True
    End If

    DefinirIcone Tipo

    ModalDialog.Show vbModal

    RetroMsgBox = ModalDialog.resposta
End Function


Private Sub DefinirIcone(ByVal Tipo As Integer)

    Select Case Tipo And &H70
        Case vbCritical
            LoadPNG ModalDialog.imgIcone, "C:\Projects\VB6\ModalDialog\img\erro.png"
        Case vbQuestion
            LoadPNG ModalDialog.imgIcone, "C:\Projects\VB6\ModalDialog\img\interrogacao.png"
        Case vbExclamation
            LoadPNG ModalDialog.imgIcone, "C:\Projects\VB6\ModalDialog\img\alerta.png"
        Case vbInformation
            LoadPNG ModalDialog.imgIcone, "C:\Projects\VB6\ModalDialog\img\concluido.png"
    End Select
End Sub

Public Sub LoadPNG(ByRef P_ComponenteImagem As Image, P_CaminhoImagem As String)
    Dim StdPictureExInstance As New StdPictureEx
    
    Set P_ComponenteImagem.Picture = StdPictureExInstance.LoadPicture(P_CaminhoImagem)
End Sub


