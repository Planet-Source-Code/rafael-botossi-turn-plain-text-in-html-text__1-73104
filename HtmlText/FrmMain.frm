VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "FrmHtmlText"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
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
   ScaleHeight     =   4170
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTurn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Turn in Html Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TxtOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2160
      Width           =   7935
   End
   Begin VB.TextBox TxtInput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label LblInput 
      BackStyle       =   0  'Transparent
      Caption         =   "Input:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Make the text with special chars
'To get more, go to http://www.ime.usp.br/~glauber/html/acentos.htm
Private Sub CmdTurn_Click()
    CmdTurn.Caption = "Wait..."
    Dim TextoNovo As String
    For i = 1 To Len(TxtInput.Text)
        If Mid(TxtInput.Text, i, 1) = "á" Then
            TextoNovo = TextoNovo & "&aacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "é" Then
            TextoNovo = TextoNovo & "&eacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "í" Then
            TextoNovo = TextoNovo & "&iacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ó" Then
            TextoNovo = TextoNovo & "&oacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ú" Then
            TextoNovo = TextoNovo & "&uacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Á" Then
            TextoNovo = TextoNovo & "&Aacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "É" Then
            TextoNovo = TextoNovo & "&Eacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Í" Then
            TextoNovo = TextoNovo & "&Iacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ó" Then
            TextoNovo = TextoNovo & "&Oacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ú" Then
            TextoNovo = TextoNovo & "&Uacute;"
        ElseIf Mid(TxtInput.Text, i, 1) = "â" Then
            TextoNovo = TextoNovo & "&acirc;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ê" Then
            TextoNovo = TextoNovo & "&ecirc;"
        ElseIf Mid(TxtInput.Text, i, 1) = "î" Then
            TextoNovo = TextoNovo & "&icirc;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ô" Then
            TextoNovo = TextoNovo & "&ocirc;"
        ElseIf Mid(TxtInput.Text, i, 1) = "û" Then
            TextoNovo = TextoNovo & "&ucirc;"
            ElseIf Mid(TxtInput.Text, i, 1) = "Â" Then
            TextoNovo = TextoNovo & "&Acirc;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ê" Then
            TextoNovo = TextoNovo & "&Ecirc;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Î" Then
            TextoNovo = TextoNovo & "&Icirc;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ô" Then
            TextoNovo = TextoNovo & "&Ocirc;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Û" Then
            TextoNovo = TextoNovo & "&Ucirc;"
        ElseIf Mid(TxtInput.Text, i, 1) = "à" Then
            TextoNovo = TextoNovo & "&agrave;"
        ElseIf Mid(TxtInput.Text, i, 1) = "è" Then
            TextoNovo = TextoNovo & "&egrave;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ì" Then
            TextoNovo = TextoNovo & "&igrave;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ò" Then
            TextoNovo = TextoNovo & "&ograve;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ù" Then
            TextoNovo = TextoNovo & "&ugrave;"
        ElseIf Mid(TxtInput.Text, i, 1) = "À" Then
            TextoNovo = TextoNovo & "&Agrave;"
        ElseIf Mid(TxtInput.Text, i, 1) = "È" Then
            TextoNovo = TextoNovo & "&Egrave;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ì" Then
            TextoNovo = TextoNovo & "&Igrave;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ò" Then
            TextoNovo = TextoNovo & "&Ograve;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ù" Then
            TextoNovo = TextoNovo & "&Ugrave;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ã" Then
            TextoNovo = TextoNovo & "&atilde;"
        ElseIf Mid(TxtInput.Text, i, 1) = "õ" Then
            TextoNovo = TextoNovo & "&otilde;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ã" Then
            TextoNovo = TextoNovo & "&Atilde;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Õ" Then
            TextoNovo = TextoNovo & "&Otilde;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ç" Then
            TextoNovo = TextoNovo & "&ccedil;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ç" Then
            TextoNovo = TextoNovo & "&Ccedil;"
        ElseIf Mid(TxtInput.Text, i, 1) = "ñ" Then
            TextoNovo = TextoNovo & "&ntilde;"
        ElseIf Mid(TxtInput.Text, i, 1) = "Ñ" Then
            TextoNovo = TextoNovo & "&Ntilde;"
        ElseIf Mid(TxtInput.Text, i, 1) = """" Then
            TextoNovo = TextoNovo & "&quot;"
        ElseIf Mid(TxtInput.Text, i, 1) = "'" Then
            TextoNovo = TextoNovo & "&#39;"
        Else
            TextoNovo = TextoNovo & Mid(TxtInput.Text, i, 1)
        End If
    Next i
    
    TxtOutput.Text = TextoNovo
    
    CmdTurn.Caption = "Turn in HTML Text"
End Sub
