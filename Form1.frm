VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Gerador de senhas"
   ClientHeight    =   5445
   ClientLeft      =   2145
   ClientTop       =   2670
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   4125
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   405
      Left            =   2760
      TabIndex        =   2
      Top             =   4920
      Width           =   1245
   End
   Begin VB.CommandButton cmdAutor 
      Caption         =   "&Autor"
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   4920
      Width           =   1245
   End
   Begin VB.CommandButton cmdGerarSenha 
      Caption         =   "&Gerar Senha"
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label lblSenhas 
      BorderStyle     =   1  'Fixed Single
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAutor_Click()
    MsgBox "Autor" & vbCrLf & "Herley Nicolas Ramos Sanchez" & vbCrLf & _
            "e-mail: Nicolas.MCP@gmail.com" & vbCrLf & vbCrLf & _
            "Licença GNU GPL (Software Livre)", vbOKOnly, "Autor"

End Sub

Private Sub cmdGerarSenha_Click()
Dim x As Integer, q As Integer, i As Integer
Dim sSen As String

x = Val(InputBox("Quantos caracteres quer na senha?", , 12))
q = Val(InputBox("Quantas senhas?", , 10))

lblSenhas.Caption = ""

For i = 1 To q
   sSen = GeraSenha(IIf(x > 1, x, 8))
   
   lblSenhas.Caption = lblSenhas.Caption & sSen & vbCrLf & vbCrLf
   
Next i

MsgBox "Senhas geradas!"


End Sub

'---------------------------------------------------------------------------------------
' Procedure : GeraSenha
' Data Hora : 08/07/2004 13:51
' Autor     : Nicolás Ramos
' Propósito : Generar una senha para uso comum
' Sintaxe   : GeraSenha([iTamaño])
'---------------------------------------------------------------------------------------
'
Function GeraSenha(Optional ByVal iTamaño As Integer = 8) As String
   Dim x As Integer, y As Integer
   Dim iIni As Integer, iFin As Integer
   Dim sAux As String
   Dim bEspecial As Boolean, bNumero As Boolean, bLetra As Boolean
   
   Randomize
Inicio:
   bEspecial = True: bNumero = True: bLetra = True
   iIni = 35: iFin = 122: sAux = ""
   
   For x = 1 To iTamaño
GenerarCaracter:
      y = Int((iFin - iIni + 1) * Rnd + iIni)
      If y = 38 Or y = 39 Or y = 44 Or y = 46 Or y = 58 _
         Or y = 59 Or (y >= 91 And y <= 96) Then GoTo GenerarCaracter
         
      If y <= 47 Or (y >= 60 And y <= 64) Then bEspecial = False
      If y >= 48 And y <= 57 Then bNumero = False
      If y >= 65 Then bLetra = False
      sAux = sAux & Chr(y)
   Next x
   If bEspecial Or bNumero Or bLetra Then GoTo Inicio
   GeraSenha = sAux
   DoEvents
End Function

Private Sub cmdSair_Click()
    Unload Me
        
End Sub
