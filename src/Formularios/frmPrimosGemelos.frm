VERSION 5.00
Begin VB.Form frmPrimosGemelos 
   Caption         =   "Primos Gemelos --- Twins Prime"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrimosGemelos 
      Caption         =   "Primos Gemelos"
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrimosGemelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' AL CARGAR EL FORMULARIO
Private Sub Form_Load()

End Sub

' EJECUTAR PRIMOS GEMELOS
Private Sub cmdPrimosGemelos_Click()
  Dim p As Long
  Dim multiplo As Long
  Dim resto As Long

  For p = 3 To 1000
    If Primo(p) Then
      If Primo(p - 2) Then
        multiplo = p * (p - 2)
        resto = multiplo Mod 9
        Print p; "*"; (p - 2); "="; multiplo; " mod 9  = "; resto
      End If
    End If

  Next p
End Sub


' FUNCION PARA CALCULAR SI EL NUMERO ES PRIMO
Public Function Primo(ByVal pN As Long) As Boolean
  Dim i As Long
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function

' FUNCION PARA CALCULAR SI EL NUMERO ES PRIMO
Public Function Tabulado(ByVal pT As String, ByVal pA As Integer) As String
  Dim i As Integer
  Dim miAncho As Integer
  miAncho = Len(Trim(pT))

  For i = 1 To (pA - miAncho)
    'pT = pT + " "
    pT = " " + pT
  Next i
  Tabulado = pT
End Function
