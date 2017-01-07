VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1260
   ClientLeft      =   1230
   ClientTop       =   4065
   ClientWidth     =   2340
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   2340
   Begin VB.CommandButton cmdExibeFormulario 
      Caption         =   "&Exibe Formulário"
      Height          =   525
      Left            =   150
      TabIndex        =   1
      Top             =   480
      Width           =   1995
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   90
      Width           =   2025
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declara o objeto FormObjects como uma variável do tipo Collection
Private FormObjects As New Collection

' Carrega a coleção com os objetos formularios usando seus nomes como chave
Private Sub Form_Load()

    'inclui na combo o nome dos formulários
    Combo1.AddItem "Form2"
    Combo1.AddItem "Form3"
    Combo1.ListIndex = 0

    'incluimos na coleção os itens e a chave
    FormObjects.Add Form2, "Form2"
    FormObjects.Add Form3, "Form3"
End Sub

' Exibe o formulário de acordo com o nome na coleção
Private Sub cmdExibeFormulario_Click()
    Dim frm As Form
    On Error Resume Next
    Set frm = FormObjects(Combo1.Text)
    
    If Err.Number Then
        MsgBox "Nome do formulário desconhecido"
        Exit Sub
    End If

    On Error GoTo 0
    frm.Show
End Sub

' Descarrega todos os formularios
Private Sub Form_Unload(Cancel As Integer)

Dim frm As Form

   'para cada objeto frm na coleção Forms
    For Each frm In Forms
        Unload frm
    Next frm

End Sub
