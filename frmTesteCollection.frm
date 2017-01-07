VERSION 5.00
Begin VB.Form frmTesteCollection 
   Caption         =   "Teste Collection"
   ClientHeight    =   945
   ClientLeft      =   1305
   ClientTop       =   2265
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   945
   ScaleWidth      =   4485
   Begin VB.CommandButton cmdForm1 
      Caption         =   "&Form1"
      Height          =   705
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2085
   End
   Begin VB.CommandButton cmdCarrega 
      Caption         =   "&Carrega (Loads) - Veja no (See in) Immediate - CTRL+G"
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2085
   End
End
Attribute VB_Name = "frmTesteCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obj As Collection
Dim obj2 As Collection
Dim obj3 As Collection
Dim clsCol As clsCollection

Private Sub cmdCarrega_Click()
Dim arrString() As String
Dim i As Byte
    
    cmdCarrega.Enabled = False
    
    arrString = Split("I6,I5,I1,I4,I3,I2", ",")
    
    obj.Add "Melancia", "I6"
    obj.Add "Banana", "I5"
    obj.Add "Maçã", "I1"
    obj.Add "Pera", "I4"
    obj.Add "Mexerica", "I3"
    obj.Add "Laranja", "I2"
    Debug.Print "----- Inicio desordenado (Messy start)"

    For i = 1 To obj.Count
       Debug.Print obj.Item(i)
    Next

    Set obj2 = clsCol.SortCollection(obj, arrString, True)
    Set obj3 = clsCol.SortCollection(obj, arrString, False)

    Debug.Print "----- Ordenado por Chave (Sorted by Key):"

    For i = 1 To obj3.Count
       Debug.Print obj.Item(obj3.Item(i))
    Next

    Debug.Print "----- Ordenado por Nome (Sorted by Name):"

    For i = 1 To obj2.Count
       Debug.Print obj2.Item(i)
    Next

    Debug.Print "-----"
    Debug.Print "Quantidade de Itens da obj2 (Number of Items in obj2): " & obj2.Count
    
    clsCol.LimpaCollection obj2
    
    Debug.Print "Limpar obj2 (Clear obj2): " & obj2.Count

    Debug.Print "----- Fim"
End Sub

Private Sub cmdForm1_Click()
    Form1.Show
End Sub

Private Sub Form_Load()
    Set obj = New Collection
    Set obj2 = New Collection
    Set obj3 = New Collection
    Set clsCol = New clsCollection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set obj = Nothing
    Set obj2 = Nothing
    Set obj3 = Nothing
    Set clsCol = Nothing
End Sub
