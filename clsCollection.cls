VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function SortCollection(ByVal ColToSort As Collection, ByVal arrCol As Variant, ByVal blnItemKey, Optional ByVal ColSortType As VbCompareMethod = vbTextCompare) As Collection
    Dim EachItem As Long
    Dim ColTemp As New Collection
    Dim OrderNotChanged As Boolean
    Dim ArrayItems() As String
    Dim FirstStringInArray As String
    Dim NumberOfItems As Long
    Dim strCol As String

    NumberOfItems = ColToSort.Count
    
    If NumberOfItems = 0 Then
        Exit Function
    End If

    ReDim ArrayItems(NumberOfItems - 1)

    For EachItem = 1 To NumberOfItems
        
        If blnItemKey Then  'Trabalha por Item
            ArrayItems(EachItem - 1) = ColToSort.Item(EachItem)
        Else                'Trabalha por Key
            ArrayItems(EachItem - 1) = arrCol(EachItem - 1)
        End If
        
    Next EachItem

    Do
      OrderNotChanged = True
          
        For EachItem = 1 To NumberOfItems - 1

            If Strings.StrComp(ArrayItems(EachItem - 1), ArrayItems(EachItem), ColSortType) = 1 Then
                FirstStringInArray = ArrayItems(EachItem - 1)
                ArrayItems(EachItem - 1) = ArrayItems(EachItem)
                ArrayItems(EachItem) = FirstStringInArray
                OrderNotChanged = False
            End If
            
          Next EachItem
    
    Loop Until OrderNotChanged

    For EachItem = 0 To NumberOfItems - 1
        ColTemp.Add ArrayItems(EachItem)
    Next EachItem
    
    Set SortCollection = ColTemp
End Function

Sub LimpaControles(frm As Form)
   
    For i = 0 To frm.Controls.Count - 1
        
        If TypeOf frm.Controls(i) Is TextBox Then
            frm.Controls(i).Text = ""
        End If
    
    Next

End Sub

Public Function getKeyCollection(ByVal obj As Collection, ByVal texto As String) As String
Dim i As Byte

    For i = 1 To obj.Count
        
        If UCase(obj.Item(i)) = UCase(texto) Then
            getKeyCollection = i
            Exit Function
        End If
    
    Next

End Function

Public Sub LimpaCollection(ByRef coll As Collection)
Dim a As Long

    For a = coll.Count To 1 Step -1
        coll.Remove 1
    Next

    'Set coll = Nothing
End Sub
