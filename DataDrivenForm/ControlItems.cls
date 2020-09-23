VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum CntrlType
    lwTextBox
    lwLabel
    lwCommandButton
    lwTreeView
End Enum
Event Validate(ctrlItem As ControlItem, CtlType As CntrlType, Cancel As Boolean)
Event DragDrop(ctrlItem As ControlItem, CtlType As CntrlType, Source As Control, x As Single, Y As Single)
Event LostFocus(Item As ControlItem, CtlType As CntrlType)
Event Click(Item As ControlItem, CtlType As CntrlType)

Private m_ControlItems As New Collection

' Add a new ControlItem item to the collection
Function Add(ctrl As Control, CtlType As CntrlType) As ControlItem
    Dim newItem As New ControlItem
    
    newItem.Init ctrl, Me, CtlType
    m_ControlItems.Add newItem
    Set Add = newItem
End Function

Friend Sub Notify_Validate(Item As ControlItem, CtlType As CntrlType, Cancel As Boolean)
    If CtlType = lwTextBox Then
        If Item.txtCtrl.Text = "" Then
            MsgBox "Invalid Input"
            Cancel = True
        End If
    End If
    '   RaiseEvent Validate(Item, Cancel)
End Sub
Friend Sub Notify_Click(Item As ControlItem, CtlType As CntrlType)
    Dim x As ControlItem, strMsg As String
    If Item.cmdCtrl.Name = "cmdOk" Then
        For i = 1 To m_ControlItems.Count
            Set x = m_ControlItems.Item(i)
            If x.m_ParentType = lwTextBox Then
               strMsg = strMsg & x.txtCtrl.Text & vbCrLf
            End If
        Next
        MsgBox strMsg
    ElseIf Item.cmdCtrl.Name = "cmdCancel" Then
        Unload frmParent
    End If
End Sub

Public Sub Notify_LostFocus(Item As ControlItem, CtlType As CntrlType)
    If CtlType = lwTextBox Then
        If Item.txtCtrl.Name = "txtCategory" Then
            UpdateTree Item.txtCtrl.Text, Item
        End If
    End If
End Sub

Public Sub UpdateTree(strText As String, Item As ControlItem)
   Dim i As Integer, aString() As String
   Dim ctrlItem As ControlItem
   Set ctrlItem = GetItem("tv", lwTreeView)
   aString = Split(strText, "/")
   ctrlItem.ctlDynamic.object.nodes.Clear
    ctrlItem.ctlDynamic.object.nodes.Add Key:=aString(0) & Str(0), Text:=aString(0)
   For i = 1 To UBound(aString)
        ctrlItem.ctlDynamic.object.nodes.Add Relative:=aString(i - 1) & Str(i - 1), Relationship:=4, Key:=aString(i) & Str(i), Text:=aString(i)
        ctrlItem.ctlDynamic.object.nodes.Item(aString(i - 1) & Str(i - 1)).Expanded = True
   Next
   ctrlItem.ctlDynamic.Visible = True
End Sub

Private Function GetItem(ctrlName As String, CtlType As CntrlType) As ControlItem
    
    Dim TempItem As ControlItem
    For i = 1 To m_ControlItems.Count
        
        Set TempItem = m_ControlItems.Item(i)
        
        If TempItem.m_ParentType = lwTreeView Then
        Select Case CtlType
            Case lwTreeView
                If UCase(TempItem.ctlDynamic.Name) = UCase(ctrlName) Then
                    Set GetItem = TempItem
                    Exit For
                End If
        End Select
        End If
    Next

End Function
