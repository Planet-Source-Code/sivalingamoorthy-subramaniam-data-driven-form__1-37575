VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public WithEvents ctrl As VBControlExtender
Public WithEvents txtCtrl As TextBox
Attribute txtCtrl.VB_VarHelpID = -1
Public WithEvents lblCtrl As Label
Attribute lblCtrl.VB_VarHelpID = -1
Public WithEvents cmdCtrl As CommandButton
Attribute cmdCtrl.VB_VarHelpID = -1
Public WithEvents ctlDynamic As VBControlExtender
Attribute ctlDynamic.VB_VarHelpID = -1

Dim m_Parent As ControlItems
Public m_ParentType As CntrlType

' Handling type of objects which is been added to the collection
Sub Init(ctl As Object, parnt As ControlItems, CtlType As CntrlType)
    Select Case CtlType
        Case lwTextBox
            Set txtCtrl = ctl
        Case lwCommandButton
            Set cmdCtrl = ctl
        Case lwLabel
            Set lblCtrl = ctl
        Case lwTreeView
            Set ctlDynamic = ctl
    End Select
    m_ParentType = CtlType
    Set m_Parent = parnt
End Sub
Private Sub txtCtrl_LostFocus()
    m_Parent.Notify_LostFocus Me, m_ParentType
End Sub

Private Sub txtCtrl_Validate(Cancel As Boolean)
    m_Parent.Notify_Validate Me, m_ParentType, Cancel
End Sub

Private Sub cmdCtrl_Click()
    m_Parent.Notify_Click Me, m_ParentType
End Sub
