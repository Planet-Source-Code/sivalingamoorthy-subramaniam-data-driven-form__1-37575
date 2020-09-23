VERSION 5.00
Begin VB.Form frmParent 
   Caption         =   "Data Driven Form Example"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Sivalingamoorthy Subramaniam

'Feedback appreciated
'If you have any suggestions for additions or changes to the program, please let me know.
'*--------------------------------*
' I tryed lot to make a Data Driven forms dynamically, after working for few
' hours i came out with the solution. You can try more with this basic codings
'*--------------------------------*

'Category Value = "PSCode/Abc/Testing"

Private Sub Form_Load()
    ' Calling the Function to create New Controls in the form
    AddFormControls Me
End Sub

Public Sub AddFormControls(frm As Form)
    
    
    Dim CtrlItems As New ControlItems
    Dim ctrl As Control

    Set ctrl = Controls.Add("VB.Label", "lblFile", frm)
    ctrl.Move 150, 200, 1000, 300
    ctrl.Caption = "File Path"
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwLabel
   
    Set ctrl = Controls.Add("VB.TextBox", "txtPath", frm)
    ctrl.Move 2000, 200, 4000, 315
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwTextBox
    

    Set ctrl = Controls.Add("VB.Label", "lblLabel", frm)
    ctrl.Move 150, 700, 1000, 300
    ctrl.Caption = "File Label"
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwLabel

    Set ctrl = Controls.Add("VB.TextBox", "txtLabel", frm)
    ctrl.Move 2000, 700, 4000, 315
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwTextBox
    
    Set ctrl = Controls.Add("VB.Label", "lblDesc", frm)
    ctrl.Move 150, 1200, 4000, 300
    ctrl.Caption = "File Describtion"
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwLabel

    Set ctrl = Controls.Add("VB.TextBox", "txtDescribtion", frm)
    ctrl.Move 2000, 1200, 4000, 315
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwTextBox
    
    Set ctrl = Controls.Add("VB.Label", "lblCategory", frm)
    ctrl.Move 150, 1700, 4000, 300
    ctrl.Caption = "File Category"
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwLabel

    Set ctrl = Controls.Add("VB.TextBox", "txtCategory", frm)
    'ctrl.MultiLine = True
    ctrl.Move 2000, 1700, 4000, 315
    'ctr1.Text = "PSCode\VB\Testing\DataDrivenForms"
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwTextBox
    
'    Add a New License to use VB Control
    Licenses.Add "MSComctlLib.TreeCtrl"
    Set ctrl = Controls.Add("MSComctlLib.TreeCtrl", "Tv", frm)
   ' Add some nodes to the control.
    ctrl.Move 2000, 2200, 4000, 2000
    ctrl.object.Style = 7
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwTreeView

    Set ctrl = Controls.Add("VB.CommandButton", "cmdOk", frm)
    ctrl.Move 2000, 4400, 1000, 315
    ctrl.Caption = "&Ok"
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwCommandButton
    
    Set ctrl = Controls.Add("VB.CommandButton", "cmdCancel", frm)
    ctrl.Move 3500, 4400, 1000, 315
    ctrl.Caption = "&Cancel"
    ctrl.Visible = True
    ' Adding new Object into the Control Items Collection
    CtrlItems.Add ctrl, lwCommandButton

End Sub
