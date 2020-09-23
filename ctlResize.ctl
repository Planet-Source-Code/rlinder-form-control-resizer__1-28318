VERSION 5.00
Begin VB.UserControl ctlResize 
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   ControlContainer=   -1  'True
   PropertyPages   =   "ctlResize.ctx":0000
   ScaleHeight     =   4065
   ScaleWidth      =   5880
   ToolboxBitmap   =   "ctlResize.ctx":000E
End
Attribute VB_Name = "ctlResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Private Type Size
    lonX As Long
    lonY As Long
    width As Long
    height As Long
End Type


Dim ControlInit As Boolean
Dim ControlCount As Long
Dim ControlSizes() As Size

Dim dblX As Double
Dim dblY As Double

'Default Property Values:
Const m_def_orgWidth = 2000
Const m_def_orgHeight = 2000
'Property Variables:
Dim m_orgWidth As Single
Dim m_orgHeight As Single



Private Sub SizeProc()
    Dim i
    
    On Error Resume Next
    
    
    If Ambient.UserMode = True And UserControl.ContainedControls.Count > 0 Then
       
       If ControlInit = False Then
          ControlCount = UserControl.ContainedControls.Count
          
          ReDim ControlSizes(ControlCount)
          
          For i = 0 To ControlCount - 1
                With ControlSizes(i)
                    .lonX = UserControl.ContainedControls(i).Left
                    .lonY = UserControl.ContainedControls(i).Top
                    .width = UserControl.ContainedControls(i).width
                    .height = UserControl.ContainedControls(i).height
                End With
          Next i
             
          ControlInit = True
          
       End If
       
       
          dblX = m_orgWidth / UserControl.width
          dblY = m_orgHeight / UserControl.height
                
          For i = 0 To ControlCount - 1
              UserControl.ContainedControls(i).Left = ControlSizes(i).lonX / dblX
              UserControl.ContainedControls(i).Top = ControlSizes(i).lonY / dblY
              UserControl.ContainedControls(i).width = ControlSizes(i).width / dblX
              UserControl.ContainedControls(i).height = ControlSizes(i).height / dblY
          Next
        
    End If
End Sub

Private Sub UserControl_Initialize()
    ControlInit = False
End Sub

Private Sub UserControl_Paint()
    Call SizeProc
End Sub

Private Sub UserControl_Resize()
    If Ambient.UserMode = False Then
       m_orgWidth = UserControl.width
       m_orgHeight = UserControl.height
       PropertyChanged "orgWidth"
       PropertyChanged "orgHeight"
    End If
    
    Call SizeProc
End Sub

Private Sub UserControl_Terminate()
    ControlInit = False
End Sub
Public Property Get orgWidth() As Single
    orgWidth = m_orgWidth
End Property

Public Property Let orgWidth(ByVal New_orgWidth As Single)
    If Ambient.UserMode Then Err.Raise 393
    m_orgWidth = New_orgWidth
    PropertyChanged "orgWidth"
End Property

Public Property Get orgHeight() As Single
    orgHeight = m_orgHeight
End Property

Public Property Let orgHeight(ByVal New_orgHeight As Single)
    m_orgHeight = New_orgHeight
    PropertyChanged "orgHeight"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_orgWidth = m_def_orgWidth
    m_orgHeight = m_def_orgHeight
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_orgWidth = PropBag.ReadProperty("orgWidth", m_def_orgWidth)
    m_orgHeight = PropBag.ReadProperty("orgHeight", m_def_orgHeight)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H40C0&)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("orgWidth", m_orgWidth, m_def_orgWidth)
    Call PropBag.WriteProperty("orgHeight", m_orgHeight, m_def_orgHeight)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H40C0&)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

