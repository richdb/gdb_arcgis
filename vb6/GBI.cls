VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GBI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements ICommand
Public m_pMXDoc As IMxDocument
Public m_pApp As IApplication
Private m_pResources As FrmPic
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
ICommand_Bitmap = FrmPic.Picture1.Picture.Handle
End Property
Private Property Get ICommand_Caption() As String
ICommand_Caption = "ArcGIS GDB Tool Versie 2.0"
End Property
Private Property Get ICommand_Category() As String
ICommand_Category = "Provincie Drenthe"
End Property
Private Property Get ICommand_Checked() As Boolean
ICommand_Checked = False
End Property
Private Property Get ICommand_Enabled() As Boolean
ICommand_Enabled = True
End Property
Private Property Get ICommand_HelpContextID() As Long

End Property
Private Property Get ICommand_HelpFile() As String

End Property
Private Property Get ICommand_Message() As String
ICommand_Message = "ArcGIS GDB Tool Versie 2.0"
End Property
Private Property Get ICommand_Name() As String
ICommand_Name = "ArcGIS GDB Tool Versie 2.0"
End Property
Private Sub ICommand_OnClick()
Form1.Show vbModal
End Sub
Private Sub ICommand_OnCreate(ByVal hook As Object)
'Set m_pApp = hook
Form1.Application = hook
End Sub
Private Property Get ICommand_Tooltip() As String
ICommand_Tooltip = "ArcGIS GDB Tool Versie 2.0"
End Property
