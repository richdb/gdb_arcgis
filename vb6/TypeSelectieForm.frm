VERSION 5.00
Begin VB.Form TypeSelectieForm 
   Caption         =   "Form2"
   ClientHeight    =   3122
   ClientLeft      =   56
   ClientTop       =   336
   ClientWidth     =   6356
   LinkTopic       =   "Form2"
   ScaleHeight     =   3122
   ScaleWidth      =   6356
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKBtn 
      Caption         =   "&OK"
      Height          =   392
      Left            =   3402
      TabIndex        =   7
      Top             =   2520
      Width           =   1274
   End
   Begin VB.CommandButton AnnuleerBtn 
      Caption         =   "&Annuleren"
      Height          =   392
      Left            =   4788
      TabIndex        =   6
      Top             =   2520
      Width           =   1274
   End
   Begin VB.CommandButton VerwijderBtn 
      Caption         =   "< Verwijder"
      Height          =   392
      Left            =   2646
      TabIndex        =   5
      Top             =   1512
      Width           =   1022
   End
   Begin VB.CommandButton SelecteerBtn 
      Caption         =   "Selecteer >"
      Height          =   392
      Left            =   2646
      TabIndex        =   4
      Top             =   756
      Width           =   1022
   End
   Begin VB.ListBox List2 
      Height          =   1876
      ItemData        =   "TypeSelectieForm.frx":0000
      Left            =   3906
      List            =   "TypeSelectieForm.frx":0002
      TabIndex        =   3
      Top             =   378
      Width           =   2156
   End
   Begin VB.ListBox List1 
      Height          =   1876
      ItemData        =   "TypeSelectieForm.frx":0004
      Left            =   252
      List            =   "TypeSelectieForm.frx":0006
      TabIndex        =   2
      Top             =   378
      Width           =   2156
   End
   Begin VB.Label Label2 
      Caption         =   "Geselecteerd"
      Height          =   266
      Left            =   4536
      TabIndex        =   1
      Top             =   126
      Width           =   1022
   End
   Begin VB.Label Label1 
      Caption         =   "Te Selecteren"
      Height          =   266
      Left            =   756
      TabIndex        =   0
      Top             =   126
      Width           =   1400
   End
End
Attribute VB_Name = "TypeSelectieForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents db As ADODB.Connection
Attribute db.VB_VarHelpID = -1
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Sub vullentypelijst()
Dim SQL As String

On Error GoTo ErrorHandler

Set db = New ADODB.Connection
db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir

SQL = "SELECT NAAM FROM TYPETABEL ORDER BY TYPE;"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

With rs
    If .RecordCount > 1 Then
        .MoveFirst
        Do
            List1.AddItem .Fields("NAAM")
            .MoveNext
        Loop Until .EOF
    End If
    .Close
End With
db.Close
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
    db.Close
End Sub
Private Sub AnnuleerBtn_Click()
Unload TypeSelectieForm
End Sub
Private Sub Form_Load()
TypeSelectieForm.Caption = Form1.versie + " - Type Selectie"
vullentypelijst
End Sub
Private Sub OKBtn_Click()
Dim selectie As String
Dim tussenvoeg As String
Dim bestand As String
Dim i As Integer

On Error GoTo ErrorHandler

selectie = ""
tussenvoeg = ";"

For i = 0 To (List2.ListCount - 1)
    If (i = (List2.ListCount - 1)) Then
        bestand = List2.List(i)
        selectie = selectie + bestand
    Else
        bestand = List2.List(i)
        selectie = selectie + bestand + tussenvoeg
    End If
Next i

SelectieForm.TypeTextbox.Text = selectie
Unload TypeSelectieForm
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Private Sub SelecteerBtn_Click()
Dim nummer As Integer
Dim waarde As String

On Error GoTo ErrorHandler

If List1.ListIndex = -1 Then
    MsgBox "Selecteer eerst een selectietype uit de lijst!", vbOKOnly, Form1.versie
    Exit Sub
End If

nummer = List1.ListIndex
waarde = List1.List(nummer)

List2.AddItem waarde
List1.RemoveItem nummer
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Private Sub VerwijderBtn_Click()
Dim nummer As Integer
Dim waarde As String

On Error GoTo ErrorHandler
If List2.ListIndex = -1 Then
    MsgBox "Selecteer eerst een selectietype uit de lijst!", vbOKOnly, Form1.versie
    Exit Sub
End If

nummer = List2.ListIndex
waarde = List2.List(nummer)

List1.AddItem waarde
List2.RemoveItem nummer
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
