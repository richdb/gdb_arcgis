VERSION 5.00
Begin VB.Form VariabelenSelectieForm 
   Caption         =   "Form2"
   ClientHeight    =   4452
   ClientLeft      =   56
   ClientTop       =   336
   ClientWidth     =   7448
   LinkTopic       =   "Form2"
   ScaleHeight     =   4452
   ScaleWidth      =   7448
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton WisBtn 
      Caption         =   "&Wissen"
      Height          =   392
      Left            =   3276
      TabIndex        =   13
      Top             =   3906
      Width           =   1274
   End
   Begin VB.CommandButton OKBtn 
      Caption         =   "&OK"
      Height          =   392
      Left            =   4662
      TabIndex        =   12
      Top             =   3906
      Width           =   1274
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Annuleren"
      Height          =   392
      Left            =   6048
      TabIndex        =   11
      Top             =   3906
      Width           =   1274
   End
   Begin VB.CommandButton OfBtn 
      Caption         =   "Of"
      Enabled         =   0   'False
      Height          =   392
      Left            =   4536
      TabIndex        =   10
      Top             =   2016
      Width           =   896
   End
   Begin VB.CommandButton EnBtn 
      Caption         =   "En"
      Enabled         =   0   'False
      Height          =   392
      Left            =   4536
      TabIndex        =   9
      Top             =   1512
      Width           =   896
   End
   Begin VB.ListBox List2 
      Height          =   1876
      ItemData        =   "VariabelenSelectieForm.frx":0000
      Left            =   126
      List            =   "VariabelenSelectieForm.frx":0002
      TabIndex        =   8
      Top             =   1764
      Width           =   2030
   End
   Begin VB.ComboBox Combo2 
      Height          =   294
      Left            =   4536
      TabIndex        =   6
      Top             =   882
      Width           =   2786
   End
   Begin VB.ListBox List1 
      Height          =   2786
      ItemData        =   "VariabelenSelectieForm.frx":0004
      Left            =   2394
      List            =   "VariabelenSelectieForm.frx":0006
      TabIndex        =   5
      Top             =   882
      Width           =   1904
   End
   Begin VB.ComboBox Combo1 
      Height          =   294
      Left            =   126
      TabIndex        =   4
      Top             =   882
      Width           =   2030
   End
   Begin VB.TextBox Text1 
      Height          =   266
      Left            =   126
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   126
      Width           =   7196
   End
   Begin VB.Label Label4 
      Caption         =   "Veld"
      Height          =   266
      Left            =   126
      TabIndex        =   7
      Top             =   1386
      Width           =   896
   End
   Begin VB.Label Label3 
      Caption         =   "Bereik"
      Height          =   266
      Left            =   4536
      TabIndex        =   3
      Top             =   504
      Width           =   1022
   End
   Begin VB.Label Label2 
      Caption         =   "Operator"
      Height          =   266
      Left            =   2394
      TabIndex        =   2
      Top             =   504
      Width           =   1022
   End
   Begin VB.Label Label1 
      Caption         =   "Tabel"
      Height          =   266
      Left            =   126
      TabIndex        =   1
      Top             =   504
      Width           =   896
   End
End
Attribute VB_Name = "VariabelenSelectieForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents db As ADODB.Connection
Attribute db.VB_VarHelpID = -1
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim querytekst As String
Dim tabelgekozen As String
Dim veldgekozen As String
Dim operatorgekozen As String
Dim bereikgekozen As String
Dim extrakomma As String
Dim welkomma As Boolean
Sub vullentabellen()
Dim SQL As String

On Error GoTo ErrorHandler
Set db = New ADODB.Connection
db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir

SQL = "SELECT DISTINCT TABLENAME FROM FIELDINFO;"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

With rs
    If .RecordCount > 1 Then
        .MoveFirst
        Do
            Combo1.AddItem .Fields("TABLENAME")
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
Private Sub Combo1_Click()
Dim nummer As Integer
Dim SQL As String
Dim RSField As ADODB.Field
Dim ColCnt As Long

On Error GoTo ErrorHandler

List2.Clear
nummer = Combo1.ListIndex
tabelgekozen = Combo1.List(nummer)

Set db = New ADODB.Connection
db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir

SQL = "SELECT * FROM " + tabelgekozen + ";"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

'ColCnt = rs.Fields.Count

For Each RSField In rs.Fields
    List2.AddItem RSField.Name
Next RSField

rs.Close
db.Close
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
    db.Close
End Sub
Private Sub Combo2_Click()
Dim nummer As Integer

On Error GoTo ErrorHandler

If Not Combo2.ListIndex = -1 Then
    nummer = Combo2.ListIndex
    bereikgekozen = Combo2.List(nummer)
    
    If welkomma = True Then
        querytekst = querytekst + bereikgekozen + extrakomma
        Text1.Text = querytekst
    Else
        querytekst = querytekst + bereikgekozen
        Text1.Text = querytekst
    End If
    
    EnBtn.Enabled = True
    OfBtn.Enabled = True
End If
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Private Sub Command3_Click()
Unload VariabelenSelectieForm
End Sub
Private Sub EnBtn_Click()
On Error GoTo ErrorHandler
querytekst = querytekst + " " + "AND "
Text1.Text = querytekst
OfBtn.Enabled = False
EnBtn.Enabled = False
Combo1.Enabled = True
List2.Enabled = True
List1.Enabled = True
List1.Clear
Combo2.Enabled = True
Combo2.Clear
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Private Sub Form_Load()
On Error GoTo ErrorHandler
VariabelenSelectieForm.Caption = Form1.versie + " - Variabelen Selectie"
vullentabellen
querytekst = Text1.Text
EnBtn.Enabled = False
OfBtn.Enabled = False
extrakomma = "'"
welkomma = False
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Private Sub List1_DblClick()
Dim nummer As Integer
Dim SQL As String

On Error GoTo ErrorHandler
welkomma = False
nummer = List1.ListIndex
operatorgekozen = List1.List(nummer)

If (operatorgekozen = "IS") Then
    operatorgekozen = "LIKE '"
    welkomma = True
ElseIf (operatorgekozen = "ONGELIJK") Then
    operatorgekozen = "NOT LIKE '"
    welkomma = True
ElseIf (operatorgekozen = "BEVAT") Then
    operatorgekozen = "LIKE *'"
    welkomma = True
ElseIf (operatorgekozen = "BEVAT NIET") Then
    operatorgekozen = "NOT LIKE *'"
    welkomma = True
End If

If welkomma = True Then
    querytekst = querytekst + " " + operatorgekozen
Else
    querytekst = querytekst + " " + operatorgekozen + " "
End If
Text1.Text = querytekst

List1.Enabled = False

Combo2.Clear

Set db = New ADODB.Connection
db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir

SQL = "SELECT DISTINCT " + veldgekozen + " FROM " + tabelgekozen + " ORDER BY " + veldgekozen + ";"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

With rs
    If .RecordCount > 1 Then
        .MoveFirst
        Do
            Combo2.AddItem .Fields(veldgekozen)
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
Private Sub List2_DblClick()
Dim nummer As Integer
Dim SQL As String
Dim RSField As ADODB.Field
Dim soort As String

On Error GoTo ErrorHandler

Set db = New ADODB.Connection
db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir

nummer = List2.ListIndex
veldgekozen = List2.List(nummer)

SQL = "SELECT " + veldgekozen + " FROM " + tabelgekozen + ";"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

For Each RSField In rs.Fields
     soort = RSField.Type
Next RSField

rs.Close
db.Close

If (soort = "3" Or soort = "5") Then
    List1.Clear
    List1.AddItem "<"
    List1.AddItem "<="
    List1.AddItem "="
    List1.AddItem ">"
    List1.AddItem ">="
    List1.AddItem "<>"
End If

If (soort = "202") Then
    List1.Clear
    List1.AddItem "IS"
    List1.AddItem "ONGELIJK"
    List1.AddItem "BEVAT"
    List1.AddItem "BEVAT NIET"
End If

querytekst = querytekst + tabelgekozen + "." + veldgekozen
Text1.Text = querytekst
Combo1.Enabled = False
List2.Enabled = False
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Private Sub OfBtn_Click()
On Error GoTo ErrorHandler
querytekst = querytekst + " " + "OR "
Text1.Text = querytekst
OfBtn.Enabled = False
EnBtn.Enabled = False
Combo1.Enabled = True
List2.Enabled = True
List1.Enabled = True
List1.Clear
Combo2.Enabled = True
Combo2.Clear
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Private Sub OKBtn_Click()
SelectieForm.VariabelenTextbox.Text = Text1.Text
Unload VariabelenSelectieForm
End Sub
Private Sub WisBtn_Click()
On Error GoTo ErrorHandler
Text1.Text = ""
Combo1.Enabled = True
List2.Enabled = True
List1.Enabled = True
List1.Clear
Combo2.Enabled = True
Combo2.Clear
EnBtn.Enabled = False
OfBtn.Enabled = False
querytekst = ""
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
