VERSION 5.00
Begin VB.Form SelectieForm 
   Caption         =   "Form2"
   ClientHeight    =   2464
   ClientLeft      =   56
   ClientTop       =   336
   ClientWidth     =   8554
   LinkTopic       =   "Form2"
   ScaleHeight     =   2464
   ScaleWidth      =   8554
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKBtn 
      Caption         =   "&OK"
      Height          =   392
      Left            =   5292
      TabIndex        =   10
      Top             =   1890
      Width           =   1526
   End
   Begin VB.CommandButton AnnuleerBtn 
      Caption         =   "&Annuleren"
      Height          =   392
      Left            =   6930
      TabIndex        =   9
      Top             =   1890
      Width           =   1400
   End
   Begin VB.TextBox VariabelenTextbox 
      Height          =   266
      Left            =   3150
      TabIndex        =   8
      Top             =   1260
      Width           =   5180
   End
   Begin VB.TextBox TrefwoordTextbox 
      Height          =   266
      Left            =   3150
      TabIndex        =   7
      Top             =   756
      Width           =   5180
   End
   Begin VB.TextBox TypeTextbox 
      Height          =   266
      Left            =   3150
      TabIndex        =   6
      Top             =   252
      Width           =   5180
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Wijzig..."
      Height          =   392
      Left            =   1512
      TabIndex        =   5
      Top             =   1197
      Width           =   1274
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Wijzig..."
      Height          =   392
      Left            =   1512
      TabIndex        =   4
      Top             =   693
      Width           =   1274
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Wijzig..."
      Height          =   392
      Left            =   1512
      TabIndex        =   3
      Top             =   189
      Width           =   1274
   End
   Begin VB.Label Label3 
      Caption         =   "Variabelen"
      Height          =   266
      Left            =   126
      TabIndex        =   2
      Top             =   1260
      Width           =   1400
   End
   Begin VB.Label Label2 
      Caption         =   "Trefwoord"
      Height          =   266
      Left            =   126
      TabIndex        =   1
      Top             =   756
      Width           =   1274
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
      Height          =   266
      Left            =   126
      TabIndex        =   0
      Top             =   252
      Width           =   1148
   End
End
Attribute VB_Name = "SelectieForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents db As ADODB.Connection
Attribute db.VB_VarHelpID = -1
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private Sub AnnuleerBtn_Click()
Unload SelectieForm
End Sub
Private Sub Command1_Click()
TypeSelectieForm.Show vbModal
End Sub
Private Sub Command2_Click()
TrefwoordSelectieForm.Show vbModal
End Sub
Private Sub Command3_Click()
VariabelenSelectieForm.Show vbModal
End Sub
Private Sub Form_Load()
SelectieForm.Caption = Form1.versie + " - Selectie"
End Sub
Private Sub OKBtn_Click()
Dim lengte As Long
Dim j As Long
Dim waarde As String
Dim positie As Long
Dim nw As String
Dim querywaarde As String
Dim querywaarde1 As String
Dim querywaarde2 As String
Dim zoekwaarde As String
Dim thema As Integer
Dim SQL As String
Dim nwtabel As String
Dim meer As Boolean
Dim tabelatt As String
Dim linkatt As String
Dim tabelcd As String
Dim linkcd As String
Dim tabelgeo As String
Dim linkgeo As String
Dim tabelit As String
Dim linkit As String
Dim tabelprog As String
Dim linkprog As String

On Error GoTo ErrorHandler

meer = False
If Not TypeTextbox.Text = "" And Not TrefwoordTextbox.Text = "" And Not VariabelenTextbox.Text = "" Then
    zoekwaarde = ";"
    waarde = TypeTextbox.Text
    lengte = Len(waarde)
    positie = 0
    j = 0
    Do While j < lengte
        positie = InStr((1 + j), waarde, zoekwaarde)
        If (positie = 0) Then
            nw = Mid$(waarde, (1 + j), lengte)
            querywaarde1 = querywaarde1 + "b.TYPE = '" + nw + "'"
            j = lengte
        Else
            nw = Mid(waarde, (1 + j), (positie - 1))
            querywaarde1 = querywaarde1 + "b.TYPE = '" + nw + "' OR "
            j = positie
        End If
    Loop

    zoekwaarde = ";"
    waarde = TrefwoordTextbox.Text
    lengte = Len(waarde)
    positie = 0
    j = 0
    Do While j < lengte
        positie = InStr((1 + j), waarde, zoekwaarde)
        If (positie = 0) Then
            nw = Mid$(waarde, (1 + j), lengte)
            querywaarde2 = querywaarde2 + "b.TREFWOORD = '" + nw + "'"
            j = lengte
        Else
            nw = Mid(waarde, (1 + j), (positie - 1))
            querywaarde2 = querywaarde2 + "b.TREFWOORD = '" + nw + "' OR "
            j = positie
        End If
    Loop
    
    j = 0
    waarde = VariabelenTextbox.Text
    lengte = Len(waarde)
    positie = InStr((1 + j), waarde, ".")
    nwtabel = Mid(waarde, (1 + j), positie)
    
    If (InStr((1 + j), waarde, "ATTRIBUUT")) Then
        tabelatt = ", ATTRIBUUT ATTRIBUUT"
        linkatt = " AND DATASET.DATACODE = ATTRIBUUT.DATACODE"
    End If
    
    If (InStr((1 + j), waarde, "CDROM")) Then
        tabelcd = ", SELECTIE SELECTIE"
        linkcd = " AND DATASET.DATACODE = SELECTIE.DATACODE"
    End If
    
    If (InStr((1 + j), waarde, "GEOGRAFISCH")) Then
        tabelcd = ", GEOGRAFISCH GEOGRAFISCH"
        linkcd = " AND DATASET.DATACODE = GEOGRAFISCH.DATACODE"
    End If

    If (InStr((1 + j), waarde, "ITEMS")) Then
        tabelit = ", ITEMS ITEMS"
        linkit = " AND DATASET.DATACODE = ITEMS.DATACODE"
    End If

    If (InStr((1 + j), waarde, "PROGRAMMA")) Then
        tabelprog = ", PROGRAMMA PROGRAMMA"
        linkprog = " AND DATASET.DATACODE = PROGRAMMA.DATACODE"
    End If
    
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir
   
    SQL = "SELECT DATASET.BESTANDSTITEL, DATASET.DATACODE FROM DATASET DATASET," + _
        "GKTYPE b, TREFTEXT c, TREFCODE d" + tabelatt + tabelcd + tabelgeo + tabelit + tabelprog + _
        " WHERE (" + querywaarde1 + ") AND (" + querywaarde2 + ") AND DATASET.DATACODE = b.DATACODE AND " + _
        "DATASET.DATACODE = d.DATACODE AND c.TREFCODE = d.TREFCODE" + _
        linkatt + linkcd + linkgeo + linkit + linkprog + " AND " + waarde + " ORDER BY DATASET.BESTANDSTITEL;"
    
    Set rs = New ADODB.Recordset
    rs.Open SQL, db, adOpenStatic, adLockOptimistic

    Form1.List1.Clear
    thema = 0
    
    If rs.RecordCount > 0 Then
        With rs
            .MoveFirst
            Do
                Form1.List1.AddItem .Fields("BESTANDSTITEL")
                .MoveNext
                thema = thema + 1
            Loop Until .EOF
            .Close
        End With
    End If
    db.Close

    Form1.LabelNamen.Caption = "Namen (" + Str(thema) + " van " + Str(Form1.totaalmeta) + ")"
    Unload SelectieForm
    meer = True
    
End If

If Not TypeTextbox.Text = "" And meer = False Then
    zoekwaarde = ";"
    waarde = TypeTextbox.Text
    lengte = Len(waarde)
    positie = 0
    j = 0
    Do While j < lengte
        positie = InStr((1 + j), waarde, zoekwaarde)
        If (positie = 0) Then
            nw = Mid$(waarde, (1 + j), lengte)
            querywaarde = querywaarde + "b.TYPE = '" + nw + "'"
            j = lengte
        Else
            nw = Mid(waarde, (1 + j), (positie - 1))
            querywaarde = querywaarde + "b.TYPE = '" + nw + "' OR "
            j = positie
        End If
    Loop
    
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir

    SQL = "SELECT a.BESTANDSTITEL FROM DATASET a, GKTYPE b WHERE (" + querywaarde + ") AND " + _
        "a.DATACODE = b.DATACODE ORDER BY a.BESTANDSTITEL;"

    Set rs = New ADODB.Recordset
    rs.Open SQL, db, adOpenStatic, adLockOptimistic

    Form1.List1.Clear
    thema = 0

    If rs.RecordCount > 0 Then
        With rs
            .MoveFirst
            Do
                Form1.List1.AddItem .Fields("BESTANDSTITEL")
                .MoveNext
                thema = thema + 1
            Loop Until .EOF
            .Close
        End With
    End If
    db.Close

    Form1.LabelNamen.Caption = "Namen (" + Str(thema) + " van " + Str(Form1.totaalmeta) + ")"
    Unload SelectieForm
End If

If Not TrefwoordTextbox.Text = "" And meer = False Then
    zoekwaarde = ";"
    waarde = TrefwoordTextbox.Text
    lengte = Len(waarde)
    positie = 0
    j = 0
    Do While j < lengte
        positie = InStr((1 + j), waarde, zoekwaarde)
        If (positie = 0) Then
            nw = Mid$(waarde, (1 + j), lengte)
            querywaarde = querywaarde + "b.TREFWOORD = '" + nw + "'"
            j = lengte
        Else
            nw = Mid(waarde, (1 + j), (positie - 1))
            querywaarde = querywaarde + "b.TREFWOORD = '" + nw + "' OR "
            j = positie
        End If
    Loop
    
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir

    SQL = "SELECT DISTINCT a.BESTANDSTITEL FROM DATASET a, TREFTEXT b, TREFCODE c WHERE (" + querywaarde + ") AND " + _
        "a.DATACODE = c.DATACODE AND b.TREFCODE = c.TREFCODE ORDER BY a.BESTANDSTITEL;"
        
    Set rs = New ADODB.Recordset
    rs.Open SQL, db, adOpenStatic, adLockOptimistic
    
    Form1.List1.Clear
    thema = 0
    
    If rs.RecordCount > 0 Then
        With rs
            .MoveFirst
            Do
                Form1.List1.AddItem .Fields("BESTANDSTITEL")
                .MoveNext
                thema = thema + 1
            Loop Until .EOF
            .Close
        End With
    End If
    db.Close
    
    Form1.LabelNamen.Caption = "Namen (" + Str(thema) + " van " + Str(Form1.totaalmeta) + ")"
    Unload SelectieForm
End If

If Not VariabelenTextbox.Text = "" And meer = False Then
    j = 0
    waarde = VariabelenTextbox.Text
    lengte = Len(waarde)
    positie = InStr((1 + j), waarde, ".")
    nwtabel = Mid(waarde, (1 + j), positie)
    
    If (InStr((1 + j), waarde, "ATTRIBUUT")) Then
        tabelatt = ", ATTRIBUUT ATTRIBUUT"
        linkatt = " AND DATASET.DATACODE = ATTRIBUUT.DATACODE"
    End If
    
    If (InStr((1 + j), waarde, "CDROM")) Then
        tabelcd = ", SELECTIE SELECTIE"
        linkcd = " AND DATASET.DATACODE = SELECTIE.DATACODE"
    End If
    
    If (InStr((1 + j), waarde, "GEOGRAFISCH")) Then
        tabelcd = ", GEOGRAFISCH GEOGRAFISCH"
        linkcd = " AND DATASET.DATACODE = GEOGRAFISCH.DATACODE"
    End If

    If (InStr((1 + j), waarde, "ITEMS")) Then
        tabelit = ", ITEMS ITEMS"
        linkit = " AND DATASET.DATACODE = ITEMS.DATACODE"
    End If

    If (InStr((1 + j), waarde, "PROGRAMMA")) Then
        tabelprog = ", PROGRAMMA PROGRAMMA"
        linkprog = " AND DATASET.DATACODE = PROGRAMMA.DATACODE"
    End If
    
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir
        
    SQL = "SELECT DATASET.BESTANDSTITEL, DATASET.DATACODE FROM DATASET DATASET" + tabelatt + tabelcd + tabelgeo + _
    tabelit + tabelprog + " WHERE DATASET.DATACODE = DATASET.DATACODE" + linkatt + linkcd + linkgeo + linkit + linkprog + _
    " AND " + waarde + " ORDER BY DATASET.BESTANDSTITEL;"
    
    Set rs = New ADODB.Recordset
    rs.Open SQL, db, adOpenStatic, adLockOptimistic
    
    Form1.List1.Clear
    thema = 0
        
    If rs.RecordCount > 0 Then
        With rs
            .MoveFirst
            Do
                Form1.List1.AddItem .Fields("BESTANDSTITEL")
                .MoveNext
                thema = thema + 1
            Loop Until .EOF
            .Close
        End With
    End If
    db.Close
        
    Form1.LabelNamen.Caption = "Namen (" + Str(thema) + " van " + Str(Form1.totaalmeta) + ")"
    Unload SelectieForm
End If
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
    db.Close
End Sub
