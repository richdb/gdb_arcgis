VERSION 5.00
Begin VB.Form OverzichtForm 
   Caption         =   "Form2"
   ClientHeight    =   2730
   ClientLeft      =   56
   ClientTop       =   336
   ClientWidth     =   7504
   LinkTopic       =   "Form2"
   ScaleHeight     =   2730
   ScaleWidth      =   7504
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Gehele database"
      Height          =   1778
      Left            =   126
      TabIndex        =   4
      Top             =   126
      Width           =   3416
      Begin VB.OptionButton Option3 
         Caption         =   "CD gegevens genereren"
         Height          =   266
         Left            =   252
         TabIndex        =   7
         Top             =   1008
         Width           =   2912
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Niet openbare GBI overzicht genereren"
         Height          =   266
         Left            =   252
         TabIndex        =   6
         Top             =   630
         Width           =   3038
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Openbare GBI overzicht genereren"
         Height          =   266
         Left            =   252
         TabIndex        =   5
         Top             =   252
         Width           =   3038
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trefwoord selectie"
      Height          =   1778
      Left            =   3780
      TabIndex        =   2
      Top             =   126
      Width           =   3416
      Begin VB.OptionButton Option5 
         Caption         =   "Alle metadata"
         Height          =   266
         Left            =   252
         TabIndex        =   9
         Top             =   1134
         Width           =   2786
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Openbare GBI overzicht genereren"
         Height          =   266
         Left            =   252
         TabIndex        =   8
         Top             =   756
         Width           =   2912
      End
      Begin VB.ComboBox Combo1 
         Height          =   294
         Left            =   252
         TabIndex        =   3
         Text            =   "Selecteer een trefwoord"
         Top             =   252
         Width           =   2912
      End
   End
   Begin VB.CommandButton AnnuleerBtn 
      Caption         =   "&Annuleren"
      Height          =   392
      Left            =   3780
      TabIndex        =   1
      Top             =   2142
      Width           =   1400
   End
   Begin VB.CommandButton StartBtn 
      Caption         =   "&Starten"
      Height          =   392
      Left            =   2142
      TabIndex        =   0
      Top             =   2142
      Width           =   1400
   End
End
Attribute VB_Name = "OverzichtForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents db As ADODB.Connection
Attribute db.VB_VarHelpID = -1
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private WithEvents rs2 As ADODB.Recordset
Attribute rs2.VB_VarHelpID = -1
Sub vullentrefwoordlijst()
Dim SQL As String

On Error GoTo ErrorHandler

Set db = New ADODB.Connection
db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir

SQL = "SELECT TREFWOORD FROM TREFTEXT ORDER BY TREFWOORD;"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

With rs
    If .RecordCount > 1 Then
        .MoveFirst
        Do
            Combo1.AddItem .Fields("TREFWOORD")
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
Unload OverzichtForm
End Sub
Private Sub Form_Load()
OverzichtForm.Caption = Form1.versie + " - Overzicht"
vullentrefwoordlijst
End Sub
Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
End Sub
Private Sub Option2_Click()
Option1.Value = False
Option2.Value = True
Option3.Value = False
Option4.Value = False
Option5.Value = False
End Sub
Private Sub Option3_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = True
Option4.Value = False
Option5.Value = False
End Sub
Private Sub Option4_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = True
Option5.Value = False
End Sub
Private Sub Option5_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = True
End Sub
Private Sub StartBtn_Click()
Dim w1 As Word.Application
Dim otable As Word.Table
Dim SQL As String
Dim SQL2 As String
Dim bestandtitel As String
Dim nw As String
Dim teller As Long
Dim lengte As Long
Dim datacode As String
Dim aantalrij As Integer
Dim i As Integer

If Option1.Value = True Then
    Form1.MousePointer = vbHourglass
    Set w1 = New Word.Application
    w1.Visible = True
    w1.WindowState = wdWindowStateMaximize
    w1.Documents.Add
    
    w1.Selection.Font.Size = 10
    w1.Selection.Font.Bold = True
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    w1.Selection.TypeText "Provincie Drenthe"
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = False
    w1.Selection.TypeText "Westerbrink 1"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Postbus 122"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "9400 AC Assen"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Contactpersoon"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Gerk van der Ploeg"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "0592 365463"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "G.vanderPloeg@drenthe.nl"
    w1.Selection.TypeParagraph
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = True
    w1.Selection.Font.Size = 12
    w1.Selection.TypeText "Openbare GBI gegevens per " + Str(Date) + ":"
    w1.Selection.Font.Bold = False
    w1.Selection.Font.Size = 10
    w1.Selection.TypeParagraph
    
    If w1.Documents.Count < 1 Then
        MsgBox "Er zijn geen documenten open!"
        Exit Sub
    End If
    
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir
       
    SQL = "SELECT DATASET.DATACODE, DATASET.BESTANDSTITEL, B.TEXT, DATASET.NAAM FROM DATASET DATASET, MEMOTABEL B" + _
            " WHERE DATASET.OMSCHRIJVING = B.CODE AND DATASET.GEBRUIKSBEPERKING = ""openbaar"" ORDER BY DATASET.BESTANDSTITEL;"
            
    Set rs = New ADODB.Recordset
    rs.Open SQL, db, adOpenStatic, adLockOptimistic
     
    If rs.RecordCount > 0 Then
        teller = 1
        With rs
            .MoveFirst
                Do
                    If InStr(1, .Fields("BESTANDSTITEL"), "cd:") = 0 Then
                        lengte = w1.Selection.StoryLength
                        lengte = lengte - 1
                        Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(lengte, lengte), 3, 2)
                        otable.Borders.Enable = 0
                        w1.ActiveDocument.Tables(teller).Columns(1).Width = 150
                        w1.ActiveDocument.Tables(teller).Columns(2).Width = 300
                        
                        w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.InsertAfter "Bestandstitel:"
                        w1.ActiveDocument.Tables(teller).Cell(1, 2).Range.InsertAfter .Fields("BESTANDSTITEL")
                        w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.InsertAfter "Omschrijving:"
                        w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.InsertAfter .Fields("TEXT")
                        w1.ActiveDocument.Tables(teller).Cell(3, 1).Range.InsertAfter "Bestandsnaam:"
                        w1.ActiveDocument.Tables(teller).Cell(3, 2).Range.InsertAfter .Fields("NAAM")
                        
                        w1.Selection.GoTo wdGoToLine, wdGoToRelative, lengte
                        w1.Selection.TypeParagraph
        
                        teller = teller + 1
                    End If
                    .MoveNext
                Loop Until .EOF
            .Close
        End With
    End If
    db.Close
    Form1.MousePointer = vbNormal
End If

If Option2.Value = True Then
    Form1.MousePointer = vbHourglass
    Set w1 = New Word.Application
    w1.Visible = True
    w1.WindowState = wdWindowStateMaximize
    w1.Documents.Add
    
    w1.Selection.Font.Size = 10
    w1.Selection.Font.Bold = True
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    w1.Selection.TypeText "Provincie Drenthe"
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = False
    w1.Selection.TypeText "Westerbrink 1"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Postbus 122"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "9400 AC Assen"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Contactpersoon"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Gerk van der Ploeg"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "0592 365463"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "G.vanderPloeg@drenthe.nl"
    w1.Selection.TypeParagraph
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = True
    w1.Selection.Font.Size = 12
    w1.Selection.TypeText "Niet Openbare GBI gegevens per " + Str(Date) + ":"
    w1.Selection.Font.Bold = False
    w1.Selection.Font.Size = 10
    w1.Selection.TypeParagraph
    
    If w1.Documents.Count < 1 Then
        MsgBox "Er zijn geen documenten open!"
        Exit Sub
    End If
    
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir
       
    SQL = "SELECT DATASET.DATACODE, DATASET.BESTANDSTITEL, B.TEXT, DATASET.NAAM, DATASET.FYSIEKE_LOCATIE FROM DATASET DATASET, MEMOTABEL B" + _
            " WHERE DATASET.OMSCHRIJVING = B.CODE AND DATASET.GEBRUIKSBEPERKING = ""niet openbaar"" ORDER BY DATASET.BESTANDSTITEL;"
            
    Set rs = New ADODB.Recordset
    rs.Open SQL, db, adOpenStatic, adLockOptimistic
     
    If rs.RecordCount > 0 Then
        teller = 1
        With rs
            .MoveFirst
                Do
                    If InStr(1, .Fields("BESTANDSTITEL"), "cd:") = 0 Then
                        lengte = w1.Selection.StoryLength
                        lengte = lengte - 1
                        Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(lengte, lengte), 4, 2)
                        otable.Borders.Enable = 0
                        w1.ActiveDocument.Tables(teller).Columns(1).Width = 150
                        w1.ActiveDocument.Tables(teller).Columns(2).Width = 300
                        
                        w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.InsertAfter "Bestandstitel:"
                        w1.ActiveDocument.Tables(teller).Cell(1, 2).Range.InsertAfter .Fields("BESTANDSTITEL")
                        w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.InsertAfter "Omschrijving:"
                        w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.InsertAfter .Fields("TEXT")
                        w1.ActiveDocument.Tables(teller).Cell(3, 1).Range.InsertAfter "Bestandsnaam:"
                        w1.ActiveDocument.Tables(teller).Cell(3, 2).Range.InsertAfter .Fields("NAAM")
                        w1.ActiveDocument.Tables(teller).Cell(4, 1).Range.InsertAfter "Fysieke locatie:"
                        w1.ActiveDocument.Tables(teller).Cell(4, 2).Range.InsertAfter .Fields("FYSIEKE_LOCATIE")
                        
                        w1.Selection.GoTo wdGoToLine, wdGoToRelative, lengte
                        w1.Selection.TypeParagraph
        
                        teller = teller + 1
                    End If
                    .MoveNext
                Loop Until .EOF
            .Close
        End With
    End If
    db.Close
    Form1.MousePointer = vbNormal
End If

If Option3.Value = True Then
    Form1.MousePointer = vbHourglass
    Set w1 = New Word.Application
    w1.Visible = True
    w1.WindowState = wdWindowStateMaximize
    w1.Documents.Add
    
    w1.Selection.Font.Size = 10
    w1.Selection.Font.Bold = True
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    w1.Selection.TypeText "Provincie Drenthe"
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = False
    w1.Selection.TypeText "Westerbrink 1"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Postbus 122"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "9400 AC Assen"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Contactpersoon"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Gerk van der Ploeg"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "0592 365463"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "G.vanderPloeg@drenthe.nl"
    w1.Selection.TypeParagraph
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = True
    w1.Selection.Font.Size = 12
    w1.Selection.TypeText "CD-Roms GBI gegevens per " + Str(Date) + ":"
    w1.Selection.Font.Bold = False
    w1.Selection.Font.Size = 10
    w1.Selection.TypeParagraph
    
    If w1.Documents.Count < 1 Then
        MsgBox "Er zijn geen documenten open!"
        Exit Sub
    End If
    
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir
       
    SQL = "SELECT DATASET.DATACODE, DATASET.BESTANDSTITEL, B.TEXT, DATASET.NAAM, DATASET.FYSIEKE_LOCATIE FROM DATASET DATASET, MEMOTABEL B" + _
            " WHERE DATASET.OMSCHRIJVING = B.CODE ORDER BY DATASET.BESTANDSTITEL;"
            
    Set rs = New ADODB.Recordset
    rs.Open SQL, db, adOpenStatic, adLockOptimistic
     
    If rs.RecordCount > 0 Then
        teller = 1
        With rs
            .MoveFirst
                Do
                    If InStr(1, .Fields("BESTANDSTITEL"), "cd:") > 0 Then
                        lengte = w1.Selection.StoryLength
                        lengte = lengte - 1
                        Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(lengte, lengte), 4, 2)
                        otable.Borders.Enable = 0
                        w1.ActiveDocument.Tables(teller).Columns(1).Width = 150
                        w1.ActiveDocument.Tables(teller).Columns(2).Width = 300
                        
                        w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.InsertAfter "Bestandstitel:"
                        w1.ActiveDocument.Tables(teller).Cell(1, 2).Range.InsertAfter .Fields("BESTANDSTITEL")
                        w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.InsertAfter "Omschrijving:"
                        w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.InsertAfter .Fields("TEXT")
                        w1.ActiveDocument.Tables(teller).Cell(3, 1).Range.InsertAfter "Bestandsnaam:"
                        w1.ActiveDocument.Tables(teller).Cell(3, 2).Range.InsertAfter .Fields("NAAM")
                        w1.ActiveDocument.Tables(teller).Cell(4, 1).Range.InsertAfter "Fysieke locatie:"
                        w1.ActiveDocument.Tables(teller).Cell(4, 2).Range.InsertAfter .Fields("FYSIEKE_LOCATIE")
                        
                        w1.Selection.GoTo wdGoToLine, wdGoToRelative, lengte
                        w1.Selection.TypeParagraph
        
                        teller = teller + 1
                    End If
                    .MoveNext
                Loop Until .EOF
            .Close
        End With
    End If
    db.Close
    Form1.MousePointer = vbNormal
End If

If Option4.Value = True Then
    Form1.MousePointer = vbHourglass
    Set w1 = New Word.Application
    w1.Visible = True
    w1.WindowState = wdWindowStateMaximize
    w1.Documents.Add
    
    w1.Selection.Font.Size = 10
    w1.Selection.Font.Bold = True
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    w1.Selection.TypeText "Provincie Drenthe"
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = False
    w1.Selection.TypeText "Westerbrink 1"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Postbus 122"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "9400 AC Assen"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Contactpersoon"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Gerk van der Ploeg"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "0592 365463"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "G.vanderPloeg@drenthe.nl"
    w1.Selection.TypeParagraph
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = True
    w1.Selection.Font.Size = 12
    w1.Selection.TypeText "Openbare " + Combo1.Text + " GBI gegevens per " + Str(Date) + ":"
    w1.Selection.Font.Bold = False
    w1.Selection.Font.Size = 10
    w1.Selection.TypeParagraph
    
    If w1.Documents.Count < 1 Then
        MsgBox "Er zijn geen documenten open!"
        Exit Sub
    End If
    
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir
    
    nw = Combo1.Text
    SQL = "SELECT DISTINCT a.BESTANDSTITEL FROM DATASET a, TREFTEXT b, TREFCODE c WHERE b.TREFWOORD = """ + nw + """ AND " + _
        "a.DATACODE = c.DATACODE AND b.TREFCODE = c.TREFCODE ORDER BY a.BESTANDSTITEL;"
        
    Set rs = New ADODB.Recordset
    rs.Open SQL, db, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        teller = 1
        With rs
            .MoveFirst
                Do
                    bestandtitel = .Fields("BESTANDSTITEL")
                    SQL2 = "SELECT DATASET.DATACODE, DATASET.BESTANDSTITEL, B.TEXT, DATASET.NAAM FROM DATASET DATASET, MEMOTABEL B" + _
                        " WHERE DATASET.OMSCHRIJVING = B.CODE AND DATASET.BESTANDSTITEL = """ + bestandtitel + """ AND DATASET.GEBRUIKSBEPERKING = ""openbaar"" ORDER BY DATASET.BESTANDSTITEL;"
                    
                    Set rs2 = New ADODB.Recordset
                    rs2.Open SQL2, db, adOpenStatic, adLockOptimistic
                    
                    If rs2.RecordCount > 0 Then
                        With rs2
                            .MoveFirst
                                Do
                                    If InStr(1, .Fields("BESTANDSTITEL"), "cd:") = 0 Then
                                        lengte = w1.Selection.StoryLength
                                        lengte = lengte - 1
                                        Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(lengte, lengte), 3, 2)
                                        otable.Borders.Enable = 0
                                        w1.ActiveDocument.Tables(teller).Columns(1).Width = 150
                                        w1.ActiveDocument.Tables(teller).Columns(2).Width = 300
                        
                                        w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.InsertAfter "Bestandstitel:"
                                        w1.ActiveDocument.Tables(teller).Cell(1, 2).Range.InsertAfter .Fields("BESTANDSTITEL")
                                        w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.InsertAfter "Omschrijving:"
                                        w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.InsertAfter .Fields("TEXT")
                                        w1.ActiveDocument.Tables(teller).Cell(3, 1).Range.InsertAfter "Bestandsnaam:"
                                        w1.ActiveDocument.Tables(teller).Cell(3, 2).Range.InsertAfter .Fields("NAAM")
                        
                                        w1.Selection.GoTo wdGoToLine, wdGoToRelative, lengte
                                        w1.Selection.TypeParagraph
        
                                        teller = teller + 1
                                    End If
                                .MoveNext
                                Loop Until .EOF
                            .Close
                        End With
                    End If
                    .MoveNext
                Loop Until .EOF
            .Close
        End With
    End If
    db.Close
    Form1.MousePointer = vbNormal
End If

If Option5.Value = True Then
    Form1.MousePointer = vbHourglass
    Set w1 = New Word.Application
    w1.Visible = True
    w1.WindowState = wdWindowStateMaximize
    w1.Documents.Add
    
    w1.Selection.Font.Size = 10
    w1.Selection.Font.Bold = True
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    w1.Selection.TypeText "Provincie Drenthe"
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = False
    w1.Selection.TypeText "Westerbrink 1"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Postbus 122"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "9400 AC Assen"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Contactpersoon"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "Gerk van der Ploeg"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "0592 365463"
    w1.Selection.TypeParagraph
    w1.Selection.TypeText "G.vanderPloeg@drenthe.nl"
    w1.Selection.TypeParagraph
    w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    w1.Selection.TypeParagraph
    w1.Selection.Font.Bold = True
    w1.Selection.Font.Size = 12
    w1.Selection.TypeText "Openbare " + Combo1.Text + " GBI Metagegevens per " + Str(Date) + ":"
    w1.Selection.Font.Bold = False
    w1.Selection.Font.Size = 10
    w1.Selection.TypeParagraph
    
    If w1.Documents.Count < 1 Then
        MsgBox "Er zijn geen documenten open!"
        Exit Sub
    End If
    
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & Form1.dbdir
    
    nw = Combo1.Text
    SQL = "SELECT DISTINCT a.DATACODE FROM DATASET a, TREFTEXT b, TREFCODE c WHERE b.TREFWOORD = """ + nw + """ AND " + _
        "a.DATACODE = c.DATACODE AND b.TREFCODE = c.TREFCODE ORDER BY a.DATACODE;"
        
    Set rs = New ADODB.Recordset
    rs.Open SQL, db, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        teller = 1
        With rs
            .MoveFirst
                Do
                    datacode = rs.Fields("DATACODE")
                    SQL2 = "SELECT a.BESTANDSTITEL, a.NAAM, a.GEBRUIKSBEPERKING, a.STATUS, b.TEXT FROM " + _
                        "DATASET a, MEMOTABEL b WHERE a.OMSCHRIJVING = b.CODE AND a.DATACODE = " + datacode + ";"
                    
                    Set rs2 = New ADODB.Recordset
                    rs2.Open SQL2, db, adOpenStatic, adLockOptimistic
                    
                    lengte = w1.Selection.StoryLength
                    lengte = lengte - 1

                    Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(lengte, lengte), 6, 2)
                    otable.Borders.Enable = 0
                    w1.ActiveDocument.Tables(teller).Columns(1).Width = 150
                    w1.ActiveDocument.Tables(teller).Columns(2).Width = 300

                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Merge w1.ActiveDocument.Tables(teller).Cell(1, 2)
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Bold = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Size = 10
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Italic = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.InsertAfter "Algemene gegevens:"

                    w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.InsertAfter "Bestandstitel:"
                    w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.InsertAfter rs2.Fields("BESTANDSTITEL")
                    w1.ActiveDocument.Tables(teller).Cell(3, 1).Range.InsertAfter "Omschrijving:"
                    w1.ActiveDocument.Tables(teller).Cell(3, 2).Range.InsertAfter rs2.Fields("TEXT")
                    w1.ActiveDocument.Tables(teller).Cell(4, 1).Range.InsertAfter "Bestandsnaam:"
                    w1.ActiveDocument.Tables(teller).Cell(4, 2).Range.InsertAfter rs2.Fields("NAAM")
                    w1.ActiveDocument.Tables(teller).Cell(5, 1).Range.InsertAfter "Gebruiksbeperking:"
                    w1.ActiveDocument.Tables(teller).Cell(5, 2).Range.InsertAfter rs2.Fields("GEBRUIKSBEPERKING")
                    w1.ActiveDocument.Tables(teller).Cell(6, 1).Range.InsertAfter "Status bestand:"
                    w1.ActiveDocument.Tables(teller).Cell(6, 2).Range.InsertAfter rs2.Fields("STATUS")
                    
                    w1.Selection.GoTo wdGoToLine, wdGoToRelative, lengte
                    w1.Selection.TypeParagraph
                    teller = teller + 1
                    
                    SQL2 = "SELECT a.NAAM, a.FYSIEKE_LOCATIE, a.CONTACTPERSOON, a.COPYRIGHT FROM " + _
                        "DATASET a WHERE a.DATACODE = " + datacode + ";"
                     
                    Set rs2 = New ADODB.Recordset
                    rs2.Open SQL2, db, adOpenStatic, adLockOptimistic
                    
                    lengte = w1.Selection.StoryLength
                    lengte = lengte - 1

                    Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(lengte, lengte), 5, 2)
                    otable.Borders.Enable = 0
                    w1.ActiveDocument.Tables(teller).Columns(1).Width = 150
                    w1.ActiveDocument.Tables(teller).Columns(2).Width = 300

                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Merge w1.ActiveDocument.Tables(teller).Cell(1, 2)
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Bold = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Size = 10
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Italic = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.InsertAfter "Toegang gegevens:"
                    
                    w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.InsertAfter "Bestandsnaam:"
                    w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.InsertAfter rs2.Fields("NAAM")
                    w1.ActiveDocument.Tables(teller).Cell(3, 1).Range.InsertAfter "Padnaam bestand:"
                    w1.ActiveDocument.Tables(teller).Cell(3, 2).Range.InsertAfter rs2.Fields("FYSIEKE_LOCATIE")
                    w1.ActiveDocument.Tables(teller).Cell(4, 1).Range.InsertAfter "Contactpersoon inhoud:"
                    w1.ActiveDocument.Tables(teller).Cell(4, 2).Range.InsertAfter rs2.Fields("CONTACTPERSOON")
                    w1.ActiveDocument.Tables(teller).Cell(5, 1).Range.InsertAfter "Copyright:"
                    w1.ActiveDocument.Tables(teller).Cell(5, 2).Range.InsertAfter rs2.Fields("COPYRIGHT")
                                         
                    w1.Selection.GoTo wdGoToLine, wdGoToRelative, lengte
                    w1.Selection.TypeParagraph
                    teller = teller + 1
                    
                    SQL2 = "SELECT a.BRONDATUM, a.OPBOUWDATUM, a.BIJHOUDING, a.CONTACTPERSOON, a.COPYRIGHT, " + _
                        "a.CONTACT_LEVERANCIER FROM DATASET a WHERE a.DATACODE = " + datacode + ";"
                    
                    Set rs2 = New ADODB.Recordset
                    rs2.Open SQL2, db, adOpenStatic, adLockOptimistic
                    
                    lengte = w1.Selection.StoryLength
                    lengte = lengte - 1

                    Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(lengte, lengte), 6, 2)
                    otable.Borders.Enable = 0
                    w1.ActiveDocument.Tables(teller).Columns(1).Width = 150
                    w1.ActiveDocument.Tables(teller).Columns(2).Width = 300

                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Merge w1.ActiveDocument.Tables(teller).Cell(1, 2)
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Bold = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Size = 10
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Italic = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.InsertAfter "Inhoudelijke gegevens:"
                    
                    w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.InsertAfter "Brondatum:"
                    w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.InsertAfter rs2.Fields("BRONDATUM")
                    w1.ActiveDocument.Tables(teller).Cell(3, 1).Range.InsertAfter "Opbouwdatum:"
                    w1.ActiveDocument.Tables(teller).Cell(3, 2).Range.InsertAfter rs2.Fields("OPBOUWDATUM")
                    w1.ActiveDocument.Tables(teller).Cell(4, 1).Range.InsertAfter "Frequentie bijhouding:"
                    w1.ActiveDocument.Tables(teller).Cell(4, 2).Range.InsertAfter rs2.Fields("BIJHOUDING")
                    w1.ActiveDocument.Tables(teller).Cell(5, 1).Range.InsertAfter "Leverancier:"
                    w1.ActiveDocument.Tables(teller).Cell(5, 2).Range.InsertAfter rs2.Fields("COPYRIGHT")
                    w1.ActiveDocument.Tables(teller).Cell(6, 1).Range.InsertAfter "Contactpersoon leverancier:"
                    w1.ActiveDocument.Tables(teller).Cell(6, 2).Range.InsertAfter rs2.Fields("CONTACT_LEVERANCIER")
                                         
                    w1.Selection.GoTo wdGoToLine, wdGoToRelative, lengte
                    w1.Selection.TypeParagraph
                    teller = teller + 1
                    
                    SQL2 = "SELECT a.OPBOUWMETHODE, a.DEELGEBIED, a.GEOMETRIE, a.LEGENDA_FILE, a.POS_NAUWKEURIGHEID, " + _
                        "a.SCHAAL, a.MAXSCHAAL, a.STD_ITEM FROM GEOGRAFISCH a WHERE a.DATACODE = " + datacode + ";"
                    
                    Set rs2 = New ADODB.Recordset
                    rs2.Open SQL2, db, adOpenStatic, adLockOptimistic
                    
                    lengte = w1.Selection.StoryLength
                    lengte = lengte - 1

                    Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(lengte, lengte), 9, 2)
                    otable.Borders.Enable = 0
                    w1.ActiveDocument.Tables(teller).Columns(1).Width = 150
                    w1.ActiveDocument.Tables(teller).Columns(2).Width = 300

                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Merge w1.ActiveDocument.Tables(teller).Cell(1, 2)
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Bold = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Size = 10
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Italic = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.InsertAfter "Specifieke gegevens:"
                    
                    w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.InsertAfter "Opbouwmethode:"
                    w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.InsertAfter rs2.Fields("OPBOUWMETHODE")
                    w1.ActiveDocument.Tables(teller).Cell(3, 1).Range.InsertAfter "Gebied:"
                    w1.ActiveDocument.Tables(teller).Cell(3, 2).Range.InsertAfter rs2.Fields("DEELGEBIED")
                    w1.ActiveDocument.Tables(teller).Cell(4, 1).Range.InsertAfter "Geometrie:"
                    w1.ActiveDocument.Tables(teller).Cell(4, 2).Range.InsertAfter rs2.Fields("GEOMETRIE")
                    w1.ActiveDocument.Tables(teller).Cell(5, 1).Range.InsertAfter "Legenda bestand:"
                    w1.ActiveDocument.Tables(teller).Cell(5, 2).Range.InsertAfter rs2.Fields("LEGENDA_FILE")
                    w1.ActiveDocument.Tables(teller).Cell(6, 1).Range.InsertAfter "Positionele nauwkeurigheid:"
                    w1.ActiveDocument.Tables(teller).Cell(6, 2).Range.InsertAfter rs2.Fields("POS_NAUWKEURIGHEID")
                    w1.ActiveDocument.Tables(teller).Cell(7, 1).Range.InsertAfter "Schaal:"
                    w1.ActiveDocument.Tables(teller).Cell(7, 2).Range.InsertAfter rs2.Fields("SCHAAL")
                    w1.ActiveDocument.Tables(teller).Cell(8, 1).Range.InsertAfter "Tekenschaal:"
                    w1.ActiveDocument.Tables(teller).Cell(8, 2).Range.InsertAfter rs2.Fields("MAXSCHAAL")
                    w1.ActiveDocument.Tables(teller).Cell(9, 1).Range.InsertAfter "Standaard item:"
                    w1.ActiveDocument.Tables(teller).Cell(9, 2).Range.InsertAfter rs2.Fields("STD_ITEM")
                                                  
                    w1.Selection.GoTo wdGoToLine, wdGoToRelative, lengte
                    w1.Selection.TypeParagraph
                    teller = teller + 1
                    
                    SQL2 = "SELECT a.ITEMNAAM, a.ITEMDEFINITIE, a.EENHEID, b.TEXT FROM ITEMS a, MEMOTABEL b " + _
                        "WHERE a.DOMEIN = b.CODE AND a.DATACODE = " + datacode + " ORDER BY a.VOLGNR;"
                    
                    Set rs2 = New ADODB.Recordset
                    rs2.Open SQL2, db, adOpenStatic, adLockOptimistic
                    
                    lengte = w1.Selection.StoryLength
                    lengte = lengte - 1
                    
                    aantalrij = rs2.RecordCount
                    
                    Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(teller, teller), (2 + aantalrij), 4)
                    otable.Borders.Enable = 0
                    w1.ActiveDocument.Tables(teller).Columns(1).Width = 70
                    w1.ActiveDocument.Tables(teller).Columns(2).Width = 125
                    w1.ActiveDocument.Tables(teller).Columns(3).Width = 50
                    w1.ActiveDocument.Tables(teller).Columns(4).Width = 210

                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Merge w1.ActiveDocument.Tables(6).Cell(1, 4)
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Bold = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Size = 10
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.Font.Italic = True
                    w1.ActiveDocument.Tables(teller).Cell(1, 1).Range.InsertAfter "Item gegevens:"

                    w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    w1.ActiveDocument.Tables(teller).Cell(2, 1).Range.InsertAfter "Itemnaam"
                    w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    w1.ActiveDocument.Tables(teller).Cell(2, 2).Range.InsertAfter "Itemdefinitie"
                    w1.ActiveDocument.Tables(teller).Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    w1.ActiveDocument.Tables(teller).Cell(2, 3).Range.InsertAfter "Eenheid"
                    w1.ActiveDocument.Tables(teller).Cell(2, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    w1.ActiveDocument.Tables(teller).Cell(2, 4).Range.InsertAfter "Mogelijke waarden"
                    
                    i = 2
                    If rs2.RecordCount > 0 Then
                        With rs2
                            .MoveFirst
                                Do
                                    w1.ActiveDocument.Tables(teller).Cell((i + 1), 1).Range.InsertAfter rs2.Fields("ITEMNAAM")
                                    w1.ActiveDocument.Tables(teller).Cell((i + 1), 2).Range.InsertAfter rs2.Fields("ITEMDEFINITIE")
                                    w1.ActiveDocument.Tables(teller).Cell((i + 1), 3).Range.InsertAfter rs2.Fields("EENHEID")
                                    w1.ActiveDocument.Tables(teller).Cell((i + 1), 4).Range.InsertAfter rs2.Fields("TEXT")
                                    .MoveNext
                                Loop Until .EOF
                        End With
                    End If
                    
                    w1.Selection.GoTo wdGoToLine, wdGoToRelative, lengte
                    w1.Selection.TypeParagraph
                    teller = teller + 1
                    
                .MoveNext
                Loop Until .EOF
            .Close
        End With
    End If
    db.Close
    Form1.MousePointer = vbNormal
End If

End Sub
