VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmUtility_ZipCode 
   Caption         =   " ZipCode Search Utility"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   Icon            =   "frmUtility_ZipCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6525
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   5400
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8440
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search Results"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6255
      Begin MSComctlLib.ListView ListView1 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "City"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "State"
            Object.Width           =   1032
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Zipcode"
            Object.Width           =   1667
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Areacode"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "County"
            Object.Width           =   2390
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search "
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CheckBox chkExact 
         Caption         =   "Results do not have to be an exact match."
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   1605
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Print"
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "Zip"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "County"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "Area"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "State"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "City"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   1080
         Picture         =   "frmUtility_ZipCode.frx":030A
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Search Criteria"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   360
         Width           =   3015
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmUtility_ZipCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Written by L. "Mike" Trivette
' Please send me comments at mtrivette@yahoo.com
' Dont forget to give credit where credit is due.
'
' Revised 4/10/03
'
Option Explicit

Dim category As String    ' Public varibale that describes the chosen category to search from
Dim strDatabase As String ' Public varibale that describes the location of the database
Dim Response As String ' Public variable that describes any input from the user

Private Sub printfunc()
    ' just a very quick and simple print routine.
    ' this will definately need some work.
    Dim i As Long
    If ListView1.ListItems.Count = 0 Then Exit Sub
    CommonDialog1.Copies = 1
    CommonDialog1.ShowPrinter
    Printer.Copies = CommonDialog1.Copies
    Printer.Print " "
    Printer.Print Tab(95); Now()
    Printer.Print " "
    Printer.Print " "
    For i = 1 To ListView1.ListItems.Count
        Printer.Print Tab(5); ListView1.ListItems(i).Text;
        Printer.Print Tab(25); ListView1.ListItems(i).SubItems(1);
        Printer.Print Tab(35); ListView1.ListItems(i).SubItems(2);
        Printer.Print Tab(45); ListView1.ListItems(i).SubItems(3);
        Printer.Print Tab(55); ListView1.ListItems(i).SubItems(4)
    Next i
    Printer.EndDoc
End Sub

Private Sub load_zips()
    ' This sub searches and then loads the found data into the listview control.
    '
    On Error Resume Next
    
    Dim ii As Long ' Set Variable  to hold loop integer
    Dim dbs As Database
    Dim rst As Recordset
    Dim itmx As ListItem
    
    txtSearch.Text = LTrim(txtSearch.Text) ' Clean up search criteria from leading spaces
    txtSearch.Text = RTrim(txtSearch.Text) ' Clean up search criteria from trailing spaces
    
    If RTrim(txtSearch.Text) = "" Then ' Warn user before pulling all the records.
        Response = MsgBox("By leaving the search field blank this program will pull up everything in the database." & vbCrLf & "This will also take a long time, depending on your system." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbYesNo + vbInformation, "Warning!")
        If Response = 7 Then Exit Sub
        chkExact.Value = 1
    End If
    
    Me.MousePointer = 11 ' Show mouse as hourglass
    StatusBar1.Panels(1).Text = "Searching..." ' Update statusbar
        
    Set dbs = OpenDatabase(strDatabase) ' Open database
    
    If chkExact.Value = 0 Then ' Determine to use exact searches or not.
        Set rst = dbs.OpenRecordset("SELECT * FROM zipcodes where " & (category) & " = '" & (txtSearch.Text) & "';") ' Search for data
    Else
        Set rst = dbs.OpenRecordset("SELECT * FROM zipcodes where " & (category) & " LIKE '" & (txtSearch.Text) & "*';") ' Search for data
    End If
    
    ListView1.ListItems.Clear ' Clear out the listview for the new data
        
    If rst.RecordCount > 0 Then ' If recordset has anything in it then dum it into our listview.
        rst.MoveLast ' Populate the recordset - (Very important or recordset.count will only show 1)
        rst.MoveFirst ' This is part of the population
        
        For ii = 1 To rst.RecordCount ' Main loop starts here
        Set itmx = ListView1.ListItems.Add(, , "" & rst.Fields("city"))
                   itmx.SubItems(1) = "" & rst.Fields("state")
                   itmx.SubItems(2) = "" & rst.Fields("zip")
                   itmx.SubItems(3) = "" & rst.Fields("area")
                   itmx.SubItems(4) = "" & rst.Fields("county")
        rst.MoveNext
        Next ii ' Main loop ends here.
    Else
        Me.MousePointer = 0 ' Do want the hourglass
        Response = MsgBox("Sorry but no search results containing " & txtSearch.Text & " were found.", vbOKOnly, "No results!")
    End If
    
    Me.MousePointer = 0 ' Do want the hourglass
    StatusBar1.Panels(1).Text = rst.RecordCount & " results." ' Show status
    StatusBar1.Panels(2).Text = ""
    
    rst.Close ' close recordset.
    dbs.Close ' close database.
End Sub

Private Sub Command1_Click()
    If category <> "" Then load_zips
End Sub

Private Sub Command2_Click()
    If ListView1.ListItems.Count > 0 Then printfunc
End Sub

Private Sub Form_Load()
    ' Variable describes the location of the zipcodes.mdb database
    ' You may have to adjust this file location
    strDatabase = App.Path & "\zipcodes.mdb"
    category = "City"
    Me.Caption = Me.Caption & " - Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    ' Keeps user form resing form
    Me.Width = 6630
    Me.Height = 6305
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuPrint_Click()
    If ListView1.ListItems.Count > 0 Then printfunc
End Sub

Private Sub Opt_Category_Click(Index As Integer)
    ' Change the category variable for the entire program to know
    category = Opt_Category(Index).Caption
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    ' If user hits enter key then search will be set off.
    If KeyAscii = 13 Then Command1_Click
End Sub
