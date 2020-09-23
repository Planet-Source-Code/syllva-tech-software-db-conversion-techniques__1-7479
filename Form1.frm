VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database conversion demos by Millennium Software"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Export to a tab-delimited text file"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   6855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fill Listbox with data"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert to HTML"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   6960
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   6855
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1260
      Width           =   3015
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6960
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#############################################
'In order for this to work correctly, you MUST
'make a reference to the Microsoft DAO Library
'
'Click on Project --> References
'Look for Microsoft DAO 3.0 Library (or higher)
'#############################################

'You will ALWAYS need to declare
'these variables when using MS DAO
Dim db As Database
Dim rs As Recordset


Private Sub Command1_Click()

'First you need to open the database
    Set db = OpenDatabase(App.Path & "\booksales.mdb")

'Now open the table you want to use...
    Set rs = db.OpenRecordset("Titles")

'Have VB create a file on the hard disk
'to write to
Open App.Path & "\Book_Titles.htm" For Output As #1

'Now print the initial HTML document structure
    Print #1, "<HTML>"
    Print #1, "<HEAD>"
    Print #1, "<TITLE>"
    Print #1, "Just Testing..."
    Print #1, "</TITLE>"
    Print #1, "</HEAD>"
    Print #1, "<BODY>"

'First, we are going to make a table, showing a border
    Print #1, "<TABLE BORDER>"

'Here the info from the database is written
'into the document's table
    Do Until rs.EOF
        Print #1, "<TR><TD><B>"; rs!Title; "</B></TD><TD BGCOLOR=""#c0cFc0"">"; "$"; rs!Price; ".00</TD></TR>"

'Go to the next record
        rs.MoveNext

'Start over with the next record
    Loop

'Print the closing table tag
    Print #1, "</TABLE>"

'Print the closing HTML Document tags
    Print #1, "</BODY>"
    Print #1, "</HEAD>"

'The file is finished,
'tell VB to stop writing
'this file
    Close #1

'Let the user know we are done
Label1.Caption = "Finished!"
End Sub


Private Sub Command2_Click()

'Open the database
    Set db = OpenDatabase(App.Path & "\booksales.mdb")
    
'Open the table we need and open it as read only
'and forward scrolling only
    Set rs = db.OpenRecordset("Titles", dbOpenForwardOnly, dbReadOnly)
    
'Now lets populate the listbox

    Do Until rs.EOF
        List1.AddItem rs!Title
        rs.MoveNext
    Loop

'Let the user know we are done
Label2.Caption = "Finished!"

End Sub

Private Sub Command3_Click()

'Open the database
    Set db = OpenDatabase(App.Path & "\booksales.mdb")

'Now open the table you want to use...
    Set rs = db.OpenRecordset("Titles")

'Have VB create a file on the hard disk
'to write to
Open App.Path & "\Book Info.txt" For Output As #1

'Let's start writing
Do Until rs.EOF
    Print #1, rs!Title; Chr(9); _
                rs!ISBN; Chr(9); _
                rs!Pages; Chr(9); _
                rs!Price
                rs.MoveNext
'Keep going until end of table
Loop

'We are done
Close #1
End Sub
