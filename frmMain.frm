VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Access Database - For Biginners"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Add New Entry"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Settings"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Previous Entry"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Querry For Members"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   1680
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000006&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "        "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2790
      ItemData        =   "frmMain.frx":0000
      Left            =   360
      List            =   "frmMain.frx":0002
      TabIndex        =   8
      Top             =   3600
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next Entry"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "Other Data, Try For Your Self"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "Other Data, Try For Your Self"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "Other Data, Try For Your Self"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2895
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   0
      Y2              =   7560
   End
   Begin VB.Label Label6 
      Caption         =   "For Biginners, by XEN AXON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Introduction To Access Database Programming,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   6720
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   5520
      X2              =   5520
      Y1              =   0
      Y2              =   7560
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "E-Mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "UserName:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
'/////////////////////////I hope you find this useful///////////////////////
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
'////////////////////Made By XEN AXON///////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
'/////////////////////////I am curently working on a text based game////////
'/////////////////////for massive multiplayer online////////////////////////
'////////////////////////////////////////If any of you Want to take part////
'///////////////////////////////////////////////////////////////////////////
'/////////////////////////please contact me at nukerarmada@yahoo.com////////
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
'////////////////IF YOU FOUND THIS CODE USEFULL PLEASE VOTE FOR IT//////////
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
Option Explicit
'Dim as Public, which allows these variables to be used by all forms and modules within
'the VB project.
Public db As Database 'db(the database, this is the .mdb file ! the Microsoft Access file)
Public rstInfo As Recordset 'recordset(the recordset is the Table in The Access Database that holds the data)
Private Sub Command1_Click()
With rstInfo 'rstInfo, the table that holds our data.
    If rstInfo.RecordCount > 0 Then 'if the .mdb file is or not empty.
    On Error GoTo moveErr   'on error it will tell the recordset to go to entry nr. 1 to avoid the error
    .MoveNext 'this will go to the next record in the table
    On Error GoTo 0
        If rstInfo.EOF Then
            MsgBox "End of file.", vbOKOnly, "  Error!" 'this will tell the program to stop going farther if the querry has reached the final entry
            .MoveLast
            Exit Sub
        Else
            .Edit
                Text1.Text = !UserName 'this sets the text1 to the new entry's username
                Text2.Text = !Email    'same for the e-mail==>>>>>>>>>>>>^^^^
                Text3.Text = !Password 'same for the password^^
        End If
    Else
        MsgBox "No Records In The DataBase." 'this will print a messagebox on the screen, saying that the database is empty, if the database is empty. :D.
    End If
End With
moveErr:
If Err.Number = 3021 Then rstInfo.MoveFirst 'this error ocures if the user tryes to go byond the final entry of the table value.
End Sub
Private Sub Command2_Click()
'this button ads all the username values to the listbox
Timer1.Enabled = True 'enables timer1 that will put all the entry's in the listbox
List1.Clear 'it clears the list1 in case it allready had values in it
rstInfo.MoveFirst 'this moves to the first value of the database table to avoid an error. :P
End Sub
Private Sub Command3_Click()
'this command is basicly the same as command1, only difference is that this goes to the previous value instead of going to the next value :P. =))
With rstInfo
    If rstInfo.RecordCount > 0 Then
        On Error GoTo moveErr
            .MovePrevious
        If rstInfo.BOF Then
            MsgBox "End of file.", vbOKOnly, "  Error!"
            .MoveFirst
            Exit Sub
        Else
            .Edit
                Text1.Text = !UserName
                Text2.Text = !Email
                Text3.Text = !Password
        End If
    Else
        MsgBox "No Records In The DataBase."
    End If
End With
moveErr:
If Err.Number = 3021 Then rstInfo.MoveLast
'same stuff as the command1....take a look at that one and you'll understand this also.
End Sub
Private Sub Command4_Click()
'this saves the modifications for the currently selected database entry :)
With rstInfo
    If rstInfo.RecordCount > 0 Then 'if the table has values....
            .Edit 'edit so we can change the entry's.
                !UserName = Text1.Text
                !Email = Text2.Text
                !Password = Text3.Text
                .Update 'won't work with out update :P
                Call Command2_Click 'i call the command2_click, that's the button that querry'es the server,
    Else
        MsgBox "No Records In The DataBase." 'If rstInfo.RecordCount > 0 Then it displays a messagebox that tells you the table is empty
    End If
End With
End Sub
Private Sub Command5_Click()
On Error GoTo UserNameEx
    With rstInfo
        .AddNew
        If Trim(Text1.Text) <> "" Then
             !UserName = Text1.Text
        Else
             MsgBox "You Must Enter A Name!", vbCritical, "Error Adding New Entry"
             Exit Sub
        End If
        If Trim(Text2.Text) <> "" Then
             !Email = Text2.Text
        Else
             !Email = "No Entry"
        End If
        If Trim(Text3.Text) <> "" Then
             !Password = Text3.Text
        Else
            MsgBox "You Must Enter A Password!", vbCritical, "Error Adding New Entry"
            Exit Sub
        End If
        .Update
    End With
Call Form_Load
Call Command2_Click
UserNameEx:
If Err.Number = "3032" Then MsgBox "That UserName Is Already Used, And The Database Is Set To Deny Adding A Duplicate Value(UserName or E-Mail)"
End Sub
Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\" & "Accounts.mdb") 'this will open the database so we can manipulate it. =)))) LOL
With db

    Set rstInfo = .OpenRecordset("UserDat") 'this will open the table that we need. in our case it's UserDat, but you can change this acording to your .mdb(acces file)
    Label1.Caption = rstInfo.RecordCount & " records" 'this will set label1 to the curent number of entry'es (records) in the table.
    If rstInfo.RecordCount = 0 Then Exit Sub 'if it's empty it stops.
    With rstInfo
        .Edit 'set's our text boxes to the entry's from the table
        Text1.Text = !UserName
        Text2.Text = !Email
        Text3.Text = !Password
    End With
End With
End Sub
Private Sub List1_Click()
'this will get the data for the selected UserName in the ListBox
With rstInfo
    If rstInfo.RecordCount > 0 Then 'again if checks if the table has entry's or not...
    Dim itemNum As String 'dims itemNum variable as the list1.listindex(that's the index number of the selected UserName)
    itemNum = List1.ListIndex
    .MoveFirst 'you first set to the first line to avoid an error because if you rstInfo.move 3 it will move 3 more from the curent entry
                'and will not get the real entry that we need and also will return an error if it goes beyond the last list item
    .Move itemNum
            .Edit
                Text1.Text = !UserName
                Text2.Text = !Email
                Text3.Text = !Password
        End If
End With
End Sub
Private Sub Timer1_Timer()
'this timer activates when the user presses the querry for members button
'it gets all of the UserName entry's and adds them to list1
With rstInfo
    If rstInfo.RecordCount > 0 Then 'check's if the table has data.
            If rstInfo.EOF Then 'if it has reached the last entry it stops and display's the following message:
            MsgBox "Finished Querrying For Members", vbOKOnly, "  Operation Done!"
            Timer1.Enabled = False ' it stops the timer from executing. =))
            Exit Sub
            Else
            .Edit
                Text1.Text = !UserName 'changes our textboxes to the values from the table
                Text2.Text = !Email
                Text3.Text = !Password
                List1.AddItem !UserName
        End If
        .MoveNext 'it moves to the next entry(prepares for the next execution of the timer)
    End If
End With
End Sub
'////////////////////////////////////////////////////////////
'//////////Whell That's All !! I Hope It Was Useful//////////
'//////////If You Found This Usefull Please Vote For IT//////
'//////////I Would Apreciate That !! Thanks Alot////////////
'////////////////////////////////////////////////////////////
