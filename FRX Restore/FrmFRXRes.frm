VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmFRXRes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FRX Restore Program."
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "FrmFRXRes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   315
      Left            =   5400
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   1095
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Pic"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.ListBox List1 
      Columns         =   40
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\Microsoft Visual Studio\VB98\FRX Restore\FrmFRXres.frx"
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Data"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Stat 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Status"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "File Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Click on the picture you wish to view."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5175
   End
End
Attribute VB_Name = "FrmFRXRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''' Http://Home.dal.net/HoeBot/ ''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Welcome to yet agean anouther open source program from HoeBot.'
'Im sure you know all the code in here is CopyRighted (c) to me'
'And if you brake that copyright, im one hell of a whore with  '
'writing a thousand e-mails to all the abuse@'s i can find, so '
'place nicely.                                                 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This is a rather simple, yet powerful program. I made it for  '
'no reason at all other then to kill time ^_^                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function CountFH(Text As String)
''simply counts the number of times the
'' "new picture file" string shows up in the frx
times1 = 1
FoundNum = 0

Do
    IsIn = InStr(times1, Text, FHeader(), vbBinaryCompare)
    If (IsIn > 0) Then
        FoundNum = FoundNum + 1
        times1 = IsIn + 2
    ElseIf (IsIn = 0) Then
        GoTo end_:
    End If
Loop
end_:
CountFH = FoundNum
End Function
Function FHeader()
''returns what that "new pic file" in the
''.frx is.
FHeader = Chr(0) & Chr(108) & Chr(116) & Chr(0) & Chr(0)


''There seems to be multiple strings for differnt file types.
''There was no need for me to figure them all out
''so im just using a common string found in all of them.
''I left these in here incase you want to play with them.

''If (LCase(fType) = "ico") Then FHeader = Chr(198) & Chr(8) & Chr(0) & Chr(0) & Chr(108) & Chr(116) & Chr(0) & Chr(0) & Chr(190) & Chr(8) & Chr(0) & Chr(0)
''If (LCase(fType) = "bmp") Then FHeader = Chr(168) & Chr(0) & Chr(3) & Chr(0) & Chr(108) & Chr(116) & Chr(0) & Chr(0) & Chr(160) & Chr(0) & Chr(3) & Chr(0)
End Function
Private Sub Command1_Click()
''This is some rather simple code.
''all it does is count how many pictures are
''in the given .frx file then adds that many
''numbers a list box.
Dim buffer As String
Stat = "Loading Data"
    FN = FreeFile
    out$ = ""
    Open Text1.Text For Binary As #FN
        buffer = String(LOF(FN), 0)
        DoEvents
        Get #FN, 1, buffer$
    Close #FN
    z = CountFH(buffer$)
    For t = 1 To z
    List1.AddItem t
    Next t
    buffer = ""
    Stat = "Finished Loading Data"
End Sub

Private Sub Command2_Click()
''This keeps running through the file
''till it gets to the requested picture number,
''then saves it data to a file, then loads it
''into a picture box from there.

Dim g As String
If List1 = "" Then
    ''error handling code
    MsgBox "select a picture number"
    Exit Sub
End If
Dim buffer As String
Stat = "Loading pic " & List1
    FN = FreeFile
    out$ = ""
    ''lets open up that frx in binary
    Open Text1.Text For Binary As #FN
        ''make a buffer large anough to handle the whole file
        buffer = String(LOF(FN), 0)
        ''store the file size for later use
        FSize = LOF(FN)
        DoEvents
        ''load all of the file into our buffer
        Get #FN, 1, buffer$
        ''we have all the data stored in ram, so lets close the file.
    Close #FN

''reset the varibles
times1 = 1
FoundNum = 0

Do
    ''find the next first picture
    IsIn = InStr(times1, buffer$, FHeader(), vbBinaryCompare)
    ''check if it does exist
    If (IsIn > 0) Then
        ''incrament the found amount
        FoundNum = FoundNum + 1
        ''now lets make sure it dosn't pick up the same file twice.
        times1 = IsIn + 2
    ElseIf (IsIn = 0) Then
        ''if none is found, jump to the end
        GoTo end__:
    End If
    ''lets see if its the picture number they asked for
    If (FoundNum = List1) Then
        ''lets store the pos of the start of the next picture
        ''in a varible. because thats also the end of our picture
        xx = InStr(times1, buffer$, FHeader(), vbBinaryCompare)
        ''if its not there, use the rest of the file as an ending point
        If (xx = 0) Then
            xx = FSize
        End If
        ''lets slosh every thing between thouse 2 points into a varible
        g = Mid(buffer$, IsIn + 9, xx)
        GoTo end__:
    End If
Loop
end__:
    ''lets open up a file, and tos the picture into there
    Open "C:\FRXPic.bmp" For Output As #1
        Print #1, g
    Close #1
    ''lets free up some ram
    g = ""
    buffer$ = ""
    ''load the picture from that saved file
    Picture3.Picture = LoadPicture("C:\FRXPic.bmp")
    ''resize the window to make sure its all showing
    Me.Height = Picture3.Height + Picture3.Top + (Label1.Height * 3)
    Stat.Top = Me.Height - (Stat.Height * 2.5)
    If (Picture3.Width + (Picture3.Left * 2)) > Me.Width Then
        Me.Width = Picture3.Width + (Picture3.Left * 2)
    Else
        Me.Width = 6195
    End If
    Stat = "Finished Loading pic " & List1
End Sub
Private Sub Command3_Click()
CommonDialog1.Filter = "Binary VB Picture files|*.frx"
CommonDialog1.Action = 1
Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command4_Click()
CommonDialog1.Filter = "All pic types|*.bmp;*.ico|BitMaps|*.bmp|Icons|*.ico|All file types|*.*"
CommonDialog1.Action = 2
If (CommonDialog1.FileName = "") Then Exit Sub
SavePicture Picture3.Picture, CommonDialog1.FileName
Stat = "Saved"
End Sub


Private Sub Form_Load()

End Sub

Private Sub List1_DblClick()
Call Command2_Click
End Sub


