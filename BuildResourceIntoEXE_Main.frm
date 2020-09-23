VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form BuildResourceIntoExeMain 
   Caption         =   "Build a Resource file (e.g. .WAV .DBF .HTM .ASP etc.) into your EXE for easy distribution"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "OPEN BASIC MODULE IN NOTEPAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   1200
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PLEASE EXPLAIN THIS AGAIN..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BROWSE FOR RESOURCE FILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUILD THE BASIC MODULE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BROWSE FOR IT !"
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   9825
   End
End
Attribute VB_Name = "BuildResourceIntoExeMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 On Error Resume Next
 MkDir "c:\Program Files"
 MkDir "c:\Program Files\V8Software\"
 MkDir "c:\Program Files\V8Software\Temp\"
'============================================
'EG "c:\Program Files\V8Software\CashReg.wav"
'============================================
 MyResource$ = Label1.Caption
 If InStr(MyResource$, "\") = False Then
  MsgBox "Please browse for the file first!", vbApplicationModal + vbExclamation, "Whoops!"
  Exit Sub
 End If
 ShortResourceFileName$ = MyResource$
 While InStr(ShortResourceFileName$, "\")
  S = InStr(ShortResourceFileName$, "\")
  ShortResourceFileName$ = Right$(ShortResourceFileName$, Len(ShortResourceFileName$) - S)
 Wend
 Open MyResource$ For Binary Shared As #1
 Open "c:\Program Files\V8Software\Temp\MyTempResource.tmp" For Output As #2
 WW = 1
 BCount = 0
 B$ = "  ResourceData$(" & Trim$(WW) & ") = " & Chr$(34)
 For i = 1 To LOF(1)
  a$ = Input$(1, #1)
  c = Asc(a$)
  CC$ = Hex$(c)
  If Len(CC$) <> 2 Then
   CC$ = CC$ & " "
  End If
  BCount = BCount + 1
  B$ = B$ & CC$
  If BCount = 240 Then
   B$ = B$ & Chr$(34)
   Print #2, B$
   WW = WW + 1
   BCount = 0
   B$ = "  ResourceData$(" & Trim$(WW) & ") = " & Chr$(34)
  End If
 Next i
 B$ = B$ & Chr$(34)
 Print #2, B$
 Close
 Open "c:\Program Files\V8Software\Temp\MyTempResource.tmp" For Input As #1
 Open "c:\Program Files\V8Software\Temp\MyTempResource.bas" For Output As #2
'================================
'Build Extract & Create Source...
'================================
 Dim CD$(16)
 CD$(1) = " '============================================="
 CD$(2) = " ' PASTE THIS ENTIRE CODE INTO YOUR BAS MODULE "
 CD$(3) = " '============================================="
 For i = 1 To 3
  Print #2, CD$(i)
 Next i
 CD$(4) = "  Dim ResourceData$(" & Trim$(WW) & ")"
 Print #2, "Sub BuildMyResourceFile1()"
 Print #2, "  If Dir(" & Chr$(34) & "c:\Program Files\V8Software\" & ShortResourceFileName$ & Chr$(34) & ") <> " & String$(2, 34) & " Then"
 Print #2, "    Exit Sub"
 Print #2, "  End If"
 Print #2, CD$(4)
 While Not EOF(1)
  Line Input #1, a$
  Print #2, a$
 Wend
 CD$(5) = "  DF = FreeFile"
 CD$(6) = "  Open " & Chr$(34) & "c:\Program Files\V8Software\" & ShortResourceFileName$ & Chr$(34) & " For Output As #DF"
 CD$(7) = "  For i = 1 To " & Trim$(WW)
 CD$(8) = "   a$ = ResourceData$(i)"
 CD$(9) = "   While Len(a$) > 0"
 CD$(10) = "    b$ = " & Chr$(34) & "&H" & Chr$(34) & " & Left$(a$, 2)"
 CD$(11) = "    a$ = Right$(a$, Len(a$) - 2)"
 CD$(12) = "    Print #DF, Chr$(Val(b$));"
 CD$(13) = "   Wend"
 CD$(14) = "  Next i"
 CD$(15) = "  Close #DF"
 CD$(16) = " "
 For i = 5 To 16
  Print #2, CD$(i)
 Next i
 Print #2, "End Sub "
 Close
 Call Command5_Click
End Sub
Private Sub Command2_Click()
 CommonDialog1.Action = 1
 Label1.Caption = CommonDialog1.FileName
End Sub
Private Sub Command4_Click()
 MsgBox "If you wish to distribute an EXE that requires a resource, this program makes it possible to convert the resource file into DATA contained within the EXE itself!" & String$(2, 10) & "When your USER executes the program, the resource file can be created on-the-fly." & String$(2, 10) & "Very useful when you need to distribute an application over the Internet for example." & String$(2, 10) & "I used it to create a wav file in my eBay Bid Advisor for people who SELL on eBay.  Why not look at my other postings and try that program as well." & String$(2, 10) & "The eBay Bid Advisor for Sellers on eBay includes a module 'BUILT' using the application in your hands right now!", vbApplicationModal, "Why would I use this Program Writer's Tool?"
End Sub
Private Sub Command5_Click()
 Shell "Notepad c:\Program Files\V8Software\Temp\MyTempResource.bas", vbMaximizedFocus
End Sub
