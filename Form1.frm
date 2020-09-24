VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Text            =   "http://www..com"
      ToolTipText     =   "ENTER WEBADDRESS AND HIT ""ENTER"" KEY ][DBL CLICK TO RESET"
      Top             =   45
      Width           =   2940
   End
   Begin VB.Frame Frame1 
      Caption         =   "RETURN..."
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1710
      TabIndex        =   3
      Top             =   450
      Width           =   5190
      Begin VB.OptionButton optElem 
         Caption         =   "input"
         Height          =   195
         Index           =   3
         Left            =   3375
         TabIndex        =   7
         Top             =   135
         Width           =   735
      End
      Begin VB.OptionButton optElem 
         Caption         =   "images"
         Height          =   195
         Index           =   2
         Left            =   2550
         TabIndex        =   6
         Top             =   135
         Width           =   825
      End
      Begin VB.OptionButton optElem 
         Caption         =   "tables"
         Height          =   195
         Index           =   1
         Left            =   1815
         TabIndex        =   5
         Top             =   135
         Width           =   735
      End
      Begin VB.OptionButton optElem 
         Caption         =   "links"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   135
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return"
      Enabled         =   0   'False
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   495
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   5910
      Left            =   6930
      TabIndex        =   1
      Top             =   450
      Width           =   3750
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5595
      Left            =   90
      TabIndex        =   0
      Top             =   855
      Width           =   6765
      ExtentX         =   11933
      ExtentY         =   9869
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   195
      Left            =   7200
      TabIndex        =   9
      Top             =   225
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private WithEvents oWeb      As cBrowser
Attribute oWeb.VB_VarHelpID = -1

Dim m_doc_ready              As Byte
Dim opt_num_selected         As Long

Private Sub Command1_Click()

                 List1.Clear
                 
                 With Label1
                   Select Case opt_num_selected
                      Case Is = 1: .Caption = "links HREF"
                      Case Is = 2: .Caption = "tables"
                      Case Is = 3: .Caption = "images href and filesize"
                      Case Is = 4: .Caption = "input elements name & type"
                   End Select
                 End With
                 
                 oWeb.ExtractAllOfElements WebBrowser1.document, opt_num_selected
End Sub

Private Sub Form_Load()
Const LB_SETHORIZONTALEXTENT = &H194

                Putfocus Text1.hwnd
                Call Text1_DblClick
                 '-- add horizontal scroll bar to list1
                 If opt_num_selected = 1 Or 3 Then _
                     SendMessageByLong List1.hwnd, LB_SETHORIZONTALEXTENT, _
                     1000, 0
                '-- pre select the first option button
                optElem(0).Value = True
                Set oWeb = New cBrowser
                '-- this will remove a lot of errors and
                '-- headaches IE can cause and stops it
                '-- from returning error dialogs
                oWeb.StripdownBrowser WebBrowser1
End Sub

Private Sub Form_Unload(Cancel As Integer)
                
                Set oWeb = Nothing
End Sub

Private Sub optElem_Click(Index As Integer)
                
                opt_num_selected = (Index + 1)
End Sub
'-- if in webbrowser_DocumentComplete you selected
'-- oWeb.ExtractAllOfElements elemImage..then this event is raised
'-- uppon each image being extracted
Private Sub oWeb_ImageElementReturned(imageCount As Long, oIMAGE As MSHTML.HTMLImg)

                Caption = imageCount & " Total images in this page"
                List1.AddItem oIMAGE.href & _
                        "  (" & oIMAGE.fileSize & " bytes )"
                        
                '--UNCOMMENT THIS TO SEE HOW MUCH INFO YOU CAN EXTRACT!!
                'debug.Print oimage.
End Sub
'-- if in webbrowser_DocumentComplete you selected
'-- oWeb.ExtractAllOfElements elemInput..then this event is raised
'-- uppon each input element being extracted
Private Sub oWeb_InputElementReturned(inputCount As Long, oINPUT As MSHTML.HTMLInputElement)
                
                Caption = inputCount & " Total input elements in this page"
                List1.AddItem oINPUT.Name & "  (" & oINPUT.Type & ")" & _
                     " is the input element type"
                     
                '--UNCOMMENT THIS TO SEE HOW MUCH INFO YOU CAN EXTRACT!!
                'debug.Print oinput.
End Sub

'-- if in webbrowser_DocumentComplete you selected
'-- oWeb.ExtractAllOfElements elemLinks..then this event is raised
'-- uppon each link being extracted
Private Sub oWeb_LinkElementReturned(linkCount As Long, oLINK As MSHTML.HTMLAnchorElement)
                
                Caption = linkCount & " Total links in this page"
                List1.AddItem oLINK.href
                
                '--UNCOMMENT THIS TO SEE HOW MUCH INFO YOU CAN EXTRACT!!
                'debug.Print olink.
End Sub
'-- if in webbrowser_DocumentComplete you selected
'-- oWeb.ExtractAllOfElements elemTable..then this event is raised
'-- uppon each table being extracted
Private Sub oWeb_TableElementReturned(tableCount As Long, oTABLE As MSHTML.HTMLTable)
                
               Caption = tableCount & " Total tables in this page"
               List1.AddItem "there are " & oTABLE.rows.length & " rows in THIS table"
               
                '--UNCOMMENT THIS TO SEE HOW MUCH INFO YOU CAN EXTRACT!!
                'debug.Print otable.
End Sub

Private Sub Text1_DblClick()
                
                Text1 = "http://www..com"
                Text1.SelStart = 11
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
                
                If KeyAscii = vbKeyReturn Then
                  Command1.Enabled = False
                  KeyAscii = 0
                  WebBrowser1.navigate Text1
                End If
End Sub
'-- this event means the webbrowsers document is ready to be used
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
                
                Command1.Enabled = True
End Sub

