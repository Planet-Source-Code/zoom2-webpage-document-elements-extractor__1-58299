VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Enum enElemType
  elemLinks = 1
  elemTables = 2
  elemImages = 3
  elemInput = 4
End Enum



Private WithEvents oWB   As WebBrowser
Attribute oWB.VB_VarHelpID = -1
Private WithEvents oDOC  As HTMLDocument
Attribute oDOC.VB_VarHelpID = -1


Event LinkElementReturned(linkCount As Long, oLINK As HTMLAnchorElement)
Event TableElementReturned(tableCount As Long, oTABLE As HTMLTable)
Event ImageElementReturned(imageCount As Long, oIMAGE As HTMLImg)
Event InputElementReturned(inputCount As Long, oINPUT As HTMLInputElement)



Function StripdownBrowser(yourBrowser As WebBrowser, _
                   Optional BrowserCanBeDropTarget As Boolean = False) As Long
On Error GoTo local_error:

                StripdownBrowser = 1
                
                With yourBrowser
                  .AddressBar = False
                  .FullScreen = False
                  .MenuBar = False
                  .RegisterAsBrowser = False
                  .RegisterAsDropTarget = BrowserCanBeDropTarget
                  .Silent = True
                  .StatusBar = False
                  .TheaterMode = False
                End With
local_error:
   If Err.Number <> 0 Then
       StripdownBrowser = 0
       Debug.Print "cBrowser.StripdownBrowser: " & Err.Number & "." & Err.Description
       Err.Clear
       Resume Next
   End If
End Function
 

Function ExtractAllOfElements(yourBrowserDoc As HTMLDocument, _
                                   ElementType As enElemType)
On Error GoTo local_error:
Dim upper   As Long
Dim lcnt    As Long
Dim DOC     As HTMLDocument
Dim oA      As HTMLAnchorElement
Dim oIMG    As HTMLImg
Dim oTABLE  As HTMLTable
Dim oINP    As HTMLInputElement
Dim tagType As String
                  
                Set DOC = yourBrowserDoc
                tagType = Choose(ElementType, "A" _
                                            , "TABLE" _
                                            , "IMG" _
                                            , "INPUT")
                With yourBrowserDoc
                   upper = (DOC.getElementsByTagName(tagType).length - 1)
                   
                   For lcnt = 0 To upper
                      If ElementType = elemLinks Then 'LINKS
                         Set oA = DOC.getElementsByTagName(tagType)(lcnt)
                         RaiseEvent LinkElementReturned(upper + 1, oA)
                      ElseIf ElementType = elemImages Then 'IMAGES
                         Set oIMG = DOC.getElementsByTagName(tagType)(lcnt)
                         RaiseEvent ImageElementReturned(upper + 1, oIMG)
                      ElseIf ElementType = elemTables Then 'TABLES
                         Set oTABLE = DOC.getElementsByTagName(tagType)(lcnt)
                         RaiseEvent TableElementReturned(upper + 1, oTABLE)
                      ElseIf ElementType = elemInput Then 'INPUTS
                         Set oINP = DOC.getElementsByTagName(tagType)(lcnt)
                         RaiseEvent InputElementReturned(upper + 1, oINP)
                      End If
                   Next lcnt
                   
                   Set oA = Nothing
                   Set oIMG = Nothing
                   Set oTABLE = Nothing
                   Set oINP = Nothing
                End With
local_error:
   If Err.Number <> 0 Then
       Debug.Print "cBrowser.ExtractAllOfElement: " & Err.Number & "." & Err.Description
       Err.Clear
       Resume Next
   End If
End Function

Private Sub Class_Terminate()
           
           Set oWB = Nothing
           Set oDOC = Nothing
End Sub
