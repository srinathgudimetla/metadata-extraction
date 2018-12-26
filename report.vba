'Attribute VB_Name = "RegexCore"
' Module-level declaration. The object will persist between calls
Private pCachedRegexes As Dictionary

' clipboard code
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) _
   As Long
Private Declare Function GlobalAlloc Lib "kernel32" ( _
   ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function SetClipboardData Lib "user32" ( _
   ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias _
   "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Private Declare Function GlobalUnlock Lib "kernel32" ( _
   ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   pDest As Any, pSource As Any, ByVal cbLength As Long)
Private Declare Function GetClipboardData Lib "user32" ( _
   ByVal wFormat As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" ( _
   ByVal lpData As Long) As Long

Private Const m_sDescription = _
                  "Version:1.0" & vbCrLf & _
                  "StartHTML:aaaaaaaaaa" & vbCrLf & _
                  "EndHTML:bbbbbbbbbb" & vbCrLf & _
                  "StartFragment:cccccccccc" & vbCrLf & _
                  "EndFragment:dddddddddd" & vbCrLf
                  
Private m_cfHTMLClipFormat As Long

Public Sub Execute()

    Const wdStory = 6
    Const wdMove = 0
    
    Dim doc As Object
    Dim objSelection As Object
    Set doc = ActiveDocument
    Set objSelection = doc.ActiveWindow.Selection
    objSelection.PageSetup.Orientation = wdOrientLandscape
    objSelection.EndKey unit:=wdStory
    

    Dim resp As String
    Dim respSA As Object
    
    'Get all Entities and related Attributes
    resp = HelloConsensus("http://localhost:18443/usventure/api/v1/workspaces/33/artifacts?artifact_type_id=1001&include_relationships=true" _
            , GetCredentials("sgudimetla", "icc123"), "GET", "")
    
    Set respSA = JSONlib.parse(resp)
    
    For Each sa In respSA("artifacts")
    
        objSelection.EndKey wdStory, wdMove
        objSelection.InsertBreak Type:=wdPageBreak
        
        If Not Trim(sa("definition") & vbNullString) = vbNullString Then
            PutHTMLClipboard (sa("definition"))
        End If
        
        With objSelection
            .EndKey wdStory, wdMove
            .Style = "Heading 1"
            .TypeText sa("name")
        End With
        
        If Not Trim(sa("definition") & vbNullString) = vbNullString Then
            With objSelection
                .EndKey wdStory, wdMove
                .TypeParagraph
                .Range.PasteSpecial dataType:=WdPasteDataType.wdPasteHTML
                .Style = "Normal"
            End With
        Else
            With objSelection
                .EndKey wdStory, wdMove
                .TypeParagraph
                .Style = "Normal"
            End With
        End If
        
        With objSelection
            .EndKey wdStory, wdMove
            .TypeText "List of Attributes:"
        End With
        
        
        'Create table for Attributes
        Dim tblAttr As Table
        Dim objRange As Object
        Set objRange = doc.Range
        
        With objRange
            .Collapse Direction:=wdCollapseEnd
            '.InsertParagraphAfter
            .Collapse Direction:=wdCollapseEnd
        End With
        
        Set tblAttr = doc.Tables.Add(objRange, sa("relationships").count + 1, 4)
        
        Dim rowNum
        rowNum = 2
        tblAttr.Cell(1, 1).Range = "Attribute Name"
        tblAttr.Cell(1, 2).Range = "Attribute Definition"
        tblAttr.Cell(1, 3).Range = "Columns"
        tblAttr.Cell(1, 4).Range = "Sample Values"
        
        With tblAttr
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = False
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = False
            .ApplyStyleRowBands = True
            .ApplyStyleColumnBands = False
            .Style = "Grid Table 4 - Accent 5"
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
            .Columns(1).PreferredWidth = 10
            .Columns(2).PreferredWidth = 55
            .Columns(3).PreferredWidth = 10
            .Columns(4).PreferredWidth = 25
        End With
        
        For Each attr In sa("relationships")
        
            Dim respAttr As Object
            resp = HelloConsensus("http://localhost:18443/usventure/api/v1/workspaces/33/artifacts/" + CStr(attr("related_artifact_id")) + "?include_properties=true&&include_relationships=true&include_related_artifact=true" _
                    , GetCredentials("sgudimetla", "icc123"), "GET", "")

            Set respAttr = JSONlib.parse(resp)
            
            tblAttr.Cell(rowNum, 1).Range = respAttr("name")
            tblAttr.Cell(rowNum, 1).Range.Font.Bold = False
            
            If Not Trim(respAttr("definition") & vbNullString) = vbNullString Then
                PutHTMLClipboard (respAttr("definition"))
                tblAttr.Cell(rowNum, 2).Range.PasteSpecial dataType:=WdPasteDataType.wdPasteHTML
                tblAttr.Cell(rowNum, 2).Range.Style = "Normal"
            End If
            
            'relationships
            Dim colValue As String
            colValue = ""
            For Each col In respAttr("relationships")
                
                If col("related_artifact_type_id") = 11 Then
                    If Not Trim(col("related_artifact")("name") & vbNullString) = vbNullString Then
                        colValue = colValue + CStr(col("related_artifact")("name")) & Chr(11)
                    End If
                End If
            Next col
            
            tblAttr.Cell(rowNum, 3).Range = colValue
            tblAttr.Cell(rowNum, 3).Range.Style = "Normal"
            
            'properties
            For Each prop In respAttr("properties")
            
                If prop("key") = "Sample" Then
                    If Not Trim(prop("value") & vbNullString) = vbNullString Then
                        PutHTMLClipboard (prop("value"))
                        tblAttr.Cell(rowNum, 4).Range.PasteSpecial dataType:=WdPasteDataType.wdPasteHTML
                        tblAttr.Cell(rowNum, 4).Range.Style = "Normal"
                    End If
                End If
            Next prop
            
            rowNum = rowNum + 1

        Next attr
        
    Next sa
    
    MsgBox "Process complete"
    
End Sub

Public Function HelloConsensus(url2 As Variant, auth As Variant, requestType As Variant, requestText As Variant) As String

    Dim objRequest As Object
    Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    objRequest.setTimeouts 30000, 30000, 30000, 30000
    url = Replace(url2, "https://host1.balancedinsight.com:443", "http://localhost:18443")
    
    Select Case requestType
    Case "GET"
        objRequest.Open "GET", url, False
        objRequest.setRequestHeader "Authorization", auth
        objRequest.send
    Case "PATCH"
        objRequest.Open "PATCH", url, False
        objRequest.setRequestHeader "Authorization", auth
        objRequest.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        objRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        objRequest.send (requestText)
    Case "POST"
        objRequest.Open "POST", url, False
        objRequest.setRequestHeader "Authorization", auth
        objRequest.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        objRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        objRequest.send (requestText)
    End Select
    
    'objRequest.waitForResponse (30)
    Select Case objRequest.Status
        Case 200
            If requestType = "GET" Then
                HelloConsensus = objRequest.responseText
            ElseIf requestType = "PATCH" Then
                HelloConsensus = "OK"
            End If
        Case 201
            HelloConsensus = objRequest.responseText
        Case 400
            HelloConsensus = "Bad Request"
        Case 401
            HelloConsensus = "Unauthorized"
        Case 403
            HelloConsensus = "Forbidden"
        Case 404
            HelloConsensus = "Not Found"
        Case 500
            HelloConsensus = "Internal Server Error"
        Case 503
            HelloConsensus = "Busy, please retry"
        Case Else
            HelloConsensus = "Unknown Error. Further investigation is required"
    End Select

End Function

Public Function GetCredentials(uname As String, pwd As String) As String
On Error GoTo FailedState
    'GetCredentials = "Basic " + EncodeBase64(uname + ":" + pwd)
    GetCredentials = "Basic c2d1ZGltZXRsYTppY2MxMjM="
    Exit Function
FailedState:
   MsgBox Err.Number & ": " & Err.Description
End Function

Public Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument
  Set objNode = objXML.createElement("b64")

  objNode.dataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

Private Sub DeleteEmptyRows()

    Dim oTable As Table, oRow As Range, oCell As Cell, Counter As Long, NumRows As Long, TextInRow As Boolean
    
    ' Specify which table you want to work on.
    'Set oTable = Selection.Tables(1)
    
    For Each oTable In ActiveDocument.Tables
        ' Set a range variable to the first row's range
        Set oRow = oTable.Rows(1).Range
        NumRows = oTable.Rows.count
        Application.ScreenUpdating = False
        
        For Counter = 1 To NumRows
        
            StatusBar = "Row " & Counter
            TextInRow = False
        
            For Each oCell In oRow.Rows(1).Cells
                If Len(oCell.Range.text) > 2 Then
                    'end of cell marker is actually 2 characters
                    TextInRow = True
                    Exit For
                End If
            Next oCell
        
            If TextInRow Then
                Set oRow = oRow.Next(wdRow)
            Else
                oRow.Rows(1).Delete
            End If
        
        Next Counter
        
    Next oTable
    
    Application.ScreenUpdating = True

End Sub


Function RegisterCF() As Long


   'Register the HTML clipboard format
   If (m_cfHTMLClipFormat = 0) Then
      m_cfHTMLClipFormat = RegisterClipboardFormat("HTML Format")
   End If
   RegisterCF = m_cfHTMLClipFormat
   
End Function


Public Sub PutHTMLClipboard(sHtmlFragment As String, _
   Optional sContextStart As String = "<HTML><BODY>", _
   Optional sContextEnd As String = "</BODY></HTML>")
   
   Dim sData As String
   
   If RegisterCF = 0 Then Exit Sub
   
   'Add the starting and ending tags for the HTML fragment
   sContextStart = sContextStart & "<!--StartFragment -->"
   sContextEnd = "<!--EndFragment -->" & sContextEnd
   
   'Build the HTML given the description, the fragment and the context.
   'And, replace the offset place holders in the description with values
   'for the offsets of StartHMTL, EndHTML, StartFragment and EndFragment.
   sData = m_sDescription & sContextStart & sHtmlFragment & sContextEnd
   sData = Replace(sData, "aaaaaaaaaa", _
                   Format(Len(m_sDescription), "0000000000"))
   sData = Replace(sData, "bbbbbbbbbb", Format(Len(sData), "0000000000"))
   sData = Replace(sData, "cccccccccc", Format(Len(m_sDescription & _
                   sContextStart), "0000000000"))
   sData = Replace(sData, "dddddddddd", Format(Len(m_sDescription & _
                   sContextStart & sHtmlFragment), "0000000000"))

   'Add the HTML code to the clipboard
   If CBool(OpenClipboard(0)) Then
   
      Dim hMemHandle As Long, lpData As Long
      
      hMemHandle = GlobalAlloc(0, Len(sData) + 10)
      
      If CBool(hMemHandle) Then
               
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            
            CopyMemory ByVal lpData, ByVal sData, Len(sData)
            GlobalUnlock hMemHandle
            EmptyClipboard
            SetClipboardData m_cfHTMLClipFormat, hMemHandle
                        
         End If
      
      End If
   
      Call CloseClipboard
   End If

End Sub

Public Function GetHTMLClipboard() As String

   Dim sData As String
   
   If RegisterCF = 0 Then Exit Function
   
   If CBool(OpenClipboard(0)) Then
   
      Dim hMemHandle As Long, lpData As Long
      Dim nClipSize As Long
      
      GlobalUnlock hMemHandle

      'Retrieve the data from the clipboard
      hMemHandle = GetClipboardData(m_cfHTMLClipFormat)
      
      If CBool(hMemHandle) Then
               
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            nClipSize = lstrlen(lpData)
            sData = String(nClipSize + 10, 0)
            

            Call CopyMemory(ByVal sData, ByVal lpData, nClipSize)
            
            Dim nStartFrag As Long, nEndFrag As Long
            Dim nIndx As Long
            
            'If StartFragment appears in the data's description,
            'then retrieve the offset specified in the description
            'for the start of the fragment. Likewise, if EndFragment
            'appears in the description, then retrieve the
            'corresponding offset.
            nIndx = InStr(sData, "StartFragment:")
            If nIndx Then
               nStartFrag = CLng(Mid(sData, _
                                 nIndx + Len("StartFragment:"), 10))

            End If
            nIndx = InStr(sData, "EndFragment:")
            If nIndx Then
               nEndFrag = CLng(Mid(sData, nIndx + Len("EndFragment:"), 10))
            End If
            
            'Return the fragment given the starting and ending
            'offsets
            If (nStartFrag > 0 And nEndFrag > 0) Then
               GetHTMLClipboard = Mid(sData, nStartFrag + 1, _
                                 (nEndFrag - nStartFrag))
            End If
                        
         End If
      
      End If

   
      Call CloseClipboard
   End If


End Function




