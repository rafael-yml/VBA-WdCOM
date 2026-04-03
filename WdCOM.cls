VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WdCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef pguid As GUID) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32" (ByRef pguid As GUID) As Long
#End If

Private Const WDCOM_OK           As Long = 0
Private Const WDCOM_FILE_MISSING As Long = 1
Private Const WDCOM_ENCRYPTED    As Long = 2
Private Const WDCOM_WORD_FAILED  As Long = 3
Private Const WDCOM_OPEN_FAILED  As Long = 4
Private Const WDCOM_EMPTY        As Long = 5
Private Const WDCOM_GARBLED      As Long = 6

Private Const MSO_EMBEDDED_OLE   As Long = 7
Private Const MSO_LINKED_PICTURE As Long = 11
Private Const MSO_PICTURE        As Long = 13
Private Const MAX_SAMPLE_CHARS   As Long = 100000

' ---------------------------------------------------------------------------
' Instance state
' ---------------------------------------------------------------------------
Private m_lLastStatus   As Long
Private m_dGarbleThresh As Double

' ---------------------------------------------------------------------------
' Properties
' ---------------------------------------------------------------------------

Public Property Get LastStatus() As Long
    LastStatus = m_lLastStatus
End Property

' Average word length above which extracted text is considered garbled (default 25).
' Word's Content.Text uses Chr(13) for paragraph breaks; the credibility check
' counts spaces, Chr(10), and Chr(13) as separators. Raise for languages with
' long compound words; lower for stricter detection.
Public Property Get GarbleThreshold() As Double
    GarbleThreshold = m_dGarbleThresh
End Property
Public Property Let GarbleThreshold(ByVal dVal As Double)
    If dVal > 0 Then m_dGarbleThresh = dVal
End Property

Private Sub Class_Initialize()
    m_dGarbleThresh = 25#
End Sub

Public Function IsWordInstalled() As Boolean
    Static lCached As Long
    If lCached = 1 Then
        IsWordInstalled = True
        Exit Function
    End If
    Dim oWord As Object
    On Error Resume Next
    Set oWord = CreateObject("Word.Application")
    On Error GoTo 0
    If Not oWord Is Nothing Then
        On Error Resume Next
        oWord.Quit
        Set oWord = Nothing
        On Error GoTo 0
        lCached = 1
        IsWordInstalled = True
    End If
End Function

Public Function IsWordDocument(ByVal sFilePath As String) As Boolean
    Const ZIP_MAGIC As String = "PK" & Chr(3) & Chr(4)
    Const OLE_MAGIC As String = Chr(&HD0) & Chr(&HCF) & Chr(&H11) & Chr(&HE0)
    Dim iFile   As Integer
    Dim sHeader As String
    Dim sExt    As String
    IsWordDocument = False
    If Len(Dir(sFilePath)) = 0 Then Exit Function
    sExt = LCase$(Mid$(sFilePath, InStrRev(sFilePath, ".") + 1))
    On Error GoTo Done
    iFile = FreeFile
    Open sFilePath For Binary Access Read As #iFile
    sHeader = Space$(4)
    Get #iFile, 1, sHeader
    Close #iFile
    iFile = 0
    If sHeader = OLE_MAGIC Then
        IsWordDocument = (sExt = "doc")
    ElseIf sHeader = ZIP_MAGIC Then
        IsWordDocument = (sExt = "docx" Or sExt = "docm")
    End If
Done:
    On Error Resume Next
    If iFile > 0 Then Close #iFile
    On Error GoTo 0
End Function

Public Function ExtractText(ByVal sFilePath As String) As String
    Dim oWord  As Object
    Dim oDoc   As Object
    Dim sText  As String
    Dim lCheck As Long

    m_lLastStatus = WDCOM_OK
    ExtractText = ""

    If Len(Dir(sFilePath)) = 0 Then
        m_lLastStatus = WDCOM_FILE_MISSING
        Exit Function
    End If

    If WordCOM_IsEncrypted(sFilePath) Then
        m_lLastStatus = WDCOM_ENCRYPTED
        Exit Function
    End If

    On Error Resume Next
    Set oWord = CreateObject("Word.Application")
    On Error GoTo 0
    If oWord Is Nothing Then
        m_lLastStatus = WDCOM_WORD_FAILED
        Exit Function
    End If

    On Error Resume Next
    oWord.Visible = False
    oWord.DisplayAlerts = 0
    oWord.Options.ConfirmConversions = False
    oWord.Options.UpdateLinksAtOpen = False
    oWord.Options.UpdateFieldsAtPrint = False
    oWord.AutomationSecurity = 3
    On Error GoTo 0

    On Error GoTo OpenFailed
    Set oDoc = oWord.Documents.Open( _
        Filename:=sFilePath, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        OpenAndRepair:=False, _
        Visible:=False)
    On Error GoTo 0

    If oDoc Is Nothing Then
        m_lLastStatus = WDCOM_OPEN_FAILED
        GoTo Cleanup
    End If

    On Error Resume Next
    sText = oDoc.Content.Text
    On Error GoTo 0

    lCheck = WordCOM_CheckCredibility(sText)
    If lCheck <> WDCOM_OK Then
        m_lLastStatus = lCheck
        GoTo Cleanup
    End If

    ExtractText = WordCOM_CleanText(sText)
    GoTo Cleanup

OpenFailed:
    m_lLastStatus = WDCOM_OPEN_FAILED

Cleanup:
    On Error Resume Next
    If Not oDoc Is Nothing Then
        oDoc.Close SaveChanges:=False
        Set oDoc = Nothing
    End If
    oWord.Quit SaveChanges:=False
    Set oWord = Nothing
    On Error GoTo 0
End Function

Public Function ExtractDocText(ByVal oDoc As Object) As String
    Dim sText  As String
    Dim lCheck As Long
    m_lLastStatus = WDCOM_OK
    ExtractDocText = ""
    On Error Resume Next
    sText = oDoc.Content.Text
    On Error GoTo 0
    lCheck = WordCOM_CheckCredibility(sText)
    If lCheck = WDCOM_OK Then
        ExtractDocText = WordCOM_CleanText(sText)
    Else
        m_lLastStatus = lCheck
    End If
End Function

Public Function PageCount(ByVal oDoc As Object, Optional ByVal bFast As Boolean = False) As Long
    PageCount = 0
    On Error Resume Next
    If bFast Then
        PageCount = oDoc.BuiltinDocumentProperties(14)
    Else
        PageCount = oDoc.ComputeStatistics(2)
        If PageCount = 0 Then _
            PageCount = oDoc.BuiltinDocumentProperties(14)
    End If
    On Error GoTo 0
End Function

Public Sub ExportImages(ByVal oDoc As Object, ByVal sOutputFolder As String)
    If Right$(sOutputFolder, 1) <> "\" Then sOutputFolder = sOutputFolder & "\"
    If Len(Dir(sOutputFolder, vbDirectory)) = 0 Then
        On Error Resume Next
        MkDir sOutputFolder
        On Error GoTo 0
    End If

    Dim oWs   As Worksheet
    Dim oIS   As Object
    Dim oSh   As Object
    Dim sPath As String
    Dim i     As Long
    i = 1

    Set oWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    oWs.Visible = xlSheetVeryHidden

    On Error Resume Next

    For Each oIS In oDoc.InlineShapes
        Err.Clear
        oIS.Range.CopyAsPicture
        If Err.Number = 0 Then
            sPath = sOutputFolder & "word_img_" & i & ".png"
            If WordCOM_PasteToFile(oWs, oIS.Width, oIS.Height, sPath) Then i = i + 1
        End If
    Next oIS

    For Each oSh In oDoc.Shapes
        If oSh.Type = MSO_PICTURE Or _
           oSh.Type = MSO_LINKED_PICTURE Or _
           oSh.Type = MSO_EMBEDDED_OLE Then
            Err.Clear
            oSh.CopyAsPicture
            If Err.Number = 0 Then
                sPath = sOutputFolder & "word_img_" & i & ".png"
                If WordCOM_PasteToFile(oWs, oSh.Width, oSh.Height, sPath) Then i = i + 1
            End If
        End If
    Next oSh

    On Error GoTo 0

Cleanup:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not oWs Is Nothing Then oWs.Delete
    Application.DisplayAlerts = True
    Set oWs = Nothing
    On Error GoTo 0
End Sub

Public Function GetImages(ByVal oDoc As Object) As Collection
    Dim oResult  As New Collection
    Set GetImages = oResult
    Dim oWs      As Worksheet
    Dim oIS      As Object
    Dim oSh      As Object
    Dim sTmp     As String
    Dim aBytes() As Byte

    Set oWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    oWs.Visible = xlSheetVeryHidden

    On Error Resume Next

    For Each oIS In oDoc.InlineShapes
        Err.Clear
        oIS.Range.CopyAsPicture
        If Err.Number = 0 Then
            sTmp = WordCOM_TempPath()
            If WordCOM_PasteToFile(oWs, oIS.Width, oIS.Height, sTmp) Then
                aBytes = WordCOM_ReadBytes(sTmp)
                If UBound(aBytes) > 0 Then oResult.Add aBytes
            End If
        End If
    Next oIS

    For Each oSh In oDoc.Shapes
        If oSh.Type = MSO_PICTURE Or _
           oSh.Type = MSO_LINKED_PICTURE Or _
           oSh.Type = MSO_EMBEDDED_OLE Then
            Err.Clear
            oSh.CopyAsPicture
            If Err.Number = 0 Then
                sTmp = WordCOM_TempPath()
                If WordCOM_PasteToFile(oWs, oSh.Width, oSh.Height, sTmp) Then
                    aBytes = WordCOM_ReadBytes(sTmp)
                    If UBound(aBytes) > 0 Then oResult.Add aBytes
                End If
            End If
        End If
    Next oSh

    On Error GoTo 0

    On Error Resume Next
    Application.DisplayAlerts = False
    If Not oWs Is Nothing Then oWs.Delete
    Application.DisplayAlerts = True
    Set oWs = Nothing
    On Error GoTo 0

    Set GetImages = oResult
End Function

Private Function WordCOM_PasteToFile(ByVal oWs As Worksheet, _
                                   ByVal dWidth As Double, _
                                   ByVal dHeight As Double, _
                                   ByVal sPath As String) As Boolean
    Dim oChart As ChartObject
    On Error GoTo Failed
    Set oChart = oWs.ChartObjects.Add(0, 0, dWidth, dHeight)
    oChart.Activate
    oChart.Chart.Paste
    oChart.Chart.Export Filename:=sPath, Filtername:="PNG"
    oChart.Delete
    Set oChart = Nothing
    WordCOM_PasteToFile = (Len(Dir(sPath)) > 0)
    Exit Function
Failed:
    On Error Resume Next
    If Not oChart Is Nothing Then oChart.Delete
    Set oChart = Nothing
    WordCOM_PasteToFile = False
End Function

Private Function WordCOM_ReadBytes(ByVal sPath As String) As Byte()
    Dim aEmpty(0) As Byte
    Dim iFile     As Integer
    Dim lSize     As Long
    Dim aBytes()  As Byte
    WordCOM_ReadBytes = aEmpty
    On Error GoTo Failed
    iFile = FreeFile
    Open sPath For Binary Access Read As #iFile
    lSize = LOF(iFile)
    If lSize = 0 Then Close #iFile: GoTo Failed
    ReDim aBytes(0 To lSize - 1)
    Get #iFile, , aBytes
    Close #iFile
    iFile = 0
    On Error Resume Next
    Kill sPath
    On Error GoTo 0
    WordCOM_ReadBytes = aBytes
    Exit Function
Failed:
    On Error Resume Next
    If iFile > 0 Then Close #iFile
    Kill sPath
    On Error GoTo 0
End Function

Private Function WordCOM_TempPath() As String
    Dim g As GUID
    CoCreateGuid g
    WordCOM_TempPath = Environ("TEMP") & "\wdcom_" & Right$("00000000" & Hex$(g.Data1), 8) & ".png"
End Function

Private Function WordCOM_CheckCredibility(ByVal sText As String) As Long
    Dim sSample As String
    Dim lSeps   As Long
    Dim dAvg    As Double

    If Len(sText) < 2 Then
        WordCOM_CheckCredibility = WDCOM_EMPTY
        Exit Function
    End If

    sSample = Left$(sText, MAX_SAMPLE_CHARS)
    lSeps = Len(sSample) - Len(Replace(sSample, " ", "")) _
          + Len(sSample) - Len(Replace(sSample, Chr(13), "")) _
          + Len(sSample) - Len(Replace(sSample, Chr(10), ""))

    If lSeps > 0 Then
        dAvg = Len(sSample) / CDbl(lSeps + 1)
    Else
        dAvg = Len(sSample)
    End If

    If dAvg > m_dGarbleThresh Then
        WordCOM_CheckCredibility = WDCOM_GARBLED
    Else
        WordCOM_CheckCredibility = WDCOM_OK
    End If
End Function

Private Function WordCOM_IsEncrypted(ByVal sFilePath As String) As Boolean
    Const HEAD_READ As Long = 512
    Const TAIL_READ As Long = 2048
    Dim iFile    As Integer
    Dim lSize    As Long
    Dim sHead    As String
    Dim sTail    As String
    Dim lTailPos As Long

    On Error GoTo NotEncrypted
    iFile = FreeFile
    Open sFilePath For Binary Access Read As #iFile
    lSize = LOF(iFile)
    sHead = Space$(IIf(lSize < HEAD_READ, lSize, HEAD_READ))
    Get #iFile, 1, sHead
    lTailPos = IIf(lSize > TAIL_READ, lSize - TAIL_READ + 1, 1)
    sTail = Space$(IIf(lSize < TAIL_READ, lSize, TAIL_READ))
    Get #iFile, lTailPos, sTail
    Close #iFile
    WordCOM_IsEncrypted = WordCOM_HasEncryptKey(sHead) Or WordCOM_HasEncryptKey(sTail)
    Exit Function
NotEncrypted:
    On Error Resume Next
    If iFile > 0 Then Close #iFile
    On Error GoTo 0
    WordCOM_IsEncrypted = False
End Function

Private Function WordCOM_HasEncryptKey(ByVal s As String) As Boolean
    Const key As String = "/Encrypt"
    Dim p   As Long
    Dim pos As Long
    Dim nxt As String
    p = 1
    Do
        p = InStr(p, s, key, vbBinaryCompare)
        If p = 0 Then Exit Do
        pos = p + Len(key)
        If pos > Len(s) Then
            WordCOM_HasEncryptKey = True
            Exit Function
        End If
        nxt = Mid$(s, pos, 1)
        If (nxt < "A" Or nxt > "Z") And (nxt < "a" Or nxt > "z") Then
            WordCOM_HasEncryptKey = True
            Exit Function
        End If
        p = p + 1
    Loop
End Function

Private Function WordCOM_CleanText(ByVal sText As String) As String
    Dim s          As String
    Dim aLines()   As String
    Dim aOut()     As String
    Dim i          As Long
    Dim nOut       As Long
    Dim bLastEmpty As Boolean

    s = Replace(sText, Chr(13) & Chr(10), Chr(10))
    s = Replace(s, Chr(13), Chr(10))

    aLines = Split(s, Chr(10))
    ReDim aOut(0 To UBound(aLines))
    nOut = 0
    bLastEmpty = False

    For i = 0 To UBound(aLines)
        If Len(Trim(aLines(i))) = 0 Then
            If Not bLastEmpty Then
                aOut(nOut) = ""
                nOut = nOut + 1
            End If
            bLastEmpty = True
        Else
            aOut(nOut) = aLines(i)
            nOut = nOut + 1
            bLastEmpty = False
        End If
    Next i

    If nOut > 0 Then
        ReDim Preserve aOut(0 To nOut - 1)
        WordCOM_CleanText = Trim(Join(aOut, Chr(10)))
    End If
End Function
