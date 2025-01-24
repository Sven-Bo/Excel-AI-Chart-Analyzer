Attribute VB_Name = "mAIObjectAnalyzer"
'=======================================
' Analyze Image or Chart with OpenAI API
' Author: Sven Bosau
' Website: https://pythonandvba.com
' YouTube: https://youtube.com/@codingisfun
'
' Requirements:
'   - OpenAI API Key
'   - VBA-JSON for JSON parsing: https://github.com/VBA-tools/VBA-JSON
'   - Select an image, shape, or chart in Excel before running this macro
'   - Two named ranges in the workbook (optional):
'       1) "PromptCell"  to store or retrieve the prompt text
'       2) "OutputCell"  to place the analysis result
'=======================================

'=======================================
' CONFIGURABLE VARIABLES
'=======================================

' 1) Your OpenAI API Key (Obtain at: https://platform.openai.com/api-keys )
Const API_KEY As String = "YOUR_OPENAI_API_KEY"  ' <<=== REPLACE WITH A REAL KEY!

' 2) OpenAI Model Name (e.g., gpt-4, gpt-4o, etc.)
Const MODEL_NAME As String = "gpt-4o"

' 3) Fallback prompt if "PromptCell" is missing or empty
Const DEFAULT_PROMPT As String = "Provide the key insights from this image briefly."

' 4) Maximum tokens for the response
Const MAX_TOKENS As Long = 300

' 5) Retry settings for clipboard operations (occasionally VBA fails to paste)
Const MAX_RETRIES As Long = 10
Const RETRY_DELAY As Double = 0.1  ' in seconds

' 6) Named range for the output cell (if missing or invalid, result is shown in a MsgBox)
Const OUTPUT_CELL_RANGE As String = "OutputCell"

' 7) OpenAI Endpoint and XML HTTP ProgID (in case you need to change them)
Const OPENAI_ENDPOINT As String = "https://api.openai.com/v1/chat/completions"
Const XMLHTTP_PROG_ID As String = "MSXML2.XMLHTTP"

'=======================================
' MAIN PROCEDURE
'=======================================
Sub AnalyzeSelectedObjectWithOpenAI()
    Dim selectedChart As chart
    Dim selectedShape As Shape
    Dim finalPrompt As String
    Dim tmpFilePath As String
    Dim base64Image As String
    Dim resultText As String
    
    '--- 0) Check API Key first ---
    If InStr(1, API_KEY, "YOUR_OPENAI_API_KEY", vbTextCompare) > 0 Then
        MsgBox "No valid OpenAI API Key found. Please visit " & _
               "https://platform.openai.com/api-keys " & _
               "to create/get your key, then paste it into the code.", vbExclamation
        Exit Sub
    End If
    
    '--- 1) Detect user selection (chart or shape/picture) ---
    GetSelectedChartOrShape selectedChart, selectedShape
    If selectedChart Is Nothing And selectedShape Is Nothing Then
        MsgBox "Unsupported selection. Please select an image, shape, or chart.", vbExclamation
        Exit Sub
    End If
    
    '--- 2) Get the final prompt text (from PromptCell or fallback, with an InputBox) ---
    finalPrompt = GetPromptText(DEFAULT_PROMPT)
    If finalPrompt = vbNullString Then Exit Sub ' user canceled
    
    '--- 3) Export the chart/picture to a temporary PNG file ---
    tmpFilePath = Environ("Temp") & "\selected_image_or_chart.png"
    If Not selectedChart Is Nothing Then
        If ExportChartAsPNG(selectedChart, tmpFilePath) = False Then
            MsgBox "Failed to export the selected chart.", vbCritical
            Exit Sub
        End If
    ElseIf Not selectedShape Is Nothing Then
        If ExportShapeAsPNG(selectedShape, tmpFilePath) = False Then
            MsgBox "Failed to export the selected shape/picture.", vbCritical
            Exit Sub
        End If
    End If
    
    '--- 4) Convert the PNG file to Base64 ---
    base64Image = EncodeImageToBase64(tmpFilePath)
    
    '--- 5) Send the image + prompt to OpenAI and retrieve the analysis ---
    resultText = SendOpenAIImageRequest(base64Image, finalPrompt, MODEL_NAME, MAX_TOKENS)
    
    '--- 6) Output the result (named range or fallback to MsgBox) ---
    OutputAnalysisResult resultText
    
    '--- 7) Cleanup (delete temp file) ---
    On Error Resume Next
    Kill tmpFilePath
    On Error GoTo 0
End Sub


'=======================================
' HELPER PROCEDURES & FUNCTIONS
'=======================================

'---------------------------------------------------------
' GetSelectedChartOrShape
'   Determines whether the user selected a ChartObject or a Shape
'   Returns the references by setting selectedChart / selectedShape
'---------------------------------------------------------
Private Sub GetSelectedChartOrShape(ByRef selectedChart As chart, ByRef selectedShape As Shape)
    On Error Resume Next
    
    Select Case TypeName(Selection)
        Case "ChartObject"
            ' The user directly selected a ChartObject
            Set selectedChart = Selection.chart

        Case "ChartArea", "PlotArea", "Legend", "Series"
            ' The user selected a part of a chart, so get the parent chart
            Set selectedChart = Selection.Parent
        
        Case "Picture", "Shape"
            ' "Picture" is used by older Excel versions, "Shape" by newer
            Set selectedShape = Selection.ShapeRange(1)
            
        Case "ShapeRange"
            If Selection.ShapeRange.Count = 1 Then
                Set selectedShape = Selection.ShapeRange(1)
            End If
    End Select
    
    On Error GoTo 0
End Sub

'---------------------------------------------------------
' GetPromptText
'   Tries to read a prompt from the "PromptCell" named range.
'   Falls back to fallbackPrompt if empty/not found.
'   Shows an InputBox with usage instructions and the current prompt.
'   Returns the final prompt, or vbNullString if user cancels.
'---------------------------------------------------------
Private Function GetPromptText(fallbackPrompt As String) As String
    Dim promptCellValue As String
    Dim instructions As String
    Dim userPrompt As String
    
    ' Try reading from named range "PromptCell"
    On Error Resume Next
    promptCellValue = Range("PromptCell").Value
    On Error GoTo 0
    
    If Trim(promptCellValue) = "" Then
        promptCellValue = fallbackPrompt
    End If
    
    ' Provide user-friendly instructions:
    instructions = _
        "HOW TO USE:" & vbCrLf & _
        "1) Update the prompt if desired (text below). Click OK to proceed, Cancel to abort." & vbCrLf & _
        "2) The prompt is also stored in the named range ""PromptCell"" for future runs." & vbCrLf & _
        "3) The analysis result will be placed in the named range ""OutputCell"", if it exists." & vbCrLf & vbCrLf & _
        "PROMPT:"
    
    ' Show an InputBox with multiline instructions + the current prompt
    userPrompt = InputBox(instructions, "OpenAI Prompt", promptCellValue)
    
    If userPrompt = "" Then
        ' User canceled or cleared
        GetPromptText = vbNullString
    Else
        ' Update PromptCell for future runs, ignoring errors if named range doesn't exist
        On Error Resume Next
        Range("PromptCell").Value = userPrompt
        On Error GoTo 0
        
        GetPromptText = userPrompt
    End If
End Function

'---------------------------------------------------------
' ExportChartAsPNG
'   Exports a Chart to a PNG file.
'   Returns True if successful, False otherwise.
'---------------------------------------------------------
Private Function ExportChartAsPNG(chrt As chart, ByVal filePath As String) As Boolean
    On Error GoTo errHandler
    chrt.Export Filename:=filePath, FilterName:="PNG"
    ExportChartAsPNG = True
    Exit Function
errHandler:
    ExportChartAsPNG = False
End Function

'---------------------------------------------------------
' ExportShapeAsPNG
'   Exports a Shape (picture, shape, etc.) to a PNG file.
'
'   VBA can sometimes fail on copy/paste operations if the
'   clipboard isn't ready. So we use a retry loop to attempt
'   multiple times before giving up.
'
'   Returns True if successful, False otherwise.
'---------------------------------------------------------
Private Function ExportShapeAsPNG(shp As Shape, ByVal filePath As String) As Boolean
    Dim tempChartObj As ChartObject
    Dim retryCount As Long
    
    On Error GoTo errHandler
    
    ' Create a temporary chart object the same size as the shape
    Set tempChartObj = ActiveSheet.ChartObjects.Add( _
        Left:=shp.Left, _
        Top:=shp.Top, _
        Width:=shp.Width, _
        Height:=shp.Height)
    
    ' Retry logic for copying/pasting the shape into the chart
    ' (Clipboard operations can fail intermittently in VBA)
    For retryCount = 1 To MAX_RETRIES
        On Error Resume Next
        shp.Copy
        DoEvents  ' allow time for the clipboard to catch up
        
        tempChartObj.chart.Paste
        DoEvents
        
        On Error GoTo errHandler
        
        If tempChartObj.chart.Shapes.Count > 0 Then
            ' Successful paste
            Exit For
        Else
            Sleep (RETRY_DELAY * 1000)
        End If
    Next retryCount
    
    ' If no shapes pasted after retry attempts, fail
    If tempChartObj.chart.Shapes.Count = 0 Then
        tempChartObj.Delete
        ExportShapeAsPNG = False
        Exit Function
    End If
    
    ' Export the chart to PNG
    tempChartObj.chart.Export Filename:=filePath, FilterName:="PNG"
    
    ' Delete the temporary chart object
    tempChartObj.Delete
    
    ExportShapeAsPNG = True
    Exit Function
errHandler:
    ExportShapeAsPNG = False
End Function

'---------------------------------------------------------
' SendOpenAIImageRequest
'   Sends the Base64 image + prompt to OpenAI, returns result text.
'---------------------------------------------------------
Private Function SendOpenAIImageRequest(base64Image As String, _
                                        prompt As String, _
                                        model As String, _
                                        maxTokens As Long) As String
    Dim http As Object
    Dim payload As String
    Dim response As String
    Dim JSON As Object
    
    ' Build JSON payload
    payload = "{""model"":""" & model & """,""messages"":[{" & _
                """role"":""user"",""content"":[{" & _
                """type"":""text"",""text"":""" & Replace(prompt, """", "\""") & """},{" & _
                """type"":""image_url"",""image_url"":{""url"":""data:image/png;base64," & base64Image & """}}" & _
                "]}],""max_tokens"":" & maxTokens & "}"
    
    ' Create HTTP object and send POST
    Set http = CreateObject(XMLHTTP_PROG_ID)
    http.Open "POST", OPENAI_ENDPOINT, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & API_KEY
    http.Send payload
    
    response = http.responseText
    
    ' Parse JSON response
    On Error Resume Next
    Set JSON = JsonConverter.ParseJson(response)
    If Not JSON Is Nothing Then
        SendOpenAIImageRequest = JSON("choices")(1)("message")("content")
    Else
        SendOpenAIImageRequest = "No valid JSON response from OpenAI."
    End If
    On Error GoTo 0
End Function

'---------------------------------------------------------
' OutputAnalysisResult
'   Attempts to place the result in the named range "OutputCell".
'   If that fails, shows a MsgBox instead.
'---------------------------------------------------------
Private Sub OutputAnalysisResult(ByVal analysisResult As String)
    On Error GoTo fallback
    Range(OUTPUT_CELL_RANGE).Value = analysisResult
    Exit Sub

fallback:
    MsgBox "Analysis result: " & analysisResult, vbInformation
End Sub

'---------------------------------------------------------
' EncodeImageToBase64
'   Reads a file in binary and encodes its bytes as Base64
'---------------------------------------------------------
Private Function EncodeImageToBase64(imagePath As String) As String
    Dim bytes() As Byte
    Dim stream As Object
    Dim base64String As String
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.LoadFromFile imagePath
    bytes = stream.Read
    stream.Close
    Set stream = Nothing
    
    base64String = EncodeBase64(bytes)
    EncodeImageToBase64 = base64String
End Function

'---------------------------------------------------------
' EncodeBase64
'---------------------------------------------------------
Private Function EncodeBase64(bytes() As Byte) As String
    Dim XML As Object
    Dim node As Object
    
    Set XML = CreateObject("MSXML2.DOMDocument")
    Set node = XML.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    EncodeBase64 = Replace(node.Text, vbLf, "")
    
    Set node = Nothing
    Set XML = Nothing
End Function

'---------------------------------------------------------
' Sleep
'   Custom Sleep function using Timer (milliseconds resolution)
'---------------------------------------------------------
Private Sub Sleep(milliseconds As Long)
    Dim start As Double
    start = Timer
    Do While Timer < start + milliseconds / 1000
        DoEvents
    Loop
End Sub


