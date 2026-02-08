Attribute VB_Name = "ChatGPT"
Option Explicit

' Writes a key-value pair to the user environment variables
Public Function SetUserEnvironmentVariable(keyName As String, value As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    
    ' Write to the user environment variables
    WshShell.Environment("USER")(keyName) = value
    
    SetUserEnvironmentVariable = True
    Exit Function

ErrHandler:
    SetUserEnvironmentVariable = False
End Function

Sub TestEnvVar()
    Dim success As Boolean
    success = SetUserEnvironmentVariable("MY_ENV_VAR", "TestValue")
    
    If success Then
        MsgBox "Environment variable set successfully!"
    Else
        MsgBox "Failed to set environment variable."
    End If
End Sub

' Reads an environment variable (checks User first, then System)
Public Function GetEnvironmentVariable(keyName As String) As String
    On Error GoTo ErrHandler
    
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    
    Dim value As String
    
    ' Try user-level environment variable
    value = WshShell.Environment("USER")(keyName)
    
    ' If not found, try system-level
    If value = "" Then
        value = WshShell.Environment("SYSTEM")(keyName)
    End If
    
    GetEnvironmentVariable = value
    Exit Function

ErrHandler:
    GetEnvironmentVariable = ""
End Function

Sub TestGetEnv()
    MsgBox "MY_ENV_VAR = " & GetEnvironmentVariable("MY_ENV_VAR")
End Sub

Function CallChatGPT(apiKey As String, userPrompt As String, Optional imagePathOrUrl As String = "") As String
    On Error GoTo ErrHandler

    Dim http As Object
    Dim json As Object
    Dim url As String
    Dim requestBody As String
    Dim responseText As String
    Dim contentBlock As String
    Dim imageBlock As String
    Dim imageData As String
    Dim mimeType As String
    Dim escapedPrompt As String
    Dim dataUri As String

    Debug.Print "CallChatGPT started..."
    Debug.Print "Prompt: " & Left(userPrompt, 100) & "..."
    If imagePathOrUrl <> "" Then Debug.Print "Image path or URL: " & imagePathOrUrl

    escapedPrompt = EscapeForJson(userPrompt)

    ' Determine the content block based on whether an image is included
    If imagePathOrUrl = "" Then
        Debug.Print "No image provided. Constructing text-only message."
        contentBlock = "{""role"":""user"",""content"":""" & escapedPrompt & """}"
    Else
        ' Check if it's a URL or local file
        If LCase(Left(imagePathOrUrl, 4)) = "http" Then
            Debug.Print "Image detected as URL."
            imageBlock = "{""type"":""image_url"",""image_url"":{""url"":""" & imagePathOrUrl & """}}"
        Else
            Debug.Print "Image detected as local file. Encoding to base64..."
            imageData = EncodeImageToBase64(imagePathOrUrl)

            Select Case LCase(Right(imagePathOrUrl, 4))
                Case ".png": mimeType = "image/png"
                Case "jpeg": mimeType = "image/jpeg"
                Case ".jpg": mimeType = "image/jpeg"
                Case "webp": mimeType = "image/webp"
                Case Else: mimeType = "application/octet-stream"
            End Select
            Debug.Print "MIME type: " & mimeType

            dataUri = "data:" & mimeType & ";base64," & imageData
            Debug.Print "Encoded data URI (start): " & Left(dataUri, 100) & "..."
            imageBlock = "{""type"":""image_url"",""image_url"":{""url"":""" & EscapeForJson(dataUri) & """}}"
        End If

        contentBlock = "{""role"":""user"",""content"":[" & _
                       "{""type"":""text"",""text"":""" & escapedPrompt & """}," & _
                       imageBlock & _
                       "]}"
        Debug.Print "Content block with image constructed."
    End If

    requestBody = "{""model"":""gpt-4o"",""messages"":[" & contentBlock & "],""temperature"":0.7}"
    Debug.Print "Request body prepared."

    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://api.openai.com/v1/chat/completions"
    Debug.Print "Sending request to OpenAI API..."

    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        WriteTextToFile GetPromptsPath & "\request.txt", requestBody
        .Send requestBody
        responseText = .responseText
    End With

    Debug.Print "Response received. Length: " & Len(responseText)
    Set json = JsonConverter.ParseJson(responseText)

    If json.Exists("error") Then
        Debug.Print "API Error: " & json("error")("message")
        CallChatGPT = "API ERROR: " & json("error")("message")
        Exit Function
    End If

    CallChatGPT = CleanJsonCodeFence(json("choices")(1)("message")("content"))
    Debug.Print "Parsed response content: " & Left(CallChatGPT, 100) & "..."
    WriteTextToFile GetPromptsPath & "\response.txt", CallChatGPT
    Debug.Print "Response saved to file."

    Exit Function

ErrHandler:
    Debug.Print "Unexpected error: " & Err.Description
    CallChatGPT = "ERROR: " & Err.Description
End Function





Sub TestChatGPT()
    Dim apiKey As String
    Dim Prompt As String
    Dim response As String

    apiKey = GetEnvironmentVariable("OPENAI_API_KEY")
    Prompt = "Summarize the key points of Newton's laws of motion."

    response = CallChatGPT(apiKey, Prompt)
    MsgBox response
End Sub

Sub TestMultimodal()
    Dim response As String
    Dim imageUrl As String
    imageUrl = "https://upload.wikimedia.org/wikipedia/commons/4/47/PNG_transparency_demonstration_1.png"
    response = CallChatGPT(GetEnvironmentVariable("OPENAI_API_KEY"), "Describe this image", imageUrl)
    MsgBox response
End Sub

Function LoadPromptFromFile(promptName As String) As String
    Dim fso As Object, file As Object
    Dim promptFileURL As String

    promptFileURL = GetPromptsPath & "\" & promptName & ".txt"
    Debug.Print "Loading prompt from: " & promptFileURL

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(promptFileURL) Then
        Debug.Print "Prompt file found."
        Set file = fso.OpenTextFile(promptFileURL, 1)
        LoadPromptFromFile = file.ReadAll
        Debug.Print "Prompt loaded: " & Left(LoadPromptFromFile, 100) & "..."
        file.Close
    Else
        Debug.Print "Prompt file NOT found."
        LoadPromptFromFile = "ERROR: File not found"
    End If
End Function



Sub TestLoadPromptFromFile()
    Dim promptText As String
    Dim promptName As String
    
    promptName = "GET_TBOX_PROJECT_INFO_FROM_SCREENSHOT"
    promptText = LoadPromptFromFile(promptName)

    If Left(promptText, 6) = "ERROR:" Then
        MsgBox promptText, vbExclamation, "Prompt Load Failed"
    Else
        MsgBox promptText, vbInformation, "Prompt Loaded Successfully"
    End If
End Sub

' Helper: Encode local image to base64
Function EncodeImageToBase64(filePath As String) As String
    Dim inputStream As Object
    Set inputStream = CreateObject("ADODB.Stream")

    inputStream.Type = 1 ' Binary
    inputStream.Open
    inputStream.LoadFromFile filePath

    Dim bytes() As Byte
    bytes = inputStream.Read(-1)
    inputStream.Close

    Dim dom As Object
    Set dom = CreateObject("Microsoft.XMLDOM")
    Dim element As Object
    Set element = dom.createElement("b64")

    element.DataType = "bin.base64"
    element.nodeTypedValue = bytes
    EncodeImageToBase64 = element.text
End Function

Function AskGPTWithImage(apiKey As String, Prompt As String, imagePath As String) As String
    AskGPTWithImage = CallChatGPT(apiKey, Prompt, imagePath)
End Function

Sub TestAskGPTWithImage()
    Dim Prompt As String
    Dim imagePath As String
    Dim apiKey As String
    Dim response As String

    Prompt = LoadPromptFromFile("GET_TBOX_PROJECT_INFO_FROM_SCREENSHOT")
    imagePath = GetOneDriveRoot & "\2526\Computers\TBox Projects\Screenshot_20-6-2025_73939_www.tboxplanet.com.jpeg"
    apiKey = GetEnvironmentVariable("OPENAI_API_KEY")

    response = AskGPTWithImage(apiKey, Prompt, imagePath)
    MsgBox response
End Sub

Sub GetTBoxProjectInfo()
    Dim Prompt As String
    Dim imagePath As String
    Dim apiKey As String
    Dim response As String

    Prompt = LoadPromptFromFile("GET_TBOX_PROJECT_INFO_FROM_SCREENSHOT")
    imagePath = GetOneDriveRoot & "\2526\Computers\TBox Projects\Screenshot_19-6-2025_142959_www.tboxplanet.com.jpeg"
    apiKey = GetEnvironmentVariable("OPENAI_API_KEY")

    response = AskGPTWithImage(apiKey, Prompt, imagePath)
    MsgBox response
End Sub

Sub ListUserFolders()
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder("C:\Users")

    For Each subfolder In folder.SubFolders
        Debug.Print subfolder.Name
    Next
End Sub

Function GetOneDriveRoot() As String
    ' 1. Commercial / school tenant
    GetOneDriveRoot = Environ$("OneDriveCommercial")
    
    ' 2. Personal OneDrive (fallback)
    If GetOneDriveRoot = "" Then GetOneDriveRoot = Environ$("OneDrive")
    
    ' 3. Last-chance check
    If GetOneDriveRoot = "" Then _
        Err.Raise vbObjectError + 1, , "Cannot locate OneDrive folder on this PC."
End Function


Function GetPromptsPath() As String
    Dim root As String
    root = GetOneDriveRoot
    
    GetPromptsPath = root & Application.PathSeparator & _
                     "2526" & Application.PathSeparator & _
                     "Tools" & Application.PathSeparator & _
                     "Prompts"
End Function

Function EscapeForJson(text As String) As String
    text = Replace(text, "\", "\\")
    text = Replace(text, """", "\""")
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    EscapeForJson = text
End Function

Sub WriteTextToFile(filePath As String, textToWrite As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    On Error GoTo ErrHandler
    Open filePath For Output As #fileNum
    Print #fileNum, textToWrite
    Close #fileNum
    Exit Sub

ErrHandler:
    MsgBox "Error writing file: " & Err.Description, vbExclamation
    If fileNum > 0 Then Close #fileNum
End Sub

Sub ProcessFolderImagesWithGPT()
    Dim folderPath As String
    Dim fDialog As FileDialog
    Dim fso As Object
    Dim file As Object
    Dim Prompt As String
    Dim apiKey As String
    Dim imagePath As String
    Dim response As String
    Dim successCount As Long
    Dim failureCount As Long
    Dim logText As String
    Dim ext As String

    ' Ask user to select a folder
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select Folder with Images"
        If .Show <> -1 Then
            MsgBox "Operation cancelled.", vbInformation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    ' Declare Sub Level Variables and Objects
    Dim Counter As Long
    Dim totalFiles As Long

    ' Initialize the Variables and Objects
    totalFiles = ContarArchivosImagen(folderPath)

    ' Declare the ProgressBar Object
    Dim MyProgressbar As ProgressBar
    Set MyProgressbar = New ProgressBar

    With MyProgressbar
        .Title = "Updating Gradebooks"
        .ExcelStatusBar = True
        .StartColour = rgbMediumSeaGreen
        .EndColour = rgbGreen
        .TotalActions = totalFiles
    End With

    MyProgressbar.ShowBar

    ' Load prompt and API key
    Prompt = LoadPromptFromFile("GET_TBOX_PROJECT_INFO_FROM_SCREENSHOT")
    apiKey = GetEnvironmentVariable("OPENAI_API_KEY")
    Set fso = CreateObject("Scripting.FileSystemObject")

    successCount = 0
    failureCount = 0
    logText = ""

    For Each file In fso.GetFolder(folderPath).Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "jpg" Or ext = "jpeg" Or ext = "png" Then
            MyProgressbar.NextAction "Processing '" & file.Name & "'", True

            Dim baseName As String
            baseName = fso.GetBaseName(file.Name)
            Dim txtPath As String
            txtPath = folderPath & "\" & baseName & ".txt"

            If Not fso.FileExists(txtPath) Then
                imagePath = file.path

                Do
                    response = AskGPTWithImage(apiKey, Prompt, imagePath)

                    If InStr(1, response, "tokens per min", vbTextCompare) > 0 Then
                        Debug.Print "Rate limit hit. Waiting 60 seconds..."
                        Application.Wait Now + TimeSerial(0, 1, 0) ' wait 1 minute
                    Else
                        Exit Do
                    End If
                Loop

                If IsValidGPTResponse(response) Then
                    SaveStringToFile response, txtPath
                    successCount = successCount + 1
                Else
                    logText = logText & baseName & " -> Invalid response: " & response & vbCrLf
                    failureCount = failureCount + 1
                End If
            End If
        End If
    Next file

    MyProgressbar.Complete

    Dim TotalCount As Long
    TotalCount = successCount + failureCount
    Dim successRate As Double
    If TotalCount > 0 Then
        successRate = successCount / TotalCount
        MsgBox "Processing complete." & vbCrLf & _
               "Success: " & successCount & vbCrLf & _
               "Failures: " & failureCount & vbCrLf & _
               "Success rate: " & Format(successRate, "0.00%"), vbInformation
    Else
        MsgBox "No images found to process.", vbExclamation
    End If

    If Len(logText) > 0 Then
        SaveStringToFile logText, folderPath & "\gpt_errors.log"
    End If
End Sub

Function SaveStringToFile(text As String, filePath As String)
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write text
    ts.Close
End Function

' -----------------------------
'  IsValidGPTResponse Function
' -----------------------------
' Return True only when:
'   1) response parses as JSON
'   2) the required keys exist
'   3) the value types look right (basic checks)

Public Function IsValidGPTResponse(ByVal response As String) As Boolean
    Const REQUIRED_KEYS As String = _
        "project_title,main_idea,academic_areas,digital_citizenship," & _
        "technology_tools,competencies,indicators,activities"

    On Error GoTo BadResponse            ' <- any error means “invalid”

    ' 1) Parse
    Dim j As Object                      ' Dictionary/Collection from JsonConverter
    Set j = JsonConverter.ParseJson(response)

    ' 2) Mandatory keys
    Dim keyArr() As String: keyArr = Split(REQUIRED_KEYS, ",")
    Dim k As Variant
    For Each k In keyArr
        If Not j.Exists(k) Then GoTo BadResponse
    Next k

    ' -- All checks passed
    IsValidGPTResponse = True
    Exit Function

BadResponse:
    IsValidGPTResponse = False
End Function


Function ContarArchivosImagen(ByVal carpeta As String) As Integer
    Dim archivo As String
    Dim contador As Integer
    
    ' Asegurarse de que la carpeta termine en "\"
    If Right(carpeta, 1) <> "\" Then carpeta = carpeta & "\"
    
    ' Contar archivos .jpg
    archivo = Dir(carpeta & "*.jpg")
    Do While archivo <> ""
        contador = contador + 1
        archivo = Dir
    Loop
    
    ' Contar archivos .jpeg
    archivo = Dir(carpeta & "*.jpeg")
    Do While archivo <> ""
        contador = contador + 1
        archivo = Dir
    Loop
    
    ' Contar archivos .png
    archivo = Dir(carpeta & "*.png")
    Do While archivo <> ""
        contador = contador + 1
        archivo = Dir
    Loop
    
    ContarArchivosImagen = contador
End Function

Sub ProcessSingleImageWithGPT()
    Dim fDialog As FileDialog
    Dim Prompt As String
    Dim apiKey As String
    Dim imagePath As String
    Dim response As String
    Dim txtPath As String
    Dim logText As String
    Dim fso As Object
    Dim ext As String
    Dim baseName As String

    Debug.Print "Starting ProcessSingleImageWithGPT..."

    ' Ask user to select a single image file
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select an Image File"
        .Filters.Clear
        .Filters.Add "Image Files", "*.jpg; *.jpeg; *.png"
        If .Show <> -1 Then
            MsgBox "Operation cancelled.", vbInformation
            Debug.Print "User cancelled file selection."
            Exit Sub
        End If
        imagePath = .SelectedItems(1)
    End With
    Debug.Print "Selected file: " & imagePath

    ' Get extension and validate
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase(fso.GetExtensionName(imagePath))
    Debug.Print "File extension: " & ext
    If ext <> "jpg" And ext <> "jpeg" And ext <> "png" Then
        MsgBox "Selected file is not a supported image format.", vbExclamation
        Debug.Print "Unsupported file type."
        Exit Sub
    End If

    ' Prepare paths
    baseName = fso.GetBaseName(imagePath)
    txtPath = fso.GetParentFolderName(imagePath) & "\" & baseName & ".txt"
    Debug.Print "Base name: " & baseName
    Debug.Print "Text path for response: " & txtPath

    ' Load prompt and API key
    Prompt = LoadPromptFromFile("GET_TBOX_PROJECT_INFO_FROM_SCREENSHOT")
    Debug.Print "Prompt loaded."
    apiKey = GetEnvironmentVariable("OPENAI_API_KEY")
    If apiKey = "" Then
        MsgBox "API key not found in environment variables.", vbCritical
        Debug.Print "Missing API key."
        Exit Sub
    End If
    Debug.Print "API key loaded."

    ' Skip if response already exists
    If fso.FileExists(txtPath) Then
        MsgBox "This image has already been processed.", vbInformation
        Debug.Print "Skipping file, response already exists."
        Exit Sub
    End If

    ' Call GPT
    Debug.Print "Sending image to GPT..."
    response = AskGPTWithImage(apiKey, Prompt, imagePath)
    Debug.Print "Response received: " & Left(response, 100) & "..."

    ' Process response
    If IsValidGPTResponse(response) Then
        SaveStringToFile response, txtPath
        MsgBox "Processing successful. Response saved.", vbInformation
        Debug.Print "Response saved successfully."
    Else
        logText = baseName & " -> Invalid response: " & response & vbCrLf
        SaveStringToFile logText, fso.GetParentFolderName(imagePath) & "\gpt_errors.log"
        MsgBox "Invalid response received. Logged to gpt_errors.log.", vbExclamation
        Debug.Print "Invalid response logged."
    End If
End Sub





Function ReadFileContents(filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    On Error GoTo ErrHandler

    With stream
        .charset = "utf-8"
        .Open
        .LoadFromFile filePath
        ReadFileContents = .ReadText(-1) ' Read all text
        .Close
    End With

    Set stream = Nothing
    Exit Function

ErrHandler:
    ReadFileContents = ""
    If Not stream Is Nothing Then
        If stream.State = 1 Then stream.Close
    End If
    Set stream = Nothing
End Function



Function ExtractGrade(fileName As String) As String
    Dim matches As Object
    With CreateObject("VBScript.RegExp")
        .pattern = "([A-Za-z]+)\s+Grade"
        .Global = False
        .IgnoreCase = True
        If .test(fileName) Then
            Set matches = .Execute(fileName)
            ExtractGrade = matches(0).SubMatches(0)
        Else
            ExtractGrade = ""
        End If
    End With
End Function

Function ReadTextAuto(filePath As String) As String
    Const adTypeBinary = 1
    Const adTypeText = 2
    
    Dim binSt As Object, txtSt As Object
    Dim bom     As Variant     ' <-- Variant, NOT fixed Byte array
    Dim charset As String
    
    '--- STEP 1 : read first few bytes to detect BOM ---------------------
    Set binSt = CreateObject("ADODB.Stream")
    With binSt
        .Type = adTypeBinary
        .Open
        .LoadFromFile filePath
        If .Size >= 4 Then
            bom = .Read(4)      'returns Variant(byte())
        Else
            bom = .Read         'smaller files
        End If
        .Close
    End With
    
    'Detect encoding
    charset = "utf-8"           'default
    If UBound(bom) >= 1 Then
        If bom(0) = &HEF And bom(1) = &HBB And bom(2) = &HBF Then
            charset = "utf-8"           'UTF-8 with BOM
        ElseIf bom(0) = &HFF And bom(1) = &HFE Then
            charset = "utf-16"          'UTF-16 LE
        ElseIf bom(0) = &HFE And bom(1) = &HFF Then
            charset = "utf-16"          'UTF-16 BE – ADODB handles BOM
        End If
    End If
    
    '--- STEP 2 : load whole file using the detected charset -------------
    Set txtSt = CreateObject("ADODB.Stream")
    With txtSt
        .Type = adTypeText
        .charset = charset
        .Open
        .LoadFromFile filePath
        ReadTextAuto = .ReadText(-1)    'all text
        .Close
    End With
End Function

'--------------------------------------------------------------------
' Reads a file (auto-detect UTF-8 / UTF-16), removes any Markdown
' code-fence lines  ``` or ```json, and returns a parsed JSON object.
' Requires:  ReadTextAuto  and  JsonConverter.ParseJson
'--------------------------------------------------------------------
Function ParseJsonFileClean(filePath As String) As Object
    Dim raw As String, cleaned As String
    
    '1) read with correct encoding
    raw = ReadTextAuto(filePath)
    
    '2) strip leading / trailing fence lines
    Dim line As Variant, buffer As String
    For Each line In Split(raw, vbCrLf)
        If Trim$(line) <> "```" And _
           LCase$(Trim$(line)) <> "```json" Then
            buffer = buffer & line & vbCrLf
        End If
    Next line
    cleaned = Trim$(buffer)
    
    '3) parse
    Set ParseJsonFileClean = JsonConverter.ParseJson(cleaned)
End Function

'--------------------------------------------------------------------
' CleanJsonCodeFence
'   Removes the leading / trailing ``` code-fence that ChatGPT may add
'   around a JSON payload.  Works even when the fence is ```json.
'
'   Example:
'       Dim raw$, clean$
'       raw = "```json" & vbCrLf & "{""a"":1}" & vbCrLf & "```"
'       clean = CleanJsonCodeFence(raw)   '? {"a":1}
'--------------------------------------------------------------------
Public Function CleanJsonCodeFence(ByVal txt As String) As String
    Dim work$, first3$, last3$
    
    work = Trim$(txt)                     'Strip outer whitespace first
    If Len(work) < 6 Then                 'Shorter than "```x```" ? nothing to do
        CleanJsonCodeFence = work
        Exit Function
    End If
    
    first3 = Left$(work, 3)
    last3 = Right$(work, 3)
    
    '--- remove opening fence -------------------------------------------------
    If first3 = "```" Then
        work = Mid$(work, 4)              'drop the ```
        'If the fence is ```json (or ```text, etc.) skip the language tag
        If InStr(1, work, vbLf, vbBinaryCompare) > 0 Then
            Dim firstLfPos As Long
            firstLfPos = InStr(1, work, vbLf, vbBinaryCompare)
            work = Mid$(work, firstLfPos + 1)
        End If
        work = Trim$(work)
    End If
    
    '--- remove closing fence --------------------------------------------------
    If last3 = "```" Then work = Left$(work, Len(work) - 3)
    
    CleanJsonCodeFence = Trim$(work)
End Function



