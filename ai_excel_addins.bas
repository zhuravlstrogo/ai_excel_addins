' =============================================================================
'  AI for Microsoft Excel (Windows) via corporate LiteLLM  -  version 3, one worksheet function
'
'  WORKSHEET FUNCTION (results are cached for the session - no repeated paid calls)
'    =LLM(prompt; [data]; [instructions]; [model]; [temperature]; [max_tokens])
'
'  MACROS
'    LLM_Chat              - chat window with agent mode (Ctrl+Shift+K)
'    LLM_ProcessSelection  - one prompt for every selected cell (Ctrl+Shift+J)
'    LLM_Undo              - undo the last AI changes
'    LLM_ClearCache        - forget cached answers of worksheet functions
'
'  The chat window needs UserForm "frmLLMChat" (code in frmLLMChat_code_v3.txt).
'  API key: entered by the user in chat window -> Settings (stored in the registry, never in this file)
'  Other settings: chat window -> Settings (stored in the Windows registry).
'  User-facing Russian text is stored as \uXXXX escapes and decoded by RU(),
'  because the VBA editor breaks Cyrillic on import on some Windows locales.
' =============================================================================
Option Explicit

' ---- Defaults (used until changed in chat window -> Settings) ---------------
Private Const DEF_URL As String = "https://litellm.geotech.user.events/v1/chat/completions"
Private Const DEF_MODEL As String = "openai/gpt-5.6-terra"
Private Const DEF_MODEL_LIST As String = "gti-local/Qwen3.8,geotech/zai/glm-5.3"
Private Const DEF_LANGUAGE As String = "Russian"
Private Const DEF_CHAT_MAX_TOKENS As String = "16000"
Private Const DEF_CELL_MAX_TOKENS As String = "1000"
' Rows per sheet in the "workbook overview" context (0 = no row limit, only MAX_CONTEXT_CHARS)
Private Const DEF_OVERVIEW_ROWS As String = "500"
' "1" = the agent writes plain values instead of lookup formulas (see SystemPrompt)
Private Const DEF_VALUES_NOT_FORMULAS As String = "0"
' -----------------------------------------------------------------------------
Private Const APP_NAME As String = "ExcelLiteLLM"
Private Const MAX_CONTEXT_CHARS As Long = 100000
Private Const MAX_HISTORY_MESSAGES As Long = 20
' Hard cap on the answer text. Must stay well above ChatMaxTokens: Cyrillic runs
' 2-3 characters per token, and a cut answer means a cut ```actions block.
Private Const MAX_ANSWER_CHARS As Long = 200000
Private Const MAX_UNDO_CELLS As Long = 200000
Private Const MAX_UNDO_FORMAT_CELLS As Long = 5000
' A reasoning model writing a few thousand tokens can easily need more than 3 minutes
Private Const TIMEOUT_MS As Long = 600000

Public LastAnswer As String
Public LastUsage As Long
' True when the model hit max_tokens: the answer (and any actions block) is incomplete
Public LastTruncated As Boolean
' Token counters of the last call, shown when an answer is cut off
Public LastMaxTokens As Long
Public LastPromptTokens As Long
Public LastCompletionTokens As Long
Public LastReasoningTokens As Long

Private mRoles() As String
Private mTexts() As String
Private mCount As Long
Private mChat As Object
Private mCache As Object
Private mUndo As Collection


' ============================== SETTINGS =====================================

Public Function Setting(ByVal key As String) As String
    Dim def As String
    Select Case key
        Case "URL": def = DEF_URL
        Case "Model": def = DEF_MODEL
        Case "ModelList": def = DEF_MODEL_LIST
        Case "Language": def = DEF_LANGUAGE
        Case "ChatMaxTokens": def = DEF_CHAT_MAX_TOKENS
        Case "CellMaxTokens": def = DEF_CELL_MAX_TOKENS
        Case "OverviewRows": def = DEF_OVERVIEW_ROWS
        Case "ValuesNotFormulas": def = DEF_VALUES_NOT_FORMULAS
    End Select
    Setting = GetSetting(APP_NAME, "Settings", key, def)
    If Len(Setting) = 0 And Len(def) > 0 Then Setting = def
End Function

Public Sub SaveSettingValue(ByVal key As String, ByVal value As String)
    SaveSetting APP_NAME, "Settings", key, value
End Sub

Private Function SettingTemp() As Double
    Dim s As String
    s = Trim$(Setting("Temperature"))
    If Len(s) = 0 Then SettingTemp = -1 Else SettingTemp = Val(Replace(s, ",", "."))
End Function

Private Function SettingLong(ByVal key As String) As Long
    SettingLong = CLng(Val(Setting(key)))
End Function


' =========================== WORKSHEET FUNCTIONS =============================

Public Function LLM(ByVal prompt As String, Optional ByVal data As Variant, _
                    Optional ByVal instructions As String = "", Optional ByVal model As String = "", _
                    Optional ByVal temperature As Double = -1, Optional ByVal maxTokens As Long = 0) As Variant
    Dim usr As String, ctx As String
    usr = prompt
    If Not IsMissing(data) Then
        ctx = DataArgToText(data)
        If Len(ctx) > 0 Then usr = usr & vbLf & vbLf & "Data:" & vbLf & ctx
    End If
    If Len(instructions) > 0 Then instructions = "Additional instructions: " & instructions
    LLM = CellCall(usr, instructions, model, temperature, maxTokens)
End Function

Private Function CellCall(ByVal usr As String, ByVal extraSys As String, ByVal model As String, _
                          ByVal temperature As Double, ByVal maxTokens As Long) As String
    Dim sys As String
    sys = SystemPrompt(False)
    sys = sys & " Your answer will be placed into ONE Excel cell: keep it short, plain text, no markdown."
    If Len(extraSys) > 0 Then sys = sys & vbLf & extraSys
    If maxTokens <= 0 Then maxTokens = SettingLong("CellMaxTokens")
    If temperature < 0 Then temperature = SettingTemp()
    CellCall = CachedRequest("[" & Msg("system", sys) & "," & Msg("user", usr) & "]", model, maxTokens, temperature)
End Function

Private Function CachedRequest(ByVal json As String, ByVal model As String, _
                               ByVal maxTokens As Long, ByVal temperature As Double) As String
    Dim key As String
    If mCache Is Nothing Then Set mCache = CreateObject("Scripting.Dictionary")
    key = model & "|" & maxTokens & "|" & temperature & "|" & json
    If mCache.Exists(key) Then
        CachedRequest = mCache.Item(key)
        Exit Function
    End If
    CachedRequest = RequestLLM(json, model, maxTokens, temperature)
    If Left$(CachedRequest, 6) <> "#ERROR" Then mCache.Item(key) = CachedRequest
End Function

Private Function DataArgToText(ByVal data As Variant) As String
    If TypeName(data) = "Range" Then
        If data.Cells.CountLarge = 1 Then
            DataArgToText = CellText(data.Value)
        Else
            DataArgToText = BuildContext(data, False, False)
        End If
    Else
        DataArgToText = ArrayToText(data)
    End If
End Function


' ================================= MACROS ====================================

Public Sub LLM_Chat()
    On Error GoTo NoForm
    If mChat Is Nothing Then Set mChat = VBA.UserForms.Add("frmLLMChat")
    mChat.Show vbModeless
    Exit Sub
NoForm:
    Set mChat = Nothing
    MsgBox RU("\u0412 \u043D\u0430\u0434\u0441\u0442\u0440\u043E\u0439\u043A\u0435 \u043D\u0435\u0442 \u0444\u043E\u0440\u043C\u044B frmLLMChat (\u0441\u043C. \u0438\u043D\u0441\u0442\u0440\u0443\u043A\u0446\u0438\u044E \u043F\u043E \u0443\u0441\u0442\u0430\u043D\u043E\u0432\u043A\u0435).") & vbLf & _
           Err.Description, vbExclamation, RU("\u0418\u0418")
End Sub

Public Sub LLM_Undo()
    MsgBox UndoLast(), vbInformation, RU("\u0418\u0418")
End Sub

Public Sub LLM_ClearCache()
    MsgBox ClearCache(), vbInformation, RU("\u0418\u0418")
End Sub

Public Function ClearCache() As String
    Dim n As Long
    If Not mCache Is Nothing Then n = mCache.Count
    Set mCache = Nothing
    ClearCache = RU("\u041A\u044D\u0448 \u043E\u0447\u0438\u0449\u0435\u043D (\u043E\u0442\u0432\u0435\u0442\u043E\u0432: ") & n & RU("). \u041D\u0430\u0436\u043C\u0438\u0442\u0435 Ctrl+Alt+F9 \u0434\u043B\u044F \u043F\u0435\u0440\u0435\u0441\u0447\u0451\u0442\u0430.")
End Function

Public Sub LLM_ProcessSelection()
    If TypeName(Selection) <> "Range" Then Exit Sub
    ' answers go to the column on the right: with several columns they would overwrite the input
    If Selection.Areas.Count > 1 Or Selection.Columns.Count > 1 Then
        MsgBox RU("\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u044F\u0447\u0435\u0439\u043A\u0438 \u0432 \u043E\u0434\u043D\u043E\u043C \u0441\u0442\u043E\u043B\u0431\u0446\u0435. \u041E\u0442\u0432\u0435\u0442\u044B \u0437\u0430\u043F\u0438\u0441\u044B\u0432\u0430\u044E\u0442\u0441\u044F \u0432 \u0441\u043E\u0441\u0435\u0434\u043D\u0438\u0439 \u0441\u0442\u043E\u043B\u0431\u0435\u0446 \u0441\u043F\u0440\u0430\u0432\u0430, ") & _
               RU("\u043F\u043E\u044D\u0442\u043E\u043C\u0443 \u043F\u0440\u0438 \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u0438\u0445 \u0441\u0442\u043E\u043B\u0431\u0446\u0430\u0445 \u043E\u043D\u0438 \u0437\u0430\u0442\u0440\u0443\u0442 \u0438\u0441\u0445\u043E\u0434\u043D\u044B\u0435 \u0434\u0430\u043D\u043D\u044B\u0435."), vbExclamation, RU("\u0418\u0418")
        Exit Sub
    End If
    If Selection.Column >= Selection.Worksheet.Columns.Count Then Exit Sub

    Dim target As Range
    Set target = Intersect(Selection, Selection.Worksheet.UsedRange)
    If target Is Nothing Then Exit Sub

    Dim prompt As String
    prompt = InputBox(RU("\u0417\u0430\u043F\u0440\u043E\u0441 \u0434\u043B\u044F \u043A\u0430\u0436\u0434\u043E\u0439 \u0432\u044B\u0434\u0435\u043B\u0435\u043D\u043D\u043E\u0439 \u044F\u0447\u0435\u0439\u043A\u0438.") & vbLf & _
                      RU("\u041E\u0442\u0432\u0435\u0442 \u0437\u0430\u043F\u0438\u0448\u0435\u0442\u0441\u044F \u0432 \u0441\u043E\u0441\u0435\u0434\u043D\u0438\u0439 \u0441\u0442\u043E\u043B\u0431\u0435\u0446 \u0441\u043F\u0440\u0430\u0432\u0430."), RU("\u0418\u0418"))
    If Len(prompt) = 0 Then Exit Sub

    Dim total As Long, n As Long, cell As Range, v As String, sys As String, batch As New Collection
    total = target.Cells.Count
    If total > 100 Then
        If MsgBox(RU("\u0412\u044B\u0434\u0435\u043B\u0435\u043D\u043E \u044F\u0447\u0435\u0435\u043A: ") & total & RU(" - \u044D\u0442\u043E \u0434\u043E ") & total & RU(" \u0437\u0430\u043F\u0440\u043E\u0441\u043E\u0432. \u041F\u0440\u043E\u0434\u043E\u043B\u0436\u0438\u0442\u044C?"), _
                  vbYesNo + vbQuestion, RU("\u0418\u0418")) <> vbYes Then Exit Sub
    End If

    SnapshotCells target.Offset(0, 1), batch, False
    PushUndo batch

    sys = SystemPrompt(False) & " Your answer will be placed into ONE Excel cell: keep it short, plain text, no markdown."
    For Each cell In target.Cells
        n = n + 1
        Application.StatusBar = "AI: " & n & " / " & total
        v = CellText(cell.Value)
        If Len(v) > 0 Then
            cell.Offset(0, 1).Value = SafeCellValue(RequestLLM("[" & Msg("system", sys) & "," & _
                Msg("user", prompt & vbLf & vbLf & "Data:" & vbLf & v) & "]", "", _
                SettingLong("CellMaxTokens"), SettingTemp()))
        End If
        DoEvents
    Next cell

    Application.StatusBar = False
    MsgBox RU("\u0413\u043E\u0442\u043E\u0432\u043E, \u043E\u0431\u0440\u0430\u0431\u043E\u0442\u0430\u043D\u043E \u044F\u0447\u0435\u0435\u043A: ") & n & RU(". \u041E\u0442\u043C\u0435\u043D\u0438\u0442\u044C: \u043A\u043D\u043E\u043F\u043A\u0430 \u00AB\u041E\u0442\u043C\u0435\u043D\u0438\u0442\u044C \u043F\u0440\u0430\u0432\u043A\u0438 \u0418\u0418\u00BB \u0432 \u043E\u043A\u043D\u0435 \u0447\u0430\u0442\u0430."), _
           vbInformation, RU("\u0418\u0418")
End Sub


' ============================ CHAT (used by form) ============================

Public Function ChatSend(ByVal userText As String, ByVal context As String, _
                         ByVal model As String, ByVal agentMode As Boolean) As String
    Dim content As String, json As String, ans As String, i As Long
    content = userText
    If Len(context) > 0 Then content = content & vbLf & vbLf & "Excel data:" & vbLf & context
    DropOldContexts

    json = "[" & Msg("system", SystemPrompt(agentMode))
    For i = 1 To mCount
        json = json & "," & Msg(mRoles(i), mTexts(i))
    Next i
    json = json & "," & Msg("user", content) & "]"

    ans = RequestLLM(json, model, SettingLong("ChatMaxTokens"), SettingTemp())
    If Left$(ans, 6) <> "#ERROR" Then
        AddHistory "user", content
        AddHistory "assistant", ans
        LastAnswer = ans
    End If
    ChatSend = ans
End Function

Public Sub ChatReset()
    mCount = 0
    Erase mRoles
    Erase mTexts
    LastAnswer = ""
End Sub

Public Function MakeFormula(ByVal description As String, ByVal context As String, _
                            ByVal targetAddress As String, ByVal model As String) As String
    Dim sys As String, usr As String, ans As String
    Dim lines() As String, i As Long, ln As String

    sys = "You write Microsoft Excel formulas. Reply with exactly ONE formula that starts with =, " & _
          "using English function names and commas as argument separators. No explanations, no markdown."
    usr = "Task: " & description & vbLf & "The formula will be placed in cell " & targetAddress & "."
    If Len(context) > 0 Then usr = usr & vbLf & vbLf & "Sheet data (sample):" & vbLf & context

    ans = RequestLLM("[" & Msg("system", sys) & "," & Msg("user", usr) & "]", model, 500)
    If Left$(ans, 6) = "#ERROR" Then
        MakeFormula = ans
        Exit Function
    End If

    lines = Split(Replace(ans, vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(Replace(lines(i), "`", ""))
        If Left$(ln, 1) = "=" Then
            MakeFormula = ln
            Exit Function
        End If
    Next i
    MakeFormula = ans
End Function

Public Function InsertFormula(ByVal target As Range, ByVal f As String) As Boolean
    Dim batch As New Collection
    SnapshotCells target, batch, False
    PushUndo batch
    On Error Resume Next
    SetFormulaSafe target, f
    InsertFormula = (Err.Number = 0)
End Function

Public Sub InsertTextToCells(ByVal text As String, ByVal target As Range)
    Dim grid As Variant, dest As Range, batch As New Collection
    If Len(text) = 0 Then Exit Sub
    grid = TextToGrid(text, False)
    If IsArray(grid) Then
        Set dest = target.Cells(1, 1).Resize(UBound(grid, 1), UBound(grid, 2))
    Else
        Set dest = target.Cells(1, 1)
    End If
    If Not ConfirmOverwrite(dest) Then Exit Sub
    SnapshotCells dest, batch, False
    PushUndo batch
    If IsArray(grid) Then
        dest.Value = grid
    Else
        dest.Value = SafeCellValue(Left$(Replace(text, vbCr, ""), 32000))
    End If
End Sub

Public Sub SaveTranscript(ByVal text As String)
    Dim ws As Worksheet, lines() As String, arr() As Variant, i As Long, n As Long
    If Len(text) = 0 Then Exit Sub
    lines = Split(Replace(text, vbCr, ""), vbLf)
    n = UBound(lines) + 1
    ReDim arr(1 To n, 1 To 1)
    For i = 1 To n
        arr(i, 1) = SafeCellValue(Left$(lines(i - 1), 32000))
    Next i
    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    On Error Resume Next
    ws.Name = RU("\u0427\u0430\u0442 \u0418\u0418 ") & Format$(Now, "hhmmss")
    On Error GoTo 0
    ws.Range("A1").Resize(n, 1).Value = arr
    ws.Columns(1).ColumnWidth = 120
    ws.Columns(1).WrapText = True
End Sub

Public Function WorkbookInfoLine() As String
    Dim ws As Worksheet, names As String
    On Error Resume Next
    For Each ws In ActiveWorkbook.Worksheets
        If Len(names) > 0 Then names = names & ", "
        names = names & ws.Name
    Next ws
    WorkbookInfoLine = "Workbook: " & ActiveWorkbook.Name & " | Active sheet: " & ActiveSheet.Name & _
                       " | Active cell: " & ActiveCell.Address(False, False) & " | Sheets: " & names
End Function

Public Function BuildWorkbookOverview() As String
    Dim ws As Worksheet, buf As String, used As Long, maxRows As Long
    maxRows = SettingLong("OverviewRows")
    If maxRows < 0 Then maxRows = 0
    For Each ws In ActiveWorkbook.Worksheets
        SbAdd buf, used, BuildContext(DataRange(ws), True, True, maxRows) & vbLf
        If used > MAX_CONTEXT_CHARS Then
            SbAdd buf, used, "...[context limit reached, remaining sheets not shown]" & vbLf
            Exit For
        End If
    Next ws
    BuildWorkbookOverview = Left$(buf, used)
End Function

' UsedRange without leading empty rows/columns (UsedRange often starts at formatted but empty cells)
Private Function DataRange(ByVal ws As Worksheet) As Range
    Dim first As Range, lastCell As Range, firstCol As Range
    Set DataRange = ws.UsedRange
    On Error GoTo Done
    Set lastCell = ws.UsedRange.Cells(ws.UsedRange.Cells.CountLarge)
    Set first = ws.UsedRange.Find("*", After:=lastCell, LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Set firstCol = ws.UsedRange.Find("*", After:=lastCell, LookIn:=xlFormulas, _
                                     SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    If first Is Nothing Or firstCol Is Nothing Then Exit Function
    Set DataRange = ws.Range(ws.Cells(first.Row, firstCol.Column), lastCell)
Done:
End Function


' ============================ AGENT: ACTIONS =================================

' Splits the model answer into display text and a collection of actions (or Nothing).
Public Function SplitActions(ByVal answer As String, ByRef displayText As String) As Collection
    Dim p1 As Long, p2 As Long, jsonText As String, v As Variant, p As Long, tail As String, tag As String
    displayText = answer
    tag = "```actions"
    p1 = InStr(1, answer, tag, vbTextCompare)
    If p1 = 0 Then
        ' models sometimes label the block ```json - accept it if it contains actions
        tag = "```json"
        p1 = InStr(1, answer, tag, vbTextCompare)
        If p1 > 0 Then
            p2 = InStr(p1 + Len(tag), answer, "```")
            If p2 = 0 Then p2 = Len(answer) + 1
            If InStr(Mid$(answer, p1, p2 - p1), """action""") = 0 Then p1 = 0
        End If
    End If
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + Len(tag), answer, "```")
    If p2 = 0 Then
        ' no closing fence: the answer was cut off mid-block, the actions are incomplete
        displayText = Trim$(Left$(answer, p1 - 1)) & vbLf & TruncatedNote()
        Exit Function
    End If
    jsonText = Mid$(answer, p1 + Len(tag), p2 - p1 - Len(tag))
    If p2 + 3 <= Len(answer) Then tail = Mid$(answer, p2 + 3)
    displayText = Trim$(Left$(answer, p1 - 1) & tail)

    If LastTruncated Then
        displayText = displayText & vbLf & TruncatedNote()
        Exit Function
    End If

    On Error GoTo Bad
    p = 1
    JsonValue jsonText, p, v
    If IsObject(v) Then
        If TypeName(v) = "Collection" Then
            Set SplitActions = v
        ElseIf TypeName(v) = "Dictionary" Then
            Set SplitActions = New Collection
            SplitActions.Add v
        End If
    End If
    Exit Function
Bad:
    Set SplitActions = Nothing
    displayText = displayText & vbLf & RU("[\u043D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u043F\u0440\u043E\u0447\u0438\u0442\u0430\u0442\u044C \u043F\u0440\u0435\u0434\u043B\u043E\u0436\u0435\u043D\u043D\u044B\u0435 \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u044F]")
End Function

Private Function TruncatedNote() As String
    Dim d As String
    d = RU("[!] \u043E\u0442\u0432\u0435\u0442 \u043C\u043E\u0434\u0435\u043B\u0438 \u043E\u0431\u043E\u0440\u0432\u0430\u043D \u043F\u043E \u043B\u0438\u043C\u0438\u0442\u0443 \u0442\u043E\u043A\u0435\u043D\u043E\u0432, \u043F\u0440\u0430\u0432\u043A\u0438 \u041D\u0415 \u043F\u0440\u0438\u043C\u0435\u043D\u0435\u043D\u044B.")
    d = d & vbLf & RU("      \u043B\u0438\u043C\u0438\u0442 max_tokens: ") & LastMaxTokens & _
            RU(", \u043C\u043E\u0434\u0435\u043B\u044C \u043D\u0430\u043F\u0438\u0441\u0430\u043B\u0430: ") & LastCompletionTokens
    If LastReasoningTokens > 0 Then d = d & RU(" (\u0438\u0437 \u043D\u0438\u0445 \u0440\u0430\u0441\u0441\u0443\u0436\u0434\u0435\u043D\u0438\u044F: ") & LastReasoningTokens & ")"
    d = d & RU(", \u0432 \u0437\u0430\u043F\u0440\u043E\u0441 \u0443\u0448\u043B\u043E: ") & LastPromptTokens
    d = d & vbLf & RU("      \u0415\u0441\u043B\u0438 \u0432 \u0437\u0430\u043F\u0440\u043E\u0441 \u0443\u0448\u043B\u043E \u043E\u0447\u0435\u043D\u044C \u043C\u043D\u043E\u0433\u043E - \u043D\u0430\u0436\u043C\u0438\u0442\u0435 \u00AB\u041D\u043E\u0432\u044B\u0439 \u0447\u0430\u0442\u00BB \u0438\u043B\u0438 \u0443\u043C\u0435\u043D\u044C\u0448\u0438\u0442\u0435 \u00AB\u0421\u0442\u0440\u043E\u043A \u043D\u0430 \u043B\u0438\u0441\u0442 (\u043E\u0431\u0437\u043E\u0440)\u00BB. \u0418\u043D\u0430\u0447\u0435 \u043F\u043E\u0434\u043D\u0438\u043C\u0438\u0442\u0435 \u00AB\u041C\u0430\u043A\u0441. \u0442\u043E\u043A\u0435\u043D\u043E\u0432 \u0432 \u0447\u0430\u0442\u0435\u00BB.")
    TruncatedNote = d
End Function

Public Function ActionsPreview(ByVal actions As Collection) As String
    Dim a As Variant, n As Long, s As String
    For Each a In actions
        n = n + 1
        If n > 25 Then
            s = s & vbLf & RU("... \u0438 \u0435\u0449\u0451 ") & (actions.Count - 25)
            Exit For
        End If
        If IsObject(a) Then s = s & vbLf & n & ". " & DescribeAction(a, True)
    Next a
    ActionsPreview = s
End Function

Public Function ApplyActions(ByVal actions As Collection) As String
    Dim a As Variant, batch As New Collection, report As String, errs As String, ok As Long
    Dim startSheet As Object
    Set startSheet = ActiveSheet
    Application.ScreenUpdating = False
    For Each a In actions
        If IsObject(a) Then
            If TypeName(a) = "Dictionary" Then
                On Error Resume Next
                Err.Clear
                ApplyOne a, batch
                If Err.Number <> 0 Then
                    errs = errs & vbLf & "  ! " & DescribeAction(a, False) & ": " & Err.Description
                Else
                    ok = ok + 1
                    report = report & vbLf & "  + " & DescribeAction(a, False)
                End If
                On Error GoTo 0
            End If
        End If
    Next a
    On Error Resume Next
    startSheet.Activate
    On Error GoTo 0
    Application.ScreenUpdating = True
    PushUndo batch
    ApplyActions = RU("\u041F\u0440\u0438\u043C\u0435\u043D\u0435\u043D\u043E \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u0439: ") & ok & RU(" \u0438\u0437 ") & actions.Count & ":" & report & errs & _
                   vbLf & RU("(\u043A\u043D\u043E\u043F\u043A\u0430 \u00AB\u041E\u0442\u043C\u0435\u043D\u0438\u0442\u044C \u043F\u0440\u0430\u0432\u043A\u0438 \u0418\u0418\u00BB \u0432\u0435\u0440\u043D\u0451\u0442 \u043A\u0430\u043A \u0431\u044B\u043B\u043E)")
End Function

Private Sub ApplyOne(ByVal a As Object, ByVal batch As Collection)
    Dim act As String, rng As Range, grid As Variant, ws As Worksheet, src As Range, tgt As Range
    Dim co As Object, keyCell As Range, ord As Long, hdr As Long, prev As Object, u As Object

    act = LCase$(DictStr(a, "action"))
    Select Case act
        Case "set_values"
            grid = ValuesToGrid(DictItem(a, "values"))
            Set rng = ResolveRange(DictStr(a, "range"))
            CheckValuesFit rng, grid
            Set rng = rng.Cells(1, 1).Resize(UBound(grid, 1), UBound(grid, 2))
            CheckUndoSize rng
            SnapshotCells rng, batch, False
            SetFormulaSafe rng, grid

        Case "set_formula"
            Set rng = ResolveRange(DictStr(a, "range"))
            CheckUndoSize rng
            SnapshotCells rng, batch, False
            SetFormulaSafe rng, DictStr(a, "formula")

        Case "clear"
            ' cells outside UsedRange are already empty: "A:A" becomes a small range that can be undone
            Set rng = ResolveRange(DictStr(a, "range"))
            Set rng = Intersect(rng, rng.Worksheet.UsedRange)
            If Not rng Is Nothing Then
                CheckUndoSize rng
                SnapshotCells rng, batch, False
                rng.ClearContents
            End If

        Case "format"
            Set rng = ResolveRange(DictStr(a, "range"))
            CheckUndoSize rng
            SnapshotCells rng, batch, True
            ApplyFormat rng, a

        Case "sort"
            Set rng = ResolveRange(DictStr(a, "range"))
            Set keyCell = Intersect(rng, rng.Worksheet.Columns(DictStr(a, "key")))
            If keyCell Is Nothing Then Err.Raise vbObjectError + 3, , RU("\u0441\u0442\u043E\u043B\u0431\u0435\u0446 \u0441\u043E\u0440\u0442\u0438\u0440\u043E\u0432\u043A\u0438 \u0432\u043D\u0435 \u0434\u0438\u0430\u043F\u0430\u0437\u043E\u043D\u0430")
            CheckUndoSize rng
            SnapshotCells rng, batch, False
            If LCase$(DictStr(a, "order")) = "desc" Then ord = xlDescending Else ord = xlAscending
            hdr = xlYes
            If a.Exists("header") Then
                If Not CBool(a.Item("header")) Then hdr = xlNo
            End If
            rng.Sort Key1:=keyCell.Cells(1, 1), Order1:=ord, Header:=hdr

        Case "add_sheet"
            Set prev = ActiveSheet
            Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
            Set u = CreateObject("Scripting.Dictionary")
            u.Item("type") = "sheet"
            Set u.Item("obj") = ws
            batch.Add u
            If Len(DictStr(a, "name")) > 0 Then ws.Name = Left$(DictStr(a, "name"), 31)
            prev.Activate

        Case "chart"
            Set src = ResolveRange(DictStr(a, "source"))
            If Len(DictStr(a, "target")) > 0 Then
                Set tgt = ResolveRange(DictStr(a, "target"))
            Else
                Set tgt = src.Cells(1, src.Columns.Count).Offset(0, 2)
            End If
            Set co = tgt.Worksheet.ChartObjects.Add(tgt.Left, tgt.Top, 420, 250)
            Set u = CreateObject("Scripting.Dictionary")
            u.Item("type") = "chart"
            Set u.Item("obj") = co
            batch.Add u
            co.Chart.SetSourceData Source:=src
            co.Chart.ChartType = ChartTypeFrom(DictStr(a, "type"))
            If Len(DictStr(a, "title")) > 0 Then
                co.Chart.HasTitle = True
                co.Chart.ChartTitle.Text = DictStr(a, "title")
            End If

        Case Else
            Err.Raise vbObjectError + 1, , RU("\u043D\u0435\u0438\u0437\u0432\u0435\u0441\u0442\u043D\u043E\u0435 \u0434\u0435\u0439\u0441\u0442\u0432\u0438\u0435 '") & act & "'"
    End Select
End Sub

' Declared a real block of cells but supplied fewer/more values: almost always a cut-off
' or lazy answer, so refuse instead of silently filling only the first rows.
Private Sub CheckValuesFit(ByVal rng As Range, ByVal grid As Variant)
    Dim nR As Long, nC As Long
    If rng.Cells.CountLarge = 1 Then Exit Sub          ' anchor cell: array size wins
    nR = UBound(grid, 1): nC = UBound(grid, 2)
    If nR = rng.Rows.Count And nC = rng.Columns.Count Then Exit Sub
    Err.Raise vbObjectError + 4, , RU("values \u043D\u0435 \u0441\u043E\u0432\u043F\u0430\u0434\u0430\u0435\u0442 \u0441 \u0434\u0438\u0430\u043F\u0430\u0437\u043E\u043D\u043E\u043C: ") & _
              rng.Address(False, False) & " " & rng.Rows.Count & "x" & rng.Columns.Count & _
              RU(", \u0430 \u0437\u043D\u0430\u0447\u0435\u043D\u0438\u0439 - ") & nR & "x" & nC
End Sub

' Changes that could not be undone are not applied at all
Private Sub CheckUndoSize(ByVal rng As Range)
    If rng.Cells.CountLarge > MAX_UNDO_CELLS Then
        Err.Raise vbObjectError + 6, , RU("\u043D\u0435 \u043F\u0440\u0438\u043C\u0435\u043D\u0435\u043D\u043E: \u0434\u0438\u0430\u043F\u0430\u0437\u043E\u043D \u0431\u043E\u043B\u044C\u0448\u0435 ") & MAX_UNDO_CELLS & _
                  RU(" \u044F\u0447\u0435\u0435\u043A, \u0442\u0430\u043A\u043E\u0435 \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u0435 \u043D\u0435\u043B\u044C\u0437\u044F \u0431\u044B\u043B\u043E \u0431\u044B \u043E\u0442\u043C\u0435\u043D\u0438\u0442\u044C")
    End If
End Sub

Private Sub ApplyFormat(ByVal rng As Range, ByVal a As Object)
    Dim b As Variant
    If a.Exists("bold") Then rng.Font.Bold = CBool(a.Item("bold"))
    If a.Exists("italic") Then rng.Font.Italic = CBool(a.Item("italic"))
    If a.Exists("font_color") Then rng.Font.Color = HexColor(DictStr(a, "font_color"))
    If a.Exists("fill") Then
        If LCase$(DictStr(a, "fill")) = "none" Then
            rng.Interior.ColorIndex = xlNone
        Else
            rng.Interior.Color = HexColor(DictStr(a, "fill"))
        End If
    End If
    If a.Exists("number_format") Then rng.NumberFormat = DictStr(a, "number_format")
    If a.Exists("wrap") Then rng.WrapText = CBool(a.Item("wrap"))
    If a.Exists("align") Then
        Select Case LCase$(DictStr(a, "align"))
            Case "left": rng.HorizontalAlignment = xlLeft
            Case "center": rng.HorizontalAlignment = xlCenter
            Case "right": rng.HorizontalAlignment = xlRight
        End Select
    End If
    If a.Exists("borders") Then
        On Error Resume Next
        For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
            If CBool(a.Item("borders")) Then
                rng.Borders(b).LineStyle = xlContinuous
                rng.Borders(b).Weight = xlThin
            Else
                rng.Borders(b).LineStyle = xlNone
            End If
        Next b
        On Error GoTo 0
    End If
    If a.Exists("autofit") Then
        If CBool(a.Item("autofit")) Then rng.EntireColumn.AutoFit
    End If
End Sub

' detailed = True: for the confirmation dialog (values, formulas, undo warnings)
Private Function DescribeAction(ByVal a As Object, ByVal detailed As Boolean) As String
    Dim act As String, d As String, f As String, n As Double
    act = LCase$(DictStr(a, "action"))
    Select Case act
        Case "set_values"
            d = RU("\u0437\u0430\u043F\u0438\u0441\u0430\u0442\u044C \u0432 ") & ValuesTarget(a)
            If detailed Then d = d & ValuesPreview(DictItem(a, "values"))
        Case "set_formula"
            f = DictStr(a, "formula")
            d = RU("\u0444\u043E\u0440\u043C\u0443\u043B\u0430 ") & f & " -> " & DictStr(a, "range")
            If detailed And IsRiskyFormula(f) Then d = d & vbLf & RiskyNote()
        Case "format":      d = RU("\u043E\u0444\u043E\u0440\u043C\u0438\u0442\u044C ") & DictStr(a, "range")
        Case "clear":       d = RU("\u043E\u0447\u0438\u0441\u0442\u0438\u0442\u044C ") & DictStr(a, "range")
        Case "sort":        d = RU("\u0441\u043E\u0440\u0442\u0438\u0440\u043E\u0432\u0430\u0442\u044C ") & DictStr(a, "range") & RU(" \u043F\u043E \u0441\u0442\u043E\u043B\u0431\u0446\u0443 ") & _
                                DictStr(a, "key") & " " & DictStr(a, "order")
        Case "add_sheet":   d = RU("\u0434\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u043B\u0438\u0441\u0442 '") & DictStr(a, "name") & "'"
        Case "chart":       d = RU("\u0434\u0438\u0430\u0433\u0440\u0430\u043C\u043C\u0430 (") & DictStr(a, "type") & RU(") \u043F\u043E \u0434\u0430\u043D\u043D\u044B\u043C ") & DictStr(a, "source")
        Case Else:          d = act
    End Select

    If detailed Then
        Select Case act
            Case "set_formula", "format", "sort"
                n = CellCountOf(DictStr(a, "range"), False)
            Case "clear"
                n = CellCountOf(DictStr(a, "range"), True)
        End Select
        If n > MAX_UNDO_CELLS Then
            d = d & vbLf & RU("      [!] \u0431\u043E\u043B\u044C\u0448\u0435 ") & MAX_UNDO_CELLS & RU(" \u044F\u0447\u0435\u0435\u043A - \u043D\u0435 \u0431\u0443\u0434\u0435\u0442 \u043F\u0440\u0438\u043C\u0435\u043D\u0435\u043D\u043E")
        ElseIf act = "format" And n > MAX_UNDO_FORMAT_CELLS Then
            d = d & vbLf & RU("      [!] \u0431\u043E\u043B\u044C\u0448\u0435 ") & MAX_UNDO_FORMAT_CELLS & _
                RU(" \u044F\u0447\u0435\u0435\u043A - \u043E\u0444\u043E\u0440\u043C\u043B\u0435\u043D\u0438\u0435 \u043D\u0435\u043B\u044C\u0437\u044F \u0431\u0443\u0434\u0435\u0442 \u043E\u0442\u043C\u0435\u043D\u0438\u0442\u044C")
        End If
    End If
    DescribeAction = d
End Function

' Number of cells in an address; usedOnly: only the part inside UsedRange. -1 if the address is bad.
Private Function CellCountOf(ByVal addr As String, ByVal usedOnly As Boolean) As Double
    Dim rng As Range
    CellCountOf = -1
    On Error GoTo Done
    Set rng = ResolveRange(addr)
    If usedOnly Then Set rng = Intersect(rng, rng.Worksheet.UsedRange)
    If rng Is Nothing Then CellCountOf = 0 Else CellCountOf = rng.Cells.CountLarge
Done:
End Function

' The address set_values really writes to: the declared range shrinks/grows to the array
Private Function ValuesTarget(ByVal a As Object) As String
    Dim rng As Range, grid As Variant
    ValuesTarget = DictStr(a, "range")
    On Error GoTo Done
    grid = ValuesToGrid(DictItem(a, "values"))
    Set rng = ResolveRange(ValuesTarget).Cells(1, 1).Resize(UBound(grid, 1), UBound(grid, 2))
    ValuesTarget = "'" & rng.Worksheet.Name & "'!" & rng.Address(False, False)
Done:
End Function

' What set_values will write: first values and every formula (up to 5), so the user sees it before confirming
Private Function ValuesPreview(ByVal v As Variant) As String
    Dim g As Variant, r As Long, c As Long, s As String, vals As String, fmls As String
    Dim nVals As Long, nFmls As Long, risky As Boolean
    On Error GoTo Bad
    g = ValuesToGrid(v)
    For r = 1 To UBound(g, 1)
        For c = 1 To UBound(g, 2)
            s = CellText(g(r, c))
            If Left$(s, 1) = "=" Then
                nFmls = nFmls + 1
                If nFmls <= 5 Then fmls = fmls & vbLf & "      " & Left$(s, 80)
                If IsRiskyFormula(s) Then risky = True
            ElseIf Len(s) > 0 Then
                nVals = nVals + 1
                If nVals <= 5 Then
                    If nVals > 1 Then vals = vals & "; "
                    vals = vals & Chr$(34) & Left$(s, 30) & Chr$(34)
                End If
            End If
        Next c
    Next r

    s = " (" & UBound(g, 1) & " x " & UBound(g, 2) & ")"
    If nVals > 0 Then
        s = s & ": " & vals
        If nVals > 5 Then s = s & RU(" ... \u0438 \u0435\u0449\u0451 ") & (nVals - 5)
    End If
    If nFmls > 0 Then
        s = s & vbLf & RU("      \u0444\u043E\u0440\u043C\u0443\u043B\u044B (") & nFmls & "):" & fmls
        If nFmls > 5 Then s = s & vbLf & "      ..."
    End If
    If risky Then s = s & vbLf & RiskyNote()
    ValuesPreview = s
    Exit Function
Bad:
    ValuesPreview = RU(" (\u043D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u043F\u0440\u043E\u0447\u0438\u0442\u0430\u0442\u044C \u0437\u043D\u0430\u0447\u0435\u043D\u0438\u044F)")
End Function

' Functions that send data outside the workbook or run external code
Private Function IsRiskyFormula(ByVal f As String) As Boolean
    Dim w As Variant
    f = UCase$(f)
    For Each w In Array("WEBSERVICE(", "FILTERXML(", "HYPERLINK(", "RTD(", "CALL(", "REGISTER", "HTTP", "\\")
        If InStr(f, w) > 0 Then
            IsRiskyFormula = True
            Exit Function
        End If
    Next w
End Function

Private Function RiskyNote() As String
    RiskyNote = RU("      [!] \u0444\u043E\u0440\u043C\u0443\u043B\u0430 \u043E\u0431\u0440\u0430\u0449\u0430\u0435\u0442\u0441\u044F \u043A \u0432\u043D\u0435\u0448\u043D\u0438\u043C \u0430\u0434\u0440\u0435\u0441\u0430\u043C \u0438\u043B\u0438 \u0444\u0430\u0439\u043B\u0430\u043C - \u043F\u0440\u043E\u0432\u0435\u0440\u044C\u0442\u0435, \u043E\u0442\u043A\u0443\u0434\u0430 \u044D\u0442\u0438 \u0434\u0430\u043D\u043D\u044B\u0435")
End Function


' ================================= UNDO ======================================

Public Function UndoLast() As String
    Dim batch As Collection, i As Long, u As Object
    If mUndo Is Nothing Then
        UndoLast = RU("\u041E\u0442\u043C\u0435\u043D\u044F\u0442\u044C \u043D\u0435\u0447\u0435\u0433\u043E.")
        Exit Function
    End If
    If mUndo.Count = 0 Then
        UndoLast = RU("\u041E\u0442\u043C\u0435\u043D\u044F\u0442\u044C \u043D\u0435\u0447\u0435\u0433\u043E.")
        Exit Function
    End If

    Set batch = mUndo(mUndo.Count)
    mUndo.Remove mUndo.Count
    Application.ScreenUpdating = False
    On Error Resume Next
    For i = batch.Count To 1 Step -1
        Set u = batch(i)
        Select Case u.Item("type")
            Case "cells"
                RestoreCells u
            Case "sheet"
                Application.DisplayAlerts = False
                u.Item("obj").Delete
                Application.DisplayAlerts = True
            Case "chart"
                u.Item("obj").Delete
        End Select
    Next i
    On Error GoTo 0
    Application.ScreenUpdating = True
    UndoLast = RU("\u041E\u0442\u043C\u0435\u043D\u0435\u043D\u043E \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u0439: ") & batch.Count & RU(". \u041C\u043E\u0436\u043D\u043E \u043E\u0442\u043C\u0435\u043D\u0438\u0442\u044C \u0435\u0449\u0451 \u0448\u0430\u0433\u043E\u0432: ") & mUndo.Count & "."
End Function

Private Sub PushUndo(ByVal batch As Collection)
    If batch.Count = 0 Then Exit Sub
    If mUndo Is Nothing Then Set mUndo = New Collection
    mUndo.Add batch
    If mUndo.Count > 20 Then mUndo.Remove 1
End Sub

Private Sub SnapshotCells(ByVal rng As Range, ByVal batch As Collection, ByVal withFormats As Boolean)
    Dim u As Object, n As Long, i As Long, c As Range, w() As Variant
    Dim nf() As Variant, bo() As Variant, it() As Variant, fc() As Variant
    Dim ci() As Variant, co() As Variant, ha() As Variant, wr() As Variant

    If rng.Cells.CountLarge > MAX_UNDO_CELLS Then Exit Sub
    Set u = CreateObject("Scripting.Dictionary")
    u.Item("type") = "cells"
    Set u.Item("range") = rng
    u.Item("formulas") = rng.Formula

    If withFormats And rng.Cells.CountLarge <= MAX_UNDO_FORMAT_CELLS Then
        n = rng.Cells.Count
        ReDim nf(1 To n): ReDim bo(1 To n): ReDim it(1 To n): ReDim fc(1 To n)
        ReDim ci(1 To n): ReDim co(1 To n): ReDim ha(1 To n): ReDim wr(1 To n)
        For Each c In rng.Cells
            i = i + 1
            nf(i) = c.NumberFormat: bo(i) = c.Font.Bold: it(i) = c.Font.Italic: fc(i) = c.Font.Color
            ci(i) = c.Interior.ColorIndex: co(i) = c.Interior.Color
            ha(i) = c.HorizontalAlignment: wr(i) = c.WrapText
        Next c
        ReDim w(1 To rng.Columns.Count)
        For i = 1 To rng.Columns.Count
            w(i) = rng.Columns(i).ColumnWidth
        Next i
        u.Item("fmt") = True
        u.Item("nf") = nf: u.Item("bo") = bo: u.Item("it") = it: u.Item("fc") = fc
        u.Item("ci") = ci: u.Item("co") = co: u.Item("ha") = ha: u.Item("wr") = wr
        u.Item("w") = w
    End If
    batch.Add u
End Sub

Private Sub RestoreCells(ByVal u As Object)
    Dim rng As Range, c As Range, i As Long
    Dim nf As Variant, bo As Variant, it As Variant, fc As Variant
    Dim ci As Variant, co As Variant, ha As Variant, wr As Variant, w As Variant

    Set rng = u.Item("range")
    rng.Formula = u.Item("formulas")
    If Not u.Exists("fmt") Then Exit Sub

    nf = u.Item("nf"): bo = u.Item("bo"): it = u.Item("it"): fc = u.Item("fc")
    ci = u.Item("ci"): co = u.Item("co"): ha = u.Item("ha"): wr = u.Item("wr"): w = u.Item("w")
    For Each c In rng.Cells
        i = i + 1
        c.NumberFormat = nf(i)
        If Not IsNull(bo(i)) Then c.Font.Bold = bo(i)
        If Not IsNull(it(i)) Then c.Font.Italic = it(i)
        If Not IsNull(fc(i)) Then c.Font.Color = fc(i)
        If ci(i) = xlColorIndexNone Then c.Interior.ColorIndex = xlNone Else c.Interior.Color = co(i)
        c.HorizontalAlignment = ha(i)
        If Not IsNull(wr(i)) Then c.WrapText = wr(i)
    Next c
    For i = 1 To rng.Columns.Count
        rng.Columns(i).ColumnWidth = w(i)
    Next i
    ' note: borders are not restored by undo
End Sub


' =============================== CONTEXT =====================================

' Range -> text: first line = column letters, each row starts with its row number.
' withFormulas: cells with formulas are shown as  value {=formula}
Public Function BuildContext(ByVal rng As Range, ByVal withFormulas As Boolean, _
                             ByVal withHeader As Boolean, Optional ByVal maxRows As Long = 0) As String
    On Error GoTo Fail
    Dim area As Range
    Set area = Intersect(rng, rng.Worksheet.UsedRange)
    If area Is Nothing Then Exit Function
    Set area = area.Areas(1)

    Dim vals As Variant, fmls As Variant, nR As Long, nC As Long
    nR = area.Rows.Count
    nC = area.Columns.Count
    If nR = 1 And nC = 1 Then
        ReDim vals(1 To 1, 1 To 1)
        vals(1, 1) = area.Value
        If withFormulas Then
            ReDim fmls(1 To 1, 1 To 1)
            fmls(1, 1) = area.Formula
        End If
    Else
        vals = area.Value
        If withFormulas Then fmls = area.Formula
    End If

    Dim buf As String, used As Long, r As Long, c As Long, f As String, cellStr As String
    If withHeader Then
        SbAdd buf, used, "Sheet: " & area.Worksheet.Name & " | Range: " & area.Address(False, False) & vbLf
    End If

    SbAdd buf, used, "Row"
    For c = 1 To nC
        SbAdd buf, used, vbTab & ColLetter(area.Column + c - 1)
    Next c
    SbAdd buf, used, vbLf

    For r = 1 To nR
        SbAdd buf, used, CStr(area.Row + r - 1)
        For c = 1 To nC
            cellStr = CellText(vals(r, c))
            If withFormulas Then
                f = CStr(fmls(r, c))
                If Left$(f, 1) = "=" Then cellStr = cellStr & " {" & f & "}"
            End If
            SbAdd buf, used, vbTab & cellStr
        Next c
        SbAdd buf, used, vbLf

        If maxRows > 0 And r >= maxRows And r < nR Then
            SbAdd buf, used, "...[" & (nR - r) & " more rows, last row is " & (area.Row + nR - 1) & "]" & vbLf
            Exit For
        End If
        If used > MAX_CONTEXT_CHARS Then
            SbAdd buf, used, "...[truncated after row " & (area.Row + r - 1) & "]" & vbLf
            Exit For
        End If
    Next r

    BuildContext = Left$(buf, used)
    Exit Function
Fail:
    BuildContext = ""
End Function


' ============================== INTERNALS ====================================

Private Function SystemPrompt(ByVal agentMode As Boolean) As String
    Dim s As String
    s = "You are an AI assistant working inside Microsoft Excel. "
    s = s & "Answer in " & Setting("Language") & " unless the user writes in another language. "
    s = s & "Be concise and practical. "
    s = s & "Excel data is given as tab-separated text: the line starting with 'Row' lists column letters, "
    s = s & "every next line starts with the Excel row number. A cell with a formula is shown as value {=formula}. "
    s = s & "Refer to cells by address (e.g. B17). Formulas you suggest must work in Excel. "
    s = s & "If the user asks for a table, output it as a markdown table."
    If Len(Setting("Instructions")) > 0 Then s = s & vbLf & "User instructions: " & Setting("Instructions")

    If agentMode Then
        s = s & vbLf & vbLf & "You CAN change the workbook yourself. When (and only when) the user asks to change, create, "
        s = s & "fill, clean, format, sort or chart something, do it in THIS answer: explain in 1-3 sentences what "
        s = s & "you are doing, then add ONE fenced block that starts with ```actions and contains a JSON array of actions:" & vbLf
        s = s & "{""action"":""set_values"",""range"":""Sheet1!A1"",""values"":[[""Name"",""Total""],[""A"",10]]}"
        s = s & "  - writes a 2D array starting at the top-left cell; a string starting with = is a formula" & vbLf
        s = s & "{""action"":""set_formula"",""range"":""Sheet1!C2:C100"",""formula"":""=A2*B2""}"
        s = s & "  - one formula for the whole range, relative references adjust like fill-down" & vbLf
        s = s & "{""action"":""format"",""range"":""Sheet1!A1:D1"",""bold"":true,""italic"":false,"
        s = s & """fill"":""#FFF2CC"",""font_color"":""#000000"",""number_format"":""#,##0.00"","
        s = s & """align"":""center"",""wrap"":true,""borders"":true,""autofit"":true}"
        s = s & "  - any subset of these properties; fill ""none"" removes the fill" & vbLf
        s = s & "{""action"":""clear"",""range"":""Sheet1!E2:E50""}" & vbLf
        s = s & "{""action"":""add_sheet"",""name"":""Summary""}" & vbLf
        s = s & "{""action"":""sort"",""range"":""Sheet1!A1:F200"",""key"":""C"",""order"":""desc"",""header"":true}" & vbLf
        s = s & "{""action"":""chart"",""source"":""Sheet1!A1:B13"",""type"":""column"",""title"":""Sales"","
        s = s & """target"":""Sheet1!H2""}  - types: column, bar, line, pie, scatter, area" & vbLf
        s = s & "set_values must cover EVERY row you were asked about: the number of value rows must equal "
        s = s & "the number of rows in the range (E2:E39 = 38 rows = 38 values). Never shorten the list, never "
        s = s & "write '...' or 'and so on'. If the answer would be very long, split it into several set_values "
        s = s & "actions in the same block (E2:E20, E21:E39) so that together they cover the whole range. "
        s = s & "If the data you were given is cut off (a '...[N more rows]' marker), say so and ask for the rest "
        s = s & "instead of guessing. " & vbLf
        s = s & "Rules: always include the sheet name in every range (quote names with spaces: 'My sheet'!A1); "
        s = s & "formulas use English function names and commas; never change cells the user did not ask about. "
        If Setting("ValuesNotFormulas") = "1" Then
            s = s & "Write plain values with set_values. Do NOT use set_formula and do not put a formula string "
            s = s & "into set_values unless the user explicitly asked for a formula. "
        Else
            s = s & "Use set_formula for things that must recalculate: sums, totals, percentages, dates, "
            s = s & "anything derived from numbers the user may still edit. "
            s = s & "But when the user asks you to FILL a column by matching, classifying or looking up rows "
            s = s & "against a reference sheet, write the resulting values with set_values: that is a one-off "
            s = s & "mapping, and a lookup formula would break as soon as the reference is re-sorted. "
            s = s & "Avoid XLOOKUP, LET, LAMBDA and other functions missing from Excel 2016/2019 unless the user "
            s = s & "says their Excel supports them; VLOOKUP or INDEX/MATCH work everywhere. "
        End If
        s = s & "The actions block is applied to the workbook IMMEDIATELY, without any confirmation: never ask the user "
        s = s & "to confirm and never say that changes are waiting for confirmation. Without the block nothing changes. "
        s = s & "If the target cells are unclear, ask a question instead of adding the block."
    Else
        s = s & vbLf & vbLf & "You CANNOT change the workbook in this mode. Never claim that you wrote, changed "
        s = s & "or checked cells. If the user asks to change something, say that the '" & RU("\u0420\u0430\u0437\u0440\u0435\u0448\u0438\u0442\u044C \u043F\u0440\u0430\u0432\u043A\u0438") & "' checkbox "
        s = s & "must be enabled, or give the exact cell address and value so the user can enter it."
    End If
    SystemPrompt = s
End Function

Public Function RequestLLM(ByVal messagesJson As String, ByVal model As String, _
                           ByVal maxTokens As Long, Optional ByVal temperature As Double = -1) As String
    On Error GoTo Fail
    LastUsage = 0
    LastTruncated = False
    LastPromptTokens = 0
    LastCompletionTokens = 0
    LastReasoningTokens = 0

    Dim apiKey As String
    apiKey = Trim$(Setting("ApiKey"))
    If Len(apiKey) = 0 Then
        RequestLLM = RU("#ERROR: \u043D\u0435 \u0437\u0430\u0434\u0430\u043D API-\u043A\u043B\u044E\u0447. \u0412\u043F\u0438\u0448\u0438\u0442\u0435 \u0435\u0433\u043E \u0432 \u043E\u043A\u043D\u0435 \u0447\u0430\u0442\u0430 -> \u041D\u0430\u0441\u0442\u0440\u043E\u0439\u043A\u0438 -> API-\u043A\u043B\u044E\u0447.")
        Exit Function
    End If
    If Len(Trim$(model)) = 0 Then model = Setting("Model")
    If maxTokens <= 0 Then maxTokens = 1000
    LastMaxTokens = maxTokens

    Dim body As String
    body = "{""model"":""" & JsonEscape(Trim$(model)) & """,""max_tokens"":" & maxTokens
    If temperature >= 0 Then body = body & ",""temperature"":" & Replace(CStr(temperature), ",", ".")
    body = body & ",""messages"":" & messagesJson & "}"

    ' ServerXMLHTTP uses WinHTTP proxy settings. If requests fail behind a corporate
    ' proxy, replace with "MSXML2.XMLHTTP.6.0" and delete the setTimeouts line.
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 10000, 10000, TIMEOUT_MS, TIMEOUT_MS
    http.Open "POST", Setting("URL"), False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send body

    Dim resp As String
    resp = http.responseText

    If http.Status < 200 Or http.Status >= 300 Then
        Dim errMsg As String
        errMsg = ExtractJsonString(resp, """message""", 1)
        If Len(errMsg) = 0 Then errMsg = Left$(resp, 300)
        RequestLLM = "#ERROR " & http.Status & ": " & Left$(errMsg, 300)
        Exit Function
    End If

    LastUsage = JsonNumber(resp, """total_tokens""")
    LastPromptTokens = JsonNumber(resp, """prompt_tokens""")
    LastCompletionTokens = JsonNumber(resp, """completion_tokens""")
    LastReasoningTokens = JsonNumber(resp, """reasoning_tokens""")

    Dim pos As Long, answer As String, fin As String
    pos = InStr(1, resp, """choices""")
    If pos = 0 Then pos = 1
    fin = LCase$(ExtractJsonString(resp, """finish_reason""", pos))
    If fin = "length" Or fin = "max_tokens" Then LastTruncated = True
    answer = ExtractJsonString(resp, """content""", pos)
    If Len(answer) = 0 Then
        RequestLLM = RU("#ERROR: \u043C\u043E\u0434\u0435\u043B\u044C \u0432\u0435\u0440\u043D\u0443\u043B\u0430 \u043F\u0443\u0441\u0442\u043E\u0439 \u043E\u0442\u0432\u0435\u0442")
    Else
        answer = Replace(answer, vbCrLf, vbLf)
        answer = Replace(answer, vbCr, vbLf)
        RequestLLM = Left$(answer, MAX_ANSWER_CHARS)
    End If
    Exit Function

Fail:
    RequestLLM = "#ERROR: " & Err.Description
End Function

' Decodes \uXXXX escapes: keeps the source ASCII-only (see the header)
Public Function RU(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW$(CLng("&H" & Mid$(s, p + 2, 4))) & Mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    RU = s
End Function

Private Function Msg(ByVal role As String, ByVal content As String) As String
    Msg = "{""role"":""" & role & """,""content"":""" & JsonEscape(content) & """}"
End Function

' Fresh Excel data is attached to every message, so the copies stored in earlier turns
' are dead weight: they multiply the prompt until the model has no room left to answer.
' Only the newest message keeps its data block.
Private Sub DropOldContexts()
    Dim i As Long, p As Long
    For i = 1 To mCount
        If mRoles(i) = "user" Then
            p = InStr(mTexts(i), vbLf & vbLf & "Excel data:" & vbLf)
            If p > 0 Then mTexts(i) = Left$(mTexts(i), p - 1) & vbLf & "[Excel data omitted, see the latest message]"
        End If
    Next i
End Sub

Private Sub AddHistory(ByVal role As String, ByVal text As String)
    Dim i As Long
    mCount = mCount + 1
    ReDim Preserve mRoles(1 To mCount)
    ReDim Preserve mTexts(1 To mCount)
    mRoles(mCount) = role
    mTexts(mCount) = text
    Do While mCount > 2 And (mCount > MAX_HISTORY_MESSAGES Or HistoryChars() > 3 * MAX_CONTEXT_CHARS)
        For i = 3 To mCount
            mRoles(i - 2) = mRoles(i)
            mTexts(i - 2) = mTexts(i)
        Next i
        mCount = mCount - 2
        ReDim Preserve mRoles(1 To mCount)
        ReDim Preserve mTexts(1 To mCount)
    Loop
End Sub

Private Function HistoryChars() As Long
    Dim i As Long
    For i = 1 To mCount
        HistoryChars = HistoryChars + Len(mTexts(i))
    Next i
End Function

Private Function ResolveRange(ByVal addr As String) As Range
    Dim p As Long, sh As String
    addr = Trim$(addr)
    If Len(addr) = 0 Then Err.Raise vbObjectError + 4, , RU("\u043D\u0435 \u0443\u043A\u0430\u0437\u0430\u043D \u0434\u0438\u0430\u043F\u0430\u0437\u043E\u043D")
    p = InStrRev(addr, "!")
    If p > 0 Then
        sh = Left$(addr, p - 1)
        If Left$(sh, 1) = "'" And Right$(sh, 1) = "'" Then sh = Mid$(sh, 2, Len(sh) - 2)
        sh = Replace(sh, "''", "'")
        Set ResolveRange = ActiveWorkbook.Worksheets(sh).Range(Mid$(addr, p + 1))
    Else
        Set ResolveRange = ActiveSheet.Range(addr)
    End If
End Function

Private Sub SetFormulaSafe(ByVal rng As Range, ByVal value As Variant)
    On Error Resume Next
    CallByName rng, "Formula2", VbLet, value      ' Excel 365/2021 (dynamic arrays)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        rng.Formula = value                        ' older Excel
    End If
End Sub

Private Function ValuesToGrid(ByVal v As Variant) As Variant
    Dim g() As Variant, r As Long, c As Long, maxC As Long, row As Variant, item As Variant
    If IsObject(v) Then
        If TypeName(v) <> "Collection" Then Err.Raise vbObjectError + 2, , RU("values \u0434\u043E\u043B\u0436\u0435\u043D \u0431\u044B\u0442\u044C \u043C\u0430\u0441\u0441\u0438\u0432\u043E\u043C")
        If v.Count = 0 Then Err.Raise vbObjectError + 2, , RU("\u043F\u0443\u0441\u0442\u043E\u0439 \u0441\u043F\u0438\u0441\u043E\u043A values")
        If IsObject(v.Item(1)) Then
            For Each row In v
                If IsObject(row) Then
                    If row.Count > maxC Then maxC = row.Count
                End If
            Next row
            If maxC = 0 Then maxC = 1
            ReDim g(1 To v.Count, 1 To maxC)
            For Each row In v
                r = r + 1
                If IsObject(row) Then
                    c = 0
                    For Each item In row
                        c = c + 1
                        g(r, c) = JsonToCell(item)
                    Next item
                Else
                    g(r, 1) = JsonToCell(row)
                End If
            Next row
        Else
            ReDim g(1 To 1, 1 To v.Count)
            For Each item In v
                c = c + 1
                g(1, c) = JsonToCell(item)
            Next item
        End If
    Else
        ReDim g(1 To 1, 1 To 1)
        g(1, 1) = JsonToCell(v)
    End If
    ValuesToGrid = g
End Function

Private Function JsonToCell(ByVal x As Variant) As Variant
    If IsObject(x) Then
        JsonToCell = ""
    ElseIf IsNull(x) Then
        JsonToCell = Empty
    Else
        JsonToCell = x
    End If
End Function

Private Function DictStr(ByVal d As Object, ByVal key As String) As String
    If Not d.Exists(key) Then Exit Function
    If IsObject(d.Item(key)) Then Exit Function
    If IsNull(d.Item(key)) Then Exit Function
    DictStr = CStr(d.Item(key))
End Function

Private Function DictItem(ByVal d As Object, ByVal key As String) As Variant
    If Not d.Exists(key) Then Exit Function
    If IsObject(d.Item(key)) Then
        Set DictItem = d.Item(key)
    Else
        DictItem = d.Item(key)
    End If
End Function

Private Function HexColor(ByVal s As String) As Long
    s = Replace(Trim$(s), "#", "")
    If Len(s) <> 6 Then Err.Raise vbObjectError + 5, , RU("\u043D\u0435\u0432\u0435\u0440\u043D\u044B\u0439 \u0446\u0432\u0435\u0442 ") & s
    HexColor = RGB(CLng("&H" & Mid$(s, 1, 2)), CLng("&H" & Mid$(s, 3, 2)), CLng("&H" & Mid$(s, 5, 2)))
End Function

Private Function ChartTypeFrom(ByVal t As String) As Long
    Select Case LCase$(t)
        Case "bar":     ChartTypeFrom = xlBarClustered
        Case "line":    ChartTypeFrom = xlLine
        Case "pie":     ChartTypeFrom = xlPie
        Case "scatter": ChartTypeFrom = xlXYScatter
        Case "area":    ChartTypeFrom = xlArea
        Case Else:      ChartTypeFrom = xlColumnClustered
    End Select
End Function

' Markdown table or tab-separated text -> 2D array; returns Empty if text is not a table.
' numbersAsNumbers: "12.5" -> 12.5 (for worksheet functions)
Private Function TextToGrid(ByVal text As String, ByVal numbersAsNumbers As Boolean) As Variant
    Dim lines() As String, i As Long, ln As String, isMd As Boolean
    Dim rowsCol As New Collection, parts As Variant, maxC As Long, r As Long, c As Long
    Dim arr() As Variant, s As String

    text = Replace(text, vbCr, "")
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 1) = "|" Then
            isMd = True
            If Not IsMdSeparator(ln) Then
                ln = Mid$(ln, 2)
                If Right$(ln, 1) = "|" Then ln = Left$(ln, Len(ln) - 1)
                rowsCol.Add Split(Replace(ln, "**", ""), "|")
            End If
        End If
    Next i
    If Not isMd And InStr(text, vbTab) > 0 Then
        For i = 0 To UBound(lines)
            If Len(Trim$(lines(i))) > 0 Then rowsCol.Add Split(lines(i), vbTab)
        Next i
    End If
    If rowsCol.Count = 0 Then Exit Function

    For Each parts In rowsCol
        If UBound(parts) + 1 > maxC Then maxC = UBound(parts) + 1
    Next parts
    If maxC = 0 Then Exit Function

    ReDim arr(1 To rowsCol.Count, 1 To maxC)
    For r = 1 To rowsCol.Count
        For c = 1 To maxC
            arr(r, c) = ""
        Next c
    Next r
    r = 0
    For Each parts In rowsCol
        r = r + 1
        For c = 0 To UBound(parts)
            s = Trim$(parts(c))
            If numbersAsNumbers And IsPlainNumber(s) Then
                arr(r, c + 1) = Val(s)
            ElseIf numbersAsNumbers Then
                arr(r, c + 1) = s
            Else
                arr(r, c + 1) = SafeCellValue(s)
            End If
        Next c
    Next parts
    TextToGrid = arr
End Function

Private Function IsPlainNumber(ByVal s As String) As Boolean
    Dim i As Long, ch As String, dots As Long, digits As Long
    If Len(s) = 0 Or Len(s) > 20 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "." Then
            dots = dots + 1
        ElseIf ch = "-" And i = 1 Then
            ' leading minus is fine
        Else
            Exit Function
        End If
    Next i
    IsPlainNumber = (digits > 0 And dots <= 1)
End Function

Private Function ArrayToText(ByVal v As Variant) As String
    Dim r As Long, c As Long, buf As String, used As Long
    If Not IsArray(v) Then
        ArrayToText = CellText(v)
        Exit Function
    End If
    If ArrayDims(v) = 1 Then
        For c = LBound(v) To UBound(v)
            If c > LBound(v) Then SbAdd buf, used, vbTab
            SbAdd buf, used, CellText(v(c))
        Next c
    Else
        For r = LBound(v, 1) To UBound(v, 1)
            For c = LBound(v, 2) To UBound(v, 2)
                If c > LBound(v, 2) Then SbAdd buf, used, vbTab
                SbAdd buf, used, CellText(v(r, c))
            Next c
            SbAdd buf, used, vbLf
        Next r
    End If
    ArrayToText = Left$(buf, used)
End Function

Private Function CellText(ByVal x As Variant) As String
    If IsError(x) Then
        CellText = ErrText(x)
    ElseIf IsEmpty(x) Or IsNull(x) Then
        CellText = ""
    Else
        CellText = Replace(Replace(Replace(CStr(x), vbTab, " "), vbCr, ""), vbLf, " / ")
    End If
End Function

Private Function ErrText(ByVal x As Variant) As String
    Select Case x
        Case CVErr(xlErrDiv0):  ErrText = "#DIV/0!"
        Case CVErr(xlErrNA):    ErrText = "#N/A"
        Case CVErr(xlErrName):  ErrText = "#NAME?"
        Case CVErr(xlErrNull):  ErrText = "#NULL!"
        Case CVErr(xlErrNum):   ErrText = "#NUM!"
        Case CVErr(xlErrRef):   ErrText = "#REF!"
        Case CVErr(xlErrValue): ErrText = "#VALUE!"
        Case Else:              ErrText = "#ERROR"
    End Select
End Function

Private Function SafeCellValue(ByVal s As String) As String
    Dim first As String
    If Len(s) > 0 Then
        first = Left$(s, 1)
        If first = "=" Or ((first = "+" Or first = "-" Or first = "@") And Not IsNumeric(s)) Then s = "'" & s
    End If
    SafeCellValue = s
End Function

Private Function ConfirmOverwrite(ByVal dest As Range) As Boolean
    If Application.WorksheetFunction.CountA(dest) = 0 Then
        ConfirmOverwrite = True
    Else
        ConfirmOverwrite = (MsgBox(RU("\u042F\u0447\u0435\u0439\u043A\u0438 ") & dest.Address(False, False) & RU(" \u043D\u0435 \u043F\u0443\u0441\u0442\u044B\u0435. \u041F\u0435\u0440\u0435\u0437\u0430\u043F\u0438\u0441\u0430\u0442\u044C?"), _
                                   vbYesNo + vbExclamation, RU("\u0418\u0418")) = vbYes)
    End If
End Function

Private Function IsMdSeparator(ByVal ln As String) As Boolean
    Dim t As String
    t = Replace(Replace(Replace(Replace(ln, "|", ""), "-", ""), ":", ""), " ", "")
    IsMdSeparator = (Len(t) = 0 And InStr(ln, "-") > 0)
End Function

Private Function ColLetter(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(65 + (n - 1) Mod 26) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function ArrayDims(ByVal v As Variant) As Long
    Dim d As Long, tmp As Long
    On Error GoTo Done
    Do
        d = d + 1
        tmp = UBound(v, d)
    Loop
Done:
    ArrayDims = d - 1
End Function

Private Sub SbAdd(ByRef buf As String, ByRef used As Long, ByVal piece As String)
    Dim n As Long
    n = Len(piece)
    If n = 0 Then Exit Sub
    If used + n > Len(buf) Then buf = buf & Space$(IIf(Len(buf) > n, Len(buf), n) + 4096)
    Mid$(buf, used + 1, n) = piece
    used = used + n
End Sub


' ================================ JSON =======================================

Private Function JsonEscape(ByVal s As String) As String
    Dim i As Long, code As Long, ch As String, buf As String, used As Long
    buf = Space$(Len(s) + 16)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch) And &HFFFF&
        Select Case code
            Case 34:        SbAdd buf, used, "\"""
            Case 92:        SbAdd buf, used, "\\"
            Case 10:        SbAdd buf, used, "\n"
            Case 13:        SbAdd buf, used, "\r"
            Case 9:         SbAdd buf, used, "\t"
            Case 32 To 126: SbAdd buf, used, ch
            Case Else:      SbAdd buf, used, "\u" & Right$("000" & Hex$(code), 4)
        End Select
    Next i
    JsonEscape = Left$(buf, used)
End Function

Private Function JsonNumber(ByVal json As String, ByVal key As String) As Long
    Dim p As Long
    p = InStr(1, json, key)
    If p = 0 Then Exit Function
    p = InStr(p + Len(key), json, ":")
    If p = 0 Then Exit Function
    JsonNumber = CLng(Val(Mid$(json, p + 1, 20)))
End Function

Private Function ExtractJsonString(ByVal json As String, ByVal key As String, _
                                   ByVal startPos As Long) As String
    Dim p As Long
    p = InStr(startPos, json, key)
    Do While p > 0
        p = SkipSpaces(json, p + Len(key))
        If Mid$(json, p, 1) = ":" Then
            p = SkipSpaces(json, p + 1)
            If Mid$(json, p, 1) = """" Then
                ExtractJsonString = ReadJsonStringAt(json, p)
                Exit Function
            End If
        End If
        p = InStr(p, json, key)
    Loop
End Function

Private Function SkipSpaces(ByVal s As String, ByVal p As Long) As Long
    Do While p <= Len(s)
        Select Case Mid$(s, p, 1)
            Case " ", vbTab, vbLf, vbCr: p = p + 1
            Case Else: Exit Do
        End Select
    Loop
    SkipSpaces = p
End Function

' p points to the opening quote; on return p points after the closing quote
Private Function ReadJsonStringAt(ByVal s As String, ByRef p As Long) As String
    Dim ch As String, buf As String, used As Long
    If Mid$(s, p, 1) <> """" Then Err.Raise vbObjectError + 10, , "JSON: string expected"
    p = p + 1
    Do While p <= Len(s)
        ch = Mid$(s, p, 1)
        If ch = """" Then Exit Do
        If ch = "\" Then
            p = p + 1
            ch = Mid$(s, p, 1)
            Select Case ch
                Case "n": SbAdd buf, used, vbLf
                Case "r": SbAdd buf, used, vbCr
                Case "t": SbAdd buf, used, vbTab
                Case "b": SbAdd buf, used, Chr$(8)
                Case "f": SbAdd buf, used, Chr$(12)
                Case "u"
                    SbAdd buf, used, ChrW$(CLng(Val("&H" & Mid$(s, p + 1, 4) & "&")))
                    p = p + 4
                Case Else: SbAdd buf, used, ch
            End Select
        Else
            SbAdd buf, used, ch
        End If
        p = p + 1
    Loop
    p = p + 1
    ReadJsonStringAt = Left$(buf, used)
End Function

' Mini JSON parser: objects -> Scripting.Dictionary, arrays -> Collection
Private Sub JsonValue(ByVal s As String, ByRef p As Long, ByRef result As Variant)
    p = SkipSpaces(s, p)
    If p > Len(s) Then Err.Raise vbObjectError + 11, , "JSON: unexpected end"
    Select Case Mid$(s, p, 1)
        Case "{": Set result = JsonObject(s, p)
        Case "[": Set result = JsonArray(s, p)
        Case """": result = ReadJsonStringAt(s, p)
        Case "t": result = True: p = p + 4
        Case "f": result = False: p = p + 5
        Case "n": result = Null: p = p + 4
        Case Else: result = JsonNumberAt(s, p)
    End Select
End Sub

Private Function JsonObject(ByVal s As String, ByRef p As Long) As Object
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    p = p + 1
    Do
        p = SkipSpaces(s, p)
        If p > Len(s) Then Exit Do
        Select Case Mid$(s, p, 1)
            Case "}"
                p = p + 1
                Exit Do
            Case ","
                p = p + 1
            Case Else
                k = ReadJsonStringAt(s, p)
                p = SkipSpaces(s, p)
                If Mid$(s, p, 1) <> ":" Then Err.Raise vbObjectError + 12, , "JSON: ':' expected"
                p = p + 1
                JsonValue s, p, v
                If IsObject(v) Then Set d.Item(k) = v Else d.Item(k) = v
        End Select
    Loop
    Set JsonObject = d
End Function

Private Function JsonArray(ByVal s As String, ByRef p As Long) As Collection
    Dim c As New Collection, v As Variant
    p = p + 1
    Do
        p = SkipSpaces(s, p)
        If p > Len(s) Then Exit Do
        Select Case Mid$(s, p, 1)
            Case "]"
                p = p + 1
                Exit Do
            Case ","
                p = p + 1
            Case Else
                JsonValue s, p, v
                c.Add v
        End Select
    Loop
    Set JsonArray = c
End Function

Private Function JsonNumberAt(ByVal s As String, ByRef p As Long) As Variant
    Dim st As Long
    st = p
    Do While p <= Len(s)
        If InStr("+-.eE0123456789", Mid$(s, p, 1)) = 0 Then Exit Do
        p = p + 1
    Loop
    If p = st Then Err.Raise vbObjectError + 13, , "JSON: unexpected character"
    JsonNumberAt = Val(Mid$(s, st, p - st))
End Function
