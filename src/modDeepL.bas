Attribute VB_Name = "modDeepL"
Option Explicit

'========================
' ユーザー設定
'========================
Private Const DEEPL_AUTH_KEY As String = "ここにDeepLのAPIキー"
Private Const DEEPL_ENDPOINT As String = "https://api-free.deepl.com/v2/translate"
' Pro版ならこっち
' Private Const DEEPL_ENDPOINT As String = "https://api.deepl.com/v2/translate"

' 空欄ならすべてのワークシートが対象。
' 例: "Sheet1" とすると、そのシートだけを対象にします。
Private Const TARGET_SHEET_NAME As String = ""

' 自分のExcelレイアウトに合わせて変更してください。
Private Const TRIGGER_CELL As String = "A2"
Private Const SOURCE_CELL As String = "B2"
Private Const OUTPUT_CELL As String = "C2"

'========================
' 対象シート判定
'========================
Public Function IsDeepLTargetSheet(ByVal Sh As Worksheet) As Boolean
    Dim configuredName As String

    configuredName = Trim$(TARGET_SHEET_NAME)

    If configuredName = "" Then
        IsDeepLTargetSheet = True
    Else
        IsDeepLTargetSheet = (StrComp(Sh.Name, configuredName, vbTextCompare) = 0)
    End If
End Function

'========================
' トリガーセル取得
'========================
Public Function GetDeepLTriggerRange(ByVal Sh As Worksheet) As Range
    On Error GoTo EH

    Set GetDeepLTriggerRange = Sh.Range(TRIGGER_CELL)
    Exit Function

EH:
    Debug.Print "GetDeepLTriggerRange error: " & Err.Number & " - " & Err.Description
    Set GetDeepLTriggerRange = Nothing
End Function

'========================
' 氏名元セル再計算
'========================
Public Sub RecalculateDeepLSource(ByVal Sh As Worksheet)
    On Error GoTo EH

    Sh.Range(SOURCE_CELL).Calculate
    Exit Sub

EH:
    Debug.Print "RecalculateDeepLSource error: " & Err.Number & " - " & Err.Description
End Sub

'========================
' メイン処理
'========================
Public Sub UpdateKana_FromDeepL(ByVal Sh As Worksheet)
    Dim srcName As String
    Dim kana As String

    srcName = Trim$(CStr(Sh.Range(SOURCE_CELL).Value))

    If srcName = "" Then
        Sh.Range(OUTPUT_CELL).Value = ""
        Exit Sub
    End If

    kana = GetKanaFromDeepL(srcName)

    If kana <> "" Then
        Sh.Range(OUTPUT_CELL).Value = kana
    Else
        Sh.Range(OUTPUT_CELL).Value = ""
    End If
End Sub

'========================
' DeepL呼び出し
'========================
Private Function GetKanaFromDeepL(ByVal personName As String) As String
    Dim http As Object
    Dim body As String
    Dim responseText As String
    Dim resultText As String

    On Error GoTo EH

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    body = BuildTranslateRequestBody(personName)

    http.Open "POST", DEEPL_ENDPOINT, False
    http.SetRequestHeader "Authorization", "DeepL-Auth-Key " & DEEPL_AUTH_KEY
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetTimeouts 10000, 10000, 15000, 15000
    http.Send body

    responseText = Utf8BytesToString(http.ResponseBody)

    If http.Status <> 200 Then
        Debug.Print "DeepL HTTP Error: " & http.Status
        Debug.Print responseText
        GetKanaFromDeepL = ""
        Exit Function
    End If

    resultText = ExtractFirstTranslatedText(responseText)
    resultText = NormalizeKanaResult(resultText)

    GetKanaFromDeepL = resultText
    Exit Function

EH:
    Debug.Print "GetKanaFromDeepL error: " & Err.Number & " - " & Err.Description
    GetKanaFromDeepL = ""
End Function

'========================
' JSON作成
'========================
Private Function BuildTranslateRequestBody(ByVal personName As String) As String
    Dim escapedName As String
    Dim escapedContext As String
    Dim instruction1 As String
    Dim instruction2 As String
    Dim instruction3 As String

    escapedName = JsonEscape(personName)

    escapedContext = JsonEscape( _
        "This text is a Japanese person's full name used in HR paperwork. " & _
        "Return the reading in full-width Katakana suitable for furigana entry." _
    )

    instruction1 = JsonEscape("Return the person's name reading in full-width Katakana.")
    instruction2 = JsonEscape("Output only the Katakana result with no explanation.")
    instruction3 = JsonEscape("Do not add spaces unless the original name clearly requires them.")

    BuildTranslateRequestBody = _
        "{" & _
            """text"": [""" & escapedName & """]," & _
            """target_lang"":""JA""," & _
            """context"":""" & escapedContext & """," & _
            """custom_instructions"":[" & _
                """" & instruction1 & """," & _
                """" & instruction2 & """," & _
                """" & instruction3 & """" & _
            "]" & _
        "}"
End Function

'========================
' JSONレスポンスから最初のtextを抜く
'========================
Private Function ExtractFirstTranslatedText(ByVal json As String) As String
    Dim re As Object
    Dim matches As Object
    Dim rawText As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = """text""\s*:\s*""((?:\\""|[^""])*)"""

    If re.Test(json) Then
        Set matches = re.Execute(json)
        rawText = matches(0).SubMatches(0)
        ExtractFirstTranslatedText = JsonUnescape(rawText)
    Else
        ExtractFirstTranslatedText = ""
    End If
End Function

'========================
' UTF-8バイト列 → 文字列
'========================
Private Function Utf8BytesToString(ByVal bytes As Variant) As String
    Dim stm As Object

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "utf-8"
    Utf8BytesToString = stm.ReadText
    stm.Close
    Set stm = Nothing
End Function

'========================
' 出力整形
'========================
Private Function NormalizeKanaResult(ByVal s As String) As String
    Dim t As String

    t = Trim$(s)
    t = Replace(t, ChrW(&H3000), " ")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")

    Do While Len(t) > 0 And (Left$(t, 1) = """" Or Left$(t, 1) = "「" Or Left$(t, 1) = "『")
        t = Mid$(t, 2)
    Loop

    Do While Len(t) > 0 And (Right$(t, 1) = """" Or Right$(t, 1) = "」" Or Right$(t, 1) = "』")
        t = Left$(t, Len(t) - 1)
    Loop

    NormalizeKanaResult = Trim$(t)
End Function

'========================
' JSONエスケープ
'========================
Private Function JsonEscape(ByVal s As String) As String
    Dim t As String

    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, "/", "\/")
    t = Replace(t, vbBack, "\b")
    t = Replace(t, vbFormFeed, "\f")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")

    JsonEscape = t
End Function

'========================
' JSONアンエスケープ
'========================
Private Function JsonUnescape(ByVal s As String) As String
    Dim t As String

    t = s
    t = Replace(t, "\""","""")
    t = Replace(t, "\/", "/")
    t = Replace(t, "\\", "\")
    t = Replace(t, "\b", vbBack)
    t = Replace(t, "\f", vbFormFeed)
    t = Replace(t, "\n", vbLf)
    t = Replace(t, "\r", vbCr)
    t = Replace(t, "\t", vbTab)

    JsonUnescape = t
End Function
