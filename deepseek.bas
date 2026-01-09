Attribute VB_Name = "模块1"
Option Explicit

' ================= 主入口：逐行分析 =================
Sub AnalyzeReviewsWithDeepSeek()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim reviewText As String
    Dim apiResult As String
    Dim parts() As String
    Dim totalRows As Long
    Dim doneCount As Long
    
    ' 按你的表结构：A=内容, B=情感极性, C=情绪类型, D=玩家体验受损类型
    Const colReview As Long = 1  ' A
    Const colPolarity As Long = 2 ' B
    Const colEmotion As Long = 3  ' C
    Const colDamage As Long = 4   ' D
    
    Set ws = ActiveSheet
    
    ' 以“内容”所在列（A 列）为准找最后一行
    lastRow = ws.Cells(ws.Rows.Count, colReview).End(xlUp).Row
    ' 有效数据行数（除去表头）
    totalRows = Application.Max(lastRow - 1, 0)
    doneCount = 0
    
    ' 初始化状态栏
    Application.StatusBar = "准备开始处理，共 " & totalRows & " 行差评数据……"
    
    ' 从第 2 行开始（第 1 行是表头）
    For i = 2 To lastRow
        DoEvents    ' 让 Excel 有机会响应界面和中断（Esc/Ctrl+Break）
        
        reviewText = Trim$(CStr(ws.Cells(i, colReview).Value))
        
        If reviewText <> "" Then
            ' 不想覆盖已有结果就保留这个判断
            If Trim$(CStr(ws.Cells(i, colPolarity).Value)) = "" _
               And Trim$(CStr(ws.Cells(i, colEmotion).Value)) = "" _
               And Trim$(CStr(ws.Cells(i, colDamage).Value)) = "" Then
               
                apiResult = DeepSeekAnalyzeTranslation(reviewText)
                
                If apiResult <> "" Then
                    parts = Split(apiResult, "|")
                    
                    ' 情感极性（B）
                    If UBound(parts) >= 0 Then
                        ws.Cells(i, colPolarity).Value = Trim$(parts(0))
                    End If
                    
                    ' 情绪类型（C）
                    If UBound(parts) >= 1 Then
                        ws.Cells(i, colEmotion).Value = Trim$(parts(1))
                    End If
                    
                    ' 玩家体验受损类型（D）
                    If UBound(parts) >= 2 Then
                        ws.Cells(i, colDamage).Value = Trim$(parts(2))
                    End If
                Else
                    ws.Cells(i, colPolarity).Value = "调用失败"
                End If
            End If
            
            ' 这行算已处理（无论是否已有标注）
            doneCount = doneCount + 1
            
            ' 每处理若干行更新一次进度（比如每 5 行）
            If doneCount Mod 5 = 0 Or i = lastRow Then
                Application.StatusBar = _
                    "正在处理：第 " & doneCount & " / " & totalRows & " 行……"
            End If
        End If
    Next i
    
    ' 处理完所有行后，恢复状态栏并提示
    Application.StatusBar = False
    MsgBox "处理完成！共遍历 " & totalRows & " 行差评数据。", vbInformation, "DeepSeek 情感分析"
End Sub

' ================= 调用 DeepSeek API =================
Function DeepSeekAnalyzeTranslation(reviewText As String) As String
    Dim http As Object
    Dim url As String
    Dim apiKey As String
    Dim sysPrompt As String
    Dim payload As String
    Dim respText As String
    Dim resultContent As String
    
    ' TODO：改成你的真实 API Key（不要在外面截图/发送）
    apiKey = "sk-a85e15d664c8461c9a98d0048ec4c1fc"
    
    If apiKey = "" Or apiKey = "YOUR_DEEPSEEK_API_KEY_HERE" Then
        MsgBox "请先在代码中设置正确的 DeepSeek API Key。", vbExclamation
        DeepSeekAnalyzeTranslation = ""
        Exit Function
    End If
    
    url = "https://api.deepseek.com/chat/completions"
    
    ' —— 系统提示词：只分析翻译相关部分 ——
    sysPrompt = ""
    sysPrompt = sysPrompt & "你是一个专门分析游戏文本本地化问题的情感分析助手。" _
                          & "用户给你的是玩家对某游戏的差评，其中只有部分内容在抱怨翻译质量。" _
                          & "你的任务：只分析与翻译/本地化质量相关的那些句子，不要被数值、玩法、Bug 等其他内容干扰。" _
                          & "如果玩家完全没有提到翻译或本地化，则认为对翻译“无明显情绪评价”。" _
                          & "情感极性只允许：负面、正面、中性。" _
                          & "情绪类型只输出 1 个：愤怒、失望、困惑、其他。" _
                          & "玩家受损体验类别从以下 6 类中，选出 0~3 个，用英文逗号连接：" _
                          & "comprehension（看不懂任务目标/技能说明/道具效果）," _
                          & "immersion_narrative（台词出戏/语气不对/世界观术语混乱）," _
                          & "aesthetic_tone（机翻味重/风格不统一/文辞粗糙违和）," _
                          & "cultural_issues（忽略本地文化/冒犯或误读/完全直译无本地化）," _
                          & "usability_playability（UI 文本被截断/按钮翻译不一致/提示含糊影响操作效率和流程）," _
                          & "trust（玩家认为开发或发行方不重视该语言区）。" _
                          & "如果完全看不出和翻译/本地化有关，则情感极性统一输出“中性”，情绪类型统一输出“其他”，类别输出“none”。" _
                          & "最终只输出一行字符串，用“|”分隔三个字段，格式严格为：" _
                          & "情感极性|情绪类型|受损体验类别列表。" _
                          & "受损体验类别列表中多个类别用英文逗号(,)分隔；如果没有翻译相关问题则写 none。" _
                          & "示例 1：负面|愤怒|comprehension,trust。" _
                          & "示例 2：中性|其他|none。" _
                          & "不要输出解释，不要换行，也不要添加其它文字。"
    
    ' —— 组装请求 JSON ——
    payload = "{""model"":""deepseek-chat"",""stream"":false,""messages"":["
    payload = payload & "{""role"":""system"",""content"":""" & JsonEscape(sysPrompt) & """},"
    payload = payload & "{""role"":""user"",""content"":""" & JsonEscape(reviewText) & """}"
    payload = payload & "]}"
    
    On Error GoTo ErrHandler
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send payload
    
    If http.readyState = 4 Then
        If http.Status = 200 Then
            respText = http.responseText
            resultContent = ExtractAssistantContent(respText)
            DeepSeekAnalyzeTranslation = Trim$(resultContent)
        Else
            Debug.Print "HTTP 错误状态码: " & http.Status & " - " & http.responseText
            DeepSeekAnalyzeTranslation = ""
        End If
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "DeepSeek 调用异常: " & Err.Number & " - " & Err.Description
    DeepSeekAnalyzeTranslation = ""
End Function

' ================= JSON 转义（防止报错） =================
Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function

' ================= 从返回 JSON 中抽出 content =================
Private Function ExtractAssistantContent(ByVal json As String) As String
    Dim idxRole As Long
    Dim idxContent As Long
    Dim startPos As Long
    Dim i As Long
    Dim ch As String
    
    idxRole = InStr(1, json, """role"":""assistant""", vbTextCompare)
    If idxRole = 0 Then Exit Function
    
    idxContent = InStr(idxRole, json, """content"":""", vbTextCompare)
    If idxContent = 0 Then Exit Function
    
    startPos = idxContent + Len("""content"":""")
    i = startPos
    
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        If ch = """" Then
            If Mid$(json, i - 1, 1) <> "\" Then
                Exit Do
            End If
        End If
        i = i + 1
    Loop
    
    ExtractAssistantContent = Mid$(json, startPos, i - startPos)
End Function


