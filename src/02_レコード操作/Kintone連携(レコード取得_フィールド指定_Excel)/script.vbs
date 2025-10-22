' -----------------------------------------------------------------------
' Main
' -----------------------------------------------------------------------
'Kintone連携(アクセストークン)が配置されているかの確認
If IsEmpty(kntn_client_id) Then
    Err.Raise 1, "", "WinActor for kintone ver1.1.1 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
End If

If IsEmpty(kntn_userAgent) Then
    Err.Raise 1, "", "WinActor for kintone ver1.1.1 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
End If

' トークンのチェック
Call Kntn_CheckAccessToken(kntn_client_id)

' Kintoneのレコードを取得
Call Kntn_GetCustomFieldsRecords()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_GetCustomFieldsRecords()
    Dim kntn_api_uri
    Dim responseText
    Dim sendData
    Dim kntn_app_id
    Dim excelFilePath
    
    kntn_app_id = !*アプリID!
    kntn_guestspace_id = !ゲストスペースID!
    nameOrCode = !*ヘッダー|フィールド名,フィールドコード!
    fields = !*取得するフィールド!
    kugirimoji = !*区切り文字!
    excelFilePath = !*保存先Excelファイルパス!
    excelSheetName = !*シート名!
    blnCloseExcel = !*保存先Excelファイルを閉じる|閉じる,閉じない!
    
    If blnCloseExcel = "閉じる" Then
        blnCloseExcel = True
    Else
        blnCloseExcel = False
    End If
    
    canOverWriteFile = !*保存先にExcelファイルが既に存在するとき|上書き,エラー!
    
    Dim objFSO
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(excelFilePath) = True And canOverWriteFile = "エラー" Then
        Err.Raise 1, "", "保存先にExcelファイルが既に存在しています。既存ファイルを移動する、または保存先Excelファイルパスを変更してください。"
    End If
    
    recordCount = $*取得レコード件数$
    field1 = !項目1!
    condition1 = !項目1の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
    value1 = !項目1の値!
    field2 = !項目2!
    condition2 = !項目2の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
    value2 = !項目2の値!
    field3 = !項目3!
    condition3 = !項目3の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
    value3 = !項目3の値!
    field4 = !項目4!
    condition4 = !項目4の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
    value4 = !項目4の値!
    field5 = !項目5!
    condition5 = !項目5の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
    value5 = !項目5の値!
    freequery = !カスタムクエリ!
    
    If kntn_app_id = "" Then
        Err.Raise 1, "", "「アプリID」の入力は必須です。"
    End If
    
    If fields = "" Then
        Err.Raise 1, "", "「取得するフィールド名またはフィールドコード」の入力は必須です。"
    End If
    
    If kugirimoji = "" Then
        Err.Raise 1, "", "「区切り文字」の入力は必須です。"
    End If
    
    '区切り文字がvbCrlfなら改行区切りとする
    If kugirimoji = "vbCrlf" Then
        kugirimoji = vbCrLf
    End If
    
    If excelFilePath = "" Then
        Err.Raise 1, "", "「保存先Excelファイルパス」の入力は必須です。"
    End If
    
    Select Case objFSO.GetExtensionName(excelFilePath)
        Case "xlsx", "xlsm", "xls"
            ' OK
        Case Else
            Err.Raise 1, "", "「保存先Excelファイルパス」の拡張子は「xlsx」「xlsm」「xls」のいずれかを指定してください。"
    End Select
    
    If excelSheetName = "" Then
        Err.Raise 1, "", "「シート名」の入力は必須です。"
    End If
    
    Dim array_fields
    If kugirimoji = vbCrLf Then
        'いったんvblfにしてからvbcrlfに統一
        fields = Replace(fields, vbCrLf, vbLf)
        fields = Replace(fields, vbLf, vbCrLf)
    End If
    array_fields = Split(fields, kugirimoji)
    
    'フィールド情報一覧の情報をAPIで取得する。
    Dim json_fieldsInfo, json_properties
    Dim fieldsInfoText
    fieldsInfoText = Kntn_GetFieldsInfo(kntn_app_id, kntn_guestspace_id)
    Set json_fieldsInfo = kntn_ScriptEngine.CodeObject.Parse(fieldsInfoText)
    Set json_properties = json_fieldsInfo.properties
    Dim array_fieldsinfo(), array_header()
    
    'property項目のkey一覧を取得する
    Dim KeysObject
    Set KeysObject = kntn_ScriptEngine.Run("getKeys", json_properties)
    
    ReDim array_fieldsinfo(2, 0)
    array_fieldsinfo(0, 0) = "$id"
    array_fieldsinfo(1, 0) = "レコードID"
    
    If nameOrCode = "フィールド名" Then
        rowNum = 1
        recordIdColName = "レコードID"
    Else
        rowNum = 0
        recordIdColName = "$id"
    End If
    
    colnum = 1
    fields = """$id"""
    
    For i = 0 To UBound(array_fields)
        target_field = array_fields(i)
        existsField = False
        
        '該当のフィールドの情報を取得する。存在しない場合はエラー
        Dim keysLength
        keysLength = KeysObject.Length
        
        For j = 0 To keysLength - 1
            'JScript9対応: getArrayItemでキーを取得
            Dim key
            key = kntn_ScriptEngine.Run("getArrayItem", KeysObject, j)
            
            'フィールド名・フィールドコード・フィールドタイプを取得する。
            Dim json_field
            Set json_field = kntn_ScriptEngine.Run("getProperty", json_properties, key)
            fieldtype = json_field.type
            fieldname = json_field.label
            fieldcode = json_field.code
            
            If nameOrCode = "フィールド名" And target_field = fieldname Then
                existsField = True
                Exit For
            ElseIf nameOrCode <> "フィールド名" And target_field = fieldcode Then
                existsField = True
                Exit For
            End If
        Next
        
        If existsField = True Then
            If fieldtype = "SUBTABLE" Then
                Err.Raise 1, "", "フィールド「" & target_field & "」はテーブルのため、情報を取得できません"
            End If
            ReDim Preserve array_fieldsinfo(2, colnum)
            array_fieldsinfo(0, colnum) = fieldcode
            array_fieldsinfo(1, colnum) = fieldname
            array_fieldsinfo(2, colnum) = fieldtype
            colnum = colnum + 1
            fields = fields & ",""" & fieldcode & """"
            
        'レコードID列は1列目に設定するので、フィールドに指定してもスキップする
        ElseIf target_field = recordIdColName Then
            ' スキップ
        Else
            Err.Raise 1, "", "フィールド「" & target_field & "」は存在しません。"
        End If
    Next
    
    'Kintone API のエンドポイント
    If kntn_guestspace_id = "" Then
        kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/records.json"
    Else
        kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/records.json"
    End If
    
    'queryの作成
    If freequery <> "" Then
        query = freequery
    Else
        query = ""
        If field1 <> "" Then
            tmpQuery = Kntn_CreateQuery(field1, condition1, value1, kugirimoji, nameOrCode, json_properties)
            query = tmpQuery
        End If
        
        If field2 <> "" Then
            tmpQuery = Kntn_CreateQuery(field2, condition2, value2, kugirimoji, nameOrCode, json_properties)
            query = query & " and " & tmpQuery
        End If
        
        If field3 <> "" Then
            tmpQuery = Kntn_CreateQuery(field3, condition3, value3, kugirimoji, nameOrCode, json_properties)
            query = query & " and " & tmpQuery
        End If
        
        If field4 <> "" Then
            tmpQuery = Kntn_CreateQuery(field4, condition4, value4, kugirimoji, nameOrCode, json_properties)
            query = query & " and " & tmpQuery
        End If
        
        If field5 <> "" Then
            tmpQuery = Kntn_CreateQuery(field5, condition5, value5, kugirimoji, nameOrCode, json_properties)
            query = query & " and " & tmpQuery
        End If
    End If
    
    Dim array_records
    
    'ヘッダーの配列を格納する辞書型配列の作成
    Dim kntn_rootHeaderArray
    Set kntn_rootHeaderArray = WScript.CreateObject("Scripting.Dictionary")
    
    count = 0
    lastRecordID = ""
    
    'レコードがなくなるまで繰り返す(500件ずつ取得のため)
    Do
        '500件以上あった場合は前回のレコードIDの次から取得する
        If lastRecordID = "" Then
            If query = "" Then
                sendquery = """query"":""order by $id asc limit 500"""
            Else
                sendquery = """query"":""" & query & " order by $id asc limit 500"""
            End If
        Else
            If query = "" Then
                sendquery = """query"":""$id > \""" & lastRecordID & "\"" order by $id asc limit 500"""
            Else
                sendquery = """query"":""" & query & " and $id > \""" & lastRecordID & "\"" order by $id asc limit 500"""
            End If
        End If
        
        ' リクエストデータを作成 Queryがない場合は全データ
        If field1 = "" And freequery = "" Then
            sendData = "{""app"": " & kntn_app_id & "," _
                & """totalCount"":true," _
                & """fields"": [" & fields & "]," _
                & sendquery _
                & "}"
        Else
            sendData = "{""app"": " & kntn_app_id & "," _
                & """totalCount"":true," _
                & """fields"": [" & fields & "]," _
                & sendquery _
                & "}"
        End If
        
        'アクセストークンの有効性を確認
        Call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token)
        
        ' API リクエストのヘッダーとデータを設定
        With WScript.CreateObject("MSXML2.XMLHTTP")
            .Open "Post", kntn_api_uri, False
            .setRequestHeader "Authorization", "Bearer " & kntn_access_token
            .setRequestHeader "Content-Type", "application/json"
            .setRequestHeader "X-HTTP-Method-Override", "GET"
            .setRequestHeader "User-Agent", kntn_userAgent
            .send sendData
            
            ' レスポンステキストを取得
            responseText = .responseText
            statusCode = .status
        End With
        
        ' レスポンスの処理を行う
        If statusCode <> 200 Then
            Err.Raise statusCode, "", "Kintoneのレコード取得操作に失敗しました。" & vbCrLf & _
                "ステータスコード:" & statusCode & vbCrLf & _
                "レスポンス: " & Kntn_GetErrorMessage(responseText)
        End If
        
        'レコードの数を返却する
        Set json = kntn_ScriptEngine.CodeObject.Parse(responseText)
        If count = 0 Then
            SetUmsVariable recordCount, json.totalCount
        End If
        
        '残りレコードが501件以上なら繰り返し取得
        If json.totalCount > 500 Then
            blnNextRecord = True
        Else
            blnNextRecord = False
        End If
        
        'recordsプロパティの取得
        Set json_records = json.records
        
        '1回目のループの場合はヘッダー情報を入力する
        If count = 0 Then
            'headerの情報を取得する。Fileタイプに注意する
            ReDim array_header(0, UBound(array_fieldsinfo, 2))
            colnum = 0
            For i = 0 To UBound(array_fieldsinfo, 2)
                If array_fieldsinfo(2, i) = "FILE" Then
                    '配列の列を増やす
                    ReDim Preserve array_header(0, UBound(array_header, 2) + 2)
                    'FILEタイプはファイル名列とファイルキー列、デフォルトの3列に分ける
                    array_header(0, colnum) = array_fieldsinfo(rowNum, i) & "_Name"
                    colnum = colnum + 1
                    array_header(0, colnum) = array_fieldsinfo(rowNum, i) & "_Key"
                    colnum = colnum + 1
                    array_header(0, colnum) = array_fieldsinfo(rowNum, i)
                Else
                    array_header(0, colnum) = array_fieldsinfo(rowNum, i)
                End If
                colnum = colnum + 1
            Next
            'ヘッダーのみ入力
            Call Kntn_SetArrayToExcel(excelFilePath, excelSheetName, array_header, "A1", False)
        End If
        
        array_records = Kntn_GetMainRecord(array_header, array_fieldsinfo, json_records, outputType, kugirimoji)
        'レコードを設定する。
        Call Kntn_SetArrayToExcel(excelFilePath, excelSheetName, array_records, "", False)
        '最終行から最終レコードIDを取得する。
        lastRecordID = array_records(499, 0)
        count = count + 1
    Loop While blnNextRecord = True
    
    'メインシートを一番左に移動し、閉じる
    Call KNTN_MoveSheet(excelFilePath, excelSheetName, blnCloseExcel)
End Sub