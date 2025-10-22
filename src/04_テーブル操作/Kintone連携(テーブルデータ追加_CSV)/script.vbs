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

' Kintoneの指定アプリのフィールド情報を取得
Call Kntn_AddTableData()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_AddTableData()
    Dim kntn_api_uri
    Dim responseText
    Dim sendData
    Dim kntn_app_id
    Dim excelFilePath
    
    kntn_app_id = !*アプリID!
    kntn_guestspace_id = !ゲストスペースID!
    nameOrCode = !*ヘッダー|フィールド名,フィールドコード!
    fieldtable = !*テーブルのフィールド名またはフィールドコード!
    kugirimoji = !*複数設定値の区切り文字!
    brank = !*ブランクとして認識する値!
    excelFilePath = !*入力ファイルパス!
    excelSheetName = !*シート名!
    outputFilePath = !*処理結果ファイルパス!
    canOverWriteFile = !*処理結果ファイルが既に存在するとき|上書き,エラー!
    
    Dim objFSO
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(outputFilePath) = True And canOverWriteFile = "エラー" Then
        Err.Raise 1, "", "処理結果ファイルが既に存在しています。既存ファイルを移動する、または処理結果ファイルパスを変更してください。"
    End If
    
    successRecordCount = 0
    errorRecordCount = 0
    
    Dim recordCount
    If kntn_app_id = "" Then
        Err.Raise 1, "", "「アプリID」の入力は必須です。"
    End If
    
    If fieldtable = "" Then
        Err.Raise 1, "", "「テーブルのフィールド名またはフィールドコード」の入力は必須です。"
    End If
    
    If kugirimoji = "" Then
        Err.Raise 1, "", "「複数設定値の区切り文字」の入力は必須です。"
    End If
    
    '区切り文字がvbCrlfなら改行区切りとする
    If kugirimoji = "vbCrlf" Then
        kugirimoji = vbCrLf
    End If
    
    If brank = "" Then
        Err.Raise 1, "", "「ブランクとして認識する値」の入力は必須です。"
    End If
    
    If excelFilePath = "" Then
        Err.Raise 1, "", "「入力ファイルパス」の入力は必須です。"
    End If
    
    If excelSheetName = "" Then
        Err.Raise 1, "", "「シート名」の入力は必須です。"
    End If
    
    If outputFilePath = "" Then
        Err.Raise 1, "", "「処理結果ファイルパス」の入力は必須です。"
    End If
    
    'Kintone API のエンドポイント
    If kntn_guestspace_id = "" Then
        kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/record.json"
    Else
        kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/record.json"
    End If
    
    'フィールド情報一覧の情報をAPIで取得する。
    Dim json_fieldsInfo, json_properties
    ' 修正1: CodeObject.Parse → Run("Parse", ...)
    Set json_fieldsInfo = kntn_ScriptEngine.Run("Parse", Kntn_GetFieldsInfo(kntn_app_id, kntn_guestspace_id))
    Set json_properties = json_fieldsInfo.properties
    
    '該当のテーブルフィールドの情報を取得する。
    On Error Resume Next
    Dim json_TableInfo
    Set json_TableInfo = Kntn_getFieldJson(fieldtable, nameOrCode, json_properties)
    If json_TableInfo Is Nothing Then
        On Error GoTo 0
        Err.Raise 1, "", "テーブルのフィールド名またはフィールドコードが正しくありません。"
    End If
    tableName = json_TableInfo.label
    ' 修正2: Code → code (小文字)
    tableCode = json_TableInfo.code
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise 1, "", "テーブルのフィールド名またはフィールドコードが正しくありません。"
    End If
    On Error GoTo 0
    
    'Lookupのmappingの項目の一覧を取得する
    Dim array_fieldsinfo, kntn_rootArray, lookUpArray, array_header
    Set kntn_rootArray = WScript.CreateObject("Scripting.Dictionary")
    lookUpArray = Kntn_CreateLookUpArray(json_TableInfo.fields)
    
    '入力ファイルのデータを配列に格納する
    Dim Array_EntryData
    Array_EntryData = Kntn_GetArraybyExcel(excelFilePath, excelSheetName)
    
    'ヘッダー情報をまとめたRootArrayを求める
    ' 修正3: outputType変数を明示的に定義
    outputType = "全フィールド取得"
    Call Kntn_getHeaderArray(json_TableInfo.fields, outputType, nameOrCode, kntn_rootArray, fieldtable, lookUpArray, True)
    
    If nameOrCode = "フィールド名" Then
        rowNum = 1
        recordIdColName = "レコードID"
    Else
        rowNum = 0
        recordIdColName = "$id"
    End If
    
    'レコードIDまたは$id行が存在しない場合はエラー出力
    If Array_EntryData(1, 1) <> recordIdColName Then
        Err.Raise 1, "", "入力ファイルに" & recordIdColName & "フィールドが存在しません。"
    End If
    
    array_header = kntn_rootArray.Item(fieldtable)
    
    'フィールドコードが入力ファイルの何列目かを辞書型配列にまとめる(毎回求めるのは大変なため)
    Dim headerDictionary
    Set headerDictionary = WScript.CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(Array_EntryData, 2)
        tmpHeader = Array_EntryData(1, i)
        If tmpHeader = "rowId" Or tmpHeader = "テーブル行ID" Then
            headerDictionary.Add "rowId", i
        Else
            For j = 0 To UBound(array_header, 2)
                If array_header(rowNum, j) = tmpHeader Then
                    '辞書にフィールドコードと列名を追加する
                    headerDictionary.Add array_header(0, j), i
                    Exit For
                End If
            Next
        End If
    Next
    
    'レコードIDの一覧を辞書型配列に持たせる。
    Dim recordIdsDictionary
    Set recordIdsDictionary = WScript.CreateObject("Scripting.Dictionary")
    Dim array_Rows()
    For i = 2 To UBound(Array_EntryData, 1)
        recordId = Array_EntryData(i, 1)
        ' 修正4: 空レコードIDのフィルタリング
        If recordId <> "" And recordId <> brank Then
            If recordIdsDictionary.Exists(recordId) Then
                exist_array_Rows = recordIdsDictionary.Item(recordId)
                idx = UBound(exist_array_Rows) + 1
                ReDim Preserve exist_array_Rows(idx)
                exist_array_Rows(idx) = i
                recordIdsDictionary.Item(recordId) = exist_array_Rows
            Else
                ReDim array_Rows(0)
                array_Rows(0) = i
                'レコードIDと行一覧を持たせる
                recordIdsDictionary.Add recordId, array_Rows
            End If
        End If
    Next
    
    '処理結果の配列とエラー出力の配列のヘッダー部のみ作成
    Dim Array_OutputData(), Array_ErrorData()
    ReDim Array_OutputData(UBound(Array_EntryData, 1) - 1, UBound(Array_EntryData, 2) - 1 + 3)
    ReDim Array_ErrorData(UBound(Array_EntryData, 1) - 1, UBound(Array_EntryData, 2) - 1)
    
    For i = 1 To UBound(Array_EntryData, 2)
        Array_OutputData(0, i - 1) = Array_EntryData(1, i)
        Array_ErrorData(0, i - 1) = Array_EntryData(1, i)
    Next
    
    resultCol = UBound(Array_EntryData, 2)
    Array_OutputData(0, resultCol) = "処理結果"
    Array_OutputData(0, resultCol + 1) = "処理実行日時"
    Array_OutputData(0, resultCol + 2) = "エラー内容"
    outputRow = 1
    erroutputRow = 1
    
    'レコードIDの一覧を繰り返す
    ' 修正5: VBScript DictionaryのKeysを配列として直接取得
    Dim recordIdKeysArray
    recordIdKeysArray = recordIdsDictionary.Keys
    Dim keysLength
    keysLength = UBound(recordIdKeysArray) + 1
    
    For keyIdx = 0 To keysLength - 1
        On Error Resume Next
        flgFirstRow = True
        errmsg = ""
        
        ' 修正6: VBScript Dictionaryから直接キーを取得
        recordId = recordIdKeysArray(keyIdx)
        
        'アクセストークンの有効性を確認
        Call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token)
        
        '既存のIDのレコード情報を取得する。
        With WScript.CreateObject("MSXML2.XMLHTTP")
            .Open "Get", kntn_api_uri & "?app=" & kntn_app_id & "&id=" & recordId, False
            .setRequestHeader "Authorization", "Bearer " & kntn_access_token
            .setRequestHeader "User-Agent", kntn_userAgent
            .send
            
            ' レスポンステキストを取得
            responseText = .responseText
            statusCode = .status
            ' 修正7: CodeObject.Parse → Run("Parse", ...)
            Set json = kntn_ScriptEngine.Run("Parse", responseText)
            
            If .status = 200 Then
                '既存のテーブルのデータをjsonパラメータにする。
                Set json_record = json.record
                Set json_Table = kntn_ScriptEngine.Run("getProperty", json_record, tableCode)
                Set json_Table = json_Table.value
                recordInfo = ""
                
                'JScript9対応: json_Tableの長さを取得
                ' 修正8: .Length → Run("getArrayLength", ...)
                Dim tableLength
                tableLength = kntn_ScriptEngine.Run("getArrayLength", json_Table)
                
                For tableIdx = 0 To tableLength - 1
                    'JScript9対応: 配列要素を取得
                    Dim json_Row
                    Set json_Row = kntn_ScriptEngine.Run("getArrayItem", json_Table, tableIdx)
                    rowId = json_Row.id
                    
                    tmpRecordInfo = "{""id"":""" & rowId & """,""value"":["
                    Set json_RowValue = json_Row.value
                    rowValues = ""
                    
                    For i = 0 To UBound(array_header, 2)
                        fieldcode = array_header(0, i)
                        fieldtype = array_header(2, i)
                        If fieldcode = "" Then Exit For
                        fileKeys = ""
                        
                        '既存の値を取得する。
                        If fieldtype = "FILE" Then
                            '既存のファイルキー一覧を取得する
                            fieldvalue = Kntn_getFieldValue("FILE_KEY", fieldcode, json_RowValue, kugirimoji)
                            
                            '区切り文字に従って区切り配列化
                            array_ExistsFiles = Kntn_SplitFiles(fieldvalue, kugirimoji)
                            
                            For idx = 0 To UBound(array_ExistsFiles)
                                If fileKeys = "" Then
                                    fileKeys = "{""fileKey"":""" & array_ExistsFiles(idx) & """}"
                                Else
                                    fileKeys = fileKeys & ",{""fileKey"":""" & array_ExistsFiles(idx) & """}"
                                End If
                            Next
                        Else
                            fieldvalue = Kntn_getFieldValue(fieldtype, fieldcode, json_RowValue, kugirimoji)
                        End If
                        
                        'json形式に変換する
                        If fieldtype = "FILE" Then
                            keyAndValue = """" & fieldcode & """:{""value"":[" & fileKeys & "]}"
                        Else
                            keyAndValue = Kntn_CreateJsonKeyAndValue(fieldcode, fieldtype, fieldvalue, "", kugirimoji)
                        End If
                        
                        If rowValues = "" Then
                            rowValues = keyAndValue
                        ElseIf keyAndValue <> "" Then
                            rowValues = rowValues & "," & keyAndValue
                        End If
                    Next
                    
                    tmpRecordInfo = "{""id"":" & rowId & ",""value"":{" & rowValues & "}}"
                    If recordInfo = "" Then
                        recordInfo = tmpRecordInfo
                    Else
                        recordInfo = recordInfo & "," & tmpRecordInfo
                    End If
                Next
                
                '追加したいテーブルの行のデータをjsonに追加する
                Array_RowNumAndRowId = recordIdsDictionary.Item(recordId)
                
                For Each rowNum In Array_RowNumAndRowId
                    rowValues = ""
                    For i = 2 To UBound(array_header, 2)
                        fieldcode = array_header(0, i)
                        fieldtype = array_header(2, i)
                        If fieldcode = "" Then Exit For
                        
                        '入力ファイルに存在しない列は初期値を設定する。
                        If headerDictionary.Exists(fieldcode) Then
                            fieldvalue = Array_EntryData(rowNum, headerDictionary.Item(fieldcode))
                        Else
                            fieldvalue = ""
                        End If
                        
                        '空白なら初期値のため、設定しない。
                        If fieldvalue <> "" Then
                            If fieldtype = "FILE" Then
                                '配列に分解する。
                                array_Files = Kntn_SplitFiles(fieldvalue, kugirimoji)
                                For idx = 0 To UBound(array_Files)
                                    'ファイルを一時領域にアップロードし、ファイルキーを取得する。
                                    array_Files(idx) = Kntn_tmpUpload(array_Files(idx))
                                Next
                                
                                fileKeys = ""
                                For idx = 0 To UBound(array_Files)
                                    If fileKeys = "" Then
                                        fileKeys = "{""fileKey"":""" & array_Files(idx) & """}"
                                    Else
                                        fileKeys = fileKeys & ",{""fileKey"":""" & array_Files(idx) & """}"
                                    End If
                                Next
                                fieldKeyValue = """" & fieldcode & """:{""value"":[" & fileKeys & "]}"
                            Else
                                fieldKeyValue = Kntn_CreateJsonKeyAndValue(fieldcode, fieldtype, fieldvalue, brank, kugirimoji)
                            End If
                            
                            If rowValues = "" Then
                                rowValues = fieldKeyValue
                            ElseIf fieldKeyValue <> "" Then
                                rowValues = rowValues & "," & fieldKeyValue
                            End If
                        End If
                    Next
                    
                    tmpRecordInfo = "{""value"":{" & rowValues & "}}"
                    If recordInfo = "" Then
                        recordInfo = tmpRecordInfo
                    Else
                        recordInfo = recordInfo & "," & tmpRecordInfo
                    End If
                Next
                
                sendData = "{""app"":" & kntn_app_id & ",""id"":" & recordId & ",""record"": {""" _
                    & tableCode & """:{""value"":[" & recordInfo & "]}}}"
                
                If Err.Number <> 0 Then
                    errmsg = Err.Description
                    statusCode = 0
                Else
                    'アクセストークンの有効性を確認
                    Call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token)
                    
                    '更新APIを呼び出す。
                    .Open "Put", kntn_api_uri, False
                    .setRequestHeader "Authorization", "Bearer " & kntn_access_token
                    .setRequestHeader "Content-Type", "application/json"
                    .setRequestHeader "User-Agent", kntn_userAgent
                    .send sendData
                    
                    ' レスポンステキストを取得
                    responseText = .responseText
                    statusCode = .status
                End If
            End If
            
            On Error GoTo 0
            'StatusCodeが200なら成功。
            If statusCode = 200 Then
                successRecordCount = successRecordCount + 1
                
                '各行のデータを処理結果の配列に記載する。同一レコードの頭の行のみに結果を記載する。
                For Each rowNum In recordIdsDictionary.Item(recordId)
                    '先頭行のみ結果を記載する
                    If flgFirstRow = True Then
                        Array_OutputData(outputRow, resultCol) = "成功"
                        Array_OutputData(outputRow, resultCol + 1) = Kntn_GetNow
                        Array_OutputData(outputRow, resultCol + 2) = ""
                        flgFirstRow = False
                    Else
                        Array_OutputData(outputRow, resultCol) = "-"
                        Array_OutputData(outputRow, resultCol + 1) = "-"
                        Array_OutputData(outputRow, resultCol + 2) = "-"
                    End If
                    
                    'OutPutDataは既存と同じデータを配置
                    For j = 1 To UBound(Array_EntryData, 2)
                        Array_OutputData(outputRow, j - 1) = Array_EntryData(rowNum, j)
                    Next
                    outputRow = outputRow + 1
                Next
            ElseIf errmsg = "" Then
                errmsg = "Kintoneのテーブル追加操作に失敗しました。" & vbCrLf & _
                    "ステータスコード:" & statusCode & vbCrLf & _
                    "レスポンス: " & Kntn_GetErrorMessage(responseText)
            End If
            
            If errmsg <> "" Then
                errorRecordCount = errorRecordCount + 1
                '各行のデータを処理結果の配列に記載する。同一レコードの頭の行のみに結果を記載する。
                For Each rowNum In recordIdsDictionary(recordId)
                    '先頭行のみ結果を記載する
                    If flgFirstRow = True Then
                        Array_OutputData(outputRow, resultCol) = "失敗"
                        Array_OutputData(outputRow, resultCol + 1) = Kntn_GetNow
                        Array_OutputData(outputRow, resultCol + 2) = errmsg
                        flgFirstRow = False
                    Else
                        Array_OutputData(outputRow, resultCol) = "-"
                        Array_OutputData(outputRow, resultCol + 1) = "-"
                        Array_OutputData(outputRow, resultCol + 2) = "-"
                    End If
                    
                    'OutPutDataは既存と同じデータを配置
                    For j = 1 To UBound(Array_EntryData, 2)
                        Array_OutputData(outputRow, j - 1) = Array_EntryData(rowNum, j)
                        Array_ErrorData(erroutputRow, j - 1) = Array_EntryData(rowNum, j)
                    Next
                    outputRow = outputRow + 1
                    erroutputRow = erroutputRow + 1
                Next
            End If
        End With
    Next
    
    SetUmsVariable $*成功レコード件数$, successRecordCount
    SetUmsVariable $*失敗レコード件数$, errorRecordCount
    
    '処理結果ファイルを作成する
    Call Kntn_SetArrayToExcel(outputFilePath, "エラーデータ", Array_ErrorData, "A1", False)
    Call Kntn_SetArrayToExcel(outputFilePath, "処理結果", Array_OutputData, "A1", True)
End Sub