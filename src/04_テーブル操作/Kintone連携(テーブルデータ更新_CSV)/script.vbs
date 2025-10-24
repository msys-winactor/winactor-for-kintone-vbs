' -----------------------------------------------------------------------
' Main
' -----------------------------------------------------------------------
'Kintone連携(アクセストークン)が配置されているかの確認
If IsEmpty(kntn_client_id) Then
    Err.Raise 1, "", "WinActor for kintone ver1.1.0 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
End If

If IsEmpty(kntn_userAgent) Then
    Err.Raise 1, "", "WinActor for kintone ver1.1.0 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
End If

' トークンのチェック
Call Kntn_CheckAccessToken(kntn_client_id)

Call Kntn_updateTableDataByCsv()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_updateTableDataByCsv()
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
    csvFilePath = !*入力CSVファイルパス!
    charcode = !*文字コード|shift-jis,utf-8!
    outputFolderPath = !*処理結果出力フォルダ!
    
    successRecordCount = 0
    errorRecordCount = 0
    
    Dim recordCount
    If kntn_app_id = "" Then
        Err.Raise 1, "", "アプリIDが指定されていません。"
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
    
    If csvFilePath = "" Then
        Err.Raise 1, "", "「入力CSVファイルパス」の入力は必須です。"
    End If
    
    If outputFolderPath = "" Then
        Err.Raise 1, "", "「処理結果出力フォルダ」の入力は必須です。"
    End If
    
    Dim objFSO
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FileExists(csvFilePath) Then
        Err.Raise 1, "", "入力CSVファイルが存在しません。ファイルパスを確認してください。"
    End If
    
    'いったんファイル名の.csvは削除する
    csvFileName = objFSO.GetFileName(csvFilePath)
    extension = objFSO.GetExtensionName(csvFileName)
    If StrComp(extension, "csv", 0) = 0 Then
        csvFileName = Left(csvFileName, Len(csvFileName) - (Len(extension) + 1))
    End If
    
    If Not objFSO.FolderExists(outputFolderPath) Then
        Call KNTN_CreateIntermediateFolders(outputFolderPath)
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
    Array_EntryData = KNTN_ReadCsv(csvFilePath, charcode, 1)
    
    'ヘッダー情報をまとめたRootArrayを求める
    ' 修正3: outputType変数を明示的に定義
    outputType = "全フィールド取得"
    Call Kntn_getHeaderArray(json_TableInfo.fields, outputType, nameOrCode, kntn_rootArray, fieldtable, lookUpArray, True)
    
    If nameOrCode = "フィールド名" Then
        rowNum = 1
        recordIdColName = "レコードID"
        rowIdColName = "テーブル行ID"
    Else
        rowNum = 0
        recordIdColName = "$id"
        rowIdColName = "rowId"
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
    
    'レコードIDの一覧とそれに対応する行とテーブルIDを辞書型配列に持たせる。
    Dim recordIdsDictionary
    Set recordIdsDictionary = WScript.CreateObject("Scripting.Dictionary")
    Dim array_records()
    For i = 2 To UBound(Array_EntryData, 1)
        recordId = Array_EntryData(i, 1)
        rowId = Array_EntryData(i, headerDictionary.Item("rowId"))
        If recordIdsDictionary.Exists(recordId) Then
            exist_array_records = recordIdsDictionary.Item(recordId)
            idx = UBound(exist_array_records, 2) + 1
            ReDim Preserve exist_array_records(1, idx)
            exist_array_records(0, idx) = i
            exist_array_records(1, idx) = rowId
            recordIdsDictionary.Item(recordId) = exist_array_records
        Else
            ReDim array_records(1, 0)
            array_records(0, 0) = i
            array_records(1, 0) = rowId
            'レコード行一覧を持たせる
            recordIdsDictionary.Add recordId, array_records
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
    
    Dim array_existsRowIds()
    
    'レコードIDの一覧を繰り返す
    ' 修正4: VBScript DictionaryのKeysを配列として直接取得
    Dim recordIdKeysArray
    recordIdKeysArray = recordIdsDictionary.Keys
    Dim keysLength
    keysLength = UBound(recordIdKeysArray) + 1
    
    For keyIdx = 0 To keysLength - 1
        On Error Resume Next
        flgFirstRow = True
        errmsg = ""
        
        ' 修正5: VBScript Dictionaryから直接キーを取得
        recordId = recordIdKeysArray(keyIdx)
        
        '処理した行IDの一覧をし配列化する。
        ReDim array_existsRowIds(0)
        
        'アクセストークンの有効性を確認
        Call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token)
        
        '既存のIDのレコード情報を取得する。
        With WScript.CreateObject("MSXML2.XMLHTTP")
            .Open "Get", kntn_api_uri & "?app=" & kntn_app_id & "&id=" & recordId, False
            If kntn_proxy_url <> "" Then .setProxy 2, kntn_proxy_url
            .setRequestHeader "User-Agent", kntn_userAgent
            .setRequestHeader "Authorization", "Bearer " & kntn_access_token
            .send
            
            ' レスポンステキストを取得
            responseText = .responseText
            statusCode = .status
            ' 修正6: CodeObject.Parse → Run("Parse", ...)
            Set json = kntn_ScriptEngine.Run("Parse", responseText)
        End With
        
        If statusCode = 200 Then
            '既存のテーブルのデータをjsonパラメータにする。
            Set json_record = json.record
            Set json_Table = kntn_ScriptEngine.Run("getProperty", json_record, tableCode)
            Set json_Table = json_Table.value
            recordInfo = ""
            
            Array_RowNumAndRowId = recordIdsDictionary.Item(recordId)
            
            ' 修正7: .Length → Run("getArrayLength", ...)
            Dim tableLengthCheck
            tableLengthCheck = kntn_ScriptEngine.Run("getArrayLength", json_Table)
            
            If tableLengthCheck > 0 Then
                'JScript9対応: json_Tableの長さを取得
                Dim tableLength
                tableLength = tableLengthCheck
                
                For tableIdx = 0 To tableLength - 1
                    'JScript9対応: 配列要素を取得
                    Dim json_Row
                    Set json_Row = kntn_ScriptEngine.Run("getArrayItem", json_Table, tableIdx)
                    rowId = json_Row.id
                    Set json_RowValue = json_Row.value
                    rowValues = ""
                    
                    '更新対象のrowIDかを確認する
                    rowNum = -1
                    For i = 0 To UBound(Array_RowNumAndRowId, 2)
                        If CLng(Array_RowNumAndRowId(1, i)) = CLng(rowId) Then
                            rowNum = Array_RowNumAndRowId(0, i)
                            '存在した行ID一覧の配列に記載する
                            array_existsRowIds(UBound(array_existsRowIds)) = rowId
                            ReDim Preserve array_existsRowIds(UBound(array_existsRowIds) + 1)
                            Exit For
                        End If
                    Next
                    
                    For i = 2 To UBound(array_header, 2)
                        fieldcode = array_header(0, i)
                        fieldtype = array_header(2, i)
                        fileKeys = ""
                        If fieldcode = "" Then Exit For
                        
                        '既存の値を取得する。
                        If fieldtype = "FILE" Then
                            fieldvalue = Kntn_getFieldValue("FILE_KEY", fieldcode, json_RowValue, kugirimoji)
                            '既存のファイルキー一覧を取得する
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
                        
                        '更新対象のrowIDのとき
                        If rowNum >= 0 Then
                            If headerDictionary.Exists(fieldcode) Then
                                tmpData = Array_EntryData(rowNum, headerDictionary.Item(fieldcode))
                                If tmpData <> "" Then
                                    If fieldtype = "FILE" Then
                                        '配列に分解する。
                                        array_Files = Kntn_SplitFiles(tmpData, kugirimoji)
                                        For idx = 0 To UBound(array_Files)
                                            'ファイルを一時領域にアップロードし、ファイルキーを取得する。
                                            array_Files(idx) = Kntn_tmpUpload(array_Files(idx))
                                        Next
                                        '今回アップロードしたファイルを追加する。
                                        For idx = 0 To UBound(array_Files)
                                            If fileKeys = "" Then
                                                fileKeys = "{""fileKey"":""" & array_Files(idx) & """}"
                                            Else
                                                fileKeys = fileKeys & ",{""fileKey"":""" & array_Files(idx) & """}"
                                            End If
                                        Next
                                        keyAndValue = """" & fieldcode & """:{""value"":[" & fileKeys & "]}"
                                    Else
                                        keyAndValue = Kntn_CreateJsonKeyAndValue(fieldcode, fieldtype, tmpData, brank, kugirimoji)
                                    End If
                                Else
                                    'セルが空白ならそのまま値を設定する。
                                    keyAndValue = Kntn_CreateJsonKeyAndValue(fieldcode, fieldtype, fieldvalue, "", kugirimoji)
                                End If
                            Else
                                keyAndValue = Kntn_CreateJsonKeyAndValue(fieldcode, fieldtype, fieldvalue, "", kugirimoji)
                            End If
                        Else
                            'json形式に変換する
                            If fieldtype = "FILE" Then
                                keyAndValue = """" & fieldcode & """:{""value"":[" & fileKeys & "]}"
                            Else
                                keyAndValue = Kntn_CreateJsonKeyAndValue(fieldcode, fieldtype, fieldvalue, "", kugirimoji)
                            End If
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
            End If
            
            '存在しない行IDを連携していないか確認する
            notExistsRowIds = ""
            For i = 0 To UBound(Array_RowNumAndRowId, 2)
                blnExistsRowId = False
                rowId = CLng(Array_RowNumAndRowId(1, i))
                '行IDが空なら処理を抜ける
                If rowId = 0 Then
                    Exit For
                End If
                
                For j = 0 To UBound(array_existsRowIds)
                    If CLng(array_existsRowIds(j)) = rowId Then
                        blnExistsRowId = True
                        Exit For
                    End If
                Next
                
                'もし存在しない行IDなら、エラーメッセージに追加する。
                If blnExistsRowId = False Then
                    notExistsRowIds = notExistsRowIds & "," & rowId
                End If
            Next
            
            If Err.Number <> 0 Then
                errmsg = Err.Description
                statusCode = 0
            ElseIf rowId = 0 Then
                errmsg = "テーブルの行IDがブランクのデータが存在します。"
                statusCode = 0
            ElseIf notExistsRowIds <> "" Then
                errmsg = "テーブルの行ID『" & Right(notExistsRowIds, Len(notExistsRowIds) - 1) & "』が存在しません。行IDを確認してください。"
                statusCode = 0
            Else
                'アクセストークンの有効性を確認
                Call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token)
                sendData = "{""app"":" & kntn_app_id & ",""id"":" & recordId & ",""record"": {""" _
                    & tableCode & """:{""value"":[" & recordInfo & "]}}}"
                
                '更新APIを呼び出す。
                With WScript.CreateObject("MSXML2.XMLHTTP")
                    .Open "Put", kntn_api_uri, False
                    If kntn_proxy_url <> "" Then .setProxy 2, kntn_proxy_url
                    .setRequestHeader "Authorization", "Bearer " & kntn_access_token
                    .setRequestHeader "Content-Type", "application/json"
                    .setRequestHeader "User-Agent", kntn_userAgent
                    .send sendData
                    ' レスポンステキストを取得
                    responseText = .responseText
                    statusCode = .status
                End With
            End If
        End If
        
        On Error GoTo 0
        'StatusCodeが200なら成功。
        If statusCode = 200 Then
            successRecordCount = successRecordCount + 1
            '各行のデータを処理結果の配列に記載する。同一レコードの頭の行のみに結果を記載する。
            Array_RowNumAndRowId = recordIdsDictionary.Item(recordId)
            For k = 0 To UBound(Array_RowNumAndRowId, 2)
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
                    Array_OutputData(outputRow, j - 1) = Array_EntryData(Array_RowNumAndRowId(0, k), j)
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
            Array_RowNumAndRowId = recordIdsDictionary.Item(recordId)
            
            flgFirstRow = True
            For k = 0 To UBound(Array_RowNumAndRowId, 2)
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
                    Array_OutputData(outputRow, j - 1) = Array_EntryData(Array_RowNumAndRowId(0, k), j)
                    Array_ErrorData(erroutputRow, j - 1) = Array_EntryData(Array_RowNumAndRowId(0, k), j)
                Next
                outputRow = outputRow + 1
                erroutputRow = erroutputRow + 1
            Next
        End If
    Next
    
    SetUmsVariable $*成功レコード件数$, successRecordCount
    SetUmsVariable $*失敗レコード件数$, errorRecordCount
    
    '処理結果ファイルとエラーファイルのパスを作成する。
    dateTimeNow = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & _
        Right("0" & Hour(Now()), 2) & Right("0" & Minute(Now()), 2) & Right("0" & Second(Now()), 2)
    outputCsvPath = objFSO.BuildPath(outputFolderPath, csvFileName & "_処理結果_" & dateTimeNow & ".csv")
    errorCsvPath = objFSO.BuildPath(outputFolderPath, csvFileName & "_エラーデータ_" & dateTimeNow & ".csv")
    
    '処理結果ファイルを作成する
    If errorRecordCount > 0 Then
        Call KNTN_SaveCsv(Array_ErrorData, errorCsvPath, charcode, 0)
    End If
    Call KNTN_SaveCsv(Array_OutputData, outputCsvPath, charcode, 0)
End Sub