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
Call Kntn_GetAppsInfo()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_GetAppsInfo()

    Dim kntn_api_uri
    Dim responseText
    Dim kntn_app_name
    
    kntn_app_name = !アプリ名(部分一致)!
    kntn_guestspace_id = !ゲストスペースID!
    kntn_app_id = ""
    
    'Excelに出力する
    fname = !*出力先Excel!
    canOverWriteFile = !*出力先にExcelファイルが既に存在するとき|上書き,エラー!
    If fname = "" Then
        Err.Raise 1, "", "*出力先Excelの入力は必須です。"
    End If
    
    Dim objFSO
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(fname) = True And canOverWriteFile = "エラー" Then
        Err.Raise 1, "", "出力先Excelファイルが既に存在しています。既存ファイルを移動する、または出力先Excelパスを変更してください。"
    End If
    
    'QueryのEncode
    Dim encoded
    
    'Kintone API のエンドポイント
    If kntn_guestspace_id = "" Then
        kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/apps.json"
    Else
        kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/apps.json"
    End If

    If kntn_app_name = "" Then
        kntn_api_uri = kntn_api_uri & "?offset="
    Else
        Dim sc
        Set sc = CreateObject("ScriptControl")
        sc.Language = "JScript"
        encoded = sc.CodeObject.encodeURIComponent(kntn_app_name)
        'Kintone API のエンドポイント
        kntn_api_uri = kntn_api_uri & "?name=" & encoded & "&offset="
    End If
    
    Dim json
    Dim json_apps
    
    '二次元配列で行は増やせないため、いったん列を増やす形にする
    Dim array_2jigen()
    ReDim array_2jigen(3, 0)
    array_2jigen(0, 0) = "アプリ名"
    array_2jigen(1, 0) = "アプリID"
    array_2jigen(2, 0) = "アプリコード"
    array_2jigen(3, 0) = "スペースID"
    
    For i = 0 To 100
        offset = i * 100
        '100件ずつ取得するため、offsetを利用する
        api_uri = kntn_api_uri & offset
        
        'アクセストークンの有効性を確認
        Call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token)
        
        ' API リクエストのヘッダーとデータを設定
        With WScript.CreateObject("MSXML2.XMLHTTP")
            .Open "Get", api_uri, False
            .setRequestHeader "Authorization", "Bearer " & kntn_access_token
            .setRequestHeader "Accept", "application/json"
            .setRequestHeader "User-Agent", kntn_userAgent
            .send
            
            ' レスポンステキストを取得
            responseText = .responseText
            statusCode = .status
            
            ' レスポンスの処理を行う
            Select Case statusCode
                Case 200
                    Set json = kntn_ScriptEngine.CodeObject.Parse(responseText)
                    Set json_apps = json.apps
                    jsonlength = json_apps.Length
                    
                    '0件なら終了
                    If jsonlength = 0 Then
                        Exit For
                    End If

                    ReDim Preserve array_2jigen(3, UBound(array_2jigen, 2) + jsonlength)

                    ' JScript9用: 配列要素へのアクセス方法を変更
                    For j = 0 To jsonlength - 1
                        ' JavaScript側で配列要素を取得する関数を使用
                        array_2jigen(0, i * 100 + j + 1) = kntn_ScriptEngine.Eval("(" & responseText & ").apps[" & j & "].name")
                        array_2jigen(1, i * 100 + j + 1) = kntn_ScriptEngine.Eval("(" & responseText & ").apps[" & j & "].appId")
                        array_2jigen(2, i * 100 + j + 1) = kntn_ScriptEngine.Eval("(" & responseText & ").apps[" & j & "].code")
                        array_2jigen(3, i * 100 + j + 1) = kntn_ScriptEngine.Eval("(" & responseText & ").apps[" & j & "].spaceId")
                    Next
                    
                    '100件未満なら終了
                    If jsonlength < 100 Then
                        Exit For
                    End If
                    
                Case Else
                    Err.Raise 1, "", _
                        "Kintoneのアプリ一覧取得操作に失敗しました。" & vbCrLf & _
                        "ステータスコード:" & statusCode & vbCrLf & _
                        "レスポンス: " & Kntn_GetErrorMessage(responseText)
            End Select
        End With
    Next
    
    '行列を入れ替える
    Dim array_transpose()
    ReDim array_transpose(UBound(array_2jigen, 2), 3)
    For i = 0 To UBound(array_2jigen, 2)
        array_transpose(i, 0) = array_2jigen(0, i)
        array_transpose(i, 1) = array_2jigen(1, i)
        array_transpose(i, 2) = array_2jigen(2, i)
        array_transpose(i, 3) = array_2jigen(3, i)
    Next
    
    Call Kntn_SetArrayToExcel(fname, sheetName, array_transpose, "A1", True)
End Sub
