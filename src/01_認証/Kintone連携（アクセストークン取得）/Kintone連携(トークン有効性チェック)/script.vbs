'========================================================
'  ライセンス認証
'========================================================
Dim kntn_license

Const kntn_MSYSADAPTER_PRODUCT_NAME = "Kintone"
Const W4kntn_PRODUCT_NAME = "WinActor for Kintone"
Const W4kntn_LICENSE_FILE_NAME = "w4kntn.lic"

' ライセンス認証
kntn_license = kntn_CheckLicense()


' -----------------------------------------------------------------------
' トークンの確認
' -----------------------------------------------------------------------
' 変数の宣言
Dim kntn_crypt_keyword
Dim Kntn_checkResult
Dim kntn_result
Dim kntn_token_info
Dim kntn_tokens_folder
Dim kntn_file_name
Dim kntn_expires_date

kntn_result = False
Kntn_checkResult = $KNTN_TOKEN_CHECK$

' トークンファイルから情報取得
On Error Resume Next
kntn_result = KNTN_CheckAccessToken(kntn_client_id)
If Err.Number <> 0 Then
  kntn_result = False
End If
On Error GoTo 0

If kntn_result Then
  kntn_result = Kntn_CheckConnectAPI(kntn_access_Token)
End If

' トークンのチェック結果を返却
SetUmsVariable Kntn_checkResult, kntn_result


'========================================================
' Functions
'========================================================

' ライセンスをチェックする関数
Function kntn_CheckLicense()
  Dim kntn_folder_path
  Dim kntn_lic_file_path
  Dim kntn_mydoc_path
  Dim objLicFile
  Dim cryptKeyword
  Dim licenseInfo
  Dim arrLicInfo
  Dim last_date
  Dim lic_warn_message
  Dim lic_error_message

  ' ファイルシステムオブジェクトを作成
  Set objFs = CreateObject("Scripting.FileSystemObject")

  ' ドキュメントフォルダのパスを取得
  kntn_mydoc_path = kntn_GetDocumentFolderPath()
  kntn_folder_path = objFs.BuildPath(kntn_mydoc_path, "WinActor\MsysAdapters\" & kntn_MSYSADAPTER_PRODUCT_NAME)
  kntn_lic_file_path = objFs.BuildPath(kntn_folder_path, W4kntn_LICENSE_FILE_NAME)

  ' ライセンスファイルが存在しない場合、エラーメッセージを表示して終了
  If Not(objFs.FileExists(kntn_lic_file_path)) Then
    lic_error_message = W4kntn_PRODUCT_NAME & "のライセンス認証が未実施です。" & vbCrLf & "ライセンス認証を実施した後、再度実行してください。"
    Err.Raise 1, "", lic_error_message
  End If

  ' ライセンスファイルを読み込む
  Set objLicFile = objFs.OpenTextFile(kntn_lic_file_path)
  cryptKeyword = objLicFile.ReadLine()
  objLicFile.Close()

  ' ライセンス情報を復号化
  licenseInfo = kntn_DecryptString(cryptKeyword, "Msys#w4b!")
  arrLicInfo = Split(licenseInfo, ",")

  ' 製品名が一致しない場合、エラーメッセージを表示して終了
  If arrLicInfo(2) <> kntn_MSYSADAPTER_PRODUCT_NAME Then
    lic_error_message = W4kntn_PRODUCT_NAME & "のライセンス認証を実施してください。"
    Err.Raise 1, "", lic_error_message
  End If

  ' トライアル版の場合、ライセンス期限をチェック
  If arrLicInfo(0) = "TRIAL" Then
    If DateDiff("d", Date(), arrLicInfo(1)) < 0 Then
      ' トライアル版で期限切れの場合、エラーメッセージを表示して終了
      lic_error_message = W4kntn_PRODUCT_NAME & "のトライアルライセンス期限が切れております。" & vbCrLf & _
                          "製品版ライセンスを登録してください。"
      Err.Raise 1, "", lic_error_message
    Else
      ' トライアル版で期限内の場合、Trueを返す
      kntn_CheckLicense = True
      Exit Function
    End If
  End If

  ' フローティングライセンスの場合、ホスト名をチェックしない
  If arrLicInfo(0) <> "FLA" Then
    ' ホスト名が一致しない場合、エラーメッセージを表示して終了
    If arrLicInfo(0) <> kntn_GetComputerName() Then
      lic_error_message = W4kntn_PRODUCT_NAME & "のライセンス認証を実施してください。"
      Err.Raise 1, "", lic_error_message
    End If
  End If

  ' ライセンス期限が90日を過ぎている場合、エラーメッセージを表示して終了
  If DateDiff("d", Date(), arrLicInfo(1)) < -90 Then
    lic_error_message = W4kntn_PRODUCT_NAME & "のライセンスが無効です。" & vbCrLf & "有効なライセンスで認証を実施し、再度実行してください。"
    Err.Raise 1, "", lic_error_message
  End If

  ' ライセンス期限が90日以内の場合、警告メッセージを表示してTrueを返す
  If DateDiff("d", Date(), arrLicInfo(1)) < 0 Then
    ' 最終利用日を取得
    last_date = DateAdd("d", 90, arrLicInfo(1))

    ' ライセンス期限切れ警告メッセージを表示
    lic_warn_message = W4kntn_PRODUCT_NAME & "のライセンス期限が切れております。" & vbCrLf & _
                       "ライセンス期限：～" & arrLicInfo(1) & vbCrLf & _
                       "恐れ入りますが丸紅情報システムズ営業担当までご連絡をお願いいたします。" & vbCrLf & vbCrLf & _
                       "更新をご希望ではない場合、本ライブラリにつきましては" & vbCrLf & _
                       last_date & "をもって利用が出来なくなります為ご注意ください。"

    Call kntn_DisplayDialog(lic_warn_message, 10)

    ' Trueを返す
    kntn_CheckLicense = True
    Exit Function
  End If

  kntn_CheckLicense = True
  Set objFs = Nothing
End Function

' ドキュメントフォルダのパスを取得する関数
Function kntn_GetDocumentFolderPath()
  Dim objShell
  Set objShell = CreateObject("WScript.Shell")
  kntn_GetDocumentFolderPath = objShell.SpecialFolders("mydocuments")
  Set objShell = Nothing
End Function

' コンピュータ名を取得する関数
Function kntn_GetComputerName()
  On Error Resume Next
  Err.Clear
  Dim strRet
  Dim objNetWork
  strRet = ""
  Set objNetWork = CreateObject("WScript.Network")
  strRet = objNetWork.ComputerName
  Set objNetWork = Nothing
  kntn_GetComputerName = strRet
  Err.Clear
End Function

' ダイアログを表示する関数
Sub kntn_DisplayDialog(message, timeoutSeconds)
  Dim dialogTitle
  Dim shell
  Dim buttonPressed
  
  dialogTitle = "UMS（タイムアウト時間：" & timeoutSeconds & "秒）"
  Set shell = WScript.CreateObject("WScript.Shell")
  buttonPressed = shell.Popup(message, timeoutSeconds, dialogTitle, vbOKonly)
  Set shell = Nothing
End Sub

' 文字列を復号化する関数
Function kntn_DecryptString(crypt, pwd)
  Dim plain
  Dim num
  Dim salt
  Dim saltlen
  Dim saltIndex
  Dim l_mask
  Dim length
  Dim i
  Dim bit
  Dim l_num
  
  salt = pwd
  saltlen = Len(salt)
  saltIndex = 1
  l_mask = 65535
  plain = ""
  length = Len(crypt)

  ' 4文字ずつ処理して復号化
  For i = 1 To length Step 4
    bit = Asc(Mid(salt, saltIndex, 1))
    saltIndex = saltIndex + 1
    If (saltIndex > saltLen) Then
      saltIndex = 1
    End If

    l_num = (CLng("&h" & Mid(crypt, i, 4)) Xor bit) And l_mask
    plain = plain & Chr(l_num)
  Next

  ' 復号化した文字列を返す
  kntn_DecryptString = plain
End Function


' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Function Kntn_CheckAccessToken(kntnClientId)
  ' 変数の宣言
  Dim kntn_first_line
  Dim kntn_token_array
  Dim crypt_access_token
  Dim crypt_refresh_token
  Dim fso
  Dim tokenInfo

  ' FSOの作成
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' トークンファイルが存在しなければ空文字を返す
  If Not fso.FileExists(kntn_token_file) Then
    KNTN_CheckAccessToken = False
    Exit Function
  End If
  
  ' トークンファイルの読み込み
  With CreateObject("ADODB.Stream")
    .Type = 2
    .Charset = "utf-8"
    .Open
    .LoadFromFile(kntn_token_file)
    kntn_first_line = .ReadText(-2)
    .Close
  End With

  ' kntn_first_lineが空文字の場合、空文字を返す
  If kntn_first_line = "" Then
    Kntn_CheckAccessToken = False
    Exit Function
  End If

  ' kntn_first_lineをカンマで分割
  kntn_token_array = Split(kntn_first_line, ",")
  crypt_access_token = kntn_token_array(0)
  kntn_expires_date = kntn_token_array(1)
  crypt_refresh_token = kntn_token_array(2)

  ' 戻り値にcrypt_access_token、crypt_refresh_tokenを設定
  tokenInfo = Kntn_DecodeCryptText(crypt_access_token, kntn_crypt_keyword) & "," & _
              kntn_expires_date & "," & _
              Kntn_DecodeCryptText(crypt_refresh_token, kntn_crypt_keyword)

  kntn_token_array = Split(tokenInfo, ",")
  kntn_access_token = kntn_token_array(0)
  kntn_expires_date = kntn_token_array(1)
  kntn_refresh_token = kntn_token_array(2)

  ' アクセストークンの有効性を確認
  If Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) Then
    KNTN_CheckAccessToken = True
  Else
    If Kntn_RenewAccessToken(kntn_refresh_token) Then
      KNTN_CheckAccessToken = True
    Else
      KNTN_CheckAccessToken = False
    End If
  End If

  ' メモリの解放
  Set fso = Nothing
End Function

Function Kntn_CheckAccessTokenValidity(accessToken, expiresDate, refreshToken)
  Dim kntn_api_uri
  Dim responseText

  ' 現在日時とkntn_expires_dateを比較し、現在日時がkntn_expires_date-10分よりも後の場合
  If DateDiff("n", Now, expiresDate) < 10 Then
    ' リフレッシュトークンを使用してアクセストークンを再取得
    If Kntn_RenewAccessToken(refreshToken) Then
      Kntn_CheckAccessTokenValidity = True
    Else
      Kntn_CheckAccessTokenValidity = False
    End If
    Exit Function
  End If
  
  Kntn_CheckAccessTokenValidity = True
End Function

Function Kntn_CheckConnectAPI(accessToken)
  Dim kntn_api_uri
  
  ' Kintone API のエンドポイント
  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/apps.json?ids[0]=1"

  ' API リクエストのヘッダーとデータを設定
  With WScript.CreateObject("MSXML2.XMLHTTP")
    .Open "Get", kntn_api_uri, False
    If kntn_proxy_url <> "" Then .setProxy 2, kntn_proxy_url
    .setRequestHeader "Authorization", "Bearer " & accessToken
    .setRequestHeader "Accept", "application/json"
    .setRequestHeader "User-Agent", kntn_userAgent
    .send

    Select Case .status
      Case 200
        ' 200 OK
        Kntn_CheckConnectAPI = True
      Case Else
        ' その他のステータスコード
        Kntn_CheckConnectAPI = False
    End Select
  End With
End Function

Sub Kntn_UpdateTokenFile(tokenFile, cryptAccessToken, expiresDate, cryptRefreshToken)
  ' グローバル変数に持たせる
  kntn_expires_date = expiresDate
  
  ' トークンファイルにトークン情報を書き込み(UTF-8)
  With CreateObject("ADODB.Stream")
    .Type = 2
    .Charset = "utf-8"
    .Open
    .WriteText cryptAccessToken & "," & expiresDate & "," & cryptRefreshToken
    .SaveToFile tokenFile, 2
    .Close
  End With
End Sub

Function Kntn_RenewAccessToken(refreshToken)
  ' 変数の宣言
  Dim kntn_api_uri
  Dim responseText
  Dim sendData
  Dim expires_date
  Dim crypt_access_token
  Dim crypt_refresh_token
  Dim kntn_token_info
  Dim encoded
  Dim sc
  Dim getData

  ' Kintone API のエンドポイント
  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/oauth2/token"

  ' QueryのEncode
  encoded = Kntn_EncodeBase64(kntn_client_id & ":" & kntn_secret_key)

  ' POSTデータを作成する
  sendData = "grant_type=refresh_token&refresh_token=" & refreshToken

  ' API リクエストのヘッダーとデータを設定
  With WScript.CreateObject("MSXML2.XMLHTTP")
    .Open "POST", kntn_api_uri, False
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    .setRequestHeader "Authorization", "Basic " & encoded
    .setRequestHeader "User-Agent", kntn_userAgent
    .send sendData

    ' レスポンステキストを取得
    responseText = .responseText

    Select Case .status
      Case 200
        ' ステータスコードが200の場合、処理を行う
        ' JSONをパースする（ScriptControlを使用）
        Set sc = CreateObject("ScriptControl")
        sc.Language = "JScript"
        sc.AddCode "function parseJSON(jsonString) { return eval('(' + jsonString + ')'); }"
        Set getData = sc.CodeObject.parseJSON(responseText)

        ' Token情報を取得する
        kntn_access_token = getData.access_token
        kntn_expires_in = getData.expires_in
        kntn_refresh_token = refreshToken

        crypt_access_token = Kntn_EncodePlainText(kntn_access_token, kntn_crypt_keyword)
        crypt_refresh_token = Kntn_EncodePlainText(refreshToken, kntn_crypt_keyword)

        ' 現在日時にexpores_inを加算した日時を取得
        expires_date = DateAdd("s", kntn_expires_in, Now)
        Call Kntn_UpdateTokenFile(kntn_token_file, crypt_access_token, expires_date, crypt_refresh_token)

        ' Trueを返却する
        Kntn_RenewAccessToken = True

      Case Else
        ' それ以外の場合、Falseを返却する
        Kntn_RenewAccessToken = False
    End Select
  End With
End Function