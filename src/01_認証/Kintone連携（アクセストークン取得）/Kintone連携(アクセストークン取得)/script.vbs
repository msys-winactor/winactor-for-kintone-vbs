' -----------------------------------------------------------------------
' Main
' -----------------------------------------------------------------------
' WinActorの変数値を取得
Dim kntn_crypt_keyword
Dim kntn_access_code

kntn_crypt_keyword = !暗号化キーワード!
kntn_access_code = !アクセスコード!
kntn_redirect_uri = !リダイレクトURI!

'  トークン取得URL
Dim kntn_token_uri
kntn_token_uri ="https://" & kntn_subdomain &  ".cybozu.com/oauth2/token"

  'QueryのEncode
  Dim encoded
  encoded =  Kntn_EncodeBase64(kntn_client_id & ":" & kntn_secret_key)

' POSTデータを作成する
sendData = "grant_type=authorization_code&redirect_uri=" & kntn_redirect_endpoint  & "&code=" &  kntn_access_code

' MSXML2.XMLHTTPオブジェクトを作成し、POSTリクエストを送信する
With wscript.CreateObject("MSXML2.XMLHTTP")
  .Open "POST", kntn_token_uri, False
  If kntn_proxy_url <> "" Then .setProxy 2, kntn_proxy_url
  .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  .setRequestHeader "Authorization", "Basic " & encoded
  .setRequestHeader "User-Agent", kntn_userAgent
  .send sendData

  ' レスポンステキストを取得
    responseText = .responseText
    statusCode = .status

    ' レスポンスの処理を行う
    Select Case statusCode
      Case 200

        kntn_access_token = Kntn_GetJSONValueByKey(responseText, "json.access_token")
        kntn_expires_in = Kntn_GetJSONValueByKey(responseText, "json.expires_in")
        kntn_refresh_token = Kntn_GetJSONValueByKey(responseText, "json.refresh_token")

        crypt_access_token = Kntn_EncodePlainText(kntn_access_token, kntn_crypt_keyword)
        crypt_refresh_token = Kntn_EncodePlainText(kntn_refresh_token, kntn_crypt_keyword)

        ' 現在日時にexpores_inを加算した日時を取得
        kntn_expires_date = DateAdd("s", kntn_expires_in, Now)

        Call Kntn_CreateTokenFile(kntn_token_file, crypt_access_token, kntn_expires_date, crypt_refresh_token)

    Case Else
      ' それ以外の場合、エラーを発生させる
      Err.Raise statusCode, "", "Kintone操作に失敗しました。(" & statusCode & ")" & vbCrLf & "レスポンス: " & responseText

  End Select
End With

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_CreateTokenFile(tokenFile,cryptAccessToken,expiresDate,cryptRefreshToken)
  ' 変数の宣言
  Dim objFile
  Dim objStream

  ' FSOの作成
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' 既にファイルが存在する場合は削除
  If fso.FileExists(tokenFile) Then
    fso.DeleteFile tokenFile
  End If

  ' ファイルを新規作成
  fso.CreateTextFile tokenFile

  ' トークンファイルにトークン情報を書き込み(UTF-8)
  With CreateObject("ADODB.Stream")
    .Type = 2
    .Charset = "utf-8"
    .Open
    .WriteText CryptAccessToken & "," & expiresDate & "," & CryptRefreshToken
    .SaveToFile tokenFile, 2
    .Close 'Close stream
  End With

End Sub
