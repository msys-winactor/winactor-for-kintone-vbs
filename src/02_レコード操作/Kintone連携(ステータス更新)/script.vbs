' -----------------------------------------------------------------------
' Main
' -----------------------------------------------------------------------
'Kintone連携（アクセストークン）が配置されているかの確認
if isempty(kntn_client_id) then
  err.raise 1,"","WinActor for kintone ver1.1.0 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
end if

if isempty(kntn_userAgent) then
  err.raise 1,"","WinActor for kintone ver1.1.0 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
end if

' トークンのチェック
Call Kntn_CheckAccessToken(kntn_client_id)

' Kintoneの指定アプリのフィールド情報を取得
Call Kntn_UpdateStatus()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_UpdateStatus()
  kntn_app_id = !*アプリID!
  kntn_guestspace_id = !ゲストスペースID!
  kntn_record_id = !*レコードID!
  kntn_action= !*アクション名!
  kntn_assignee = !作業者（ログイン名）!

  If kntn_app_id = "" Then
    Err.Raise 1, "", "アプリIDの入力は必須です。"
  End If  

  If kntn_record_id = "" Then
    Err.Raise 1, "", "レコードIDの入力は必須です。"
  End If  

  If kntn_action = "" Then
    Err.Raise 1, "", "アクション名の入力は必須です。"
  End If  


  'Kintone API のエンドポイント
	if kntn_guestspace_id = "" then
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/record/status.json"
	else
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/record/status.json"
	end if

  if kntn_assignee ="" then
  	sendData = "{""app"":" & kntn_app_id & ",""id"":" & kntn_record_id & _ 
	             ",""action"": """ & kntn_action & """}" 
  else
  	sendData = "{""app"":" & kntn_app_id & ",""id"":" & kntn_record_id & _ 
	             ",""action"": """ & kntn_action & """,""assignee"":""" &  kntn_assignee &  """}"
  end if

		'アクセストークンの有効性を確認
		call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

  ' API リクエストのヘッダーとデータを設定
  With wscript.CreateObject("MSXML2.XMLHTTP")
    .Open "Put", kntn_api_uri, False
    .setRequestHeader "Authorization", "Bearer " & kntn_access_token
    .setRequestHeader "Content-Type", "application/json"
    .setRequestHeader "User-Agent", kntn_userAgent
    .send sendData

    ' レスポンステキストを取得
    responseText = .responseText
    statusCode = .status

    ' レスポンスの処理を行う
    Select Case statusCode
      Case 200
		     '成功。戻り値は特に渡さない
      Case Else
        Err.Raise 1, "", _
				  "Kintoneのステータス更新操作に失敗しました。" & vbCrLf & _
					"ステータスコード：" & statusCode  & vbCrLf & _
					"レスポンス: " & Kntn_GetErrorMessage(responseText)
    End Select
  End With
End Sub
