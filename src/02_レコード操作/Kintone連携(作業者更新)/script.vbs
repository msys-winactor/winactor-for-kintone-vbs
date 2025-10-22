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
Call KNTN_CheckAccessToken(kntn_client_id)

' Kintoneの作業者を変更する
Call KNTN_UpdateAssignees()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub KNTN_UpdateAssignees()
  kntn_app_id = !*アプリID!
  kntn_guestspace_id = !ゲストスペースID!
  kntn_record_id = !*レコードID!
  kntn_assignee = !*作業者（ログイン名）!
  kntn_kugirimoji = !*複数作業者設定時の区切り文字!

  If kntn_app_id = "" Then
    Err.Raise 1, "", "アプリIDの入力は必須です。"
  End If  

  If kntn_record_id = "" Then
    Err.Raise 1, "", "レコードIDの入力は必須です。"
  End If  

  If kntn_assignee = "" Then
    Err.Raise 1, "", "作業者（ログイン名）の入力は必須です。"
  End If  

  if kntn_kugirimoji = "" then
    Err.Raise 1, "", "複数作業者設定時の区切り文字の入力は必須です。"
  end if

  '区切り文字がvbCrlfなら改行区切りとする
  if kntn_kugirimoji = "vbCrlf" then
    kntn_kugirimoji=vbCrLf
  end if

  'Kintone API のエンドポイント
	if kntn_guestspace_id = "" then
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/record/assignees.json"
	else
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/record/assignees.json"
	end if
  
	sendData = "{""app"":" & kntn_app_id & ",""id"":" & kntn_record_id & _ 
	             ",""assignees"":[" &  KNTN_SplitKugirimojiForJson(kntn_assignee,kntn_kugirimoji) &  "]}"

  'アクセストークンの有効性を確認
  call KNTN_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

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
				  "Kintoneの作業者更新操作に失敗しました。" & vbCrLf & _
					"ステータスコード：" & statusCode  & vbCrLf & _
					"レスポンス: " & KNTN_GetErrorMessage(responseText)
    End Select
  End With
End Sub

