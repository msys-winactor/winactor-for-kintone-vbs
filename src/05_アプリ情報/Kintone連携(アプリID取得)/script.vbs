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
Call Kntn_GetAppID()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_GetAppID()
  Dim kntn_api_uri
  Dim responseText
  Dim kntn_app_name
  
  kntn_app_name = !*アプリ名（完全一致）!
  kntn_guestspace_id = !ゲストスペースID!
  kntn_app_id = ""

  if kntn_app_name = "" then
			Err.Raise 1, "","*アプリ名（完全一致）の入力は必須です。"
	end if

  'QueryのEncode
  Dim encoded
  Dim sc
  Set sc = CreateObject("ScriptControl")
  sc.Language = "JScript"
  encoded = sc.CodeObject.encodeURIComponent(kntn_app_name)
  
  'Kintone API のエンドポイント
	if kntn_guestspace_id = "" then
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/apps.json"
	else
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/apps.json"
	end if

  'Kintone API のエンドポイント
  kntn_api_uri = kntn_api_uri & "?name=" & encoded

  Dim json
  Dim json_apps

  for i = 0 to 100
    offset = i*100
    '100件ずつ取得するため、offsetを利用する
    api_uri =kntn_api_uri & "&offset=" & offset

  
    'アクセストークンの有効性を確認
    call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

    ' API リクエストのヘッダーとデータを設定
    With wscript.CreateObject("MSXML2.XMLHTTP")
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
        Set json =  kntn_ScriptEngine.CodeObject.Parse(responseText)
        Set json_apps = json.apps
        jsonlength = json_apps.Length
        '配列が0のときの処理
        if jsonlength = 0 then 
          if kntn_app_id = "" then
            Err.Raise "1", "", "アプリ名に該当するアプリが見つかりません"
          else 
            '見つかったので抜ける（アプリがちょうど100個目、200個目といったケース）
            exit for
          end if 
        end if 

        '配列内のアプリネームを取得し、完全一致となるものを取得。みつかっても重複があるかもしれないので、全件回す
        dim counter
        for counter= 0 to jsonlength-1
          set appInfo = Eval("json_apps.[" & counter &"]")
          if appInfo.name = kntn_app_name then
            if kntn_app_id <> "" then
              Err.Raise "1", "", "アプリ名に該当するアプリが複数存在します"
            else
              kntn_app_id = appInfo.appId
            end if  
          end if
        next

        '該当アプリが100個未満でアプリが見つからない場合はエラー
        if jsonlength < 100 and kntn_app_id = "" then
          Err.Raise "1", "", "アプリ名に該当するアプリが見つかりません"
        '該当アプリが100個未満でアプリが見つかった場合は処理を抜ける
        elseif  jsonlength < 100 and kntn_app_id <> "" then
          exit for
        end if

      Case Else
        Err.Raise 1, "", _
				  "KintoneのアプリID取得操作に失敗しました。" & vbCrLf & _
					"ステータスコード：" & statusCode  & vbCrLf & _
					"レスポンス: " & Kntn_GetErrorMessage(responseText)
      End Select
    End With
  next
  call SetUmsVariable($*アプリID$,kntn_app_id)


End Sub
