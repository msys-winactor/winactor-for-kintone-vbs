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
Call Kntn_GetAppsInfoToCsv()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_GetAppsInfoToCsv()
  Dim kntn_api_uri
  Dim responseText
  Dim kntn_app_name
  
  kntn_app_name = !アプリ名（部分一致）!
  kntn_guestspace_id = !ゲストスペースID!
  kntn_app_id = ""

	'CSVに出力する
  csvFilePath = !*出力先CSVパス!
	canOverWriteFile = !*出力先にCSVファイルが既に存在するとき|上書き,エラー!
  charcode =!*文字コード|shift-jis,utf-8!

	if csvFilePath = "" then
		Err.Raise 1, "","*出力先CSVの入力は必須です。"
	end if

  Dim objFSO
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
  If objFSO.FileExists(csvFilePath) = True and canOverWriteFile = "エラー" Then
    Err.Raise 1, "", "出力先CSVファイルが既に存在しています。既存ファイルを移動する、または出力先CSVパスを変更してください。"
  end if

  'QueryのEncode
  Dim encoded
  
  'Kintone API のエンドポイント
	if kntn_guestspace_id = "" then
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/apps.json"
	else
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/apps.json"
	end if
	if kntn_app_name = "" then 
		kntn_api_uri = kntn_api_uri & "?offset="
	else
		Dim sc
		Set sc = CreateObject("ScriptControl")
		sc.Language = "JScript"
		encoded = sc.CodeObject.encodeURIComponent(kntn_app_name)
		'Kintone API のエンドポイント
		kntn_api_uri = kntn_api_uri & "?name=" & encoded & "&offset="
	end if

  Dim json
  Dim json_apps

	'二次元配列で行は増やせないため、いったん列を増やす形にする
  dim array_2jigen()
  redim array_2jigen(3,0)
	array_2jigen(0,0) = "アプリ名"
	array_2jigen(1,0) = "アプリID"
	array_2jigen(2,0) = "アプリコード"    
	array_2jigen(3,0) = "スペースID"    



  for i = 0 to 100
    offset = i*100
    '100件ずつ取得するため、offsetを利用する
    api_uri =kntn_api_uri  & offset
		
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
				'0件なら終了
				if jsonlength = 0 then
					exit for 
				end if


			  redim preserve array_2jigen(3,ubound(array_2jigen,2) + jsonlength)
				c = 1
				for each app in json_apps
						array_2jigen(0,i*100+c) = app.name
						array_2jigen(1,i*100+c) =  app.appId
						array_2jigen(2,i*100+c) =  app.code
						array_2jigen(3,i*100+c) =  app.spaceId
						c= c+1     
				next

				'100件未満なら終了
				if jsonlength < 100 then
					exit for
				end if

      Case Else
        Err.Raise 1, "", _
				  "Kintoneのアプリ一覧取得操作に失敗しました。" & vbCrLf & _
					"ステータスコード：" & statusCode  & vbCrLf & _
					"レスポンス: " & Kntn_GetErrorMessage(responseText)
      End Select
    End With
  next

	'行列を入れ替える
	dim array_transpose()
	redim array_transpose(ubound(array_2jigen,2),3)
	for i = 0 to ubound(array_2jigen,2)
		array_transpose(i,0) = array_2jigen(0,i)
		array_transpose(i,1) = array_2jigen(1,i)
		array_transpose(i,2) = array_2jigen(2,i)
		array_transpose(i,3) = array_2jigen(3,i)
	next 

	'ヘッダーのみ入力
	call KNTN_SaveCsv(array_transpose,csvFilePath,charcode,0)   

End Sub
