' -----------------------------------------------------------------------
' Main
' -----------------------------------------------------------------------
'Kintone連携（アクセストークン）が配置されているかの確認
if isempty(kntn_client_id) then
  err.raise 1,"","WinActor for kintone ver1.1.1 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
end if

if isempty(kntn_userAgent) then
  err.raise 1,"","WinActor for kintone ver1.1.1 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
end if

' トークンのチェック
Call KNTN_CheckAccessToken(kntn_client_id)

' Kintoneの指定アプリの指定レコード・指定フィールドを更新
Call KNTN_UpdateField()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub KNTN_UpdateField()
  kntn_app_id = !*アプリID!
  kntn_guestspace_id = !ゲストスペースID!
	recordId = !*レコードID!
  nameOrCode = !*フィールド名またはフィールドコード選択|フィールド名,フィールドコード!
  field = !*フィールド!
  kugirimoji = !*複数設定値の区切り文字!
	brank = !*ブランクとして認識する値!
  fieldvalue = !*設定値!

  dim recordCount 
  If kntn_app_id = "" Then
    Err.Raise 1, "", "アプリIDの入力は必須です。"
  End If  

  If recordid = "" Then
    Err.Raise 1, "", "レコードIDの入力は必須です。"
  End If  

  If field = "" Then
    Err.Raise 1, "", "フィールドの入力は必須です。"
  End If  

  If kugirimoji = "" Then
    Err.Raise 1, "", "複数設定値の区切り文字の入力は必須です。"
	elseif kugirimoji = "vbCrlf" then
		kugirimoji = vbCrLf
  End If  


	if brank = "" then
		Err.Raise 1, "","「ブランクとして認識する値」の入力は必須です。"
	end if

  If fieldvalue = "" Then
    Err.Raise 1, "", "設定値の入力は必須です。"
  End If  

  'Kintone レコード更新API のエンドポイント
  if kntn_guestspace_id = "" then
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/record.json"
  else
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/record.json"
  end if


	'フィールドのコードを取得する
		'フィールド情報一覧の情報をAPIで取得する。
		dim json_fieldsInfo,json_properties
		Set json_fieldsInfo= kntn_ScriptEngine.CodeObject.Parse(KNTN_GetFieldsInfo(kntn_app_id,kntn_guestspace_id))
		set json_properties = json_fieldsInfo.properties
		on error resume next
		set json_fieldInfo = KNTN_getFieldJson(field,nameOrCode,json_properties)
		fieldCode = json_fieldInfo.Code
		fieldType = json_fieldInfo.Type
		if Err.Number <> 0 then
			on error goto 0
			err.raise 1,"","フィールド「" & field & "」が正しくありません。また、テーブル内のフィールドは選択できません。"
		end if
		on error goto 0

		if fieldType = "FILE" then
			err.raise 1,"","添付ファイルの操作はできません。ファイル操作ライブラリをご利用ください。"
		else
			fieldKeyValue = KNTN_CreateJsonKeyAndValue(fieldCode,fieldType,fieldvalue,brank,kugirimoji)
		end if

		sendData = "{app:" &  kntn_app_id & ",""id"":" & recordId   & ",record:{" & fieldKeyValue  & "}}" 

		'アクセストークンの有効性を確認
		call KNTN_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

		' API リクエストのヘッダーとデータを設定
		With wscript.CreateObject("MSXML2.XMLHTTP")
			.Open "PUT", kntn_api_uri, False
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

			Case Else
				'処理結果配列に格納する
				err.raise 1,"","Kintoneのレコード更新操作に失敗しました。" & vbCrLf & _
								"ステータスコード：" & statusCode  & vbCrLf & _
								"レスポンス: " & KNTN_GetErrorMessage(responseText) 
			End Select
		End With

End Sub


