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

Call KNTN_AddFiles()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub KNTN_AddFiles()
  Dim kntn_api_uri
  Dim responseText
  Dim sendData
  Dim kntn_app_id
  dim excelFilePath

  kntn_app_id = !*アプリID!
  kntn_guestspace_id = !ゲストスペースID!
	recordId = !*レコードID!
  nameOrCode = !*フィールド名またはフィールドコード選択|フィールド名,フィールドコード!
  field = !*フィールド!
  kugirimoji = !*複数ファイル選択時の区切り文字!
  addFileKeys = !*アップロードファイルパス!

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
    Err.Raise 1, "", "複数ファイル選択時の区切り文字の入力は必須です。"
  End If  
  '区切り文字がvbCrlfなら改行区切りとする
  if kugirimoji = "vbCrlf" then
    kugirimoji=vbCrLf
  end if
  If addFileKeys = "" Then
    Err.Raise 1, "", "アップロードファイルパスの入力は必須です。"
  End If  

	'Kintone API のエンドポイント
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
		if Err.Number <> 0 then
			on error goto 0
			err.raise 1,"","フィールド「" & field & "」が正しくありません。また、テーブル内のフィールドは選択できません。"
		end if
		on error goto 0




	'アクセストークンの有効性を確認
	call KNTN_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

	'既存のIDのレコード情報を取得する。
	With wscript.CreateObject("MSXML2.XMLHTTP")
		.Open "Get", kntn_api_uri & "?app=" & kntn_app_id & "&id=" & recordId, False
		.setRequestHeader "Authorization", "Bearer " & kntn_access_token
		.setRequestHeader "User-Agent", kntn_userAgent
		.send 

		' レスポンステキストを取得
		responseText = .responseText
		statusCode = .status
		Set json = kntn_ScriptEngine.CodeObject.Parse(responseText)
	end with

	if statusCode = 200 then
		'既存のレコードのファイルキーを取得する。
		set json_record = json.record
		fileKeys  = ""

		'既存のファイルキー情報を取得する。
		tmpData = KNTN_getFieldValue("FILE_KEY",fieldcode,json_record,kugirimoji)

		if tmpData <> "" then
			array_ExistsFiles = Kntn_SplitFiles(tmpData,kugirimoji)

			'既存のファイルキーを設定する
			for idx = 0 to ubound(array_ExistsFiles)
				if fileKeys = "" then 
					fileKeys = "{""fileKey"":""" & array_ExistsFiles(idx) & """}"
				else
					fileKeys = fileKeys & ",{""fileKey"":""" & array_ExistsFiles(idx) & """}"
				end if
			next			
		end if

		array_Files = Kntn_SplitFiles(addFileKeys,kugirimoji)
			
		for idx = 0 to ubound(array_Files)
			'ファイルを一時領域にアップロードし、ファイルキーを取得する。
			array_Files(idx) = KNTN_tmpUpload(array_Files(idx))
		next

		'追加のファイルキーを設定する
		for idx = 0 to ubound(array_Files)
			if fileKeys = "" then 
				fileKeys = "{""fileKey"":""" & array_Files(idx) & """}"
			else
				fileKeys = fileKeys & ",{""fileKey"":""" & array_Files(idx) & """}"
			end if
		next

			sendData = "{""app"":" & kntn_app_id & ",""id"":"& recordId &",""record"": {""" _
									& fieldCode & """:{""value"":["  &	fileKeys & "]}}}"


		'アクセストークンの有効性を確認
		call KNTN_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

			'更新APIを呼び出す。
		With wscript.CreateObject("MSXML2.XMLHTTP")
			.Open "Put", kntn_api_uri , False
			.setRequestHeader "Authorization", "Bearer " & kntn_access_token
			.setRequestHeader "Content-Type", "application/json"
			.send sendData

			' レスポンステキストを取得
			responseText = .responseText
			statusCode = .status
		end with
	end if

	'StatusCodeが200なら成功。
	if statusCode = 200 then

	else
		err.raise 1,"","Kintoneのファイル追加操作に失敗しました。" & vbCrLf & _
		"ステータスコード：" & statusCode  & vbCrLf & _
		"レスポンス: " & KNTN_GetErrorMessage(responseText)
	end if

End Sub


