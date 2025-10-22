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

Call KNTN_DownloadFilesByName()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub KNTN_DownloadFilesByName()
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
	dlFolderPath = !*ダウンロードフォルダパス!
  dlFileNames = !ダウンロードファイル名!
	kugirimoji = !*複数ファイル指定時の区切り文字!
	
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

  If dlFolderPath = "" Then
    Err.Raise 1, "", "ダウンロードフォルダパスの入力は必須です。"
  End If  

  '区切り文字がvbCrlfなら改行区切りとする
  if kugirimoji = "vbCrlf" then
    kugirimoji=vbCrLf
  end if

	If kugirimoji = "" Then
    Err.Raise 1, "", "複数ファイル選択時の区切り文字の入力は必須です。"
  End If  

	dim objFso
	Set objFso = CreateObject("Scripting.FileSystemObject")
	'ダウンロードフォルダがない場合は作成
	If Not objfso.FolderExists(dlFolderPath) Then
		Call KNTN_CreateIntermediateFolders(dlFolderPath)
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
	end with

	'StatusCodeが200なら成功。
	if statusCode <> 200 then
		err.raise 1,"","Kintoneの既存レコード情報取得処理に失敗しました。" & vbCrLf & _
		"ステータスコード：" & statusCode  & vbCrLf & _
		"レスポンス: " & KNTN_GetErrorMessage(responseText)
	end if

	Set json = kntn_ScriptEngine.CodeObject.Parse(responseText)

	'既存のレコードのファイルキーを取得する。
	set json_record = json.record

	if kntn_ScriptEngine.Run("checkKey", json_record,fieldcode) = false then
			err.raise 1,"","フィールド「" & field & "」が正しくありません。また、テーブル内のフィールドは選択できません。"
	end if

	'既存のファイルキーとファイル名の情報を取得する。
	tmpKeyData = KNTN_getFieldValue("FILE_KEY",fieldcode,json_record,kugirimoji)
	tmpNameData = KNTN_getFieldValue("FILE_NAME",fieldcode,json_record,kugirimoji)

	'既存のファイルキーとファイル名の一覧を取得する
	array_ExistsFileKeys = Kntn_SplitFiles(tmpKeyData ,kugirimoji)
	array_ExistsFileNames = Kntn_SplitFiles(tmpNameData ,kugirimoji)

	if ubound(array_ExistsFileKeys) <> ubound(array_ExistsFileNames) then
		Err.Raise 1, "", _
			"「複数ファイル指定時の区切り文字」をファイル名やファイルキーに利用されていない文字を指定してください" & vbCrLf & _
			"例：/や|など"
	end if

	'Kintone API のエンドポイント
	if kntn_guestspace_id = "" then
		kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/file.json"
	else
		kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/file.json"
	end if

	'ダウンロード対象のファイル名一覧を配列化する
	array_dlFileNames =  Kntn_SplitFiles(dlFileNames,kugirimoji)

	'同ファイル名が複数存在するときに連番をつけるため、ダウンロード回数を格納する辞書型配列を作成する
  dim dicDLCount
	set dicDLCount = WScript.CreateObject("Scripting.Dictionary")

	'既存のファイルキーを繰り返す
	for idx = 0 to ubound(array_ExistsFileKeys)
		dlFileKey = array_ExistsFileKeys(idx)
		dlFileName = array_ExistsFileNames(idx)

		blnDLTarget = False
		'ファイル名が空なら、全ファイルダウンロード
		if ubound(array_dlFileNames) = -1 then 
			blnDLTarget = true
		else
			'ダウンロード対象のファイルか確認する
			for i = 0 to ubound(array_dlFileNames)
				if array_dlFileNames(i) = dlFileName then 
					blnDLTarget = true
					exit for
				end if
			next
		end if

		'ダウンロード対象ならダウンロードする
		if blnDLTarget = True then 
			'同ファイル名のDL回数のカウントを数える
			if dicDLCount.exists(dlFileName) then
				DLCount = dicDLCount.item(dlfileName)
				dicDLCount.item(dlfileName) = dLCount +1
			else
				DLCount=0
				dicDLCount.item(dlfileName) = dLCount +1
			end if

			'すでにダウンロードしている場合は連番をつける
			if dLCount > 0 then
				extention =  objFso.GetExtensionName(dlFileName)
				'拡張子を除いたファイル名を取得する。
				if len(extention) > 0 then
					extention = "." & extention
					dlFileName = left(dlFileName,len(dlFileName)-len(extention))
				end if
				dlFileName = dlFileName & "_" & dLCount &  extention
			end if
			savePath = objFso.BuildPath(dlFolderPath, dlFileName)

			'Kintone API のエンドポイント
			kntn_api_DLuri = kntn_api_uri & "?fileKey=" & dlFileKey 

			'アクセストークンの有効性を確認
			call KNTN_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

			With wscript.CreateObject("MSXML2.XMLHTTP")
				.Open "Get", kntn_api_DLuri, False
				.setRequestHeader "Authorization", "Bearer " & kntn_access_token
				.setRequestHeader "User-Agent", kntn_userAgent
				.send 

				statusCode = .status

				' レスポンスの処理を行う
				Select Case statusCode
				Case 200
					'バイナリデータを取得
					FileData = .responseBody

					'バイナリデータを生成する
					Dim adoStr
					Set adoStr = CreateObject("ADODB.Stream")

					With adoStr
					.Type = 1
					.Open()
					.Write FileData
					.SaveToFile savePath, 2
					.Close
					End With

				Case Else
					' レスポンステキストを取得
					responseText = .responseText
					Err.Raise 1, "", _
						"Kintoneのファイルダウンロード操作に失敗しました。" & vbCrLf & _
						"ステータスコード：" & statusCode  & vbCrLf & _
						"レスポンス: " & KNTN_GetErrorMessage(responseText)
				End Select
			End With
		end if
	next			
End Sub

