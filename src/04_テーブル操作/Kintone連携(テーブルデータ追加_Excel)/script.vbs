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
Call Kntn_AddTableData()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_AddTableData()
  Dim kntn_api_uri
  Dim responseText
  Dim sendData
  Dim kntn_app_id
  dim excelFilePath

  kntn_app_id = !*アプリID!
  kntn_guestspace_id = !ゲストスペースID!
  nameOrCode = !*ヘッダー|フィールド名,フィールドコード!
  fieldtable = !*テーブルのフィールド名またはフィールドコード!
  kugirimoji = !*複数設定値の区切り文字!
	brank =!*ブランクとして認識する値!
  excelFilePath = !*入力ファイルパス!
  excelSheetName = !*シート名!
  outputFilePath = !*処理結果ファイルパス!
  canOverWriteFile = !*処理結果ファイルが既に存在するとき|上書き,エラー!

  Dim objFSO
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
  If objFSO.FileExists(outputFilePath) = True and canOverWriteFile = "エラー" Then
    Err.Raise 1, "", "処理結果ファイルが既に存在しています。既存ファイルを移動する、または処理結果ファイルパスを変更してください。"
  end if

  successRecordCount=0
  errorRecordCount=0

  dim recordCount 
  If kntn_app_id = "" Then
    Err.Raise 1, "", "「アプリID」の入力は必須です。"
  End If  

  If fieldtable = "" Then
    Err.Raise 1, "", "「テーブルのフィールド名またはフィールドコード」の入力は必須です。"
  End If  

	if kugirimoji = "" then
		Err.Raise 1, "","「複数設定値の区切り文字」の入力は必須です。"
	end if
  
	'区切り文字がvbCrlfなら改行区切りとする
  if kugirimoji = "vbCrlf" then
    kugirimoji=vbCrLf
  end if
	
	if brank = "" then
		Err.Raise 1, "","「ブランクとして認識する値」の入力は必須です。"
	end if

	if excelFilePath = "" then
		Err.Raise 1, "","「入力ファイルパス」の入力は必須です。"
	end if

	if excelSheetName = "" then
		Err.Raise 1, "","「シート名」の入力は必須です。"
	end if

  if outputFilePath = "" then
		Err.Raise 1, "","「処理結果ファイルパス」の入力は必須です。"
	end if


			'Kintone API のエンドポイント
		if kntn_guestspace_id = "" then
			kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/record.json"
		else
			kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/record.json"
		end if

  'フィールド情報一覧の情報をAPIで取得する。
  dim json_fieldsInfo,json_properties
  Set json_fieldsInfo= kntn_ScriptEngine.CodeObject.Parse(Kntn_GetFieldsInfo(kntn_app_id,kntn_guestspace_id))
  set json_properties = json_fieldsInfo.properties

	'該当のテーブルフィールドの情報を取得する。
	on error resume next
	set json_TableInfo = Kntn_getFieldJson(fieldtable,nameOrCode,json_properties)
	tableName = json_TableInfo.label
	tableCode = json_TableInfo.Code

	if err.number <> 0 then
		on error goto 0
		err.raise 1,"","テーブルのフィールド名またはフィールドコードが正しくありません。"
	end if
	on error goto 0


  'Lookupのmappingの項目の一覧を取得する
  dim array_fieldsinfo,kntn_rootArray,lookUpArray,array_header
	set kntn_rootArray = WScript.CreateObject("Scripting.Dictionary")
  lookUpArray = Kntn_CreateLookUpArray(json_TableInfo.fields)

	'入力ファイルのデータを配列に格納する
  Dim Array_EntryData
	Array_EntryData = Kntn_GetArraybyExcel(ExcelFilePath,excelSheetName)


  'ヘッダー情報をまとめたRootArrayを求める
  call Kntn_getHeaderArray(json_TableInfo.fields,outputType,nameOrCode,kntn_rootArray,fieldtable,lookUpArray,True)

  if nameOrCode = "フィールド名" then
    rowNum = 1
		recordIdColName="レコードID"
	else 
    rowNum = 0   
		recordIdColName="$id"
  end if

	'レコードIDまたは$id行が存在しない場合はエラー出力	
	if Array_EntryData(1,1) <> recordIdColName then
		err.raise 1,"","入力ファイルに" & recordIdColName  & "フィールドが存在しません。"
	end if	



	array_header = kntn_rootArray.Item(fieldtable)

	'フィールドコードが入力ファイルの何列目かを辞書型配列にまとめる（毎回求めるのは大変なため）
	dim headerDictionary
	Set headerDictionary = WScript.CreateObject("Scripting.Dictionary")
	for i = 2 to ubound(Array_EntryData,2)
		tmpHeader = Array_EntryData(1,i)
		if tmpHeader = "rowId" or tmpHeader = "テーブル行ID" then
			headerDictionary.Add "rowId",i
		else
			for j = 0 to ubound(array_header,2)
				if array_header(rownum,j) = tmpHeader then
					'辞書にフィールドコードと列名を追加する
					headerDictionary.Add array_header(0,j),i
					exit for
				end if	
			next
		end if
	next


	'レコードIDの一覧を辞書型配列に持たせる。
	Dim recordIdsDictionary
	Set recordIdsDictionary = WScript.CreateObject("Scripting.Dictionary")
	dim array_Rows()
	for i = 2 to ubound(Array_EntryData,1)
		recordId = Array_EntryData(i,1)
		if recordIdsDictionary.Exists(recordId) Then
			exist_array_Rows = recordIdsDictionary.Item(recordId)
			idx = ubound(exist_array_Rows)+1
			redim preserve exist_array_Rows(idx)
			exist_array_Rows(idx) = i
			recordIdsDictionary.Item(recordId)=exist_array_Rows
		else
			redim array_Rows(0)
			array_Rows(0) = i
			'レコードIDと行一覧を持たせる
			recordIdsDictionary.Add recordId,array_Rows
		end if
	next

	'処理結果の配列とエラー出力の配列のヘッダー部のみ作成
	dim Array_OutputData(),Array_ErrorData()
	redim Array_OutputData(ubound(Array_EntryData,1)-1,ubound(Array_EntryData,2)-1+3)
	redim Array_ErrorData(ubound(Array_EntryData,1)-1,ubound(Array_EntryData,2)-1)

	for i = 1 to ubound(Array_EntryData,2)
		Array_OutputData(0,i-1) = Array_EntryData(1,i)
		Array_ErrorData(0,i-1) = Array_EntryData(1,i)
	next 

	resultCol = ubound(Array_EntryData,2)
	Array_OutputData(0,resultCol) = "処理結果"
	Array_OutputData(0,resultCol+1) = "処理実行日時"
	Array_OutputData(0,resultCol+2) = "エラー内容"
	outputRow =1
	erroutputRow = 1


	'レコードIDの一覧を繰り返す
	for each recordId in recordIdsDictionary.Keys
		on error resume next
		flgFirstRow = True
		errmsg = ""


		'アクセストークンの有効性を確認
		call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

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

			if .status = 200 then
				'既存のテーブルのデータをjsonパラメータにする。
				set json_record = json.record
				set json_Table = kntn_ScriptEngine.Run("getProperty", json_record,tableCode)
				set json_Table = json_Table.value
				recordInfo=""


					for each json_Row in json_Table
						rowId = json_Row.id
						
						tmpRecordInfo = "{""id"":""" & rowid & """,""value"":["
						set json_RowValue = json_Row.value
						rowValues =""
						for i = 0 to ubound(array_header,2)
							fieldcode = array_header(0,i)
							fieldType = array_header(2,i)
							if fieldcode = "" then exit for							
							fileKeys  = ""
							'既存の値を取得する。
							if fieldType = "FILE" then
							'既存のファイルキー一覧を取得する
								fieldvalue = Kntn_getFieldValue("FILE_KEY",fieldcode,json_RowValue,kugirimoji)
									
								'区切り文字に従って区切り配列化
								array_ExistsFiles = Kntn_SplitFiles(fieldValue,kugirimoji)


								for idx = 0 to ubound(array_ExistsFiles)
									if fileKeys = "" then 
										fileKeys = "{""fileKey"":""" & array_ExistsFiles(idx) & """}"
									else
										fileKeys = fileKeys & ",{""fileKey"":""" & array_ExistsFiles(idx) & """}"
									end if
								next

							else
								fieldvalue = Kntn_getFieldValue(fieldtype,fieldcode,json_RowValue,kugirimoji)
							end if


						'json形式に変換する
						if fieldtype = "FILE" then
							keyAndValue = """" & fieldcode & """:{""value"":[" & fileKeys & "]}" 
						else
							keyAndValue =  Kntn_CreateJsonKeyAndValue(fieldCode,fieldType,fieldvalue,"",kugirimoji)
						end if

							if rowValues = "" then
								rowValues = keyAndValue
							elseif keyAndValue <> "" then
								rowValues = rowValues & "," & keyAndValue
							end if
						next

						tmpRecordInfo = "{""id"":" & rowid & ",""value"":{" & rowValues & "}}"
						if recordInfo = "" then
							recordInfo = tmpRecordInfo
						else
							recordInfo	= recordInfo & "," & tmpRecordInfo
						end if
					next


				'追加したいテーブルの行のデータをjsonに追加する
				Array_RowNumAndRowId = recordIdsDictionary.Item(recordId)

				for each rownum in Array_RowNumAndRowId
					rowValues =""
					for i = 2 to ubound(array_header,2)
						fieldcode = array_header(0,i)
						fieldType = array_header(2,i)
						if fieldcode= "" then exit for

						'入力ファイルに存在しない列は初期値を設定する。
						if headerDictionary.exists(fieldcode) then
							fieldvalue=Array_EntryData(rowNum,headerDictionary.Item(fieldcode))
						else
							fieldvalue = ""
						end if
						'空白なら初期値のため、設定しない。
						if fieldValue <> "" then
							if fieldType = "FILE" then
								'配列に分解する。
								array_Files = Kntn_SplitFiles(fieldvalue,kugirimoji)
								for idx = 0 to ubound(array_Files)
									'ファイルを一時領域にアップロードし、ファイルキーを取得する。
									array_Files(idx) = Kntn_tmpUpload(array_Files(idx))
								next

								fileKeys  = ""
								for idx = 0 to ubound(array_Files)
									if fileKeys = "" then 
										fileKeys = "{""fileKey"":""" & array_Files(idx) & """}"
									else
										fileKeys = fileKeys & ",{""fileKey"":""" & array_Files(idx) & """}"
									end if
								next
								fieldKeyValue = """" & fieldcode & """:{""value"":[" & fileKeys & "]}" 
							else
								fieldKeyValue = Kntn_CreateJsonKeyAndValue(fieldCode,fieldType,fieldvalue,brank,kugirimoji)
							end if


							if rowValues = "" then
								rowValues = fieldKeyValue
							elseif fieldKeyValue <> "" then
								rowValues = rowValues & "," & fieldKeyValue
							end if
						end if
					next
					tmpRecordInfo = "{""value"":{" & rowValues & "}}"
					if recordInfo = "" then
						recordInfo = tmpRecordInfo
					else
						recordInfo	= recordInfo & "," & tmpRecordInfo
					end if
				next
	
				sendData = "{""app"":" & kntn_app_id & ",""id"":"& recordId &",""record"": {""" _
										& tableCode & """:{""value"":["  &	recordInfo & "]}}}"
	
				if err.Number <> 0 then
					errmsg =err.description
					statuscode = 0
				else 
					'アクセストークンの有効性を確認
					call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

					'更新APIを呼び出す。
					.Open "Put", kntn_api_uri , False
					.setRequestHeader "Authorization", "Bearer " & kntn_access_token
					.setRequestHeader "Content-Type", "application/json"
					.send sendData

					' レスポンステキストを取得
					responseText = .responseText
					statusCode = .status
				end if
			end if


			on error goto 0
			'StatusCodeが200なら成功。
			if statusCode = 200 then
				successRecordCount= successRecordCount + 1

				'各行のデータを処理結果の配列に記載する。同一レコードの頭の行のみに結果を記載する。
				for each rowNum in recordIdsDictionary.Item(recordId)
					'先頭行のみ結果を記載する
					if 	flgFirstRow = True then
						Array_OutputData(outputRow,resultCol) = "成功"
						Array_OutputData(outputRow,resultCol+1) = Kntn_GetNow
						Array_OutputData(outputRow,resultCol+2) = ""
						flgFirstRow = False
					else
						Array_OutputData(outputRow,resultCol) = "-"
						Array_OutputData(outputRow,resultCol+1) = "-"
						Array_OutputData(outputRow,resultCol+2) = "-"
						flgFirstRow = False
					end if

					'OutPutDataは既存と同じデータを配置
					for j = 1 to ubound(Array_EntryData,2)
						Array_OutputData(outputRow,j-1) = Array_EntryData(rowNum,j)
					next
					outputRow = outputRow+1
				next 
			elseif errmsg = "" then 
				errmsg = "Kintoneのテーブル追加操作に失敗しました。" & vbCrLf & _
				"ステータスコード：" & statusCode  & vbCrLf & _
				"レスポンス: " & Kntn_GetErrorMessage(responseText)
			end if

			if errmsg <> "" then
				errorRecordCount = 	errorRecordCount +1
				'各行のデータを処理結果の配列に記載する。同一レコードの頭の行のみに結果を記載する。
				for each rowNum in recordIdsDictionary(recordId)
					'先頭行のみ結果を記載する
					if 	flgFirstRow = True then
						Array_OutputData(outputRow,resultCol) = "失敗"
						Array_OutputData(outputRow,resultCol+1) = Kntn_GetNow
						Array_OutputData(outputRow,resultCol+2) =errMsg
						flgFirstRow = False
					else
						Array_OutputData(outputRow,resultCol) = "-"
						Array_OutputData(outputRow,resultCol+1) = "-"
						Array_OutputData(outputRow,resultCol+2) = "-"
						flgFirstRow = False
					end if

					'OutPutDataは既存と同じデータを配置
					for j = 1 to ubound(Array_EntryData,2)
						Array_OutputData(outputRow,j-1) = Array_EntryData(rowNum,j)
						Array_ErrorData(erroutputRow,j-1) = Array_EntryData(rowNum,j)
					next
					outputRow = outputRow+1
					erroutputRow = erroutputRow +1
				next 
			end if
		end With
	next

	SetUmsVariable $*成功レコード件数$,successRecordCount
	SetUmsVariable $*失敗レコード件数$,errorRecordCount
	
	'処理結果ファイルを作成する
	call Kntn_SetArrayToExcel(outputFilePath,"エラーデータ",Array_ErrorData,"A1",false)
	call Kntn_SetArrayToExcel(outputFilePath,"処理結果",Array_OutputData,"A1",true)

End Sub

