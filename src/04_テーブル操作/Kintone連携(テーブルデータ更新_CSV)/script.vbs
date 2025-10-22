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
Call Kntn_CheckAccessToken(kntn_client_id)

Call Kntn_updateTableDataByCsv()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_updateTableDataByCsv()
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
  csvFilePath = !*入力CSVファイルパス!
	charcode =!*文字コード|shift-jis,utf-8!
  outputFolderPath = !*処理結果出力フォルダ!


  successRecordCount=0
  errorRecordCount=0

  dim recordCount 
  If kntn_app_id = "" Then
    Err.Raise 1, "", "アプリIDが指定されていません。"
  End If  

	if fieldtable = "" then
		Err.Raise 1, "","「テーブルのフィールド名またはフィールドコード」の入力は必須です。"
	end if

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

	if csvFilePath = "" then
		Err.Raise 1, "","「入力CSVファイルパス」の入力は必須です。"
	end if


  if outputFolderPath = "" then
		Err.Raise 1, "","「処理結果出力フォルダ」の入力は必須です。"
	end if


  Dim objFSO
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
  If not objFSO.FileExists(csvFilePath)  Then
    Err.Raise 1, "", "入力CSVファイルが存在しません。ファイルパスを確認してください。"
  end if

	'いったんファイル名の.csvは削除する
	csvFileName =  objfso.GetFileName(csvFilePath)
	extension = objFso.GetExtensionName(csvFileName)
	if StrComp(extension,"csv",0) = 0 then
		csvFileName = left(csvFileName,len(csvFileName)-(len(extension)+1))
	end if 

	If not objFSO.folderExists(outputFolderPath)  Then
			call KNTN_CreateIntermediateFolders(outputFolderPath)
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
	on error resume Next
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
	Array_EntryData = KNTN_ReadCsv(csvFilePath, charcode,1)

  'ヘッダー情報をまとめたRootArrayを求める
  call Kntn_getHeaderArray(json_TableInfo.fields,outputType,nameOrCode,kntn_rootArray,fieldtable,lookUpArray,True)

	if nameOrCode = "フィールド名" then
    rowNum = 1
		recordIdColName="レコードID"
		rowIdColName = "テーブル行ID"
	else 
    rowNum = 0   
		recordIdColName="$id"
		rowIdColName = "rowId"
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


	'レコードIDの一覧とそれに対応する行とテーブルIDを辞書型配列に持たせる。
	Dim recordIdsDictionary
	Set recordIdsDictionary = WScript.CreateObject("Scripting.Dictionary")
	dim array_records()
	for i = 2 to ubound(Array_EntryData,1)
		recordId = Array_EntryData(i,1)
		rowId = Array_EntryData(i,headerDictionary.Item("rowId"))
		if recordIdsDictionary.Exists(recordId) Then
			exist_array_records = recordIdsDictionary.Item(recordId)
			idx = ubound(exist_array_records,2)+1
			redim preserve exist_array_records(1,idx)
			exist_array_records(0,idx) = i
			exist_array_records(1,idx) = rowId
			recordIdsDictionary.Item(recordId)=exist_array_records
		else
			redim array_records(1,0)
			array_records(0,0) = i
			array_records(1,0) = rowId

			'レコード行一覧を持たせる
			recordIdsDictionary.Add recordId,array_records
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

	dim array_existsRowIds()

	'レコードIDの一覧を繰り返す
	for each recordId in recordIdsDictionary.Keys
		on error resume next
		flgFirstRow = True
		errmsg = ""

		'処理した行IDの一覧をし配列化する。
		redim array_existsRowIds(0)

		'アクセストークンの有効性を確認
		call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 
		
		'既存のIDのレコード情報を取得する。
	  With wscript.CreateObject("MSXML2.XMLHTTP")
			.Open "Get", kntn_api_uri & "?app=" & kntn_app_id & "&id=" & recordId, False
			If kntn_proxy_url <> "" Then .setProxy 2, kntn_proxy_url
			.setRequestHeader "User-Agent", kntn_userAgent
			.setRequestHeader "Authorization", "Bearer " & kntn_access_token
			.send 

			' レスポンステキストを取得
			responseText = .responseText
			statusCode = .status
			Set json = kntn_ScriptEngine.CodeObject.Parse(responseText)
		end with

		if statusCode = 200 then
			'既存のテーブルのデータをjsonパラメータにする。
			set json_record = json.record
			set json_Table = kntn_ScriptEngine.Run("getProperty", json_record,tableCode)
			set json_Table = json_Table.value
			recordInfo=""

			Array_RowNumAndRowId = recordIdsDictionary.Item(recordId)

			if json_Table.length > 0 then
				for each json_Row in json_Table
					rowId = json_Row.id
					set json_RowValue = json_Row.value
	        rowValues =""
					'更新対象のrowIDかを確認する
					rowNum = -1
					for  i = 0 to ubound(Array_RowNumAndRowId,2)
						if clng(Array_RowNumAndRowId(1,i)) = clng(rowId) then
							rowNum = Array_RowNumAndRowId(0,i)
							'存在した行ID一覧の配列に記載する
							array_existsRowIds(ubound(array_existsRowIds)) = rowid	
							redim preserve array_existsRowIds(ubound(array_existsRowIds)+1)	
							exit for
						end if
					next

					

					for i = 2 to ubound(array_header,2)
						fieldcode = array_header(0,i)
						fieldType = array_header(2,i)
						fileKeys  = ""
						if fieldcode = "" then exit for
						'既存の値を取得する。
						if fieldType = "FILE" then
							fieldvalue = Kntn_getFieldValue("FILE_KEY",fieldcode,json_RowValue,kugirimoji)
							'既存のファイルキー一覧を取得する
							array_ExistsFiles = Kntn_SplitFiles(fieldvalue,kugirimoji)
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


						'更新対象のrowIDのとき
						if rowNum >= 0 then
							if headerDictionary.Exists(fieldcode) then
								tmpData = Array_EntryData(rownum,headerDictionary.Item(fieldcode))
								if tmpData <> "" then
									if fieldType = "FILE" then
										'配列に分解する。
										array_Files = Kntn_SplitFiles(tmpData,kugirimoji)
										for idx = 0 to ubound(array_Files)
											'ファイルを一時領域にアップロードし、ファイルキーを取得する。
											array_Files(idx) = Kntn_tmpUpload(array_Files(idx))
										next
										'今回アップロードしたファイルを追加する。
										for idx = 0 to ubound(array_Files)
											if fileKeys = "" then 
												fileKeys = "{""fileKey"":""" & array_Files(idx) & """}"
											else
												fileKeys = fileKeys & ",{""fileKey"":""" & array_Files(idx) & """}"
											end if
										next
										keyAndValue = """" & fieldcode & """:{""value"":[" & fileKeys & "]}" 
									else
									
										keyAndValue  = Kntn_CreateJsonKeyAndValue(fieldCode,fieldType,tmpData,brank,kugirimoji)							
									end if
								else
									'セルが空白ならそのまま値を設定する。
									keyAndValue  = Kntn_CreateJsonKeyAndValue(fieldCode,fieldType,fieldvalue,"",kugirimoji)
								end if
							else 
								keyAndValue  = Kntn_CreateJsonKeyAndValue(fieldCode,fieldType,fieldvalue,"",kugirimoji)
							end if

						else
							'json形式に変換する
							if fieldtype = "FILE" then
								keyAndValue = """" & fieldcode & """:{""value"":[" & fileKeys & "]}" 
							else
								keyAndValue =  Kntn_CreateJsonKeyAndValue(fieldCode,fieldType,fieldvalue,"",kugirimoji)
							end if
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
			end if	

			'存在しない行IDを連携していないか確認する
			notExistsRowIds = ""
			for i = 0 to ubound(Array_RowNumAndRowId,2)
				blnExistsRowId = false
				rowid = clng(Array_RowNumAndRowId(1,i))
				'行IDが空なら処理を抜ける
				if rowid = 0 then 
					exit for
				end if	

				for j = 0 to ubound(array_existsRowIds)
					if clng(array_existsRowIds(j)) = rowid then
						blnExistsRowId  = True
						exit for	
					end if
				next
		
				'もし存在しない行IDなら、エラーメッセージに追加する。
				if blnExistsRowId  = false then 
					notExistsRowIds =  notExistsRowIds & "," & rowid
				end if			
			next



			if err.Number <> 0 then
				errmsg =err.description
				statuscode = 0

			elseif rowid = 0 then  
				errmsg = "テーブルの行IDがブランクのデータが存在します。"
				statuscode = 0

			elseif notExistsRowIds <> "" then
				errmsg = "テーブルの行ID『" & Right(notExistsRowIds,len(notExistsRowIds)-1) & "』が存在しません。行IDを確認してください。"
				statuscode = 0
			else 
				'アクセストークンの有効性を確認
				call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 
				sendData = "{""app"":" & kntn_app_id & ",""id"":"& recordId &",""record"": {""" _
										& tableCode & """:{""value"":["  &	recordInfo & "]}}}"

				'更新APIを呼び出す。
				With wscript.CreateObject("MSXML2.XMLHTTP")
					.Open "Put", kntn_api_uri , False
					If kntn_proxy_url <> "" Then .setProxy 2, kntn_proxy_url
					.setRequestHeader "Authorization", "Bearer " & kntn_access_token
					.setRequestHeader "Content-Type", "application/json"
					.setRequestHeader "User-Agent", kntn_userAgent
					.send sendData
					' レスポンステキストを取得
					responseText = .responseText
					statusCode = .status
				end with
			end if
		end if

		on error goto 0
		'StatusCodeが200なら成功。
		if statusCode = 200 then
			successRecordCount= successRecordCount + 1
			'各行のデータを処理結果の配列に記載する。同一レコードの頭の行のみに結果を記載する。
			Array_RowNumAndRowId = recordIdsDictionary.Item(recordId) 
			for k = 0 to ubound(Array_RowNumAndRowId,2)
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
					Array_OutputData(outputRow,j-1) = Array_EntryData(Array_RowNumAndRowId(0,k),j)
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
			Array_RowNumAndRowId = recordIdsDictionary.Item(recordId) 

			flgFirstRow = true
			for k = 0 to ubound(Array_RowNumAndRowId,2)
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
					Array_OutputData(outputRow,j-1) = Array_EntryData(Array_RowNumAndRowId(0,k),j)
					Array_ErrorData(erroutputRow,j-1) = Array_EntryData(Array_RowNumAndRowId(0,k),j)
				next
				outputRow = outputRow+1
				erroutputRow = erroutputRow +1
			next 
		end if

	next

	SetUmsVariable $*成功レコード件数$,successRecordCount
	SetUmsVariable $*失敗レコード件数$,errorRecordCount
	
'処理結果ファイルとエラーファイルのパスを作成する。
	dateTimeNow = Year(now()) & Right("0" & Month(Now()) , 2) & Right("0" & Day(Now()) , 2) & _
								Right("0" & Hour(Now()) , 2)  & Right("0" & Minute(Now()) , 2) &  Right("0" & Second(Now()) , 2)
	outputCsvPath = objfso.BuildPath(outputFolderPath,csvFileName & "_処理結果_" &dateTimeNow & ".csv")
	errorCsvPath= objfso.BuildPath(outputFolderPath,csvFileName & "_エラーデータ_" &dateTimeNow & ".csv")

	'処理結果ファイルを作成する
	if errorRecordCount > 0 then
		call KNTN_SaveCsv(Array_ErrorData,errorCsvPath,charcode,0)
	end if
	call KNTN_SaveCsv(Array_OutputData,outputCsvPath,charcode,0)

End Sub


