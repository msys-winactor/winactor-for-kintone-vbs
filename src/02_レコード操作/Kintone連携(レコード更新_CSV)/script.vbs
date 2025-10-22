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

' Kintoneの指定アプリのフィールド情報を取得
Call KNTN_UpdateRecordsByCsv()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub KNTN_UpdateRecordsByCsv()
  Dim kntn_api_uri
  Dim responseText
  Dim sendData
  Dim kntn_app_id

  kntn_app_id = !*アプリID!
  kntn_guestspace_id = !ゲストスペースID!
  nameOrCode = !*ヘッダー|フィールド名,フィールドコード!
  kugirimoji = !*複数設定値の区切り文字!
  brank = !*ブランクとして認識する値!
  csvFilePath = !*入力CSVファイルパス!
	charcode =!*文字コード|shift-jis,utf-8!
  outputFolderPath = !*処理結果出力フォルダ!

  If kntn_app_id = "" Then
    Err.Raise 1, "", "「アプリID」の入力は必須です。"
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

	'初期化
	successRecordCount = 0
	errorRecordCount = 0

  'Kintone レコード更新API のエンドポイント
  if kntn_guestspace_id = "" then
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/record.json"
  else
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/record.json"
  end if

	'入力ファイルのデータを配列に格納する
  Dim Array_EntryData
	Array_EntryData = KNTN_ReadCsv(csvFilePath, charcode,1)


  'フィールド情報一覧の情報をAPIで取得する。
  dim json_fieldsInfo,json_field
  Set json_fieldsInfo= kntn_ScriptEngine.CodeObject.Parse(KNTN_GetFieldsInfo(kntn_app_id,kntn_guestspace_id))
  Set json_properties =json_fieldsInfo.properties

  dim array_fieldsinfo,RootArray_UodateApiFields,outputType_rootArray,lookUpArray,array_header
	set RootArray_UodateApiFields = WScript.CreateObject("Scripting.Dictionary")

  'Lookupのmappingの項目の一覧を取得する
  lookUpArray = KNTN_CreateLookUpArray(json_properties)

  'APIで利用するヘッダー情報をまとめたArrayを求める
  call KNTN_getHeaderArray(json_properties,"レコード更新用フィールド取得",nameOrCode,RootArray_UodateApiFields,csvFileName,lookUpArray,False)
  array_header = RootArray_UodateApiFields.item(csvFileName)

  'APIで連携するフィールドのコードと該当の列数を辞書型配列に格納する。
  dim dicTargetfields
  set dicTargetFields = WScript.CreateObject("Scripting.Dictionary")

  if nameOrCode = "フィールド名" then
    rowNum = 1
  else
    rowNum = 0
  end if

	'フィールドコードが入力ファイルの何列目かを辞書型配列にまとめる（毎回求めるのは大変なため）
	dim headerDictionary
	Set headerDictionary = WScript.CreateObject("Scripting.Dictionary")
	for i = 1 to ubound(Array_EntryData,2)
		tmpHeader = Array_EntryData(1,i)
    for j = 0 to ubound(array_header,2)
      if array_header(rownum,j) = tmpHeader then
        '辞書にフィールドコードと列名を追加する
        headerDictionary.Add array_header(0,j),i
        exit for
      end if	
    next
	next
	

  '重複する列名が存在するかをチェック
  call KNTN_checkDuplicateFields(Array_EntryData,1)

	'処理結果の配列とエラー出力の配列のヘッダー部のみ作成
	dim Array_OutputData(),Array_ErrorData()
	redim Array_OutputData(ubound(Array_EntryData,1)-1,ubound(Array_EntryData,2)-1+4)
	redim Array_ErrorData(ubound(Array_EntryData,1)-1,ubound(Array_EntryData,2))

	for i = 1 to ubound(Array_EntryData,2)
		Array_ErrorData(0,i-1) = Array_EntryData(1,i)

		'OutPutDataは既存と同じデータを配置
		for j = 1 to ubound(Array_EntryData,1)
			Array_OutputData(j-1,i-1) = Array_EntryData(j,i)
		next
	next 
  
	resultCol = ubound(Array_EntryData,2)
	Array_OutputData(0,resultCol) = "処理結果"
	Array_OutputData(0,resultCol+1) = "処理実行日時"
	Array_OutputData(0,resultCol+2)= "エラー内容"

  dim aryStrings
	dim json_record

	'配列を繰り返す
	for i = 2 to ubound(Array_EntryData,1)
    '配列内の要素を繰り返し、bodyを作成する。
    recordInfo = ""
    fieldKeyValue = ""
    errMsg =""
		set json_record =  Nothing
    
    for j = 0 to ubound(Array_Header,2)
      fieldCode = Array_Header(0,j)
      fieldType = Array_Header(2,j)
      
      '入力ファイルに存在しない列は初期値を設定する。
      if headerDictionary.exists(fieldcode) then
        fieldvalue=Array_EntryData(i,headerDictionary.Item(fieldcode))
      else
        fieldvalue = ""
      end if

      if fieldcode ="" then exit for
			if fieldcode ="$id" then recordID = fieldvalue
      fieldKeyValue =""
			fileKeys = ""

      on error resume next
      if fieldType = "FILE" and fieldvalue <> "" then
				'既存のレコード情報が空なら取得する。
				if json_record is nothing then
					'既存のIDのレコード情報を取得する。
					With wscript.CreateObject("MSXML2.XMLHTTP")
						.Open "Get", kntn_api_uri & "?app=" & kntn_app_id & "&id=" & recordId, False
						.setRequestHeader "Authorization", "Bearer " & kntn_access_token
						.setRequestHeader "User-Agent", kntn_userAgent
            .send 

						' レスポンステキストを取得
						responseText = .responseText
						statusCode = .status
						if statuscode = 200 then
							Set json_record = kntn_ScriptEngine.CodeObject.Parse(responseText)
							set json_record = json_record.record
						else
								errmsg = "Kintoneの既存レコード取得操作に失敗しました。" & vbCrLf & _
											"ステータスコード：" & statusCode  & vbCrLf & _
											"レスポンス: " & KNTN_GetErrorMessage(responseText)
								exit for			
						end if				
					end with
				end if
   
        fileKeys  = ""
				'既存のファイルキー情報を取得する。
					tmpData = KNTN_getFieldValue("FILE_KEY",fieldcode,json_record,kugirimoji)
					'既存のファイルキー一覧を取得する
          if tmpData <> "" then 
            array_ExistsFiles = Kntn_SplitFiles(tmpData,kugirimoji)
            for idx = 0 to ubound(array_ExistsFiles)
              if fileKeys = "" then 
                fileKeys = "{""fileKey"":""" & array_ExistsFiles(idx) & """}"
              else
                fileKeys = fileKeys & ",{""fileKey"":""" & array_ExistsFiles(idx) & """}"
              end if
            next
          end if

        '配列に分解する。
        array_Files = Kntn_SplitFiles(fieldvalue,kugirimoji)
        for idx = 0 to ubound(array_Files)
          'ファイルを一時領域にアップロードし、ファイルキーを取得する。
          array_Files(idx) = KNTN_tmpUpload(array_Files(idx))
        next

        for idx = 0 to ubound(array_Files)
          if fileKeys = "" then 
            fileKeys = "{""fileKey"":""" & array_Files(idx) & """}"
          else
            fileKeys = fileKeys & ",{""fileKey"":""" & array_Files(idx) & """}"
          end if
        next
        fieldKeyValue = """" & fieldcode & """:{""value"":[" & fileKeys & "]}" 
      else
        fieldKeyValue = KNTN_CreateJsonKeyAndValue(fieldCode,fieldType,fieldvalue,brank,kugirimoji)
      end if

      if err.Number <> 0 then
        errMsg = err.description
        exit for
      end if

      if fieldKeyValue = "" then
        '初期値を渡す場合は何もしない      
      elseif recordInfo = "" then
        recordInfo = fieldKeyValue
      else
        recordInfo = recordInfo & "," & fieldKeyValue
      end if
    next

    on error goto 0

    'もしエラーが出ているなら処理結果を失敗とする。
    if errMsg <> "" then
      '処理結果配列に格納する
      errorRecordCount = errorRecordCount +1
      Array_OutputData(i-1,resultCol) = "失敗"
      Array_OutputData(i-1,resultCol+1) = KNTN_GetNow
      Array_OutputData(i-1,resultCol+2) = errmsg
      
      'エラー配列に既存データを格納
      for j = 1 to ubound(Array_EntryData,2)
        Array_ErrorData(errorRecordCount,j-1) = Array_EntryData(i,j) 
      next

    else 
      sendData = "{app:" &  kntn_app_id & ",""id"":" & recordId   & ",record:{" & recordInfo  & "}}" 

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
          '処理結果配列に格納する
          set result_json = kntn_ScriptEngine.CodeObject.Parse(responseText)
          successRecordCount = successRecordCount +1
          Array_OutputData(i-1,resultCol) = "成功"
          Array_OutputData(i-1,resultCol+1) = KNTN_GetNow
          Array_OutputData(i-1,resultCol+2) = ""

        Case Else
          '処理結果配列に格納する
          errorRecordCount = errorRecordCount +1
          errmsg = "Kintoneのレコード更新操作に失敗しました。" & vbCrLf & _
                  "ステータスコード：" & statusCode  & vbCrLf & _
                  "レスポンス: " & KNTN_GetErrorMessage(responseText)
          Array_OutputData(i-1,resultCol) = "失敗"
          Array_OutputData(i-1,resultCol+1) = KNTN_GetNow
          Array_OutputData(i-1,resultCol+2) = errmsg
          
          'エラー配列に既存データを格納
          for j = 1 to ubound(Array_EntryData,2)
            Array_ErrorData(errorRecordCount,j-1) = Array_EntryData(i,j) 
          next 
        End Select
      End With
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

