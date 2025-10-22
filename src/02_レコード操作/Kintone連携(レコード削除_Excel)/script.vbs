' -----------------------------------------------------------------------
' Main
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
'Kintone連携（アクセストークン）が配置されているかの確認
if isempty(kntn_client_id) then
  err.raise 1,"","WinActor for kintone ver1.1.0 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
end if

if isempty(kntn_userAgent) then
  err.raise 1,"","WinActor for kintone ver1.1.0 以降の『01_認証>Kintone連携(アクセストークン取得)』ライブラリを配置してください。"
end if

' トークンのチェック
Call kntn_CheckAccessToken(kntn_client_id)

' Kintoneの指定アプリのフィールド情報を取得
Call Kntn_DeleteRecords()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_DeleteRecords()
  Dim kntn_api_uri
  Dim responseText
  Dim sendData
  Dim kntn_app_id
  dim excelFilePath

  kntn_app_id = !*アプリID!
  kntn_guestspace_id = !ゲストスペースID!
  nameOrCode = !*ヘッダー|フィールド名,フィールドコード!
  excelFilePath = !*入力ファイルパス!
  excelSheetName = !*シート名!
  outputFilePath = !*処理結果ファイルパス!
	canOverWriteFile = !*処理結果ファイルが既に存在するとき|上書き,エラー!

  Dim objFSO
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
  If objFSO.FileExists(outputFilePath) = True and canOverWriteFile = "エラー" Then
    Err.Raise 1, "", "処理結果ファイルが既に存在しています。既存ファイルを移動する、または処理結果ファイルパスを変更してください。"
  end if


  If kntn_app_id = "" Then
    Err.Raise 1, "", "「アプリID」の入力は必須です。"
  End If  

	if excelFilePath = "" then
		Err.Raise 1, "","「入力ファイルパス」の入力は必須です。"
	end if

	if excelSheetName = "" then
		Err.Raise 1, "","「シート名」の入力は必須です。"
	end if

  if outputFilePath = "" then
		Err.Raise 1, "","「処理結果ファイルパス」の入力は必須です。"
	end if

	'初期化
	successRecordCount = 0
	errorRecordCount = 0

  'Kintone API のエンドポイント
  if kntn_guestspace_id = "" then
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/records.json"
  else
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/records.json"
  end if

	'入力ファイルのデータを配列に格納する
  Dim Array_EntryData
	Array_EntryData = Kntn_GetArraybyExcel(ExcelFilePath,excelSheetName)

	'1行目にレコードIDが存在していることを確認する
	Dim i
	dim idCol
	idCol = -1

	if  nameOrCode  = "フィールド名" then
		recordIdColName = "レコードID"
	else
		recordIdColName = "$id"
	end if 

	for i = 1 to ubound(Array_EntryData,2)
		 if Array_EntryData(1,i) = recordIdColName then
		 	idCol = i
			exit for
		 end if
	next

	if idCol < 0 then
		Err.Raise 1, "", "『" & recordIdColName & "』列が存在しません。"
	end if

	'処理結果の配列とエラー出力の配列のヘッダー部のみ作成
	dim Array_OutputData(),Array_ErrorData()
	redim Array_OutputData(ubound(Array_EntryData,1)-1,ubound(Array_EntryData,2)-1+3)
	redim Array_ErrorData(ubound(Array_EntryData,1)-1,ubound(Array_EntryData,2)-1)

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
	Array_OutputData(0,resultCol+2) = "エラー内容"


	'配列を繰り返す
	for i = 2 to ubound(Array_EntryData,1)
		recordId = Array_EntryData(i,idCol)
		sendData = "{app:" &  kntn_app_id & ",ids:[" & recordId  & "]}" 
		'アクセストークンの有効性を確認
		call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 
		' API リクエストのヘッダーとデータを設定
		With wscript.CreateObject("MSXML2.XMLHTTP")
			.Open "DELETE", kntn_api_uri, False
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
				successRecordCount = successRecordCount +1
				Array_OutputData(i-1,resultCol) = "成功"
				Array_OutputData(i-1,resultCol+1) = Kntn_GetNow
				Array_OutputData(i-1,resultCol+2) = ""

			Case Else
				'処理結果配列に格納する
				errorRecordCount = errorRecordCount +1
				errmsg = "Kintoneのレコード削除操作に失敗しました。" & vbCrLf & _
								"ステータスコード：" & statusCode  & vbCrLf & _
								"レスポンス: " & Kntn_GetErrorMessage(responseText)
				Array_OutputData(i-1,resultCol) = "失敗"
				Array_OutputData(i-1,resultCol+1) = Kntn_GetNow
				Array_OutputData(i-1,resultCol+2) = errmsg
				
				'エラー配列に既存データを格納
				for j = 1 to ubound(Array_EntryData,2)
					Array_ErrorData(errorRecordCount,j-1) = Array_EntryData(i,j) 
				next 
			End Select
		End With
	next

	SetUmsVariable $*成功レコード件数$,successRecordCount
	SetUmsVariable $*失敗レコード件数$,errorRecordCount

	'処理結果ファイルを作成する
	call Kntn_SetArrayToExcel(outputFilePath,"エラーデータ",Array_ErrorData,"A1",false)
	call Kntn_SetArrayToExcel(outputFilePath,"処理結果",Array_OutputData,"A1",true)



End Sub

