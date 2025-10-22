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

' Kintoneのレコードを取得
Call Kntn_GetRecordsToCsv()

' -----------------------------------------------------------------------
' Sub / Function
' -----------------------------------------------------------------------
Sub Kntn_GetRecordsToCsv()
  dim unUsedCheckMark 
  unUsedCheckMark ="（×）"

  Dim kntn_api_uri
  Dim responseText
  Dim sendData
  Dim kntn_app_id
  dim csvFilePath 
  kntn_app_id = !*アプリID!
  kntn_guestspace_id = !ゲストスペースID!
  nameOrCode = !*ヘッダー|フィールド名,フィールドコード!
  outputType = !*出力フィールドタイプ|全フィールド取得,レコード（更新・削除用）フィールド取得,レコード登録用フィールド取得（ヘッダーのみ）!
  flgtable = !テーブルデータ有無|有り,無し!
  kugirimoji = !*複数設定値の区切り文字!
  saveFolderPath  = !*保存先フォルダパス!
  csvFileName = !*CSVファイル名!
	charcode =!*文字コード|shift-jis,utf-8!
	canOverWriteFile = !*保存先にCSVファイルが既に存在するとき|上書き,エラー!
  recordCount = $*取得レコード件数$
  field1 = !項目1!
  condition1 = !項目1の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
  value1 = !項目1の値!
  field2 = !項目2!
  condition2 = !項目2の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
  value2 = !項目2の値!
  field3 = !項目3!
  condition3 = !項目3の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
  value3 = !項目3の値!
  field4 = !項目4!
  condition4= !項目4の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
  value4 = !項目4の値!
  field5 = !項目5!
  condition5 = !項目5の条件|等しい(=),等しくない(<>),含む(like),含まない(not like),以上(≧),以下(≦),より大きい(>),より小さい(<)!
  value5 = !項目5の値!
  freequery = !カスタムクエリ!


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

	if saveFolderPath  = "" then
		Err.Raise 1, "","「保存先フォルダパス」の入力は必須です。"
	end if

	if csvFileName  = "" then
		Err.Raise 1, "","「CSVファイル名」の入力は必須です。"
	end if



  Dim objFSO
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

		'いったんファイル名の.csvは削除する
	extension = objFso.GetExtensionName(csvFileName)
	if extension ="" then
    Err.Raise 1, "", "CSVファイル名の拡張子を指定してください。"
  else
		csvFileNameNoExtention = left(csvFileName,len(csvFileName)-(len(extension)+1))
	end if 
	
  If objFSO.FileExists(objfso.BuildPath(saveFolderPath,csvFileName)) = True and canOverWriteFile = "エラー" Then
    Err.Raise 1, "", "保存先にCSVファイルが既に存在しています。既存ファイルを移動する、または保存先を変更してください。"
  end if


  if flgtable  = "有り" then
    flgtable  = True
  else
    flgTable = false
  end if

  'フィールド情報一覧の情報をAPIで取得する。
  dim json_fieldsInfo,json_properties
  Set json_fieldsInfo= kntn_ScriptEngine.CodeObject.Parse(Kntn_GetFieldsInfo(kntn_app_id,kntn_guestspace_id))
  set json_properties = json_fieldsInfo.properties

  dim array_fieldsinfo,kntn_rootArray,outputType_rootArray,lookUpArray,array_header()
	set kntn_rootArray = WScript.CreateObject("Scripting.Dictionary")
  set outputType_rootArray = WScript.CreateObject("Scripting.Dictionary")

    'Lookupのmappingの項目の一覧を取得する
  lookUpArray = Kntn_CreateLookUpArray(json_properties)

  'SUBTABLEを含めてヘッダー情報をまとめたRootArrayを求める
  call Kntn_getHeaderArray(json_properties,"全フィールド取得",nameOrCode,kntn_rootArray,csvFileName,lookUpArray,False)
  call Kntn_getHeaderArray(json_properties,outputType,nameOrCode,outputType_rootArray,csvFileName,lookUpArray,False)

  if nameOrCode  = "フィールド名" then
    rowNum = 1
  else 
    rowNum = 0    
  end if

  'レコード登録用フィールド取得（ヘッダーのみ）のときはフィールド情報から取得する
  if outputType = "レコード登録用フィールド取得（ヘッダーのみ）" then
		
    for each key in Kntn_rootArray.keys
      'メイン情報はレコードIDを除く
      if key = csvFileName then
				csvFilePath	= objfso.BuildPath(saveFolderPath,csvFileNameNoExtention & ".csv")
        array_fieldsinfo = Kntn_rootArray.Item(key)  
        array_usefieldsinfo = outputType_rootArray.Item(Key) 
        redim array_header(0,ubound(array_fieldsinfo,2))
        colNum = 0
        for i = 0 to ubound(array_fieldsinfo,2)
            if array_fieldsinfo(2,i) ="FILE" then
              'FILEタイプはファイル名列とファイルキー列に分ける
              array_header(0,colnum) =unUsedCheckMark & array_fieldsinfo(rowNum,i)& "_Name"
              '配列の列を増やす
              redim preserve array_header(0,ubound(array_header,2)+1)
              colNum = colNum + 1 
              array_header(0,colNum) = unUsedCheckMark & array_fieldsinfo(rowNum,i) & "_Key"

              '配列の列を増やす
              redim preserve array_header(0,ubound(array_header,2)+1)
              colNum = colNum + 1 
              array_header(0,colnum) =array_fieldsinfo(rowNum,i)

            else
              blnUsedField = KNTN_CheckUseField(array_fieldsinfo(0,i),array_usefieldsinfo)
              if blnUsedField then 
                array_header(0,colnum) = array_fieldsinfo(rowNum,i)
              else  
                array_header(0,colnum) = unUsedCheckMark & array_fieldsinfo(rowNum,i)
              end if
            end if
            colNum = colNum + 1 
        next
        call KNTN_SaveCsv(array_header,csvFilePath,charcode,0)  

      elseif flgtable= false  then 
        'テーブル取得なしでメインのシート名ではない場合はスキップ

      else
				csvFilePath	= objfso.BuildPath(saveFolderPath,csvFileNameNoExtention & "_" &KNTN_ReplaceForbiddenChar(key) & ".csv")
        '各テーブルのフィールド情報を取得し、1行目に格納し、Excelに貼り付ける。
        array_fieldsinfo = Kntn_rootArray.Item(key)
        array_usefieldsinfo = outputType_rootArray.Item(Key) 
        redim array_header(0,ubound(array_fieldsinfo,2))
        colNum = 0

        for i = 0 to ubound(array_fieldsinfo,2)
            if array_fieldsinfo(2,i) ="FILE" then
              'FILEタイプはファイル名列とファイルキー列に分ける
              array_header(0,colnum) =unUsedCheckMark & array_fieldsinfo(rowNum,i)& "_Name"
              '配列の列を増やす
              redim preserve array_header(0,ubound(array_header,2)+1)
              colNum = colNum + 1 
              array_header(0,colNum) = unUsedCheckMark & array_fieldsinfo(rowNum,i) & "_Key"

              '配列の列を増やす
              redim preserve array_header(0,ubound(array_header,2)+1)
              colNum = colNum + 1 
              array_header(0,colnum) =array_fieldsinfo(rowNum,i)

            else
              blnUsedField = KNTN_CheckUseField(array_fieldsinfo(0,i),array_usefieldsinfo)
              if blnUsedField then 
                array_header(0,colnum) = array_fieldsinfo(rowNum,i)
              else  
                array_header(0,colnum) = unUsedCheckMark & array_fieldsinfo(rowNum,i)
              end if
            end if
            colNum = colNum + 1 
        next
        call KNTN_SaveCsv(array_header,csvFilePath,charcode,0)       
      end if
    next

    exit sub
  end if

  'Kintone API のエンドポイント
	if kntn_guestspace_id = "" then
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/records.json"
	else
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/records.json"
	end if

  'queryの作成
  if freequery <> "" then
      query = freeQuery
  else
    if field1 <> "" then    
      tmpQuery = Kntn_CreateQuery(field1,condition1,value1,kugirimoji,nameOrCode,json_properties)
      query = tmpQuery
    end if

    if field2 <> "" then
      tmpQuery = Kntn_CreateQuery(field2,condition2,value2,kugirimoji,nameOrCode,json_properties)
      query = query & " and " & tmpQuery
    end if

    if field3 <> "" then
      tmpQuery = Kntn_CreateQuery(field3,condition3,value3,kugirimoji,nameOrCode,json_properties)
      query = query & " and " & tmpQuery
    end if

    if field4 <> "" then
      tmpQuery = Kntn_CreateQuery(field4,condition4,value4,kugirimoji,nameOrCode,json_properties)
      query = query & " and " & tmpQuery
    end if

    if field5 <> "" then
      tmpQuery = Kntn_CreateQuery(field5,condition5,value5,kugirimoji,nameOrCode,json_properties)
      query = query & " and " & tmpQuery
    end if
  end if


  
  dim  array_records
  dim  array_table(),array_transposeTable()

  'ヘッダーの配列を格納する辞書型配列の作成
  dim kntn_rootHeaderArray
  set kntn_rootHeaderArray = WScript.CreateObject("Scripting.Dictionary")
  tableRecordCount = 0
  count = 0
  lastRecordID = ""
  
'レコードがなくなるまで繰り返す（500件ずつ取得のため）
  do 
    '500件以上あった場合は前回のレコードIDの次から取得する
    if lastRecordID = "" then
      if query = "" then 
        sendquery = """query"":""order by $id asc limit 500"""
      else
        sendquery = """query"":""" &  query  & " order by $id asc limit 500"""
      end if
    else 
      if query = "" then 
        sendquery = """query"":""$id > \""" & lastRecordID & "\"" order by $id asc limit 500"""
      else
        sendquery = """query"":""" &  query  & " and $id > \""" & lastRecordID  & "\"" order by $id asc limit 500"""
      end if
    end if

    ' リクエストデータを作成 Queryがない場合は全データ
    if field1 = "" and freequery = "" then
      sendData = "{""app"": " &  kntn_app_id  & "," _
              & """totalCount"":true," _
              & sendquery _
              & "}"
    else
      sendData = "{""app"": " &  kntn_app_id  & "," _
              & """totalCount"":true," _
              & sendquery _
              & "}"
    end if  

    'アクセストークンの有効性を確認
    call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

    ' API リクエストのヘッダーとデータを設定
    With WScript.CreateObject("MSXML2.XMLHTTP")
      .Open "Post", kntn_api_uri, False
      .setRequestHeader "Authorization", "Bearer " & kntn_access_token
      .setRequestHeader "Content-Type", "application/json"
      .setRequestHeader "X-HTTP-Method-Override", "GET"
      .setRequestHeader "User-Agent", kntn_userAgent
      .send sendData

      ' レスポンステキストを取得
      responseText = .responseText
      statusCode = .status
    end with 

    ' レスポンスの処理を行う
    if statusCode <> 200 then
      Err.Raise statusCode, "",  "Kintoneのレコード取得操作に失敗しました。" & vbCrLf & _
                "ステータスコード：" & statusCode  & vbCrLf & _
                "レスポンス: " & Kntn_GetErrorMessage(responseText)
    End if

    'レコードの数を返却する
    Set json = kntn_ScriptEngine.CodeObject.Parse(responseText)
    if count = 0 then SetUmsVariable recordCount,json.totalCount

    '残りレコードが501件以上なら繰り返し取得
    if json.totalCount > 500 then
        blnNextRecord = True						
    else
        blnNextRecord = false						
    end if

    'recordsプロパティの取得
    set json_records = json.records

    for each key in Kntn_rootArray.keys  
			'保存先csvのファイルパスを作成する。テーブルはテーブル名を付ける（禁則文字は全角化） 
			if key = csvFileName then
				csvFilePath	= objfso.BuildPath(saveFolderPath,csvFileNameNoExtention & ".csv")
			else
				csvFilePath	= objfso.BuildPath(saveFolderPath,csvFileNameNoExtention & "_" & KNTN_ReplaceForbiddenChar(key) & ".csv")
			end if

      'フィールド情報を取得する
      array_fieldsinfo = Kntn_rootArray.Item(key)
      array_usefieldsinfo = outputType_rootArray.Item(Key) 

      '1回目のループの場合はヘッダー情報を入力する
      if count = 0 then
        'headerの情報を取得する。Fileタイプに注意する
        redim  array_header(0,ubound(array_fieldsinfo,2))
        colnum = 0
        for i = 0 to ubound(array_fieldsinfo,2)
          if array_fieldsinfo(2,i) ="FILE" then
            '配列の列を増やす
            redim preserve array_header(0,ubound(array_header,2)+2)
            'FILEタイプはファイル名列とファイルキー列、デフォルトの3列に分ける
            array_header(0,colNum) =array_fieldsinfo(rowNum,i)& "_Name"
            colNum = colNum+1 
            array_header(0,colNum) = array_fieldsinfo(rowNum,i) & "_Key"
            colNum = colNum+1 
            array_header(0,colNum) = array_fieldsinfo(rowNum,i) 
          else
              blnUsedField = KNTN_CheckUseField(array_fieldsinfo(0,i),array_usefieldsinfo)
              if blnUsedField then 
                array_header(0,colnum) = array_fieldsinfo(rowNum,i)
              else  
                array_header(0,colnum) = unUsedCheckMark & array_fieldsinfo(rowNum,i)
              end if
          end if
          colNum=colNum+1
        next 

        '辞書型配列に格納
        kntn_rootHeaderArray.add key,array_header

        'Keyが引数のシート名と等しい（メイン情報）、もしくはテーブル取得ありのとき
        if key = csvFileName or flgtable = true then
          'ヘッダーのみ入力
        	call KNTN_SaveCsv(array_header,csvFilePath,charcode,0)    
        end if
      end if

      'メインレコードの取得 kntn_rootHeaderArray
      if key = csvFileName then
        array_records = Kntn_GetMainRecord(kntn_rootHeaderArray.Item(key),array_fieldsinfo,json_records,outputType,kugirimoji)
        'レコードを設定する。
				call KNTN_AddCsv(array_records ,csvFilePath,charcode,0,False)   
        '最終行から最終レコードIDを取得する。
        lastRecordID = array_records(499,0)
      elseif flgtable= false  then 
        '何もしない
      else
        'テーブルのデータ取得
        if nameOrCode = "フィールド名" then 
          set tmpJson =  Kntn_getFieldJson(key,nameOrCode,json_properties)
          tableCode = tmpJson.code
        else 
          tableCode = key
        end if
        array_records = Kntn_GetTableRecord(tableCode,kntn_rootHeaderArray.Item(key),array_fieldsinfo,json_records,outputType,kugirimoji)

        'レコードを設定する。
				call KNTN_AddCsv(array_records ,csvFilePath,charcode,0,False)   
      end if
    next 
    count =count + 1 
  loop while blnNextRecord=True
End Sub

' -----------------------------------------------------------------------
' Sub / Function
'登録時や更新時に利用するフィールドかの確認
' -----------------------------------------------------------------------
function KNTN_CheckUseField(fieldcode,Array_usefields)
  dim cnt 
  dim result
  result = false
  for cnt = 0 to ubound(Array_usefields,2)
    if fieldcode = Array_usefields(0,cnt) then
      result = true
      exit for
    end if
  next
  KNTN_CheckUseField= result
end function