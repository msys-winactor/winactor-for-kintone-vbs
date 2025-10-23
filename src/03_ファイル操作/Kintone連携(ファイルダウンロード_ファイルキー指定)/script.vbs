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


' Kintoneのファイルダウンロード
Call KNTN_DownloadFileByKey()

' -----------------------------------------------------------------------
' Sub / Function
' Kintoneのファイルダウンロード
' -----------------------------------------------------------------------
Sub KNTN_DownloadFileByKey()
  Dim kntn_api_uri
  dim savePath
  kntn_guestspace_id = !ゲストスペースID!
  dlFileKeys = !*ファイルキー! 
  dlFileNames = !*ファイル名!
  kugirimoji = !*複数ファイル指定時の区切り文字!
  dlFolderPath = !*保存先フォルダパス!
  
  If dlFileKeys = "" Then
    Err.Raise 1, "", "ファイルキーの入力は必須です。"
  End If  

  if dlFileNames = "" then 
    Err.Raise 1, "", "ファイル名の入力は必須です。"
  end if

  If dlFolderPath = "" Then
    Err.Raise 1, "", "保存先フォルダパスの入力は必須です。"
  End If  

 If kugirimoji = "" Then
    Err.Raise 1, "", "複数ファイル選択時の区切り文字の入力は必須です。"
  End If  

  '区切り文字がvbCrlfなら改行区切りとする
  if kugirimoji = "vbCrlf" then
    kugirimoji=vbCrLf
  end if

	dim objFso
	Set objFso = CreateObject("Scripting.FileSystemObject")
	'ダウンロードフォルダがない場合は作成
	If Not objfso.FolderExists(dlFolderPath) Then
		Call KNTN_CreateIntermediateFolders(dlFolderPath)
	End If

	'ダウンロード対象のファイル名とキーの一覧を配列化する
	array_dlFileNames =  Kntn_SplitFiles(dlFileNames,kugirimoji)
  array_dlFileKeys =  Kntn_SplitFiles(dlFileKeys,kugirimoji)

  if ubound(array_dlFileNames) <> ubound(array_dlFileKeys) then
    Err.Raise 1, "", "ファイルキーとファイル名の数が一致していません。区切り文字に注意し、ファイルキーとファイル名の再設定をしてください。"
  end if 

  'Kintone API のエンドポイント
  if kntn_guestspace_id = "" then
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/file.json"
  else
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/file.json"
  end if


  '同ファイル名が複数存在するときに連番をつけるため、ダウンロード回数を格納する辞書型配列を作成する
  dim dicDLCount
  set dicDLCount = WScript.CreateObject("Scripting.Dictionary")


  for idx = 0 to ubound(array_dlFileKeys)
    
    dlfileKey = array_dlFileKeys(idx)
    dlfileName = array_dlFileNames(idx)

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
  next
End Sub

