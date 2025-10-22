' -----------------------------------------------------------------------
' Main
' -----------------------------------------------------------------------
' 認証用情報
kntn_user_id = !Kintoneユーザー名!
kntn_redirect_endpoint = !リダイレクトエンドポイント!
kntn_client_id = !クライアントID!
kntn_secret_key =  !クライアントシークレット!
kntn_subdomain = !サブドメイン!
kntn_crypt_keyword = !暗号化キーワード!
'プロキシサーバーはxmlhttprequestでは利用不可
'kntn_proxy_url = !プロキシサーバURL!
kntn_proxy_url = ""

' 共通変数の定義
Dim kntn_access_token
Dim kntn_expires_in
Dim kntn_refresh_token
Dim kntn_tokens_folder
Dim kntn_token_file

'User-Agentに渡す情報
kntn_userAgent = "WinActor for kintone" 

'json解析用のスクリプトエンジンの作成
dim kntn_ScriptEngine
call KNTN_InitScriptEngine()

' ドキュメントフォルダの取得
Set objWshShell = WScript.CreateObject("WScript.Shell")
documentFolder = objWshShell.SpecialFolders("MyDocuments")

' FSOの作成
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' ドキュメントフォルダの取得
Set objWshShell = WScript.CreateObject("WScript.Shell")
documentFolder = objWshShell.SpecialFolders("MyDocuments")

' トークンフォルダのパス作成
kntn_tokens_folder = fso.BuildPath(documentFolder, "WinActor\MsysAdapters\Kintone\tokens")
kntn_tokens_folder = fso.BuildPath(kntn_tokens_folder,kntn_user_id)

' トークンファイルのパス作成
kntn_file_name =  kntn_client_id & ".token"
kntn_token_file = fso.BuildPath(kntn_tokens_folder, kntn_file_name)

If Not fso.FolderExists(kntn_tokens_folder) Then
  Call KNTN_CreateIntermediateFolders(kntn_tokens_folder)
End If

' メモリの解放
Set fso = Nothing



'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★アクセストークン取得★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
' -----------------------------------------------------------------------
' Sub / Function
'jsonのバリューをkey指定で取得する
' -----------------------------------------------------------------------
Function KNTN_GetJSONValueByKey(jsonString, key)
    ' Create a ScriptControl object
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    
    ' Add the JSON string to the ScriptControl
    sc.AddCode "var json = " & jsonString & ";"
    
    ' Evaluate the key in JSON and return the value
    KNTN_GetJSONValueByKey = sc.Eval(key)
End Function


' -----------------------------------------------------------------------
' Sub / Function
'jsonのtagを取得する
' -----------------------------------------------------------------------
Function KNTN_GetTagValues(jsonString)
    Dim sc, i, entry, tagValue
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    
    ' Add the JSON string to the ScriptControl
    sc.AddCode "var json = " & jsonString & ";"
    
    ' Get the length of the entries array
    Dim length
    length = sc.Eval("json.entries.length")
    
    ' Iterate through the entries array and get the .tag values
    For i = 0 To length - 1
        KNTN_GetTagValues = sc.Eval("json.entries[" & i & "]['.tag']")
    Next
End Function

' -----------------------------------------------------------------------
' Sub / Function
'文字列を暗号化する
' -----------------------------------------------------------------------
Function KNTN_EncodePlainText(plainText, password)
  Dim plain
  Dim crypt

  salt = password
  saltlen = Len(salt)
  saltIndex = 1

  crypt = ""

  plain = plainText
  length = Len(plain)

  For i = 1 to length

    bit = asc(Mid(salt, saltIndex, 1))
    saltIndex = saltIndex + 1
    if ( saltIndex > saltLen ) Then
      saltIndex = 1
    End if

    num = asc(Mid(plain, i, 1)) xor bit

    padding = "0000" + hex(num)

    crypt = crypt + Right(padding,4)
  Next

  ' 変数に値を設定する
  KNTN_EncodePlainText = crypt

End Function

' -----------------------------------------------------------------------
' Sub / Function
'暗号文字列を復号化する
' -----------------------------------------------------------------------
Function KNTN_DecodeCryptText(encryptedText, password)
    Dim plain
    Dim crypt
    Dim num
    salt = password
    saltlen = Len(salt)
    saltIndex = 1
    l_mask = 65535

    plain = ""

    crypt = encryptedText
    length = Len(crypt)

    For i = 1 to length step 4

    bit = asc(Mid(salt, saltIndex, 1))
    saltIndex = saltIndex + 1
    if ( saltIndex > saltLen ) Then
        saltIndex = 1
    End if

    l_num = (int("&h" + Mid(crypt, i, 4)) Xor bit) And l_mask
    plain = plain + Chr(l_num)

    Next

    ' 変数に値を設定する
    KNTN_DecodeCryptText = plain
End Function

' -----------------------------------------------------------------------
' Sub / Function
'unicodeのエスケープ処理
' -----------------------------------------------------------------------
Function KNTN_DecodeUnicodeEscapes(str)
    Dim i, result, code
    result = ""
    i = 1
    Do While i <= Len(str)
        If Mid(str, i, 2) = "\u" Then
            code = Mid(str, i + 2, 4)
            result = result & ChrW("&H" & code)
            i = i + 6
        Else
            result = result & Mid(str, i, 1)
            i = i + 1
        End If
    Loop
    KNTN_DecodeUnicodeEscapes = result
End Function

'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★



'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★汎用的関数★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
' -----------------------------------------------------------------------
' Sub / Function
'フォルダを前階層含めて作成
' -----------------------------------------------------------------------
Sub KNTN_CreateIntermediateFolders(folderPath)
  Dim folderArray, folder, currentPath
  ' FSOの作成
  Dim objfso
  Set objfso = CreateObject("Scripting.FileSystemObject")

  ' フォルダパスを分割
  folderArray = Split(folderPath, "\")

  ' 初期化
  currentPath = ""

  ' 各フォルダを順番に処理
  For Each folder In folderArray
    ' 現在のフォルダパスを更新
      currentPath = currentPath & folder & "\"
    
    ' フォルダが存在しない場合は作成
    If Not objfso.FolderExists(currentPath) Then
        objfso.CreateFolder currentPath
    End If
  Next
End Sub

' -----------------------------------------------------------------------
' Sub / Function
'■文字列をBase64へエンコードを行う関数
' -----------------------------------------------------------------------
Function KNTN_EncodeBase64(text) 
    '■参照設定不要、オブジェクト準備
    Dim node,obj 
    Set node = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
    Set obj = CreateObject("ADODB.Stream")
  
    '■エンコード(textをBASE64へ変換)
    node.DataType = "bin.base64"
    With obj
        .Type = 2
        .Charset = "us-ascii"
        .Open
        .WriteText text
        .Position = 0
        .Type = 1
        .Position = 0
    End With
    node.nodeTypedValue = obj.Read
  
    '■改行を削除して返却(上記で取り除けない為)
    KNTN_EncodeBase64 = Replace(node.text, vbLf, "")
End Function

' -----------------------------------------------------------------------
' Sub / Function
' adobestreamのtype変更関数
' -----------------------------------------------------------------------
Function KNTN_ChangeStreamType(stream, t)
    p = stream.Position
    stream.Position = 0
    stream.Type = t
    stream.Position = p
    Set KNTN_ChangeStreamType = stream
End Function

' -----------------------------------------------------------------------
' Sub / Function
' KintoneのAPI実行結果のエラーメッセージを取得する。messageがない場合は、そのまま返す。
' -----------------------------------------------------------------------
function KNTN_GetErrorMessage(response)
  on error resume next
  dim json
  set json = kntn_ScriptEngine.CodeObject.Parse(response)

  if err.Number <> 0 then
    KNTN_GetErrorMessage = response
    exit function
  end if
  on error goto 0

  Dim objReg,errMsg
  Set objReg = CreateObject("VBScript.RegExp")

  existErrmsg = kntn_ScriptEngine.Run("checkKey", json,"message")
  if existErrmsg = True then
    errMsg= json.message
    'もしKintoneのエラーメッセージが「フィールド「○○」の編集権限がありません。」ならメッセージを追記する。    
    objReg.Pattern = "フィールド「.*」の編集権限がありません。"
    if objReg.Test(errMsg) then response=response & "編集権限のないフィールドのセルの値を削除してください。"    
  end if  
  KNTN_GetErrorMessage = response
end function

'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★


'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★json関連★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
' -----------------------------------------------------------------------
' Sub / Function
' jsonを扱う便利な関数の一覧を作成する
' -----------------------------------------------------------------------
Sub KNTN_InitScriptEngine()
    Set kntn_ScriptEngine = CreateObject("ScriptControl")
    kntn_ScriptEngine.Language = "JScript"
    'stringをjsonに変換する関数
    kntn_ScriptEngine.AddCode "function Parse(str) { return eval('(' + str + ')'); };"
    'keyに対応するvalueを返却する関数
    kntn_ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    'keyの一覧を取得する関数（配列として返す）
    kntn_ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
    'keyの数を取得する関数
    kntn_ScriptEngine.AddCode "function getKeysCount(jsonObj) { var count = 0; for (var i in jsonObj) { count++; } return count; } "
    'インデックス指定でkeyを取得する関数
    kntn_ScriptEngine.AddCode "function getKeyByIndex(jsonObj, index) { var keys = getKeys(jsonObj); return keys[index]; } "
    'keyが存在するかをチェックする関数
    kntn_ScriptEngine.AddCode "function checkKey(jsonObj,key) { return key in jsonObj; } "
    '配列の長さを取得する関数
    kntn_ScriptEngine.AddCode "function getArrayLength(arr) { return arr.length; } "
    'インデックス指定で配列要素を取得する関数
    kntn_ScriptEngine.AddCode "function getArrayItem(arr, index) { return arr[index]; } "
End Sub

'=====================================================
'APIに渡すjson文字列のvalueに複数値を設定する関数
'=====================================================
function KNTN_SplitKugirimojiForJson(target,kugirimoji)
  dim aryStrings,after
  after = ""
  if kugirimoji = vbcrlf then
    'いったんvblfにしてからvbcrlfに統一
    target = Replace(target,vbcrlf,vblf)
    target = Replace(target,vblf,vbcrlf)
  end if
  aryStrings = Split(target,kugirimoji)

  dim i 
  for i = 0 to ubound(aryStrings)
    if after ="" then
      'ダブルクオーテーションで囲う
      after = """" & aryStrings(i) & """"
    else    
      'ダブルクオーテーションで囲い、カンマ区切りでつなげる
      after = after & ",""" & aryStrings(i) & """"
    end if
  next
  KNTN_SplitKugirimojiForJson =  after
end  function


'=====================================================
'APIに渡すjson文字列のvalueのobjectに複数値を設定する関数
'=====================================================
function KNTN_SplitKugirimojiForJsonObject(target,kugirimoji)
  dim aryStrings,after
  after = ""
  if kugirimoji = vbcrlf then
    'いったんvblfにしてからvbcrlfに統一
    target = Replace(target,vbcrlf,vblf)
    target = Replace(target,vblf,vbcrlf)
  end if
  aryStrings = Split(target,kugirimoji)
 
  dim i 
  for i = 0 to ubound(aryStrings)
    if after ="" then
      'ダブルクオーテーションで囲う
      after = "{""code"":""" & aryStrings(i) & """}"
    else    
      'ダブルクオーテーションで囲う
      after = after & ",{""code"":""" & aryStrings(i) & """}"
    end if
  next
  KNTN_SplitKugirimojiForJsonObject =  after
end  function


'=====================================================
'フィールド名またはフィールドコードに該当するフィールド情報のjsonを取得
'=====================================================
function KNTN_getFieldJson(field,nameOrCode,json_properties)
  dim json_field
  'フィールド名指定の場合は繰り返し取得する
  if nameOrCode = "フィールド名" then
    Dim key, i, keysCount
    'keyの数を取得する
    keysCount = kntn_ScriptEngine.Run("getKeysCount", json_properties)
    For i = 0 To keysCount - 1
      'インデックス指定でkeyを取得する
      key = kntn_ScriptEngine.Run("getKeyByIndex", json_properties, i)

      'それぞれのkeyに該当するフィールドを取得する
      set json_field = kntn_ScriptEngine.Run("getProperty", json_properties,key)

      'json_fieldがnullでないことを確認してからlabelプロパティにアクセス
      if Not (json_field Is Nothing) then
        'フィールド名を取得し、一致していたら処理を抜ける
        if json_field.label = field then
          set KNTN_getFieldJson = json_field
          exit for
        end if
      end if
    Next
  else
    'フィールドコードの場合はjsonを取得する
    if kntn_ScriptEngine.Run("checkKey", json_properties, field) then
      set json_field = kntn_ScriptEngine.Run("getProperty", json_properties,field)
      set KNTN_getFieldJson = json_field
    else
      set KNTN_getFieldJson = Nothing
    end if
  end if
end function

'=====================================================
'yyyy-mm-ddTHH:mm:ssZの型に変換
'=====================================================
function KNTN_FormatDateTime(datetime)
	Dim objReg
	Set objReg = CreateObject("VBScript.RegExp")
	' 正規表現の指定
	With objReg
			.Pattern = "\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z"
	End With

	blnRtn = objReg.Test(datetime)
  '元から形式が正しい場合はそのまま返す
	if blnRtn then
    KNTN_FormatDateTime = datetime
  '日付形式でない場合はそのまま返す（APIでエラー検知）
  elseif isdate(datetime) = false then
    KNTN_FormatDateTime = datetime
	else
		' フォーマットされた日付と時刻を作成
		KNTN_FormatDateTime = Year(datetime) & "-" & Right("0" & Month(datetime), 2) & "-" & Right("0" & Day(datetime), 2) _
      & "T" &  Right("0" & Hour(datetime), 2) & ":" & Right("0" & Minute(datetime), 2) & ":" & Right("0" & Second(datetime), 2) &"Z" 
	end if

end function

'=====================================================
'yyyy-mm-ddに変換
'=====================================================
function KNTN_FormatDate(date)
  '日付形式でない場合はそのまま返す（APIでエラー検知）
  if isdate(date) = false then
    KNTN_FormatDate = date
  else
    KNTN_FormatDate = Year(date) & "-" & Right("0" & Month(date), 2) & "-" & Right("0" & Day(date), 2) 
  end if
end function


'=====================================================
'現在時刻をyyyy/mm/dd HH:mm:ssの形で取得する
'=====================================================
function KNTN_GetNow()
  dateNow = now
  KNTN_GetNow = Year(datenow) & "/" & Right("0" & Month(datenow), 2) & "/" & Right("0" & Day(datenow), 2) _ 
                    & " " & Right("0" & Hour(datenow), 2) & ":" & Right("0" & Minute(datenow), 2) & ":" & Right("0" & Second(datenow), 2)
end function


' -----------------------------------------------------------------------
' Sub / Function
' KintoneのTypeに応じてAPIに渡すjson文字列のvalueを作成する。
' -----------------------------------------------------------------------
function KNTN_CreateJsonKeyAndValue(fieldCode,fieldType,fieldvalue,brank,kugirimoji)
  dim result
  result =""
  'もしブランクなら初期値を設定するため、項目自体が不要
  'ラジオボタンはブランクは設定できないため、初期値を設定
  if (fieldvalue  = "" and brank <> "") or (fieldType = "RADIO_BUTTON" and fieldvalue = brank) then
    KNTN_CreateJsonKeyAndValue = ""
    exit function
  end if

  'jsonの値なのでエスケープ処理
  if fieldvalue <> brank then
    fieldvalue = Replace(fieldvalue,"\","\\")
    fieldvalue = Replace(fieldvalue,"""","\""")
  end if

  'blank（空白またはNull）はTypeに応じて異なる。
  select case fieldType
  case "RICH_TEXT","RADIO_BUTTON","DROP_DOWN","SINGLE_LINE_TEXT","MULTI_LINE_TEXT","NUMBER","LINK","SINGLE_LINE_TEXT","NUMBER"
    if fieldvalue = brank then
      result =  """" & fieldcode & """:{""value"":""""}" 
    else
      result =  """" & fieldcode & """:{""value"":""" & fieldvalue & """}" 
    end if

  
  case "CHECK_BOX","MULTI_SELECT"
    if fieldvalue = brank then
      result =  """" & fieldcode & """:{""value"":[]}" 
    else
      result =  """" & fieldcode & """:{""value"":[" & KNTN_SplitKugirimojiForJson(fieldValue,kugirimoji) & "]}" 
    end if

  case "USER_SELECT","ORGANIZATION_SELECT","GROUP_SELECT"
    if fieldvalue = brank then
      result =  """" & fieldcode & """:{""value"":[]}" 
    else
      result =  """" & fieldcode & """:{""value"":[" & KNTN_SplitKugirimojiForJsonObject(fieldvalue,kugirimoji) & "]}" 
    end if

  case "DATETIME"
    if fieldvalue = brank then
      result =  """" & fieldcode & """:{""value"":""""}" 
    else
      result =  """" & fieldcode & """:{""value"":""" & KNTN_FormatDateTime(fieldValue) & """}" 
    end if

  case "DATE"
    if fieldvalue = brank then
      result =  """" & fieldcode & """:{""value"":null}" 
    else
      result =  """" & fieldcode & """:{""value"":""" & KNTN_FormatDate(fieldValue) & """}" 
    end if

  case "TIME"
    if fieldvalue = brank then
      result =  """" & fieldcode & """:{""value"":null}" 
    else
      result =  """" & fieldcode & """:{""value"":""" & fieldvalue & """}" 
    end if
  end select
  KNTN_CreateJsonKeyAndValue = result
end function


'=====================================================
'複数ファイルの設定値を区切り文字で区切り、配列化する
'=====================================================
function Kntn_SplitFiles(target,kugirimoji)
  if kugirimoji = vbcrlf then
    'いったんvblfにしてからvbcrlfに統一
    target = Replace(target,vbcrlf,vblf)
    target = Replace(target,vblf,vbcrlf)
  end if
  Kntn_SplitFiles = Split(target,kugirimoji)
end  function
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★



'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★ヘッダー・フィールド情報関連★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'=====================================================
'dictionaryにヘッダーのフィールドのフィールドコードとフィールド名を配列で取得する
'key：配列名（メイン情報やテーブル名（またはテーブルのコード）などシート名が入る）
'value：2次元配列。0行目にフィールドコード、1行目にフィールド名、2行目にフィールドタイプ
'=====================================================
sub KNTN_getHeaderArray(json_properties,outputType,nameOrCode,rootArray,arrayName,lookUpArray,flgTable)
  dim json_field,json_lookup

  'property項目のkey一覧を取得する
  Dim key, i, keysCount
  keysCount = kntn_ScriptEngine.Run("getKeysCount", json_properties)
  

  dim tmpArray()
  dim colNum 
  'テーブルではない場合は、rowidを設定しない
  if flgTable = False then
    '登録ではレコードIDを利用しない
    if outputType="レコード登録用フィールド取得（ヘッダーのみ）" then
      redim tmpArray(2,keysCount)
      colNum =0
    else
      redim tmpArray(2,keysCount+1)
      tmpArray(0,0) = "$id"
      tmpArray(1,0) = "レコードID" 
      colNum =1
    end if
  else
    'テーブルの登録では行IDを利用しない
    if outputType="レコード登録用フィールド取得（ヘッダーのみ）" then
      redim tmpArray(2,keysCount+1)
      tmpArray(0,0) = "$id"
      tmpArray(1,0) = "レコードID"
      colNum =1 
    else 
      redim tmpArray(2,keysCount+2)
      tmpArray(0,0) = "$id"
      tmpArray(1,0) = "レコードID"
      tmpArray(0,1) = "rowId"
      tmpArray(1,1) = "テーブル行ID"
      colNum =2
    end if
  end if

  For i = 0 To keysCount - 1
    'インデックス指定でkeyを取得する
    key = kntn_ScriptEngine.Run("getKeyByIndex", json_properties, i)
    'フィールド名・フィールドコード・フィールドタイプを取得する。
    set json_field =kntn_ScriptEngine.Run("getProperty", json_properties,key)
    fieldtype = json_field.type 
    fieldname = json_field.label
    fieldcode = json_field.code
    
    select case outputType 
    case "全フィールド取得","カスタム"
      'skip対象のフィールドはfieldcodeをブランクにする
      select case fieldtype
'      case "__ID__","REFERENCE_TABLE","LABEL","SPACER","HR","GROUP"
      'カテゴリー、ステータス、作業者は取得できるとドキュメントに記載があるが、実際には取得できない（理由は不明）
      case "__ID__","REFERENCE_TABLE","LABEL","SPACER","HR","GROUP","CATEGORY","STATUS","STATUS_ASSIGNEE"
        fieldcode = ""
      end select

    case "レコード（更新・削除用）フィールド取得"
      select case fieldtype
      case "RECORD_NUMBER","__ID__","__REVISION__","CREATOR","CREATED_TIME","MODIFIER","UPDATED_TIME","CATEGORY","STATUS","STATUS_ASSIGNEE","REFERENCE_TABLE","LABEL","SPACER","HR","GROUP","CALC"
        fieldcode = ""
      case else
        'もしLookUpのMappingフィールドなら取得しない
        if kntn_checkLookUpMappingField(fieldcode,lookUpArray) then
          fieldcode = ""
        end if
      end select

    case "レコード登録用フィールド取得（ヘッダーのみ）"
      select case fieldtype
      case "RECORD_NUMBER","__ID__","__REVISION__","CREATOR","CREATED_TIME","MODIFIER","UPDATED_TIME","CATEGORY","STATUS","STATUS_ASSIGNEE","REFERENCE_TABLE","LABEL","SPACER","HR","GROUP","CALC"
        fieldcode = ""
      case else
        'もしLookUpのMappingフィールドなら取得しない
        if kntn_checkLookUpMappingField(fieldcode,lookUpArray) then
          fieldcode = ""
        end if
      end select

    end select
  
    'サブテーブルは別の配列に格納するため、再帰的に呼び出す
    if fieldType = "SUBTABLE" then
        set table_field = json_field.fields
        if nameOrCode = "フィールド名" then
          table_name = fieldname
        else
          table_name = fieldcode
        end if
        'サブテーブルはrootArrayにkeyをテーブル名にして格納するため、再帰的に関数を呼び出す。
        call KNTN_getHeaderArray(table_field,outputType,nameOrCode,rootArray,table_name,lookUpArray,True)
    
    elseif fieldcode <> "" then
      tmpArray(0,colNum) = fieldcode
      tmpArray(1,colNum ) = fieldName
      tmpArray(2,colNum) = fieldType
      colNum =colNum + 1
    end if
  next

  for i = 0 to ubound( tmpArray,2)
    fieldcode =  tmpArray(0,i)
    'もし不要な列が入っていたら削除する
    if fieldCode = "" then 
      redim preserve tmpArray(ubound( tmpArray,1),i-1) 
      exit for
    end if
  next

  rootArray.Add arrayname,tmpArray
end sub


' -----------------------------------------------------------------------
' Sub / Function
' ヘッダーのフィールド名やフィールドコードが複数存在するか確認する
' -----------------------------------------------------------------------
sub KNTN_checkDuplicateFields(array_header,rowNum)
  dim i ,j
  dim blnDuplicateFields
  blnDuplicateFields=False

  for i = lbound(array_header,2) to ubound(array_header,2)
    field = array_header(rowNum,i)
    if field ="" then exit for
    for j = i+1 to ubound(array_header,2)
      'もし同じコードが存在するならエラー
      if array_header(rowNum,j) = field then
        blnDuplicateFields=True
      end if
    next
  next

  if blnDuplicateFields then
    err.raise 1,"","同名のフィールド（" & field & "）が複数存在しています。"
  end if
end sub

'=====================================================
'LookUpのMappingFieldに該当するかを判別する
'=====================================================
function KNTN_checkLookUpMappingField(fieldcode,lookUpArray)
  dim i,result
  result = false
  for i = 1 to ubound(lookUpArray)
    if fieldcode = lookUpArray(i) then
      result = True
      exit for
    end if
  next 
  KNTN_checkLookUpMappingField = result
end function


'=====================================================
'LookUpのMappingFieldsになっているFieldの一覧を配列に格納
'=====================================================
function KNTN_CreateLookUpArray(json_properties)
  dim json_field,json_lookup

  Dim key, i, keysCount, mappingField
  keysCount = kntn_ScriptEngine.Run("getKeysCount", json_properties)
  
  'ルックアップの一覧を初めに取得する
  dim lookUpArray()
  redim lookUpArray(0)
  For i = 0 To keysCount - 1
    'インデックス指定でkeyを取得する
    key = kntn_ScriptEngine.Run("getKeyByIndex", json_properties, i)
    set json_field =kntn_ScriptEngine.Run("getProperty", json_properties,key)
    'lookupキーが存在するかで判断。タイプでは判断できない。
    if kntn_ScriptEngine.Run("checkKey", json_field,"lookup") then
      set json_lookup =  json_field.lookup
      Dim mappingFieldsCount, j
      mappingFieldsCount = kntn_ScriptEngine.Run("getArrayLength", json_lookup.fieldMappings)
      For j = 0 To mappingFieldsCount - 1
        Set mappingField = kntn_ScriptEngine.Run("getArrayItem", json_lookup.fieldMappings, j)
        redim preserve lookUpArray(ubound(lookUpArray)+1)
        lookUpArray(ubound(lookUpArray)) = mappingField.field
      next
    end if

		'テーブル内にもルックアップが存在する場合がある
		if json_field.type ="SUBTABLE" then
      set json_field = json_field.fields
			Dim tableKey, tableKeysCount, k
			tableKeysCount = kntn_ScriptEngine.Run("getKeysCount", json_field)
			For k = 0 To tableKeysCount - 1
				'インデックス指定でtableKeyを取得する
				tableKey = kntn_ScriptEngine.Run("getKeyByIndex", json_field, k)
				set table_field =kntn_ScriptEngine.Run("getProperty", json_field,tableKey)
				'lookupキーが存在するかで判断。タイプでは判断できない。
				if kntn_ScriptEngine.Run("checkKey", table_field,"lookup") then
					set json_lookup =  table_field.lookup
					Dim tableMappingFieldsCount, l
					tableMappingFieldsCount = kntn_ScriptEngine.Run("getArrayLength", json_lookup.fieldMappings)
					For l = 0 To tableMappingFieldsCount - 1
						Set mappingField = kntn_ScriptEngine.Run("getArrayItem", json_lookup.fieldMappings, l)
						redim preserve lookUpArray(ubound(lookUpArray)+1)
						lookUpArray(ubound(lookUpArray)) = mappingField.field
					next
				end if
			next
		end if
  next

  KNTN_CreateLookUpArray = lookUpArray
end function

'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★


'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★API利用★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'=====================================================
'APIを利用し、フィールドの情報を取得する
'=====================================================
function KNTN_getFieldsInfo(kntn_app_id,kntn_guestspace_id)
  Dim kntn_api_uri
  Dim responseText
  Dim sendData

  'Kintone API のエンドポイント
	if kntn_guestspace_id = "" then
    kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/v1/app/form/fields.json?app=" & kntn_app_id
	else
	  kntn_api_uri = "https://" & kntn_subdomain & ".cybozu.com/k/guest/" & kntn_guestspace_id & "/v1/app/form/fields.json?app=" & kntn_app_id
	end if

  'アクセストークンの有効性を確認
  call KNTN_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

  ' API リクエストのヘッダーとデータを設定
  With wscript.CreateObject("MSXML2.XMLHTTP")
    .Open "Get", kntn_api_uri, False
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
         KNTN_getFieldsInfo = responseText

      Case Else
        Err.Raise 1,"","Kintoneのフィールド情報取得操作に失敗しました。" & vbCrLf & _
					"ステータスコード：" & statusCode  & vbCrLf & _
					"レスポンス: " & KNTN_GetErrorMessage(responseText)
    End Select
  End With
End function


' -----------------------------------------------------------------------
' Sub / Function
' Kintoneの一時保管領域へのファイルアップロード
' -----------------------------------------------------------------------
function Kntn_tmpUpload(file_Path)
  'ファイルパスからファイル名のみ抽出
  dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  fileName = fso.GetFileName(file_Path)

  'ファイルサイズが1GBを超えているものは添付不可
  Set file = fso.GetFile(file_Path)
  fileSizeInBytes = file.Size
  fileSizeInGB = fileSizeInBytes / (1024 ^ 3) ' バイトをGBに変換
  if fileSizeInGB > 1 then err.raise 1,"","1GBを超えたサイズの添付ファイルはアップロードできません。"

	'URLEncodeしないと文字化けする
  Dim sc, js
	Set sc = CreateObject("ScriptControl")
	sc.Language = "JScript"
	Set js = sc.CodeObject
	URLEncode = js.encodeURIComponent(fileName)
  URLEncode = Replace(URLEncode, "(", "%28")
  URLEncode = Replace(URLEncode, ")", "%29")

  strBoundary = "boundary"

  adTypeBinary = 1
  adTypeText = 2

  Set fso = CreateObject("Scripting.FileSystemObject")
  fsize = fso.GetFile(File_Path).Size

  ' アップロードファイルをバイナリ形式で読込
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = adTypeBinary
  stream.Open
  stream.LoadFromFile File_Path
  fileContents = stream.Read
  stream.Close

  stream.Type = adTypeText
  stream.Charset = "shift-jis"
  stream.Open

  ' バイナリデータの前まで
  Kntn_ChangeStreamType stream, adTypeText
  params = ""
  params = params & "--" & strboundary & vbCrLf
  '日本語のファイル名で文字化けしたため、filename*=UTF-8''" & URLEncode & "" を追加
  params = params &  "Content-Disposition: form-data;  name=""file"";filename="""& FILEname &"""; filename*=UTF-8''" & URLEncode & "" & vbCrLf
  params = params &  "Content-Type: application/octet-stream" & vbCrLf  
  
  params = params & vbCrLf
  stream.WriteText params

  ' バイナリデータ
  Kntn_ChangeStreamType stream, adTypeBinary
  
  '0バイトの場合は本文を書きこまない
  if fsize <> 0 then
    stream.Write fileContents
  end if

  ' 最後
  Kntn_ChangeStreamType stream, adTypeText
  stream.WriteText vbCrLf & "--" & strboundary & "--" & vbCrLf

  Kntn_ChangeStreamType stream, adTypeBinary
  stream.Position = 0
  formData = stream.Read
  stream.Close
  strPost = "https://" & kntn_subdomain & ".cybozu.com/k/v1/file.json"

  'アクセストークンの有効性を確認
  call Kntn_CheckAccessTokenValidity(kntn_access_token, kntn_expires_date, kntn_refresh_token) 

  Dim oHTTP : Set oHTTP = wscript.CreateObject("MSXML2.XMLHTTP")
  oHTTP.Open "POST", strPost, False
  oHTTP.setRequestHeader "Authorization", "Bearer " & kntn_access_token
  oHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary="&strBoundary
  oHTTP.setRequestHeader "User-Agent", kntn_userAgent
  oHTTP.send(formData)

  ' レスポンステキストを取得
  responseText = oHTTP.responseText
  statusCode = oHTTP.status

  '成功したら、ファイルキーをレスポンスから取得する
  if oHTTP.status = 200 then
    dim json
    set json = kntn_ScriptEngine.CodeObject.Parse(responseText)
    Kntn_tmpUpload = json.fileKey
  else 
    Err.Raise 1, "", _
      "Kintoneのファイルアップロード操作に失敗しました。" & vbCrLf & _
      "ステータスコード：" & statusCode  & vbCrLf & _
      "レスポンス: " & Kntn_GetErrorMessage(responseText)
  end if

  Set oHTTP = Nothing
end function
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★



'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★取得レコード整形★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
' -----------------------------------------------------------------------
' Sub / Function
' Kintoneのレコード取得時にTypeに応じてvalueを取得する。
' -----------------------------------------------------------------------
function KNTN_getFieldValue(fieldtype,fieldcode,record,kugirimoji)
  dim values
  values = ""

  '閲覧不可のフィールドの場合、取得できないため、空文字を返す
  if kntn_ScriptEngine.Run("checkKey", record,fieldcode) = false then
    KNTN_getFieldValue = values
    exit function
  end if

  select case fieldtype
  case "RECORD_NUMBER","__ID__","__REVISION__","CREATED_TIME","UPDATED_TIME","SINGLE_LINE_TEXT","MULTI_LINE_TEXT","RICH_TEXT","NUMBER","CALC","RADIO_BUTTON","DROP_DOWN","DATE","TIME","DATETIME","LINK","STATUS"
    'NULLの場合はブランクを入れる
    if isNull(kntn_ScriptEngine.Run("getProperty", record,fieldcode)) then
      values =""
    else
      values =kntn_ScriptEngine.Run("getProperty", record,fieldcode).value
    end if

  case "CREATOR","MODIFIER"
    set obj =kntn_ScriptEngine.Run("getProperty", record,fieldcode).value
    values = obj.code

  case "CHECK_BOX","MULTI_SELECT","カテゴリー"
    set obj =kntn_ScriptEngine.Run("getProperty", record,fieldcode).value
    Dim objCount, m
    objCount = kntn_ScriptEngine.Run("getArrayLength", obj)
    For m = 0 To objCount - 1
      Dim tmpValue
      tmpValue = kntn_ScriptEngine.Run("getArrayItem", obj, m)
      if values = "" then      
        values = tmpValue
      else 
        values = values & kugirimoji & tmpValue
      end if
    next

  case "USER_SELECT","ORGANIZATION_SELECT","GROUP_SELECT","STATUS_ASSIGNEE"
    set obj =kntn_ScriptEngine.Run("getProperty", record,fieldcode).value
    Dim objCount2, n
    objCount2 = kntn_ScriptEngine.Run("getArrayLength", obj)
    For n = 0 To objCount2 - 1
      Dim codeObj
      Set codeObj = kntn_ScriptEngine.Run("getArrayItem", obj, n)
      if values = "" then      
        values = codeObj.code
      else 
        values = values & kugirimoji & codeObj.code
      end if
    next

  'TYPEがFILEとして送られてきた場合は更新用のもののため、何も設定しない。
  case "FILE"
    values =""

  case "FILE_NAME"
    set obj =kntn_ScriptEngine.Run("getProperty", record,fieldcode).value
    Dim objCount3, p, fileObj
    objCount3 = kntn_ScriptEngine.Run("getArrayLength", obj)
    For p = 0 To objCount3 - 1
      Set fileObj = kntn_ScriptEngine.Run("getArrayItem", obj, p)
      if values = "" then      
        values = fileObj.name
      else 
        values = values & kugirimoji & fileObj.name
      end if
    next

  case "FILE_KEY"
    set obj =kntn_ScriptEngine.Run("getProperty", record,fieldcode).value
    Dim objCount4, q
    objCount4 = kntn_ScriptEngine.Run("getArrayLength", obj)
    For q = 0 To objCount4 - 1
      Set fileObj = kntn_ScriptEngine.Run("getArrayItem", obj, q)
      if values = "" then      
        values = fileObj.fileKey
      else 
        values = values & kugirimoji & fileObj.fileKey
      end if
    next

  case "__ID__"
    values = kntn_ScriptEngine.Run("getProperty", record,"$id").value
  end select
  if isNull(values) then values =""
  KNTN_getFieldValue = values
end function

'=====================================================
'レコードのメイン情報を二次元配列で取得する
'=====================================================
function Kntn_GetMainRecord(array_header,array_fieldsinfo,json_records,outputType,kugirimoji)
  dim array_records()
  redim array_records(499,ubound(array_header,2))

  recordRow=0
  i = 0
  
  Dim recordsCount, recordIndex
  recordsCount = kntn_ScriptEngine.Run("getArrayLength", json_records)
  For recordIndex = 0 To recordsCount - 1
    Dim record
    Set record = kntn_ScriptEngine.Run("getArrayItem", json_records, recordIndex)
    colNum = 0
    for i = 0 to ubound(array_fieldsinfo,2)
      fieldcode = array_fieldsinfo(0,i)
      fieldName = array_fieldsinfo(1,i)
      fieldtype = array_fieldsinfo(2,i)
      if fieldcode ="" then exit for
      if fieldType ="FILE" then
        array_records(recordRow,colNum)=  Kntn_getFieldValue("FILE_NAME",fieldcode,record,kugirimoji)
        colnum = colnum + 1
        array_records(recordRow,colNum) =  Kntn_getFieldValue("FILE_KEY",fieldcode,record,kugirimoji)
        colnum = colnum + 1
        array_records(recordRow,colNum) =  ""
      elseif fieldcode ="$id" then
        array_records(recordRow,colNum) = kntn_ScriptEngine.Run("getProperty", record,"$id").value
      else
        array_records(recordRow,colNum) = Kntn_getFieldValue(fieldtype,fieldcode,record,kugirimoji)
      end if
			colNum = colNum+1
    next
    recordRow=recordRow+1
  next
  Kntn_GetMainRecord = array_records
end function


'=====================================================
'レコードのテーブル情報を二次元配列で取得する
'=====================================================
function Kntn_GetTableRecord(tableCode,array_header,array_fieldsinfo,json_records,outputType,kugirimoji)
  dim array_Table()
  redim array_table(ubound(array_header,2),0)
  tableRecordCount = 0 

  '各レコードの情報を繰り返す
  Dim recordsCount2, recordIndex2, tableRecord
  recordsCount2 = kntn_ScriptEngine.Run("getArrayLength", json_records)
  For recordIndex2 = 0 To recordsCount2 - 1
    Set tableRecord = kntn_ScriptEngine.Run("getArrayItem", json_records, recordIndex2)
		'もしテーブルの閲覧権限がなければデータは取得しない。
		if kntn_ScriptEngine.Run("checkKey", tableRecord,tableCode) = false then
			exit for
		end if
    'recordからtable情報を取得する                
    set json_Table = kntn_ScriptEngine.Run("getProperty", tableRecord,tableCode)
    set json_Table = json_Table.value
    'テーブル内の各行を繰り返す
    Dim tableRowsCount, tableRowIndex
    tableRowsCount = kntn_ScriptEngine.Run("getArrayLength", json_Table)
    For tableRowIndex = 0 To tableRowsCount - 1
      Dim json_row
      Set json_row = kntn_ScriptEngine.Run("getArrayItem", json_Table, tableRowIndex)
      rowId = json_row.id
      set rowValues = json_row.Value
      colnum = 2

      'テーブルの行が増えるたびに配列のサイズを調整する
      redim preserve array_table(ubound(array_header,2),tablerecordCount)
      array_table(0,tablerecordCount) = kntn_ScriptEngine.Run("getProperty", tableRecord,"$id").value
      array_table(1,tablerecordCount) = rowId
      
      for i = 2 to ubound(array_fieldsinfo,2)
        fieldType = array_fieldsinfo(2,i)
        fieldcode = array_fieldsinfo(0,i)
        if fieldcode ="" then exit for
        if fieldType = "FILE"  then
          array_table(Colnum,tablerecordCount) =  Kntn_getFieldValue("FILE_NAME",fieldcode,rowValues,kugirimoji)
          colnum = colnum + 1
          array_table(Colnum,tablerecordCount) =  Kntn_getFieldValue("FILE_KEY",fieldcode,rowValues,kugirimoji)
          colnum = colnum + 1
          array_table(colNum,tablerecordCount) =  ""
        else
          array_table(colNum,tablerecordCount) =  Kntn_getFieldValue(fieldtype,fieldcode,rowValues,kugirimoji)
        end if
        colnum = colnum + 1
      next
      tablerecordCount = tablerecordCount + 1
    next
  next
  dim array_transposeTable()
  '行列を入れ替えたテーブルを作成する                
  redim array_transposeTable(ubound(array_table,2),ubound(array_table,1))
  for r =  0 to ubound(array_table,1)
      for c = 0 to ubound(array_table,2)
        array_transposeTable(c,r) = array_Table(r,c)
      Next
  next
  Kntn_GetTableRecord = array_transposeTable
end function
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★



'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★QUERY関連★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'=====================================================
'フィールド条件のqueryを作成
'=====================================================
function Kntn_CreateQuery(field,condition,value,kugirimoji,nameOrCode,json)
  dim fieldname,fieldcode,fieldtype,json_field
  if field = "レコードID" or field = "$id" then
    fieldname = "レコードID"
    fieldcode = "$id"
    fieldtype = "__ID__"

  else
    on error resume next
    set json_field = Kntn_getFieldJson(field,nameOrCode,json)
    fieldname = json_field.label
    fieldcode = json_field.code
    fieldtype = json_field.type
    if err.Number <> 0 then
      on error goto 0
      err.raise 1,"","クエリに指定したフィールド「" & field & "」が正しくありません。"
    end if
    on error goto 0
  end if

  dim tmpQuery
	if nameOrCode = "フィールド名" then 
		'エラー時に表示するフィールドをフィールド名かフィールドコードどちらかを設定
		errField = fieldName 
	else
		errField = fieldCode
	end if

  '\と"でエスケープ処理が必要なものはエスケープする
  select case fieldtype
  case "SINGLE_LINE_TEXT","MULTI_LINE_TEXT","RICH_TEXT","CHECK_BOX","RADIO_BUTTON","DROP_DOWN","MULTI_SELECT","STATUS"
    value = Replace(value,"""","\\\""")
    value = Replace(value,"\","\\\\")
  case "CREATED_TIME","UPDATED_TIME","DATETIME"
    value = KNTN_FormatDateTime(value)
  case "DATE"
    value = KNTN_FormatDate(value)
  end select

  '条件で分岐
  select case condition
  case "等しい(=)"
    'inを遣えるものはinを利用する。=のみのものは=
    select case fieldtype
    case "RECORD_NUMBER","__ID__","CREATOR","MODIFIER","SINGLE_LINE_TEXT","LINK","NUMBER","CHECK_BOX","RADIO_BUTTON","DROP_DOWN","MULTI_SELECT","USER_SELECT","STATUS"
      value  = Kntn_SplitKugirimojiForQuery(value,kugirimoji)  
      tmpQuery = fieldcode & " in (" & value & ")"
    case "CREATED_TIME","UPDATED_TIME","DATE","DATETIME","TIME"
      tmpQuery = fieldcode & " = \""" & value & "\"""
    case else   
       Err.Raise 1,"",errfield & "の条件は「等しい(=)」の指定ができません。"
    end select    

  case "等しくない(<>)"
    select case fieldtype
    case "RECORD_NUMBER","__ID__","CREATOR","MODIFIER","SINGLE_LINE_TEXT","LINK","NUMBER","CHECK_BOX","RADIO_BUTTON","DROP_DOWN","MULTI_SELECT","USER_SELECT","STATUS"
      value  = Kntn_SplitKugirimojiForQuery(value,kugirimoji)  
      tmpQuery = fieldcode & " not in (" & value & ")"
    case "CREATED_TIME","UPDATED_TIME","DATE","DATETIME","TIME"
      tmpQuery = fieldcode & " != \""" & value & "\"""
    case else   
       Err.Raise 1,"",errfield & "の条件は「等しくない(<>)」の指定ができません。"
    end select    

  case "含む(like)"
    select case fieldtype
    case "SINGLE_LINE_TEXT","LINK","MULTI_LINE_TEXT","RICH_TEXT","FILE"
      tmpQuery = fieldcode & " like \""" & value & "\"""
    case else   
       Err.Raise 1,"",errfield & "の条件は「含む(like)」の指定ができません。"
    end select  

  case "含まない(not like)"
    select case fieldtype
    case "SINGLE_LINE_TEXT","LINK","MULTI_LINE_TEXT","RICH_TEXT","FILE"
      tmpQuery = fieldcode & " like \""" & value & "\"""
    case else   
       Err.Raise 1,"",errfield & "の条件は「含まない(not like)」の指定ができません。"
    end select 

  case "以上(≧)"
    select case fieldtype
    case "RECORD_NUMBER","__ID__","CREATED_TIME","UPDATED_TIME","NUMBER","DATE","DATETIME","TIME"
      tmpQuery = fieldcode & " >= \""" & value & "\"""
    case else   
       Err.Raise 1,"",errfield & "の条件は「以上(≧)」の指定ができません。"
    end select 

  case "以下(≦)"
    select case fieldtype
    case "RECORD_NUMBER","__ID__","CREATED_TIME","UPDATED_TIME","NUMBER","DATE","DATETIME","TIME"
      tmpQuery = fieldcode & " <= \""" & value & "\"""
    case else   
       Err.Raise 1,"",errfield & "の条件は「以下(≦)」の指定ができません。"
    end select 

  case "より大きい(>)"
    select case fieldtype
    case "RECORD_NUMBER","__ID__","CREATED_TIME","UPDATED_TIME","NUMBER","DATE","DATETIME","TIME"
      tmpQuery = fieldcode & " > \""" & value & "\"""
    case else   
       Err.Raise 1,"",errfield & "の条件は「より大きい(>)」の指定ができません。"
    end select 

  case "より小さい(<)"
    select case fieldtype
    case "RECORD_NUMBER","__ID__","CREATED_TIME","UPDATED_TIME","NUMBER","DATE","DATETIME","TIME"
      tmpQuery = fieldcode & " < \""" & value & "\"""
    case else   
       Err.Raise 1,"",errfield & "の条件は「より小さい(<)」の指定ができません。"
    end select 
  end select
  Kntn_CreateQuery = tmpQuery
end function


'=====================================================
'query用の分割関数　ダブルクオーテーションに\マークが必要となる。
'=====================================================
function Kntn_SplitKugirimojiForQuery(target,kugirimoji)
  dim aryStrings,after
  after = ""
  if kugirimoji = vbcrlf then
    'いったんvblfにしてからvbcrlfに統一
    target = Replace(target,vbcrlf,vblf)
    target = Replace(target,vblf,vbcrlf)
  end if
  aryStrings = Split(target,kugirimoji)

  dim i 
  for i = 0 to ubound(aryStrings)
    if after ="" then
      after = "\""" & aryStrings(i) & "\"""
    else    
      'ダブルクオーテーション区切りでつなげる
      after = after & ",\""" & aryStrings(i) & "\"""
    end if
  next
  Kntn_SplitKugirimojiForQuery =  after
end  function
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★



'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★CSV関連★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'=====================================================
'CSVを読込、配列として返す
'=====================================================
function KNTN_ReadCsv(csvPath, charcode,minIdx)
	dim arrTemp(),arrHeader(),arrTransport()
	redim arrHeader(0)

	Dim fs: Set fs = CreateObject("Scripting.FileSystemObject")
	'ADODBオブジェクト作成
	dim objAdo:Set objAdo = CreateObject("ADODB.Stream")

	'ストリームオブジェクトをオープン
	objAdo.Open
	objAdo.Type = 2               ' テキストファイル
	objAdo.Charset = charcode      ' 文字エンコーディング
	objAdo.LoadFromFile csvPath   ' ファイルをストリームに読み込む

	strItem = ""
	strLine = ""
	lngQuote = 0

	'ヘッダー行を初めに読み込む。
	Do While objAdo.EOS <> True
		strLine = strLine + objAdo.ReadText(-2) 
		dCnt = (Len(strLine) - Len(Replace(strLine,"""","")))
		If (dCnt = 0) Or (dCnt Mod 2 = 0) Then Exit Do
		'ダブルクォーテーションの数が偶数でない場合、改行含むデータとみなし次行を読み込む
		strLine = strLine + vbLf
	Loop

	For k = 1 To Len(strLine)
		char=Mid(strLine, k, 1)
		Select Case char
		Case "," '「"」が偶数なら区切り、奇数ならただの文字
			If lngQuote Mod 2 = 0 Then
				arrHeader(ubound(arrHeader)) = strItem
				redim preserve arrHeader(ubound(arrHeader)+1)
				strItem = ""
				lngQuote = 0
			Else
				strItem = strItem & char
			End If
		Case """" '「"」のカウントをとる
			lngQuote = lngQuote + 1
			strItem = strItem & char
		Case Else
			strItem = strItem & char
		End Select
	Next

	'末尾の項目は,がないため、最後に設定する
	arrHeader(ubound(arrHeader)) = strItem
	rColsNum=ubound(arrHeader)
	redim preserve arrTemp(rColsNum,0)

	for i = 0 to rColsNum
		arrTemp(i,0) = KNTN_trimCsvDoblequate(arrHeader(i))
	next

	'CSVの2行目以降を読み込む
	Do while objAdo.EOS <> True
		strLine=""
		'CSVの１行分（改行含む）を読み込む
		Do While objAdo.EOS <> True
			strLine = strLine + objAdo.ReadText(-2) 
			dCnt = (Len(strLine) - Len(Replace(strLine,"""","")))
			If (dCnt = 0) Or (dCnt Mod 2 = 0) Then Exit Do
			'ダブルクォーテーションの数が偶数でない場合、改行含むデータとみなし次行を読み込む
			strLine = strLine + vbLf
		Loop

		' CSVの１行分（改行含む）のカラム値の配列を作成
		rRows = Split(strLine, ",", -1, 1)
		rRowsNum = Ubound(rRows)

		If rColsNum <> rRowsNum Then
			' カンマ区切りしたデータ数がヘッダーカラム数と等しくない場合、rRowsを作り直す
			Dim tmpData: tmpData = ""
			i = 0
			For Each val in rRows
				' 1カラム分のデータを取得
				tmpData = tmpData & val
				dCnt = (Len(tmpData) - Len(Replace(tmpData,"""","")))
				If (dCnt > 0) And (dCnt Mod 2 <> 0) Then
					'ダブルクォーテーションの数が偶数でない場合、カンマを含むデータとみなし次のデータを読み込む
					tmpData = tmpData & ","
				Else
					'ダブルクォーテーションの数が偶数の場合、rRowsにデータを格納
					rRows(i) = tmpData
					i = i + 1
					tmpData = ""
				End If
			Next
			ReDim Preserve rRows(rColsNum)
		End If

    '空行のチェックを行う。
		for i = 0 to ubound(rRows)
      strval = rRows(i)
			if KNTN_trimCsvDoblequate(strval) <> "" then exit for
    next
    '上の繰り返しでexit forを通っていない場合は、全項目ブランクで空行と判定しLoopを抜ける
    if i > ubound(rRows) then
      exit do
    end if

		'入力行数を取得する。arrTempでは行列逆なのでわかりづらくなってしまっている。
		rowsCnt = ubound(arrTemp,2)+1
		redim preserve arrTemp(ubound(arrTemp,1),rowsCnt) 
		for i = 0 to ubound(rRows)
      strval = rRows(i)
			arrTemp(i,rowsCnt) =  KNTN_trimCsvDoblequate(strVal)
		next
	Loop

	'最後に配列の行列を逆さにする。同時に最小のidxを変更する（高速配列では最小のインデックスが1となるため）
	redim arrTransport(ubound(arrTemp,2)+minIdx,ubound(arrTemp,1)+minIdx)
	for i = 0 to ubound(arrTemp,1)
		for k = 0 to ubound(arrTemp,2)
			arrTransport(k+minIdx,i+minIdx) =arrTemp(i,k) 
		next
	next 
	KNTN_ReadCsv = arrTransport
end function

'=====================================================
'CSVの各項目の前後のダブルクオーテーションを削除したり、項目内のダブルクオーテーションの重複を単独にする
'=====================================================
function KNTN_trimCsvDoblequate(strValue)
    '入力引数がブランクなら何もしない
    if strValue ="" then 
      KNTN_trimCsvDoblequate = strValue
      exit function 
    end if

    '前後の「"」を削除
    If strValue = """" Then
        strValue = ""
    ElseIf Left(strValue, 1) = """" And Right(strValue, 1) = """" Then
        strValue = Mid(strValue, 2, Len(strValue) - 2)
        '「""」を「"」で置換
        strValue = Replace(strValue, """""", """")
    End If
		KNTN_trimCsvDoblequate = strValue
end function

'=====================================================
'配列をCSVとして保存する
'=====================================================
sub KNTN_SaveCsv(arr,csvPath,charcode,minIdx)
	strCsv =""
	for i = minIdx to ubound(arr,1)
		strLine =""
		for K = minIdx to ubound(arr,2)
      if isnull(arr(i,k)) then
        strVal = ""
      else
        strVal = arr(i,k)
      end if
			'ダブルクオーテーションを2つにする
			strVal = Replace(strVal,"""","""""")
			'各項目の前後にダブルクオーテーションを付ける
			strVal = """" & strVal & """"
			if strLine = "" then 
	 			strLine = strVal
			else
	 			strLine = strLine & "," & strVal
			end if
		Next
    '空行まで達したら抜ける
    if Replace(strLine,""""",","") = """""" then exit for
 		if strCsv = "" then 
			strCsv = strLine
		else
			strCsv = strCsv & vbcrlf & strLine
		end if
	next	
	
	'ADODBオブジェクト作成
	Set objAdo = CreateObject("ADODB.Stream")

	'ストリームオブジェクトをオープン
	objAdo.Open
	objAdo.Type = 2               ' テキストファイル
	objAdo.Charset = charcode      ' 文字コード
	'ストリームオブジェクトに書き込む
	objado.WriteText strCsv, 1
	'ストリームの内容をファイルに保存
	objado.SaveToFile csvPath, 2
	' クローズ
	objAdo.Close
end sub

'=====================================================
'配列を既存のCSVに追記する
'=====================================================
sub KNTN_AddCsv(arr,csvPath,charcode,minIdx,blnHeader)
  if blnHeader then
    diff = 1
  else
    diff = 0
  end if

	strCsv =""
  'ヘッダー行を抜く場合、minIdx+1からスタート
	for i = minIdx + diff to ubound(arr,1)

		strLine =""
		for K = minIdx to ubound(arr,2)
      strVal = arr(i,k)
			'ダブルクオーテーションを2つにする
			strVal = Replace(strVal,"""","""""")
			'各項目の前後にダブルクオーテーションを付ける
			strVal = """" & strVal & """"
			if strLine = "" then 
	 			strLine = strVal
			else
	 			strLine = strLine & "," & strVal
			end if
		Next

    '空行まで達したら抜ける
    if Replace(strLine,""""",","") = """""" then exit for

		if strCsv = "" then 
			strCsv = strLine
		else
			strCsv = strCsv & vbcrlf & strLine
		end if
	next	
	
	'ADODBオブジェクト作成
	Set objAdo = CreateObject("ADODB.Stream")

	'ストリームオブジェクトをオープン
	objAdo.Open
	objAdo.Type = 2               ' テキストファイル
	objAdo.Charset = charcode      ' 文字コード
  objAdo.LoadFromFile(csvPath)
  objAdo.Position = objAdo.size

	'ストリームオブジェクトに書き込む
	objado.WriteText strCsv, 1

	'ストリームの内容をファイルに保存
	objado.SaveToFile csvPath, 2

	' クローズ
	objAdo.Close
end sub

'=====================================================
'CSVファイル名の禁止文字を全角に変換する
'テーブル名で保存する際
'=====================================================
function KNTN_ReplaceForbiddenChar(strValue)
  '文字列置換
  strValue = Replace(strValue,"\","￥")
  strValue = Replace(strValue,":","：")
  strValue = Replace(strValue,"/","／")
  strValue = Replace(strValue,"*","＊")
  strValue = Replace(strValue,"?","？")
  strValue = Replace(strValue,"""","”")
  strValue = Replace(strValue,"<","＜")
  strValue = Replace(strValue,">","＞")
  strValue = Replace(strValue,"|","｜")
  KNTN_ReplaceForbiddenChar=strValue
end function
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★



'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★EXCEL関連★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'=====================================================
'2次元配列を指定のExcelファイルに入力。ファイルが存在しない場合は新規作成する。
'=====================================================
Sub KNTN_SetArrayToExcel(fname,sheetName,arraydata,range,flgClose)
  ' workbookオブジェクトを取得する
  SetUmsVariable "$CLEAR_ARGUMENT" , ""
  SetUMSVariable "$FILE_PATH_TYPE", "1"
  SetUMSVariable "$PARSE_FILE_PATH", fname
  filePath = GetUMSVariable("$PARSE_FILE_PATH")

  If filePath = "" Then
    SetUmsVariable "$CLEAR_ARGUMENT" , ""
    SetUMSVariable "$FILE_PATH_TYPE", "2"
    SetUMSVariable "$PARSE_FILE_PATH", fname
    filePath = GetUMSVariable("$PARSE_FILE_PATH")
  End If

  If filePath = "" Then
    Err.Raise 1, "", "指定されたファイルを開くことができません。" 
  End If

  'シート名に利用できない文字があったら削除する
  tmpname =""
  For i = 1 To Len(sheetName)
    Charstr = Mid(sheetName, i, 1)
    select case Charstr
    case ":","：","\","￥","?","？","[","［","]","］","/","／","*","＊"
    case else
      tmpname = tmpname & Charstr
    end select
  next
  sheetName = tmpname

  'シート名が32文字以上ある場合は31文字までにする
  if len(sheetName) > 31 then
    sheetName = left(sheetName,31)
  end if 

  ' workbookオブジェクトを取得する
  Set workbook = Nothing
  On Error Resume Next
  ' 既存のエクセルが起動されていれば警告を抑制する
  Set existingXlsApp = Nothing
  Set existingXlsApp = GetObject(, "Excel.Application")
  existingXlsApp.DisplayAlerts = False

  Set wash = CreateObject("WinActor7.ScriptHelper")
  For Each book in wash.GetExcelWorkbooks
    SetUMSVariable "$FILE_PATH_TYPE", 0
    SetUMSVariable "$PARSE_FILE_PATH", book.FullName
    bookPath = GetUMSVariable("$PARSE_FILE_PATH")
    If StrComp(bookPath, filePath, 1) = 0 Then
      Set workbook = book
      Set xlsApp = workbook.Parent
      xlsApp.Visible = True
      Exit For
    End If
  Next
  Set wash = Nothing

  ' Workbookが存在しない場合は、新たに開く。
  If workbook Is Nothing Then
    Set xlsApp = Nothing

    ' Excelが既に開かれていたならそれを再利用する
    If Not existingXlsApp Is Nothing Then
      Set xlsApp = existingXlsApp
      xlsApp.Visible = True
    Else
      Set xlsApp = CreateObject("Excel.Application")
      xlsApp.Visible = True
    End If

    '既存ファイルが存在する場合は利用する
    Set objFS = CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(filePath) = True Then
      Set workbook = xlsApp.Workbooks.Open(filePath)
    else
      '新規ワークシートを作成
      xlsApp.Workbooks.Add
      Set workbook = xlsApp.ActiveWorkbook
      workbook.Activesheet.name = sheetName
      workbook.SaveAs(filePath)
    end if

    xlsApp.DisplayAlerts = False
    xlsApp.DisplayAlerts = True
  End If

  ' 警告の抑制を元に戻す
  existingXlsApp.DisplayAlerts = True
  Set existingXlsApp = Nothing
  On Error Goto 0

  If workbook Is Nothing Then
    Err.Raise 1, "", "指定されたファイルを開くことができません。"
  End If

  workbook.Activate

  if sheetName <> "" then 
    '指定のシートが存在しなければ追加する
    on error resume next
    Set worksheet = workbook.Worksheets(sheetName)
    If Err.Number <> 0 Then
      Set worksheet = workbook.Sheets.Add
      worksheet.Name = sheetName
    end if
    On Error Goto 0
  else
    set worksheet=workbook.ActiveSheet
  end if

  ' ====ハイライトを表示する========================================================
  ' HwndプロパティはExcel2002以降のみ対応
  On Error Resume Next
    ShowUMSHighlight(xlsApp.Hwnd)
  On Error Goto 0

  'もしRangeが空ならA列の最終行+1を設定する。
  if range ="" then
    endrow = worksheet.Cells(worksheet.Rows.Count, 1).End(-4162).row
    range = "A" & endrow + 1
  end if
  

  ' ====配列の値をセルに書き込む==============================================================   
  if range = "A1" then
    worksheet.Cells.CLEAR
    worksheet.Cells.NumberFormatLocal = "@"
  end if

  '配列の最小インデックスが0の場合は調整する
  if lbound(arraydata,2) = 0 then
    diff = 1
  else 
    diff=0
  end if

  worksheet.Range(range).Resize(ubound(arraydata,1)+diff,ubound(arraydata,2)+diff).Value =arraydata

  'もし初めの行の最後の列が「エラー内容」なら列幅を調節する。
  if arraydata(lbound(arraydata,1),ubound(arraydata,2)) = "エラー内容" then
    worksheet.columns(ubound(arraydata,2)+1).ColumnWidth = 80
  end if  
  workbook.Save

  if flgClose = true then
    workbook.close True
    If xlsApp.Workbooks.Count = 0 Then
      xlsApp.Quit
    End If
  end if

  Set objRe = Nothing
  Set xlsApp = Nothing
  Set worksheet = Nothing
  Set workbook = Nothing

End Sub

'=====================================================
'APIで連携しないフィールドをグレーアウトする
'=====================================================
Sub KNTN_GrayOutExcel(fname,sheetName,array_header,array_outputtype,nameOrCode,flgClose)
  ' workbookオブジェクトを取得する
  SetUmsVariable "$CLEAR_ARGUMENT" , ""
  SetUMSVariable "$FILE_PATH_TYPE", "1"
  SetUMSVariable "$PARSE_FILE_PATH", fname
  filePath = GetUMSVariable("$PARSE_FILE_PATH")

  'シート名に利用できない文字があったら削除する
  tmpname =""
  For i = 1 To Len(sheetName)
    Charstr = Mid(sheetName, i, 1)
    select case Charstr
    case ":","：","\","￥","?","？","[","［","]","］","/","／","*","＊"
    case else
      tmpname = tmpname & Charstr
    end select
  next
  sheetName = tmpname

  'シート名が32文字以上ある場合は31文字までにする
  if len(sheetName) > 31 then
    sheetName = left(sheetName,31)
  end if 

  If filePath = "" Then
    SetUmsVariable "$CLEAR_ARGUMENT" , ""
    SetUMSVariable "$FILE_PATH_TYPE", "2"
    SetUMSVariable "$PARSE_FILE_PATH", fname
    filePath = GetUMSVariable("$PARSE_FILE_PATH")
  End If

  If filePath = "" Then
    Err.Raise 1, "", "指定されたファイルを開くことができません。" 
  End If


  ' workbookオブジェクトを取得する
  Set workbook = Nothing
  On Error Resume Next
  ' 既存のエクセルが起動されていれば警告を抑制する
  Set existingXlsApp = Nothing
  Set existingXlsApp = GetObject(, "Excel.Application")
  existingXlsApp.DisplayAlerts = False

  Set wash = CreateObject("WinActor7.ScriptHelper")
  For Each book in wash.GetExcelWorkbooks
    SetUMSVariable "$FILE_PATH_TYPE", 0
    SetUMSVariable "$PARSE_FILE_PATH", book.FullName
    bookPath = GetUMSVariable("$PARSE_FILE_PATH")
    If StrComp(bookPath, filePath, 1) = 0 Then
      Set workbook = book
      Set xlsApp = workbook.Parent
      xlsApp.Visible = True
      Exit For
    End If
  Next
  Set wash = Nothing

  ' Workbookが存在しない場合は、新たに開く。
  If workbook Is Nothing Then
    Set xlsApp = Nothing

    ' Excelが既に開かれていたならそれを再利用する
    If Not existingXlsApp Is Nothing Then
      Set xlsApp = existingXlsApp
      xlsApp.Visible = True
    Else
      Set xlsApp = CreateObject("Excel.Application")
      xlsApp.Visible = True
    End If

    '既存ファイルが存在する場合は利用する
    Set objFS = CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(filePath) = True Then
      Set workbook = xlsApp.Workbooks.Open(filePath)
    else
      '新規ワークシートを作成
      xlsApp.Workbooks.Add
      Set workbook = xlsApp.ActiveWorkbook
      workbook.Activesheet.name = sheetName
      workbook.SaveAs(filePath)
    end if

    xlsApp.DisplayAlerts = False
    xlsApp.DisplayAlerts = True
  End If

  ' 警告の抑制を元に戻す
  existingXlsApp.DisplayAlerts = True
  Set existingXlsApp = Nothing
  On Error Goto 0

  If workbook Is Nothing Then
    Err.Raise 1, "", "指定されたファイルを開くことができません。"
  End If

  workbook.Activate

  if sheetName <> "" then 
    Set worksheet = workbook.Worksheets(sheetName)
  else
    set worksheet=workbook.ActiveSheet
  end if

  ' ====ハイライトを表示する========================================================
  ' HwndプロパティはExcel2002以降のみ対応
  On Error Resume Next
    ShowUMSHighlight(xlsApp.Hwnd)
  On Error Goto 0

  'もしRangeが空ならA列の最終行+1を設定する。
  dim i,j,colNum
  colNum = 1

  if nameOrCode = "フィールド名" then 
    rownum =1 
  else 
    rownum =0
  end if
  for i = lbound(array_header,2) to ubound(array_header,2)
    fieldName =  array_header(0,i) 
    blnTargetCol = False
    'フィールドがAPI連携の利用列に存在するか確認する。
    for j = lbound(array_outputtype,2) to ubound(array_outputtype,2)
      if array_outputtype(rownum,j) = fieldName then
        blnTargetCol = true
        exit for
      end if 
    next
    'API連携で利用する列ではないため、グレーアウトする。
    if blnTargetCol = False then
        worksheet.columns(colnum).Interior.Color = RGB(128, 128, 128)
    end if
    colnum = colnum + 1
  next

  workbook.Save

  if flgClose = true then
    workbook.close True
    If xlsApp.Workbooks.Count = 0 Then
      xlsApp.Quit
    End If
  end if

  Set objRe = Nothing
  Set xlsApp = Nothing
  Set worksheet = Nothing
  Set workbook = Nothing

End Sub

'=====================================================
'Excelからデータを取得する
'=====================================================
Function KNTN_GetArraybyExcel(ExcelFilename,sheetName)
  'シート名に利用できない文字があったら削除する
  tmpname =""
  For i = 1 To Len(sheetName)
    Charstr = Mid(sheetName, i, 1)
    select case Charstr
    case ":","：","\","￥","?","？","[","［","]","］","/","／","*","＊"
    case else
      tmpname = tmpname & Charstr
    end select
  next
  sheetName = tmpname

  'シート名が32文字以上ある場合は31文字までにする
  if len(sheetName) > 31 then
    sheetName = left(sheetName,31)
  end if 

  ' ファイルのパスをフルパスに変換する
  Set fso = CreateObject("Scripting.FileSystemObject")
  filePath = fso.GetAbsolutePathName(ExcelFilename)

  ' workbookオブジェクトを取得する
  Set workbook = Nothing
  On Error Resume Next
  ' 既存のエクセルが起動されていれば警告を抑制する
  Set existingXlsApp = Nothing
  Set existingXlsApp = GetObject(, "Excel.Application")
  existingXlsApp.DisplayAlerts = False

  ' 一先ずWorkbookオブジェクトをGetObjectしてみる
  Set workbook = GetObject(filePath)
  Set xlsApp = workbook.Parent

  Set workbook = Nothing

  ' Workbookがまだ存在するか確認する
  For Each book In xlsApp.Workbooks
    If StrComp(book.FullName, filePath, 1) = 0 Then
      ' Workbookがまだ存在するので、このWorkbookは既に開かれていたもの
      Set workbook = book
    End If
  Next

  ' Workbookが存在しない場合は、新たに開く。
  If workbook Is Nothing Then
    Set xlsApp = Nothing

    ' Excelが既に開かれていたならそれを再利用する
    If Not existingXlsApp Is Nothing Then
      Set xlsApp = existingXlsApp
      xlsApp.Visible = True
    Else
      Set xlsApp = CreateObject("Excel.Application")
      xlsApp.Visible = True
    End If

    Set workbook = xlsApp.Workbooks.Open(filePath)
  End If

  ' 警告の抑制を元に戻す
  existingXlsApp.DisplayAlerts = True
  Set existingXlsApp = Nothing
  On Error Goto 0

  If workbook Is Nothing Then
    Err.Raise 1, "", "指定されたファイルを開くことができません。"
  End If

  ' ====指定されたシートを取得する==================================================
  Set worksheet = Nothing
  On Error Resume Next
    ' シート名が指定されていない場合は、アクティブシートを対象とする
    If sheetName = "" Then
      Set worksheet = workbook.ActiveSheet
    Else
      Set worksheet = workbook.Worksheets(sheetName)
    End If
  On Error Goto 0

  If worksheet Is Nothing Then
    Err.Raise 1, "", "指定されたシートが見つかりません。"
  End If

  worksheet.Activate

  ' ====指定されたセルを取得する==================================================
  Dim rawData, formattedData()
  rawData = worksheet.Range("A1").CurrentRegion.Value
  
  ' 配列のサイズを取得
  Dim rowCount, colCount
  If IsArray(rawData) Then
    On Error Resume Next
    rowCount = UBound(rawData, 1)
    colCount = UBound(rawData, 2)
    On Error GoTo 0
    
    ' 日付フォーマットを保持するために配列を再構築
    ReDim formattedData(rowCount, colCount)
    
    Dim r, c, cellValue
    For r = 1 To rowCount
      For c = 1 To colCount
        cellValue = rawData(r, c)
        
        ' 日付型の場合はYYYY/MM/DD形式に変換
        If IsDate(cellValue) Then
          formattedData(r, c) = Year(CDate(cellValue)) & "/" & _
                                Right("0" & Month(CDate(cellValue)), 2) & "/" & _
                                Right("0" & Day(CDate(cellValue)), 2)
        Else
          formattedData(r, c) = cellValue
        End If
      Next
    Next
    
    KNTN_GetArraybyExcel = formattedData
  Else
    ' 単一セルの場合
    KNTN_GetArraybyExcel = rawData
  End If

  workbook.close True
  If xlsApp.Workbooks.Count = 0 Then
    xlsApp.Quit
  End If

  ' HwndプロパティはExcel2002以降のみ対応
  On Error Resume Next
  ShowUMSHighlight(xlsApp.Hwnd)
  On Error Goto 0
End Function
' -----------------------------------------------------------------------
' Sub / Function
' 指定のシートを一番前に移動する
' -----------------------------------------------------------------------
sub KNTN_MoveSheet(fname,sheetName,flgClose)
  ' workbookオブジェクトを取得する
  SetUmsVariable "$CLEAR_ARGUMENT" , ""
  SetUMSVariable "$FILE_PATH_TYPE", "1"
  SetUMSVariable "$PARSE_FILE_PATH", fname
  filePath = GetUMSVariable("$PARSE_FILE_PATH")

  If filePath = "" Then
    SetUmsVariable "$CLEAR_ARGUMENT" , ""
    SetUMSVariable "$FILE_PATH_TYPE", "2"
    SetUMSVariable "$PARSE_FILE_PATH", fname
    filePath = GetUMSVariable("$PARSE_FILE_PATH")
  End If

  If filePath = "" Then
    Err.Raise 1, "", "指定されたファイルを開くことができません。" 
  End If

  'シート名に利用できない文字があったら削除する
  tmpname =""
  For i = 1 To Len(sheetName)
    Charstr = Mid(sheetName, i, 1)
    select case Charstr
    case ":","：","\","￥","?","？","[","［","]","］","/","／","*","＊"
    case else
      tmpname = tmpname & Charstr
    end select
  next
  sheetName = tmpname

  'シート名が32文字以上ある場合は31文字までにする
  if len(sheetName) > 31 then
    sheetName = left(sheetName,31)
  end if 

  ' workbookオブジェクトを取得する
  Set workbook = Nothing
  On Error Resume Next
  ' 既存のエクセルが起動されていれば警告を抑制する
  Set existingXlsApp = Nothing
  Set existingXlsApp = GetObject(, "Excel.Application")
  existingXlsApp.DisplayAlerts = False

  Set wash = CreateObject("WinActor7.ScriptHelper")
  For Each book in wash.GetExcelWorkbooks
    SetUMSVariable "$FILE_PATH_TYPE", 0
    SetUMSVariable "$PARSE_FILE_PATH", book.FullName
    bookPath = GetUMSVariable("$PARSE_FILE_PATH")
    If StrComp(bookPath, filePath, 1) = 0 Then
      Set workbook = book
      Set xlsApp = workbook.Parent
      xlsApp.Visible = True
      Exit For
    End If
  Next
  Set wash = Nothing

  ' Workbookが存在しない場合は、新たに開く。
  If workbook Is Nothing Then
    Set xlsApp = Nothing

    ' Excelが既に開かれていたならそれを再利用する
    If Not existingXlsApp Is Nothing Then
      Set xlsApp = existingXlsApp
      xlsApp.Visible = True
    Else
      Set xlsApp = CreateObject("Excel.Application")
      xlsApp.Visible = True
    End If

    '既存ファイルが存在する場合は利用する
    Set objFS = CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(filePath) = True Then
      Set workbook = xlsApp.Workbooks.Open(filePath)
    else
      '新規ワークシートを作成
      xlsApp.Workbooks.Add
      Set workbook = xlsApp.ActiveWorkbook
      workbook.Activesheet.name = sheetName
      workbook.SaveAs(filePath)
    end if

    xlsApp.DisplayAlerts = False
    xlsApp.DisplayAlerts = True
  End If

  ' 警告の抑制を元に戻す
  existingXlsApp.DisplayAlerts = True
  Set existingXlsApp = Nothing
  On Error Goto 0

  If workbook Is Nothing Then
    Err.Raise 1, "", "指定されたファイルを開くことができません。"
  End If

  workbook.Activate

  Set worksheet = workbook.Worksheets(sheetName)

  ' ====ハイライトを表示する========================================================
  ' HwndプロパティはExcel2002以降のみ対応
  On Error Resume Next
    ShowUMSHighlight(xlsApp.Hwnd)
  On Error Goto 0

  worksheet.move workbook.Worksheets(1)
  worksheet.Activate
  workbook.save

  if flgClose = true then
    workbook.close True
    If xlsApp.Workbooks.Count = 0 Then
      xlsApp.Quit
    End If
  end if

  Set xlsApp = Nothing
  Set worksheet = Nothing
  Set workbook = Nothing

end sub
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
