kntn_result = _
  "https://" & kntn_subdomain & ".cybozu.com/oauth2/authorization?" & _
  "client_id=" & kntn_client_id & _
  "&redirect_uri=" & kntn_redirect_endpoint & _
  "&state=state1&response_type=code" & _
  "&scope=k:app_record:read k:app_record:write k:app_settings:read k:app_settings:write k:file:read k:file:write"

SetUmsVariable $連結結果$, kntn_result