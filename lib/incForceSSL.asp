<%
	'강제 SSL 전환(2017.05.24; 허진원)
	If Request.ServerVariables("SERVER_PORT")=80 and application("Svr_Info")<>"Dev" Then
		Dim strSecureURL
		strSecureURL = "https://"
		strSecureURL = strSecureURL & Request.ServerVariables("SERVER_NAME")
		strSecureURL = strSecureURL & Request.ServerVariables("URL")
		if Request.ServerVariables("QUERY_STRING")<>"" then
			strSecureURL = strSecureURL & "?" & Request.ServerVariables("QUERY_STRING")
		end if
		Response.Redirect strSecureURL
	End If
%>