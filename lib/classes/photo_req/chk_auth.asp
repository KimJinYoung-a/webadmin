<%
	Dim g_TeamJang, g_PartJang, g_MyTeam, g_MyPart, g_MenuPos
	IF application("Svr_Info")="Dev" THEN
		g_MenuPos   = "1404"		'### 메뉴번호 지정.
	Else
		g_MenuPos   = "1404"		'### 메뉴번호 지정.
	End If
	g_TeamJang = "x"
	g_PartJang = "x"
	
	
	'####### 관리자, 마스터 권한 #######
	If CInt(session("ssAdminLsn")) =< CInt("2") Then
		g_TeamJang = "o"
	End If
	
	'####### 파트선임권한 #######
	If CInt(session("ssAdminLsn")) = CInt("3") Then
		g_PartJang = "o"
	End If
	



	g_MyPart = session("ssAdminPsn")
	
	'####### 시스템팀, 경영관리팀, 패션사업팀 팀구분 #######
	If CInt(g_MyPart) = CInt("7") OR CInt(g_MyPart) = CInt("8") OR CInt(g_MyPart) = CInt("17") Then
		g_MyTeam = g_MyPart
	End IF

	'####### 아이띵소팀구분 #######
	If CInt(g_MyPart) = CInt("15") OR CInt(g_MyPart) = CInt("19") Then
		g_MyTeam = "15,19"
	End IF

	'####### 오프라인팀구분 #######
	If CInt(g_MyPart) = CInt("13") OR CInt(g_MyPart) = CInt("18") Then
		g_MyTeam = "13,18"
	End IF

	'####### 운영관리팀 팀구분 #######
	If CInt(g_MyPart) = CInt("9") OR CInt(g_MyPart) = CInt("10") Then
		g_MyTeam = "9,10"
	End IF
	
	'####### 텐바이텐사업팀(마케팅팀,MD,WD) 팀구분 #######
	If CInt(g_MyPart) = CInt("11") OR CInt(g_MyPart) = CInt("12") OR CInt(g_MyPart) = CInt("14") OR CInt(g_MyPart) = CInt("16") Then
		g_MyTeam = "11,12,14,16"
	End IF
	
	

	If Request.ServerVariables("REMOTE_ADDR") = "61.252.133.15" Then
		'### test용
		'g_TeamJang = "o"
		'g_PartJang = "o"
	End If
%>
