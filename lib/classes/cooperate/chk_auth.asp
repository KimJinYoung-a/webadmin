<%
	Dim g_TeamJang, g_PartJang, g_MyTeam, g_MyPart, g_MenuPos, g_VertiHoriz
	IF application("Svr_Info")="Dev" THEN
		g_MenuPos   = "1106"		'### �޴���ȣ ����.
	Else
		g_MenuPos   = "1109"		'### �޴���ȣ ����.
	End If
	g_VertiHoriz = Request.cookies("scmcooperatevertihoriz")
	
	g_TeamJang = "x"
	g_PartJang = "x"
	
	
	'####### ������, ������ ���� #######
	If CInt(session("ssAdminLsn")) =< CInt("2") Then
		g_TeamJang = "o"
	End If
	
	'####### ��Ʈ���ӱ��� #######
	If CInt(session("ssAdminLsn")) = CInt("3") Then
		g_PartJang = "o"
	End If
	



	g_MyPart = session("ssAdminPsn")
	
	'####### �ý�����, �мǻ���� ������ #######
	If CInt(g_MyPart) = CInt("7") OR CInt(g_MyPart) = CInt("30") OR CInt(g_MyPart) = CInt("31") OR CInt(g_MyPart) = CInt("17") Then
		g_MyTeam = g_MyPart
	End IF

	'####### �濵������ - �繫ȸ����, �濵������ - �λ米����Ʈ #######
	If CInt(g_MyPart) = CInt("8") OR CInt(g_MyPart) = CInt("20") Then
		g_MyTeam = "8,20"
	End IF

	'####### ���̶�������� #######
	If CInt(g_MyPart) = CInt("15") OR CInt(g_MyPart) = CInt("19") Then
		g_MyTeam = "15,19"
	End IF

	'####### �������������� #######
	If CInt(g_MyPart) = CInt("13") OR CInt(g_MyPart) = CInt("18") OR CInt(g_MyPart) = CInt("24") OR CInt(g_MyPart) = CInt("25") OR CInt(g_MyPart) = CInt("26") OR CInt(g_MyPart) = CInt("27") OR CInt(g_MyPart) = CInt("28") OR CInt(g_MyPart) = CInt("29") Then
		g_MyTeam = "13,18,24,25,26,27,28,29"
	End IF

	'####### ������� ������ #######
	If CInt(g_MyPart) = CInt("9") OR CInt(g_MyPart) = CInt("10") Then
		g_MyTeam = "9,10"
	End IF
	
	'####### �ٹ����ٻ����(��������,MD,WD) ������ #######
	If CInt(g_MyPart) = CInt("11") OR CInt(g_MyPart) = CInt("12") OR CInt(g_MyPart) = CInt("14") OR CInt(g_MyPart) = CInt("33") OR CInt(g_MyPart) = CInt("16") OR CInt(g_MyPart) = CInt("21") OR CInt(g_MyPart) = CInt("22") OR CInt(g_MyPart) = CInt("23") Then
		g_MyTeam = "11,12,14,16,21,22,23"
	End IF
	
	

	If Request.ServerVariables("REMOTE_ADDR") = "61.252.133.15" Then
		'### test��
		'g_TeamJang = "o"
		'g_PartJang = "o"
	End If
%>
