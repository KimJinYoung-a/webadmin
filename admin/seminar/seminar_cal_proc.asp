<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim mode, sqlStr, strsql, idx, Fusername
	Dim roomidx, reserdate, start_time, end_time, start_date, end_date, groupname, usepurpose, usercell, useSu, etc, lecnum, isusing, adminID
	mode			= request("mode")
	idx				= request("idx")
	roomidx			= request("roomidx")
	reserdate		= request("reserdate")
	start_time		= request("start_time")
	end_time		= request("end_time")
	start_date		= reserdate&" "&start_time
	end_date		= reserdate&" "&end_time
	groupname		= request("groupname")
	usepurpose 		= request("usepurpose")
	usercell		= request("usercell")
	useSu			= request("useSu")
	etc 			= request("etc")
	lecnum			= request("lecnum")
	isusing			= request("isusing")
	adminID 		= session("ssBctId")

	if (checkNotValidHTML(groupname) = True) Then
		response.write "<script>alert('그룹명에는 HTML을 사용하실 수 없습니다.');</script>"
		dbget.close()	:	response.End
	End If

	if (checkNotValidHTML(usercell) = True) Then
		response.write "<script>alert('연락처에는 HTML을 사용하실 수 없습니다.');</script>"
		dbget.close()	:	response.End
	End If

	if (checkNotValidHTML(etc) = True) Then
		response.write "<script>alert('기타사항에는 HTML을 사용하실 수 없습니다.');</script>"
		dbget.close()	:	response.End
	End If


If mode = "write" Then
	isusing = "Y"
	sqlStr = "select adminID FROM [db_partner].[dbo].[tbl_seminar_schedule] " + VBCrlf
	sqlStr = sqlStr + " where start_date < '"&end_date&"' " & vbcrlf
	sqlStr = sqlStr + " and end_date > '"&start_date&"' " & vbcrlf
	sqlStr = sqlStr + " and roomidx = '" + roomidx + "' and isusing = 'Y' "

	rsget.Open sqlStr,dbget,1
	If  not rsget.EOF  Then
		Fusername   =  rsget("adminID")
	End If
	rsget.close
	
	If Fusername <> "" Then
		response.write "<script language='JavaScript'>alert('" + Fusername + "님의 예약과 겹칩니다.\n다시 확인하시고 선택해주세요...');history.back(-1);</script>"
		dbget.close()	:	response.End
	End If

	strsql = ""
	strsql = strsql &" insert into [db_partner].[dbo].tbl_seminar_schedule "& vbCrLf
	strsql = strsql &" (roomidx, start_date, end_date, groupname, usepurpose, usercell, useSu, etc, lecnum, isusing, adminID, regdate) " & vbCrLf
	strsql = strsql &" values " & vbCrLf
	strsql = strsql &" ('"&roomidx&"', '"&start_date&"', '"&end_date&"', '"&groupname&"', '"&usepurpose&"', '"&usercell&"', '"&useSu&"', '"&etc&"', '"&lecnum&"', '"&isusing&"', '"&adminID&"', getdate())" & vbCrLf
	dbget.execute strsql

ElseIf mode="modify" Then

		sqlStr = "select adminID FROM [db_partner].[dbo].[tbl_seminar_schedule] " + VBCrlf
		sqlStr = sqlStr + " where start_date < '"&end_date&"' " & vbcrlf
		sqlStr = sqlStr + " and end_date > '"&start_date&"' " & vbcrlf
		sqlStr = sqlStr + " and roomidx = '" + roomidx + "' and isusing = 'Y' and idx <> '"&idx&"' "

		rsget.Open sqlStr,dbget,1
		If  not rsget.EOF  Then
			Fusername   =  rsget("adminID")
		End If
		rsget.close
	
		If Fusername <> "" Then
			response.write "<script language='JavaScript'>alert('" + Fusername + "님의 예약과 겹칩니다.\n다시 확인하시고 선택해주세요...');history.back(-1);</script>"
			dbget.close()	:	response.End
		End If
	
		sqlStr = ""
		sqlStr = sqlStr & " Update [db_partner].[dbo].[tbl_seminar_schedule] set " & vbCrLf
		sqlStr = sqlStr & " roomidx = '"&roomidx&"', start_date = '"&start_date&"', end_date = '"&end_date&"', groupname = '"&groupname&"',usepurpose = '"&usepurpose&"',usercell = '"&usercell&"',useSu = '"&useSu&"',etc = '"&etc&"',lecnum = '"&lecnum&"',isusing = '"&isusing&"' " & vbCrLf
		sqlStr = sqlStr & " where idx = '"& idx &"' "
		dbget.execute sqlStr
End If
%>
<script language="javascript">
	alert('OK');
	window.opener.location.reload();
	self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->