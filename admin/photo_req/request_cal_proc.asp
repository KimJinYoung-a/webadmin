<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 촬영 요청 수정 & 뷰 페이지
' History : 2012.03.15 김진영 생성
'			2018.02.09 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim mode, sqlStr, strsql2,  Fusername, req_no, useyn
	Dim req_photo_user, req_date, start_time, end_time, start_date, end_date, req_comment, query1, max_req_no

	req_photo_user	= request("req_photo_user")
	req_date		= request("req_date")
	start_time		= request("start_time")
	end_time		= request("end_time")
	start_date		= req_date&" "&start_time
	end_date		= req_date&" "&end_time
	req_comment		= request("req_comment")
	req_no			= request("rno")
	mode 			= request("mode")
	useyn			= request("useyn")

If mode = "write" Then

	sqlStr = "select req_name  from [db_partner].[dbo].[tbl_photo_req] a" + VBCrlf
	sqlStr = sqlStr + " left join [db_partner].[dbo].[tbl_photo_schedule] b" + vbcrlf
	sqlStr = sqlStr + " on a.req_no = b.req_no" + vbcrlf
	sqlStr = sqlStr + " left join [db_partner].[dbo].[tbl_photo_user] c" + vbcrlf
	sqlStr = sqlStr + " on b.req_photo = c.user_id" + vbcrlf
	sqlStr = sqlStr + " where b.start_date < '"&end_date&"' " & vbcrlf
	sqlStr = sqlStr + " and b.end_date > '"&start_date&"' " & vbcrlf
	sqlStr = sqlStr + " and c.user_id = '" + req_photo_user + "' and c.user_type = '1' and a.use_yn = 'Y'"

	rsget.Open sqlStr,dbget,1
	If  not rsget.EOF  Then
		Fusername   =  rsget("req_name")
	End If
	rsget.close
	
	If Fusername <> "" Then
		response.write "<script language='JavaScript'>alert('" + Fusername + "님의 예약과 겹칩니다.\n다시 확인하시고 선택해주세요...');history.back(-1);</script>"
		dbget.close()	:	response.End
	End If

	sqlStr = ""
	sqlStr = sqlStr & "Insert into [db_partner].[dbo].tbl_photo_req" & vbCrLf
	sqlStr = sqlStr & "(req_gubun, req_use, prd_name, prd_type, " & vbCrLf
	sqlStr = sqlStr & " req_date, req_etc1, use_yn, req_name, req_comment)" & vbCrLf
	sqlStr = sqlStr & " Values " & vbCrLf
	sqlStr = sqlStr & "('', '', '', '', " & vbCrLf
	sqlStr = sqlStr & " getdate(), '', 'S', '"&session("ssBctid")&"', '"&Cstr(req_comment)&"')" & vbCrLf
	dbget.execute sqlStr

	query1 = " select max(req_no)as req_no from [db_partner].[dbo].tbl_photo_req"
	rsget.Open query1,dbget
	IF not rsget.EOF THEN
		max_req_no 	= rsget("req_no")
	End IF			
	rsget.Close	

	strsql2 = ""
	strsql2 = strsql2 &" insert into [db_partner].[dbo].tbl_photo_schedule "& vbCrLf
	strsql2 = strsql2 &" (req_no, start_date, end_date, schedule_regdate, req_photo) " & vbCrLf
	strsql2 = strsql2 &" values " & vbCrLf
	strsql2 = strsql2 &" ('"&max_req_no&"', '"&start_date&"', '"&end_date&"', getdate(), '"&Cstr(req_photo_user)&"')" & vbCrLf
	dbget.execute strsql2

ElseIf mode="modify" Then

	If useyn = "Y" Then
		sqlStr = ""
		sqlStr = sqlStr & " Update [db_partner].[dbo].tbl_photo_req set " & vbCrLf
		sqlStr = sqlStr & " use_yn = 'N' " & vbCrLf
		sqlStr = sqlStr & " where req_no = '"& req_no &"' "
		dbget.execute sqlStr
	Else
		sqlStr = "select req_name  from [db_partner].[dbo].[tbl_photo_req] a" + VBCrlf
		sqlStr = sqlStr + " left join [db_partner].[dbo].[tbl_photo_schedule] b" + vbcrlf
		sqlStr = sqlStr + " on a.req_no = b.req_no" + vbcrlf
		sqlStr = sqlStr + " left join [db_partner].[dbo].[tbl_photo_user] c" + vbcrlf
		sqlStr = sqlStr + " on b.req_photo = c.user_id" + vbcrlf
		sqlStr = sqlStr + " where b.start_date < '"&end_date&"' " & vbcrlf
		sqlStr = sqlStr + " and b.end_date > '"&start_date&"' " & vbcrlf
		sqlStr = sqlStr + " and c.user_id = '" + req_photo_user + "' and c.user_type = '1' and a.use_yn = 'Y'"
		rsget.Open sqlStr,dbget,1
		If  not rsget.EOF  Then
			Fusername   =  rsget("req_name")
		End If
		rsget.close
	
		If Fusername <> "" Then
			response.write "<script language='JavaScript'>alert('" + Fusername + "님의 예약과 겹칩니다.\n다시 확인하시고 선택해주세요...');history.back(-1);</script>"
			dbget.close()	:	response.End
		End If
	
		sqlStr = ""
		sqlStr = sqlStr & " Update [db_partner].[dbo].tbl_photo_req set " & vbCrLf
		sqlStr = sqlStr & " req_comment = '"&Cstr(req_comment)&"'" & vbCrLf
		sqlStr = sqlStr & " where req_no = '"& req_no &"' "
		dbget.execute sqlStr
	
		sqlStr = ""
		sqlStr = sqlStr & " Update [db_partner].[dbo].tbl_photo_schedule set " & vbCrLf
		sqlStr = sqlStr & " start_date = '"&start_date&"', end_date = '"&end_date&"', req_photo = '"&Cstr(req_photo_user)&"'" & vbCrLf
		sqlStr = sqlStr & " where req_no = '"& req_no &"' "
		dbget.execute sqlStr
	End If
End If
%>
<script language="javascript">
	alert('OK');
	window.opener.location.reload();
	self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->