<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  이벤트 코멘트 참여자 Excel 다운로드
' History : 2007.10.12 허진원 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim eCode, Sdate, Edate, limitLevel
dim strSql

eCode = Request("eC")	'이벤트코드
Sdate = Request("Sdate")	'적용시작일
Edate = Request("Edate")	'적용종료일
limitLevel = Request("limitLevel")	'회원등급제한

	'// DB에서 목록접수
	strSql = "select " &_
			"	t1.sub_idx, t1.userid, t2.username " &_
			"	, Case When left(t2.juminno,2) < 20 Then (year(getdate())-('20' + left(t2.juminno,2))+1) " &_
			"		When left(t2.juminno,2) >= 20 Then (year(getdate())-('19' + left(t2.juminno,2))+1) " &_
			"		end as userAge " &_
			"	, t2.usercell, t2.userphone " &_
			"	, Case t3.userlevel " &_
			"		When '0' Then 'Yellow' " &_
			"		When '1' Then 'Green' " &_
			"		When '2' Then 'Blue' " &_
			"		When '3' Then 'VIP Siver' " &_
			"		When '4' Then 'VIP Gold' " &_
			"		When '5' Then 'Orange' " &_
			"		When '6' Then 'Friends' " &_
			"		When '7' Then 'Staff' " &_
			"		When '9' Then '감성매니아' " &_
			"	 End as userlevel " &_
			"	, t2.zipcode " &_
			"	, ( " &_
			"		Select top 1 t4.Addr_Si + ' ' + t4.Addr_Gu " &_
			"		From db_zipcode.[dbo].ADDR080TL as t4 " &_
			"		Where t4.Addr_zip1=left(t2.zipcode,3) and  t4.Addr_zip2=right(t2.zipcode,3) " &_
			"	) as useraddr1 " &_
			"	, t2.useraddr as useraddr2 " &_
			"	, t1.regdate, t2.regdate as joindate " &_
			"	, t1.sub_opt1 " &_
			"from [db_culture_station].[dbo].[tbl_culturestation_event_subscript] as t1 " &_
			"	Join db_user.[dbo].tbl_user_n as t2 " &_
			"		on t1.userid=t2.userid " &_
			"	Join db_user.[dbo].tbl_logindata as t3 " &_
			"		on t2.userid=t3.userid " &_
			"where t1.evt_code=" & eCode &_
			"	and t1.regdate between '" & Sdate & "' and dateadd(d,1,'" & Edate & "') "
		
		Select Case limitLevel
			Case "orange"
				strSql = strSql & "	and t3.userlevel not in ('5') "
			Case "yellow"
				strSql = strSql & "	and t3.userlevel not in ('0','5') "
			Case "white"
				strSql = strSql & "	and t3.userlevel not in ('0') "
		end Select

		rsget.Open strSql, dbget, 1
%>
<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=culturestation_event" & eCode & "_" & Date() & ".xls"
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>번호</td>
	<td>회원ID</td>
	<td>회원이름</td>
	<td>나이</td>
	<td>휴대폰</td>
	<td>전화번호</td>
	<td>회원등급</td>
	<td>우편번호</td>
	<td>행정주소</td>
	<td>번지주소</td>
	<td>응모일</td>
	<td>회원가입일</td>
	<td>선택 및 입력란 1</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td><%=rsget("sub_idx")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("userid")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("username")%></td>
	<td><%=rsget("userAge")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("usercell")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("userphone")%></td>
	<td><%=rsget("userlevel")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("zipcode")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("useraddr1")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("useraddr2")%></td>
	<td><%=rsget("regdate")%></td>
	<td><%=rsget("joindate")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("sub_opt1")%></td>
</tr>
<%
		rsget.MoveNext
		loop
	else
%>
<tr><td colspan="13" align="center">조건에 맞는 참여자가 없습니다</td></tr>
<%	end if %>
</table>
</body>
</html>
<% rsget.close %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
