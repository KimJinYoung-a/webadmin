<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  이벤트 코멘트 참여자 Excel 다운로드
' History : 2007.10.12 허진원 생성
'           2014.03.03 허진원 ; 개인정보 데이터 제거
'			2014.03.10 한용민 수정
'			2014.08.13 이종화 비회원 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
response.write "사용중지 매뉴입니다. 신매뉴로 사용 부탁드립니다."
response.end

dim eCode, Sdate, Edate, limitLevel
dim strSql

eCode = Request("eC")	'이벤트코드
Sdate = Request("Sdate")	'적용시작일
Edate = Request("Edate")	'적용종료일

	'// DB에서 목록접수
	strSql = "select " &_
			"	t1.sub_idx " &_
			"	, t1.regdate " &_
			"	, t1.sub_opt1, t1.sub_opt2, t1.sub_opt3 " &_
			" from [db_event].[dbo].[tbl_event_subscript] as t1 " &_
			" left join db_user.dbo.tbl_invalid_user iu" &_
			" 	on t1.userid=iu.invaliduserid" &_
			" 	and iu.isusing='Y'" &_
			" 	and iu.gubun='ONEVT'" &_			
			" where iu.idx is null and t1.evt_code=" & eCode &_
			"	and t1.regdate between '" & Sdate & "' and dateadd(d,1,'" & Edate & "') "

		rsget.Open strSql, dbget, 1
%>
<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=event" & eCode & "_" & Date() & ".xls"
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>번호</td>
	<td>응모일</td>
	<td>선택 및 입력란 1</td>
	<td>선택 및 입력란 2</td>
	<td>선택 및 입력란 3</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td><%=rsget("sub_idx")%></td>
	<td><%=rsget("regdate")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("sub_opt1")%></td>
	<td><%=rsget("sub_opt2")%></td>
	<td><%=rsget("sub_opt3")%></td>
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
