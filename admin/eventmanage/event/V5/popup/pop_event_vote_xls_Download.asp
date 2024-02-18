<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  이벤트 코멘트 참여자 Excel 다운로드
' History : 2007.10.12 허진원 생성
'			2014.03.10 한용민 수정
'			2016.03.02 원승현 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->

<%
dim eCode, Sdate, Edate, limitLevel, strSql, oevent
dim intLoop : intLoop = 0
	eCode = Request("eC")	'이벤트코드
	Sdate = Request("Sdate")	'적용시작일
	Edate = Request("Edate")	'적용종료일
	limitLevel = Request("limitLevel")	'회원등급제한

set oevent = new ClsEventbbs
	oevent.frecteCode = eCode
	oevent.frectSdate = Sdate
	oevent.frectEdate = Edate
	oevent.frectlimitLevel = limitLevel
	oevent.fevent_subscript_notpaging()

downPersonalInformation_rowcnt=oevent.ftotalcount

%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
'엑셀 출력시작
'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=event " & eCode & "_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>번호</td>
	<td>회원ID</td>
	<td>나이</td>
	<td>회원등급</td>
	<td>응모일</td>
	<td>회원가입일</td>
	<td>선택 및 입력란 1</td>
	<td>선택 및 입력란 2</td>
	<td>선택 및 입력란 3</td>
	<td>사이트구분</td>
</tr>

<% if oevent.FresultCount>0 then %>
<% for intLoop=0 to oevent.FresultCount-1 %>
<tr>
	<td><%= oevent.FItemList(intLoop).fsub_idx %></td>
	<td style='mso-number-format:"\@"'><%= oevent.FItemList(intLoop).fuserid %></td>
	<td><%= oevent.FItemList(intLoop).fuserAge %></td>
	<td><%= getUserLevelStrByDate(oevent.FItemList(intLoop).fuserlevel, left(oevent.FItemList(intLoop).fregdate,10)) %></td>
	<td><%= oevent.FItemList(intLoop).fregdate %></td>
	<td><%= oevent.FItemList(intLoop).fjoindate %></td>
	<td style='mso-number-format:"\@"'><%= oevent.FItemList(intLoop).fsub_opt1 %></td>
	<td><%= oevent.FItemList(intLoop).fsub_opt2 %></td>
	<td><%= oevent.FItemList(intLoop).fsub_opt3 %></td>
	<td><%= oevent.FItemList(intLoop).fsitegubun %></td>
</tr>
<%
	if (intLoop+1) mod 500 = 0 then
		Response.Flush		' 버퍼리플래쉬
	end if
next
%>
<% else %>
	<tr>
		<td colspan="13" align="center">조건에 맞는 참여자가 없습니다</td>
	</tr>
<% end if %>

</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
