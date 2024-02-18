<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  이벤트 코멘트 참여자 Excel 다운로드
' History : 2007.10.12 허진원 생성
'			2014.03.10 한용민 수정
'			2015.10.02 이종화 라이트 버전
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
dim eCode, Sdate, Edate, limitLevel, oevent
dim strSql
dim intLoop : intLoop = 0
eCode = Request("eC")	'이벤트코드
Sdate = Request("Sdate")	'적용시작일
Edate = Request("Edate")	'적용종료일

set oevent = new ClsEventbbs
	oevent.frecteCode = eCode
	oevent.frectSdate = Sdate
	oevent.frectEdate = Edate
	oevent.fevent_subscriptlite_notpaging()

downPersonalInformation_rowcnt=oevent.ftotalcount

%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=event" & eCode & "_" & Date() & "_lite.xls"
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td colspan="8">=RANDBETWEEN(BOTTOM,TOP) 최소수 , 최대수 중 1개 출력</th>
</tr>
<tr align="center">
	<td>번호</td>
	<td>회원ID</td>
	<td>전화번호</td>
	<td>응모일</td>
	<td>선택 및 입력란 1</td>
	<td>선택 및 입력란 2</td>
	<td>선택 및 입력란 3</td>
	<td>접속경로</td>
</tr>
<% if oevent.FresultCount>0 then %>
<% for intLoop=0 to oevent.FresultCount-1 %>
<tr align="center">
	<td><%= oevent.FItemList(intLoop).frownum %></td>
	<td><%= oevent.FItemList(intLoop).fuserid %></td>
	<td><%= oevent.FItemList(intLoop).fusercell %></td>
	<td><%= oevent.FItemList(intLoop).fregdate %></td>
	<td style='mso-number-format:"\@"'><%= oevent.FItemList(intLoop).fsub_opt1 %></td>
	<td><%= oevent.FItemList(intLoop).fsub_opt2 %></td>
	<td><%= oevent.FItemList(intLoop).fsub_opt3 %></td>
	<td><%= oevent.FItemList(intLoop).fdevice %></td>
</tr>
<%
	intLoop = intLoop + 1
	if intLoop mod 1000 = 0 then
		Response.Flush		' 버퍼리플래쉬
	end if
next
%>
<% else %>
<tr><td colspan="13" align="center">조건에 맞는 참여자가 없습니다</td></tr>
<%	end if %>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
