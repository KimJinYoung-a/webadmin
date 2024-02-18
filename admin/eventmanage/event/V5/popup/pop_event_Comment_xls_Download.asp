<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  이벤트 코멘트 참여자 Excel 다운로드
' History : 2007.10.12 허진원 생성
'           2014.03.10 한용민 수정
'			2019.11.14 정태훈 수정 (휴대폰 정보 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
dim eCode, Sdate, Edate, limitLevel, strSql, oevent, intLoop
	eCode = Request("eC")	'이벤트코드
	Sdate = Request("Sdate")	'적용시작일
	Edate = Request("Edate")	'적용종료일
	limitLevel = Request("limitLevel")	'회원등급제한

set oevent = new ClsEventbbs
	oevent.frecteCode = eCode
	oevent.frectSdate = Sdate
	oevent.frectEdate = Edate
	oevent.frectlimitLevel = limitLevel
	oevent.fevent_comment_notpaging()

downPersonalInformation_rowcnt=oevent.ftotalcount

%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=event" & eCode & "_" & Date() & ".xls"
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>번호</td>
	<td>회원ID</td>
	<td>이름</td>
	<td>나이</td>
	<td>회원등급</td>
	<td>이메일</td>
	<td>휴대폰</td>
	<td>작성일</td>
	<td>회원가입일</td>
	<td>코멘트 내용</td>
	<td>선택번호</td>
	<td>블로그주소</td>
	<td>최근당첨일</td>
	<td>이벤트당첨횟수</td>
</tr>
<% if oevent.FresultCount>0 then %>
<% for intLoop=0 to oevent.FresultCount-1 %>
<tr>
	<td><%= oevent.FItemList(intLoop).fevtcom_idx %></td>
	<td><%= oevent.FItemList(intLoop).fuserid %></td>
	<td><%= oevent.FItemList(intLoop).fusername %></td>
	<td><%= oevent.FItemList(intLoop).fuserAge %></td>
	<td><%= getUserLevelStrByDate(oevent.FItemList(intLoop).fuserlevel, left(oevent.FItemList(intLoop).fevtcom_regdate,10)) %></td>
	<td><%= oevent.FItemList(intLoop).fusermail %></td>
	<td><%= oevent.FItemList(intLoop).fusercell %></td>
	<td><%= oevent.FItemList(intLoop).fevtcom_regdate %></td>
	<td><%= oevent.FItemList(intLoop).fjoindate %></td>
	<td><%= oevent.FItemList(intLoop).fevtcom_txt %></td>
	<td><%= oevent.FItemList(intLoop).fevtcom_point %></td>
	<td><%= oevent.FItemList(intLoop).fblogurl %></td>
	<td><%= oevent.FItemList(intLoop).fwindate %></td>
	<td><%= oevent.FItemList(intLoop).fwincnt %></td>
</tr>
<%
	if intLoop mod 1000 = 0 then
		Response.Flush		' 버퍼리플래쉬
	end if
next
%>
<% else %>
<tr><td colspan="13" align="center">조건에 맞는 참여자가 없습니다</td></tr>
<% end if %>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->