<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : LMS발송관리
' Hieditor : 2020.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->

<%
dim page, reload, ridx, clmsReport, tot_cnt, tot_ordercnt, tot_subtotalprice, tot_pushycnt, tot_sendafterpushycnt, i
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
	reload = requestcheckvar(request("reload"),2)
	ridx = requestcheckvar(getNumeric(request("ridx")),10)

if page = "" then page = 1

Set clmsReport = New clms_msg_list
	clmsReport.FPageSize = 100
	clmsReport.FCurrPage = page
	clmsReport.Frectridx = ridx
    clmsReport.fLmsMsgListRealTime

tot_cnt=0
tot_ordercnt=0
tot_subtotalprice=0
tot_pushycnt=0
tot_sendafterpushycnt=0
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
        번호 : <%= ridx %>
		&nbsp;&nbsp;<font color="red">※ 30분 지연 데이터 입니다.</font>
		<br> 발송후 24시간까지 집계된 데이터 입니다.
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= clmsReport.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= clmsReport.FTotalPage %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>발송방법</td>
	<td>회원등급</td>
	<td>발송모수</td>
    <td>발송이후주문건수</td>
    <td>발송이후매출</td>
    <td>푸시수신Y건수</td>
	<td>발송이후푸시수신Y건수</td>
</tr>
<%
dim lastsendmethodresult, sub_cnt, sub_ordercnt, sub_subtotalprice, sub_pushycnt, sub_sendafterpushycnt
%>
<% if clmsReport.FresultCount>0 then %>
	<%
	for i=0 to clmsReport.FresultCount-1
	%>
	<%
	tot_cnt = tot_cnt + clmsReport.FItemList(i).fcnt
    tot_ordercnt = tot_ordercnt + clmsReport.FItemList(i).fordercnt
    tot_subtotalprice = tot_subtotalprice + clmsReport.FItemList(i).fsubtotalprice
    tot_pushycnt = tot_pushycnt + clmsReport.FItemList(i).fpushycnt
	tot_sendafterpushycnt = tot_sendafterpushycnt + clmsReport.FItemList(i).fsendafterpushycnt

	if i<>0 and lastsendmethodresult<>clmsReport.FItemList(i).fsendmethodresult then
	%>
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td colspan=2><%= Selectsendmethodname(lastsendmethodresult) %> 합계</td>
			<td><%= FormatNumber(sub_cnt,0) %></td>
			<td><%= FormatNumber(sub_ordercnt,0) %></td>
			<td><%= FormatNumber(sub_subtotalprice,0) %></td>
			<td><%= FormatNumber(sub_pushycnt,0) %></td>
			<td><%= FormatNumber(sub_sendafterpushycnt,0) %></td>
		</tr>
	<%
		sub_cnt=0
		sub_ordercnt=0
		sub_subtotalprice=0
		sub_pushycnt=0
		sub_sendafterpushycnt=0
	end if
	%>
		<tr bgcolor="#FFFFFF" align="center">
			<td>
				<%= Selectsendmethodname(clmsReport.FItemList(i).fsendmethodresult) %>
			</td>
            <td><%= getUserLevelStr(clmsReport.FItemList(i).fuserlevel) %></td>
            <td><%= FormatNumber(clmsReport.FItemList(i).fcnt,0) %></td>
            <td><%= FormatNumber(clmsReport.FItemList(i).fordercnt,0) %></td>
            <td><%= FormatNumber(clmsReport.FItemList(i).fsubtotalprice,0) %></td>
            <td><%= FormatNumber(clmsReport.FItemList(i).fpushycnt,0) %></td>
			<td><%= FormatNumber(clmsReport.FItemList(i).fsendafterpushycnt,0) %></td>
		</tr>
	<%
	lastsendmethodresult=clmsReport.FItemList(i).fsendmethodresult
	sub_cnt = sub_cnt + clmsReport.FItemList(i).fcnt
	sub_ordercnt = sub_ordercnt + clmsReport.FItemList(i).fordercnt
	sub_subtotalprice = sub_subtotalprice + clmsReport.FItemList(i).fsubtotalprice
	sub_pushycnt = sub_pushycnt + clmsReport.FItemList(i).fpushycnt
	sub_sendafterpushycnt = sub_sendafterpushycnt + clmsReport.FItemList(i).fsendafterpushycnt
	Next
	%>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td colspan=2><%= Selectsendmethodname(lastsendmethodresult) %> 합계</td>
		<td><%= FormatNumber(sub_cnt,0) %></td>
		<td><%= FormatNumber(sub_ordercnt,0) %></td>
		<td><%= FormatNumber(sub_subtotalprice,0) %></td>
		<td><%= FormatNumber(sub_pushycnt,0) %></td>
		<td><%= FormatNumber(sub_sendafterpushycnt,0) %></td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td colspan=2>총 합계</td>
		<td><%= FormatNumber(tot_cnt,0) %></td>
		<td><%= FormatNumber(tot_ordercnt,0) %></td>
		<td><%= FormatNumber(tot_subtotalprice,0) %></td>
		<td><%= FormatNumber(tot_pushycnt,0) %></td>
		<td><%= FormatNumber(tot_sendafterpushycnt,0) %></td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% End If %>
</table>

<%
Set clmsReport = Nothing
%>
<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->