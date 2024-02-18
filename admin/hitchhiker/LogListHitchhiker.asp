<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 히치하이커 재발송 로그팝업
'	History		: 2012.04.23 김진영 생성
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhikerCls.asp"-->
<%
Dim sUserID
Dim pMode, iHVol
Dim clsUInfo
Dim sZip, sAdd1, sAdd2, sP, sC, sChk
Dim iAV, clsLogList,page, i
sUserID = Request("sUID")
pMode	= Request("pMode")
iHVol = Request("iHV")
iAV = Request("iAV")

page = request("page")
If page = "" Then page = 1

set clsLogList = new Chitchhiker
	clsLogList.FHVol = iHVol
	clsLogList.FPageSize = 20
	clsLogList.FCurrPage = page
	clsLogList.fnHitchLoglist()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body leftmargin=0 topmargin=0>
<div style="padding:10 10 0 10">
<img src="/images/icon_star.gif" align="absmiddle"> <font color="red"><strong> 히치하이커 Vol.<%=iHVol%> 재발송 LOG 리스트 </strong></font><br>
<hr>
</div>
<script>
function gosubmit(page){
    document.frmuser.page.value=page;
	document.frmuser.submit();
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmuser" method="post" onSubmit="return jsSubmit(this);">
<input type="hidden" name="pMode" value="<%=pMode%>">
<input type="hidden" name="iHV" value="<%=iHVol%>">
<input type="hidden" name="iAV" value="<%=iAV%>">
<input type="hidden" name="page" value="<%=page%>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>ID</td>
	<td>초기 발송회차</td>
	<td>변경 발송회차</td>
	<td>변경일</td>
	<td>변경ID</td>
</tr>
<%
	If clsLogList.FResultcount <> 0 Then
		For i = 0 to clsLogList.FResultcount -1
%>
<tr align="center" bgcolor="ffffff">
	<td><%=clsLogList.FHitchLogList(i).Fuserid%></td>
	<td><%=chkIIF(clsLogList.FHitchLogList(i).FiAvol="0","<font color=blue>새발송</font>",clsLogList.FHitchLogList(i).FiAvol)%></td>
	<td><%=clsLogList.FHitchLogList(i).FiAvol2%></td>
	<td><%=clsLogList.FHitchLogList(i).FRegdate%></td>
	<td><%=clsLogList.FHitchLogList(i).FAdminId & "(" & clsLogList.FHitchLogList(i).FAdminNm & ")"%></td>
</tr>
<%
		Next
%>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="5" align="center">
       	<% If clsLogList.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= ohistory.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + clsLogList.StartScrollPage to clsLogList.StartScrollPage + clsLogList.FScrollCount - 1 %>
			<% If (i > clsLogList.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(clsLogList.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If clsLogList.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<%
	Else
%>
<tr align="center" bgcolor="ffffff">
	<td colspan="5">데이터가 없습니다</td>
</tr>
<%End if%>
</form>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->