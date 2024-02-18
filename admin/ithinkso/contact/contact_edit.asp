<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : 아이띵소 제휴 관리
' Hieditor : 2013.05.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language='javascript'>

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<% if session("sslgnMethod")<>"S" then %>
	<!-- USB키 처리 시작 (2008.06.23;허진원) -->
	<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
	<script language="javascript" src="/js/check_USBToken.js"></script>
	<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body bgcolor="#F4F4F4" onload="checkUSBKey()">

<!-- #include virtual="/lib/classes/ithinkso/contact/contact_cls_ithinkso.asp"-->

<%
dim idx, contact_gubun, username, email, hp, countryname, title, contents, uploadfileurl, isusing, regdate
	idx = request("idx")

if idx = "" then
	response.write "<script language='javascript'>"
	response.write "	alert('IDX가 없습니다.');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end
end if

dim ocontact, i
set ocontact = new Ccontact_list
	ocontact.frectidx = idx
	
	if idx <> "" then
		ocontact.fcontact_one()
	end if
	
	if ocontact.ftotalcount > 0 then
		idx = ocontact.FOneItem.fidx
		contact_gubun = ocontact.FOneItem.fcontact_gubun
		username = ocontact.FOneItem.fusername
		email = ocontact.FOneItem.femail
		hp = ocontact.FOneItem.fhp
		countryname = ocontact.FOneItem.fcountryname
		title = ocontact.FOneItem.ftitle
		contents = ocontact.FOneItem.fcontents
		uploadfileurl = ocontact.FOneItem.fuploadfileurl
		isusing = ocontact.FOneItem.fisusing
		regdate = ocontact.FOneItem.fregdate
	end if
set ocontact = nothing
%>

<script language="javascript">

//등록
function contactreg(){
	if(frm.isusing.value == ""){
		alert("사용여부를 선택하세요.");
		frm.isusing.focus();
		return;
	}
	
	frm.submit();	
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="/admin/ithinkso/contact/contact_process.asp">
<input type="hidden" name="mode" value="contactreg">
<tr bgcolor="#FFFFFF">
	<td align="center">번호</td>
	<td>
		<%=idx%><input type="hidden" name="idx" value="<%=idx%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">고객명</td>
	<td>
		<%=username%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">이메일</td>
	<td>
		<%=email%>
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td align="center">전화번호</td>
	<td>
		<%=hp%>
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td align="center">국가명</td>
	<td>
		<%=countryname%>
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td align="center">제목</td>
	<td>
		<%=title%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">내용</td>
	<td>
		<%= nl2br(contents) %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">등록일</td>
	<td>
		<%=regdate%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">파일</td>
	<td>
		<% if uploadfileurl <> "" then %>
			<a href="<%= uploadfileurl %>" onfocus="this.blur();"><%= uploadfileurl %></a>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">사용여부</td>
	<td>
		<% drawSelectBoxisusingYN "isusing", isusing, "" %>
	</td>
</tr>
</form>	
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" value="수정" onclick="contactreg();" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
