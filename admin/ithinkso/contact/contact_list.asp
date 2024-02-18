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
Dim ocontact, i, page, username, title, isusing, menupos
	username = request("username")
	title = request("title")
	isusing = request("isusing")
	page = request("page")
	menupos = request("menupos")

if page = "" then page = 1

'// 이벤트 리스트
set ocontact = new Ccontact_list
	ocontact.FPageSize = 50
	ocontact.FCurrPage = page
	ocontact.frectusername = username
	ocontact.frecttitle = title
	ocontact.frectisusing = isusing
	ocontact.fcontact_list()
%>

<script language="javascript">

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

function pop_edit(idx){
	var pop_edit = window.open('/admin/ithinkso/contact/contact_edit.asp?idx='+idx,'pop_edit','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_edit.focus();
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
			<tr>
				<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
					<font color="#333333"><b>[ON]해외상품관리&gt;&gt;아이띵소해외제휴문의</b></font>
				</td>
				
				<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">
					<!-- 마스터이상 메뉴권한 설정 -->
					
					<a href="Javascript:PopMenuEdit('1491');"><img src="/images/icon_chgauth.gif" border="0" valign="bottom"></a>
					
					<!-- Help 설정 -->
					
					<a href="Javascript:PopMenuHelp('1491');"><img src="/images/icon_help.gif" border="0" valign="bottom"></a>
					
				</td>
				
			</tr>
		</table>
	</td>
</tr>
</table>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="editor_no">
<input type="hidden" name="page" value=1>
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 고객명 : <input type="text" name="username" value="<%= username%>" size="10">
		&nbsp;&nbsp;
		* 제목 : <input type="text" name="title" value="<%= title%>" size="20">
		&nbsp;&nbsp;
		* 사용여부 : <% drawSelectBoxisusingYN "isusing", isusing," onchange='frmsubmit("""");'" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<Br>		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ocontact.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ocontact.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>번호</td>
	<td>고객명</td>
	<td>국가명</td>
	<td>제목</td>
	<td>사용여부</td>
	<td>등록일</td>
	<td>비고</td>				
</tr>
<% if ocontact.FresultCount>0 then %>
	
<% for i=0 to ocontact.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			

<% if ocontact.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#ffffff" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% else %>    
<tr align="center" bgcolor="#e1e1e1" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#e1e1e1';>
<% end if %>

	<td>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td>
		<%= ocontact.FItemList(i).fidx %>
	</td>		
	<td>
		<%= ocontact.FItemList(i).fusername %>
	</td>
	<td>
		<%= ocontact.FItemList(i).fcountryname %>
	</td>
	<td align="left">
		<%= chrbyte(ocontact.FItemList(i).ftitle,50,"Y") %>
	</td>
		
	<td>
		<%= ocontact.FItemList(i).fisusing %>
	</td>
	<td>
		<%= left(ocontact.FItemList(i).fregdate,10) %>
	</td>	
	<td width=70>
		<input type="button" value="보기" onclick="pop_edit('<%= ocontact.FItemList(i).fidx %>');" class="button">
	</td>	
</tr>   
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ocontact.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= ocontact.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ocontact.StartScrollPage to ocontact.StartScrollPage + ocontact.FScrollCount - 1 %>
			<% if (i > ocontact.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ocontact.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ocontact.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
