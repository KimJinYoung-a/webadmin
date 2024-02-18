<%
'###########################################################
' Description : 헤더 UTF8 버전
' History : 2016.11.24 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/lib/classes/admin/MenuCls.asp"-->

<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// 즐겨찾기
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)
%>
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

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;

	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "즐겨찾기에서 제외하시겠습니까?";
	} else {
		msg = "즐겨찾기에 추가하시겠습니까?";
	}

	ret = confirm(msg);

	if (ret) {
		frm.submit();
	}
}
</script>
<% if session("sslgnMethod")<>"S" then %>
<!-- USB키 처리 시작 (2008.06.23;허진원) -->
<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
<script language="javascript" src="/js/check_USBToken.js"></script>
<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body bgcolor="#F4F4F4" <% if session("sslgnMethod")<>"S" then %>onload="checkUSBKey()"<% end if %>>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<% if (imenuposStr<>"") then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="400" style="padding:5; border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>"  background="/images/menubar_1px.gif">
						<font color="#333333"><b><%= imenuposStr %></b></font>
					</td>
					<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
					<input type="hidden" name="mode" value="">
					<input type="hidden" name="menu_id" value="<%= menupos %>">
					</form>
					<td align="right" style="border-bottom:1px solid <%= adminColor("tablebg") %>" bgcolor="#F4F4F4">
						<% if (menupos > 1) then %>
							<% if (IsMenuFavoriteAdded) then %>
								<a href="javascript:fnMenuFavoriteAct('delonefavorite')"><font color="blue">[즐겨찾기]</font></a>
							<% else %>
								<a href="javascript:fnMenuFavoriteAct('addonefavorite')"><font color="black">[즐겨찾기]</font></a>
							<% end if %>
						<% end if %>
						<!-- 마스터이상 메뉴권한 설정 -->
						<% if C_ADMIN_AUTH then %>
						<a href="Javascript:PopMenuEdit('<%= menupos %>');"><img src="/images/icon_chgauth.gif" border="0" valign="bottom"></a>
						<% end if %>
						<!-- Help 설정 -->
						<% if (imenuposhelp<>"") or (C_ADMIN_AUTH) then %>
						<a href="Javascript:PopMenuHelp('<%= menupos %>');"><img src="/images/icon_help.gif" border="0" valign="bottom"></a>
						<% end if %>
					</td>

				</tr>
			</table>
		</td>
	</tr>

	<!--	설명 있으면 들어갑니다.	-->
	<% if imenuposnotice<>"" then %>
	<tr bgcolor="#FFFFFF">
		<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
			<%= nl2br(imenuposnotice) %>
		</td>
	</tr>
	<% end if %>
</table>

<p>
<% end if %>
