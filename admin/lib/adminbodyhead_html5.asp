<%
'###########################################################
' Description : html5 ���
' History : 2016.03.15 �ѿ�� ����(html5���� ����� scm ��ü�� ǥ�� �԰����� �ڵ��� ���� �ʾƼ� ����. �׷��� ������ ���ζ�)
'###########################################################
%>
<!-- #include virtual="/lib/classes/admin/MenuCls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// ���ã��
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge" /><% '�����ͺм� �Ŵ����� ��� %>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
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
		msg = "���ã�⿡�� �����Ͻðڽ��ϱ�?";
	} else {
		msg = "���ã�⿡ �߰��Ͻðڽ��ϱ�?";
	}

	ret = confirm(msg);

	if (ret) {
		frm.submit();
	}
}
</script>
<% if session("sslgnMethod")<>"S" then %>
<!-- USBŰ ó�� ���� (2008.06.23;������) -->
<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
<script language="javascript" src="/js/check_USBToken.js"></script>
<!-- USBŰ ó�� �� -->
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
								<a href="javascript:fnMenuFavoriteAct('delonefavorite')"><font color="blue">[���ã��]</font></a>
							<% else %>
								<a href="javascript:fnMenuFavoriteAct('addonefavorite')"><font color="black">[���ã��]</font></a>
							<% end if %>
						<% end if %>
						<!-- �������̻� �޴����� ���� -->
						<% if C_ADMIN_AUTH then %>
						<a href="Javascript:PopMenuEdit('<%= menupos %>');"><img src="/images/icon_chgauth.gif" border="0" valign="bottom"></a>
						<% end if %>
						<!-- Help ���� -->
						<% if (imenuposhelp<>"") or (C_ADMIN_AUTH) then %>
						<a href="Javascript:PopMenuHelp('<%= menupos %>');"><img src="/images/icon_help.gif" border="0" valign="bottom"></a>
						<% end if %>
					</td>

				</tr>
			</table>
		</td>
	</tr>

	<!--	���� ������ ���ϴ�.	-->
	<% if imenuposnotice<>"" then %>
	<tr bgcolor="#FFFFFF">
		<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
			<%= nl2br(imenuposnotice) %>
		</td>
	</tr>
	<% end if %>
</table>
<% end if %>
