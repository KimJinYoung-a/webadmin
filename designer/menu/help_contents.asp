<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/admin/OLDmenucls.asp"-->
<%
dim menupos
menupos = request("menupos")


dim imenupos, menuposStr
set imenupos = new CMenu
imenupos.FRectID = menupos

if menupos<>"" then
	imenupos.getOneMenu
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<SCRIPT LANGUAGE="Javascript" SRC="/js/xmlTree.js"></SCRIPT>
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none}
-->
</style>
<script language='javascript'>
function SaveMenuEtc(frm){
	if (frm.menuname.value.length<1){
		alert('�޴����� �Է��ϼ���.');
		frm.menuname.focus();
		return;
	}

	if (frm.viewidx.value.length<1){
		alert('���ü����� �Է��ϼ���.');
		frm.viewidx.focus();
		return;
	}


	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}
</script>
</head>
<body topmargin="0">
<% if imenupos.FResultCount>0 then %>
	<!-- ǥ ��ܹ� ����-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
		</tr>
		<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	<%= imenupos.FOneItem.FMenuStr %>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
	</table>
	<!-- ǥ ��ܹ� ��-->
	
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#bababa">

		<% if imenupos.FOneItem.Fmenuposnotice<>"" then %>
		<tr bgcolor="<%= adminColor("tabletop") %>">
			<td>
				<img src="/images/icon_arrow_down.gif" align="absbottom">
				<font color="red"><strong>���ܼ���</strong></font><br>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>
				<%= nl2br(imenupos.FOneItem.Fmenuposnotice) %>
			</td>
		</tr>
		<% end if %>
	
		<% if imenupos.FOneItem.Fmenuposhelp<>"" then %>
		<tr bgcolor="<%= adminColor("tabletop") %>">
			<td>
				<img src="/images/icon_arrow_down.gif" align="absbottom">
				<font color="red"><strong>�󼼼���</strong></font><br>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>
				<%= nl2br(imenupos.FOneItem.Fmenuposhelp) %>
			</td>
		</tr>
		<% end if %>
		
		<% if imenupos.FOneItem.Fmenuposnotice="" and imenupos.FOneItem.Fmenuposhelp="" then %>
		<tr align="center" bgcolor="#FFFFFF">
			<td>
				������ �����ϴ�.
			</td>
		</tr>
		<% end if %>
		
	</table>
	
	<!-- ǥ �ϴܹ� ����-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	    <tr valign="bottom" height="25">
	        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="bottom" align="center">&nbsp;</td>
	        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	    </tr>
	    <tr valign="top" height="10">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_08.gif"></td>
	        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	    </tr>
	</table>
	<!-- ǥ �ϴܹ� ��-->


<p>

	<% if C_ADMIN_AUTH then %>
	<!-- ǥ ��ܹ� ����-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
		</tr>
		<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	* ������ �޴�
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
	</table>
	<!-- ǥ ��ܹ� ��-->
	
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frmSubmit" method="post" action="do_menuhelpedit.asp">
		<input type="hidden" name="id" value="<%= imenupos.FOneItem.FMenuid %>">
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">�޴���</td>
			<td bgcolor="#FFFFFF"><input type="text" name="menuname" size="50" value="<%= imenupos.FOneItem.FMenuName %>"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">���ü���</td>
			<td bgcolor="#FFFFFF"><input type="text" name="viewidx" size="6" value="<%= imenupos.FOneItem.FViewIndex %>"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ũ</td>
			<td bgcolor="#FFFFFF"><input type="text" name="linkurl" size="50" value="<%= imenupos.FOneItem.FLinkURL %>"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">Color</td>
			<td bgcolor="#FFFFFF"><input type="text" name="menucolor" size="7" value="<%= imenupos.FOneItem.Fmenucolor %>"></td>
		</tr>
	
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">�޴���뿩��</td>
			<td bgcolor="#FFFFFF">
				<% if imenupos.FOneItem.Fisusing="Y" then %>
				<input type="radio" name="isusing" value="Y" checked > �����
				<input type="radio" name="isusing" value="N"> ������
				<% else %>
				<input type="radio" name="isusing" value="Y" > �����
				<input type="radio" name="isusing" value="N" checked > <font color="red">������</font>
				<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">���ܼ���</td>
			<td bgcolor="#FFFFFF">
				<textarea name="menuposnotice" cols="90" rows="8"><%= imenupos.FOneItem.Fmenuposnotice %></textarea>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">�󼼼���</td>
			<td bgcolor="#FFFFFF">
				<textarea name="menuposhelp" cols="90" rows="8"><%= imenupos.FOneItem.Fmenuposhelp %></textarea>
			</td>
		</tr>
		</form>
	</table>
	
	<!-- ǥ �ϴܹ� ����-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	    <tr valign="bottom" height="25">
	        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="bottom" align="center">
	        	<input type="button" value=" �� �� " onclick="SaveMenuEtc(frmSubmit);">
	        </td>
	        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	    </tr>
	    <tr valign="top" height="10">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_08.gif"></td>
	        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	    </tr>
	</table>
	<!-- ǥ �ϴܹ� ��-->
	
	<% end if %>
<% end if %>
</body>
</html>


<%
set imenupos = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->