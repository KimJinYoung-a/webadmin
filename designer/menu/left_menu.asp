<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/OLDmenucls.asp"-->
<%
	dim allMenuItem,i,j, strFTree, strColor

	set allMenuItem = new CMenu
	allMenuItem.FrectUsingOnly="Y"
	allMenuItem.getMenuItems 9999

	dim url

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none;color:#000000;}
-->
</style>
<SCRIPT language="javascript" SRC="/js/jsTree_new.js"></SCRIPT>
<SCRIPT language="javascript">
	// �⺻�ɼ� ����
	USETEXTLINKS = 1
	STARTALLOPEN = 0
	HIGHLIGHT = 1
	PRESERVESTATE = 1
	GLOBALTARGET="R"

	// ��Ʈ�޴�
	foldersTree = gFld('Admin', '')

	// �����޴�
<%
		for i=0 to allMenuItem.FMenuCount-1
			if allMenuItem.FMenuitemlist(i).IsHasChild then
				'���� ����
				url = allMenuItem.FMenuitemlist(i).getLinkURL
				if url="#" then url=""
				if Not(url="" or isNull(url)) then url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).FMenuID
				strColor = allMenuItem.FMenuitemlist(i).FMenuColor
				if Not(strColor="" or isNull(strColor)) then strColor= ", '" & strColor & "'": else strColor="": end if
				Response.Write "a" & i & " = gFld('&nbsp;" & allMenuItem.FMenuitemlist(i).FMenuName & "', '" & url & "'" & strColor & ")" & vbCrLf
				Response.Write "a" & i & ".xID='f" & i & "'" & vbCrLf

				'���� ����ǥ��
				Response.Write "a" & i & ".addChildren(["
				for j=0 to allMenuItem.FMenuitemlist(i).FChildCount-1
				url = allMenuItem.FMenuitemlist(i).FChildItem(j).getLinkURL
				if url="#" then url=""
				strColor = allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuColor
				if Not(strColor="" or isNull(strColor)) then strColor= ", '" & strColor & "'": else strColor="": end if
				if Not(url="" or isNull(url)) then url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuID
					Response.Write "['" & allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuName & "', '" & url & "'" & strColor & "]"

					'������ ǥ��
					if j<allMenuItem.FMenuitemlist(i).FChildCount-1 then
						Response.Write ", "
					end if
				next
				'���� ��ǥ��
				Response.Write "])" & vbCrLf & vbCrLf

				strFTree = strFTree & "a" & i
			else
				'���� ����
				url = allMenuItem.FMenuitemlist(i).getLinkURL
				if url="#" then url=""
				if Not(url="" or isNull(url)) then url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).FMenuID
				strFTree = strFTree & "['&nbsp;" & allMenuItem.FMenuitemlist(i).FMenuName & "', '" & url & "', '" & allMenuItem.FMenuitemlist(i).FMenuColor & "']"
			end if

			if i<allMenuItem.FMenuCount-1 then
				strFTree = strFTree & ", "
			end if
		next

		'�ֻ����� ���� �޴� �߰�
		Response.Write vbCrLf & "foldersTree.addChildren([" & strFTree & "])" & vbCrLf
	%>
	foldersTree.treeID = "L1"
	foldersTree.xID = "bigtree"

</SCRIPT>
</head>

<body topmargin="15" leftmargin=3  bgcolor="#F4F4F4">
<script language='javascript'>
<!--
	function PopMenuHelp(menupos){
		var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'admin_PopMenuHelp_d','width=800, height=600, scrollbars=yes,resizable=yes');
		popwin.focus();
	}
//-->
</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td valign="top">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("menubar_left") %>">
					<td>
						<b>MENU</b>
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td>
						<SCRIPT>
							// �޴� ��� ����
							initializeDocument();
						</SCRIPT>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="10">
		<td></td>
	</tr>
	<!--
	<tr>
		<td valign="top">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("menubar_left") %>">
						<td>
							<b>��Ÿ����</b>
						</td>
					</tr>
				<tr bgcolor="#FFFFFF">
					<td>
						<img src="/images/icon_num01.gif">&nbsp;�����籸�ż���<br>
						<img src="/images/icon_num02.gif">&nbsp;�����籸�ż���<br>
						<img src="/images/icon_num03.gif">&nbsp;�����籸�ż���
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="10">
		<td></td>
	</tr>
	<tr>
		<td valign="top">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("menubar_left") %>">
						<td>
							<b>HELP</b>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size="3">>>more</font>
						</td>
					</tr>
				<tr bgcolor="#FFFFFF">
					<td>
						<img src="/images/icon_star.gif">&nbsp;���ݰ�꼭������<br>
						<img src="/images/icon_star.gif">&nbsp;��ü�������ó�����<br>
						<img src="/images/icon_star.gif">&nbsp;������
					</td>
				</tr>
			</table>
		</td>
	</tr>
	-->
	<tr>
		<td>
			<div style="margin-top:25px; padding:12px 5px; text-align:center; background-color:#e8e8e8; font-family:'malgun Gothic','�������', Dotum, '����', sans-serif; border:1px dashed #ddd">
				<strong style="font-size:13px;">��Ʈ�� ���� ������</strong><br /><span style="font-size:11px; color:#666;">(���ֹ� ���� ����)</span>
				<div style="background-color:#fff; padding:10px; margin-top:10px;">
					<strong style="font-family:'malgun Gothic','�������', Dotum, '����', sans-serif; font-size:18px; color:#00cccc; text-shadow:1px 1px rgba(0,51,51,0.4);">070-4868-1799</strong>
					<table style="width:97%; font-size:11px; color:#999; margin:10px auto 0 auto; line-height:13px;">
						<tr>
							<th style="text-align:left;">����</th>
							<td style="text-align:right;">09:00 ~ 06:00</td>
						</tr>
						<tr>
							<th style="text-align:left;">���ɽð�</th>
							<td style="text-align:right;">12:00 ~ 01:00</td>
						</tr>
						<tr>
							<td colspan="2"  style="text-align:center;">��/�Ϥ������� �޹�</td>
						</tr>
					</table>
				</div>
			</div>
		</td>
	</tr>
	<tr>
		<td style="padding-top:12px"><a href="http://webadmin.10x10.co.kr/partner/index.asp" target="_blank"><img src="http://webadmin.10x10.co.kr/images/partner/partner_btn_newver.png" alt="�� ���� �ٷΰ���" border="0" /></a></td>
	</tr>
</table>

</body>
</html>
<%
set allMenuItem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
