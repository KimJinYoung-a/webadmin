<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/MenuCls.asp"-->
<%
dim allMenuItem,i,j, strFTree, strColor , shopdiv

shopdiv = getoffshopdiv(session("SSBctID"))

set allMenuItem = new CMenuList

	if (session("ssBctDiv")=509) then
	    allMenuItem.FRectUserDiv = "509"
	else
	    allMenuItem.FRectUserDiv = session("ssBctDiv")
	end if

	'//���� ǥ�� ��� ó��
	if getadmindisplaylanguage(session("SSBctID")) <> "KOR" then
	    allMenuItem.FRectUsingEnMenuName="on"
	end if

	allMenuItem.FRectshopdiv = shopdiv
	allMenuItem.GetLeftMenuList_offLine

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
				url = allMenuItem.FMenuitemlist(i).Fmenu_linkurl
				if Not(url="" or isNull(url)) then url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).Fmenu_id
				strColor = allMenuItem.FMenuitemlist(i).Fmenu_color
				if Not(strColor="" or isNull(strColor)) then strColor= ", '" & strColor & "'": else strColor="": end if
				Response.Write "a" & i & " = gFld('&nbsp;" & allMenuItem.FMenuitemlist(i).Fmenu_name & "', '" & url & "'" & strColor & ")" & vbCrLf
				Response.Write "a" & i & ".xID='f" & i & "'" & vbCrLf

				'���� ����ǥ��
				Response.Write "a" & i & ".addChildren(["
				for j=0 to allMenuItem.FMenuitemlist(i).FChildCount-1
				url = allMenuItem.FMenuitemlist(i).FChildItem(j).Fmenu_linkurl
				strColor = allMenuItem.FMenuitemlist(i).FChildItem(j).Fmenu_color
				if Not(strColor="" or isNull(strColor)) then strColor= ", '" & strColor & "'": else strColor="": end if
				if Not(url="" or isNull(url)) then url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).FChildItem(j).Fmenu_id
					Response.Write "['" & allMenuItem.FMenuitemlist(i).FChildItem(j).Fmenu_name & "', '" & url & "'" & strColor & "]"

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
				url = allMenuItem.FMenuitemlist(i).Fmenu_linkurl
				if Not(url="" or isNull(url)) then url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).Fmenu_id
				strFTree = strFTree & "['&nbsp;" & allMenuItem.FMenuitemlist(i).Fmenu_name & "', '" & url & "', '" & allMenuItem.FMenuitemlist(i).Fmenu_color & "']"
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
<body topmargin="0" leftmargin=0>
<script language='javascript'>
<!--
	function PopMenuHelp(menupos){
		var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'admin_PopMenuHelp_d','width=800, height=600, scrollbars=yes,resizable=yes');
		popwin.focus();
	}
//-->
</script>


<table width="100%" border="0" cellSpacing="0" cellPadding="0">
	<tr>
		<td valign="top">
			<img src="/images/icon_help.gif" width="50" height="20" onclick="PopMenuHelp('');" style="cursor:pointer">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<SCRIPT>
				// �޴� ��� ����
				initializeDocument();
			</SCRIPT>
		</td>
	</tr>
</table>
</body>
</html>
<%
set allMenuItem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
