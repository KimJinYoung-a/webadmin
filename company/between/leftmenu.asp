<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/admin/OLDmenucls.asp"-->
<%
response.charset = "euc-kr"
%>
<html>
<head>
<SCRIPT language="javascript" SRC="/js/jsTree_new.js"></SCRIPT>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none;color:#000000;}
-->
</style>
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
	foldersTree.addChildren([['��������', '/company/between/sellreport.asp', '#000000'], ['��ǰ��ȸ', '/company/between/itemlist.asp', '#000000']])
	foldersTree.treeID = "L1" 
	foldersTree.xID = "bigtree"

</SCRIPT>
</head>
<body topmargin="0">
<table width="100%" height="100%" border="0" cellSpacing="0" cellPadding="0">
<tr>
<td valign="top">
		<SCRIPT>
			// �޴� ��� ����
			initializeDocument();
		</SCRIPT>
</td>
<td width="1" bgcolor="#DDDDDD" ></td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->