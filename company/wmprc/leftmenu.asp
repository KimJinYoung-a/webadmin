<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/admin/OLDmenucls.asp"-->
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
	// 기본옵션 지정
	USETEXTLINKS = 1
	STARTALLOPEN = 0
	HIGHLIGHT = 1
	PRESERVESTATE = 1
	GLOBALTARGET="R"

	// 루트메뉴
	foldersTree = gFld('Admin', '')
	
	// 하위메뉴
	a2 = gFld('&nbsp;매출관리', '')
	a2.xID='f2'
	a2.addChildren([['매출집계', '/company/wmprc/sellreport.asp', '#000000']])

	foldersTree.addChildren([a2])

	foldersTree.treeID = "L1" 
	foldersTree.xID = "bigtree"

</SCRIPT>
</head>
<body topmargin="0">
<table width="100%" height="100%" border="0" cellSpacing="0" cellPadding="0">
<tr>
<td valign="top">
		<SCRIPT>
			// 메뉴 출력 실행
			initializeDocument();
		</SCRIPT>
</td>
<td width="1" bgcolor="#DDDDDD" ></td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->