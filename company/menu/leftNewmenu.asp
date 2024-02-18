<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/admin/OLDmenucls.asp"-->
<html>
<head>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link rel="StyleSheet" href="/css/dtree.css" type="text/css" />
<link rel="stylesheet" href="/js/jqueryui/css/jquery-ui.css">
<% '<link rel="stylesheet" href="//code.jquery.com/ui/1.11.1/themes/smoothness/jquery-ui.css"> %>
<script language="JavaScript" src="/cscenter/js/jquery-1.8.3.js"></script>
<script language="JavaScript" src="/cscenter/js/jquery-ui-1.9.2.min.js"></script>
<script language="javascript" src="/js/dtree.js"></script>
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none}
-->
</style>
</head>
<body topmargin="0">
<%
Dim allMenuItem,i,j
Dim strFTree, strColor, tmpMenuName
Dim url, admSelPosit, admSelLevel
Dim searchString
set allMenuItem = new CMenu
	allMenuItem.FrectUsingOnly="Y"
	allMenuItem.getMenuNewItems
%>
<table width="100%" border="0" cellSpacing="0" cellPadding="0">
<tr>
	<td valign="top">
		<script type="text/javascript">
		var menuAllNameArray = new Array(<%= allMenuItem.FResultCount %>);
		var menuAll = new dTree("menuAll");
		menuAll.config.useCookies = false;
		menuAll.add(0,-1,"Admin");
		<%
		For i=0 to allMenuItem.FResultCount - 1
			url = allMenuItem.FMenuitemlist(i).FLinkURL
			If IsNull(url) Then
				url = ""
			End if

			If (url <> "") Then
				If instr(url,"?")>0 then
					url=url & "&menupos=" & allMenuItem.FMenuitemlist(i).FMenuID
				Else
					url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).FMenuID
				End if
			End If

			tmpMenuName = allMenuItem.FMenuitemlist(i).FMenuName
			%>menuAll.add(<%= allMenuItem.FMenuitemlist(i).FMenuID %>, <%= allMenuItem.FMenuitemlist(i).FParentID %>, "<%= tmpMenuName %>", "<%= url %>"); <%
			If (url <> "") Then
				%>menuAllNameArray[<%= i %>] = "<%= allMenuItem.FMenuitemlist(i).FMenuName %>"; <%
			Else
				%>menuAllNameArray[<%= i %>] = "XXX"; <%
			End if
		Next
		%>

		document.write(menuAll);

		var menuAllNameArrayUniq = [];
		$.each(menuAllNameArray, function(i, el){
			if($.inArray(el, menuAllNameArrayUniq) === -1) menuAllNameArrayUniq.push(el);
		});
		</script>
		<br>
	</td>
</tr>
</table>
</body>
</html>
<% set allMenuItem = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->