<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/menucls.asp"-->
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
</head>
<body topmargin="0">
<%
dim allMenuItem,i,j

set allMenuItem = new CMenu
allMenuItem.FrectUsingOnly="Y"
allMenuItem.getMenuItems 9999

dim url
%>
<table width="100%" height="100%" border="0" cellSpacing="0" cellPadding="0">
<tr>
<td valign="top">
<DIV id="folderTree" STYLE="padding-top: 8px;">
	<TABLE cellSpacing=0 cellPadding=0 border=0>
		<TBODY>
		<TR>
			<TD valign="center"><IMG src="/images/mycomputer.gif"></TD>
			<TD valign="center">&nbsp;Admin</TD>
		</TR>
		</TBODY>
	</TABLE>
	<% for i=0 to allMenuItem.FMenuCount-1 %>
	<%
		url = allMenuItem.FMenuitemlist(i).getLinkURL
		if (url<>"") and (url<>"#") then url = url + "?menupos=" + CStr(allMenuItem.FMenuitemlist(i).FMenuID)
	%>
	<DIV onselectstart="return false" id=f<%= allMenuItem.FMenuitemlist(i).FMenuID %> ondragstart="return false" style="CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" xmlns:dt="urn:schemas-microsoft-com:datatypes" open="false" imageOpen="<%= allMenuItem.FMenuitemlist(i).getOpenIconURL %>" image="<%= allMenuItem.FMenuitemlist(i).getCloseIconURL %>" url="<%= url %>">
		<TABLE cellSpacing=0 cellPadding=0 border=0>
			<TBODY>
			<TR>
				<TD><IMG id=stateImagef800 src="<%= allMenuItem.FMenuitemlist(i).getCloseTimageUrl %>" _closed="<%= allMenuItem.FMenuitemlist(i).getCloseTimageUrl %>" _open="<%= allMenuItem.FMenuitemlist(i).getOpenTimageUrl %>"></TD>
				<TD vAlign=center><IMG id=image src="<%= allMenuItem.FMenuitemlist(i).getCloseIconURL %>" border=0></TD>
				<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = '<%= allMenuItem.FMenuitemlist(i).FMenuColor %>'; this.style.fontWeight = 'normal';" vAlign=center noWrap><font color="<%= allMenuItem.FMenuitemlist(i).FMenuColor %>"><%= allMenuItem.FMenuitemlist(i).FMenuName %></font></TD>
			</TR>
			</TBODY>
		</TABLE>
		<% if allMenuItem.FMenuitemlist(i).IsHasChild then %>
		<% for j=0 to allMenuItem.FMenuitemlist(i).FChildCount-1  %>
		<%
			url = allMenuItem.FMenuitemlist(i).FChildItem(j).getLinkURL
			if (url<>"") and (url<>"#") then url = url + "?menupos=" + CStr(allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuID)
		%>
		<DIV onselectstart="return false" id=f603 ondragstart="return false" style="DISPLAY: none; CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" open="false" imageOpen="/images/paper2.gif" image="/images/paper2.gif" url="<%= url %>">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<TBODY>
				<TR>
					<TD><IMG src="<%= allMenuItem.FMenuitemlist(i).getCloseIimageUrl %>"><IMG src="<%= allMenuItem.FMenuitemlist(i).FChildItem(j).getOpenTimageUrl %>"></TD>
					<TD vAlign=center><IMG id=image src="/images/paper2.gif" border=0></TD>
					<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: <%= allMenuItem.FMenuitemlist(i).FMenuColor %>" onmouseout="this.style.color = '<%= allMenuItem.FMenuitemlist(i).FMenuColor %>'; this.style.fontWeight = 'normal';" vAlign=center noWrap><font color="<%= allMenuItem.FMenuitemlist(i).FMenuColor %>"><%= allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuname %></font></TD>
				</TR>
				</TBODY>
			</TABLE>
		</DIV>
		<% next %>
		<% end if %>
	</DIV>
	<% next %>

</DIV>
</td>
<td width="1" bgcolor="#DDDDDD" ></td>
</tr>
</table>
</body>
</html>
<%
set allMenuItem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->