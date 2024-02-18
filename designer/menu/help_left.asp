<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/OLDmenucls.asp"-->
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
<body topmargin="0" >
<%
dim menupos, udiv
menupos = requestCheckVar(request("menupos"),10)
udiv = requestCheckVar(request("udiv"),20)

if udiv="" then udiv="9999"

dim allMenuItem,i,j

set allMenuItem = new CMenu
allMenuItem.FrectUsingOnly="Y"
allMenuItem.getMenuItems udiv

dim url
dim parentmenuid
%>
<script language='javascript'>
function ResearchMenu(comp){
	document.reloadfrm.submit();
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<form name="reloadfrm" method=get >
<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	도움말
				<% if C_ADMIN_AUTH then %>
				<select name="udiv" onChange="ResearchMenu(this)">
				<option value="9999" <% if udiv="9999" then response.write "selected" %> > 업체(9999)
				<option value="999" <% if udiv="999" then response.write "selected" %> > 제휴사(999)
				<option value="9" <% if udiv="9" then response.write "selected" %> > 관리자(9)
				<option value="501" <% if udiv="501" then response.write "selected" %> > 직영점(501)
				<option value="502" <% if udiv="502" then response.write "selected" %> > 가맹점(502)
				<option value="503" <% if udiv="503" then response.write "selected" %> > 기타매장(503)
				<option value="101" <% if udiv="101" then response.write "selected" %> > 오프샾(101)
				<option value="301" <% if udiv="301" then response.write "selected" %> > 컬리지(301)
				</select>
				<% end if %>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</form>

<!-- 표 상단바 끝-->



<tr>
	<td background="/images/tbl_blue_round_04.gif"></td>
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
			if (url<>"") and (url<>"#") then url = "/designer/menu/help_contents.asp" + "?menupos=" + CStr(allMenuItem.FMenuitemlist(i).FMenuID)
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
				if (url<>"") and (url<>"#") then url = "/designer/menu/help_contents.asp" + "?menupos=" + CStr(allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuID)

				if CStr(allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuID)=menupos then parentmenuid = allMenuItem.FMenuitemlist(i).FMenuID
			%>
			<DIV onselectstart="return false" id=g<%= allMenuItem.FMenuitemlist(i).FMenuID %> ondragstart="return false" style="DISPLAY: none; CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" open="false" imageOpen="/images/paper2.gif" image="/images/paper2.gif" url="<%= url %>">
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
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>


<!-- 표 하단바 시작-->
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
<!-- 표 하단바 끝-->
<script language='javascript'>


window.onload = getOnload;
function getOnload(){
	var entity = document.all.f<%= parentmenuid %>;

	if (entity){
		expand(entity, true);
	}
}
</script>
</body>
</html>
<%
set allMenuItem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->