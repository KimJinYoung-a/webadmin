<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/admin/OLDmenucls.asp"-->
<html>
<head>
<SCRIPT LANGUAGE="Javascript" SRC="/js/xmlTree.js"></SCRIPT>
<script language="JavaScript" src="/cscenter/js/jquery-1.8.3.js"></script>
<script language="JavaScript" src="/cscenter/js/jquery-ui-1.9.2.min.js"></script>
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
allMenuItem.getMenuItems 999

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
	<% if session("ssBctId")="tingmart" then %>
	<DIV onselectstart="return false" id=ftingmart ondragstart="return false" style="CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" xmlns:dt="urn:schemas-microsoft-com:datatypes" open="false" imageOpen="/images/openfolder.png" image="/images/closedfolder.png" url="">
		<TABLE cellSpacing=0 cellPadding=0 border=0>
			<TBODY>
			<TR>
				<TD><IMG id=stateImagef800 src="/images/Lplus.png" _closed="/images/Lplus.png" _open="/images/Lminus.png"></TD>
				<TD vAlign=center><IMG id=image src="/images/closedfolder.png" border=0></TD>
				<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = 'black'; this.style.fontWeight = 'normal';" vAlign=center noWrap>팅마트전용</TD>
			</TR>
			</TBODY>
		</TABLE>
		<DIV onselectstart="return false" id=f603 ondragstart="return false" style="DISPLAY: none; CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" open="false" imageOpen="/images/paper2.gif" image="/images/paper2.gif" url="/company/tingmart/boardlist.asp">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<TBODY>
				<TR>
					<TD><IMG src="/images/blank.png"><IMG src="/images/L.png"></TD>
					<TD vAlign=center><IMG id=image src="/images/paper2.gif" border=0></TD>
					<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = 'black'; this.style.fontWeight = 'normal';" vAlign=center noWrap>질문답변</TD>
				</TR>
				</TBODY>
			</TABLE>
		</DIV>
	</DIV>
	<% else %>
	<% for i=0 to allMenuItem.FMenuCount-1 %> 
	<%
		url = allMenuItem.FMenuitemlist(i).getLinkURL
		if (url<>"") and (url<>"#") then url = url + "?menupos=" + CStr(allMenuItem.FMenuitemlist(i).FMenuID)
	%>
	<DIV onselectstart="return false" id=f<%= allMenuItem.FMenuitemlist(i).FMenuID %> ondragstart="return false" style="CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" xmlns:dt="urn:schemas-microsoft-com:datatypes" open="false" imageOpen="<%= allMenuItem.FMenuitemlist(i).getOpenIconURL %>" image="<%= allMenuItem.FMenuitemlist(i).getCloseIconURL %>" url="<%= url %>">
		<TABLE cellSpacing=0 cellPadding=0 border=0>
			<TBODY>
			<TR>
				<TD><IMG id=stateImagef<%= allMenuItem.FMenuitemlist(i).FMenuID %> src="<%= allMenuItem.FMenuitemlist(i).getCloseTimageUrl %>" _closed="<%= allMenuItem.FMenuitemlist(i).getCloseTimageUrl %>" _open="<%= allMenuItem.FMenuitemlist(i).getOpenTimageUrl %>"></TD>
				<TD vAlign=center><IMG id=imagef<%= allMenuItem.FMenuitemlist(i).FMenuID %> src="<%= allMenuItem.FMenuitemlist(i).getCloseIconURL %>" border=0></TD>
				<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = 'black'; this.style.fontWeight = 'normal';" vAlign=center noWrap><%= allMenuItem.FMenuitemlist(i).FMenuName %></TD>
			</TR>
			</TBODY>
		</TABLE>
		<% if allMenuItem.FMenuitemlist(i).IsHasChild then %>
		<% for j=0 to allMenuItem.FMenuitemlist(i).FChildCount-1  %>
		<%
			url = allMenuItem.FMenuitemlist(i).FChildItem(j).getLinkURL
			if (url<>"") and (url<>"#") then url = url + "?menupos=" + CStr(allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuID)
		%>
		<DIV onselectstart="return false" id=f<%= allMenuItem.FMenuitemlist(i).FMenuID %> ondragstart="return false" style="DISPLAY: none; CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" open="false" imageOpen="/images/paper2.gif" image="/images/paper2.gif" url="<%= url %>">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<TBODY>
				<TR>
					<TD><IMG src="<%= allMenuItem.FMenuitemlist(i).getCloseIimageUrl %>"><IMG src="<%= allMenuItem.FMenuitemlist(i).FChildItem(j).getOpenTimageUrl %>"></TD>
					<TD vAlign=center><IMG id=imagef<%= allMenuItem.FMenuitemlist(i).FMenuID %> src="/images/paper2.gif" border=0></TD>
					<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = 'black'; this.style.fontWeight = 'normal';" vAlign=center noWrap><%= allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuname %></TD>
				</TR>
				</TBODY>
			</TABLE>
		</DIV>
		<% next %>
		<% end if %>
	</DIV>
	<% next %>

	<% if session("ssBctId")="nanishow" then %>
	<DIV onselectstart="return false" id=fnanishow ondragstart="return false" style="CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" xmlns:dt="urn:schemas-microsoft-com:datatypes" open="false" imageOpen="/images/openfolder.png" image="/images/closedfolder.png" url="">
		<TABLE cellSpacing=0 cellPadding=0 border=0>
			<TBODY>
			<TR>
				<TD><IMG id=stateImagef800 src="/images/Lplus.png" _closed="/images/Lplus.png" _open="/images/Lminus.png"></TD>
				<TD vAlign=center><IMG id=image src="/images/closedfolder.png" border=0></TD>
				<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = 'black'; this.style.fontWeight = 'normal';" vAlign=center noWrap>나니쇼샾전용</TD>
			</TR>
			</TBODY>
		</TABLE>
		<DIV onselectstart="return false" id=f603 ondragstart="return false" style="DISPLAY: none; CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" open="false" imageOpen="/images/paper2.gif" image="/images/paper2.gif" url="http://www.10x10.co.kr/ext/nanishow/naniadmini/noticsadmin.asp">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<TBODY>
				<TR>
					<TD><IMG src="/images/blank.png"><IMG src="/images/T.png"></TD>
					<TD vAlign=center><IMG id=image src="/images/paper2.gif" border=0></TD>
					<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = 'black'; this.style.fontWeight = 'normal';" vAlign=center noWrap>공지사항관리</TD>
				</TR>
				</TBODY>
			</TABLE>
		</DIV>
		<DIV onselectstart="return false" id=f603 ondragstart="return false" style="DISPLAY: none; CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" open="false" imageOpen="/images/paper2.gif" image="/images/paper2.gif" url="http://www.10x10.co.kr/ext/nanishow/naniadmini/choicemonthadmin.asp">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<TBODY>
				<TR>
					<TD><IMG src="/images/blank.png"><IMG src="/images/T.png"></TD>
					<TD vAlign=center><IMG id=image src="/images/paper2.gif" border=0></TD>
					<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = 'black'; this.style.fontWeight = 'normal';" vAlign=center noWrap>메인페이지</TD>
				</TR>
				</TBODY>
			</TABLE>
		</DIV>

		<DIV onselectstart="return false" id=f603 ondragstart="return false" style="DISPLAY: none; CURSOR: hand" onclick="window.event.cancelBubble = true;clickOnEntity(this);" open="false" imageOpen="/images/paper2.gif" image="/images/paper2.gif" url="/company/nanishow/bestnaniadmin.asp">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<TBODY>
				<TR>
					<TD><IMG src="/images/blank.png"><IMG src="/images/T.png"></TD>
					<TD vAlign=center><IMG id=image src="/images/paper2.gif" border=0></TD>
					<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = 'black'; this.style.fontWeight = 'normal';" vAlign=center noWrap>BestNaniShow</TD>
				</TR>
				</TBODY>
			</TABLE>
		</DIV>
		<DIV onselectstart="return false" id=f603 ondragstart="return false" style="DISPLAY: none; CURSOR: hand" onclick="NnPopAdm()" open="false" imageOpen="/images/paper2.gif" image="/images/paper2.gif" url="">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
				<TBODY>
				<TR>
					<TD><IMG src="/images/blank.png"><IMG src="/images/L.png"></TD>
					<TD vAlign=center><IMG id=image src="/images/paper2.gif" border=0></TD>
					<TD onmouseover="this.style.color = '#FF0000'; this.style.fontWeight = 'bold'; this.style.fontFamily = 'Tahoma'" style="PADDING-LEFT: 7px; FONT-WEIGHT: normal; FONT-SIZE: 9pt; COLOR: black; FONT-FAMILY: Tahoma; font-color: black" onmouseout="this.style.color = 'black'; this.style.fontWeight = 'normal';" vAlign=center noWrap>커뮤니티Admin</TD>
				</TR>
				</TBODY>
			</TABLE>
		</DIV>
		<script language='javascript'>
		    function NnPopAdm(){
		    	var popwin;
		    	popwin = window.open('','naniadmin');
		    	document.NnFrm.target = 'naniadmin';
		    	document.NnFrm.submit();
		    }
		</script>
		<form name="NnFrm" method="post" action="http://www.10x10.co.kr/ext/nanishow/extlogin.asp" >
			<input type="hidden" name="NnUid" value="<%= session("ssBctId") %>">
		</form>
	</DIV>
	<% end if %>
	<% end if %>
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