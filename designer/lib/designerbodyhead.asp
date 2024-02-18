<!-- #include virtual="/lib/classes/admin/menucls.asp"-->


<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script type="text/javascript" src="/js/xl.js"></script>
<script type="text/javascript" src="/js/common.js"></script>
<script type="text/javascript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/calendar.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript">
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_d','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
</head>

<body bgcolor="#F4F4F4">
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="400" style="padding:5; border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>"  background="/images/menubar_1px.gif">
						<font color="#333333">[<%= session("ssBctId") %>]&nbsp;<b><%= imenuposStr %></b></font>
					</td>
					<td align="right" style="border-bottom:1px solid <%= adminColor("tablebg") %>" bgcolor="#F4F4F4">
						<% if imenuposhelp<>"" then %>
						<a href="Javascript:PopMenuHelp('<%= menupos %>');"><img src="/images/icon_help.gif" border="0" valign="bottom"></a>
						<!--
						<input type="button" class="button" value="? 메뉴설명" onclick="PopMenuHelp('<%= menupos %>');" style="cursor:pointer">
						-->
						<% end if %>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	
	<!--	설명 있으면 들어갑니다.	-->
	<% if imenuposnotice<>"" then %>
	<tr bgcolor="#FFFFFF">
		<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
			<%= nl2br(imenuposnotice) %>
		</td>
	</tr>
	<% end if %>
</table>
<div style="padding:5px;"></div>