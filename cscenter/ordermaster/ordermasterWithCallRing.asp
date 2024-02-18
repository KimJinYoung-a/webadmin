<%@ language=vbscript %>
<% option Explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<HTML>
<head>
<title>주문내역조회 - with CALL</title>
</head>
<FRAMESET border=1 frameSpacing=0 cols=*,460 scrolling=auto>
<FRAMESET border=1 frameSpacing=0 rows=305,* scrolling=yes>
    <% if (request("orderserial")<>"") then %>
    <FRAME name="listFrame" src="ordermaster_list.asp?orderserial=<%= request("orderserial") %>&searchfield=orderserial" scrolling=auto>
    <% elseif (request("userid")<>"") then %>
	<FRAME name="listFrame" src="ordermaster_list.asp?userid=<%= request("userid") %>&searchfield=userid" scrolling=auto>
	<% else %>
	<FRAME name="listFrame" src="ordermaster_list.asp" scrolling=auto>
	<% end if %>
	<FRAME name="detailFrame" src="ordermaster_detail.asp" scrolling=auto>
</FRAMESET>
<FRAME name="callring" src="/cscenter/ippbxmng/CallRingWithOrderFrame.asp?ippbxuser=<%= request("ippbxuser") %>&intel=<%= request("intel") %>&phoneNumber=<%= request("phoneNumber") %>&id=<%= request("id") %>" scrolling=yes>
</FRAMESET>
</HTML>
