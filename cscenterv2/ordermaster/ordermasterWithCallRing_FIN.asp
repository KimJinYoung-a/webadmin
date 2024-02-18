<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim sitename, ordermasterURL, orderdetailURL

sitename = requestCheckVar(request("sitename"),32)


ordermasterURL = "/cscenterv2/order/ordermaster_list.asp"
orderdetailURL = "/cscenterv2/order/orderdetail_view.asp"

if (sitename <> "diyitem") then
	sitename = "academy"
	ordermasterURL = "/cscenterv2/lecture/lecturemaster_list.asp"
	orderdetailURL = "/cscenterv2/lecture/lecturedetail_view.asp"
end if

%>
<HTML>
<head>
	<title>주문내역조회(<%= sitename %>)</title>
</head>
<FRAMESET border=1 frameSpacing=0 cols=*,460 scrolling=auto>
<FRAMESET border=1 frameSpacing=0 rows=305,* scrolling=yes>
    <% if (request("orderserial")<>"") then %>
    <FRAME name="listFrame" src="<%= ordermasterURL %>?orderserial=<%= request("orderserial") %>&searchfield=orderserial" scrolling=auto>
    <% elseif (request("userid")<>"") then %>
	<FRAME name="listFrame" src="<%= ordermasterURL %>?userid=<%= request("userid") %>&searchfield=userid" scrolling=auto>
	<% else %>
	<FRAME name="listFrame" src="<%= ordermasterURL %>" scrolling=auto>
	<% end if %>
	<FRAME name="detailFrame" src="<%= orderdetailURL %>" scrolling=auto>
</FRAMESET>
<FRAME name="callring" src="/cscenterv2/order/CallRingWithOrderFrame.asp?sitename=<%= sitename %>&ippbxuser=<%= request("ippbxuser") %>&intel=<%= request("intel") %>&phoneNumber=<%= request("phoneNumber") %>&id=<%= request("id") %>" scrolling=yes>
</FRAMESET>
</HTML>
