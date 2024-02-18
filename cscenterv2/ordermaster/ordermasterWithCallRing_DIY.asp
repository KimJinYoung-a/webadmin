<!-- #include virtual="/lib/util/htmllib.asp"-->
<HTML>
<head>
<title>주문내역조회(DIY)</title>
</head>
<FRAMESET border=1 frameSpacing=0 cols=*,460 scrolling=auto>
<FRAMESET border=1 frameSpacing=0 rows=305,* scrolling=yes>
    <% if (requestCheckVar(request("orderserial"),16)<>"") then %>
    <FRAME name="listFrame" src="/cscenterv2/order/ordermaster_list.asp?orderserial=<%= requestCheckVar(request("orderserial"),16) %>&searchfield=orderserial" scrolling=auto>
    <% elseif (requestCheckVar(request("userid"),32)<>"") then %>
	<FRAME name="listFrame" src="/cscenterv2/order/ordermaster_list.asp?userid=<%= requestCheckVar(request("userid"),32) %>&searchfield=userid" scrolling=auto>
	<% else %>
	<FRAME name="listFrame" src="/cscenterv2/order/ordermaster_list.asp" scrolling=auto>
	<% end if %>
	<FRAME name="detailFrame" src="/cscenterv2/order/orderdetail_view.asp" scrolling=auto>
</FRAMESET>
<FRAME name="callring" src="/cscenterv2/order/CallRingWithOrderFrame.asp?ippbxuser=<%= requestCheckVar(request("ippbxuser"),32) %>&intel=<%= requestCheckVar(request("intel"),32) %>&phoneNumber=<%= requestCheckVar(request("phoneNumber"),16) %>&id=<%= requestCheckVar(request("id"),32) %>" scrolling=yes>
</FRAMESET>
</HTML>
