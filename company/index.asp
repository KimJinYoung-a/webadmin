<html>
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<frameset rows="56,*" frameborder="NO" border="1" framespacing="0" cols="*">
    <frame name="header" scrolling="NO" noresize src="/company/lib/frameheader.asp" >
    <frameset cols="200,*" frameborder="NO" border="1" framespacing="0">
	<% If session("ssBctId") = "accommate" Then %>
		<frame name="menu" noresize src="/company/ch/leftmenu.asp" scrolling="NO">
		<frame name="contents" src="/company/ch/target_itemlist.asp">
	<% ElseIf session("ssBctId") = "nvshop" Then %>
		<frame name="menu" noresize src="/company/nv/leftmenu.asp" scrolling="NO">
		<frame name="contents" src="/company/nv/sellreport.asp">
	<% ElseIf session("ssBctId") = "betweenshop" Then %>
		<frame name="menu" noresize src="/company/between/leftmenu.asp" scrolling="NO">
		<frame name="contents" src="/company/between/sellreport.asp">
	<% ElseIf session("ssBctId") = "daumshop" Then %>
		<frame name="menu" noresize src="/company/dm/leftmenu.asp" scrolling="NO">
		<frame name="contents" src="/company/dm/sellreport.asp">
	<% ElseIf session("ssBctId") = "nateshop" Then %>
		<frame name="menu" noresize src="/company/nt/leftmenu.asp" scrolling="NO">
		<frame name="contents" src="/company/nt/sellreport.asp">
	<% ElseIf session("ssBctId") = "jikbang" Then %>
		<frame name="menu" noresize src="/company/jikbang/leftmenu.asp" scrolling="NO">
		<frame name="contents" src="/company/jikbang/sellreport.asp">
	<% ElseIf session("ssBctId") = "shodoc" Then %>
		<frame name="menu" noresize src="/company/shodoc/leftmenu.asp" scrolling="NO">
		<frame name="contents" src="/company/shodoc/sellreport.asp">
	<% ElseIf (session("ssBctId") = "tingmart") OR (session("ssBctId") = "nanishow") Then %>
    	<frame name="menu" noresize src="/company/menu/leftmenu.asp" scrolling="NO">
    	<frame name="contents" src="/company/notice/notics.asp?menupos=50">
	<% ElseIf session("ssBctId") = "wmprc" Then %>
		<frame name="menu" noresize src="/company/wmprc/leftmenu.asp" scrolling="NO">
		<frame name="contents" src="/company/wmprc/sellreport.asp">
	<% Else %>
    	<frame name="menu" noresize src="/company/menu/leftNewmenu.asp" scrolling="NO">
    	<frame name="contents" src="/company/notice/notics.asp?menupos=50">
	<% End If %>
    </frameset>
    <!-- <frame name="footer" src="/admin/lib/frametailer.asp" noresize scrolling="no"> -->
</frameset>
<noframes>
<body bgcolor="#FFFFFF" text="#000000">
</body></noframes>
</html>