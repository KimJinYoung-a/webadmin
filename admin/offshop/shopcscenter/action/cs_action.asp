<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
'###########################################################
%>

<HTML>
<head>
<title>CS LIST</title>
</head>

<FRAMESET border=1 frameSpacing=0 rows=325,* scrolling=yes>
	<FRAME name="listFrame" src="cs_action_list.asp?orderno=<%=request("orderno")%>&masteridx=<%=request("masteridx")%>&searchtype=<%=request("searchtype")%>&searchfield=<%=request("searchfield")%>&searchstring=<%=request("searchstring")%>&divcd=<%=request("divcd")%>&currstate=<%=request("currstate")%>&delYN=<%=request("delYN")%>&shopid=<%=request("shopid")%>" scrolling=no>
	<FRAME name="detailFrame" src="cs_action_detail.asp" scrolling=yes>
</FRAMESET>

</HTML>
