<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.09 한용민 생성
'###########################################################
%>
<HTML>
<head>
<title>CS LIST</title>
</head>

<FRAMESET border=1 frameSpacing=0 rows=325,* scrolling=yes>
	<FRAME name="listFrame" src="cs_action_list.asp?masteridx=<%=request("masteridx")%>&searchtype=<%=request("searchtype")%>&searchfield=<%=request("searchfield")%>&searchstring=<%=request("searchstring")%>&divcd=<%=request("divcd")%>&currstate=<%=request("currstate")%>&delYN=<%=request("delYN")%>" scrolling=no>
	<FRAME name="detailFrame" src="cs_action_detail.asp" scrolling=yes>
</FRAMESET>

</HTML>
