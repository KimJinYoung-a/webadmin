<%@ language=vbscript %>
<% option explicit %>
<%

dim startdt, enddt
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2

startdt = request("startdt")
enddt = request("enddt")

if Len(startdt) = 10 then
    startdt = "&yyyy1=" & Left(startdt, 4) & "&mm1=" & Right(Left(startdt, 7), 2) & "&dd1=" & Right(startdt, 2)
else
    startdt = ""
end if

if Len(enddt) = 10 then
    enddt = "&yyyy2=" & Left(enddt, 4) & "&mm2=" & Right(Left(enddt, 7), 2) & "&dd2=" & Right(enddt, 2)
else
    enddt = ""
end if

%>
<HTML>
<head>
<title>CS LIST</title>
</head>

<FRAMESET border=1 frameSpacing=0 rows=325,* scrolling=yes>
	<FRAME name="listFrame" src="cs_action_list.asp?orderserial=<%=request("orderserial")%>&searchtype=<%=request("searchtype")%>&searchfield=<%=request("searchfield")%>&searchstring=<%=request("searchstring")%>&divcd=<%=request("divcd")%>&currstate=<%=request("currstate")%>&delYN=<%=request("delYN")%><%= startdt %><%= enddt %>" scrolling=no>
	<FRAME name="detailFrame" src="cs_action_detail.asp" scrolling=yes>
</FRAMESET>

</HTML>
