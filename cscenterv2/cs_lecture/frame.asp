<!-- #include virtual="/lib/util/htmllib.asp"-->
<HTML>
<head>
<title>CS LIST</title>
</head>

<FRAMESET border=1 frameSpacing=0 rows=325,* scrolling=yes>
	<FRAME name="listFrame" src="lec_csmaster_list.asp?orderserial=<%=RequestCheckvar(request("orderserial"),16)%>&searchtype=<%=RequestCheckvar(request("searchtype"),16)%>&searchfield=<%=RequestCheckvar(request("searchfield"),16)%>&searchstring=<%=RequestCheckvar(request("searchstring"),32)%>&divcd=<%=RequestCheckvar(request("divcd"),4)%>&currstate=<%=RequestCheckvar(request("currstate"),4)%>&delYN=<%=RequestCheckvar(request("delYN"),2)%>" scrolling=no>
	<FRAME name="detailFrame" src="lec_csdetail_view.asp" scrolling=yes>
</FRAMESET>

</HTML>
