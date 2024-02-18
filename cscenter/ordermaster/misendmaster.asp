<%

response.write "사용중지메뉴 : 관리자 문의 요망."
response.end

'response.redirect("newmisendlist.asp?notincludeupchecheck=on&delaydate=4")
'dbget.close()	:	response.End
%>



<HTML>
<FRAMESET border=1 frameSpacing=0 rows=300,* frameBorder=yes scrolling=yes>
	<FRAME name=topFrame src="misendmaster_top.asp?inputyn=<%= request("inputyn") %>" scrolling=yes>
	<FRAME name=mainFrame src="misendmaster_main.asp"  scrolling=yes>
</FRAMESET>
</HTML>
