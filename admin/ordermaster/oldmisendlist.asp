<%
response.write "������ �޴�, ������ ���� ���."
response.end
'response.redirect("newmisendlist.asp?notincludeupchecheck=on&delaydate=4")
'dbget.close()	:	response.End
%>


<HTML>
<FRAMESET border=1 frameSpacing=0 rows=300,* frameBorder=yes scrolling=yes>
	<FRAME name=topFrame src="oldmisendinput_top.asp" scrolling=yes>
	<FRAME name=mainFrame src="oldmisendinput_main.asp"  scrolling=yes>
</FRAMESET>
</HTML>
