 <%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전자결재 폼 선택
' History : 2011.03.14 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
Dim sReDirectURL, ireportState, ireportidx
sReDirectURL = requestCheckvar(Request("sRDURL"),100)
ireportState =  requestCheckvar(Request("iRS"),4) 
ireportidx =  requestCheckvar(Request("iridx"),10)

IF sReDirectURL = "" THEN 
	sReDirectURL = "/admin/approval/eapp/main.asp"
ELSE
	sReDirectURL = sReDirectURL&"?iRS="&ireportState&"&iridx="&ireportidx
END IF	
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->   
<frameset   cols="180,*"  border="1" framespacing="0" scrolling="yes">
   <frame name="eLMenu" src="/admin/approval/eapp/leftmenu.asp" scrolling="auto">
   <frame name="eConts" src="<%=sReDirectURL%>" scrolling="auto">
</frameset>  
</html>
 	
 