<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̹��� �̸����� 
' History : 2008.04.22 �ѿ�� ����
'###########################################################
%>

<%
dim image
	image = request("image")
%>
<img src="<%=image%>">