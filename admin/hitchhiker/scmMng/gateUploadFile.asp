<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̹��� ���ó��
' History : 2011.03.16 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
	Dim  strImgUrl 
 
	strImgUrl	= requestCheckVar(Request("sImgUrl"),100)  
	 
%> 
<script type="text/javascript"> 
	alert("������ ���ε� �Ǿ����ϴ�.");
	opener.document.getElementById("sfimg").value = "<%=strImgUrl%>";		
	opener.document.all.dvFUrl.innerHTML = "<%=strImgUrl%>";		
	window.close(); 
</script>
 