<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̹��� ���ó��
' History : 2018.08.16 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
	Dim sName,sSpan,strImgUrl, sWidth , sOpt
	sName	= requestCheckVar(Request("sName"),50) 
	sSpan	= requestCheckVar(Request("sSpan"),50)  
	sWidth  = requestCheckVar(Request("sWidth"),10)  
	strImgUrl	= requestCheckVar(Request("sImgUrl"),100) 
	sOpt	= requestCheckVar(Request("sOpt"),1) 

%>
<script language="javascript">
<!--	 
	var sName, sSpan;
	sName = "<%=sName%>";	
	sSpan = "<%=sSpan%>";
	
	alert("�̹����� ��ϵǾ����ϴ�.\n\n�̹��� ����� �����ư�� ������ ó���Ϸ�˴ϴ�.");
	opener.eval("document.all."+sName).value = "<%=strImgUrl%>";
	opener.eval("document.all."+sSpan).innerHTML ="<img src='<%=strImgUrl%>' width='100'>";
	opener.eval("document.all."+sSpan).style.display = "";
	window.close();
//-->
</script>