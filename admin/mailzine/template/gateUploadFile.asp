<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ������
' History : 2018.04.27 �̻� ����(���Ϸ� ���� ���� ���Ϸ��� �߼� ���� ����. ���� �������� ����.)
'			2019.06.24 ������ ����(���ø� ��� �ű� �߰�)
'			2020.05.28 �ѿ�� ����(TMS ���Ϸ� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
	Dim sName, strImgUrl
	sName	= requestCheckVar(Request("sName"),50) 
	strImgUrl	= requestCheckVar(Request("sImgUrl"),100)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script language="javascript">
<!--
	alert("�̹����� ��ϵǾ����ϴ�.\n\n�̹��� ����� �����ư�� ������ ó���Ϸ�˴ϴ�.");
	$("input[name='<%=sName%>']",opener.document).val("<%=strImgUrl%>");
	$("#<%=sName%>",opener.document).html("<%=strImgUrl%>");
	window.close();
//-->
</script>