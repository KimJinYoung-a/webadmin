<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̺�Ʈ �׷� ���  '������ ������ ��ȸ ��Ŵ
' History : 2010.09.28 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim strUrl
	strUrl = request("strUrl")
%>

<script language="javascript">

	alert("��ϵǾ����ϴ�.");
	//opener.location.href='<%=strUrl%>';
	opener.location.reload(); 		
	window.close();

</script>

