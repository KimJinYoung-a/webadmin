<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
dim orderserial, didxarr
dim mode
orderserial = request("orderserial")
didxarr = request("didxarr")
mode = request("mode")

1
'// ��������. 2016-12-15, skyer9
'// �����ϴ� �������̸� �����Ѵ�.

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('���� �Ǿ����ϴ�.');
location.replace('oldmisendinput_main.asp?orderserial=<%= orderserial %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
