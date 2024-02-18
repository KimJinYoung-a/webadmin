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
'// 에러낸다. 2016-12-15, skyer9
'// 사용안하는 페이지이면 삭제한다.

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('oldmisendinput_main.asp?orderserial=<%= orderserial %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
