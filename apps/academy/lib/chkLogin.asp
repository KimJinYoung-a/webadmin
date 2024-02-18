<!-- #include virtual="/apps/academy/lib/tenSessionLib.asp" -->
<%
If request.cookies("partner")("userid")="" Then
%>
<script type="text/javascript">
<!--
	fnAPPclosePopup();
//-->
</script>
<%
    response.end
End If
%>