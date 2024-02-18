<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim sqlStr, voutmallorderseq
voutmallorderseq = requestcheckvar(request("outmallorderseq"),10)

If voutmallorderseq="" Then
	response.write "<script type='text/javascript'>alert('해당되는 주문건이 없습니다.'); self.close();</script>"
	dbget.close()	:	response.end
End If

sqlStr = ""
sqlStr = sqlStr & " UPDATE db_temp.dbo.tbl_xSite_TMPOrder " + vbcrlf
sqlStr = sqlStr & " SET sendState='951'" + vbcrlf
sqlStr = sqlStr & " WHERE outmallorderseq = '"&voutmallorderseq&"'" + vbcrlf
dbget.execute sqlStr
response.write "<script type='text/javascript'>alert('처리 되었습니다.'); opener.location.reload(); self.close();</script>"
dbget.close()	:	response.end
%>
<script>
	opener.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->