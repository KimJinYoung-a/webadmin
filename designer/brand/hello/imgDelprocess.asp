<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/helloCls.asp"-->
<%
Dim duplid, strSql
duplid = requestCheckVar(request("duplid"),100)
strSql = ""
strSql = strSql & " UPDATE db_brand.dbo.tbl_street_Hello SET bgImageURL = NULL WHERE makerid = '"&duplid&"' "
dbget.execute strSql
response.write "<script>alert('이미지가 삭제되었습니다');parent.location.reload();</script>"
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->