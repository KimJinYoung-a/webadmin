<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.29 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/helloCls.asp"-->
<%
Dim duplid, strSql
duplid = request("duplid")
strSql = ""
strSql = strSql & " UPDATE db_brand.dbo.tbl_street_Hello SET bgImageURL = NULL WHERE makerid = '"&duplid&"' "
dbget.execute strSql
response.write "<script>alert('�̹����� �����Ǿ����ϴ�');parent.location.reload();</script>"
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->