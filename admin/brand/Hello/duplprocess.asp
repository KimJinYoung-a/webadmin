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
strSql = strSql & " SELECT count(*) as cnt FROM db_brand.dbo.tbl_street_Hello WHERE makerid = '"&duplid&"' "
rsget.Open strSql,dbget,1
If rsget("cnt") > 0 Then
	response.write "<script>alert('�ش� �귣�尡 ��ϵǾ� �ֽ��ϴ�.\n����ص� ������� �ʽ��ϴ�');</script>"
Else
	response.write "<script>alert('��ϰ����� �귣�� �Դϴ�.');</script>"
End If
rsget.Close
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->