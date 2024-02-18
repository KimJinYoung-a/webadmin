<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.29 김진영 생성
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
	response.write "<script>alert('해당 브랜드가 등록되어 있습니다.\n등록해도 저장되지 않습니다');</script>"
Else
	response.write "<script>alert('등록가능한 브랜드 입니다.');</script>"
End If
rsget.Close
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->