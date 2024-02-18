<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/CategoryMaster/displaycate/classes/displaycateCls.asp"-->
<%
	Dim vQuery, vSubQ, i, vDepth, vCateCode, vCateCode_s, vSortNo, vRealTotal, vTotalCount, vChangeContentsCa, vSCMChangeSQLCa
	vDepth 		= NullFillWith(RequestCheckvar(Request("depth"),16), "1")
	vCateCode 	= Replace(Trim(RequestCheckvar(Request("catecode"),16))," ","")
	vSortNo		= Replace(Trim(RequestCheckvar(Request("sortno"),10))," ","")
	vTotalCount	= Request("totalcount")
	vCateCode_s	= Request("catecode_s")

	vSubQ = vSubQ & " AND c.depth = '" & vDepth & "' "
	IF vDepth <> "1" Then
		vSubQ = vSubQ & " AND Left(c.catecode," & (3*(vDepth-1)) & ") = '" & vCateCode_s & "' "
	End If
	
	vQuery = "SELECT count(c.catecode) FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS c WHERE 1=1 " & vSubQ & " "
	rsACADEMYget.Open vQuery,dbACADEMYget,1
	vRealTotal = rsACADEMYget(0)
	rsACADEMYget.Close
	'rw vQuery
	If CStr(vRealTotal) <> CStr(vTotalCount) Then
		Response.Write "<script>alert('현재 해당 depth의 카테고리 갯수와 다릅니다.\n다시 확인해주세요.');parent.location.reload();</script>"
		dbACADEMYget.close()
		Response.End
	End If
	
	vQuery = ""
	For i = 0 To vTotalCount-1
		vQuery = vQuery & " UPDATE [db_academy].[dbo].[tbl_display_cate_Academy] SET sortNo = '" & Split(vSortNo,",")(i) & "' WHERE catecode = '" & Split(vCateCode,",")(i) & "' " & vbCrLf
		
	Next
	
	If vQuery <> "" Then
		dbACADEMYget.execute vQuery
	End If
%>
<script>
parent.opener.location.reload();
parent.window.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->