<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	Dim vQuery, vSubQ, i, vDepth, vCateCode, vCateCode_s, vSortNo, vRealTotal, vTotalCount
	vDepth 		= NullFillWith(Request("depth"), "1")
	vCateCode 	= Replace(Trim(Request("catecode"))," ","")
	vSortNo		= Replace(Trim(Request("sortno"))," ","")
	vTotalCount	= Request("totalcount")
	vCateCode_s	= Request("catecode_s")
	
	
	vSubQ = vSubQ & " AND depth = '" & vDepth & "' "
	IF vDepth <> "1" Then
		vSubQ = vSubQ & " AND Left(catecode," & (3*(vDepth-1)) & ") = '" & vCateCode_s & "' "
	End If
	
	vQuery = "SELECT COUNT(catecode) FROM db_outmall.dbo.tbl_between_cate WHERE 1=1 " & vSubQ & " "
	rsCTget.Open vQuery,dbCTget,1
	vRealTotal = rsCTget(0)
	rsCTget.Close
	'rw vQuery
	If CStr(vRealTotal) <> CStr(vTotalCount) Then
		Response.Write "<script>alert('현재 해당 depth의 카테고리 갯수와 다릅니다.\n다시 확인해주세요.');parent.location.reload();</script>"
		dbCTget.close()
		Response.End
	End If
	
	vQuery = ""
	For i = 0 To vTotalCount-1
		vQuery = vQuery & " UPDATE db_outmall.dbo.tbl_between_cate SET sortNo = '" & Split(vSortNo,",")(i) & "' WHERE catecode = '" & Split(vCateCode,",")(i) & "' " & vbCrLf
	Next
	
	If vQuery <> "" Then
		dbCTget.execute vQuery
	End If
%>
<script>
parent.opener.location.reload();
parent.window.close();
</script>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->