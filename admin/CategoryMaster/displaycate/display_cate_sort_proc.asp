<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	Dim vQuery, vSubQ, i, vDepth, vCateCode, vCateCode_s, vSortNo, vRealTotal, vTotalCount, vChangeContentsCa, vSCMChangeSQLCa
	vDepth 		= NullFillWith(Request("depth"), "1")
	vCateCode 	= Replace(Trim(Request("catecode"))," ","")
	vSortNo		= Replace(Trim(Request("sortno"))," ","")
	vTotalCount	= Request("totalcount")
	vCateCode_s	= Request("catecode_s")

	vChangeContentsCa = "- 전시카테고리 " & vDepth & " Depth 정렬 수정 : " & vbCrLf

	vSubQ = vSubQ & " AND c.depth = '" & vDepth & "' "
	IF vDepth <> "1" Then
		vSubQ = vSubQ & " AND Left(c.catecode," & (3*(vDepth-1)) & ") = '" & vCateCode_s & "' "
	End If
	
	vQuery = "SELECT count(c.catecode) FROM [db_item].[dbo].[tbl_display_cate] AS c WHERE 1=1 " & vSubQ & " "
	rsget.Open vQuery,dbget,1
	vRealTotal = rsget(0)
	rsget.Close
	'rw vQuery
	If CStr(vRealTotal) <> CStr(vTotalCount) Then
		Response.Write "<script>alert('현재 해당 depth의 카테고리 갯수와 다릅니다.\n다시 확인해주세요.');parent.location.reload();</script>"
		dbget.close()
		Response.End
	End If
	
	vQuery = ""
	For i = 0 To vTotalCount-1
		vQuery = vQuery & " UPDATE [db_item].[dbo].[tbl_display_cate] SET sortNo = '" & Split(vSortNo,",")(i) & "' WHERE catecode = '" & Split(vCateCode,",")(i) & "' " & vbCrLf
		
		vChangeContentsCa = vChangeContentsCa & "- SET sortNo = " & Split(vSortNo,",")(i) & " WHERE catecode = " & Split(vCateCode,",")(i) & ", " & vbCrLf
	Next
	
	If vQuery <> "" Then
		dbget.execute vQuery

    	vSCMChangeSQLCa = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    	vSCMChangeSQLCa = vSCMChangeSQLCa & "VALUES('" & session("ssBctId") & "', 'dispcate', '" & Left(vCateCode_s,3) & "', '" & Request("menupos") & "', "
    	vSCMChangeSQLCa = vSCMChangeSQLCa & "'" & vChangeContentsCa & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQLCa)
	End If
%>
<script>
parent.opener.location.reload();
parent.window.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->