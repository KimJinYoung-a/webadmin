<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 비트윈
' History : 2014.10.02 원승현 생성
'			2015.08.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim vQuery, vCateCode, vCateName, vDepth, vUseYN, vSortNo, vParentCateCode, vCompleteDel, vdispyn
	vCateCode		= Request("catecode")
	vCateName 		= html2db(Request("catename"))
	vDepth			= Request("depth")
	vUseYN			= Request("useyn")
	vSortNo			= Request("sortno")
	vParentCateCode	= Request("parentcatecode")
	vCompleteDel	= Request("completedel")
	vdispyn	= Request("dispyn")
	
	if vdispyn="" then vdispyn="N"
	If vDepth = "" Then
		dbCTget.close()
		Response.End
	End If

	
	If vCompleteDel = "o" Then
		vQuery = ""
		vQuery = vQuery & "DELETE [db_outmall].[dbo].[tbl_between_cate] WHERE catecode = '" & vCateCode & "'; " & vbCrLf
		vQuery = vQuery & "DELETE [db_outmall].[dbo].[tbl_between_cate_item] WHERE catecode = '" & vCateCode & "'; " & vbCrLf
		dbCTget.execute vQuery
		
		If Len(vCateCode) = 3 Then
			vCateCode = ""
		Else
			vCateCode = Left(vCateCode,(Len(vCateCode)-3))
		End IF
		Response.Write "<script>parent.location.href='/admin/etc/between/category/cate_list.asp?menupos=1582&depth_s="&CHKIIF((Len(Request("catecode"))/3)=1,"1",(Len(Request("catecode"))/3))&"&catecode_s="&vCateCode&"';</script>"
		dbCTget.close()
		Response.End
	Else
		If vCateCode = "" Then
			If vDepth = "1" Then
				vQuery = "SELECT TOP 1 catecode FROM db_outmall.dbo.tbl_between_cate WHERE depth = '" & vDepth & "' ORDER BY catecode DESC"
				rsCTget.Open vQuery,dbCTget,1
				If Not rsCTget.Eof Then
					vCateCode = CInt(rsCTget("catecode")) + 1
				Else
					vCateCode = "101"
				End If
				rsCTget.close()
			Else
				vQuery = "SELECT TOP 1 catecode FROM db_outmall.dbo.tbl_between_cate WHERE depth = '" & vDepth & "' AND Left(catecode, "&(3*(vDepth-1))&") = '" & vParentCateCode & "' ORDER BY catecode DESC"
				rsCTget.Open vQuery,dbCTget,1
				If Not rsCTget.Eof Then
					vCateCode = CInt(Right(rsCTget("catecode"),3)) + 1
					vCateCode = vParentCateCode & vCateCode
				Else
					vCateCode = vParentCateCode & "101"
				End If
				rsCTget.close()
			End IF
			
			vQuery = "INSERT INTO db_outmall.dbo.tbl_between_cate (catecode, depth, catename, useyn, sortno, dispyn) "
			vQuery = vQuery & " VALUES('" & vCateCode & "', '" & vDepth & "', '" & vCateName & "', '" & vUseYN & "', '" & vSortNo & "', '" & vdispyn & "')"
			rw vQuery
			dbCTget.execute vQuery
		Else
			vQuery = "UPDATE db_outmall.dbo.tbl_between_cate SET "
			vQuery = vQuery & " 	catename = '" & vCateName & "'"
			vQuery = vQuery & " 	,useyn = '" & vUseYN & "'"
			vQuery = vQuery & " 	,sortno = '" & vSortNo & ""
			vQuery = vQuery & " 	,dispyn = '" & vdispyn & "'"
			vQuery = vQuery & " WHERE catecode = '" & vCateCode & "'"
			dbCTget.execute vQuery
		End If
	End If
%>
<script>
parent.location.reload();
</script>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->