<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim vQuery, vCateCode, vCateName, vCateName_E, vDepth, vUseYN, vSortNo, vParentCateCode, vCompleteDel, vJaehuname, vIsNew
	Dim vChangeContents, vSCMChangeSQL

	vCateCode		= RequestCheckvar(Request("catecode"),10)
	vCateName 		= html2db(RequestCheckvar(Request("catename"),32))
	vCateName_E		= html2db(RequestCheckvar(Request("catename_e"),32))
	vJaehuname		= html2db(RequestCheckvar(Request("jaehuname"),32))
	vDepth			= RequestCheckvar(Request("depth"),10)
	vUseYN			= RequestCheckvar(Request("useyn"),10)
	vSortNo			= RequestCheckvar(Request("sortno"),10)
	vParentCateCode	= RequestCheckvar(Request("parentcatecode"),10)
	vCompleteDel	= RequestCheckvar(Request("completedel"),10)
	vIsNew			= RequestCheckvar(Request("isnew"),10)
	
	If vDepth = "" Then
		dbACADEMYget.close()
		Response.End
	End If

	If vCompleteDel = "o" Then
		vQuery = "SELECT count(catecode) FROM [db_academy].[dbo].[tbl_display_cate_Academy] where Left(catecode," & Len(vCateCode) & ") = '" & vCateCode & "' and catecode <> '" & vCateCode & "'"
		rsACADEMYget.Open vQuery,dbACADEMYget,1
		If rsACADEMYget(0) > 0 Then
			Response.Write "<script>alert('삭제하려는 카테고리에 하위카테고리가 존재하여\n하위카테고리를 모두 삭제 후 진행이 가능합니다.');</script>"
			rsACADEMYget.close()
			dbACADEMYget.close()
			Response.End
		Else
			rsACADEMYget.close()
		End If
		
		vQuery = "SELECT i.itemid, isNull((SELECT TOP 1 catecode FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] where itemid = i.itemid AND isDefault = 'n' ORDER BY depth ASC, sortno ASC),0) as cate "
		vQuery = vQuery & "FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] as i WHERE i.catecode = '" & vCateCode & "' AND i.isDefault = 'y'"
		rsACADEMYget.Open vQuery,dbACADEMYget,1
		Do Until rsACADEMYget.Eof
			IF CStr(rsACADEMYget("cate")) = "0" Then
				vQuery = vQuery & "UPDATE [db_academy].[dbo].[tbl_diy_item] SET dispcate1 = null WHERE itemid = '" & rsACADEMYget("itemid") & "'; " & vbCrLf
			Else
				vQuery = vQuery & "UPDATE [db_academy].[dbo].[tbl_diy_item] SET dispcate1 = '" & Left(rsACADEMYget("cate"),3) & "' WHERE itemid = '" & rsACADEMYget("itemid") & "'; " & vbCrLf
			End IF
		rsACADEMYget.MoveNext
		Loop
		rsACADEMYget.close()
		
		vQuery = vQuery & "DELETE [db_academy].[dbo].[tbl_display_cate_Academy] WHERE catecode = '" & vCateCode & "'; " & vbCrLf
		vQuery = vQuery & "DELETE [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE catecode = '" & vCateCode & "'; " & vbCrLf
		dbACADEMYget.execute vQuery
		

		If Len(vCateCode) = 3 Then
			vCateCode = ""
		Else
			vCateCode = Left(vCateCode,(Len(vCateCode)-3))
		End IF
		
		Response.Write "<script>parent.location.href='/academy/CategoryMaster/displaycate/display_cate_list.asp?menupos=1790&depth_s="&CHKIIF((Len(Request("catecode"))/3)=1,"1",(Len(Request("catecode"))/3))&"&catecode_s="&vCateCode&"';</script>"
		dbACADEMYget.close()
		Response.End
	Else
		If vCateCode = "" Then
			If vDepth = "1" Then
				vQuery = "SELECT Top 1 c.catecode FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS c WHERE c.depth = '" & vDepth & "' ORDER BY c.catecode DESC"
				rsACADEMYget.Open vQuery,dbACADEMYget,1
				If Not rsACADEMYget.Eof Then
					vCateCode = CInt(rsACADEMYget("catecode")) + 1
				Else
					vCateCode = "101"
				End If
				rsACADEMYget.close()
			Else
				vQuery = "SELECT Top 1 c.catecode FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS c WHERE c.depth = '" & vDepth & "' AND Left(c.catecode, "&(3*(vDepth-1))&") = '" & vParentCateCode & "' ORDER BY c.catecode DESC"
				rsACADEMYget.Open vQuery,dbACADEMYget,1
				If Not rsACADEMYget.Eof Then
					vCateCode = CInt(Right(rsACADEMYget("catecode"),3)) + 1
					vCateCode = vParentCateCode & vCateCode
				Else
					vCateCode = vParentCateCode & "101"
				End If
				rsACADEMYget.close()
			End IF
			
			vQuery = "INSERT INTO [db_academy].[dbo].[tbl_display_cate_Academy](catecode, depth, catename, catename_e, useyn, sortno, isnew) "
			vQuery = vQuery & " VALUES('" & vCateCode & "', '" & vDepth & "', '" & vCateName & "', '" & vCateName_E & "', '" & vUseYN & "', '" & vSortNo & "', '" & vIsNew & "')"
			rw vQuery
			dbACADEMYget.execute vQuery

		Else
			vQuery = "UPDATE [db_academy].[dbo].[tbl_display_cate_Academy] SET "
			vQuery = vQuery & " 	catename = '" & vCateName & "', "
			vQuery = vQuery & " 	catename_e = '" & vCateName_E & "', "
			vQuery = vQuery & " 	jaehuname = '" & vJaehuname & "', "
			vQuery = vQuery & " 	useyn = '" & vUseYN & "', "
			vQuery = vQuery & " 	isnew = '" & vIsNew & "', "
			vQuery = vQuery & " 	sortno = '" & vSortNo & "' "
			vQuery = vQuery & " WHERE catecode = '" & vCateCode & "'"
			dbACADEMYget.execute vQuery
			
			vQuery = "UPDATE [db_academy].[dbo].[tbl_display_cate_Academy] SET useyn = '" & vUseYN & "' "
			vQuery = vQuery & " WHERE Left(catecode,'" & Len(vCateCode) & "') = '" & vCateCode & "' AND depth > '" & (Len(vCateCode)/3) & "'"
			dbACADEMYget.execute vQuery

		End If

	End If
%>

<script>
parent.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->