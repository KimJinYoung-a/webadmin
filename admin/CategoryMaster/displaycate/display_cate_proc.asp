<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%

function fnRemoveSpecialChar(str)
	dim result : result = str
	result = Replace(result, Chr(34), "")
	result = Replace(result, Chr(39), "")
	result = Replace(result, Chr(44), "")
	fnRemoveSpecialChar = result
end function

	Dim vQuery, vCateCode, vCateName, vCateName_E, vDepth, vUseYN, vSortNo, vParentCateCode, vCompleteDel, vJaehuname, vIsNew, vCateKeywords, vSafetyInfoType, vSearchKeywords
	Dim vChangeContents, vSCMChangeSQL
	'vChangeContents = "- HTTP_REFERER : " & request.ServerVariables("HTTP_REFERER") & vbCrLf

	vCateCode		= Request("catecode")
	vCateName 		= html2db(fnRemoveSpecialChar(Request("catename")))
	vCateName_E		= html2db(fnRemoveSpecialChar(Request("catename_e")))
	vJaehuname		= html2db(fnRemoveSpecialChar(Request("jaehuname")))
	vDepth			= Request("depth")
	vUseYN			= Request("useyn")
	vSortNo			= Request("sortno")
	vParentCateCode	= Request("parentcatecode")
	vCompleteDel	= Request("completedel")
	vIsNew			= Request("isnew")
	vCateKeywords	= html2db(fnRemoveSpecialChar(Request("keywords")))
	vCateKeywords = Replace(vCateKeywords, "/", ",")
	vSafetyInfoType = Request("safetyinfotype")
	vSearchKeywords = Request("searchKeywords")

	If vDepth = "" Then
		dbget.close()
		Response.End
	End If

	If vCompleteDel = "o" Then
		vQuery = "SELECT count(catecode) FROM [db_item].[dbo].[tbl_display_cate] where Left(catecode," & Len(vCateCode) & ") = '" & vCateCode & "' and catecode <> '" & vCateCode & "'"
		rsget.Open vQuery,dbget,1
		If rsget(0) > 0 Then
			Response.Write "<script>alert('삭제하려는 카테고리에 하위카테고리가 존재하여\n하위카테고리를 모두 삭제 후 진행이 가능합니다.');</script>"
			rsget.close()
			dbget.close()
			Response.End
		Else
			rsget.close()
		End If

		vChangeContents = vChangeContents & "- 전시카테고리 삭제 catecode = " & vCateCode & "" & vbCrLf
		vQuery = "SELECT i.itemid, isNull((SELECT TOP 1 catecode FROM [db_item].[dbo].[tbl_display_cate_item] where itemid = i.itemid AND isDefault = 'n' ORDER BY depth ASC, sortno ASC),0) as cate "
		vQuery = vQuery & "FROM [db_item].[dbo].[tbl_display_cate_item] as i WHERE i.catecode = '" & vCateCode & "' AND i.isDefault = 'y'"
		rsget.Open vQuery,dbget,1
		Do Until rsget.Eof
			IF CStr(rsget("cate")) = "0" Then
				vQuery = vQuery & "UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = null WHERE itemid = '" & rsget("itemid") & "'; " & vbCrLf
				vChangeContents = vChangeContents & "- tbl_item.dispcate1(itemid:" & rsget("itemid") & ") = null" & vbCrLf
			Else
				vQuery = vQuery & "UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = '" & Left(rsget("cate"),3) & "' WHERE itemid = '" & rsget("itemid") & "'; " & vbCrLf
				vChangeContents = vChangeContents & "- tbl_item.dispcate1(itemid:" & rsget("itemid") & ") = " & Left(rsget("cate"),3) & vbCrLf
			End IF
		rsget.MoveNext
		Loop
		rsget.close()

		vQuery = vQuery & "DELETE [db_item].[dbo].[tbl_display_cate] WHERE catecode = '" & vCateCode & "'; " & vbCrLf
		vQuery = vQuery & "DELETE [db_item].[dbo].[tbl_display_cate_item] WHERE catecode = '" & vCateCode & "'; " & vbCrLf
		vQuery = vQuery & "DELETE [db_partner].[dbo].[tbl_partner_dispcate] WHERE catecode = '" & vCateCode & "'; " & vbCrLf
		dbget.execute vQuery


    	'### 수정 로그 저장(dispcate)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'dispcate', '" & Left(vCateCode,3) & "', '" & vCateCode & "', '" & Request("menupos") & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)


		If Len(vCateCode) = 3 Then
			vCateCode = ""
		Else
			vCateCode = Left(vCateCode,(Len(vCateCode)-3))
		End IF

		Response.Write "<script>parent.location.href='/admin/CategoryMaster/displaycate/display_cate_list.asp?menupos=1582&depth_s="&CHKIIF((Len(Request("catecode"))/3)=1,"1",(Len(Request("catecode"))/3))&"&catecode_s="&vCateCode&"';</script>"
		dbget.close()
		Response.End
	Else
		If vCateCode = "" Then
			If vDepth = "1" Then
				vQuery = "SELECT Top 1 c.catecode FROM [db_item].[dbo].[tbl_display_cate] AS c WHERE c.depth = '" & vDepth & "' ORDER BY c.catecode DESC"
				rsget.Open vQuery,dbget,1
				If Not rsget.Eof Then
					vCateCode = CInt(rsget("catecode")) + 1
				Else
					vCateCode = "101"
				End If
				rsget.close()
			Else
				vQuery = "SELECT Top 1 c.catecode FROM [db_item].[dbo].[tbl_display_cate] AS c WHERE c.depth = '" & vDepth & "' AND Left(c.catecode, "&(3*(vDepth-1))&") = '" & vParentCateCode & "' ORDER BY c.catecode DESC"
				rsget.Open vQuery,dbget,1
				If Not rsget.Eof Then
					vCateCode = CInt(Right(rsget("catecode"),3)) + 1
					vCateCode = vParentCateCode & vCateCode
				Else
					vCateCode = vParentCateCode & "101"
				End If
				rsget.close()
			End IF

			vQuery = "INSERT INTO [db_item].[dbo].[tbl_display_cate](catecode, depth, catename, catename_e, useyn, sortno, isnew, keywords, safetyinfotype, searchKeywords) "
			vQuery = vQuery & " VALUES('" & vCateCode & "', '" & vDepth & "', '" & vCateName & "', '" & vCateName_E & "', '" & vUseYN & "', '" & vSortNo & "', '" & vIsNew & "', '" & vCateKeywords & "', '" & vSafetyInfoType & "', '" & vSearchKeywords & "')"
			rw vQuery
			dbget.execute vQuery

			vChangeContents = vChangeContents & "- 전시카테고리 생성 catecode = " & vCateCode & "" & vbCrLf
			vChangeContents = vChangeContents & "- 한글명 catename = " & vCateName & "" & vbCrLf
			vChangeContents = vChangeContents & "- 영문명 catename_e = " & vCateName_E & "" & vbCrLf
			vChangeContents = vChangeContents & "- 사용유무 useyn = " & vUseYN & "" & vbCrLf
			vChangeContents = vChangeContents & "- 정렬번호 sortno = " & vSortNo & "" & vbCrLf
			vChangeContents = vChangeContents & "- new아이콘사용 isnew = " & vIsNew & "" & vbCrLf
			vChangeContents = vChangeContents & "- 카테고리 키워드 keywords = " & vCateKeywords & "" & vbCrLf
		Else
			vQuery = "UPDATE [db_item].[dbo].[tbl_display_cate] SET "
			vQuery = vQuery & " 	catename = '" & vCateName & "', "
			vQuery = vQuery & " 	catename_e = '" & vCateName_E & "', "
			vQuery = vQuery & " 	jaehuname = '" & vJaehuname & "', "
			vQuery = vQuery & " 	useyn = '" & vUseYN & "', "
			vQuery = vQuery & " 	isnew = '" & vIsNew & "', "
			vQuery = vQuery & " 	keywords = '" & vCateKeywords & "', "
			vQuery = vQuery & " 	sortno = '" & vSortNo & "', "
			vQuery = vQuery & " 	safetyinfotype = '" & vSafetyInfoType & "', "
			vQuery = vQuery & " 	searchKeywords = '" & vSearchKeywords & "' "
			vQuery = vQuery & " WHERE catecode = '" & vCateCode & "'"
			dbget.execute vQuery

			vQuery = "UPDATE [db_item].[dbo].[tbl_display_cate] SET useyn = '" & vUseYN & "' "
			vQuery = vQuery & " WHERE Left(catecode,'" & Len(vCateCode) & "') = '" & vCateCode & "' AND depth > '" & (Len(vCateCode)/3) & "'"
			dbget.execute vQuery

			vChangeContents = vChangeContents & "- 전시카테고리 수정 catecode = " & vCateCode & "" & vbCrLf
			vChangeContents = vChangeContents & "- 한글명 catename = " & vCateName & "" & vbCrLf
			vChangeContents = vChangeContents & "- 영문명 catename_e = " & vCateName_E & "" & vbCrLf
			vChangeContents = vChangeContents & "- 사용유무 useyn = " & vUseYN & "" & vbCrLf
			vChangeContents = vChangeContents & "- 정렬번호 sortno = " & vSortNo & "" & vbCrLf
			vChangeContents = vChangeContents & "- new아이콘사용 isnew = " & vIsNew & "" & vbCrLf
			vChangeContents = vChangeContents & "- 카테고리 키워드 keywords = " & vCateKeywords & "" & vbCrLf
		End If

    	'### 수정 로그 저장(dispcate)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'dispcate', '" & Left(vCateCode,3) & "', '" & vCateCode & "', '" & Request("menupos") & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)
	End If
%>

<script>
parent.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->