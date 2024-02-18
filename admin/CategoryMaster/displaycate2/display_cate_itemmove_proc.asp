<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/categorymaster/displaycate2/classes/displaycateCls.asp"-->

<%
	Dim vQuery, vCateCode_A, vCateCode_B, vDepth, vOnlyThisCate, vCateCodeQuery, vTempItem, vChangeContents, vSCMChangeSQL, vMode
	vCateCode_A		= Request("catecode_a")
	vCateCode_B		= Request("catecode_b")
	vDepth			= (Len(vCateCode_B)/3)
	vOnlyThisCate	= Request("onlythiscate")
	vMode			= Request("mode")
	
	vChangeContents = "- 전시카테고리 전상품 이동 : " & vCateCode_A & " ---> " & vCateCode_B & "" & vbCrLf
	
	If vOnlyThisCate = "N" Then
		vCateCodeQuery = "Left(i.catecode," & Len(vCateCode_A) & ") = '" & vCateCode_A & "'"
		vChangeContents = vChangeContents & "- 이동할 카테고리 하위 뎁스 상품 전부 이동" & vbCrLf
	Else
		vCateCodeQuery = "i.catecode = '" & vCateCode_A & "'"
		vChangeContents = vChangeContents & "- 이동할 카테고리 상품만 이동" & vbCrLf
	End If
	
	'####### _A -------> 이동해야할곳
	'####### _B -------> 옮겨질곳
	
	If vCateCode_A = "" OR vCateCode_B = "" OR vOnlyThisCate = "" Then
		dbget.close()
		Response.End
	End If
	
	if vMode="getOldCate" then
		'// 기존 카테고리에서 상품 지정
		vQuery = "INSERT INTO [db_item].[dbo].[tbl_display_cate2_item]" & vbCrLf & _
				"		SELECT '" & vCateCode_B & "', i.itemid, '" & vDepth & "', 9999, " & vbCrLf & _
				"			(select case when count(itemid) > 0 then 'n' else 'y' end from [db_item].[dbo].[tbl_display_cate2_item] " & vbCrLf & _
				"	 		where itemid = i.itemid and isDefault = 'y') " & vbCrLf & _
				"		FROM [db_item].[dbo].[tbl_display_cate_item] as i " & vbCrLf & _
				"		WHERE " & vCateCodeQuery & " and i.itemid not in(select itemid from [db_item].[dbo].[tbl_display_cate2_item] where catecode = '" & vCateCode_B & "') " & vbCrLf & _
				"		GROUP BY i.itemid"
		dbget.execute vQuery
	else
		'// 카테고리내 상품 이동
		vQuery = "INSERT INTO [db_item].[dbo].[tbl_display_cate2_item]" & vbCrLf & _
				"		SELECT '" & vCateCode_B & "', i.itemid, '" & vDepth & "', 9999, " & vbCrLf & _
				"			(select case when count(itemid) > 0 then 'n' else 'y' end from [db_item].[dbo].[tbl_display_cate2_item] " & vbCrLf & _
				"	 		where itemid = i.itemid and catecode <> '" & vCateCode_A & "' and isDefault = 'y') " & vbCrLf & _
				"		FROM [db_item].[dbo].[tbl_display_cate2_item] as i " & vbCrLf & _
				"		WHERE " & vCateCodeQuery & " and i.itemid not in(select itemid from [db_item].[dbo].[tbl_display_cate2_item] where catecode = '" & vCateCode_B & "') " & vbCrLf & _
				"		GROUP BY i.itemid"
		dbget.execute vQuery
		
		vQuery = "DELETE [db_item].[dbo].[tbl_display_cate2_item] WHERE " & Replace(vCateCodeQuery,"i.","") & ""
		dbget.execute vQuery
	end if
%>

<script>
parent.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->