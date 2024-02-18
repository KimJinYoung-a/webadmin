<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/categorymaster/displaycate2/classes/displaycateCls.asp"-->

<%
	Dim cDisp, i, vQuery, vAction, vDepth, vCateCode, vAllItemID, vItemID, vSortNo, vIsDefault, vTemp, vCount, vDefCate, vChgCateL3, vChgCate, vIsDef
	Dim vChangeContents, vSCMChangeSQL
	vAction		= Request("action")
	vCateCode	= Request("catecode")
	vDepth		= (Len(vCateCode)/3)
	vAllItemID	= Request("allitemid")
	vSortNo		= Request("sortno")
	vIsDefault	= Request("isdefault")
	
	If vAllItemID = "" Then
		dbget.close()
		Response.End
	End IF
	
	If vAction<>"delete" and vCateCode = "" Then
		dbget.close()
		Response.End
	End IF

	If vSortNo = "" Then
		vSortNo = 9999
	End If
	
	vAllItemID = Replace(Trim(vAllItemID)," ","")
	If Right(vAllItemID,1) = "," Then
		vAllItemID = Left(vAllItemID,(Len(vAllItemID)-1))
	End IF

	For i = 0 To UBound(Split(vAllItemID,","))
		vItemID = Split(vAllItemID,",")(i)
		
		
		If vAction = "delete" Then
			vQuery = "DELETE [db_item].[dbo].[tbl_display_cate2_item] WHERE itemid = '" & vItemID & "' "  & vbCrLf
			dbget.execute vQuery
		Else
			vQuery = ""
			vQuery = vQuery & "IF NOT EXISTS(SELECT catecode FROM [db_item].[dbo].[tbl_display_cate2_item] WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "') " & vbCrLf
			vQuery = vQuery & "	BEGIN " & vbCrLf
			vQuery = vQuery & "		IF NOT EXISTS(SELECT catecode FROM [db_item].[dbo].[tbl_display_cate2_item] WHERE itemid = '" & vItemID & "' AND isDefault = 'y') " & vbCrLf
			vQuery = vQuery & "		BEGIN " & vbCrLf
			vQuery = vQuery & "			INSERT INTO [db_item].[dbo].[tbl_display_cate2_item](catecode, itemid, depth, sortNo, isDefault) " & vbCrLf
			vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'y') " & vbCrLf
			vQuery = vQuery & "		END " & vbCrLf
			vQuery = vQuery & "		ELSE " & vbCrLf
			vQuery = vQuery & "		BEGIN " & vbCrLf
			vQuery = vQuery & "			INSERT INTO [db_item].[dbo].[tbl_display_cate2_item](catecode, itemid, depth, sortNo, isDefault) " & vbCrLf
			vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'n')" & vbCrLf
			vQuery = vQuery & "		END " & vbCrLf
			vQuery = vQuery & "	END " & vbCrLf
			dbget.execute vQuery
		End If
		
	Next
%>
<script>parent.location.reload();</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->