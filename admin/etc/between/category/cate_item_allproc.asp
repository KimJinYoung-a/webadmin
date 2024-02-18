<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/betweenItemcls.asp"-->
<%
	Dim cDisp, i, vQuery, vAction, vDepth, vCateCode, vAllItemID, vItemID, vSortNo, vIsDefault, vTemp, vCount, vDefCate, vChgCateL3, vChgCate, vIsDef
	vAction		= Request("action")
	vCateCode	= Request("catecode")
	vDepth		= (Len(vCateCode)/3)
	vAllItemID	= Request("allitemid")
	vSortNo		= Request("sortno")
	vIsDefault	= Request("isdefault")
	If vAllItemID = "" Then
		dbCTget.close()
		Response.End
	End IF
	If vCateCode = "" Then
		dbCTget.close()
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
			vQuery = "DELETE [db_outmall].[dbo].[tbl_between_cate_item] WHERE itemid = '" & vItemID & "' "  & vbCrLf
			dbCTget.execute vQuery
		Else
			vQuery = ""
			vQuery = vQuery & "IF NOT EXISTS(SELECT catecode FROM [db_outmall].[dbo].[tbl_between_cate_item] WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "') " & vbCrLf
			vQuery = vQuery & "	BEGIN " & vbCrLf
			vQuery = vQuery & "		IF NOT EXISTS(SELECT catecode FROM [db_outmall].[dbo].[tbl_between_cate_item] WHERE itemid = '" & vItemID & "' AND isDefault = 'y') " & vbCrLf
			vQuery = vQuery & "		BEGIN " & vbCrLf
			vQuery = vQuery & "			INSERT INTO [db_outmall].[dbo].[tbl_between_cate_item](catecode, itemid, depth, sortNo, isDefault, regdate) " & vbCrLf
			vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'y', getdate()) " & vbCrLf
			vQuery = vQuery & "		END " & vbCrLf
			vQuery = vQuery & "		ELSE " & vbCrLf
			vQuery = vQuery & "		BEGIN " & vbCrLf
			vQuery = vQuery & "			INSERT INTO [db_outmall].[dbo].[tbl_between_cate_item](catecode, itemid, depth, sortNo, isDefault, regdate) " & vbCrLf
			vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'n', getdate())" & vbCrLf
			vQuery = vQuery & "		END " & vbCrLf
			vQuery = vQuery & "	END " & vbCrLf
			dbCTget.execute vQuery
		End If
		
	Next
%>
<script>parent.location.reload();</script>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->