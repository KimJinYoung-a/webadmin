<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/CategoryMaster/displaycate/classes/displaycateCls.asp"-->

<%
	Dim cDisp, i, vQuery, vAction, vDepth, vCateCode, vAllItemID, vItemID, vSortNo, vIsDefault, vTemp, vCount, vDefCate, vChgCateL3, vChgCate, vIsDef
	Dim vChangeContents, vSCMChangeSQL
	vAction		= RequestCheckvar(Request("action"),16)
	vCateCode	= RequestCheckvar(Request("catecode"),10)
	vDepth		= (Len(vCateCode)/3)
	vAllItemID	= Request("allitemid")
	vSortNo		= RequestCheckvar(Request("sortno"),10)
	vIsDefault	= RequestCheckvar(Request("isdefault"),1)
	
	If vAllItemID = "" Then
		dbACADEMYget.close()
		Response.End
	End IF
	
	If vCateCode = "" Then
		dbACADEMYget.close()
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
			vQuery = "DELETE [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE catecode='" & Cstr(vCateCode) & "' and itemid = '" & vItemID & "' and isDefault='n'"  & vbCrLf
			'vQuery = vQuery & " UPDATE [db_academy].[dbo].[tbl_diy_item] SET dispcate1 = null WHERE itemid = '" & vItemID & "' "
			dbACADEMYget.execute vQuery
		Else
			vQuery = ""
			vQuery = vQuery & "IF NOT EXISTS(SELECT catecode FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "') " & vbCrLf
			vQuery = vQuery & "	BEGIN " & vbCrLf
			vQuery = vQuery & "		IF NOT EXISTS(SELECT catecode FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE itemid = '" & vItemID & "' AND isDefault = 'y') " & vbCrLf
			vQuery = vQuery & "		BEGIN " & vbCrLf
			vQuery = vQuery & "			INSERT INTO [db_academy].[dbo].[tbl_display_cate_item_Academy](catecode, itemid, depth, sortNo, isDefault) " & vbCrLf
			vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'y') " & vbCrLf
			vQuery = vQuery & "			UPDATE [db_academy].[dbo].[tbl_diy_item] SET dispcate1 = '" & Left(vCateCode,3) & "' WHERE itemid = '" & vItemID & "' " & vbCrLf
			vQuery = vQuery & "		END " & vbCrLf
			vQuery = vQuery & "		ELSE " & vbCrLf
			vQuery = vQuery & "		BEGIN " & vbCrLf
			vQuery = vQuery & "			INSERT INTO [db_academy].[dbo].[tbl_display_cate_item_Academy](catecode, itemid, depth, sortNo, isDefault) " & vbCrLf
			vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'n')" & vbCrLf
			vQuery = vQuery & "		END " & vbCrLf
			vQuery = vQuery & "	END " & vbCrLf
			dbACADEMYget.execute vQuery
		End If
		
	Next
%>
<script>parent.fnDispCateAddEnd();</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->