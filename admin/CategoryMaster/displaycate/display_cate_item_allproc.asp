<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->

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
			vQuery = "DELETE [db_item].[dbo].[tbl_display_cate_item] WHERE itemid = '" & vItemID & "' "  & vbCrLf
			vQuery = vQuery & " UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = null WHERE itemid = '" & vItemID & "' "
			dbget.execute vQuery
		Else
			vQuery = ""
			vQuery = vQuery & "IF NOT EXISTS(SELECT catecode FROM [db_item].[dbo].[tbl_display_cate_item] WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "') " & vbCrLf
			vQuery = vQuery & "	BEGIN " & vbCrLf
			vQuery = vQuery & "		IF NOT EXISTS(SELECT catecode FROM [db_item].[dbo].[tbl_display_cate_item] WHERE itemid = '" & vItemID & "' AND isDefault = 'y') " & vbCrLf
			vQuery = vQuery & "		BEGIN " & vbCrLf
			vQuery = vQuery & "			INSERT INTO [db_item].[dbo].[tbl_display_cate_item](catecode, itemid, depth, sortNo, isDefault) " & vbCrLf
			vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'y') " & vbCrLf
			vQuery = vQuery & "			UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = '" & Left(vCateCode,3) & "' WHERE itemid = '" & vItemID & "' " & vbCrLf
			vQuery = vQuery & "		END " & vbCrLf
			vQuery = vQuery & "		ELSE " & vbCrLf
			vQuery = vQuery & "		BEGIN " & vbCrLf
			vQuery = vQuery & "			INSERT INTO [db_item].[dbo].[tbl_display_cate_item](catecode, itemid, depth, sortNo, isDefault) " & vbCrLf
			vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'n')" & vbCrLf
			vQuery = vQuery & "		END " & vbCrLf
			vQuery = vQuery & "	END " & vbCrLf
			dbget.execute vQuery
		End If
		
	Next
	
	
	If vAction = "delete" Then
    	'### 수정 로그 저장(dispcate)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'dispcate', '" & Request("menupos") & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'- 선택한 상품에 등록된 카테고리 모두 삭제 : " & vAllItemID & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)
	Else
    	'### 수정 로그 저장(dispcate)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'dispcate', '" & Left(vCateCode,3) & "', '" & vCateCode & "', '" & Request("menupos") & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'- 선택한것 모두 등록 : " & vAllItemID & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)
	End If
%>
<script>parent.location.reload();</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->