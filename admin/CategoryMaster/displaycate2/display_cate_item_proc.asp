<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/categorymaster/displaycate2/classes/displaycateCls.asp"-->

<%
	Response.CharSet = "euc-kr"
	
	Dim cDisp, vQuery, vAction, vDepth, vCateCode, vItemID, vSortNo, vIsDefault, vTemp
	Dim vChangeContents, vSCMChangeSQL
	vAction		= Request("action")
	vCateCode	= Request("catecode")
	vDepth		= (Len(vCateCode)/3)
	vItemID		= Request("itemid")
	vSortNo		= Request("sortno")
	vIsDefault	= Request("isdefault")
	
	If vItemID = "" Then
		dbget.close()
		Response.End
	End IF
	
	If vCateCode = "" Then
		dbget.close()
		Response.End
	End IF
	
	If vAction = "" Then
		vAction = "insert"
	End IF
	
	If vSortNo = "" Then
		vSortNo = 9999
	End If
	
	vQuery = ""
	If vAction = "update" OR vAction = "delete" Then
		vQuery = "SELECT count(catecode) FROM [db_item].[dbo].[tbl_display_cate2_item] WHERE itemid = '" & vItemID & "'"
		rsget.Open vQuery,dbget,1
		vTemp = rsget(0)
		rsget.close()
	End IF
	
	If vAction = "update" Then
		
		'#################################################################################################################################################################
		'
		'#################################################################################################################################################################
		If vTemp = 1 Then
			vIsDefault = "y"	'### 무조건 한개는 기본이어야함. 총갯수가 1개이므로 n으로 변경 불가.
		ElseIf vTemp > 1 Then
			vQuery = "SELECT catecode FROM [db_item].[dbo].[tbl_display_cate2_item] WHERE itemid = '" & vItemID & "' AND isDefault = 'y'"
			rsget.Open vQuery,dbget,1
			vTemp = rsget(0)
			rsget.close()
			If CStr(vTemp) = CStr(vCateCode) AND vIsDefault = "n" Then
				vQuery = "SELECT TOP 1 catecode FROM [db_item].[dbo].[tbl_display_cate2_item] where itemid = '" & vItemID & "' AND isDefault = 'n' ORDER BY depth ASC, sortno ASC"
				rsget.Open vQuery,dbget,1
				vTemp = rsget(0)
				rsget.close()
				vQuery = "UPDATE [db_item].[dbo].[tbl_display_cate2_item] SET isDefault = 'y' WHERE catecode = '" & vTemp & "' AND itemid = '" & vItemID & "' " & vbCrLf
				vQuery = vQuery & "UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = '" & Left(vTemp,3) & "' WHERE itemid = '" & vItemID & "'"
				dbget.execute vQuery
			End If
			If CStr(vTemp) <> CStr(vCateCode) AND vIsDefault = "y" Then		'### 이미 y가 있는데 다른카테고리를 y로 지정할경우 일단 같은 itemid 모두 n으로 변경.
				vQuery = "UPDATE [db_item].[dbo].[tbl_display_cate2_item] SET isDefault = 'n' WHERE itemid = '" & vItemID & "'"
				vQuery = vQuery & " UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = '" & Left(vCateCode,3) & "' WHERE itemid = '" & vItemID & "'"
				dbget.execute vQuery
			End If
		End If
		'#################################################################################################################################################################
		
		vQuery = "IF EXISTS(SELECT catecode FROM [db_item].[dbo].[tbl_display_cate2_item] WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "') " & vbCrLf
		vQuery = vQuery & "	BEGIN " & vbCrLf
		vQuery = vQuery & "		UPDATE [db_item].[dbo].[tbl_display_cate2_item] SET " & vbCrLf
		vQuery = vQuery & "			sortNo = '" & vSortNo & "', " & vbCrLf
		vQuery = vQuery & "			isDefault = '" & vIsDefault & "' " & vbCrLf
		vQuery = vQuery & "		WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "' " & vbCrLf
		vQuery = vQuery & "	END " & vbCrLf
		dbget.execute vQuery
		
		vChangeContents = "- 전시카테고리 update " & vbCrLf
		vChangeContents = vChangeContents & "- itemid = " & vItemID & " " & vbCrLf
		vChangeContents = vChangeContents & "- catecode = " & vCateCode & " " & vbCrLf
		vChangeContents = vChangeContents & "- sortNo = " & vSortNo & " " & vbCrLf
		vChangeContents = vChangeContents & "- isDefault = " & vIsDefault & " " & vbCrLf
	ElseIf vAction = "delete" Then
		
		'#################################################################################################################################################################
		'isDefault = 'y' 인것을 지우려할 경우 ORDER BY depth ASC, sortno ASC 로 top 1 catecode를 기본으로 지정.
		'#################################################################################################################################################################
		If vTemp > 1 Then
			vQuery = "SELECT catecode FROM [db_item].[dbo].[tbl_display_cate2_item] WHERE itemid = '" & vItemID & "' AND isDefault = 'y'"
			rsget.Open vQuery,dbget,1
			vTemp = rsget(0)
			rsget.close()
			If CStr(vTemp) = CStr(vCateCode) Then
				vQuery = "SELECT TOP 1 catecode FROM [db_item].[dbo].[tbl_display_cate2_item] where itemid = '" & vItemID & "' AND isDefault = 'n' ORDER BY depth ASC, sortno ASC"
				rsget.Open vQuery,dbget,1
				vTemp = rsget(0)
				rsget.close()
				vQuery = "UPDATE [db_item].[dbo].[tbl_display_cate2_item] SET isDefault = 'y' WHERE catecode = '" & vTemp & "' AND itemid = '" & vItemID & "' " & vbCrLf
				vQuery = vQuery & "	UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = '" & Left(vTemp,3) & "' WHERE itemid = '" & vItemID & "' " & vbCrLf
				dbget.execute vQuery
			End If
		Else
			vQuery = "UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = null WHERE itemid = '" & vItemID & "' " & vbCrLf
			dbget.execute vQuery
		End If
		'#################################################################################################################################################################
		
		vQuery = "DELETE [db_item].[dbo].[tbl_display_cate2_item] WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "' " & vbCrLf
		dbget.execute vQuery
		
		vChangeContents = "- 전시카테고리 delete " & vbCrLf
		vChangeContents = vChangeContents & "- itemid = " & vItemID & " " & vbCrLf
		vChangeContents = vChangeContents & "- catecode = " & vCateCode & " " & vbCrLf
	Else
		vQuery = vQuery & "IF NOT EXISTS(SELECT catecode FROM [db_item].[dbo].[tbl_display_cate2_item] WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "') " & vbCrLf
		vQuery = vQuery & "	BEGIN " & vbCrLf
		vQuery = vQuery & "		IF NOT EXISTS(SELECT catecode FROM [db_item].[dbo].[tbl_display_cate2_item] WHERE itemid = '" & vItemID & "' AND isDefault = 'y') " & vbCrLf
		vQuery = vQuery & "		BEGIN " & vbCrLf
		vQuery = vQuery & "			INSERT INTO [db_item].[dbo].[tbl_display_cate2_item](catecode, itemid, depth, sortNo, isDefault) " & vbCrLf
		vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'y') " & vbCrLf
		vQuery = vQuery & "			UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = '" & Left(vCateCode,3) & "' WHERE itemid = '" & vItemID & "' " & vbCrLf
		vQuery = vQuery & "		END " & vbCrLf
		vQuery = vQuery & "		ELSE " & vbCrLf
		vQuery = vQuery & "		BEGIN " & vbCrLf
		vQuery = vQuery & "			INSERT INTO [db_item].[dbo].[tbl_display_cate2_item](catecode, itemid, depth, sortNo, isDefault) " & vbCrLf
		vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'n')" & vbCrLf
		vQuery = vQuery & "		END " & vbCrLf
		vQuery = vQuery & "	END " & vbCrLf
		dbget.execute vQuery
		
		vChangeContents = "- 전시카테고리 insert " & vbCrLf
		vChangeContents = vChangeContents & "- itemid = " & vItemID & " " & vbCrLf
		vChangeContents = vChangeContents & "- catecode = " & vCateCode & " " & vbCrLf
	End If
	
	
    	'### 로그 저장(dispcate)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'dispcate', '" & Left(vCateCode,3) & "', '" & vCateCode & "', '" & Request("menupos") & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)
	
	
	If vAction = "insert" Then
		SET cDisp = New cDispCate
		cDisp.FCurrPage = 1
		cDisp.FPageSize = 1
		cDisp.FRectDepth = vDepth
		cDisp.FRectItemID = vItemID
		cDisp.GetDispCateItemList()
		
		If cDisp.FResultCount > 0 Then
			Response.Write fnCateCodeNameSplit(cDisp.FItemList(0).FCateName, vItemID)
		End If
		
		SET cDisp = Nothing
	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->