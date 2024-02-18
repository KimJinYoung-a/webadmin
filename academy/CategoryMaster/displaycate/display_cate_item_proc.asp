<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/CategoryMaster/displaycate/classes/displaycateCls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Response.CharSet = "euc-kr"
	
	Dim cDisp, vQuery, vAction, vDepth, vCateCode, vItemID, vSortNo, vIsDefault, vTemp
	Dim vChangeContents, vSCMChangeSQL
	vAction		= RequestCheckvar(Request("action"),16)
	vCateCode	= RequestCheckvar(Request("catecode"),10)
	vDepth		= (Len(vCateCode)/3)
	vItemID		= RequestCheckvar(Request("itemid"),10)
	vSortNo		= RequestCheckvar(Request("sortno"),10)
	vIsDefault	= RequestCheckvar(Request("isdefault"),1)
	
	If vItemID = "" Then
		dbACADEMYget.close()
		Response.End
	End IF
	
	If vCateCode = "" Then
		dbACADEMYget.close()
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
		vQuery = "SELECT count(catecode) FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE itemid = '" & vItemID & "'"
		rsACADEMYget.Open vQuery,dbACADEMYget,1
		vTemp = rsACADEMYget(0)
		rsACADEMYget.close()
	End IF
	
	If vAction = "update" Then
		
		'#################################################################################################################################################################
		'
		'#################################################################################################################################################################
		If vTemp = 1 Then
			vIsDefault = "y"	'### 무조건 한개는 기본이어야함. 총갯수가 1개이므로 n으로 변경 불가.
		ElseIf vTemp > 1 Then
			vQuery = "SELECT catecode FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE itemid = '" & vItemID & "' AND isDefault = 'y'"
			rsACADEMYget.Open vQuery,dbACADEMYget,1
			vTemp = rsACADEMYget(0)
			rsACADEMYget.close()
			If CStr(vTemp) = CStr(vCateCode) AND vIsDefault = "n" Then
				vQuery = "SELECT TOP 1 catecode FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] where itemid = '" & vItemID & "' AND isDefault = 'n' ORDER BY depth ASC, sortno ASC"
				rsACADEMYget.Open vQuery,dbACADEMYget,1
				vTemp = rsACADEMYget(0)
				rsACADEMYget.close()
				vQuery = "UPDATE [db_academy].[dbo].[tbl_display_cate_item_Academy] SET isDefault = 'y' WHERE catecode = '" & vTemp & "' AND itemid = '" & vItemID & "' " & vbCrLf
				vQuery = vQuery & "UPDATE [db_academy].[dbo].[tbl_diy_item] SET dispcate1 = '" & Left(vTemp,3) & "' WHERE itemid = '" & vItemID & "'"
				dbACADEMYget.execute vQuery
			End If
			If CStr(vTemp) <> CStr(vCateCode) AND vIsDefault = "y" Then		'### 이미 y가 있는데 다른카테고리를 y로 지정할경우 일단 같은 itemid 모두 n으로 변경.
				vQuery = "UPDATE [db_academy].[dbo].[tbl_display_cate_item_Academy] SET isDefault = 'n' WHERE itemid = '" & vItemID & "'"
				vQuery = vQuery & " UPDATE [db_academy].[dbo].[tbl_diy_item] SET dispcate1 = '" & Left(vCateCode,3) & "' WHERE itemid = '" & vItemID & "'"
				dbACADEMYget.execute vQuery
			End If
		End If
		'#################################################################################################################################################################
		
		vQuery = "IF EXISTS(SELECT catecode FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "') " & vbCrLf
		vQuery = vQuery & "	BEGIN " & vbCrLf
		vQuery = vQuery & "		UPDATE [db_academy].[dbo].[tbl_display_cate_item_Academy] SET " & vbCrLf
		vQuery = vQuery & "			sortNo = '" & vSortNo & "', " & vbCrLf
		vQuery = vQuery & "			isDefault = '" & vIsDefault & "' " & vbCrLf
		vQuery = vQuery & "		WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "' " & vbCrLf
		vQuery = vQuery & "	END " & vbCrLf
		dbACADEMYget.execute vQuery

	ElseIf vAction = "delete" Then
		
		'#################################################################################################################################################################
		'isDefault = 'y' 인것을 지우려할 경우 ORDER BY depth ASC, sortno ASC 로 top 1 catecode를 기본으로 지정.
		'#################################################################################################################################################################
		If vTemp > 1 Then
			vQuery = "SELECT catecode FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE itemid = '" & vItemID & "' AND isDefault = 'y'"
			rsACADEMYget.Open vQuery,dbACADEMYget,1
			vTemp = rsACADEMYget(0)
			rsACADEMYget.close()
			If CStr(vTemp) = CStr(vCateCode) Then
				vQuery = "SELECT TOP 1 catecode FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] where itemid = '" & vItemID & "' AND isDefault = 'n' ORDER BY depth ASC, sortno ASC"
				rsACADEMYget.Open vQuery,dbACADEMYget,1
				vTemp = rsACADEMYget(0)
				rsACADEMYget.close()
				vQuery = "UPDATE [db_academy].[dbo].[tbl_display_cate_item_Academy] SET isDefault = 'y' WHERE catecode = '" & vTemp & "' AND itemid = '" & vItemID & "' " & vbCrLf
				vQuery = vQuery & "	UPDATE [db_academy].[dbo].[tbl_diy_item] SET dispcate1 = '" & Left(vTemp,3) & "' WHERE itemid = '" & vItemID & "' " & vbCrLf
				dbACADEMYget.execute vQuery
			End If
		Else
			vQuery = "UPDATE [db_academy].[dbo].[tbl_diy_item] SET dispcate1 = null WHERE itemid = '" & vItemID & "' " & vbCrLf
			dbACADEMYget.execute vQuery
		End If
		'#################################################################################################################################################################
		
		vQuery = "DELETE [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "' " & vbCrLf
		dbACADEMYget.execute vQuery

	Else
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
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->