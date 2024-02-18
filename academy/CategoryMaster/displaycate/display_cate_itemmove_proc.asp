<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/CategoryMaster/displaycate/classes/displaycateCls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim vQuery, vCateCode_A, vCateCode_B, vDepth, vOnlyThisCate, vCateCodeQuery, vTempItem, vChangeContents, vSCMChangeSQL
	vCateCode_A		= RequestCheckvar(Request("catecode_a"),10)
	vCateCode_B		= RequestCheckvar(Request("catecode_b"),10)
	vDepth			= (Len(vCateCode_B)/3)
	vOnlyThisCate	= RequestCheckvar(Request("onlythiscate"),1)

	If vOnlyThisCate = "N" Then
		vCateCodeQuery = "Left(i.catecode," & Len(vCateCode_A) & ") = '" & vCateCode_A & "'"
	Else
		vCateCodeQuery = "i.catecode = '" & vCateCode_A & "'"
	End If
	
	'####### _A -------> 이동해야할곳
	'####### _B -------> 옮겨질곳
	
	If vCateCode_A = "" OR vCateCode_B = "" OR vOnlyThisCate = "" Then
		dbACADEMYget.close()
		Response.End
	End If
	
	'######### n은 어차피 상관이 없고 y인것들만 itemid 골라내서 한꺼번에 [tbl_item] dispcate1 업데이트.
	vQuery = "select itemid, isDef from (SELECT i.itemid, " & vbCrLf & _
			 "			(select case when count(itemid) > 0 then 'n' else 'y' end from [db_academy].[dbo].[tbl_display_cate_item_Academy] " & vbCrLf & _
			 "	 		where itemid = i.itemid and catecode <> '" & vCateCode_A & "' and isDefault = 'y') as isDef " & vbCrLf & _
			 "		FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] as i " & vbCrLf & _
			 "		WHERE " & vCateCodeQuery & " and i.itemid not in(select itemid from [db_academy].[dbo].[tbl_display_cate_item_Academy] where catecode = '" & vCateCode_B & "') " & vbCrLf & _
			 "		GROUP BY i.itemid " & vbCrLf & _
			 "		) as a where isDef = 'y' "
	rsACADEMYget.Open vQuery,dbACADEMYget,1
	Do Until rsACADEMYget.Eof
		vTempItem = vTempItem & rsACADEMYget("itemid") & ","
	rsACADEMYget.MoveNext
	Loop
	rsACADEMYget.close()
	If vTempItem <> "" Then
		vTempItem = Left(vTempItem,Len(vTempItem)-1)
		vQuery = "UPDATE [db_academy].[dbo].[tbl_diy_item] SET dispcate1 = '" & Left(vCateCode_B,3) & "' WHERE itemid IN(" & vTempItem & ")"
		dbACADEMYget.execute vQuery
	End IF
	
	
	vQuery = "INSERT INTO [db_academy].[dbo].[tbl_display_cate_item_Academy]" & vbCrLf & _
			 "		SELECT '" & vCateCode_B & "', i.itemid, '" & vDepth & "', 9999, " & vbCrLf & _
			 "			(select case when count(itemid) > 0 then 'n' else 'y' end from [db_academy].[dbo].[tbl_display_cate_item_Academy] " & vbCrLf & _
			 "	 		where itemid = i.itemid and catecode <> '" & vCateCode_A & "' and isDefault = 'y') " & vbCrLf & _
			 "		FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] as i " & vbCrLf & _
			 "		WHERE " & vCateCodeQuery & " and i.itemid not in(select itemid from [db_academy].[dbo].[tbl_display_cate_item_Academy] where catecode = '" & vCateCode_B & "') " & vbCrLf & _
			 "		GROUP BY i.itemid"
	dbACADEMYget.execute vQuery
	
	vQuery = "DELETE [db_academy].[dbo].[tbl_display_cate_item_Academy] WHERE " & Replace(vCateCodeQuery,"i.","") & ""
	dbACADEMYget.execute vQuery
%>

<script>
parent.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->