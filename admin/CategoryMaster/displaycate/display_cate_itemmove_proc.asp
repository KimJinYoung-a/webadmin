<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->

<%
	Dim vQuery, vCateCode_A, vCateCode_B, vDepth, vOnlyThisCate, vCateCodeQuery, vTempItem, vChangeContents, vSCMChangeSQL
	vCateCode_A		= Request("catecode_a")
	vCateCode_B		= Request("catecode_b")
	vDepth			= (Len(vCateCode_B)/3)
	vOnlyThisCate	= Request("onlythiscate")
	
	vChangeContents = "- ����ī�װ� ����ǰ �̵� : " & vCateCode_A & " ---> " & vCateCode_B & "" & vbCrLf
	
	If vOnlyThisCate = "N" Then
		vCateCodeQuery = "Left(i.catecode," & Len(vCateCode_A) & ") = '" & vCateCode_A & "'"
		vChangeContents = vChangeContents & "- �̵��� ī�װ� ���� ���� ��ǰ ���� �̵�" & vbCrLf
	Else
		vCateCodeQuery = "i.catecode = '" & vCateCode_A & "'"
		vChangeContents = vChangeContents & "- �̵��� ī�װ� ��ǰ�� �̵�" & vbCrLf
	End If
	
	'####### _A -------> �̵��ؾ��Ұ�
	'####### _B -------> �Ű�����
	
	If vCateCode_A = "" OR vCateCode_B = "" OR vOnlyThisCate = "" Then
		dbget.close()
		Response.End
	End If
	
	'######### n�� ������ ����� ���� y�ΰ͵鸸 itemid ��󳻼� �Ѳ����� [tbl_item] dispcate1 ������Ʈ.
	vQuery = "select itemid, isDef from (SELECT i.itemid, " & vbCrLf & _
			 "			(select case when count(itemid) > 0 then 'n' else 'y' end from [db_item].[dbo].[tbl_display_cate_item] " & vbCrLf & _
			 "	 		where itemid = i.itemid and catecode <> '" & vCateCode_A & "' and isDefault = 'y') as isDef " & vbCrLf & _
			 "		FROM [db_item].[dbo].[tbl_display_cate_item] as i " & vbCrLf & _
			 "		WHERE " & vCateCodeQuery & " and i.itemid not in(select itemid from [db_item].[dbo].[tbl_display_cate_item] where catecode = '" & vCateCode_B & "') " & vbCrLf & _
			 "		GROUP BY i.itemid " & vbCrLf & _
			 "		) as a where isDef = 'y' "
	rsget.Open vQuery,dbget,1
	Do Until rsget.Eof
		vTempItem = vTempItem & rsget("itemid") & ","
	rsget.MoveNext
	Loop
	rsget.close()
	If vTempItem <> "" Then
		vTempItem = Left(vTempItem,Len(vTempItem)-1)
		vQuery = "UPDATE [db_item].[dbo].[tbl_item] SET dispcate1 = '" & Left(vCateCode_B,3) & "' WHERE itemid IN(" & vTempItem & ")"
		dbget.execute vQuery
		
		vChangeContents = vChangeContents & "- ��� ��ǰ : " & vTempItem & vbCrLf
	End IF
	
	
	vQuery = "INSERT INTO [db_item].[dbo].[tbl_display_cate_item]" & vbCrLf & _
			 "		SELECT '" & vCateCode_B & "', i.itemid, '" & vDepth & "', 9999, " & vbCrLf & _
			 "			(select case when count(itemid) > 0 then 'n' else 'y' end from [db_item].[dbo].[tbl_display_cate_item] " & vbCrLf & _
			 "	 		where itemid = i.itemid and catecode <> '" & vCateCode_A & "' and isDefault = 'y') " & vbCrLf & _
			 "		FROM [db_item].[dbo].[tbl_display_cate_item] as i " & vbCrLf & _
			 "		WHERE " & vCateCodeQuery & " and i.itemid not in(select itemid from [db_item].[dbo].[tbl_display_cate_item] where catecode = '" & vCateCode_B & "') " & vbCrLf & _
			 "		GROUP BY i.itemid"
	dbget.execute vQuery
	
	vQuery = "DELETE [db_item].[dbo].[tbl_display_cate_item] WHERE " & Replace(vCateCodeQuery,"i.","") & ""
	dbget.execute vQuery
	
	
	'### ���� �α� ����(dispcate)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'dispcate', '" & Request("menupos") & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)
%>

<script>
parent.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->