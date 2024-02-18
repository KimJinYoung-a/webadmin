<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모
' Hieditor : 2009.11.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
	Dim sql, vItemID, vUseYN, vGubun
	vGubun  = Request("gb")
	vItemID = Request("itemid")
	vUseYN  = Request("useyn")
	
	
	If vGubun = "u" Then
		sql = "UPDATE [db_momo].[dbo].[tbl_coin_manage_item] SET useyn = '" & vUseYN & "' WHERE itemid = '" & vItemID & "' "
		dbget.execute sql
	Else
		sql = "SELECT COUNT(*) FROM [db_momo].[dbo].[tbl_coin_manage_item] WHERE itemid = '" & vItemID & "'"
		rsget.Open sql, dbget ,1
		IF rsget(0) > 0 Then
			Response.Write "<script>alert('이미 저장된 상품입니다.');history.back();</script>"
			dbget.close()
			Response.End
		End If
	
		sql = "INSERT INTO [db_momo].[dbo].[tbl_coin_manage_item](itemid, useyn) VALUES('" & vItemID & "', '" & vUseYN & "') "
		dbget.execute sql
	End If
	
	Response.Redirect "pop_prod_list.asp"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->