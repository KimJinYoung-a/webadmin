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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
	Dim sql, vMngIdx, vIdx, vType, vItem, vOption, vUseYN, vItemDesc, vTmp
	vMngIdx = Request("mng_idx")
	vIdx = Request("idx")
	vType = Request("type")
	vItem = Request("item")
	vOption = Request("option")
	vUseYN = Request("useyn")
	vItemDesc = Request("item_desc")
	vTmp = 0
	
	If vItem = "해당 아이템 설명" OR vMngIdx = "" OR vType = "" OR vItem = "" OR vUseYN = "" OR vItemDesc = "" Then
		Response.Write "잘못 입력함."
		Response.End
	End If
	
	If vType = "i" OR vType = "s" Then
		If vOption = "" Then
			sql = "SELECT count(*) FROM [db_item].[dbo].[tbl_item_option] WHERE isusing = 'Y' AND itemid = '" & vItem & "' "
			rsget.Open sql, dbget ,1
			vTmp = rsget(0)
			rsget.Close
			If vTmp > 0 Then
				Response.Write "<script>alert('옵션코드가 있는 상품입니다. 반드시 옵션코드를 입력해야합니다.');history.back();</script>"
				dbget.close()
				Response.End
			End If
		End IF
	End IF
	
	If vIdx = "" Then
		sql = "INSERT INTO [db_momo].[dbo].[tbl_coin_manage_prod] " & _
			  "		(mng_idx, type, prod, prod_option, prod_desc, useyn) " & _
			  "		VALUES " & _
			  "		('" & vMngIdx & "', '" & vType & "', '" & vItem & "', '" & vOption & "', '" & vItemDesc & "', '" & vUseYN & "') "
		dbget.execute sql
	Else
		sql = "UPDATE [db_momo].[dbo].[tbl_coin_manage_prod] SET " & _
			  "		type = '" & vType & "', " & _
			  "		prod = '" & vItem & "', " & _
			  "		prod_option = '" & vOption & "', " & _
			  "		prod_desc = '" & vItemDesc & "', " & _
			  "		useyn = '" & vUseYN & "' " & _
			  "	WHERE idx = '" & vIdx & "' "
		dbget.execute sql
	End If
	
	dbget.close()
	Response.Write "<script>alert('저장되었습니다.');location.href='coin_manage_item.asp?mng_idx="&vMngIdx&"';</script>"
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
