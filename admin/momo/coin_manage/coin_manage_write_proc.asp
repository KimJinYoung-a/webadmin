<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������
' Hieditor : 2009.11.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
	Dim sql, vIdx, vCoin, vUseYN, vRegdate
	vIdx = Request("idx")
	vCoin = Request("coin")
	vUseYN = Request("useyn")
	
	If vIdx = "" Then
		sql = "INSERT INTO [db_momo].[dbo].[tbl_coin_manage] " & _
			  "		(coin, useyn) " & _
			  "		VALUES " & _
			  "		('" & vCoin & "', '" & vUseYN & "') "
		dbget.execute sql
	Else
		sql = "UPDATE [db_momo].[dbo].[tbl_coin_manage] SET " & _
			  "		coin = '" & vCoin & "', " & _
			  "		useyn = '" & vUseYN & "' " & _
			  "	WHERE idx = '" & vIdx & "' "
		dbget.execute sql
	End If
	
	dbget.close()
	Response.Write "<script>alert('����Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
