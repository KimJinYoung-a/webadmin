<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// ���� ��� ����
Dim mode, sqlStr, iErrMsg, categbn, brandcode, makerid, cateCnt
mode = Request("mode")
categbn = Request("categbn")

If (categbn <> "cate" and categbn <> "brand") Then
	response.write "<script>alert('�߸��� ����Դϴ�');window.close();</script>"
	response.end
End If

Dim cdl, cdm, cds, large_category, middle_category, small_category, detail_category
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
large_category	= requestCheckVar(request("large_category"),10)
middle_category	= requestCheckVar(request("middle_category"),10)
small_category	= requestCheckVar(request("small_category"),10)
detail_category	= requestCheckVar(request("detail_category"),10)

If (mode = "saveCate") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	End If
End If

'// ��庰 �б�
Select Case mode
	Case "saveCate"
		If detail_category = "0" Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
			sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_wetoo1300k_category "
			sqlStr = sqlStr & " WHERE large_category = '"& large_category &"' "
			sqlStr = sqlStr & " and middle_category = '"& middle_category &"' "
			sqlStr = sqlStr & " and small_category = '"& small_category &"' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				cateCnt = rsget("cnt")
			rsget.Close

			If cateCnt <> 1 Then
				Call Alert_return("����ī�װ��� �����մϴ�.\n\n�ٸ� ī�װ��� ��Ī�ϼ���") 
				dbget.close()	:	response.End
			End If
		End If


		sqlStr = "DELETE FROM db_etcmall.dbo.tbl_wetoo1300k_cate_mapping " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_wetoo1300k_cate_mapping  " & VbCrlf
		sqlStr = sqlStr & " (large_category, middle_category, small_category, detail_category, tenCateLarge, tenCateMid, tenCateSmall, lastupdate)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & large_category & "', '" & middle_category & "', '" & small_category & "', '" & detail_category & "' "  & VbCrlf
		sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
		dbget.execute sqlStr
	Case "delCate"
		sqlStr = "DELETE FROM db_etcmall.dbo.[tbl_wetoo1300k_brandcode] " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute(sqlStr)
	Case "savebrandcode"
		brandcode	= Trim(request("brandcode"))
		makerid		= Trim(request("makerid"))

		sqlStr = " DELETE FROM db_etcmall.dbo.[tbl_wetoo1300k_brandcode] WHERE makerid ='"& makerid &"' " & VbCrlf
		dbget.execute(sqlStr)

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.[tbl_wetoo1300k_brandcode] (makerid, brandCode, regdate) " & VbCrlf
		sqlStr = sqlStr & " SELECT userid, '"& brandCode &"', GETDATE() " & VbCrlf
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c " & VbCrlf
		sqlStr = sqlStr & " WHERE userid = '"& makerid &"' " 
		dbget.execute(sqlStr)
	Case "delbrandcode"
		makerid		= Trim(request("makerid"))
		brandcode	= Trim(request("brandcode"))
		sqlStr = " DELETE FROM db_etcmall.dbo.[tbl_wetoo1300k_brandcode] WHERE makerid ='"& makerid &"' and brandcode = '"& brandcode &"' " & VbCrlf
		dbget.execute(sqlStr)
End Select
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
	alert("<%=iErrMsg %>");
	history.back(-1);
<% Else %>
	alert("���������� ó���Ǿ����ϴ�.");
	<% If mode = "savebrandcode" or mode = "delbrandcode" Then %>
	location.replace('/admin/etc/wetoo1300k/popwetoo1300kBrandList.asp?brandcode=<%= brandCode %>');
	<% Else %>
	parent.opener.history.go(0);
	parent.self.close();
	<% End If %>
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->