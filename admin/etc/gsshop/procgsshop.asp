<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// ���� ��� ����
Dim mode, sqlStr, iErrMsg, categbn
mode = Request("mode")
'// ��ǰ��ȣ/�ɼǹ�ȣ�� �޴´� //
Dim dispNo, cdl, cdm, cds, safecode
dispNo	= requestCheckvar(Request("dspNo"),32)
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
safecode	= requestCheckvar(Request("safecode"),10)
categbn = requestCheckvar(Request("categbn"),1)

If (mode = "saveCate") OR (mode = "delGbn") OR (mode = "delCate") OR (mode = "delNewPrddiv") OR (mode = "saveNewPrdDiv") Then
	If (dispNo = "" ) OR cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	End If
End If

'// ��庰 �б�
Select Case mode
	Case "saveCate"
		'��Ī�� ���ñ��� ����
		sqlStr = ""
		sqlStr = sqlStr & " Delete From db_item.dbo.tbl_gsshop_cate_mapping " & VbCrlf
		sqlStr = sqlStr & " Where tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr & " and categbn ='" & categbn & "'"
		dbget.execute(sqlStr)

		sqlStr = ""
		sqlStr = sqlStr & " Insert into db_item.dbo.tbl_gsshop_cate_mapping  " & VbCrlf
		sqlStr = sqlStr & " (CateKey, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate, categbn)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & dispNo & "'"  & VbCrlf
		sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate(), '"& categbn &"') "
		dbget.execute sqlStr
	Case "delCate"
		'��Ī�� �ٹ����� ī�װ� ����
		sqlStr = "Delete From db_item.dbo.tbl_gsshop_cate_mapping " & VbCrlf
		sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and cateKey='" & dispNo & "'"
		dbget.execute(sqlStr)

	Case "saveNewPrdDiv"
        '�ߺ� Ȯ��
        sqlStr = ""
		sqlStr = sqlStr & " SELECT dtlCd FROM db_item.dbo.tbl_gsshop_MngDiv_mapping "  & VbCrlf
		sqlStr = sqlStr & " WHERE tenCateLarge='" & cdl & "'"  & VbCrlf
		sqlStr = sqlStr & " and tenCateMid='" & cdm & "'"  & VbCrlf
		sqlStr = sqlStr & " and tenCateSmall='" & cds & "'"  & VbCrlf
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If rsget.EOF Then
			'�űԵ��
			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_gsshop_MngDiv_mapping  " & VbCrlf
			sqlStr = sqlStr & " (dtlCd, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate, safecode)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & dispNo & "'"  & VbCrlf
			sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate(),'" & safecode & "') "
			dbget.execute sqlStr
		Else
		    iErrMsg = "�̹� ���ε� ��ǰ�з��� ["&dispNo&"] �߰��� �� �����ϴ�."
		End If
		rsget.Close

	Case "delNewPrddiv"
		'��Ī�� �ٹ����� ī�װ� ����
		sqlStr = ""
		sqlStr = sqlStr & " DELETE FROM db_item.dbo.tbl_gsshop_MngDiv_mapping " & VbCrlf
		sqlStr = sqlStr & " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr & " and dtlCd = '" & dispNo & "'"
		dbget.execute(sqlStr)
End Select

If (mode="saveCate") or (mode="delCate") then
    CALL Fn_ActOutMall_CateSummary("gsshop")
End If
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
alert("���������� ó���Ǿ����ϴ�.");
parent.opener.history.go(0);
parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->