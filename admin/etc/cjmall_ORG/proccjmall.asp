<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// ���� ��� ����
Dim mode, sqlStr, iErrMsg
mode = Request("mode")

'// ��ǰ��ȣ/�ɼǹ�ȣ�� �޴´� //
Dim dispNo, cdl, cdm, cds
dispNo	= requestCheckvar(Request("dspNo"),32)
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)

If (mode = "saveCate") OR (mode = "delGbn") OR (mode = "delCate") Then
	If (dispNo = "" ) OR cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	End If
End If

'// ��庰 �б�
Select Case mode
	Case "saveCate"
        '�ߺ� Ȯ��
        sqlStr = "Select cateKey From db_item.dbo.tbl_cjMall_cate_mapping "  & VbCrlf
		sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and cateKey='" & dispNo & "'"
		rsget.Open sqlStr,dbget,1

		If rsget.EOF Then
			'�űԵ��
			sqlStr = ""
			sqlStr = sqlStr & " Insert into db_item.dbo.tbl_cjMall_cate_mapping  " & VbCrlf
			sqlStr = sqlStr & " (CateKey, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & dispNo & "'"  & VbCrlf
			sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
			dbget.execute sqlStr
		Else
		    iErrMsg = "�̹� ���ε� ī�װ� ["&dispNo&"] �߰��� �� �����ϴ�."
		End If
		rsget.Close

	Case "delCate"
		'��Ī�� �ٹ����� ī�װ� ����
		sqlStr = "Delete From db_item.dbo.tbl_cjMall_cate_mapping " & VbCrlf
		sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and cateKey='" & dispNo & "'"
		dbget.execute(sqlStr)
End Select

If (mode="saveCate") or (mode="delCate") then
    CALL Fn_ActOutMall_CateSummary("cjmall")
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