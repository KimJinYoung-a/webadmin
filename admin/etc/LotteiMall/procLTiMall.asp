<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
	'// ���� ��� ����
	dim mode, sqlStr, iErrMsg
	mode = Request("mode")

    '// ��ǰ��ȣ/�ɼǹ�ȣ�� �޴´� //
    dim dispNo, cdl, cdm, cds, itemGbnKey '', odispNo, oitemGbnKey
    dispNo  = requestCheckvar(Request("dspNo"),32)
    ''odispNo = requestCheckvar(Request("odspNo"),32)
    ''itemGbnKey = requestCheckvar(Request("itemGbnKey"),32)
    ''oitemGbnKey = requestCheckvar(Request("oitemGbnKey"),32)
    
    cdl = requestCheckvar(Request("cdl"),10)
    cdm = requestCheckvar(Request("cdm"),10)
    cds = requestCheckvar(Request("cds"),10)

    if (mode="saveCate") or (mode="delGbn") or (mode="delCate") then
    	if (dispNo="" ) or cdl="" or cdm="" or cds=""  then
    		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
    		dbget.Close: response.End
    	end if
    end if
    
	'// ��庰 �б�
	Select Case mode
		Case "saveCate"
''			'�ߺ� Ȯ�� //������
''			sqlStr = "Select cateKey From db_item.dbo.tbl_LTiMall_cateGbn_mapping "  & VbCrlf
''			sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'"  & VbCrlf
''			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
''			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
''			sqlStr = sqlStr& " 	and cateKey='" & oitemGbnKey & "'"
''			rsget.Open sqlStr,dbget,1
''			if rsget.EOF then
''				'�űԵ��
''				sqlStr = "Insert into db_item.dbo.tbl_LTiMall_cateGbn_mapping  "  & VbCrlf
''				sqlStr = sqlStr& " (tenCateLarge,tenCateMid,tenCateSmall,CateKey,lastUpdate)"
''				sqlStr = sqlStr& " values('" & cdl & "','" & cdm & "','" & cds & "','" & itemGbnKey & "', getdate()) "
''				dbget.execute(sqlStr)
''			else
''			    '������Ʈ
''			    sqlStr = "update db_item.dbo.tbl_LTiMall_cateGbn_mapping  "  & VbCrlf
''			    sqlStr = sqlStr& " set cateKey='"&itemGbnKey&"'"
''				sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'"  & VbCrlf
''    			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
''    			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
''    			sqlStr = sqlStr& " 	and cateKey='" & oitemGbnKey & "'"
''				dbget.execute(sqlStr)
''			end if
''			rsget.Close
            
            '�ߺ� Ȯ��
            sqlStr = "Select cateKey From db_item.dbo.tbl_LTiMall_cate_mapping "  & VbCrlf
			sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and cateKey='" & dispNo & "'"
			rsget.Open sqlStr,dbget,1

			if rsget.EOF then
				'�űԵ��
				sqlStr = "Insert into db_item.dbo.tbl_LTiMall_cate_mapping  "  & VbCrlf
				sqlStr = sqlStr& " (CateKey,tenCateLarge,tenCateMid,tenCateSmall,lastUpdate)"
				sqlStr = sqlStr& " values('" & dispNo & "'"  & VbCrlf
				sqlStr = sqlStr& ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
				dbget.execute sqlStr
			else
			    iErrMsg = "�̹� ���ε� ī�װ� ["&dispNo&"] �߰��� �� �����ϴ�."
			end if
			rsget.Close

		Case "delCate"
			'��Ī�� �ٹ����� ī�װ� ����
			sqlStr = "Delete From db_item.dbo.tbl_LTiMall_cate_mapping " & VbCrlf
			sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
			sqlStr = sqlStr& " 	and cateKey='" & dispNo & "'"
			dbget.execute(sqlStr)
	    Case "delGbn"
	        '��Ī�� �ٹ����� ��ǰ�з� ����
			sqlStr = "Delete From db_item.dbo.tbl_LTiMall_cateGbn_mapping " & VbCrlf
			sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
			sqlStr = sqlStr& " 	and cateKey='" & itemGbnKey & "'"
			dbget.execute(sqlStr)
	End Select
	
	if (mode="saveCate") or (mode="delCate") then
	    CALL Fn_ActOutMall_CateSummary("lotteimall")
	end if
%>
<script language="javascript">
<% if (iErrMsg<>"") then %>
alert("<%=iErrMsg %>");
<% else %>
alert("���������� ó���Ǿ����ϴ�.");
parent.opener.history.go(0);
parent.self.close();
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->