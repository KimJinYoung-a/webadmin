<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
	'// ���� ��� ����
	dim mode, sqlStr
	mode = Request("mode")

    '// ��ǰ��ȣ/�ɼǹ�ȣ�� �޴´� //
    dim dispNo, cdl, cdm, cds
    dispNo = Request("dspNo")
    cdl = Request("cdl")
    cdm = Request("cdm")
    cds = Request("cds")

	if dispNo="" or cdl="" or cdm="" or cds="" then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	end if

	'// ��庰 �б�
	Select Case mode
		Case "save"
			'�ߺ� Ȯ��
			sqlStr = "Select DispNo From db_item.dbo.tbl_lotte_cate_mapping " &_
					" Where tenCateLarge='" & cdl & "'" &_
					" 	and tenCateMid='" & cdm & "'" &_
					" 	and tenCateSmall='" & cds & "'" &_
					" 	and DispNo='" & dispNo & "'"
			rsget.Open sqlStr,dbget,1
			if rsget.EOF then
				'�űԵ��
				sqlStr = "Insert into db_item.dbo.tbl_lotte_cate_mapping values " &_
						" ('" & dispNo & "'" &_
						", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
				dbget.execute(sqlStr)
			end if
			rsget.Close

		Case "del"
			'��Ī�� �ٹ����� ī�װ� ����
			sqlStr = "Delete From db_item.dbo.tbl_lotte_cate_mapping " &_
					" Where tenCateLarge='" & cdl & "'" &_
					" 	and tenCateMid='" & cdm & "'" &_
					" 	and tenCateSmall='" & cds & "'" &_
					" 	and DispNo='" & dispNo & "'"
			dbget.execute(sqlStr)
	End Select
	
	if (mode="save") or (mode="del") then
	    CALL Fn_ActOutMall_CateSummary("lotteCom")
	end if
%>
<script language="javascript">
alert("���������� ó���Ǿ����ϴ�.");
parent.opener.history.go(0);
parent.self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->