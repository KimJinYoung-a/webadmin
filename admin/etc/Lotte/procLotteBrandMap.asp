<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	'// ���� ��� ����
	dim mode, sqlStr
	mode = Request("mode")

    '// ��ǰ��ȣ/�ɼǹ�ȣ�� �޴´� //
    dim TenMakerid, lotteBrandCd, lotteBrandNm
    TenMakerid = Request("TenMakerid")
    lotteBrandCd = Request("lotteBrandCd")
    lotteBrandNm = Request("lotteBrandNm")

	if TenMakerid="" or lotteBrandCd="" then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	end if

	'// ��庰 �б�
	Select Case mode
		Case "save"
			'��Ͽ��� Ȯ��
			sqlStr = "Select count(*) From db_item.dbo.tbl_lotte_brand_mapping Where TenMakerid='" & TenMakerid & "'"
			rsget.Open sqlStr,dbget,1
			if rsget(0)>0 then
				'����
				sqlStr = "Update db_item.dbo.tbl_lotte_brand_mapping Set " &_
					"	lotteBrandCd='" & lotteBrandCd & "'" &_
					" 	,lotteBrandNm='" & lotteBrandNm & "'" &_
					" Where TenMakerid='" & TenMakerid & "'"
				dbget.execute(sqlStr)
			else
				'�űԵ��
				sqlStr = "Insert into db_item.dbo.tbl_lotte_brand_mapping values " &_
						" ('" & TenMakerid & "'" &_
						", '" & lotteBrandCd & "','" & lotteBrandNm & "','Y', getdate()) "
				dbget.execute(sqlStr)
			end if
			rsget.Close

		Case "del"
			'��Ī�� �ٹ����� ī�װ� ����
			sqlStr = "Delete From db_item.dbo.tbl_lotte_brand_mapping " &_
					" Where TenMakerid='" & TenMakerid & "'"
			dbget.execute(sqlStr)
	End Select
%>
<script language="javascript">
alert("���������� ó���Ǿ����ϴ�.");
parent.opener.history.go(0);
parent.self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->