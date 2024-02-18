<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
'###############################################
' PageName : popDelCate.asp
' Discription : ī�װ� ����ó�� ������
' History : 2008.03.20 ������ : ���� Admin���� ����/����
'###############################################

dim cdl, cdm, cds, mode
dim sqlstr
Dim FTotalcnt, acode

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

mode = request("mode")

If mode="mdel" Then

	sqlstr = "select count(code_large) as cnt from [db_item].dbo.tbl_Cate_small"
	sqlstr = sqlstr + " where code_large='" + Cstr(cdl) + "'"
	sqlstr = sqlstr + " and code_mid='" + Cstr(cdm) + "'"
	rsget.Open sqlStr, dbget, 1
		If Not rsget.Eof Then
			FTotalcnt = rsget("cnt")
		End If
	rsget.close

	If FTotalcnt > 0 Then
		response.write "<script>alert('�� ī�װ��� ���� �ϼž߸�\n�� ī�װ��� ���� �� �� �ֽ��ϴ�.');self.close();</script>"
		dbget.close()	:	response.End
	Else
		 '// ī�װ� �Ӽ� ����
		 sqlstr = "declare @acode varchar(4000);" & vbCrLf &_
				" Select @acode = coalesce(@acode + ',', '') + convert(varchar, attrib_Code) from [db_item].dbo.tbl_Cate_Attrib_div" & vbCrLf &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "';" & vbCrLf &_
				" Delete from [db_item].dbo.tbl_Item_Attribute" &_
				" where attrib_Code in (@acode);" & vbCrLf &_
				" Delete from [db_item].dbo.tbl_Cate_Attrib_item" &_
				" where attrib_Code in (@acode);"
		 dbget.Execute(sqlStr)
		 	
		 '�Ӽ� �ɼ� �� ��ǰ�����ɼ� ����
		 sqlstr = "Delete from [db_item].dbo.tbl_Cate_Attrib_div" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'"

		'ī�װ� ���� ��ũ���� ����
		 sqlstr = sqlstr & "Delete from [db_item].dbo.tbl_Cate_RelateLink" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'"

		 dbget.Execute(sqlStr)

		 '// ��ī�װ� ����
		 sqlstr = "Delete from [db_item].dbo.tbl_Cate_mid" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'"
		 dbget.Execute(sqlStr)

		 response.write "<script>alert('�� ī�װ��� ���� �Ͽ����ϴ�.');opener.document.location.reload();self.close();</script>"
	End If
End If

if mode="sdel" Then

	'��ǰ���̺��� Ȯ��(�⺻ ī�װ�)
	sqlstr = "select count(itemid) as cnt from [db_item].dbo.tbl_item" &_
			" where cate_large='" + Cstr(cdl) + "'" &_
			" and cate_mid='" + Cstr(cdm) + "'" &_
			" and cate_small='" + Cstr(cds) + "'"
	rsget.Open sqlStr, dbget, 1
		If Not rsget.Eof Then
			FTotalcnt = rsget("cnt")
		End If
	rsget.close

	If FTotalcnt > 0 Then
		response.write "<script>alert('�̵��������� ��ǰ�� �ֽ��ϴ�.\n�⺻ ī�װ� ���� ��ǰ�� ����߸� ī�װ��� ���� �� �� �ֽ��ϴ�.');self.close();</script>"
		dbget.close()	:	response.End
	Else
		 '// �߰� ī�װ��� ��ϵ� ��ǰ ����
		 sqlstr = "Delete from db_item.dbo.tbl_Item_category " &_
		 		" where code_div='A' " &_
		 		" and code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'" & vbCrLf

		'// ī�װ� ���� ��ũ���� ����
		 sqlstr = sqlstr & "Delete from [db_item].dbo.tbl_Cate_RelateLink" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'"

		 '// ��ī�װ� ����
		 sqlstr = sqlstr & "Delete from [db_item].dbo.tbl_Cate_small" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'"
		 dbget.Execute(sqlStr)

		'ī�װ� �ߺз� ��ǰ����� �Һз� ������ ������Ʈ(2009.07.06; ������)
		dbget.execute("exec db_const.dbo.sp_Ten_MakeCategorySmallIconList")

		 response.write "<script>alert('�� ī�װ��� ���� �Ͽ����ϴ�.');opener.document.location.reload();self.close();</script>"
	End If
end if

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->