<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/CategoryCls.asp"-->
<%
'###############################################
' PageName : popDelCate.asp
' Discription : ī�װ� ����ó�� ������
' History : 2008.03.20 ������ : ���� Admin���� ����/����
' History : 2012.08.17 ����ȭ : ���� Admin���� ����/����
'###############################################

dim cdl, cdm, cds, mode
dim sqlstr
Dim FTotalcnt, acode

cdl = RequestCheckvar(request("cdl"),10)
cdm = RequestCheckvar(request("cdm"),10)
cds = RequestCheckvar(request("cds"),10)

mode = RequestCheckvar(request("mode"),16)

If mode="mdel" Then

	sqlstr = "select count(code_large) as cnt from [db_academy].dbo.tbl_diy_item_Cate_small"
	sqlstr = sqlstr + " where code_large='" + Cstr(cdl) + "'"
	sqlstr = sqlstr + " and code_mid='" + Cstr(cdm) + "'"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		If Not rsACADEMYget.Eof Then
			FTotalcnt = rsACADEMYget("cnt")
		End If
	rsACADEMYget.close

	If FTotalcnt > 0 Then
		response.write "<script>alert('�� ī�װ��� ���� �ϼž߸�\n�� ī�װ��� ���� �� �� �ֽ��ϴ�.');self.close();</script>"
		dbACADEMYget.close()	:	response.End
	Else
		 '// ��ī�װ� ����
		 sqlstr = "Delete from [db_academy].dbo.tbl_diy_item_Cate_mid" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'"
		 dbACADEMYget.Execute(sqlStr)

		 response.write "<script>alert('�� ī�װ��� ���� �Ͽ����ϴ�.');opener.document.location.reload();self.close();</script>"
	End If
End If

if mode="sdel" Then

	'��ǰ���̺��� Ȯ��(�⺻ ī�װ�)
	sqlstr = "select count(itemid) as cnt from [db_academy].dbo.tbl_diy_item" &_
			" where cate_large='" + Cstr(cdl) + "'" &_
			" and cate_mid='" + Cstr(cdm) + "'" &_
			" and cate_small='" + Cstr(cds) + "'"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		If Not rsACADEMYget.Eof Then
			FTotalcnt = rsACADEMYget("cnt")
		End If
	rsACADEMYget.close

	If FTotalcnt > 0 Then
		response.write "<script>alert('�̵��������� ��ǰ�� �ֽ��ϴ�.\n�⺻ ī�װ� ���� ��ǰ�� ����߸� ī�װ��� ���� �� �� �ֽ��ϴ�.');self.close();</script>"
		dbACADEMYget.close()	:	response.End
	Else
		 '// �߰� ī�װ��� ��ϵ� ��ǰ ����
		 sqlstr = "Delete from [db_academy].dbo.tbl_diy_item_category " &_
		 		" where code_div='A' " &_
		 		" and code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'" & vbCrLf

		 '// ��ī�װ� ����
		 sqlstr = sqlstr & "Delete from [db_academy].dbo.tbl_diy_item_Cate_small" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'"
		 dbACADEMYget.Execute(sqlStr)

		'ī�װ� �ߺз� ��ǰ����� �Һз� ������ ������Ʈ(2009.07.06; ������)
		'dbACADEMYget.execute("exec db_const.dbo.sp_Ten_MakeCategorySmallIconList")

		 response.write "<script>alert('�� ī�װ��� ���� �Ͽ����ϴ�.');opener.document.location.reload();self.close();</script>"
	End If
end if

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->