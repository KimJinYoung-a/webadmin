<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��Ʈ������ ���μ���
' History : 2011.01.25 ������ ����
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/partpersonCls.asp"-->
<%
Dim mode, cate1, idx, sabun, name, doc_worker, j, cc
Dim sql, sql2, sql3, sql4, midx, Fidx, sortNo
mode = requestCheckVar(request("mode"),10)
cate1 = requestCheckVar(request("category1"),50)
idx = requestCheckVar(request("idx"),10)
cc = requestCheckVar(request("cc"),10)
sabun = requestCheckVar(request("sabun"),30)
name = requestCheckVar(request("name"),30)
Fidx = requestCheckVar(request("Fidx"),30)
doc_worker = requestCheckVar(request("doc_worker"),200)
doc_worker = split(doc_worker, ",")
sortNo	= requestCheckVar(request("sortNo"),30)

If cate1 <> "" Then
	if (checkNotValidHTML(cate1) = true) Then
		response.write "<script>alert('ī�װ� �̸����� HTML�� ����Ͻ� �� �����ϴ�.');history.back();</script>"
		dbget.Close

		response.End
	End If
End If

'��ī�װ� ����� ���
If mode = "insert" Then
	sql = "insert into db_partner.dbo.tbl_partperson_category (category1, gubun, isusing) values ('"& cate1 &"', '0', 'Y')"
	dbget.execute sql
	response.write "<script>alert('��ϵǾ����ϴ�.');opener.location.reload();window.close();</script>"
End If

'��ī�װ� ������ ���
If mode = "modify" Then
	sql = "update db_partner.dbo.tbl_partperson_category set category1 = '"& cate1 &"' where idx = '"& idx &"'"
	dbget.execute sql
	response.write "<script>alert('�����Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
End If

'��ī�װ� ���� ���
If mode = "hide" Then
	sql = "update db_partner.dbo.tbl_partperson_category set isusing = 'N' where idx = '"& idx &"' or gubun='"& idx &"'"
	dbget.execute sql
	response.write "<script>alert('�����Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
End If

'��ī�װ� ����� ���
If mode = "use" Then
	sql = "update db_partner.dbo.tbl_partperson_category set isusing = 'Y' where idx = '"& idx &"' or gubun='"& idx &"'"
	dbget.execute sql
	response.write "<script>alert('�����Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
End If

'��ī�װ� ����� ���
If mode = "cinsert" Then
	sql = "insert into db_partner.dbo.tbl_partperson_category (category1, gubun, isusing) values ('"& cate1 &"', '"& idx &"', 'Y')"
	dbget.execute sql

	sql2 = "select max(idx) as midx from db_partner.dbo.tbl_partperson_category"
	rsget.open sql2,dbget,1
		midx = rsget("midx")
	rsget.close

	For j = 0 to ubound(doc_worker)
		sql3 = "insert into db_partner.dbo.tbl_partperson_category2 (cidx, category1, sabun, isusing) values ('"& midx &"', '"& idx &"', '"& doc_worker(j) &"', 'Y')"
		dbget.execute sql3
	Next

	response.write "<script>alert('��ϵǾ����ϴ�.');location.href='partcate2_pop.asp?idx="& idx &"&name="& name &"';</script>"
End If

'��ī�װ� ������ ���
If mode = "cmodify" Then
	sql = "update db_partner.dbo.tbl_partperson_category set category1 = '"& cate1 &"' where idx = '"& idx &"'"
	dbget.execute sql

	sql2 = "delete from db_partner.dbo.tbl_partperson_category2 where cidx = '"& idx &"'"
	dbget.execute sql2

	For j = 0 to ubound(doc_worker)
		sql3 = "insert into db_partner.dbo.tbl_partperson_category2 (cidx, category1, sabun, isusing) values ('"& idx &"', '"& cc &"', '"& doc_worker(j) &"', 'Y')"
		dbget.execute sql3
	Next
	response.write "<script>alert('�����Ǿ����ϴ�.');location.href='partcate2_pop.asp?idx="& Fidx &"';</script>"

End If

'��ī�װ� ���� ���
If mode = "chide" Then
	sql = "update db_partner.dbo.tbl_partperson_category2 set isusing = 'N' where sabun = '"& sabun &"'"
	dbget.execute sql
	response.write "<script>alert('�����Ǿ����ϴ�.');location.href='partcate2_pop.asp?idx="& idx &"';</script>"
End If

'��ī�װ� ����� ���
If mode = "cuse" Then
	sql = "update db_partner.dbo.tbl_partperson_category2 set isusing = 'Y' where sabun = '"& sabun &"'"
	dbget.execute sql
	response.write "<script>alert('�����Ǿ����ϴ�.');location.href='partcate2_pop.asp?idx="& idx &"';</script>"
End If

'����ڰ� ��� ������ ���� �� ���
If mode = "del" Then
	sql = "delete from db_partner.dbo.tbl_partperson_category where idx = '"& idx &"'"
	dbget.execute sql

	sql2 = "delete from db_partner.dbo.tbl_partperson_category2 where cidx = '"& idx &"'"
	dbget.execute sql2
	response.write "<script>alert('�����Ǿ����ϴ�.');window.close();</script>"
End If

'���� ����
If mode = "sortNo" Then
	sql = ""
	sql = sql & " UPDATE db_partner.dbo.tbl_partperson_category "
	sql = sql & " SET sortno = '"& sortno &"' "
	sql = sql & " WHERE idx = '"& idx &"'"
	dbget.execute sql
	response.write "<script>location.href='partcate2_pop.asp?idx="& cc &"';</script>"
End If

'��ü ���� ����
If mode = "sortNoAll" Then
	if request("idx").count()>0 then
		sql = ""
		for j=1 to request("idx").count()
			sql = sql & " UPDATE db_partner.dbo.tbl_partperson_category "
			sql = sql & " SET sortno = '"& request("sortNo")(j) &"' "
			sql = sql & " WHERE idx = '"& request("idx")(j) &"';" & vbCrLf
		next
		dbget.execute sql
	end if

	response.write "<script>location.href='partcate2_pop.asp?idx="& cc &"';</script>"
End If

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->