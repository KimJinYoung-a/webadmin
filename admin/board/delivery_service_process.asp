<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%

dim mode,divcd,divname,findurl,returnURL,isusing,isTenUsing,tel
dim SQL
dim msg
mode 	=	request("mode")
divcd =	request("divcd")
findurl =	html2db(request("findurl"))
returnURL =	html2db(request("returnURL"))
divname =	html2db(request("divname"))
isusing = request("isusing")
isTenUsing = request("isTenUsing")
tel = request("tel")

'response.write mode & "<br>" & divcd & "<br>" & findurl & "<br>" & divname


on Error Resume Next

'// �ù� ��ü �ű� �Է�
if mode="add" then

	SQL = " insert into db_order.[dbo].tbl_songjang_div (divcd,tel,divname,findurl,returnURL) "&_
				" values(" &_
				"'" & divcd & "'" &_
				",'" & tel & "'" &_
				",'" & divname & "'" &_
				",'" & findurl & "'" &_
				",'" & returnURL & "'" &_
				"	) "

'// �ù� ��ü ����
elseif mode="edit" then

	'�ٹ����� ���� ��ü����� (2007�� 2�� ���� �簡���ͽ� ������)
	if isTenUsing="Y" then
	SQL = SQL + "" &_
				" update db_order.[dbo].tbl_songjang_div " &_
				" set isTenUsing='N'" &_
				" where isTenUsing='Y';"
	end if


	SQL = SQL + "" &_
				" update db_order.[dbo].tbl_songjang_div " &_
				" set divname='" & divname & "' " &_
				", findurl='" & findurl & "' " &_
				", returnURL='" & returnURL & "' " &_
				", tel='" & tel & "' " &_
				", isUsing='" & isusing & "' " &_
				", isTenUsing='" & isTenUsing & "' " &_
				" where divcd=" & divcd

end if
	dbget.beginTrans

	dbget.execute(SQL)

	if Error=0 then
		dbget.commitTrans
		msg="����Ǿ����ϴ�"
	else
		dbget.rollback
		msg="�����߻� �ٽ� �Է����ּ���."
		response.write "<script language='javascript'>alert('" & msg & "');</script>"
		dbget.close()	:	response.End
	end if

%>
<Script language="javascript">
parent.location.reload();
</script>


<!-- #include virtual="/lib/db/dbclose.asp" -->
