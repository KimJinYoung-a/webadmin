<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'''doLecwait.asp ���� ����.

dim mode,arridx,arrcnt
dim lec_idx
dim refer
refer = request.ServerVariables("HTTP_REFERER")

mode	=	RequestCheckvar(request("mode"),16)
lec_idx	=	RequestCheckvar(request("lec_idx"),10)
arridx	=	trim(request("arridx"))
if arridx <> "" then
	if checkNotValidHTML(arridx) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if Right(arridx,1)="," then arridx= Left(arridx,Len(arridx)-1)

if (arridx="") then
	response.write "<script>alert('���õ� ����� �����ϴ�.');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
	dbget.close()	:	response.End
end if

if (lec_idx="") then
	response.write "<script>alert('���õ� ���°� �����ϴ�.');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
	dbget.close()	:	response.End
end if

dim sql

if mode="open" then
	''Open , ���� 1������ ����
	sql = "update [db_academy].[dbo].tbl_lec_waiting_user" ''���µ�Ͽ���� �����Ŵ
	sql = sql + " set currstate=3" + vbcrlf
	sql = sql + " ,regEndday=(convert(varchar(10),dateadd(d,1,getdate()),21) + ' 13:00:00') " + vbcrlf
	sql = sql + " ,regrank=0 " + vbcrlf
	sql = sql + " where idx in (" + arridx + ")" + vbcrlf

'response.write sql
	rsACADEMYget.open sql,dbACADEMYget,1

	''SMS ����

	'sql = "Insert into [110.93.128.72].[db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
	'sql = sql + " select distinct '010-6324-9110'," ''user_phone
	'sql = sql + " '02-741-9070',"
	'sql = sql + " '1',"
	'sql = sql + " getdate(),"
	'sql = sql + " '[�ΰŽ�]����û ���°� ���� ���� ���·� ����Ǿ����ϴ�. �α����� ���� ���ּ���.'"
	'sql = sql + " from [db_academy].[dbo].tbl_lec_waiting_user where idx in (" + arridx + ")" + vbcrlf
	'sql = sql + " and isusing='Y'"
	'sql = sql + " and currstate=3"

	'rsACADEMYget.Open sql,dbACADEMYget,1

	''��� ���� ����
	sql = " update [db_academy].dbo.tbl_lec_waiting_user"
	sql = sql + " set regrank=T.rank"
	sql = sql + " from ("
	sql = sql + " 		select idx,"
	sql = sql + " 		("
	sql = sql + " 			select count(*) from [db_academy].dbo.tbl_lec_waiting_user"
	sql = sql + " 			where lec_idx=" + CStr(lec_idx)
	sql = sql + " 			and idx not in (" + arridx + ")"
	sql = sql + " 			and idx<A.idx"
	sql = sql + " 		) + 1 as rank"
	sql = sql + " 		from  [db_academy].dbo.tbl_lec_waiting_user A"
	sql = sql + " 		where A.lec_idx=" + CStr(lec_idx)
	sql = sql + " 		and A.idx not in (" + arridx + ")"
	sql = sql + " 		) T"
	sql = sql + " where [db_academy].dbo.tbl_lec_waiting_user.lec_idx=" + CStr(lec_idx)
	sql = sql + " and [db_academy].dbo.tbl_lec_waiting_user.idx=T.idx"

	rsACADEMYget.open sql,dbACADEMYget,1

elseif mode="del" then																	''������ ��� ����Ʈ�� ������ �����ڵ� ���� ����
	sql = "update [db_academy].[dbo].tbl_lec_waiting_user"
	sql = sql + " set isusing='N'" + vbcrlf
	sql = sql + " where idx in (" + arridx + ")" + vbcrlf

	rsACADEMYget.open sql,dbACADEMYget,1
'response.write sql & "<br>"
end if

%>

<script language='javascript'>
alert('���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>



<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->