<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminUPCHE.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<%
dim id,finishmemo, finishuser,songjangdiv, songjangno
id          = RequestCheckvar(request("id"),10)
finishmemo  = html2db(request("finishmemo"))
finishuser  = RequestCheckvar(request("finishuser"),32)
songjangdiv = RequestCheckvar(request("songjangdiv"),10)
songjangno  = RequestCheckvar(request("songjangno"),16)
if finishmemo <> "" then
	if checkNotValidHTML(finishmemo) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

dim sqlStr
dim oldcurrstate
sqlStr = "select currstate "
sqlStr = sqlStr + " from " & TABLE_CSMASTER & "" + VbCrlf
sqlStr = sqlStr + " where id =" + id
rsget.Open sqlStr,dbget,1
    oldcurrstate = rsget("currstate")
rsget.Close

if (oldcurrstate="B007") then
    response.write "<script>alert('�̹� ó�� �Ϸ�� �����Դϴ�. - �Ϸ�ó���� ���� �� �� �����ϴ�.');history.back();</script>"
    response.end
end if

sqlStr = "update " & TABLE_CSMASTER & "" + VbCrlf
sqlStr = sqlStr + " set finishuser ='" + finishuser + "'," + VbCrlf
sqlStr = sqlStr + " contents_finish ='" + finishmemo + "'," + VbCrlf
sqlStr = sqlStr + " songjangdiv ='" + songjangdiv + "'," + VbCrlf
sqlStr = sqlStr + " songjangno ='" + songjangno + "'," + VbCrlf
sqlStr = sqlStr + " finishdate=getdate()," + VbCrlf
sqlStr = sqlStr + " currstate='B006'" + VbCrlf
sqlStr = sqlStr + " where id =" + id
sqlStr = sqlStr + " and makerid='" & session("ssBctID") & "'"

rsget.Open sqlStr,dbget,1

%>
<script language='javascript'>
alert('����Ǿ����ϴ�.');
location.replace('upchecslist.asp');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->