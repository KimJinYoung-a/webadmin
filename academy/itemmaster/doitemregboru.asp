<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ��ǰ ��� ��� ��ǰ 
' Hieditor : 2010.10.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode,idx,rejectmsg ,sqlStr
	mode = RequestCheckvar(request("mode"),16)
	idx = RequestCheckvar(request("idx"),10)
	rejectmsg = html2Db(request("rejectmsg"))
  	if rejectmsg <> "" then
		if checkNotValidHTML(rejectmsg) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
if mode="waitstate" then
	sqlStr = "update [db_academy].dbo.tbl_diy_wait_item"
	sqlStr = sqlStr + " set currstate='1'"
	sqlStr = sqlStr + " where itemid=" + idx
	
	'response.write sqlStr &"<br>"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	
elseif mode="delstate" then
    ''�������
	sqlStr = "update [db_academy].dbo.tbl_diy_wait_item"
	sqlStr = sqlStr + " set currstate='0'"
	sqlStr = sqlStr + " ,rejectmsg='" + rejectmsg + "'"
	sqlStr = sqlStr + " ,rejectDate=getdate() "
	sqlStr = sqlStr + " where itemid=" + idx
	
	'response.write sqlStr &"<br>"
	dbACADEMYget.execute(sqlStr)
	
	''2016/13/08 �߰�
	sqlStr = "exec db_academy.[dbo].[sp_ACA_sendPushMsgItemConfirm_Artist] "&idx&",NULL"
	dbACADEMYget.execute(sqlStr)
	
else
    ''��Ϻ���(���û)
	sqlStr = "update [db_academy].dbo.tbl_diy_wait_item"
	sqlStr = sqlStr + " set currstate='2'"
	sqlStr = sqlStr + " ,rejectmsg='" + rejectmsg + "'"
	sqlStr = sqlStr + " ,rejectDate=getdate() "
	sqlStr = sqlStr + " where itemid=" + idx

	'response.write sqlStr &"<br>"
	dbACADEMYget.execute(sqlStr)
	
	''2016/13/08 �߰�
	sqlStr = "exec db_academy.[dbo].[sp_ACA_sendPushMsgItemConfirm_Artist] "&idx&",NULL"
	dbACADEMYget.execute(sqlStr)
end if
%>

<script language="JavaScript">

<% if mode="waitstate" then %>

	alert("��� ���� ���� �Ǿ����ϴ�.");
	history.go(-1);

<% elseif mode="delstate" then %>

	alert("������� ���� ���� �Ǿ����ϴ�.");
	opener.location.reload();
	window.close();

<% else %>

	alert("��� ����(���û) �Ǿ����ϴ�.");
	opener.location.reload();
	window.close();

<% end if %>
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->