<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/event_EntryList.asp
' Description :  �̺�Ʈ ������ ó�� ������
' History : 2007.09.21 ������
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->

<%

dim evtCode,arridx,selStr
'
evtCode =request("eC")
arridx =chkarray(request("arridx"))
selStr =request("selStr")

dim strSQL,msg

IF selStr = "N" THEN
'// ����
	strSQL =" DELETE FROM db_event.dbo.tbl_event_comment " &_
			" where evt_code='" & evtCode & "' " &_
			" and userid in( " &_
			" 	SELECT userid  " &_
			" 	FROM db_event.dbo.tbl_event_comment " &_
			" 	WHERE evtcom_idx in (" & arridx &")) "

	msg="���� �Ǿ����ϴ�"

ELSEIF selStr= "S" THEN
'// ����
	strSQL =" INSERT INTO [db_event].[dbo].[tbl_event_winner_log](evt_code,com_idx,userid) " &_
			" SELECT evt_code,com_idx,userid  " &_
			" FROM [db_event].[dbo].tbl_event_common_comment " &_
			" WHERE com_idx in (" & arridx &") " &_
			" and com_idx not in (select com_idx from db_event.dbo.tbl_event_winner_log where evt_code='" & evtCode & "')"
	msg="���õǾ����ϴ�."


END IF
	dbget.BeginTrans

	'response.write strsQL
	dbget.execute(strSQL)

	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)

		response.write	"<script language='javascript'>"
		response.write	"	alert('" & msg & "'); parent.location.reload();"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)

		response.write	"<script language='javascript'>"
		response.write	"	alert('ó���� ������ �߻��߽��ϴ�.');"
		response.write	"</script>"
		dbget.close()	:	response.End

	End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->