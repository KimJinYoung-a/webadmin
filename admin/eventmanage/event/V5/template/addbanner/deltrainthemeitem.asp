<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : deltrainthemeitem.asp
' Description :  �̺�Ʈ ������ ���ø� ������ ����
' History : 2019.02.13 ������
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim sqlStr
Dim idx : idx = requestCheckVar(Request.Form("idx"),9)
Dim reUrl : reUrl = Request.ServerVariables("HTTP_REFERER")

If idx <> "" Then
dbget.beginTrans
		sqlStr = " delete FROM [db_event].[dbo].[tbl_event_multi_contents] WHERE idx=" & idx
		dbget.execute sqlStr
	IF Err.Number <> 0 THEN
		dbget.RollBackTrans 
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
		response.End 
	END IF
dbget.CommitTrans
Response.write "<script>alert('���� �Ǿ����ϴ�.');window.document.domain = '10x10.co.kr';parent.location.reload();</script>"
Else
Response.write "<script>alert('������ ����Ȯ�Ͽ� ������ �Ұ����մϴ�.');</script>"
End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->