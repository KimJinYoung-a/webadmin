<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/itemmaster/deal/dodealitemreg.asp
' Description :  �� ��ǰ - ���, ����
' History : 2017.08.28 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<%
'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim k, sqlStr, i
Dim vCnt : vCnt = Request.Form("cksel").count
Dim idx : idx = requestCheckVar(Request.Form("idx"),9)
Dim stype : stype = requestCheckVar(Request.Form("stype"),1)
Dim upback : upback = requestCheckVar(Request.Form("upback"),1)
Dim mode : mode = requestCheckVar(Request.Form("mode"),1)
Dim itemidarr : itemidarr = Request.Form("itemidarr")

if Request.Form("cksel") <> "" then
	if checkNotValidHTML(Request.Form("cksel")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if Request.Form("sitemname") <> "" then
	if checkNotValidHTML(Request.Form("sitemname")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if mode="D" then
	sqlStr = "delete FROM [db_event].[dbo].[tbl_deal_event_item] WHERE dealcode=" & idx & " and itemid in (" & itemidarr & ")"
	dbget.execute sqlStr
else
	'�迭�� ó��
	redim arritemcode(vCnt)
	redim arritemname(vCnt)
	for i=1 to vCnt
		arritemcode(i) = Request.Form("cksel")(i)
		arritemname(i) = Request.Form("sitemname")(i)
	next

	If vCnt > 0 Then
	dbget.beginTrans
		For k=1 To vCnt
			sqlStr = "delete FROM [db_event].[dbo].[tbl_deal_event_item] WHERE dealcode=" & idx & " and itemid=" & arritemcode(k)
			dbget.execute sqlStr
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans 
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
			response.End 
		END IF
		Next
		dbget.CommitTrans
	End If
End If

%>
<script type="text/javascript">
parent.opener.fnLoadItems();
parent.opener.fnItemSelectboxLoad();
parent.jsDeleteItemReload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->