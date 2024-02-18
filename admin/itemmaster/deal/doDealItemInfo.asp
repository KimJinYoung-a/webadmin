<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/itemmaster/deal/dodealitemgroup.asp
' Description : �� ��ǰ ����
' History : 2023.01.10 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim k, sqlStr, i, oJson, itemCount, groupCount
Dim mode : mode = requestCheckVar(Request("mode"),10)
Dim idx : idx = requestCheckVar(Request("idx"),10)
Dim dealcode : dealcode = requestCheckVar(Request("dealcode"),10)

Response.ContentType = "application/json"
Set oJson = jsObject()

if Request("idx") = "" then
    oJson("response") = "err"
    oJson("message") = "���ڵ� ������ �����ϴ�. �ٽ� ���� ���ּ���"
    oJson.flush
    Set oJson = Nothing
    dbget.close() : Response.End
end if

'������� ��ϵ� ������ ����
sqlStr = "select count(idx) as cnt from [db_event].[dbo].[tbl_deal_event_item]"
sqlStr = sqlStr + " WHERE isusing='Y' and dealcode = '" & idx & "'"
rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    itemCount = rsget("cnt")
else
    itemCount = 0
end if
rsget.close

'������� ��ϵ� �׷� ����
sqlStr = "select count(group_code) as cnt from [db_event].[dbo].[tbl_deal_event_group]"
sqlStr = sqlStr + " WHERE isusing='Y' and deal_code = '" & idx & "'"
rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    groupCount = rsget("cnt")
else
    groupCount = 0
end if
rsget.close

oJson("response") = "ok"
oJson("itemCount") = itemCount
oJson("groupCount") = groupCount
oJson.flush
Set oJson = Nothing
dbget.close() : Response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->