<%@  language="VBScript" %>
<% option explicit %>
<%
Response.CharSet = "EUC-KR"
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/itemmaster/deal/selectdealitemkeywords.asp
' Description :  �� ��ǰ - Ű���� ��������
' History : 2017.11.24 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim itemid : itemid = requestCheckVar(getNumeric(trim(Request("itemid"))),9)

Dim oitem
set oitem = new CItem
    oitem.FRectItemID = itemid

if itemid<>"" then
    oitem.GetOneItem
end if

if oitem.FTotalCount > 0 then
%>
<%= db2html(oitem.FOneItem.Fkeywords) %>
<%
end if

Set oitem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->