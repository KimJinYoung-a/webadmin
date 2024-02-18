<%@  language="VBScript" %>
<% option explicit %>
<%
Response.CharSet = "EUC-KR"
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/itemmaster/deal/selectdealitemkeywords.asp
' Description :  딜 상품 - 키워드 가져오기
' History : 2017.11.24 정태훈 생성
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
' 변수선언 & 파라미터 값 받기
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