<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/itemmaster/deal/dodealitemgroup.asp
' Description : 딜 상품 복사
' History : 2023.01.10 정태훈 생성
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
' 변수선언 & 파라미터 값 받기
'--------------------------------------------------------
Dim k, sqlStr, i, oJson
Dim mode : mode = requestCheckVar(Request("mode"),10)
Dim idx : idx = requestCheckVar(Request("idx"),10)
Dim dealcode : dealcode = requestCheckVar(Request("dealcode"),10)

Response.ContentType = "application/json"
Set oJson = jsObject()

if Request.Form("dealcode") = "" then
    oJson("response") = "err"
    oJson("message") = "딜코드 정보가 없습니다. 다시 선택 해주세요"
    oJson.flush
    Set oJson = Nothing
    dbget.close() : Response.End
end if

if mode="copy" then
	sqlStr = "INSERT INTO [db_event].[dbo].[tbl_deal_event_item] (dealcode, itemid, itemname, viewidx)" & vbcrlf
	sqlStr = sqlStr + " SELECT " & idx & ", itemid, itemname, viewidx" & vbcrlf
    sqlStr = sqlStr + " FROM [db_event].[dbo].[tbl_deal_event_item]" & vbcrlf
    sqlStr = sqlStr + " WHERE dealcode=" & dealcode & " AND isusing='Y'"
	dbget.execute sqlStr


    oJson("response") = "ok"
    oJson("message") = "상품 복사 완료"
    oJson.flush
    Set oJson = Nothing
    dbget.close() : Response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->