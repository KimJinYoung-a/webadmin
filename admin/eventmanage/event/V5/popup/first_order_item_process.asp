<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   
Response.CharSet="euc-kr"
Session.codepage="949"
Response.codepage="949"
Response.ContentType="text/html;charset=euc-kr"
'###########################################################
' Page : /admin/eventmanage/event/v5/popup/first_order_item_process.asp
' Description : 첫 구매 샵 상품 - 등록, 삭제
' History : 2023.05.09 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<%
'--------------------------------------------------------
' 변수선언 & 파라미터 값 받기
'--------------------------------------------------------
Dim k, sqlStr, i, arrItemCode, arrItemName, listimage, arrItemSort, itemCount
Dim itemid : itemid = requestCheckVar(Request("itemid"),9)
Dim itemidarr : itemidarr = Request("itemidarr")
Dim sortarr : sortarr = Request("sortarr")
Dim mode : mode = requestCheckVar(Request("mode"),9)
Dim isusing : isusing = requestCheckVar(Request("isusing"),1)

if Request.Form("itemidarr") <> "" then
	if checkNotValidHTML(Request.Form("itemidarr")) then
		response.write "err2"
		Response.End
	end if
end if

if mode="add" then
	'현재까지 등록된 아이템 수량 등록
	sqlStr = "select count(itemid) as cnt from [db_event].[dbo].[tbl_event_first_order_item]"
	sqlStr = sqlStr + " WHERE isusing = 'Y'"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        itemCount = rsget("cnt")
	else
		itemCount = 0
    end if
    rsget.close

	arrItemCode = split(itemidarr,",")
	dbget.beginTrans
	For k = 0 To UBOUND(arrItemCode)
		sqlStr = " IF Not Exists(SELECT itemid FROM [db_event].[dbo].[tbl_event_first_order_item] WHERE itemid='" & arrItemCode(k) & "')"
		sqlStr = sqlStr + "		BEGIN"
		sqlStr = sqlStr + " 		INSERT INTO [db_event].[dbo].[tbl_event_first_order_item](itemid, sort)"
		sqlStr = sqlStr + "     	VALUES (" & arrItemCode(k) & "," & itemCount+k+1 & ")"
		sqlStr = sqlStr + " 	END"
		dbget.execute sqlStr
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans 
			response.write "err3"
			response.End 
		END IF
	Next
	dbget.CommitTrans
response.write "ok"
response.End
elseif mode="sort" then
	arrItemCode = split(itemidarr,",")
	arrItemSort = split(sortarr,",")
	dbget.beginTrans
	For k = 0 To UBOUND(arrItemCode)
		sqlStr = "UPDATE [db_event].[dbo].[tbl_event_first_order_item]"
		sqlStr = sqlStr + " SET sort ='" & arrItemSort(k) & "'"
		sqlStr = sqlStr + " WHERE itemid =" & arrItemCode(k) & ""
		dbget.execute sqlStr
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans 
			response.write "err3"
			response.End 
		END IF
	Next
	dbget.CommitTrans
elseif mode="edit" then
	sqlStr = "update [db_event].[dbo].[tbl_event_first_order_item]"
    sqlStr = sqlStr + " SET isusing ='" & isusing & "'"
	sqlStr = sqlStr + " WHERE itemid =" & itemid & ""
	dbget.execute sqlStr
elseif mode="delarr" then
	sqlStr = "update [db_event].[dbo].[tbl_event_first_order_item]"
	sqlStr = sqlStr + " SET isusing ='" & isusing & "'"
	sqlStr = sqlStr + " WHERE itemid in (" & itemidarr & ")"
	dbget.execute sqlStr
end if
response.write "<script>parent.location.reload();</script>"
response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->