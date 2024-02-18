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
' Page : /admin/itemmaster/deal/dodealitemreg.asp
' Description :  딜 상품 - 등록, 삭제
' History : 2017.08.28 정태훈 생성
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
Dim idx : idx = requestCheckVar(Request("idx"),9)
Dim itemid : itemid = requestCheckVar(Request("itemid"),9)
Dim itemidarr : itemidarr = Request("itemidarr")
Dim sortarr : sortarr = Request("sortval")
Dim sortdiv : sortdiv = Request("sortdiv")
Dim sitemnamearr : sitemnamearr = unescape(Trim(Request("sitemnamearr")))
Dim mode : mode = requestCheckVar(Request("mode"),9)

Dim dealcode : dealcode = requestCheckVar(Request("dealcode"),10)
Dim group_code : group_code = requestCheckVar(Request("group_code"),10)

'마스터키가 없음?? 종료
if idx="" then
	response.write "err1"
	Response.End
end if

if Request.Form("itemidarr") <> "" then
	if checkNotValidHTML(Request.Form("itemidarr")) then
		response.write "err2"
		Response.End
	end if
end if

if Request.Form("sitemnamearr") <> "" then
	if checkNotValidHTML(Request.Form("sitemnamearr")) then
		response.write "err2"
		Response.End
	end if
end if

if mode="add" then
	'현재까지 등록된 아이템 수량 등록
	sqlStr = "select count(idx) as cnt from [db_event].[dbo].[tbl_deal_event_item]"
	sqlStr = sqlStr + " WHERE dealcode = '" & idx & "'"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        itemCount = rsget("cnt")
	else
		itemCount = 0
    end if
    rsget.close

	arrItemCode = split(itemidarr,",")
	arrItemName = split(sitemnamearr,"|")
	dbget.beginTrans
	For k = 0 To UBOUND(arrItemCode)
		sqlStr = " IF Not Exists(SELECT IDX FROM [db_event].[dbo].[tbl_deal_event_item] WHERE itemid='" & arrItemCode(k) & "' and dealcode=" & idx & ")"
		sqlStr = sqlStr + "		BEGIN"
		sqlStr = sqlStr + " 		INSERT INTO [db_event].[dbo].[tbl_deal_event_item](dealcode, itemid, itemname, viewidx, group_code)"
		sqlStr = sqlStr + "     	VALUES (" & idx & ", " & arrItemCode(k) & ",'" & arrItemName(k) & "'," & itemCount+k+1 & "," & group_code & ")"
		sqlStr = sqlStr + " 	END"
		dbget.execute sqlStr
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans 
			response.write "err3"
			response.End 
		END IF
	Next
	dbget.CommitTrans
elseif mode="view" then
	arrItemCode = split(itemidarr,",")
	arrItemSort = split(sortarr,",")
	dbget.beginTrans
	For k = 0 To UBOUND(arrItemCode)
		sqlStr = "UPDATE [db_event].[dbo].[tbl_deal_event_item]"
		sqlStr = sqlStr + " SET viewidx ='" & arrItemSort(k) & "'"
		sqlStr = sqlStr + " WHERE dealcode = '" & idx & "' "
		sqlStr = sqlStr + " and itemid =" & arrItemCode(k) & ""
		dbget.execute sqlStr
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans 
			response.write "err3"
			response.End 
		END IF
	Next
	dbget.CommitTrans
elseif mode="del" then
	sqlStr = "delete from [db_event].[dbo].[tbl_deal_event_item]"
	sqlStr = sqlStr + " WHERE dealcode = '" & idx & "'"
	sqlStr = sqlStr + " and itemid =" & itemid & ""
	dbget.execute sqlStr
elseif mode="delarr" then
	sqlStr = "delete from [db_event].[dbo].[tbl_deal_event_item]"
	sqlStr = sqlStr + " WHERE dealcode = '" & idx & "'"
	sqlStr = sqlStr + " and itemid in (" & itemidarr & ")"
	dbget.execute sqlStr
end if
response.write "ok"
response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->