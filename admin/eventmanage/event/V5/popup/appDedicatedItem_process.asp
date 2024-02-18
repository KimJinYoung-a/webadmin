<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : appDedicatedItem_process.asp
' Discription : 앱전용 응모템 아이템 설정 프로세스
' History : 2023.02.07 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
dim eCode, mode, strSql, cnt, PrizeUserid, ix
dim episode, itemid, startdate, enddate, prizedate, itemidarr, prizearr
dim prize_count, prize_count_color, prizetime

eCode = requestCheckVar(Request.Form("evt_code"),10)
mode = requestCheckVar(Request.Form("mode"),10)
episode = requestCheckVar(Request.Form("episode"),2)
itemid = requestCheckVar(Request.Form("itemid"),10)
startdate = requestCheckVar(Request.Form("startdate"),10)
enddate = requestCheckVar(Request.Form("enddate"),10)
prizedate = requestCheckVar(Request.Form("prizedate"),10)
prizetime = requestCheckVar(Request.Form("prizetime"),2)
prize_count = requestCheckVar(Request.Form("prize_count"),3)
prize_count_color = requestCheckVar(Request.Form("prize_count_color"),8)
enddate = enddate & " 23:59:59"
itemidarr = Request.Form("itemidarr")
prizearr = Request.Form("prizearr")
if eCode="" then
    response.write "<script type='text/javascript'>"
    response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
    response.write "</script>"
    response.End
end if

prizedate = prizedate & " " & Num2Str(prizetime,2,"0","R") & ":00:00"

if mode="add" then
    strSql = "INSERT INTO [db_event].[dbo].[tbl_event_app_exclusive_episode](evt_code,episode,itemid,start_date,end_date,prize_date,prize_count,prize_count_color)" & vbCrlf
    strSql = strSql + " VALUES(" & eCode & "," & episode & "," & itemid & ",'" & startdate & "','" & enddate & "','" & prizedate & "'," & prize_count & ",'" & prize_count_color & "')"
    dbget.execute strSql
elseif mode="del" then
    strSql = "UPDATE [db_event].[dbo].[tbl_event_app_exclusive_episode]" & vbCrlf
    strSql = strSql + " SET isusing='N'" & vbCrlf
    strSql = strSql + " where evt_code=" & eCode
    strSql = strSql + " and idx IN (" & itemidarr & ")"
    dbget.execute strSql
elseif mode="prize" then
    PrizeUserid = split(prizearr,",")
	cnt = ubound(PrizeUserid)
    dbget.beginTrans
    for ix=0 to cnt
        '당첨 등록
        strSql = "INSERT INTO [db_event].[dbo].[tbl_event_app_exclusive_prize](evt_code,episode,prize_userid)" & vbCrlf
        strSql = strSql + " VALUES(" & eCode & "," & episode & ",'" & PrizeUserid(ix) & "')"
        dbget.execute strSql
        If err.number<>0 Then
            dbget.rollback
        End If
        '장바구니 상품 등록
        strSql = "INSERT INTO [db_my10x10].[dbo].[tbl_my_baguni] (userKey,isLoginUser,itemid,itemoption,itemea,regdate,chkOrder)" & vbCrlf
        strSql = strSql + "VALUES('" & PrizeUserid(ix) & "','Y'," & itemid & ",'0000',1,getdate(),'N')"
        dbget.execute strSql
        If err.number<>0 Then
            dbget.rollback
        End If
    Next
    '당첨 완료 플래그 수정
    strSql = "UPDATE [db_event].[dbo].[tbl_event_app_exclusive_episode]" & vbCrlf
    strSql = strSql + " SET prizeyn='Y'" & vbCrlf
    strSql = strSql + " where evt_code=" & eCode
    strSql = strSql + " and episode=" & episode
    dbget.execute strSql
    If err.number<>0 Then
        dbget.rollback
    else
        dbget.committrans
    End If
end if
	response.write "<script type='text/javascript'>"
    if mode="prize" then
        response.write "	alert('당첨 설정이 완료 되었습니다.');self.close();"
    else
        response.write "	location.replace('pop_app_event_item_regist.asp?evt_code="&eCode&"');"
    end if
	response.write "</script>"
	dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->