<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : appDedicatedItem_process.asp
' Discription : 앱전용 응모템 설정 프로세스
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
dim eCode, mode, strSql, idx, etcNotice, notice
dim bannerImg, itemid, startdate, enddate, itemidarr
dim title_color, itemlist_bg_color, button_color, prize_bg_color, sub_title
dim prize_circle_color, prize_circle_color2, bannerImg2, bannerImg3, deeplink
eCode = requestCheckVar(Request.Form("evt_code"),10)
mode = requestCheckVar(Request.Form("mode"),10)
idx = requestCheckVar(Request.Form("idx"),10)
bannerImg = requestCheckVar(Request.Form("bannerImg"),128)
bannerImg2 = requestCheckVar(Request.Form("bannerImg2"),128)
bannerImg3 = requestCheckVar(Request.Form("bannerImg3"),128)
etcNotice = Request.Form("etcNotice")
notice = Request.Form("notice")
title_color = Request.Form("title_color")
itemlist_bg_color = Request.Form("itemlist_bg_color")
button_color = Request.Form("button_color")
prize_bg_color = Request.Form("prize_bg_color")
sub_title = Request.Form("sub_title")
prize_circle_color = Request.Form("prize_circle_color")
prize_circle_color2 = Request.Form("prize_circle_color2")
deeplink = Request.Form("deeplink")

if eCode="" then
    response.write "<script type='text/javascript'>"
    response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
    response.write "</script>"
    response.End
end if

if mode="Add" then
    strSql = "INSERT INTO [db_event].[dbo].[tbl_event_app_exclusive](evt_code,main_image,main_image2,etc_notice,notice, title_color, itemlist_bg_color, button_color, prize_bg_color, sub_title, prize_circle_color, prize_circle_color2, mo_main_image, deeplink)" & vbCrlf
    strSql = strSql + " VALUES(" & eCode & ",'" & bannerImg & "','" & bannerImg2 & "','" & etcNotice & "','" & notice & "','" & title_color & "','" & itemlist_bg_color & "','" & button_color & "','" & prize_bg_color & "','" & sub_title & "','" & prize_circle_color & "','" & prize_circle_color2 & "','" & bannerImg3 & "','" & deeplink & "')"
    dbget.execute strSql
elseif mode="Modify" then
    strSql = "UPDATE [db_event].[dbo].[tbl_event_app_exclusive]" & vbCrlf
    strSql = strSql + " SET main_image='" & bannerImg & "'" & vbCrlf
    strSql = strSql + " ,main_image2='" & bannerImg2 & "'" & vbCrlf
    strSql = strSql + " ,mo_main_image='" & bannerImg3 & "'" & vbCrlf
    strSql = strSql + " ,etc_notice='" & etcNotice & "'" & vbCrlf
    strSql = strSql + " ,notice='" & notice & "'" & vbCrlf
    strSql = strSql + " ,title_color='" & title_color & "'" & vbCrlf
    strSql = strSql + " ,itemlist_bg_color='" & itemlist_bg_color & "'" & vbCrlf
    strSql = strSql + " ,button_color='" & button_color & "'" & vbCrlf
    strSql = strSql + " ,prize_bg_color='" & prize_bg_color & "'" & vbCrlf
    strSql = strSql + " ,sub_title='" & sub_title & "'" & vbCrlf
    strSql = strSql + " ,prize_circle_color='" & prize_circle_color & "'" & vbCrlf
    strSql = strSql + " ,prize_circle_color2='" & prize_circle_color2 & "'" & vbCrlf
    strSql = strSql + " ,deeplink='" & deeplink & "'" & vbCrlf
    strSql = strSql + " where evt_code=" & eCode
    strSql = strSql + " and idx=" & idx
    dbget.execute strSql
end if

	response.write "<script type='text/javascript'>"
	response.write "	location.replace('pop_app_event_setting.asp?evt_code="&eCode&"');"
	response.write "</script>"
	dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->