<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : attendance_process.asp
' Discription : 출석 체크 이벤트 설정 프로세스
' History : 2023.08.01 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
dim eCode, mode, strSql
dim idx, main_image, main_image_link, button_before_day_color, deeplink, mileage_summary
dim button_before_point_color, button_before_bg_color, button_after_day_color, button_after_point_color
dim button_after_bg_color, check_area_bg_color, check_title_color, check_button_bg_color
dim check_button_title_color, check_etc_contents, check_etc_contents_color, alarm_bg_color, alarm_etc_contents
dim mo_main_image, mo_main_image2, button_today_ring_color, popup_bubble_bg_color, popup_bubble_text_color

eCode = requestCheckVar(Request.Form("evt_code"),10)
mode = requestCheckVar(Request.Form("mode"),10)
idx = requestCheckVar(Request.Form("idx"),10)
main_image = requestCheckVar(Request.Form("main_image"),128)
main_image_link = requestCheckVar(Request.Form("main_image_link"),128)
mo_main_image = requestCheckVar(Request.Form("mo_main_image"),128)
mo_main_image2 = requestCheckVar(Request.Form("mo_main_image2"),128)
button_before_day_color = Request.Form("button_before_day_color")
button_before_point_color = Request.Form("button_before_point_color")
button_before_bg_color = Request.Form("button_before_bg_color")
button_after_day_color = Request.Form("button_after_day_color")
button_after_point_color = Request.Form("button_after_point_color")
button_after_bg_color = Request.Form("button_after_bg_color")
button_today_ring_color = Request.Form("button_today_ring_color")
check_area_bg_color = Request.Form("check_area_bg_color")
check_title_color = Request.Form("check_title_color")
check_button_bg_color = Request.Form("check_button_bg_color")
check_button_title_color = Request.Form("check_button_title_color")
check_etc_contents = Request.Form("check_etc_contents")
check_etc_contents_color = Request.Form("check_etc_contents_color")
alarm_bg_color = Request.Form("alarm_bg_color")
alarm_etc_contents = Request.Form("alarm_etc_contents")
deeplink = Request.Form("deeplink")
popup_bubble_bg_color = Request.Form("popup_bubble_bg_color")
popup_bubble_text_color = Request.Form("popup_bubble_text_color")
mileage_summary = Request.Form("mileage_summary")

if eCode="" then
    response.write "<script type='text/javascript'>"
    response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
    response.write "</script>"
    response.End
end if

if mode="Add" then
    strSql = "INSERT INTO [db_event].[dbo].[tbl_event_attendance](evt_code,main_image,main_image_link,mo_main_image,mo_main_image2,button_before_day_color,button_before_point_color,button_before_bg_color,button_after_day_color,button_after_point_color,button_after_bg_color,button_today_ring_color,check_area_bg_color,check_title_color,check_button_bg_color,check_button_title_color,check_etc_contents,check_etc_contents_color,alarm_bg_color,alarm_etc_contents,deeplink,popup_bubble_bg_color,popup_bubble_text_color,mileage_summary)" & vbCrlf
    strSql = strSql + " VALUES(" & eCode & ",'" & main_image & "','" & main_image_link & "','" & mo_main_image & "','" & mo_main_image2 & "','" & button_before_day_color & "','" & button_before_point_color & "','" & button_before_bg_color & "','" & button_after_day_color & "','" & button_after_point_color & "','" & button_after_bg_color & "','" & button_today_ring_color & "','" & check_area_bg_color & "','" & check_title_color & "','" & check_button_bg_color & "','" & check_button_title_color & "','" & check_etc_contents & "','" & check_etc_contents_color & "','" & alarm_bg_color & "','" & alarm_etc_contents & "','" & deeplink & "','" & popup_bubble_bg_color & "','" & popup_bubble_text_color & "','" & mileage_summary & "')"
    dbget.execute strSql
elseif mode="Modify" then
    strSql = "UPDATE [db_event].[dbo].[tbl_event_attendance]" & vbCrlf
    strSql = strSql + " SET main_image='" & main_image & "'" & vbCrlf
    strSql = strSql + " ,main_image_link='" & main_image_link & "'" & vbCrlf
    strSql = strSql + " ,mo_main_image='" & mo_main_image & "'" & vbCrlf
    strSql = strSql + " ,mo_main_image2='" & mo_main_image2 & "'" & vbCrlf
    strSql = strSql + " ,button_before_day_color='" & button_before_day_color & "'" & vbCrlf
    strSql = strSql + " ,button_before_point_color='" & button_before_point_color & "'" & vbCrlf
    strSql = strSql + " ,button_before_bg_color='" & button_before_bg_color & "'" & vbCrlf
    strSql = strSql + " ,button_after_day_color='" & button_after_day_color & "'" & vbCrlf
    strSql = strSql + " ,button_after_point_color='" & button_after_point_color & "'" & vbCrlf
    strSql = strSql + " ,button_after_bg_color='" & button_after_bg_color & "'" & vbCrlf
    strSql = strSql + " ,button_today_ring_color='" & button_today_ring_color & "'" & vbCrlf
    strSql = strSql + " ,check_area_bg_color='" & check_area_bg_color & "'" & vbCrlf
    strSql = strSql + " ,check_title_color='" & check_title_color & "'" & vbCrlf
    strSql = strSql + " ,check_button_bg_color='" & check_button_bg_color & "'" & vbCrlf
    strSql = strSql + " ,check_button_title_color='" & check_button_title_color & "'" & vbCrlf
    strSql = strSql + " ,check_etc_contents='" & check_etc_contents & "'" & vbCrlf
    strSql = strSql + " ,check_etc_contents_color='" & check_etc_contents_color & "'" & vbCrlf
    strSql = strSql + " ,alarm_bg_color='" & alarm_bg_color & "'" & vbCrlf
    strSql = strSql + " ,alarm_etc_contents='" & alarm_etc_contents & "'" & vbCrlf
    strSql = strSql + " ,deeplink='" & deeplink & "'" & vbCrlf
    strSql = strSql + " ,popup_bubble_bg_color='" & popup_bubble_bg_color & "'" & vbCrlf
    strSql = strSql + " ,popup_bubble_text_color='" & popup_bubble_text_color & "'" & vbCrlf
    strSql = strSql + " ,mileage_summary='" & mileage_summary & "'" & vbCrlf
    strSql = strSql + " where evt_code=" & eCode
    strSql = strSql + " and idx=" & idx
    dbget.execute strSql
end if

	response.write "<script type='text/javascript'>"
	response.write "	location.replace('pop_attendance_event_setting.asp?evt_code="&eCode&"');"
	response.write "</script>"
	dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->