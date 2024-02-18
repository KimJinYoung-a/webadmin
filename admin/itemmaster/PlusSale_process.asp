<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 주문서처리
' History : 이상구 생성
'			2020.05.12 정태훈 수정
'			2020.05.20 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode, sqlStr
dim buy_benefit_idx, benefit_type, benefit_title, benefit_subtitle, benefit_start_dt, benefit_end_dt, benefit_end_dt_time, whole_target_yn, use_yn
dim channel_www_yn, channel_mob_yn, channel_app_yn, mob_info_contents, www_info_contents, show_rank, sellcash

dim benefit_group_no, group_type, group_name, sort_no, condition_amount, delivery_type, catecode, makerid, evtcode, evt_buy_condition

dim plus_sale_item_idx, itemid, plus_sale_price, plus_sale_pct, plus_sale_buyprice, sale_burden_type, limit_yn, limit_cnt, max_buy_cnt, badge_contents, notice, opt_cnt

dim info_contents_mobile, info_contents_www

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim reguserid
reguserid = session("ssBctId")

mode = requestCheckVar(request("mode"), 32)

buy_benefit_idx = requestCheckVar(request("buy_benefit_idx"), 10)
benefit_type = requestCheckVar(request("benefit_type"), 1)
benefit_title = requestCheckVar(request("benefit_title"), 100)
benefit_subtitle = requestCheckVar(request("benefit_subtitle"), 200)
benefit_start_dt = requestCheckVar(request("benefit_start_dt"), 10)
benefit_end_dt = requestCheckVar(request("benefit_end_dt"), 10)
benefit_end_dt_time = requestCheckVar(request("benefit_end_dt_time"), 8)
whole_target_yn = requestCheckVar(request("whole_target_yn"), 1)
use_yn = requestCheckVar(request("use_yn"), 1)
channel_www_yn = requestCheckVar(request("channel_www_yn"), 1)
channel_mob_yn = requestCheckVar(request("channel_mob_yn"), 1)
channel_app_yn = requestCheckVar(request("channel_app_yn"), 1)
mob_info_contents = requestCheckVar(request("mob_info_contents"), 400)
www_info_contents = requestCheckVar(request("www_info_contents"), 400)
show_rank = requestCheckVar(request("show_rank"), 10)

benefit_group_no = requestCheckVar(request("benefit_group_no"), 10)
group_type = requestCheckVar(request("group_type"), 1)
group_name = requestCheckVar(request("group_name"), 40)
sort_no = requestCheckVar(request("sort_no"), 10)
condition_amount = requestCheckVar(request("condition_amount"), 18)
delivery_type = requestCheckVar(request("delivery_type"), 1)
catecode = requestCheckVar(request("catecode"), 18)
makerid = requestCheckVar(request("makerid"), 32)
evtcode = requestCheckVar(request("evtcode"), 10)
evt_buy_condition = requestCheckVar(request("evt_buy_condition"), 1)

plus_sale_item_idx = requestCheckVar(request("plus_sale_item_idx"), 10)
itemid = requestCheckVar(getNumeric(trim(request("itemid"))), 10)
plus_sale_price = requestCheckVar(request("plus_sale_price"), 18)
plus_sale_pct = requestCheckVar(request("plus_sale_pct"), 10)
plus_sale_buyprice = requestCheckVar(request("plus_sale_buyprice"), 18)
sale_burden_type = requestCheckVar(request("sale_burden_type"), 1)
limit_yn = requestCheckVar(request("limit_yn"), 1)
limit_cnt = requestCheckVar(request("limit_cnt"), 10)
max_buy_cnt = requestCheckVar(request("max_buy_cnt"), 10)
badge_contents = requestCheckVar(request("badge_contents"), 40)
notice = requestCheckVar(request("notice"), 100)
opt_cnt = requestCheckVar(request("opt_cnt"), 10)

info_contents_mobile = requestCheckVar(request("info_contents_mobile"), 3200)
info_contents_www = requestCheckVar(request("info_contents_www"), 3200)
sellcash = requestCheckVar(request("sellcash"), 18)

'할인율 계산 추가 (2022.06.10 정태훈)
if sellcash="" then
    plus_sale_pct = 0
elseif sellcash=0 then
    plus_sale_pct = 0
else
    plus_sale_pct = CLng((sellcash-plus_sale_price)/sellcash*100)
end if
'한정여부 강제 N (2022.06.22 정태훈)
limit_yn = "N"
limit_cnt = 0

if (mode = "insmaster") then
    sqlStr = " insert into db_sitemaster.dbo.tbl_buy_benefit(benefit_type, benefit_title, benefit_subtitle, benefit_start_dt, benefit_end_dt, whole_target_yn, use_yn, channel_www_yn, channel_mob_yn, channel_app_yn, mob_info_contents, www_info_contents, show_rank, reg_dt, reg_admin_id) "
    sqlStr = sqlStr & " values('" & benefit_type & "', '" & benefit_title & "', '" & benefit_subtitle & "', '" & benefit_start_dt & "', '" & benefit_end_dt &" "& benefit_end_dt_time & "', '" & whole_target_yn & "', '" & use_yn & "', '" & channel_www_yn & "', '" & channel_mob_yn & "', '" & channel_app_yn & "', '" & mob_info_contents & "', '" & www_info_contents & "', '" & show_rank & "', getdate(), '" & reguserid & "') "
    dbget.Execute sqlStr

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "opener.location.reload(); opener.focus(); window.close(); "
    response.write "</script>"
elseif (mode = "modimaster") then
    sqlStr = " update db_sitemaster.dbo.tbl_buy_benefit "
    sqlStr = sqlStr & " set last_update_admin_id = '" & reguserid & "', last_update_dt = getdate() "
    sqlStr = sqlStr & " ,benefit_type = '" & benefit_type & "' "
    sqlStr = sqlStr & " ,benefit_title = '" & benefit_title & "' "
    sqlStr = sqlStr & " ,benefit_subtitle = '" & benefit_subtitle & "' "
    sqlStr = sqlStr & " ,benefit_start_dt = '" & benefit_start_dt & "' "
    sqlStr = sqlStr & " ,benefit_end_dt = '" & benefit_end_dt &" "& benefit_end_dt_time & "' "
    sqlStr = sqlStr & " ,whole_target_yn = '" & whole_target_yn & "' "
    sqlStr = sqlStr & " ,use_yn = '" & use_yn & "' "
    sqlStr = sqlStr & " ,channel_www_yn = '" & channel_www_yn & "' "
    sqlStr = sqlStr & " ,channel_mob_yn = '" & channel_mob_yn & "' "
    sqlStr = sqlStr & " ,channel_app_yn = '" & channel_app_yn & "' "
    ''sqlStr = sqlStr & " ,mob_info_contents = '" & mob_info_contents & "' "
    ''sqlStr = sqlStr & " ,www_info_contents = '" & www_info_contents & "' "
    sqlStr = sqlStr & " ,show_rank = '" & show_rank & "' "
    sqlStr = sqlStr & " where buy_benefit_idx = " & buy_benefit_idx
    dbget.Execute sqlStr

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "location.replace('"&refer&"');"
    response.write "</script>"
elseif (mode = "insgroup") then

    sqlStr = " insert into db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group(buy_benefit_idx, group_type, group_name, sort_no, use_yn, condition_amount, delivery_type, catecode, makerid, evtcode, evt_buy_condition) "
    sqlStr = sqlStr & " values('" & buy_benefit_idx & "', '" & group_type & "', '" & group_name & "', '" & sort_no & "', '" & use_yn & "', '" & condition_amount & "', '" & delivery_type & "', '" & catecode & "', '" & makerid & "', '" & evtcode & "', '" & evt_buy_condition & "') "
    dbget.Execute sqlStr

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "opener.location.reload(); opener.focus(); window.close(); "
    response.write "</script>"
elseif (mode = "modigroup") then
    sqlStr = " update db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group "
    sqlStr = sqlStr & " set group_type = '" & group_type & "' "
    sqlStr = sqlStr & " ,group_name = '" & group_name & "' "
    sqlStr = sqlStr & " ,sort_no = '" & sort_no & "' "
    sqlStr = sqlStr & " ,use_yn = '" & use_yn & "' "
    sqlStr = sqlStr & " ,condition_amount = '" & condition_amount & "' "
    sqlStr = sqlStr & " ,delivery_type = '" & delivery_type & "' "
    sqlStr = sqlStr & " ,catecode = '" & catecode & "' "
    sqlStr = sqlStr & " ,makerid = '" & makerid & "' "
    sqlStr = sqlStr & " ,evtcode = '" & evtcode & "' "
    sqlStr = sqlStr & " ,evt_buy_condition = '" & evt_buy_condition & "' "
    sqlStr = sqlStr & " where benefit_group_no = " & benefit_group_no
    dbget.Execute sqlStr

    sqlStr = " update db_sitemaster.dbo.tbl_buy_benefit "
    sqlStr = sqlStr & " set last_update_admin_id = '" & reguserid & "', last_update_dt = getdate() "
    sqlStr = sqlStr & " where buy_benefit_idx = " & buy_benefit_idx
    dbget.Execute sqlStr

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "location.replace('"&refer&"');"
    response.write "</script>"
elseif (mode = "insitem") then

    sqlStr = " select top 1 o.itemid "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_item].[dbo].[tbl_item_option] o "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and o.itemid = " & itemid
    sqlStr = sqlStr & " 	and o.itemoption <> '0000' "
    rsget.Open sqlStr, dbget, 1
    if Not rsget.Eof then
        opt_cnt = 1
    else
        opt_cnt = 0
    end if
    rsget.close

    sqlStr = " insert into db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group_item(benefit_group_no, itemid, plus_sale_price, plus_sale_pct, plus_sale_buyprice, sale_burden_type, limit_yn, limit_cnt, max_buy_cnt, badge_contents, notice, sort_no, sell_cnt, use_yn, opt_cnt) "
    sqlStr = sqlStr & " values('" & benefit_group_no & "', '" & itemid & "', '" & plus_sale_price & "', '" & plus_sale_pct & "', '" & plus_sale_buyprice & "', '" & sale_burden_type & "', '" & limit_yn & "', '" & limit_cnt & "', '" & max_buy_cnt & "', '" & badge_contents & "', '" & notice & "', '" & sort_no & "', '0', '" & use_yn & "', '" & opt_cnt & "') "
    ''response.write sqlStr : response.end
    dbget.Execute sqlStr

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "opener.location.reload(); opener.focus(); window.close(); "
    response.write "</script>"
elseif (mode = "modiitem") then
    sqlStr = " update db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group_item "
    sqlStr = sqlStr & " set itemid = '" & itemid & "' "
    sqlStr = sqlStr & " ,plus_sale_price = '" & plus_sale_price & "' "
    sqlStr = sqlStr & " ,plus_sale_pct = '" & plus_sale_pct & "' "
    sqlStr = sqlStr & " ,plus_sale_buyprice = '" & plus_sale_buyprice & "' "
    sqlStr = sqlStr & " ,sale_burden_type = '" & sale_burden_type & "' "
    sqlStr = sqlStr & " ,limit_yn = '" & limit_yn & "' "
    sqlStr = sqlStr & " ,limit_cnt = '" & limit_cnt & "' "
    sqlStr = sqlStr & " ,max_buy_cnt = '" & max_buy_cnt & "' "
    sqlStr = sqlStr & " ,badge_contents = '" & badge_contents & "' "
    sqlStr = sqlStr & " ,notice = '" & notice & "' "
    sqlStr = sqlStr & " ,sort_no = '" & sort_no & "' "
    sqlStr = sqlStr & " ,use_yn = '" & use_yn & "' "
    sqlStr = sqlStr & " where plus_sale_item_idx = " & plus_sale_item_idx
    dbget.Execute sqlStr

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "location.replace('"&refer&"');"
    response.write "</script>"

elseif (mode = "modiinfo") then

    sqlStr = " update db_sitemaster.dbo.tbl_buy_benefit "
    sqlStr = sqlStr & " set last_update_admin_id = '" & reguserid & "', last_update_dt = getdate() "
    sqlStr = sqlStr & " ,info_contents_mobile = '" & html2db(info_contents_mobile) & "' "
    sqlStr = sqlStr & " ,info_contents_www = '" & html2db(info_contents_www) & "' "
    sqlStr = sqlStr & " where buy_benefit_idx = " & buy_benefit_idx
    dbget.Execute sqlStr

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "location.replace('"&refer&"');"
    response.write "</script>"

elseif (mode = "updOrderCount") then

    sqlStr = " exec [db_datamart].[dbo].[usp_TEN_Buy_Benifit_Stat_Make] " & benefit_group_no
    db3_dbget.Execute sqlStr

    response.write "<script language='javascript'>"
    response.write "alert('저장 되었습니다.');"
    response.write "location.replace('"&refer&"');"
    response.write "</script>"
else
	response.write "<script language='javascript'>"
	response.write "alert('잘못된 접근입니다.');"
	response.write "</script>"
    response.write "잘못된 접근입니다."
    dbget.close()	:	response.End
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
