<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")


dim sqlStr, found, i


dim mode
dim makerid, is_return_allow, tel_start, tel_end, is_saturday_work, vacation_startday, vacation_endday, brand_comment, last_modifyday
dim lunch_start, lunch_end, vacation_div, ret_comment, customer_return_deny
dim mduserid, isSpecialBrand
'===========스페셜 브랜드 2019-07-17 최종원===========
dim isexposure, frequency, exposure_seq, always_exposure, startdate, enddate, regdate, brand_icon

isexposure = request("isexposure")
frequency = request("frequency")
exposure_seq = request("exposure_seq")
always_exposure = request("always_exposure")
startdate = request("startDate")
enddate = request("endDate")
brand_icon = request("brand_icon")
isSpecialBrand = request("isSpecialBrand")

if enddate <> "" then enddate = enddate & " 23:59:59"
'=======================================================

mode = requestCheckVar(request("mode"),32)

makerid = requestCheckVar(request("makerid"),32)
mduserid = html2db(requestCheckVar(request("mduserid"),32))
is_return_allow = requestCheckVar(request("is_return_allow"),1)
tel_start = requestCheckVar(request("tel_start"),3)
tel_end = requestCheckVar(request("tel_end"),3)
is_saturday_work = requestCheckVar(request("is_saturday_work"),1)
vacation_startday = requestCheckVar(request("vacation_startday"),10)
vacation_endday = requestCheckVar(request("vacation_endday"),10)
brand_comment = html2db(request("brand_comment"))
customer_return_deny = requestCheckVar(request("customer_return_deny"), 1)

lunch_start = requestCheckVar(request("lunch_start"),32)
lunch_end = requestCheckVar(request("lunch_end"),32)
vacation_div = requestCheckVar(request("vacation_div"),32)
ret_comment = html2db(requestCheckVar(request("ret_comment"),32))

last_modifyday = Left(now, 10)

if isSpecialBrand > 0 then
	sqlStr = ""
	sqlStr = " update db_brand.[dbo].[tbl_special_brand] "
	sqlStr = sqlStr & " set isexposure = '"& isexposure &"'"
	sqlStr = sqlStr & " , frequency = '"& frequency &"' "
	sqlStr = sqlStr & " , exposure_seq = '"& exposure_seq &"' "
	sqlStr = sqlStr & " , always_exposure = '"& always_exposure &"' "
	sqlStr = sqlStr & " , startdate = '"& startdate &"' "
	sqlStr = sqlStr & " , enddate = '"& enddate &"' "
	sqlStr = sqlStr & " , brand_icon = '"& brand_icon &"' "
	sqlStr = sqlStr + " where brandid = '" & makerid & "' "

	rsget.Open sqlStr,dbget,1
elseif isexposure = "1" then
	sqlStr = ""
	sqlStr = sqlStr & " INSERT INTO db_brand.[dbo].[tbl_special_brand] "
	sqlStr = sqlStr & "            ([brandid] "
	sqlStr = sqlStr & "            ,[isexposure] "
	sqlStr = sqlStr & "            ,[frequency] "
	sqlStr = sqlStr & "            ,[exposure_seq] "
	sqlStr = sqlStr & "            ,[always_exposure] "
	sqlStr = sqlStr & "            ,[startdate] "
	sqlStr = sqlStr & "            ,[enddate] "
	sqlStr = sqlStr & "            ,[regdate] "
	sqlStr = sqlStr & "            ,[brand_icon]) "
	sqlStr = sqlStr & "      VALUES "
	sqlStr = sqlStr & "            ('"& makerid &"' "
	sqlStr = sqlStr & "            , '"& isexposure &"'"
	sqlStr = sqlStr & "            ,'"& frequency &"' "
	sqlStr = sqlStr & "            ,'"& exposure_seq &"' "
	sqlStr = sqlStr & "            ,'"& always_exposure &"' "
	if startdate = "" then
		sqlStr = sqlStr & "            ,null "
	else
		sqlStr = sqlStr & "            ,'"& startdate &"' "
	end if
	if enddate = "" then
		sqlStr = sqlStr & "            ,null "
	else
		sqlStr = sqlStr & "            ,'"& enddate &"' "
	end if
	sqlStr = sqlStr & "            ,getdate() "
	sqlStr = sqlStr & "            ,'"& brand_icon &"' "
	sqlStr = sqlStr & " 		   ) "

	rsget.Open sqlStr,dbget,1
end if

if (mode = "insert") then
	'
	sqlStr = ""
	sqlStr = "insert into [db_cs].[dbo].tbl_cs_brand_memo(brandid, is_return_allow, return_comment, vacation_startday, vacation_endday, tel_start, tel_end, is_saturday_work, brand_comment, last_modifyday, lunch_start, lunch_end, vacation_div, customer_return_deny) "
	sqlStr = sqlStr & " values('" & makerid & "', '" & is_return_allow & "', '" + CStr(ret_comment) + "', '" & vacation_startday & "', '" & vacation_endday & "', " & tel_start & ", " & tel_end & ", '" & is_saturday_work & "', '" & brand_comment & "', '" & last_modifyday & "', '" + CStr(lunch_start) + "', '" + CStr(lunch_end) + "', '" + CStr(vacation_div) + "', '" + CStr(customer_return_deny) + "') "
	rsget.Open sqlStr,dbget,1
	'response.write sqlStr

elseif (mode = "modify") then
	'
	sqlStr = ""
	sqlStr = "update [db_cs].[dbo].tbl_cs_brand_memo set last_modifyday = '" & last_modifyday & "'"

	sqlStr = sqlStr & " ,is_return_allow = '" & is_return_allow & "' "
	sqlStr = sqlStr & " ,return_comment = '" & ret_comment & "' "
	sqlStr = sqlStr & " ,vacation_startday = '" & vacation_startday & "' "
	sqlStr = sqlStr & " ,vacation_endday = '" & vacation_endday & "' "
	sqlStr = sqlStr & " ,tel_start = " & tel_start & " "
	sqlStr = sqlStr & " ,tel_end = " & tel_end & " "
	sqlStr = sqlStr & " ,is_saturday_work = '" & is_saturday_work & "' "
	sqlStr = sqlStr & " ,brand_comment = '" & brand_comment & "' "
	sqlStr = sqlStr & " ,lunch_start = '" & lunch_start & "' "
	sqlStr = sqlStr & " ,lunch_end = '" & lunch_end & "' "
	sqlStr = sqlStr & " ,vacation_div = '" & vacation_div & "' "
    sqlStr = sqlStr & " ,customer_return_deny = '" & customer_return_deny & "' "
	sqlStr = sqlStr + " where brandid = '" & makerid & "' "
	rsget.Open sqlStr,dbget,1
	'response.write sqlStr
else
	'
end if

if (makerid <> "") and (mduserid <> "") then
    sqlStr = "update [db_user].[dbo].tbl_user_c" + VbCrlf
    sqlStr = sqlStr + " set mduserid='" + CStr(mduserid)  + "'" + VbCrlf
    sqlStr = sqlStr + " where userid='" + makerid + "' and IsNull(mduserid, '') <> '" + CStr(mduserid) + "' " + VbCrlf
    rsget.Open sqlStr, dbget, 1
end if

%>
<script language="javascript">
alert('저장 되었습니다.');
<% if refer<>"" then %>
location.replace('<%= refer %>');
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
