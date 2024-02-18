<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")


dim sqlStr, found, i


dim mode
dim makerid, is_return_allow, tel_start, tel_end, is_saturday_work, vacation_startday, vacation_endday, brand_comment, last_modifyday



mode = requestCheckVar(request("mode"),32)

makerid = requestCheckVar(request("makerid"),32)
is_return_allow = requestCheckVar(request("is_return_allow"),1)
tel_start = requestCheckVar(request("tel_start"),3)
tel_end = requestCheckVar(request("tel_end"),3)
is_saturday_work = requestCheckVar(request("is_saturday_work"),1)
vacation_startday = requestCheckVar(request("vacation_startday"),10)
vacation_endday = requestCheckVar(request("vacation_endday"),10)
brand_comment = html2db(request("brand_comment"))
last_modifyday = Left(now, 10)

if brand_comment <> "" then
	if checkNotValidHTML(brand_comment) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if


if (mode = "insert") then
	'
	sqlStr = "insert into " & TABLE_CS_BRAND_MEMO & "(brandid, is_return_allow, vacation_startday, vacation_endday, tel_start, tel_end, is_saturday_work, brand_comment, last_modifyday) "
	sqlStr = sqlStr & " values('" & makerid & "', '" & is_return_allow & "', '" & vacation_startday & "', '" & vacation_endday & "', " & tel_start & ", " & tel_end & ", '" & is_saturday_work & "', '" & brand_comment & "', '" & last_modifyday & "') "
	rsget.Open sqlStr,dbget,1
	'response.write sqlStr
elseif (mode = "modify") then
	'
	sqlStr = "update " & TABLE_CS_BRAND_MEMO & " set last_modifyday = '" & last_modifyday & "'"

	sqlStr = sqlStr & " ,is_return_allow = '" & is_return_allow & "' "
	sqlStr = sqlStr & " ,vacation_startday = '" & vacation_startday & "' "
	sqlStr = sqlStr & " ,vacation_endday = '" & vacation_endday & "' "
	sqlStr = sqlStr & " ,tel_start = " & tel_start & " "
	sqlStr = sqlStr & " ,tel_end = " & tel_end & " "
	sqlStr = sqlStr & " ,is_saturday_work = '" & is_saturday_work & "' "
	sqlStr = sqlStr & " ,brand_comment = '" & brand_comment & "' "
	sqlStr = sqlStr + " where brandid = '" & makerid & "' "
	rsget.Open sqlStr,dbget,1
	'response.write sqlStr

else
	'
end if

%>
<script language="javascript">
alert('저장 되었습니다.');
<% if refer<>"" then %>
location.replace('<%= refer %>');
<% end if %>
</script>
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->