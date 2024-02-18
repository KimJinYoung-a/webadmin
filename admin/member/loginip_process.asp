<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 로그인 IP 관리
' Hieditor : 이상구 생성
'			 2020.07.17 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode, idx, ipaddressArr, duplicate, sqlStr, i
dim ipaddress, department_id, userid, managername, comment, usescmyn, uselogicsyn, usecustomerinfoyn, reguserid, modiuserid, useyn, regdate, lastupdate
mode		= requestCheckvar(Request("mode"),10)
idx			= requestCheckvar(Request("idx"),10)
ipaddress	= requestCheckvar(Request("ipaddress"),3200)
department_id	= requestCheckvar(Request("department_id"),32)
userid			= html2db(requestCheckvar(Request("userid"),32))
managername		= html2db(requestCheckvar(Request("managername"),32))
comment			= html2db(requestCheckvar(Request("comment"),120))
usescmyn		= requestCheckvar(Request("usescmyn"),32)
uselogicsyn		= requestCheckvar(Request("uselogicsyn"),32)
usecustomerinfoyn		= requestCheckvar(Request("usecustomerinfoyn"),32)
reguserid		= session("ssBctId")
modiuserid		= session("ssBctId")
useyn			= requestCheckvar(Request("useyn"),32)

select case mode
	case "ins"
		sqlStr = "select top 1 ipaddress "
		sqlStr = sqlStr + " from db_partner.dbo.tbl_user_loginIP with (nolock)"
		sqlStr = sqlStr + " where ipaddress in ('" & Replace(ipaddress, ",", "','") & "') "

		duplicate = ""
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
        	duplicate = rsget("ipaddress")
		end if
		rsget.Close

	    if (duplicate <> "") then
	        response.write "<script>alert('이미 등록된 아이피가 있습니다.(" & duplicate & ")');history.back();</script>"
	        dbget.close()	:	response.End
	    end if

		ipaddressArr = Split(ipaddress, ",")

		for i = 0 to UBound(ipaddressArr)
			if (Trim(ipaddressArr(i)) <> "") then
				ipaddress = Trim(ipaddressArr(i))
				sqlStr = "insert into db_partner.dbo.tbl_user_loginIP(ipaddress, department_id, userid, managername, comment, usescmyn, uselogicsyn, usecustomerinfoyn, reguserid, modiuserid, useyn, regdate, lastupdate)"
				sqlStr = sqlStr + " values('" & ipaddress & "', '" & department_id & "', '" & userid & "', '" & managername & "', '" & comment & "', '" & usescmyn & "', '" & uselogicsyn & "', '" & usecustomerinfoyn & "', '" & reguserid & "', '" & modiuserid & "', '" & useyn & "', getdate(), getdate())"
				''response.write sqlStr
				dbget.Execute sqlStr
			end if
		next

		response.write "<script>alert('저장 되었습니다.');</script>"
		response.write "<script>opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()	:	response.End

	case "modi"
		sqlStr = "select top 1 ipaddress "
		sqlStr = sqlStr + " from db_partner.dbo.tbl_user_loginIP with (nolock)"
		sqlStr = sqlStr + " where ipaddress in ('" & Replace(ipaddress, ",", "','") & "') and idx<>"& idx &""

		duplicate = ""
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
        	duplicate = rsget("ipaddress")
		end if
		rsget.Close

	    if (duplicate <> "") then
	        response.write "<script>alert('이미 등록된 아이피가 있습니다.(" & duplicate & ")');history.back();</script>"
	        dbget.close()	:	response.End
	    end if

		sqlStr = " update db_partner.dbo.tbl_user_loginIP "
		sqlStr = sqlStr + " set lastupdate = getdate()"
		sqlStr = sqlStr + " , ipaddress = '" & ipaddress & "' "
		sqlStr = sqlStr + " , department_id = '" & department_id & "' "
		sqlStr = sqlStr + " , userid = '" & userid & "' "
		sqlStr = sqlStr + " , managername = '" & managername & "' "
		sqlStr = sqlStr + " , comment = '" & comment & "' "
		sqlStr = sqlStr + " , usescmyn = '" & usescmyn & "' "
		sqlStr = sqlStr + " , uselogicsyn = '" & uselogicsyn & "' "
		sqlStr = sqlStr + " , usecustomerinfoyn = '" & usecustomerinfoyn & "' "
		sqlStr = sqlStr + " , modiuserid = '" & modiuserid & "' "
		sqlStr = sqlStr + " , useyn = '" & useyn & "' "
		sqlStr = sqlStr + " where idx = " & idx
		''response.write sqlStr
		dbget.Execute sqlStr

		response.write "<script type='text/javascript'>"
		response.write "	alert('저장 되었습니다.');"
		response.write "	opener.location.reload();"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	case else
		''
end select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
