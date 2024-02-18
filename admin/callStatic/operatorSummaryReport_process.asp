<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<script language="javascript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%

dim mode, yyyymmdd, pTenUserID
dim oJSON, datas, data, user
Dim sqlStr, AssignedCNT, addCount
dim extension, tenUserID, in_count, in_count_success, in_count_total_sec, out_count, out_count_success, out_count_total_sec
dim aj_XmlHttp, result
dim i

mode = req("mode", "")
yyyymmdd = req("yyyymmdd", "")
pTenUserID = req("tenUserID", "")


Select Case mode
	Case "recvdata"
		if (yyyymmdd = "") then
			''전날
			yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
		end if

		'// 데이타 수신
		Set aj_XmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		aj_XmlHttp.open "GET", "http://110.93.128.96/ipcc/smart/remote/tenbyten_stat_user.jsp?fromtime=" + CStr(yyyymmdd) + " 00:00:00&totime=" + CStr(yyyymmdd) + " 23:59:59", False
		aj_XmlHttp.setRequestHeader "Content-Type", "text/json"
		aj_XmlHttp.setRequestHeader "CharSet", "UTF-8"
		aj_XmlHttp.Send
		result = aj_XmlHttp.responseText
		set aj_XmlHttp = Nothing

		'// 파싱
		Set oJSON = JSON.parse(result)

		if (oJSON.RESULT = "OK") then

			sqlStr = " delete from [db_datamart].[dbo].[tbl_call_cdr_NEW] "
			sqlStr = sqlStr & " where yyyymmdd = '" + CStr(yyyymmdd) + "' "
			if (pTenUserID <> "") then
				sqlStr = sqlStr + " and tenUserID = '" & pTenUserID & "' "
			end if
			db3_dbget.Execute sqlStr,AssignedCNT

			addCount = 0

			sqlStr = " insert into [db_datamart].[dbo].[tbl_call_cdr_NEW] (yyyymmdd, extension, tenUserID, in_count, in_count_success, in_count_total_sec, out_count, out_count_success, out_count_total_sec) "
			''sqlStr = sqlStr & " values "
			for i = 0 to oJSON.DATAS.length - 1
				Set user = oJSON.DATAS.get(i)

				extension = ""
				tenUserID = user.userid
				in_count = user.in_count
				in_count_success = user.in_con_count
				in_count_total_sec = user.in_con_count * user.in_ava
				out_count = user.out_count
				out_count_success = user.out_con_count
				out_count_total_sec = user.out_con_count * user.out_ava

				if (tenUserID <> "SUM") and (pTenUserID = "" or pTenUserID = tenUserID) then
					if (addCount > 0) then
						sqlStr = sqlStr & " union all "
					end if

					sqlStr = sqlStr & " select '" + CStr(yyyymmdd) + "', '" + CStr(extension) + "', '" + CStr(tenUserID) + "', '" + CStr(in_count) + "', '" + CStr(in_count_success) + "', '" + CStr(in_count_total_sec) + "', '" + CStr(out_count) + "', '" + CStr(out_count_success) + "', '" + CStr(out_count_total_sec) + "' "
					addCount = addCount + 1
				end if

				Set user = Nothing
			next

			if (addCount > 0) then
				db3_dbget.Execute sqlStr,AssignedCNT
			else
				'// 콜서버에 없는 아이디 추가
				sqlStr = " insert into [db_datamart].[dbo].[tbl_call_cdr_NEW](yyyymmdd, extension, tenUserID, in_count, in_count_success, in_count_total_sec, out_count, out_count_success, out_count_total_sec) "
				sqlStr = sqlStr & " select T.yyyymmdd, '', T.tenUserID, 0, 0, 0, 0, 0, 0 "
				sqlStr = sqlStr & " from "
				sqlStr = sqlStr & " 	( "
				sqlStr = sqlStr & " 		select distinct '" + CStr(yyyymmdd) + "' as yyyymmdd, tenUserID "
				sqlStr = sqlStr & " 		from "
				sqlStr = sqlStr & " 		[db_datamart].[dbo].[tbl_call_cdr_NEW] "
				sqlStr = sqlStr & " 		where yyyymmdd >= '" + CStr(yyyymmdd) + "' and yyyymmdd < DateAdd(m, 1, '" + CStr(yyyymmdd) + "') and tenUserID = '" & pTenUserID & "' "
				sqlStr = sqlStr & " 	) T "
				sqlStr = sqlStr & " 	left join [db_datamart].[dbo].[tbl_call_cdr_NEW] c "
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		1 = 1 "
				sqlStr = sqlStr & " 		and c.yyyymmdd = T.yyyymmdd "
				sqlStr = sqlStr & " 		and c.tenUserID = T.tenUserID "
				sqlStr = sqlStr & " where "
				sqlStr = sqlStr & " 	1 = 1 "
				sqlStr = sqlStr & " 	and c.yyyymmdd is NULL "
				db3_dbget.Execute sqlStr
				AssignedCNT = 1
			end if

			IF (AssignedCNT>0) then
				sqlStr = " update c "
				sqlStr = sqlStr & " set c.extension = i.localcallno "
				sqlStr = sqlStr & " from "
				sqlStr = sqlStr & " 	[db_datamart].[dbo].[tbl_call_cdr_NEW] c "
				sqlStr = sqlStr & " 	join [TENDB].[db_cs].[dbo].[tbl_cs_ippbx_user] i "   ''2018/03/09 [TENDB] 추가.
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		c.tenUserID = i.userid "
				sqlStr = sqlStr & " where "
				sqlStr = sqlStr & " 	c.yyyymmdd = '" + CStr(yyyymmdd) + "' "
				db3_dbget.Execute sqlStr,AssignedCNT

				sqlStr = " update c "
				sqlStr = sqlStr & " set c.extension = T.extension "
				sqlStr = sqlStr & " from "
				sqlStr = sqlStr & " 	[db_datamart].[dbo].[tbl_call_cdr_NEW] c "
				sqlStr = sqlStr & " 	join ( "
				sqlStr = sqlStr & " 		select tenUserID, max(extension) as extension "
				sqlStr = sqlStr & " 		from "
				sqlStr = sqlStr & " 		[db_datamart].[dbo].[tbl_call_cdr_NEW] "
				sqlStr = sqlStr & " 		where "
				sqlStr = sqlStr & " 			1 = 1 "
				sqlStr = sqlStr & " 			and yyyymmdd >= '" + Left(CStr(yyyymmdd), 7) + "-01' "
				sqlStr = sqlStr & " 			and yyyymmdd < convert(varchar(10), DateAdd(month, 1, '" + Left(CStr(yyyymmdd), 7) + "-01'), 121) "
				sqlStr = sqlStr & " 			and extension <> '' "
				sqlStr = sqlStr & " 		group by tenUserID "
				sqlStr = sqlStr & " 	) T "
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		1 = 1 "
				sqlStr = sqlStr & " 		and yyyymmdd >= '" + Left(CStr(yyyymmdd), 7) + "-01' "
				sqlStr = sqlStr & " 		and yyyymmdd < convert(varchar(10), DateAdd(month, 1, '" + Left(CStr(yyyymmdd), 7) + "-01'), 121) "
				sqlStr = sqlStr & " 		and T.tenUserID = c.tenUserID "
				sqlStr = sqlStr & " 		and c.extension = '' "
				db3_dbget.Execute sqlStr

        	    if (IsAutoScript) then
        	        rw "OK|" & yyyymmdd
        	    ELSE
					Response.Write "<script language=javascript>alert('OK : " + CStr(AssignedCNT) + " 건'); opener.location.reload(); opener.focus(); window.close();</script>"
            	ENd IF
            ENd IF
		else
			if (IsAutoScript) then
				rw "[" + CStr(yyyymmdd) + "] 콜서버 데이타 수신 실패"
			else
				Response.Write "<script language=javascript>alert('콜서버 데이타 수신 실패\n나중에 다시 시도해보세요');</script>"
			end if
		end if

		Set oJSON = Nothing
		''
	Case Else
		''에러
End Select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
