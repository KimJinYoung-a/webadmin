<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode
dim idx, title, contents, startDate, endDate, importantLevel, openLevel, useYN, reguserid, modiuserid
dim department_id_Arr, empno_Arr, department_id, empno
dim sqlStr, i, j, k


mode = requestcheckvar(request("mode"),32)

idx = requestcheckvar(request("idx"),32)
title = html2db(requestcheckvar(request("title"),64))
contents = html2db(requestcheckvar(request("contents"),2000))
startDate = requestcheckvar(request("startDate"),10)
endDate = requestcheckvar(request("endDate"),10)
importantLevel = requestcheckvar(request("importantLevel"),10)
openLevel = requestcheckvar(request("openLevel"),10)
useYN = requestcheckvar(request("useYN"),1)

department_id_Arr = request("department_id")
empno_Arr = request("empno")

reguserid = session("ssBctId")
modiuserid = session("ssBctId")


''dbget.close()
''rw UBound(department_id_Arr)
''rw UBound(empno_Arr)
''Response.end

select case mode
	case "ins"
		department_id_Arr = Split(department_id_Arr, ",")
		empno_Arr = Split(empno_Arr, ",")

		sqlStr = " insert into [db_partner].[dbo].[tbl_compCalendar](title, contents, startDate, endDate, importantLevel, openLevel, useYN, reguserid, modiuserid, regdate, lastupdate) "
		sqlStr = sqlStr + " values('" & title & "', '" & contents & "', '" & startDate & "', '" & endDate & "', '" & importantLevel & "', '" & openLevel & "', '" & useYN & "', '" & reguserid & "', '" & modiuserid & "', getdate(), getdate()) "
		dbget.execute sqlStr

		idx = -1
		sqlStr = "Select @@Identity as idx"
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsget.EOF OR rsget.BOF) THEN
			idx = rsget("idx")
		END IF
		rsget.Close

		if (UBound(department_id_Arr) >= 0) then
			for i = 0 to UBound(department_id_Arr)
				department_id = department_id_Arr(i)
				if (i = 0) then
					sqlStr = " insert into [db_partner].[dbo].[tbl_compCalendar_OpenList](calIdx, department_id, empno, useYN, regdate, lastupdate) "
					sqlStr = sqlStr + " values('" & idx & "', '" & department_id & "', NULL, 'Y', getdate(), getdate())"
				else
					sqlStr = sqlStr + " ,('" & idx & "', '" & department_id & "', NULL, 'Y', getdate(), getdate())"
				end if
			next
			dbget.execute sqlStr
		end if

		if (UBound(empno_Arr) >= 0) then
			for i = 0 to UBound(empno_Arr)
				empno = empno_Arr(i)
				if (i = 0) then
					sqlStr = " insert into [db_partner].[dbo].[tbl_compCalendar_OpenList](calIdx, department_id, empno, useYN, regdate, lastupdate) "
					sqlStr = sqlStr + " values('" & idx & "', NULL, '" & empno & "', 'Y', getdate(), getdate())"
				else
					sqlStr = sqlStr + " ,('" & idx & "', NULL, '" & empno & "', 'Y', getdate(), getdate())"
				end if
			next
			dbget.execute sqlStr
		end if

		Response.Write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
	case "mod"
		''sqlStr
		sqlStr = " update [db_partner].[dbo].[tbl_compCalendar] "
		sqlStr = sqlStr + " set lastupdate = getdate() "
		sqlStr = sqlStr + " , title = '" & title & "' "
		sqlStr = sqlStr + " , contents = '" & contents & "' "
		sqlStr = sqlStr + " , startDate = '" & startDate & "' "
		sqlStr = sqlStr + " , endDate = '" & endDate & "' "
		sqlStr = sqlStr + " , importantLevel = '" & importantLevel & "' "
		sqlStr = sqlStr + " , openLevel = '" & openLevel & "' "
		sqlStr = sqlStr + " , useYN = '" & useYN & "' "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	idx = " & idx & " "
		dbget.execute sqlStr

		''있다(Y) 없다	update
		''있다(N) 있다	update
		''없다 있다		insert
		if (department_id_Arr = "") then
			sqlStr = " update [db_partner].[dbo].[tbl_compCalendar_OpenList] "
			sqlStr = sqlStr + " set lastupdate = getdate() "
			sqlStr = sqlStr + " , useYN = 'N' "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and department_id is not NULL "
			sqlStr = sqlStr + " 	and useYN = 'Y' "
			sqlStr = sqlStr + " 	and calIdx = " & idx & " "
			dbget.execute sqlStr
		else
			department_id_Arr = Replace(department_id_Arr, " ", "")
			department_id_Arr = Replace(department_id_Arr, ",", "','")
			department_id_Arr = "'" & department_id_Arr & "'"

			sqlStr = " update [db_partner].[dbo].[tbl_compCalendar_OpenList] "
			sqlStr = sqlStr + " set lastupdate = getdate() "
			sqlStr = sqlStr + " , useYN = 'N' "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and department_id not in (" & department_id_Arr & ") "
			sqlStr = sqlStr + " 	and useYN = 'Y' "
			sqlStr = sqlStr + " 	and calIdx = " & idx & " "
			dbget.execute sqlStr
			''rw sqlStr

			sqlStr = " update [db_partner].[dbo].[tbl_compCalendar_OpenList] "
			sqlStr = sqlStr + " set lastupdate = getdate() "
			sqlStr = sqlStr + " , useYN = 'Y' "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and department_id in (" & department_id_Arr & ") "
			sqlStr = sqlStr + " 	and useYN = 'N' "
			sqlStr = sqlStr + " 	and calIdx = " & idx & " "
			dbget.execute sqlStr
			''rw sqlStr

			sqlStr = " insert into [db_partner].[dbo].[tbl_compCalendar_OpenList](calIdx, department_id, empno, useYN, regdate, lastupdate) "
			sqlStr = sqlStr + " select " & idx & ", p.cid, NULL, 'Y', getdate(), getdate() "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	[db_partner].[dbo].tbl_user_department p "
			sqlStr = sqlStr + " 	left join [db_partner].[dbo].[tbl_compCalendar_OpenList] o "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and o.calIdx = " & idx & " "
			sqlStr = sqlStr + " 		and o.department_id = p.cid "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and p.cid in (" & department_id_Arr & ") "
			sqlStr = sqlStr + " 	and o.idx is NULL "
			dbget.execute sqlStr
			''rw sqlStr
		end if

		if (empno_Arr = "") then
			sqlStr = " update [db_partner].[dbo].[tbl_compCalendar_OpenList] "
			sqlStr = sqlStr + " set lastupdate = getdate() "
			sqlStr = sqlStr + " , useYN = 'N' "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and empno is not NULL "
			sqlStr = sqlStr + " 	and useYN = 'Y' "
			sqlStr = sqlStr + " 	and calIdx = " & idx & " "
			dbget.execute sqlStr
		else
			empno_Arr = Replace(empno_Arr, " ", "")
			empno_Arr = Replace(empno_Arr, ",", "','")
			empno_Arr = "'" & empno_Arr & "'"

			sqlStr = " update [db_partner].[dbo].[tbl_compCalendar_OpenList] "
			sqlStr = sqlStr + " set lastupdate = getdate() "
			sqlStr = sqlStr + " , useYN = 'N' "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and empno not in (" & empno_Arr & ") "
			sqlStr = sqlStr + " 	and useYN = 'Y' "
			sqlStr = sqlStr + " 	and calIdx = " & idx & " "
			dbget.execute sqlStr
			''rw sqlStr

			sqlStr = " update [db_partner].[dbo].[tbl_compCalendar_OpenList] "
			sqlStr = sqlStr + " set lastupdate = getdate() "
			sqlStr = sqlStr + " , useYN = 'Y' "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and empno in (" & empno_Arr & ") "
			sqlStr = sqlStr + " 	and useYN = 'N' "
			sqlStr = sqlStr + " 	and calIdx = " & idx & " "
			dbget.execute sqlStr
			''rw sqlStr

			sqlStr = " insert into [db_partner].[dbo].[tbl_compCalendar_OpenList](calIdx, department_id, empno, useYN, regdate, lastupdate) "
			sqlStr = sqlStr + " select " & idx & ", NULL, p.empno, 'Y', getdate(), getdate() "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	[db_partner].[dbo].tbl_user_tenbyten p "
			sqlStr = sqlStr + " 	left join [db_partner].[dbo].[tbl_compCalendar_OpenList] o "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and o.calIdx = " & idx & " "
			sqlStr = sqlStr + " 		and o.empno = p.empno "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and p.empno in (" & empno_Arr & ") "
			sqlStr = sqlStr + " 	and o.idx is NULL "
			dbget.execute sqlStr
			''rw sqlStr
		end if

		Response.Write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
	case else
		''
end select

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->

<script type='text/javascript'>
	opener.location.reload();
	self.close();
</script>
