<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim mode, designerid, mastercode, detailidx

mode = request("mode")
designerid = request("designerid")
mastercode = request("mastercode")
detailidx = request("detailidx")


if (mode = "") then
	mode = "chkjungsanexist"
end if

dim sqlStr


if mode = "chkjungsanexist" then
	'// 정산내역 체크
	sqlStr = " select top 1 d.mastercode "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_jungsan.dbo.tbl_designer_jungsan_master m "
	sqlStr = sqlStr + " 	join db_jungsan.dbo.tbl_designer_jungsan_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.id = d.masteridx "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and m.yyyymm >= '" + Left(DateAdd("m", -2, Now()), 7) + "' "
	sqlStr = sqlStr + " and d.mastercode = '" + CStr(mastercode) + "' "

	if (designerid <> "") then
		sqlStr = sqlStr + " and m.designerid = '" + CStr(designerid) + "' "
	end if

	if (detailidx <> "") then
		sqlStr = sqlStr + " and d.detailidx = '" + CStr(detailidx) + "' "
	end if

	rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			response.write "Y"
		else
			response.write "N"
		end if
	rsget.close

else
	''
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
