<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%

dim chk, itemid, itemoption, slcode, ipgodate, reqstr, sidx, eidx
dim stockoutarr, sdetailidx, edetailidx, makerid

dim sqlStr, i, j

''response.write "작업중"
''dbget.close : response.end

chk  = request("chk")
chk = chk + ","
chk = Split(chk, ",")

for i = 0 to UBound(chk)
	if (trim(chk(i)) <> "") then

		makerid     = request("makerid" + CStr(trim(chk(i))))
        itemid      = request("itemid" + CStr(trim(chk(i))))
		itemoption  = request("itemoption" + CStr(trim(chk(i))))
		slcode      = request("slcode" + CStr(trim(chk(i))))
		ipgodate    = request("ipgodate" + CStr(trim(chk(i))))
		reqstr      = request("reqstr" + CStr(trim(chk(i))))

		sidx      = request("sidx" + CStr(trim(chk(i))))
		eidx      = request("eidx" + CStr(trim(chk(i))))

        sdetailidx      = request("sdetailidx" + CStr(trim(chk(i))))
        edetailidx      = request("edetailidx" + CStr(trim(chk(i))))

		if (ipgodate="1900-01-01") then ipgodate=""
		'if (((slcode = "03") or (slcode = "05")) and (reqstr = "")) then
		'        reqstr = "전화요망"
		'end if

		sqlStr = " insert into [db_temp].[dbo].tbl_mibeasong_list(detailidx,orderserial,itemid,itemoption,itemname,itemoptionname,itemno,itemlackno,code,reqstr, ipgodate, reguserid) "
		sqlStr = sqlStr + " select d.idx,d.orderserial,d.itemid,d.itemoption,d.itemname,d.itemoptionname,d.itemno,d.itemno,'00','','1900-01-01', '" & session("ssBctId") & "' "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_master] m "
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d on m.orderserial = d.orderserial "
		sqlStr = sqlStr + " 	left join [db_temp].[dbo].tbl_mibeasong_list T "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		d.idx = T.detailidx "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
        sqlStr = sqlStr + " 	and m.ipkumdiv = '4' "
        sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
        sqlStr = sqlStr + " 	and d.isupchebeasong = 'N' "
        sqlStr = sqlStr + " 	and d.makerid = '" & makerid & "' "
		sqlStr = sqlStr + " 	and d.itemid = " & itemid
		sqlStr = sqlStr + " 	and d.itemoption = '" & itemoption & "' "
		sqlStr = sqlStr + " 	and d.currstate = '0' "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and T.idx is NULL "
        ''response.write sqlStr
        ''dbget.close : response.end
        dbget.Execute sqlStr

		sqlStr = " update T "
		sqlStr = sqlStr + " set T.code = '" + slcode + "' "
        if (ipgodate="") then
		    sqlStr = sqlStr + " , T.ipgodate=NULL" &VbCRLF
		else
		    sqlStr = sqlStr + " , T.ipgodate='" + ipgodate + "'" &VbCRLF
		end if
		sqlStr = sqlStr + " , T.reqstr = '" + reqstr + "' " &VbCRLF
		sqlStr = sqlStr + "	, T.modiuserid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + "	, T.modidate = getdate() "
		sqlStr = sqlStr + "	, T.reqreguserid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_master] m "
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d on m.orderserial = d.orderserial "
		sqlStr = sqlStr + " 	join [db_temp].[dbo].tbl_mibeasong_list T "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		d.idx = T.detailidx "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		sqlStr = sqlStr + " 	and m.ipkumdiv >= '4' "
		sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
		sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
		sqlStr = sqlStr + " 	and d.isupchebeasong = 'N' "
        sqlStr = sqlStr + " 	and d.makerid = '" & makerid & "' "
		sqlStr = sqlStr + " 	and d.itemid = " & itemid
		sqlStr = sqlStr + " 	and d.itemoption = '" & itemoption & "' "
		sqlStr = sqlStr + " 	and d.currstate >= '0' "
		sqlStr = sqlStr + " 	and d.currstate < '7' "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
        sqlStr = sqlStr + " 	and T.state = '0' "				''처리 안한것만.
        ''rw sqlStr
		dbget.Execute sqlStr


		if False and (slcode = "05") then
			'// 품절출고불가(CS 담당자 지정)
			stockoutarr = ""

			sqlStr = "select detailidx from [db_temp].[dbo].tbl_mibeasong_list "
			sqlStr = sqlStr + " where itemid=" + CStr(itemid) &VbCRLF
			sqlStr = sqlStr + " and itemoption='" + itemoption + "'" &VbCRLF
			sqlStr = sqlStr + " and state=0" &VbCRLF ''처리 안한것만.
            ''rw sqlStr
			rsget.Open sqlStr, dbget

	        if  not rsget.EOF  then
	            do until rsget.eof
	            	stockoutarr = stockoutarr + "," + CStr(rsget("detailidx"))
	                rsget.MoveNext
	            loop
	        end if
	        rsget.close

	        stockoutarr   = split(stockoutarr, ",")

		    if IsArray(stockoutarr) then
		        for j = 0 to Ubound(stockoutarr)
		        	if (Trim(stockoutarr(j))<>"") then
						sqlStr = " exec db_cs.[dbo].[sp_Ten_MichulgoStockout_SetChargeID] " & Trim(stockoutarr(j)) & " "
                        ''rw sqlStr
						dbget.Execute sqlStr
					end if
		        next
		    end if
		end if

	end if
next

dim refer
refer = request.ServerVariables("HTTP_REFERER")

%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
