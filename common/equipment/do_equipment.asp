<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 장비자산관리
' History : 2008년 06월 27일 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<%
dim idx ,equip_code ,equip_gubun ,equip_name ,equip_spec ,equip_mainimage ,property_gubun
dim manufacture_sn ,manufacture_company ,manufacture_manager ,manufacture_tel ,buy_company_name
dim buy_date ,buy_cost ,buy_vat ,buy_sum ,using_userid ,using_date ,out_date ,state ,etc ,part_sn
dim reguserid ,lastuserid ,isusing, info_gubun_idx, info_gubun_val, info_gubun_idx_list, info_GbnIdx_exist_list
dim mode ,sqlStr, i, j, adminuserid, accountassetcode, paymentrequestidx
dim accountGubun, department_id, yyyymm, tmpDate, found
dim BIZSECTION_CD, locate_gubun, monthlyDeprice, remainValue201412, info_gubun, info_importance_C, info_importance_I, info_importance_A
dim tmp_idx_arr, tmp_val_arr, info_gubun_idx_arr, info_gubun_val_arr, arridx
	accountassetcode = requestCheckVar(request("accountassetcode"),32)
	paymentrequestidx = requestCheckVar(request("paymentrequestidx"),10)
	idx = requestCheckVar(request("idx"),10)
	equip_code = requestCheckVar(request("equip_code"),20)
	equip_gubun = requestCheckVar(request("equip_gubun"),2)
	equip_name = requestCheckVar(request("equip_name"),64)
	equip_spec = requestCheckVar(request("equip_spec"),800)
	equip_mainimage = requestCheckVar(request("equip_mainimage"),128)
	property_gubun = requestCheckVar(request("property_gubun"),10)
	manufacture_sn = requestCheckVar(request("manufacture_sn"),32)
	manufacture_company = requestCheckVar(request("manufacture_company"),64)
	manufacture_manager = requestCheckVar(request("manufacture_manager"),32)
	manufacture_tel = requestCheckVar(request("manufacture_tel"),16)
	buy_company_name = requestCheckVar(request("buy_company_name"),64)
	buy_date = requestCheckVar(request("buy_date"),32)
	buy_cost = requestCheckVar(request("buy_cost"),20)
	buy_vat = requestCheckVar(request("buy_vat"),20)
	buy_sum = requestCheckVar(request("buy_sum"),20)
	using_userid = requestCheckVar(request("using_userid"),32)
	using_date = requestCheckVar(request("using_date"),20)
	out_date = requestCheckVar(request("out_date"),20)
	state = requestCheckVar(request("state"),10)
	etc = requestCheckVar(request("etc"),800)
	part_sn = requestCheckVar(request("part_sn"),10)
	isusing = requestCheckVar(request("isusing"),1)
	accountGubun = requestCheckVar(request("accountGubun"),5)
	department_id = requestCheckVar(request("department_id"),10)
	BIZSECTION_CD = requestCheckVar(request("BIZSECTION_CD"),10)
	locate_gubun = requestCheckVar(request("locate_gubun"),2)
	mode = requestCheckVar(request("mode"),32)
	yyyymm = requestCheckVar(request("yyyymm"),10)
	monthlyDeprice = requestCheckVar(request("monthlyDeprice"),20)
	remainValue201412 = requestCheckVar(request("remainValue201412"),20)
	info_gubun = requestCheckVar(request("info_gubun"),10)
	info_importance_C = requestCheckVar(request("info_importance_C"),10)
	info_importance_I = requestCheckVar(request("info_importance_I"),10)
	info_importance_A = requestCheckVar(request("info_importance_A"),10)
	info_gubun_idx_arr = request("info_gubun_idx_arr")
	info_gubun_val_arr = request("info_gubun_val_arr")
	arridx = requestCheckVar(request("arridx"),300)

	buy_date = Left(buy_date,10)

'response.write mode &"<Br>"
'response.end

adminuserid = session("ssBctId")
if not IsNumeric(buy_sum) then buy_sum=0

if (buy_cost = "") then
	buy_cost = CLng(buy_sum*10/11)
end if
buy_vat  = buy_sum-buy_cost

if (monthlyDeprice = "") then
	monthlyDeprice = 0
end if

if (remainValue201412 = "") then
	remainValue201412 = 0
end if

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

Select Case mode
	Case "equipmentreg"
		'/수정모드
		if idx <> "" then

			if info_gubun <> "" then
				if (info_gubun <> "-1") then
		rw "info_gubun_idx_arr:"&info_gubun_idx_arr
					if (info_gubun_idx_arr<>"") then  ''조건추가
						tmp_idx_arr = Split(info_gubun_idx_arr, "__|__")
						tmp_val_arr = Split(info_gubun_val_arr, "__|__")

						for i = 0 to UBound(tmp_idx_arr)
							if (info_gubun_idx_list = "") then
								info_gubun_idx_list = tmp_idx_arr(i)
							else
								info_gubun_idx_list = info_gubun_idx_list + "," + tmp_idx_arr(i)
							end if
						next

						sqlStr = " select top 100 info_GbnIdx "
						sqlStr = sqlStr + " from "
						sqlStr = sqlStr + " [db_partner].[dbo].[tbl_equipment_info_Dtl] "
						sqlStr = sqlStr + " where eq_idx = " & idx & " and info_GbnIdx in (" & html2db(info_gubun_idx_list) & ") "
		rw sqlStr
						rsget.pagesize = 100
						rsget.Open sqlStr, dbget,1

						if  not rsget.EOF  then
							rsget.absolutepage = 1
							do until rsget.EOF
								if (info_GbnIdx_exist_list = "") then
									info_GbnIdx_exist_list = rsget("info_GbnIdx")
								else
									info_GbnIdx_exist_list = info_GbnIdx_exist_list & "," & rsget("info_GbnIdx")
								end if

								rsget.movenext
							loop
						end if
						rsget.Close

						info_GbnIdx_exist_list = Split(info_GbnIdx_exist_list, ",")

						for i = 0 to UBound(tmp_idx_arr)
							found = False

							for j = 0 to UBound(info_GbnIdx_exist_list)
								if (tmp_idx_arr(i) = info_GbnIdx_exist_list(j)) then
									found = True
									sqlStr = " update [db_partner].[dbo].[tbl_equipment_info_Dtl] "
									sqlStr = sqlStr + " set info_value = '" & html2db(tmp_val_arr(i)) & "' "
									sqlStr = sqlStr + " where eq_idx = " & idx & " and info_GbnIdx = " & requestCheckVar(tmp_idx_arr(i),10) & " "

									'response.write sqlStr
									dbget.execute sqlStr
								end if
							next

							if found = False and tmp_idx_arr(i) <> "" then
								sqlStr = " insert into [db_partner].[dbo].[tbl_equipment_info_Dtl](eq_idx,info_GbnIdx,info_value) "
								sqlStr = sqlStr + " values(" & idx & ", " & tmp_idx_arr(i) & ",'" & html2db(tmp_val_arr(i)) & "')"

								''response.write sqlStr
								dbget.execute sqlStr
							end if
						next
					end if
				end if
			end if

			sqlStr = "insert into [db_partner].[dbo].tbl_equipment_main_log(" + VbCrlf
			sqlStr = sqlStr + " idx ,equip_code ,equip_gubun ,equip_name ,equip_spec ,equip_mainimage ,property_gubun" + VbCrlf
			sqlStr = sqlStr + " ,manufacture_sn ,manufacture_company ,manufacture_manager ,manufacture_tel" + VbCrlf
			sqlStr = sqlStr + " ,buy_company_name ,buy_date ,buy_cost ,buy_vat ,buy_sum ,using_userid ,using_date" + VbCrlf
			sqlStr = sqlStr + " ,out_date ,state ,durability_month ,etc ,part_sn ,regdate ,lastupdate ,reguserid" + VbCrlf
			sqlStr = sqlStr + " ,lastuserid ,isusing, logreguserid , logregdate, account_gubun, department_id, BIZSECTION_CD" + VbCrlf
			sqlStr = sqlStr + " , locate_gubun, monthlyDeprice, remainValue201412, accountassetcode, paymentrequestidx" + VbCrlf
			sqlStr = sqlStr + " )" + VbCrlf
			sqlStr = sqlStr + " 	select" + VbCrlf
			sqlStr = sqlStr + " 	idx ,equip_code ,equip_gubun ,equip_name ,equip_spec ,equip_mainimage ,property_gubun" + VbCrlf
			sqlStr = sqlStr + " 	,manufacture_sn ,manufacture_company ,manufacture_manager ,manufacture_tel" + VbCrlf
			sqlStr = sqlStr + " 	,buy_company_name ,buy_date ,buy_cost ,buy_vat ,buy_sum ,using_userid ,using_date" + VbCrlf
			sqlStr = sqlStr + " 	,out_date ,state ,durability_month ,etc ,part_sn ,regdate ,lastupdate ,reguserid" + VbCrlf
			sqlStr = sqlStr + " 	,lastuserid ,isusing ,'"&adminuserid&"' ,getdate(), account_gubun, department_id, BIZSECTION_CD" + VbCrlf
			sqlStr = sqlStr + " 	, locate_gubun, monthlyDeprice, remainValue201412, accountassetcode, paymentrequestidx" + VbCrlf
			sqlStr = sqlStr + " 	from [db_partner].[dbo].tbl_equipment_main" + VbCrlf
			sqlStr = sqlStr + " 	where idx = "&idx&""

			'response.write sqlStr
			dbget.execute sqlStr

			sqlStr = " update [db_partner].[dbo].tbl_equipment_main" + VbCrlf
			sqlStr = sqlStr + " set equip_gubun = '"&equip_gubun&"'" + VbCrlf
			sqlStr = sqlStr + " ,equip_name = '"&html2db(equip_name)&"'" + VbCrlf
			sqlStr = sqlStr + " ,equip_spec =  '"&html2db(equip_spec)&"'" + VbCrlf
			sqlStr = sqlStr + " ,equip_mainimage =  '"&equip_mainimage&"'" + VbCrlf
			sqlStr = sqlStr + " ,property_gubun =  '"&property_gubun&"'" + VbCrlf
			sqlStr = sqlStr + " ,manufacture_sn =  '"&html2db(manufacture_sn)&"'" + VbCrlf
			sqlStr = sqlStr + " ,manufacture_company =  '"&html2db(manufacture_company)&"'" + VbCrlf
			sqlStr = sqlStr + " ,manufacture_manager =  '"&html2db(manufacture_manager)&"'" + VbCrlf
			sqlStr = sqlStr + " ,manufacture_tel =  '"&html2db(manufacture_tel)&"'" + VbCrlf
			sqlStr = sqlStr + " ,buy_company_name =  '"&html2db(buy_company_name)&"'" + VbCrlf
			sqlStr = sqlStr + " ,using_userid =  '"&using_userid&"'" + VbCrlf
			sqlStr = sqlStr + " ,etc =  '"&html2db(etc)&"'" + VbCrlf
			sqlStr = sqlStr + " ,part_sn =  '"&part_sn&"'" + VbCrlf
			sqlStr = sqlStr + " ,lastuserid =  '"&adminuserid&"'" + VbCrlf
			sqlStr = sqlStr + " ,isusing =  '"&isusing&"'" + VbCrlf
			sqlStr = sqlStr + " ,account_gubun =  '"&accountGubun&"'" + VbCrlf
			sqlStr = sqlStr + " ,department_id =  '"& department_id &"'" + VbCrlf
			sqlStr = sqlStr + " ,BIZSECTION_CD =  '"& BIZSECTION_CD &"'" + VbCrlf
			sqlStr = sqlStr + " ,locate_gubun =  '"& locate_gubun &"'" + VbCrlf
			sqlStr = sqlStr + " ,monthlyDeprice =  '"& monthlyDeprice &"'" + VbCrlf
			sqlStr = sqlStr + " ,remainValue201412 =  '"& remainValue201412 &"'" + VbCrlf

			if using_date <> "" then
				sqlStr = sqlStr + " ,using_date =  '"&using_date&"'" + VbCrlf
			end if
			if out_date <> "" then
				sqlStr = sqlStr + " ,out_date =  '"&out_date&"'" + VbCrlf
			end if

			sqlStr = sqlStr + " ,state =  '"&state&"'" + VbCrlf

			if buy_date <> "" then
				sqlStr = sqlStr + " ,buy_date =  '"&buy_date&"'" + VbCrlf
			end if

			if info_gubun <> "" then
				if (info_gubun = "-1") then
					sqlStr = sqlStr + " ,info_gubun = NULL " + VbCrlf
				else
					sqlStr = sqlStr + " ,info_gubun =  '" & info_gubun & "'" + VbCrlf
				end if
			end if

			if info_importance_C <> "" then
				sqlStr = sqlStr + " ,info_importance_C =  '" & info_importance_C & "'" + VbCrlf
			end if

			if info_importance_I <> "" then
				sqlStr = sqlStr + " ,info_importance_I =  '" & info_importance_I & "'" + VbCrlf
			end if

			if info_importance_A <> "" then
				sqlStr = sqlStr + " ,info_importance_A =  '" & info_importance_A & "'" + VbCrlf
			end if

			if buy_cost<>"" then
				sqlStr = sqlStr + " ,buy_cost =  "&buy_cost&"" + VbCrlf
				sqlStr = sqlStr + " ,buy_vat =  "&buy_vat&"" + VbCrlf
				sqlStr = sqlStr + " ,buy_sum =  "&buy_sum&"" + VbCrlf
			end if

			sqlStr = sqlStr + " ,accountassetcode = '"& trim(html2db(accountassetcode)) &"'" + VbCrlf
			sqlStr = sqlStr + " ,paymentrequestidx = '"& trim(paymentrequestidx) &"'" + VbCrlf
			sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
			sqlStr = sqlStr + " where idx = "&idx&""

			'response.write sqlStr
			dbget.execute sqlStr

			response.write "<script type='text/javascript'>"
			response.write "	alert('OK');"
			response.write "	location.replace('/common/equipment/pop_equipmentreg.asp?idx="&idx&"');"
			response.write "	opener.location.reload();"
			response.write "</script>"

		'/신규등록
		else
			sqlStr = " select * from [db_partner].[dbo].tbl_equipment_main where 1=0" + VbCrlf

			'response.write sqlStr &"<Br>"
			rsget.Open sqlStr,dbget,1,3
			rsget.AddNew

			rsget("equip_code") = ""
			rsget("equip_gubun") = equip_gubun
			rsget("equip_name") = html2db(equip_name)
			rsget("equip_spec") = html2db(equip_spec)
			rsget("equip_mainimage") = equip_mainimage
			rsget("property_gubun") = "10"						'// property_gubun
			rsget("manufacture_sn") = html2db(manufacture_sn)
			rsget("manufacture_company") = html2db(manufacture_company)
			rsget("manufacture_manager") = html2db(manufacture_manager)
			rsget("manufacture_tel") = html2db(manufacture_tel)
			rsget("buy_company_name") = html2db(buy_company_name)
			rsget("using_userid") = using_userid
			rsget("etc") = html2db(etc)
			
			if part_sn <> "" then
				rsget("part_sn") = part_sn
			end if

			rsget("account_gubun") = accountGubun

			rsget("BIZSECTION_CD") = BIZSECTION_CD
			rsget("locate_gubun") = locate_gubun
			if department_id <> "" then
				rsget("department_id") = department_id
			end if

			rsget("reguserid") = adminuserid
			rsget("lastuserid") = adminuserid
			rsget("isusing") = isusing

			if using_date <> "" then
				rsget("using_date") = using_date
			end if
			if out_date <> "" then
				rsget("out_date") = out_date
			end if

			rsget("state") = state

			if buy_date<>"" then
				rsget("buy_date") = buy_date
			end if

			if buy_cost<>"" then
				rsget("buy_cost") = buy_cost
				rsget("buy_vat") = buy_vat
				rsget("buy_sum") = buy_sum
			end if

			rsget("durability_month") = 60

			rsget("monthlyDeprice") = monthlyDeprice
			rsget("remainValue201412") = remainValue201412

			if (info_gubun <> "") then
				rsget("info_gubun") = info_gubun
			end if

			if info_importance_C <> "" then
				rsget("info_importance_C") = info_importance_C
			end if

			if info_importance_I <> "" then
				rsget("info_importance_I") = info_importance_I
			end if

			if info_importance_A <> "" then
				rsget("info_importance_A") = info_importance_A
			end if

			if accountassetcode <> "" then
				rsget("accountassetcode") = trim(html2db(accountassetcode))
			end if

			if paymentrequestidx <> "" then
				rsget("paymentrequestidx") = trim(paymentrequestidx)
			end if

			rsget.update
				idx = rsget("idx")
			rsget.close

			equip_code = makeEquipCodeNew(idx, equip_gubun, buy_date, accountGubun)

			sqlStr = "update [db_partner].[dbo].tbl_equipment_main"
			sqlStr = sqlStr + " set equip_code='" + equip_code + "'"
			sqlStr = sqlStr + " where idx=" + CStr(idx)

			'response.write sqlStr &"<Br>"
			dbget.execute sqlStr

			if info_gubun <> "" then
				if (info_gubun <> "-1") then
					tmp_idx_arr = Split(info_gubun_idx_arr, "__|__")
					tmp_val_arr = Split(info_gubun_val_arr, "__|__")

					for i = 0 to UBound(tmp_idx_arr)
						if (Trim(tmp_idx_arr(i)) <> "") then
							sqlStr = " insert into [db_partner].[dbo].[tbl_equipment_info_Dtl](eq_idx,info_GbnIdx,info_value) "
							sqlStr = sqlStr + " values(" & idx & ", " & tmp_idx_arr(i) & ",'" & html2db(tmp_val_arr(i)) & "')"

							''response.write sqlStr
							dbget.execute sqlStr
						end if
					next
				end if
			end if

			response.write "<script type='text/javascript'>"
			response.write "	alert('ok');"
			response.write "	location.replace('/common/equipment/pop_equipmentreg.asp?idx="&idx&"');"
			response.write "	opener.location.reload();"
			response.write "</script>"
		end if

	Case "makemonthlydata"
		tmpDate = Left(DateSerial(Left(yyyymm, 4), Right(yyyymm, 2) + 1, 1), 7)

		'//신규 작성. 재작성시에도 기존 데이터는 건들이지 않음.
		sqlStr = " insert into db_partner.dbo.tbl_equipment_monthly(" & vbcrlf
		sqlStr = sqlStr + " yyyymm, idx, account_gubun, BIZSECTION_CD, buy_date, buy_sum, buy_cost, state, out_date, month_down_value" & vbcrlf
		sqlStr = sqlStr + " , month_remain_value" & vbcrlf
		sqlStr = sqlStr + " )" & vbcrlf
		sqlStr = sqlStr + " 	select " & vbcrlf
		sqlStr = sqlStr + " 	'" & yyyymm & "', e.idx, e.account_gubun , IsNull(e.BIZSECTION_CD, ''), e.buy_date, e.buy_sum, e.buy_cost " & vbcrlf
		sqlStr = sqlStr + " 	, (case " & vbcrlf
		sqlStr = sqlStr + " 			when e.state = 5 and e.out_date >= '" + CStr(tmpDate) + "-01' then 3 " & vbcrlf
		sqlStr = sqlStr + " 			else e.state end) as state " & vbcrlf
		sqlStr = sqlStr + " 	, (case " & vbcrlf
		sqlStr = sqlStr + " 			when e.state = 5 and e.out_date >= '" + CStr(tmpDate) + "-01' then NULL " & vbcrlf
		sqlStr = sqlStr + " 			else e.out_date end) as out_date " & vbcrlf
		sqlStr = sqlStr + " 	, e.monthlyDeprice as month_down_value " & vbcrlf
		sqlStr = sqlStr + " 	, (case " & vbcrlf
		sqlStr = sqlStr + " 		when IsNull(remainValue201412,0) <> 0 then " & vbcrlf
		sqlStr = sqlStr + " 			(case " & vbcrlf
		sqlStr = sqlStr + " 				when (e.remainValue201412 - (monthlyDeprice * DateDiff(m, '2014-12-01', '" & yyyymm & "-01'))) < 1000 then 1000 " & vbcrlf
		sqlStr = sqlStr + " 				else (e.remainValue201412 - (monthlyDeprice * DateDiff(m, '2014-12-01', '" & yyyymm & "-01'))) end) " & vbcrlf
		sqlStr = sqlStr + " 		else " & vbcrlf
		sqlStr = sqlStr + " 			(case " & vbcrlf
		sqlStr = sqlStr + " 				when (e.buy_cost - (monthlyDeprice * (DateDiff(m, e.buy_date, '" & yyyymm & "-01') + 1))) < 1000 then 1000 " & vbcrlf
		sqlStr = sqlStr + " 				else (e.buy_cost - (monthlyDeprice * (DateDiff(m, e.buy_date, '" & yyyymm & "-01') + 1))) end) " & vbcrlf
		sqlStr = sqlStr + " 		end) as month_remain_value " & vbcrlf
		sqlStr = sqlStr + " 	from db_partner.dbo.tbl_equipment_main e" & vbcrlf
		sqlStr = sqlStr + " 	left join db_partner.dbo.tbl_equipment_monthly m" & vbcrlf
		sqlStr = sqlStr + " 		on e.idx=m.idx" & vbcrlf
		sqlStr = sqlStr + " 		and m.yyyymm='"& yyyymm &"'" & vbcrlf
		sqlStr = sqlStr + " 	where e.isusing = 'Y' " & vbcrlf
		sqlStr = sqlStr + " 	and e.account_gubun in ('21200', '24000', '21900', '23300', '23500') " & vbcrlf
		sqlStr = sqlStr + " 	and ((e.state <> '5') or (e.state = '5' and e.out_date >= '" & yyyymm & "-01')) " & vbcrlf
		sqlStr = sqlStr + " 	and e.buy_date < '" + CStr(tmpDate) + "-01' " & vbcrlf
		sqlStr = sqlStr + " 	and m.idx is null"

		'response.write sqlStr &"<Br>"
		dbget.execute sqlStr

		response.write "<script type='text/javascript'>"
		response.write "	alert('ok');"
		response.write "	location.replace('" + CStr(refer) + "');"
		response.write "</script>"

	Case "monthlyequipmentreg"
		if yyyymm="" or idx="" then
			response.write "구분자가 없습니다."
			dbget.close() : response.end
		end if

		sqlStr = " update db_partner.dbo.tbl_equipment_monthly" + VbCrlf
		sqlStr = sqlStr + " set BIZSECTION_CD =  '"& BIZSECTION_CD &"'" + VbCrlf
		sqlStr = sqlStr + " where idx = "& idx &" and yyyymm = '" & YYYYMM & "'"

		'response.write sqlStr
		dbget.execute sqlStr

		response.write "<script type='text/javascript'>"
		response.write "	alert('OK');"
		response.write "	location.replace('/common/equipment/pop_equipmentreg_monthly.asp?yyyymm="& YYYYMM &"&idx="& idx &"');"
		response.write "	opener.location.reload();"
		response.write "</script>"

	Case "equipmentDelete"
		'// 장비자산 삭제처리
		sqlStr = " update db_partner.dbo.tbl_equipment_main " + VbCrlf
		sqlStr = sqlStr + " set isUsing='N' " + VbCrlf
		sqlStr = sqlStr + " where idx in ("& arridx &") "
		dbget.execute sqlStr

		response.write "<script type='text/javascript'>"
		response.write "	alert('OK');"
		response.write "	parent.location.reload();"
		response.write "</script>"

	Case Else
		response.write "<script type='text/javascript'>"
		response.write "	alert('구분이 지정 되지 않았습니다');"
		response.write "	self.close();"
		response.write "</script>"
end Select
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
