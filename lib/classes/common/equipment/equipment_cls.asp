<%
'####################################################
' Description : 장비자산관리 클래스
' History : 2008년 06월 27일 한용민 생성
'####################################################

class CEquipmentItem
	public fpaymentrequestidx
	public faccountassetcode
	public fidx
	public fequip_code
	public fequip_gubun
	public fequip_name
	public fequip_spec
	public fequip_mainimage
	public fproperty_gubun
	public fmanufacture_sn
	public fmanufacture_company
	public fmanufacture_manager
	public fmanufacture_tel
	public fbuy_company_name
	public Fbuy_date
	public Fbuy_cost
	public fbuy_vat
	public fbuy_sum
	public fusing_userid
	public fusing_date
	public fout_date
	public fstate
	public fdurability_month
	public fetc
	public fpart_sn
	public fregdate
	public flastupdate
	public freguserid
	public flastuserid
	public fisusing
	public fwonga_cost
	public fproperty_gubun_name
	public fstate_name
	public fstatediv
	public fusingusername
	public fpart_sn_name
	public fequip_gubun_name
	public flogreguserid
	public flogregdate
	public FaccountGubun
	public Fdepartment_id
	public FBIZSECTION_CD
	public FBIZSECTION_NM
	public Flocate_gubun
	public Flocate_gubun_name
	public FdepartmentNameFull
	public FmonthlyDeprice
	public FremainValue201412
	Public Finfo_gubun
	Public Finfo_importance_C
	Public Finfo_importance_I
	Public Finfo_importance_A
	Public Finfo_gubun_dic

	public function GetAccountGubunName()
		select case FaccountGubun
			case "21200"
				GetAccountGubunName = "비품(자산)"
			case "24000"
				GetAccountGubunName = "소프트웨어"
			case "21900"
				GetAccountGubunName = "시설장치"
			case "23300"
				GetAccountGubunName = "상표권"
			case "23500"
				GetAccountGubunName = "디자인권"
			case "81950"
				GetAccountGubunName = "임대품"
			case "83090"
				GetAccountGubunName = "소모품"
			case else
				GetAccountGubunName = FaccountGubun
		end select
	end function

	public function getDiffDate()'// 구입일 부터 현재까지 경과 개월수
	    If IsDate(Fbuy_date) then
    		if datediff("m",Fbuy_date,Now()) > 0 then
    			getDiffDate = datediff("m", Fbuy_date, Now())
    			'	datediff =		("m", 구입날짜(이전날짜),현재날짜(이후날짜))
    		end if
    	ELSE
    		getDiffDate = 0
    	end if
	end function

	'/전년 말일 대비 잔존가치	'/2016.04.19 한용민 추가
	public function getremainValue()
		dim tmpValue

		if IsNULL(Fbuy_date) or (Fbuy_date="") then exit function

		if year(Fbuy_date) <= year(dateadd("yyyy",-1,date)) then
			'21200 비품(자산)
			'24000 소프트웨어
			'21900 시설장치
			'23300 상표권
			'23500 디자인권
			'81950 임대품
			'83090 소모품
			if FaccountGubun="21200" or FaccountGubun="24000" or FaccountGubun="21900" or FaccountGubun="23300" or FaccountGubun="23500" then
				'/2014년 잔존가치가 있는거
				if FremainValue201412 <> 0 then
					tmpValue = formatNumber( FremainValue201412 - (FmonthlyDeprice * 12) ,0)
				else
					tmpValue = formatNumber( fwonga_cost - (FmonthlyDeprice * (DateDiff("m", Fbuy_date, year(dateadd("yyyy",-1,date)) & "-12-31")+1) ) ,0)
				end if
			else
				tmpValue = ""
			end if
		else
			tmpValue = ""
		end if

		getremainValue = tmpValue
	end function

	'//자산 가치 가격
	public function getCurrentValue()
		dim tmpValue

		getCurrentValue = 0
		if IsNULL(Fbuy_date) or (Fbuy_date="") then exit function

		dim ItemExpired : ItemExpired = False
		if Not IsNull(fout_date) then
			exit function
		end if

		'21200 비품(자산)
		'24000 소프트웨어
		'21900 시설장치
		'23300 상표권
		'23500 디자인권
		'81950 임대품
		'83090 소모품
		''// 구매월 말일자 기준으로 구매가(부가세) * 59/60
		''// 59개월 까지 1/60 씩 차감
		''// 60개월 이상이면 1000원(폐기처리는 경영팀에서 한다.)
		'자산구분이 자산(비품, 소프트웨어, 시설장치, 상표권, 디자인권) 일경우
		'2014년 잔존가치 가 있으면 2014년잔존가치 - (월감가상각)
		'2014년 잔존가치 가 없으면 구매일 부터 월감가상각
		'그리고 1000미만 이면 1000원
		'자산구분이 자산이 아닌 소모품 이면 0원
		select case FaccountGubun
			case "21200", "24000", "21900", "23300", "23500"
				if (FremainValue201412 <> 0) then
					tmpValue = FremainValue201412 - (FmonthlyDeprice * DateDiff("m", "2014-12-01", Now()))
					if (tmpValue < 1000) then
						tmpValue = 1000
					end if

					getCurrentValue = tmpValue
				else
					tmpValue = Fbuy_cost - (FmonthlyDeprice * (DateDiff("m", Fbuy_date, Now()) + 1))
					if (tmpValue < 1000) then
						tmpValue = 1000
					end if

					getCurrentValue = tmpValue
				end if
			case "81950"
				getCurrentValue = 0
			case "83090"
				getCurrentValue = 0
			case else
				getCurrentValue = ""
		end select

		''//getCurrentValue = fwonga_cost  - ((fwonga_cost * getDiffDate)/Fdurability_month)	'정액법 ' 차후 정률법으로 바꾸어야함 1년당 -0.451% 감소
		'현재 페이지의 자산 가치 합계 = 구입가격 - (구입가격 *구입일부터 현재까지의 날짜수/개월수)
	end function

	'// 자산가치의 총 합		'/사용안함
	public function getAllCurrentValue()
		dim SQL

		SQL = " select sum(buy_sum-(buy_sum * Datediff(m,buy_date,getdate()))/durability_month) as aaa"
		SQL = SQL + " from [db_partner].[dbo].tbl_equipment_list"

		'response.write SQL &"<Br>"
		rsget.open SQL, dbget,1

		getAllCurrentValue = rsget("aaa")
		rsget.close
	end function

	public function getEquipCode()
		getEquipCode = Fequip_code
	end function

	public function getTotalPrice()
	end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CInfoEquipmentItem
	public Fidx
	public Fequip_code
	public Fequip_gubun
	Public Faccount_gubun
	public Fequip_Name
	Public Fequip_gubun_Name
	Public Flocate_gubun_name
	Public Fstate_Name
	Public Fout_date
	Public Fisusing
    Public Fusing_userid
    Public Fusing_username
    Public Finfo_gubun
	Public Finfo_importance_C
	Public Finfo_importance_I
	Public Finfo_importance_A
	Public Finfo_gubun_dic
	Public Finfo_HOSTNM
    Public Finfo_OS
    Public Finfo_IP
    Public Finfo_place
    Public Finfo_manager
    Public Finfo_master
    Public Finfo_manageBu

	public function getCIATotalValue()
	    getCIATotalValue = Finfo_importance_C+Finfo_importance_I+Finfo_importance_A
    end function

	public function getCIATotalLevelName()
        dim ival : ival= getCIATotalValue
        if (ival>=8) then
            getCIATotalLevelName="H"
        elseif (ival>=5) then
            getCIATotalLevelName="M"
        elseif (ival>=3) then
            getCIATotalLevelName="L"
        else
            getCIATotalLevelName=""
        end if
    end function

	public function GetAccountGubunName()
		select case Faccount_gubun
			case "21200"
				GetAccountGubunName = "비품(자산)"
			case "24000"
				GetAccountGubunName = "소프트웨어"
			case "21900"
				GetAccountGubunName = "시설장치"
			case "23300"
				GetAccountGubunName = "상표권"
			case "23500"
				GetAccountGubunName = "디자인권"
			case "81950"
				GetAccountGubunName = "임대품"
			case "83090"
				GetAccountGubunName = "소모품"
			case else
				GetAccountGubunName = FaccountGubun
		end select
	end function

	public function getTotalPrice()
	end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CInfoEquipmentGubunItem
	public Finfo_GbnIdx
	public Finfo_gubun
	public Finfo_GbnName
	public Finfo_sort

	public function getTotalPrice()
	end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CEquipmentMonthlyItem
	public Fyyyymm
	public Fidx
	public FBIZSECTION_CD
	public Fbuy_date
	public Fbuy_sum
	public Fbuy_cost
	public Fstate
	public Fstate_name
	public Fout_date
	public Fprev_remain_value
	public Fmonth_down_value
	public Fmonth_remain_value
	public Fregdate
	public Faccount_gubun
	public Fequip_code
	public FBIZSECTION_NM
	public FaccountGubun

	public function GetAccMonthCount()
		GetAccMonthCount = DateDiff("m", Fbuy_date, Fyyyymm + "-01") + 1
	end function

	public function GetBuyThisMonth()
		GetBuyThisMonth = 0
		if (DateDiff("m", Fbuy_date, Fyyyymm + "-01") = 0) then
			GetBuyThisMonth = Fbuy_cost
		end if
	end function

	public function GetDiscardThisMonth()
		GetDiscardThisMonth = 0
		if Not IsNull(Fout_date) then
			if (DateDiff("m", Fout_date, Fyyyymm + "-01") = 0) then
				GetDiscardThisMonth = Fmonth_remain_value
			end if
		end if
	end function

	public function GetRemainThisMonth()
		GetRemainThisMonth = Fmonth_remain_value
		if Not IsNull(Fout_date) then
			if (DateDiff("m", Fout_date, Fyyyymm + "-01") = 0) then
				GetRemainThisMonth = 0
			end if
		end if
	end function

	public function GetAccountGubunName()
		select case Faccount_gubun
			case "21200"
				GetAccountGubunName = "비품(자산)"
			case "24000"
				GetAccountGubunName = "소프트웨어"
			case "21900"
				GetAccountGubunName = "시설장치"
			case "23300"
				GetAccountGubunName = "상표권"
			case "23500"
				GetAccountGubunName = "디자인권"
			case "81950"
				GetAccountGubunName = "임대품"
			case "83090"
				GetAccountGubunName = "소모품"
			case else
				GetAccountGubunName = FaccountGubun
		end select
	end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end class

class CEquipmentMonthlySumItem
	public Fyyyymm
	public Faccount_gubun
	public FBIZSECTION_CD
	public FBIZSECTION_NM
	public Ftot_buy_cost
	public Ftot_prev_remain_value
	public Ftot_buy_cost_this_month
	public Ftot_month_down_value
	public Ftot_month_out_value
	public Ftot_month_remain_value

	public function GetAccountGubunName()
		select case Faccount_gubun
			case "21200"
				GetAccountGubunName = "비품(자산)"
			case "24000"
				GetAccountGubunName = "소프트웨어"
			case "21900"
				GetAccountGubunName = "시설장치"
			case "23300"
				GetAccountGubunName = "상표권"
			case "23500"
				GetAccountGubunName = "디자인권"
			case "81950"
				GetAccountGubunName = "임대품"
			case "83090"
				GetAccountGubunName = "소모품"
			case else
				GetAccountGubunName = FaccountGubun
		end select
	end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end class

class CEquipment
	public FOneItem
	public FItemList()
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FTotalSum
	public FRectequip_gubun
	public FRectpart_sn
	public FRectusing_userid
	public FRectusing_username
	public Frectequip_code
	public frectequip_name
	public frectmanufacture_company
	public fRectBuyCompanyName
	public frectmanufacture_sn
	public frectbuy_startdate
	public frectbuy_enddate
	public frectout_startdate
	public frectout_enddate
	public frectproperty_gubun
	public frectstate
	public Frectidx
	public FRectIsusing
	public FRectAccountGubun
	public FRectDepartmentID
	public FRectBIZSECTION_CD
	public FRectYYYYMM
	public FRectOnly1000
	Public FRectOnlyInfoEquip
    public FRectinfo_gubun
    public FRectaccountassetcode
    public FRectpaymentrequestidx
    public frectsorttype
    
	'// 월별 자산 통계
	'/common/equipment/equipment_monthly_list.asp
	public Sub getEquipmentMonthlySUM()
		dim sqlStr, i, addSQL

		addSQL = ""
		addSQL = addSQL + " from "
		addSQL = addSQL + " 	db_partner.dbo.tbl_equipment_monthly m "
		addSQL = addSQL + " 	join db_partner.dbo.tbl_equipment_main e "
		addSQL = addSQL + " 	on "
		addSQL = addSQL + " 		1 = 1 "
		addSQL = addSQL + " 		and m.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		addSQL = addSQL + " 		and m.idx = e.idx "
		addSQL = addSQL + " 	left join db_partner.dbo.tbl_TMS_BA_BIZSECTION s on m.BIZSECTION_CD = s.BIZSECTION_CD "
		addSQL = addSQL + " 	LEFT JOIN db_partner.dbo.tbl_equipment_monthly p "
		addSQL = addSQL + " 	on "
		addSQL = addSQL + " 		1 = 1 "
		addSQL = addSQL + " 		and p.idx = m.idx "
		addSQL = addSQL + " 		AND p.yyyymm = Convert(VARCHAR(7), DateAdd(m, - 1, m.yyyymm + '-01'), 121) "
		addSQL = addSQL + " where "
		addSQL = addSQL + " 	1 = 1 "

		if FRectBIZSECTION_CD <> "" then
			addSQL = addSQL + " and m.BIZSECTION_CD = '" & FRectBIZSECTION_CD & "'"
		end if
		if FRectAccountGubun <> "" then
			addSQL = addSQL + " and e.account_gubun = '" + CStr(FRectAccountGubun) + "' "
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + addSQL

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit Sub

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)

		sqlStr = sqlStr + " 	m.yyyymm "
		sqlStr = sqlStr + " 	,e.account_gubun "
		sqlStr = sqlStr + " 	,m.BIZSECTION_CD "
		sqlStr = sqlStr + " 	,s.BIZSECTION_NM "
		sqlStr = sqlStr + " 	,sum(m.buy_cost) as tot_buy_cost "
		sqlStr = sqlStr + " 	,sum(IsNull(p.month_remain_value, 0)) AS tot_prev_remain_value "
		sqlStr = sqlStr + " 	,sum(case when DateDiff(m, m.buy_date, (m.yyyymm+'-01')) = 0 then m.buy_cost else 0 end) as tot_buy_cost_this_month "
		sqlStr = sqlStr + " 	,sum(m.month_down_value) as tot_month_down_value "
		sqlStr = sqlStr + " 	,sum(case when m.STATE = '5' then m.month_remain_value else 0 end) as tot_month_out_value "
		sqlStr = sqlStr + " 	,sum(case when m.STATE <> '5' then m.month_remain_value else 0 end) as tot_month_remain_value "
		sqlStr = sqlStr + addSQL
		sqlStr = sqlStr + " group by m.yyyymm,e.account_gubun,m.BIZSECTION_CD,s.BIZSECTION_NM "
		sqlStr = sqlStr + " order by m.yyyymm, e.account_gubun, m.BIZSECTION_CD "

		''response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEquipmentMonthlySumItem
					FItemList(i).Fyyyymm 					= rsget("yyyymm")
					FItemList(i).Faccount_gubun 			= rsget("account_gubun")
					FItemList(i).FBIZSECTION_CD 			= rsget("BIZSECTION_CD")
					FItemList(i).FBIZSECTION_NM 			= rsget("BIZSECTION_NM")
					FItemList(i).Ftot_buy_cost 				= rsget("tot_buy_cost")
					FItemList(i).Ftot_prev_remain_value 	= rsget("tot_prev_remain_value")
					FItemList(i).Ftot_buy_cost_this_month 	= rsget("tot_buy_cost_this_month")
					FItemList(i).Ftot_month_down_value 		= rsget("tot_month_down_value")
					FItemList(i).Ftot_month_out_value 		= rsget("tot_month_out_value")
					FItemList(i).Ftot_month_remain_value 	= rsget("tot_month_remain_value")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	'// 월별 자산
	'/common/equipment/equipment_monthly_list.asp
	public Sub getEquipmentMonthlyList()
		dim sqlStr, i, addSQL

		addSQL = ""
		addSQL = addSQL + " from "
		addSQL = addSQL + " 	db_partner.dbo.tbl_equipment_monthly m "
		addSQL = addSQL + " 	join db_partner.dbo.tbl_equipment_main e "
		addSQL = addSQL + " 	on "
		addSQL = addSQL + " 		1 = 1 "
		addSQL = addSQL + " 		and m.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		addSQL = addSQL + " 		and m.idx = e.idx "
		addSQL = addSQL + " 	left join db_partner.dbo.tbl_TMS_BA_BIZSECTION s on m.BIZSECTION_CD = s.BIZSECTION_CD "
		addSQL = addSQL + " where "
		addSQL = addSQL + " 	1 = 1 "
		''addSQL = addSQL + " 	and ((m.state <> '5') or (m.state = '5' and m.out_date >= '2014-01-01')) "

		if FRectBIZSECTION_CD <> "" then
			addSQL = addSQL + " and m.BIZSECTION_CD = '" & FRectBIZSECTION_CD & "'"
		end if
		if FRectAccountGubun <> "" then
			addSQL = addSQL + " and e.account_gubun = '" + CStr(FRectAccountGubun) + "' "
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + addSQL

		''response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit Sub

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " m.yyyymm, m.idx, m.BIZSECTION_CD, m.buy_date, m.buy_sum, m.buy_cost, m.state, m.out_date, m.month_down_value"
		sqlStr = sqlStr + " , m.month_remain_value, m.regdate, s.BIZSECTION_NM, e.account_gubun, e.equip_code "
		sqlStr = sqlStr + " , IsNull(( "
		sqlStr = sqlStr + " 	select top 1 p.month_remain_value as month_remain_value "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	db_partner.dbo.tbl_equipment_monthly p "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and p.yyyymm = Convert(varchar(7), DateAdd(m, -1, m.yyyymm + '-01'), 121) "
		sqlStr = sqlStr + " 		and p.idx = m.idx), 0) "
		sqlStr = sqlStr + " as prev_remain_value "
		sqlStr = sqlStr + " , ("
		sqlStr = sqlStr + " 	select top 1 gubunname"
		sqlStr = sqlStr + " 	from [db_partner].[dbo].tbl_equipment_gubun"
		sqlStr = sqlStr + "  	where gubuntype='50' and m.state=gubuncd and isusing='Y' "
		sqlStr = sqlStr + "  ) as state_name"
		sqlStr = sqlStr + addSQL
		sqlStr = sqlStr + " order by m.yyyymm desc, m.idx desc "

		''response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEquipmentMonthlyItem
					FItemList(i).Fyyyymm 			= rsget("yyyymm")
					FItemList(i).Fidx 				= rsget("idx")
					FItemList(i).FBIZSECTION_CD 	= rsget("BIZSECTION_CD")
					FItemList(i).Fbuy_date 			= rsget("buy_date")
					FItemList(i).Fbuy_sum 			= rsget("buy_sum")
					FItemList(i).Fbuy_cost 			= rsget("buy_cost")
					FItemList(i).Fstate 			= rsget("state")
					FItemList(i).Fstate_name 		= rsget("state_name")
					FItemList(i).Fout_date 			= rsget("out_date")
					FItemList(i).Fprev_remain_value 	= rsget("prev_remain_value")
					FItemList(i).Fmonth_down_value 		= rsget("month_down_value")
					FItemList(i).Fmonth_remain_value 	= rsget("month_remain_value")
					FItemList(i).Fregdate 				= rsget("regdate")
					FItemList(i).Faccount_gubun 		= rsget("account_gubun")
					FItemList(i).Fequip_code 			= rsget("equip_code")
					FItemList(i).FBIZSECTION_NM 		= rsget("BIZSECTION_NM")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	'//common/equipment/equipment_list.asp
	public Sub getEquipmentList()
		dim sqlStr, i, addSQL

		if frectaccountassetcode <> "" then
			addSQL = addSQL + " and l.accountassetcode = '" & frectaccountassetcode & "'"
		end if
		if frectpaymentrequestidx <> "" then
			addSQL = addSQL + " and l.paymentrequestidx = " & frectpaymentrequestidx & ""
		end if
		if frectstate <> "" then
			addSQL = addSQL + " and l.state = '" & frectstate & "'"
		end if
		if FRectBIZSECTION_CD <> "" then
			addSQL = addSQL + " and l.BIZSECTION_CD = '" & FRectBIZSECTION_CD & "'"
		end if
		if frectproperty_gubun <> "" then
			addSQL = addSQL + " and l.property_gubun = '" & frectproperty_gubun & "'"
		end if
		if FRectequip_gubun <> "" then
			addSQL = addSQL + " and l.equip_gubun = '" & FRectequip_gubun & "'"
		end if
		if FRectpart_sn <> "" then
			addSQL = addSQL + " and l.part_sn = '" & FRectpart_sn & "'"
		end if
		if FRectusing_userid <> "" then
			addSQL = addSQL + " and l.using_userid = '" & FRectusing_userid & "'"
		end if
		if FRectusing_username <> "" then
			if frectproperty_gubun = "2" then
				addSQL = addSQL + " and u.shopname = '" & FRectusing_username & "' "
			else
				addSQL = addSQL + " and u.username = '" & FRectusing_username & "' "
			end if
		end if
		if Frectequip_code <> "" then
			addSQL = addSQL + " and l.equip_code like '%" & Frectequip_code & "%'"
		end if
		if Frectequip_name <> "" then
			addSQL = addSQL + " and l.equip_name like '%" & Frectequip_name & "%'"
		end if
		if frectbuy_startdate <> "" and frectbuy_enddate <> "" then
			if frectbuy_startdate <>"" then
				addSQL = addSQL + " and l.buy_date>='" + frectbuy_startdate + "'"
			end if

			if frectbuy_enddate <>"" then
				addSQL = addSQL + " and l.buy_date<'" + frectbuy_enddate + "'"
			end if
		end if
		if frectout_startdate <> "" and frectout_enddate <> "" then
			if frectout_startdate <>"" then
				addSQL = addSQL + " and l.out_date>='" + frectout_startdate + "'"
			end if

			if frectout_enddate <>"" then
				addSQL = addSQL + " and l.out_date<'" + frectout_enddate + "'"
			end if
		end if
		if frectmanufacture_company <>"" then
			addSQL = addSQL + " and manufacture_company like '%" & frectmanufacture_company & "%'"
		end if
		if fRectBuyCompanyName <>"" then
			addSQL = addSQL + " and l.buy_company_name like '%" & fRectBuyCompanyName & "%'"
		end if
		if frectmanufacture_sn <>"" then
			addSQL = addSQL + " and manufacture_sn like '" & frectmanufacture_sn & "%'"
		end if
		if FRectIsusing <> "" then
			addSQL = addSQL + " and l.isusing = 'Y' "
		end if
		if FRectAccountGubun <> "" then
			addSQL = addSQL + " and l.account_gubun = '" + CStr(FRectAccountGubun) + "' "
		end if
		if FRectDepartmentID <> "" then
			addSQL = addSQL + " and l.department_id = '" + CStr(FRectDepartmentID) + "' "
		end if
		if FRectOnly1000 <> "" then
			addSQL = addSQL + " and l.account_Gubun in ('21200', '24000', '21900', '23300', '23500') "
			addSQL = addSQL + " and ( "
			addSQL = addSQL + " 	((IsNull(l.remainValue201412,0) <> 0) and ((IsNull(l.remainValue201412,0) - monthlyDeprice * DateDiff(m, '2014-12-01', getdate())) <= 1000)) "
			addSQL = addSQL + " 	or "
			addSQL = addSQL + " 	((IsNull(l.remainValue201412,0) = 0) and ((IsNull(l.buy_cost,0) - monthlyDeprice * DateDiff(m, l.buy_date, getdate())) <= 1000)) "
			addSQL = addSQL + " ) "
		end if

		sqlStr = " select count(*) as cnt, sum(l.buy_sum) as totalprice"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_main l"

		if frectproperty_gubun = "2" then
			sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user u"
			sqlStr = sqlStr + " 	on l.using_userid = u.userid"
			sqlStr = sqlStr + " 	and u.isusing='Y'"
		else
			sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten u"
			sqlStr = sqlStr + " 	on l.using_userid=u.userid"
			sqlStr = sqlStr + " 	and u.isUsing = 1"
			sqlStr = sqlStr + " 	and isnull(u.userid,'') <> ''"
		end if

		sqlStr = sqlStr + " where 1=1 " & addSQL

		'rw sqlStr
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalSum = rsget("totalprice")
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit Sub

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " l.idx, l.equip_code ,l.equip_gubun ,l.equip_name ,l.equip_spec ,l.equip_mainimage, l.property_gubun ,l.manufacture_sn"
		sqlStr = sqlStr + " , l.manufacture_company, l.manufacture_manager, l.manufacture_tel, l.buy_company_name, l.buy_date, l.buy_cost"
		sqlStr = sqlStr + " , l.buy_vat, l.buy_sum, l.using_userid, l.using_date, l.out_date, l.state, l.durability_month, l.etc, l.part_sn"
		sqlStr = sqlStr + " , l.regdate, l.lastupdate, l.reguserid, l.lastuserid, l.isusing, l.account_gubun, l.department_id, l.BIZSECTION_CD"
		sqlStr = sqlStr + " , l.locate_gubun, IsNull(l.monthlyDeprice, 0) as monthlyDeprice, IsNull(l.remainValue201412, 0) as remainValue201412 "
		''sqlStr = sqlStr + " , v.departmentNameFull, s.BIZSECTION_NM "
		sqlStr = sqlStr + " , '' as departmentNameFull, s.BIZSECTION_NM, l.accountassetcode, l.paymentrequestidx"

		if frectproperty_gubun = "2" then
			sqlStr = sqlStr + " , u.shopname as usingusername, 'Y' as statediv"
		else
			sqlStr = sqlStr + " , u.username as usingusername, u.statediv"
		end if

		sqlStr = sqlStr + " ,("
		sqlStr = sqlStr + " 	select top 1 gubunname"
		sqlStr = sqlStr + " 	from [db_partner].[dbo].tbl_equipment_gubun"
		sqlStr = sqlStr + "  	where gubuntype='10' and l.equip_gubun=gubuncd and isusing='Y'"
		sqlStr = sqlStr + "  )as equip_gubun_name"
		sqlStr = sqlStr + " ,'' as part_sn_name"
		sqlStr = sqlStr + " ,'' as property_gubun_name"
		sqlStr = sqlStr + " ,("
		sqlStr = sqlStr + " 	select top 1 gubunname"
		sqlStr = sqlStr + " 	from [db_partner].[dbo].tbl_equipment_gubun"
		sqlStr = sqlStr + "  	where gubuntype='50' and l.state=gubuncd and isusing='Y'"
		sqlStr = sqlStr + "  )as state_name"
		sqlStr = sqlStr + " ,("
		sqlStr = sqlStr + " 	select top 1 gubunname"
		sqlStr = sqlStr + " 	from [db_partner].[dbo].tbl_equipment_gubun"
		sqlStr = sqlStr + "  	where gubuntype='30' and l.locate_gubun=gubuncd and isusing='Y'"
		sqlStr = sqlStr + "  )as locate_gubun_name"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_main l"
		''sqlStr = sqlStr + " left join db_partner.dbo.vw_user_department v on l.department_id = v.cid "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_TMS_BA_BIZSECTION s on l.BIZSECTION_CD = s.BIZSECTION_CD "

		if frectproperty_gubun = "2" then
			sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user u"
			sqlStr = sqlStr + " 	on l.using_userid = u.userid"
			sqlStr = sqlStr + " 	and u.isusing='Y'"
		else
			sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten u"
			sqlStr = sqlStr + " 	on l.using_userid=u.userid"
			sqlStr = sqlStr + " 	and u.isUsing = 1"
			sqlStr = sqlStr + " 	and isnull(u.userid,'') <> ''"
		end if

		sqlStr = sqlStr + " where 1=1 " & addSQL

		if frectsorttype="1" then
			sqlStr = sqlStr + " order by l.idx desc"
		elseif frectsorttype="2" then
			sqlStr = sqlStr + " order by l.buy_date desc, l.idx desc"
		elseif frectsorttype="3" then
			sqlStr = sqlStr + " order by l.buy_date asc, l.idx asc"
		else
			sqlStr = sqlStr + " order by l.idx desc"			
		end if

		'rw sqlStr
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        
		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEquipmentItem
					FItemList(i).faccountassetcode = db2html(rsget("accountassetcode"))
					FItemList(i).fpaymentrequestidx = rsget("paymentrequestidx")
					FItemList(i).fidx = rsget("idx")
					FItemList(i).fequip_code = rsget("equip_code")
					FItemList(i).fequip_gubun = rsget("equip_gubun")
					FItemList(i).fequip_name = db2html(rsget("equip_name"))
					FItemList(i).fequip_spec = db2html(rsget("equip_spec"))
					FItemList(i).fequip_mainimage = rsget("equip_mainimage")
					FItemList(i).fproperty_gubun = rsget("property_gubun")
					FItemList(i).fmanufacture_sn = db2html(rsget("manufacture_sn"))
					FItemList(i).fmanufacture_company = db2html(rsget("manufacture_company"))
					FItemList(i).fmanufacture_manager = db2html(rsget("manufacture_manager"))
					FItemList(i).fmanufacture_tel = db2html(rsget("manufacture_tel"))
					FItemList(i).fbuy_company_name = db2html(rsget("buy_company_name"))
					FItemList(i).fbuy_date = rsget("buy_date")
					FItemList(i).fbuy_cost = rsget("buy_cost")
					FItemList(i).fbuy_vat = rsget("buy_vat")
					FItemList(i).fbuy_sum = rsget("buy_sum")
					FItemList(i).fusing_userid = rsget("using_userid")
					FItemList(i).fusing_date = rsget("using_date")
					FItemList(i).fout_date = rsget("out_date")
					FItemList(i).fstate = rsget("state")
					FItemList(i).fdurability_month = rsget("durability_month")
					FItemList(i).fetc = db2html(rsget("etc"))
					FItemList(i).fpart_sn = rsget("part_sn")
					FItemList(i).fregdate = rsget("regdate")
					FItemList(i).flastupdate = rsget("lastupdate")
					FItemList(i).freguserid = rsget("reguserid")
					FItemList(i).flastuserid = rsget("lastuserid")
					FItemList(i).fisusing = rsget("isusing")
					FItemList(i).fproperty_gubun_name = db2html(rsget("property_gubun_name"))
					FItemList(i).fstate_name = db2html(rsget("state_name"))
					FItemList(i).fwonga_cost          = rsget("buy_cost")				'// / 1.1
					FItemList(i).fstatediv = rsget("statediv")
					FItemList(i).fusingusername = rsget("usingusername")
					FItemList(i).fpart_sn_name = rsget("part_sn_name")
					FItemList(i).fequip_gubun_name = rsget("equip_gubun_name")
					FItemList(i).FaccountGubun = rsget("account_gubun")
					FItemList(i).Fdepartment_id = rsget("department_id")
					FItemList(i).FBIZSECTION_CD = rsget("BIZSECTION_CD")
					FItemList(i).FBIZSECTION_NM = rsget("BIZSECTION_NM")
					FItemList(i).Flocate_gubun = rsget("locate_gubun")
					FItemList(i).Flocate_gubun_name = rsget("locate_gubun_name")
					FItemList(i).FdepartmentNameFull = rsget("departmentNameFull")
					FItemList(i).FmonthlyDeprice = rsget("monthlyDeprice")
					FItemList(i).FremainValue201412 = rsget("remainValue201412")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	public Sub getInfoEquipmentList()
		dim sqlStr, i, addSQL

		addSQL = ""

		If (FRectOnlyInfoEquip <> "") Then
			addSQL = addSQL + " and l.info_gubun is not NULL "
		End If

		if FRectIsusing <> "" then
			addSQL = addSQL + " and l.isusing = 'Y' "
		end If

		if FRectequip_gubun <> "" then
			addSQL = addSQL + " and l.equip_gubun = '" & FRectequip_gubun & "'"
		end if

        if Frectequip_code <> "" then
			addSQL = addSQL + " and l.equip_code like '%" & Frectequip_code & "%'"
		end if
        
        if FRectinfo_gubun <> "" then
			addSQL = addSQL + " and l.info_gubun = '" & info_gubun & "'"
		end if
		
		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_main l"

		sqlStr = sqlStr + " where 1=1 " & addSQL

		'rw sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit Sub

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " l.idx, l.equip_code ,l.equip_gubun, l.account_gubun, l.equip_name, eg.gubunname as equip_gubun_name, lg.gubunname as locate_gubun_name, sg.gubunname as state_Name "
		sqlStr = sqlStr + " , l.isusing, l.out_date "
		sqlStr = sqlStr + " , l.using_userid"
		sqlStr = sqlStr + " , (select top 1 username from db_partner.dbo.tbl_user_tenbyten t where t.userid=l.using_userid and isNULL(t.userid,'')<>'') as using_username "
		sqlStr = sqlStr + " , l.info_gubun"
		sqlStr = sqlStr + " , isNULL(l.info_importance_C,0) as info_importance_C"
		sqlStr = sqlStr + " , isNULL(l.info_importance_I,0) as info_importance_I"
		sqlStr = sqlStr + " , isNULL(l.info_importance_A,0) as info_importance_A"
		sqlStr = sqlStr + " , V.info_HOSTNM"
		sqlStr = sqlStr + " , V.info_OS"
		sqlStr = sqlStr + " , V.info_IP"
		sqlStr = sqlStr + " , V.info_place"
		sqlStr = sqlStr + " , V.info_manager"
		sqlStr = sqlStr + " , V.info_master"
		sqlStr = sqlStr + " , V.info_manageBu"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_main l"
		sqlStr = sqlStr + " left join db_partner.dbo.vw_InfoEquipment_pivot V"
		sqlStr = sqlStr + " on l.idx=V.eq_idx"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun eg "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + "     1 = 1 "
		sqlStr = sqlStr + "     and eg.gubuntype='10' "
		sqlStr = sqlStr + "     and l.equip_gubun=eg.gubuncd "
		sqlStr = sqlStr + "     and eg.isusing='Y' "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun lg "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + "     1 = 1 "
		sqlStr = sqlStr + "     and lg.gubuntype='30' "
		sqlStr = sqlStr + "     and l.locate_gubun=lg.gubuncd "
		sqlStr = sqlStr + "     and lg.isusing='Y' "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_equipment_gubun sg "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + "     1 = 1 "
		sqlStr = sqlStr + "     and sg.gubuntype='50' "
		sqlStr = sqlStr + "     and l.state=sg.gubuncd "
		sqlStr = sqlStr + "     and sg.isusing='Y' "
		sqlStr = sqlStr + " where 1=1 " & addSQL
		sqlStr = sqlStr + " order by l.idx desc"

		'rw sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CInfoEquipmentItem

				FItemList(i).Fidx = rsget("idx")
				FItemList(i).Fequip_code = rsget("equip_code")
				FItemList(i).Fequip_gubun = rsget("equip_gubun")
				FItemList(i).Faccount_gubun = rsget("account_gubun")
				FItemList(i).Fequip_name = db2html(rsget("equip_Name"))
				FItemList(i).Fequip_gubun_name = rsget("equip_gubun_Name")
				FItemList(i).Flocate_gubun_name = rsget("locate_gubun_Name")
				FItemList(i).Fstate_name = db2html(rsget("state_Name"))
				FItemList(i).Fout_date = rsget("out_date")
				FItemList(i).Fisusing = rsget("isusing")
                FItemList(i).Fusing_userid      = rsget("using_userid")
                FItemList(i).Fusing_username    = rsget("using_username")
                FItemList(i).Finfo_gubun            = rsget("info_gubun")
                FItemList(i).Finfo_importance_C     = rsget("info_importance_C")
                FItemList(i).Finfo_importance_I     = rsget("info_importance_I")
                FItemList(i).Finfo_importance_A     = rsget("info_importance_A")
                FItemList(i).Finfo_HOSTNM   = rsget("info_HOSTNM")
                FItemList(i).Finfo_OS       = rsget("info_OS")
                FItemList(i).Finfo_IP       = rsget("info_IP")
                FItemList(i).Finfo_place    = rsget("info_place")
                FItemList(i).Finfo_manager  = rsget("info_manager")
                FItemList(i).Finfo_master   = rsget("info_master")
                FItemList(i).Finfo_manageBu = rsget("info_manageBu")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	public Sub getInfoEquipmentGubunList()
		dim sqlStr, i, addSQL

		sqlStr = " select top " & Cstr(FPageSize * FCurrPage) & " * "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_partner.dbo.tbl_equipment_info_Gbn "
		sqlStr = sqlStr + " order by info_gubun, info_sort "

		''response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CInfoEquipmentGubunItem

				FItemList(i).Finfo_GbnIdx 		= rsget("info_GbnIdx")
				FItemList(i).Finfo_gubun 		= rsget("info_gubun")
				FItemList(i).Finfo_GbnName 		= rsget("info_GbnName")
				FItemList(i).Finfo_sort 		= rsget("info_sort")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	'//common/equipment/equipment_list.asp
	public Sub getEquipmentlogList()
		dim sqlStr, i, addSQL

		if frectidx <>"" then
			addSQL = addSQL + " and ll.idx="&frectidx&""
		end if

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " ll.idx, ll.equip_code, ll.equip_gubun, ll.equip_name, ll.equip_spec, ll.equip_mainimage, ll.property_gubun"
		sqlStr = sqlStr + " ,ll.manufacture_sn ,ll.manufacture_company ,ll.manufacture_manager, ll.manufacture_tel,ll.buy_company_name"
		sqlStr = sqlStr + " ,ll.buy_date ,ll.buy_cost ,ll.buy_vat ,ll.buy_sum, ll.using_userid, ll.using_date, ll.out_date, ll.state"
		sqlStr = sqlStr + " ,ll.durability_month, ll.etc, ll.part_sn,ll.regdate ,ll.lastupdate ,ll.reguserid ,ll.lastuserid ,ll.isusing"
		sqlStr = sqlStr + " , ll.account_gubun, ll.department_id, ll.accountassetcode, ll.paymentrequestidx, v.departmentNameFull"
		sqlStr = sqlStr + " ,(case"
		sqlStr = sqlStr + " 	when ll.property_gubun = 2 then u.shopname"
		sqlStr = sqlStr + " 	else tu.username"
		sqlStr = sqlStr + " end) as usingusername"
		sqlStr = sqlStr + " ,(case"
		sqlStr = sqlStr + " 	when ll.property_gubun = 2 then 'Y'"
		sqlStr = sqlStr + " 	else tu.statediv"
		sqlStr = sqlStr + " end) as statediv"
		sqlStr = sqlStr + " ,("
		sqlStr = sqlStr + " 	select top 1 gubunname"
		sqlStr = sqlStr + " 	from [db_partner].[dbo].tbl_equipment_gubun"
		sqlStr = sqlStr + "  	where gubuntype='10' and ll.equip_gubun=gubuncd and isusing='Y'"
		sqlStr = sqlStr + "  )as equip_gubun_name"
		sqlStr = sqlStr + " ,("
		sqlStr = sqlStr + " 	select top 1 part_name"
		sqlStr = sqlStr + " 	from db_partner.dbo.tbl_partInfo"
		sqlStr = sqlStr + "  	where ll.part_sn=part_sn and part_isdel='N'"
		sqlStr = sqlStr + "  )as part_sn_name"
		sqlStr = sqlStr + " ,("
		sqlStr = sqlStr + " 	select top 1 gubunname"
		sqlStr = sqlStr + " 	from [db_partner].[dbo].tbl_equipment_gubun"
		sqlStr = sqlStr + "  	where gubuntype='40' and ll.property_gubun=gubuncd and isusing='Y'"
		sqlStr = sqlStr + "  )as property_gubun_name"
		sqlStr = sqlStr + " ,("
		sqlStr = sqlStr + " 	select top 1 gubunname"
		sqlStr = sqlStr + " 	from [db_partner].[dbo].tbl_equipment_gubun"
		sqlStr = sqlStr + "  	where gubuntype='50' and ll.state=gubuncd and isusing='Y'"
		sqlStr = sqlStr + "  )as state_name"
		sqlStr = sqlStr + " ,logreguserid ,logregdate"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_main_log ll"
		sqlStr = sqlStr + " left join db_partner.dbo.vw_user_department v on ll.department_id = v.cid "
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr + " 	on ll.using_userid = u.userid"
		sqlStr = sqlStr + " 	and u.isusing='Y'"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten tu"
		sqlStr = sqlStr + " 	on ll.using_userid=tu.userid"
		sqlStr = sqlStr + " 	and tu.isUsing = 1"
		sqlStr = sqlStr + " 	and isnull(tu.userid,'') <> ''"
		sqlStr = sqlStr + " where 1=1 " & addSQL
		sqlStr = sqlStr + " order by ll.logidx desc"

		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget,1

		FResultCount =  rsget.RecordCount
		ftotalcount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEquipmentItem
					FItemList(i).faccountassetcode = db2html(rsget("accountassetcode"))
					FItemList(i).fpaymentrequestidx = rsget("paymentrequestidx")
					FItemList(i).fidx = rsget("idx")
					FItemList(i).fequip_code = rsget("equip_code")
					FItemList(i).fequip_gubun = rsget("equip_gubun")
					FItemList(i).fequip_name = db2html(rsget("equip_name"))
					FItemList(i).fequip_spec = db2html(rsget("equip_spec"))
					FItemList(i).fequip_mainimage = rsget("equip_mainimage")
					FItemList(i).fproperty_gubun = rsget("property_gubun")
					FItemList(i).fmanufacture_sn = db2html(rsget("manufacture_sn"))
					FItemList(i).fmanufacture_company = db2html(rsget("manufacture_company"))
					FItemList(i).fmanufacture_manager = db2html(rsget("manufacture_manager"))
					FItemList(i).fmanufacture_tel = db2html(rsget("manufacture_tel"))
					FItemList(i).fbuy_company_name = db2html(rsget("buy_company_name"))
					FItemList(i).fbuy_date = rsget("buy_date")
					FItemList(i).fbuy_cost = rsget("buy_cost")
					FItemList(i).fbuy_vat = rsget("buy_vat")
					FItemList(i).fbuy_sum = rsget("buy_sum")
					FItemList(i).fusing_userid = rsget("using_userid")
					FItemList(i).fusing_date = rsget("using_date")
					FItemList(i).fout_date = rsget("out_date")
					FItemList(i).fstate = rsget("state")
					FItemList(i).fdurability_month = rsget("durability_month")
					FItemList(i).fetc = db2html(rsget("etc"))
					FItemList(i).fpart_sn = rsget("part_sn")
					FItemList(i).fregdate = rsget("regdate")
					FItemList(i).flastupdate = rsget("lastupdate")
					FItemList(i).freguserid = rsget("reguserid")
					FItemList(i).flastuserid = rsget("lastuserid")
					FItemList(i).fisusing = rsget("isusing")
					FItemList(i).fproperty_gubun_name = db2html(rsget("property_gubun_name"))
					FItemList(i).fstate_name = db2html(rsget("state_name"))
					FItemList(i).fwonga_cost          = rsget("buy_sum") / 1.1
					FItemList(i).fstatediv = rsget("statediv")
					FItemList(i).fusingusername = rsget("usingusername")
					FItemList(i).fpart_sn_name = rsget("part_sn_name")
					FItemList(i).fequip_gubun_name = rsget("equip_gubun_name")
					FItemList(i).flogreguserid = rsget("logreguserid")
					FItemList(i).flogregdate = rsget("logregdate")
					FItemList(i).FaccountGubun = rsget("account_gubun")
					FItemList(i).Fdepartment_id = rsget("department_id")
					FItemList(i).FdepartmentNameFull = rsget("departmentNameFull")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	'//common/equipment/pop_equipmentreg.asp
	public Sub getOneEquipment()
		dim sqlStr, i ,addSQL

		if frectidx <> "" then
			addSQL = addSQL & " and l.idx = "&frectidx&""
		end if

		sqlStr = "select top 1"
		sqlStr = sqlStr + " l.idx, l.equip_code ,l.equip_gubun ,l.equip_name ,l.equip_spec ,l.equip_mainimage,l.property_gubun"
		sqlStr = sqlStr + " ,l.manufacture_sn,l.manufacture_company ,l.manufacture_manager,l.manufacture_tel ,l.buy_company_name"
		sqlStr = sqlStr + " ,l.buy_date ,l.buy_cost ,l.buy_vat ,l.buy_sum, l.using_userid, l.using_date ,l.out_date ,l.state"
		sqlStr = sqlStr + " ,l.durability_month, l.etc ,l.part_sn ,l.regdate ,l.lastupdate ,l.reguserid ,l.lastuserid, l.isusing"
		sqlStr = sqlStr + " , l.account_gubun, l.department_id, l.BIZSECTION_CD, s.BIZSECTION_NM, l.locate_gubun, l.monthlyDeprice"
		sqlStr = sqlStr + " , l.remainValue201412, IsNull(l.info_gubun, -1) as info_gubun, info_importance_C, info_importance_I"
		sqlStr = sqlStr + " , info_importance_A, l.accountassetcode, l.paymentrequestidx"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_main l"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_TMS_BA_BIZSECTION s on l.BIZSECTION_CD = s.BIZSECTION_CD "
		sqlStr = sqlStr + " where 1=1 " & addSQL

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

		ftotalcount =  rsget.RecordCount
		set FOneItem = new CEquipmentItem
		i=0
		if not rsget.EOF  then
			FOneItem.faccountassetcode = db2html(rsget("accountassetcode"))
			FOneItem.fpaymentrequestidx = rsget("paymentrequestidx")
			FOneItem.fidx = rsget("idx")
			FOneItem.fequip_code = rsget("equip_code")
			FOneItem.fequip_gubun = rsget("equip_gubun")
			FOneItem.fequip_name = db2html(rsget("equip_name"))
			FOneItem.fequip_spec = db2html(rsget("equip_spec"))
			FOneItem.fequip_mainimage = rsget("equip_mainimage")
			FOneItem.fproperty_gubun = rsget("property_gubun")
			FOneItem.fmanufacture_sn = db2html(rsget("manufacture_sn"))
			FOneItem.fmanufacture_company = db2html(rsget("manufacture_company"))
			FOneItem.fmanufacture_manager = db2html(rsget("manufacture_manager"))
			FOneItem.fmanufacture_tel = db2html(rsget("manufacture_tel"))
			FOneItem.fbuy_company_name = db2html(rsget("buy_company_name"))
			FOneItem.fbuy_date = rsget("buy_date")
			FOneItem.fbuy_cost = rsget("buy_cost")
			FOneItem.fbuy_vat = rsget("buy_vat")
			FOneItem.fbuy_sum = rsget("buy_sum")
			FOneItem.fusing_userid = rsget("using_userid")
			FOneItem.fusing_date = rsget("using_date")
			FOneItem.fout_date = rsget("out_date")
			FOneItem.fstate = rsget("state")
			FOneItem.fdurability_month = rsget("durability_month")
			FOneItem.fetc = db2html(rsget("etc"))
			FOneItem.fpart_sn = rsget("part_sn")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.flastupdate = rsget("lastupdate")
			FOneItem.freguserid = rsget("reguserid")
			FOneItem.flastuserid = rsget("lastuserid")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.FaccountGubun = rsget("account_gubun")
			FOneItem.Fdepartment_id = rsget("department_id")
			FOneItem.FBIZSECTION_CD = rsget("BIZSECTION_CD")
			FOneItem.FBIZSECTION_NM = db2html(rsget("BIZSECTION_NM"))
			FOneItem.Flocate_gubun = rsget("locate_gubun")
			FOneItem.FmonthlyDeprice = rsget("monthlyDeprice")
			FOneItem.FremainValue201412 = rsget("remainValue201412")	'/현재년도의 바로 직전년도 12월 금액
			FOneItem.Finfo_gubun = rsget("info_gubun")
			FOneItem.Finfo_importance_C = rsget("info_importance_C")
			FOneItem.Finfo_importance_I = rsget("info_importance_I")
			FOneItem.Finfo_importance_A = rsget("info_importance_A")

		end if
		rsget.Close

		Set FOneItem.Finfo_gubun_dic = Server.CreateObject("Scripting.Dictionary")

		If FOneItem.fidx <> "" And FOneItem.Finfo_gubun <> "" And FOneItem.Finfo_gubun <> "-1" then
			sqlStr = " select top 100 info_GbnIdx,info_value "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " [db_partner].[dbo].[tbl_equipment_info_Dtl] "
			sqlStr = sqlStr + " where eq_idx = " & FOneItem.fidx & " "

			rsget.pagesize = 100
			rsget.Open sqlStr, dbget,1

			if  not rsget.EOF  then
				rsget.absolutepage = 1
				do until rsget.EOF
					FOneItem.Finfo_gubun_dic.Add CStr(rsget("info_GbnIdx")), db2html(rsget("info_value"))
					rsget.movenext
				loop
			end if
			rsget.Close
		End If
	end Sub

	'//common/equipment/pop_equipmentreg_monthly.asp
    public Sub getOneEquipment_monthly()
		dim sqlStr, i ,addSQL

		if frectidx <> "" then
			addSQL = addSQL & " and m.idx = "&frectidx&""
		end if
		if FRectYYYYMM <> "" then
			addSQL = addSQL & " and m.yyyymm = '" & FRectYYYYMM & "'"
		end if

		sqlStr = "select top 1"
		sqlStr = sqlStr + " m.yyyymm, m.idx, m.account_gubun, m.BIZSECTION_CD, m.buy_date, m.buy_sum, m.buy_cost, m.state, m.out_date"
		sqlStr = sqlStr + " , m.month_down_value, m.month_remain_value, m.regdate, s.BIZSECTION_NM"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_monthly m"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_TMS_BA_BIZSECTION s"
		sqlStr = sqlStr + " 	on m.BIZSECTION_CD = s.BIZSECTION_CD "
		sqlStr = sqlStr + " where 1=1 " & addSQL

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        ftotalcount = rsget.RecordCount
        set FOneItem = new CEquipmentMonthlyItem
        if Not rsget.Eof then
			FOneItem.fyyyymm = rsget("yyyymm")
			FOneItem.fidx = rsget("idx")
			FOneItem.FaccountGubun = rsget("account_gubun")
			FOneItem.FBIZSECTION_CD = rsget("BIZSECTION_CD")
			FOneItem.FBIZSECTION_NM = db2html(rsget("BIZSECTION_NM"))
        end if
        rsget.Close
    end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	end sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

class cequipmentcodeoneitem
	public fgubuntype
	public fgubuncd
	public ftypename
	public fgubunname
	public fisusing
	public forderno
	public fidx

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end class

class cequipmentcode
	public FOneItem
	public FItemList()
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public frectgubuntype
	public frectgubuncd
	public frectidx

	'//common/equipment/popmanagecode.asp
	public Sub getequipmentcodedetail()
		dim sqlStr, i , sqlsearch

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "&frectidx&""
		end if

		sqlStr = "select top 1"
		sqlStr = sqlStr + " idx, gubuntype, gubuncd, typename, gubunname, isusing, orderno"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_gubun"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by orderno asc"

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount =  rsget.RecordCount
		set FOneItem = new cequipmentcodeoneitem

		i=0
		if not rsget.EOF  then
			FOneItem.fidx = rsget("idx")
			FOneItem.fgubuntype = rsget("gubuntype")
			FOneItem.fgubuncd = db2html(rsget("gubuncd"))
			FOneItem.ftypename = rsget("typename")
			FOneItem.fgubunname = db2html(rsget("gubunname"))
			FOneItem.fisusing = rsget("isusing")
			FOneItem.forderno = rsget("orderno")
		end if
		rsget.Close
	end Sub

	'//common/equipment/popmanagecode.asp
	public Sub getequipmentcodelist()
		dim sqlStr, i , sqlsearch

		if frectgubuntype <> "" then
			sqlsearch = sqlsearch & " and gubuntype = '"&frectgubuntype&"'"
		end if

		'// 레코드 들의 수를 페이징 하기위해서 쿼리
		sqlStr = " select count(*) as cnt"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_gubun"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit Sub

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " idx ,gubuntype, gubuncd, typename, gubunname, isusing, orderno"
		sqlStr = sqlStr + " from db_partner.dbo.tbl_equipment_gubun"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by gubuntype desc ,orderno asc"

		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cequipmentcodeoneitem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fgubuntype = rsget("gubuntype")
				FItemList(i).fgubuncd = db2html(rsget("gubuncd"))
				FItemList(i).ftypename = rsget("typename")
				FItemList(i).fgubunname = db2html(rsget("gubunname"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).forderno = rsget("orderno")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	end sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class

function drawequipmentCodeType(selectboxname, selectid ,chplg)
%>
	<select class="select" name="<%= selectboxname %>" <%= chplg %>>
		<option name='' value='' <% if selectid = "" then response.write " selected" %>>선택</option>
		<option name='10' value='10' <% if selectid = "10" then response.write " selected" %>>물품구분</option>
		<option name='30' value='30' <% if selectid = "30" then response.write " selected" %>>장소구분</option>
		<option name='50' value='50' <% if selectid = "50" then response.write " selected" %>>상태구분</option>
	</select>
<%
End function

'// 계정과목
function drawEquipmentAccountCode(selectboxname, selectid ,chplg)
%>
	<select class="select" name="<%= selectboxname %>" <%= chplg %>>
		<option name='' value='' <% if selectid = "" then response.write " selected" %>>선택</option>
		<option name='21200' value='21200' <% if selectid = "21200" then response.write " selected" %>>비품(자산)</option>
		<option name='24000' value='24000' <% if selectid = "24000" then response.write " selected" %>>소프트웨어</option>
		<option name='21900' value='21900' <% if selectid = "21900" then response.write " selected" %>>시설장치</option>
		<option name='23300' value='23300' <% if selectid = "23300" then response.write " selected" %>>상표권</option>
		<option name='23500' value='23500' <% if selectid = "23500" then response.write " selected" %>>디자인권</option>
		<option name='' value=''>-----------</option>
		<option name='81950' value='81950' <% if selectid = "81950" then response.write " selected" %>>임대품</option>
		<option name='83090' value='83090' <% if selectid = "83090" then response.write " selected" %>>소모품</option>
	</select>
<%
End Function

'// 정보자산구분
function drawInfoEquipmentGubun(selectboxname, selectid ,chplg)
%>
	<select class="select" name="<%= selectboxname %>" <%= chplg %>>
		<option name='-1' value='-1' <% if selectid = "-1" then response.write " selected" %>>선택</option>
		<option name="10" value="10" <% if selectid = "10" then response.write " selected" %>>서버</option>
		<option name="20" value="20" <% if selectid = "20" then response.write " selected" %>>네트워크</option>
		<option name="30" value="30" <% if selectid = "30" then response.write " selected" %>>정보보호시스템</option>
		<!-- option name="40" value="40" <% if selectid = "40" then response.write " selected" %>>DBMS</option -->
		<!-- option name="50" value="50" <% if selectid = "50" then response.write " selected" %>>WAS</option -->
		<!-- option name="60" value="60" <% if selectid = "60" then response.write " selected" %>>정보자산</option -->
		<!-- option name="70" value="70" <% if selectid = "70" then response.write " selected" %>>서비스자산</option -->
		<option name="80" value="80" <% if selectid = "80" then response.write " selected" %>>PC</option>
		<option name="85" value="85" <% if selectid = "85" then response.write " selected" %>>모바일장비</option>
		<option name="90" value="90" <% if selectid = "90" then response.write " selected" %>>소프트웨어</option>
		<!-- option name="100" value="100" <% if selectid = "100" then response.write " selected" %>>문서</option -->
		<option name="110" value="110" <% if selectid = "110" then response.write " selected" %>>물리적자산</option>
	</select>
<%
End Function

function GetInfoGubunCodeName(info_gubun)
     select case info_gubun
        case "10"
			GetInfoGubunCodeName = "서버"
		case "20"
			GetInfoGubunCodeName = "네트워크"
		case "30"
			GetInfoGubunCodeName = "정보보호시스템"
		case "40"
			GetInfoGubunCodeName = "DBMS"
		case "50"
			GetInfoGubunCodeName = "WAS"
		case "60"
			GetInfoGubunCodeName = "정보자산"
		case "70"
			GetInfoGubunCodeName = "서비스자산"
		case "80"
			GetInfoGubunCodeName = "PC"
		case "85"
			GetInfoGubunCodeName = "모바일장비"
		case "90"
			GetInfoGubunCodeName = "소프트웨어"
		case "100"
			GetInfoGubunCodeName = "문서"
		case "110"
			GetInfoGubunCodeName = "물리적자산"
		case else
			GetInfoGubunCodeName = info_gubun
	end select
end function

function drawInfoImportance(selectboxname, selectid ,chplg)
%>
	<select class="select" name="<%= selectboxname %>" <%= chplg %>>
		<option name='' value='' <% if selectid = "" then response.write " selected" %>>선택</option>
		<option name="1" value="1" <% if selectid = "1" then response.write " selected" %>>L</Option>
		<option name="2" value="2" <% if selectid = "2" then response.write " selected" %>>M</Option>
		<option name="3" value="3" <% if selectid = "3" then response.write " selected" %>>H</option>
	</select>
<%
End function

function GetEquipmentAccountCodeName(selectid)
	select case selectid
		case "21200"
			GetEquipmentAccountCodeName = "비품(자산)"
		case "24000"
			GetEquipmentAccountCodeName = "소프트웨어"
		case "21900"
			GetEquipmentAccountCodeName = "시설장치"
		case "23300"
			GetEquipmentAccountCodeName = "상표권"
		case "23500"
			GetEquipmentAccountCodeName = "디자인권"
		case "81950"
			GetEquipmentAccountCodeName = "임대품"
		case "83090"
			GetEquipmentAccountCodeName = "소모품"
		case else
			GetEquipmentAccountCodeName = selectid
	end select
End function

function drawequipmentCodeType2(selectboxname, selectid ,chplg)
%>
	<select name="<%= selectboxname %>" <%= chplg %>>
		<option name='' value='' <% if selectid = "" then response.write " selected" %>>전체</option>
		<option name='10' value='10' <% if selectid = "10" then response.write " selected" %>>시스템장비</option>
		<option name='11' value='11' <% if selectid = "11" then response.write " selected" %>>오프라인장비</option>
	</select>
<%
End function

function getequipmentCodeType(gubuntype)
	if gubuntype = "" then exit function

	if gubuntype = "10" then
		getequipmentCodeType = "물품구분"
	elseif gubuntype = "30" then
		getequipmentCodeType = "장소구분"
	elseif gubuntype = "50" then
		getequipmentCodeType = "상태구분"
	end if
End function

Sub drawpartneruser(byval selectBoxName, selectedId ,chplg)
   dim tmp_str,sqlStr ,tmp_substr

	sqlStr = "select"
	sqlStr = sqlStr & " pi.part_name, t.empno  , t.username ,t.userid ,t.statediv"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr & " 	on t.userid = p.id"
	sqlStr = sqlStr & " 	and t.isUsing = 1"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partInfo pi"
	sqlStr = sqlStr & " 	on t.part_sn = pi.part_sn"
	sqlStr = sqlStr & " 	and pi.part_isdel = 'N'"
	sqlStr = sqlStr & " order by t.statediv desc ,t.part_sn desc, t.posit_sn asc ,t.username asc"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	%>
	<select class='select' name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<%

		if not rsget.EOF then
			rsget.Movefirst

			do until rsget.EOF

				tmp_substr = ""

				if selectedId <> "" then
					if selectedId = rsget("userid") then
						tmp_str = " selected"
					end if
				end if

				tmp_substr = tmp_substr + db2html(rsget("part_name")) + "-"
				tmp_substr = tmp_substr + db2html(rsget("username"))

				if rsget("userid") <> "" then tmp_substr = tmp_substr + " (" + rsget("userid") + ")"

				if rsget("statediv") <> "Y" then tmp_substr = tmp_substr + " (퇴사)"

				response.write("<option value='" + rsget("userid") + "' "&tmp_str&">" + tmp_substr + "</option>")
				tmp_str = ""
				rsget.MoveNext
			loop
		end if
	rsget.close
	response.write("</select>")
end Sub

'셀렉트 옵션 생성 함수(장비구분, 사용구분)
Sub DrawEquipMentGubun(gubuntype,selectBoxName,selectedId,chplg)
   dim tmp_str,query1, qyery2

   query1 = " select gubuncd,gubunname"
   query1 = query1 + " from [db_partner].[dbo].tbl_equipment_gubun"
   query1 = query1 + " where isusing='Y' and gubuntype='" + gubuntype + "'"
   query1 = query1 + " order by orderno asc ,idx desc"

   'response.write query1 & "<Br>"
   rsget.Open query1,dbget,1

	response.write "<select class='select' name='" & selectBoxName & "' "&chplg&">"
	response.write "<option value=''"
		if selectedId="" then
			response.write " selected"
		end if
	response.write ">선택</option>"

   if not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(trim(rsget("gubuncd"))) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&trim(rsget("gubuncd"))&"' "&tmp_str&">" + db2html(rsget("gubunname")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close

   response.write("</select>")
End Sub

'셀렉트 옵션 생성 함수(장비구분, 사용구분)
function getEquipMentGubun(gubuntype,selectedId)
   dim tmp_str,query1, qyery2

	 '옵션 내용 DB에서 가져오기
   query1 = " select gubuncd,gubunname"
   query1 = query1 + " from [db_partner].[dbo].tbl_equipment_gubun"
   query1 = query1 + " where isusing='Y' and gubuntype='" + gubuntype + "'"
   query1 = query1 + " order by orderno asc ,idx desc"

   'response.write query1 & "<Br>"
   rsget.Open query1,dbget,1

   if not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(trim(rsget("gubuncd"))) then
               tmp_str = db2html(rsget("gubunname"))
           end if

           rsget.MoveNext
       loop
   end if
   rsget.close

   getEquipMentGubun = tmp_str
End function

function textboxEquipMentGubun(gubuntype,selectBoxName,selectedId,chplg,scriptchplg)
	Dim strSql, arrList, intLoop , tmpgubuncd ,tmpgubunname

	strSql = " select top 1 gubuncd, gubunname"
	strSql = strSql & " from [db_partner].[dbo].tbl_equipment_gubun"
	strSql = strSql & " where isusing='Y' and gubuntype='" & gubuntype & "'"
	strSql = strSql & " and gubuncd = '"&selectedId&"'"

	'response.write strSql &"<br>"
	rsget.Open strSql,dbget

	IF not rsget.eof THEN
		arrList = rsget.getRows()
	End IF

	rsget.close

	IF isArray(arrList) THEN
		tmpgubuncd = arrList(0,0)
		tmpgubunname = arrList(1,0)
	end if
%>
	<input type="text" name="gubunname" value="<%= tmpgubunname %>" readonly size="20" maxlength=20 class="text" <%= chplg %>>
	<input type="button" class="button" value="선택" onClick="EquipMentGubun()">
	<input type="hidden" name="<%= selectBoxName %>" value="<%= tmpgubuncd %>">

	<script type='text/javascript'>
		function EquipMentGubun(){

			<%= scriptchplg %>

			var gubuncd = document.getElementsByName("<%= selectBoxName %>")[0].value;
			var gubunname = document.getElementsByName("gubunname")[0].value;

			var EquipMentGubun = window.open('/common/equipment/PopequipmentList.asp?gubuncd='+gubuncd+'&gubuntype=<%= gubuntype %>&boxname=<%=selectBoxName%>&gubunname='+gubunname,'EquipMentGubun','width=570,height=570,scrollbars=yes');
			EquipMentGubun.focus();
		}
	</script>
<%
end function

function textboxEquipmentGubunNew(gubuntype, gubunCdName, gubunNmName, selectedId, chplg, scriptchplg)
	Dim strSql, arrList, intLoop , tmpgubuncd ,tmpgubunname

	strSql = " select top 1 gubuncd, gubunname"
	strSql = strSql & " from [db_partner].[dbo].tbl_equipment_gubun"
	strSql = strSql & " where isusing='Y' and gubuntype='" & gubuntype & "'"
	strSql = strSql & " and gubuncd = '"&selectedId&"'"

	'response.write strSql &"<br>"
	rsget.Open strSql,dbget

	IF not rsget.eof THEN
		arrList = rsget.getRows()
	End IF

	rsget.close

	IF isArray(arrList) THEN
		tmpgubuncd = arrList(0,0)
		tmpgubunname = arrList(1,0)
	end if
%>
	<input type="text" name="<%= gubunNmName %>" value="<%= tmpgubunname %>" readonly size="20" maxlength=20 class="text" <%= chplg %>>
	<input type="button" class="button" value="선택" onClick="EquipMentGubun('<%= gubuntype %>', '<%= gubunCdName %>', '<%= gubunNmName %>')">
	<input type="hidden" name="<%= gubunCdName %>" value="<%= tmpgubuncd %>">

	<script type='text/javascript'>
		function EquipMentGubun(gubuntype, gubunCdName, gubunNmName) {

			<%= scriptchplg %>

			var EquipMentGubun = window.open('/common/equipment/PopequipmentList.asp?gubuntype=' + gubuntype + '&boxname=' + gubunCdName + '&gubunname=' + gubunNmName,'EquipMentGubun','width=570,height=570,scrollbars=yes');
			EquipMentGubun.focus();
		}
	</script>
<%
end function

function makeEquipCode(byval idx, byval equip_gubun, byval buy_date)
	if buy_date="" then buy_date="00000000"

	makeEquipCode = equip_gubun + "-" + replace(Left(CStr(buy_date),10),"-","") + "-" + format00(6,idx)
end function

function makeEquipCodeNew(byval idx, byval equip_gubun, byval buy_date, byval accountGubun)
	if buy_date="" then buy_date="00000000"

	makeEquipCodeNew = equip_gubun + "-" + Right(replace(Left(CStr(buy_date),10),"-",""), 6) & "-" & Left(accountGubun, 3) & "-" & format00(5,idx)
end function

function getCIALevelName(ciaVal)
    getCIALevelName=""
    if isNULL(ciaVal) then Exit Function

    if (ciaVal="1") then getCIALevelName="L"
    if (ciaVal="2") then getCIALevelName="M"
    if (ciaVal="3") then getCIALevelName="H"

end function
%>
