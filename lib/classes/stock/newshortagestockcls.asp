<%
'###########################################################
' Description : 물류 입출고 클래스
' History : 이상구 생성
'			2016.03.20 한용민 수정
'###########################################################

Class CShortageStockItem
	public Fjungsan_gubun
	public FOffimgMain
	public FOffimgList
	public FOffimgSmall
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public FitemOptionName
	public Fisusing
	public Foptionusing
	public Fsellyn
	public Fdanjongyn
    public Foptdanjongyn
	public Flimityn
	public Foptlimityn
	public FLimitNo
	public FLimitSold
	public Foptionlimitno
	public Foptionlimitsold
	public FOptionCnt
    public FreipgoMayDate					'재입고 예정일
	public FSellcash
	public FBuycash
	public FMwDiv
	public Fmakerid
	public FimageSmall
	public Fdeliverytype
	public FOffLineDefaultMargin
	public FOffLineDefaultSuplyMargin
	public Fipgono
	public Freipgono
	public Ftotipgono
	public Foffchulgono
	public Foffrechulgono
	public Fetcchulgono
	public Fetcrechulgono
	public Ftotchulgono
	public Fsellno
	public Fresellno
	public Ftotsellno
	public Ferrcsno
	public Ferrbaditemno
	public Ferrrealcheckno
	public Ferretcno
	public Ftoterrno
	public Ftotsysstock
	public Favailsysstock
	public Frealstock
	public Fsell7days
	public Foffchulgo7days
	public Fipkumdiv5
	public Fipkumdiv4
	public Fipkumdiv2
	public Foffconfirmno
	public Foffjupno
	public Frequireno
	public Fshortageno
	public Fpreordernofix
	public Fpreorderno
	public Foffsellno
	public Fmaxsellday
	public Fregdate
	public Flastupdate
	public FDayForSellCount
	public FDayForSafeStock
	public FDayForLeadTime
	public FDayForMaxStock
	public FOffMwMargin

	public Fsailyn
	public Forgprice
	public FsaleStr

	public FsellSTDateStr
	public FsellSTDate
	public FthreeMonthSellNo

    public FlastIpgoDate
	public FAGVStock
	public FrackcodeByOption
	public FsubRackcodeByOption

	public function IsOffContractExist()
		IsOffContractExist = Not (FOffMwMargin = "")
	end function

	public function GetOffContractMWDiv()
		dim tmpArr

		GetOffContractMWDiv = ""

		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			GetOffContractMWDiv = tmpArr(0)
		end if
	end function

	public function GetOffContractMargin()
		dim tmpArr

		GetOffContractMargin = ""

		'// M_45_0
		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			GetOffContractMargin = tmpArr(1)
		end if
	end function

	public function GetOffContractBuycash()
		dim tmpArr

		GetOffContractBuycash = FBuycash

		'// M_45_0
		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			if tmpArr(1) <> 0 and tmpArr(2) = 0 then
				'// 마진적용
				GetOffContractBuycash = CLng(Fsellcash * (100 - tmpArr(1)) / 100)
			elseif tmpArr(2) <> 0 then
				'//상품매입가
				GetOffContractBuycash = tmpArr(2)
			end if
		end if
	end function

	public function GetOffContractCenterMW()
		dim tmpArr

		GetOffContractCenterMW = "U"

		'// M_45_0
		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			GetOffContractCenterMW = tmpArr(0)
		end if
	end function

	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	public function getMwDivColor()
		if FmwDiv="M" then
			getMwDivColor = "#CC2222"
		elseif FmwDiv="W" then
			getMwDivColor = "#2222CC"
		elseif FmwDiv="U" then
			getMwDivColor = "#000000"
		end if
	end function

	public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public Function IsSoldOut()
		IsSoldOut = (FSellYn="N") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function

	public Function GetMayNo()
		GetMayNo = Favailsysstock
	end function

	public Function GetUsingStr()
		if FIsUsing="N" then
			GetUsingStr = "<font color=#00FF00>x</font>"
		end if
	end function

	public Function GetSellStr()
		if FSellYn="N" then
			GetSellStr = "<font color=#FF0000>x</font>"
		end if
	end function

	public Function GetLimitStr()
		if (Fitemoption="0000") then
			if FLimityn="Y" then
				if FLimitNo-FLimitSold<1 then
					GetLimitStr = "0"
				else
					GetLimitStr = CStr(FLimitNo-FLimitSold)
				end if
			end if
		else
			if FOptLimityn="Y" then
				if Foptionlimitno-Foptionlimitsold<1 then
					GetLimitStr = "0"
				else
					GetLimitStr = CStr(Foptionlimitno-Foptionlimitsold)
				end if
			end if
		end if
	end function

	public Function GetBigoStr()
		dim reStr
		if Fdanjongyn="Y" then
			reStr = reStr + " 단종"
		end if

		if FIsUsing="N" then
			reStr = reStr + " 사용x"
		end if

		if FSellYn="N" then
			reStr = reStr + " 판매x"
		end if

		if FLimityn="Y" then
			reStr = reStr + " 한정" + CStr(GetLimitEa()) + "개"
		end if

		GetBigoStr = reStr
	end function

	public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public function GetdeliverytypeName()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			GetdeliverytypeName = "업배"
		else
			GetdeliverytypeName = "텐배"
		end if
	end function

	public function IsInvalidOption()
		IsInvalidOption = (Fitemoption<>"0000") and (FitemoptionName="")
	end function

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fitemid) & Fitemoption
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	public function GetBarCodeBoldStr()
		GetBarCodeBoldStr = Fitemgubun & "-" & Format00(6,Fitemid) & "-" & Fitemoption
		if (Fitemid >= 1000000) then
    		GetBarCodeBoldStr = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	''N일치 부족수량
    public function GetNdayShortageNo(nday)
		GetNdayShortageNo = Fshortageno + CLng(Frequireno*(nday-7)/7)
	end function

    ''출고이전 필요수량(접수,결제완료..)
    public function GetReqNotChulgoNo()
		GetReqNotChulgoNo = Fipkumdiv5 + Foffconfirmno + Fipkumdiv4 + Fipkumdiv2 + Foffjupno
	end Function

    ''예상재고
	public function GetMaystock()
		GetMaystock = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2 + Foffjupno
	end function

    ''재고파악재고
	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

    ''금일 상품준비수량
	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function

    public function GetAGVShortageNo()
		GetAGVShortageNo = FAGVStock + Fipkumdiv5 + Fipkumdiv4 + Foffconfirmno
	end function

	public function getLimitNo()
		getLimitNo = 0
		if (Flimityn = "Y") then
			getLimitNo = Flimitno-Flimitsold
		end if

		if getLimitNo < 1 then getLimitNo = 0
	end function

	public function getOptionLimitNo()
		getOptionLimitNo = 0
		if ((Flimityn = "Y") and (Foptionusing = "Y")) then
			getOptionLimitNo = Foptionlimitno - Foptionlimitsold
		end if

		if getOptionLimitNo < 1 then getOptionLimitNo = 0

		if (FOptionCnt < 1) then
		    getOptionLimitNo = getLimitNo
		end if
	end function

	Private Sub Class_Initialize()
		Fipgono			= 0
		Freipgono		= 0
		Ftotipgono		= 0
		Foffchulgono	= 0
		Foffrechulgono	= 0
		Fetcchulgono	= 0
		Fetcrechulgono	= 0
		Ftotchulgono	= 0
		Fsellno			= 0
		Fresellno		= 0
		Ftotsellno		= 0
		Ferrcsno		= 0
		Ferrbaditemno	= 0
		Ferrrealcheckno	= 0
		Ferretcno		= 0
		Ftoterrno		= 0
		Ftotsysstock	= 0
		Favailsysstock	= 0
		Frealstock		= 0
		Fsell7days		= 0
		Foffchulgo7days	= 0
		Fipkumdiv5		= 0
		Fipkumdiv4		= 0
		Fipkumdiv2		= 0
		Foffconfirmno	= 0
		Foffjupno		= 0
		Frequireno		= 0
		Fshortageno		= 0
		Fpreordernofix	= 0
		Fpreorderno		= 0
		Fmaxsellday		= 0
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CStockBaseDayItem
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public FitemOptionName
	public Fisusing
	public Foptionusing
	public Fsellyn
	public Fdanjongyn
	public Flimityn
	public Foptlimityn
	public FLimitNo
	public FLimitSold
	public Foptionlimitno
	public Foptionlimitsold
	public FOptionCnt
    public FreipgoMayDate					'재입고 예정일
	public FSellcash
	public FBuycash
	public FMwDiv
	public Fmakerid
	public FimageSmall
	public Fdeliverytype
	public FDayForSellCount
	public FDayForSafeStock
	public FDayForLeadTime
	public FDayForMaxStock

	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	public function getMwDivColor()
		if FmwDiv="M" then
			getMwDivColor = "#CC2222"
		elseif FmwDiv="W" then
			getMwDivColor = "#2222CC"
		elseif FmwDiv="U" then
			getMwDivColor = "#000000"
		end if
	end function

	public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public Function IsSoldOut()
		IsSoldOut = (FSellYn="N") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function

	public Function GetUsingStr()
		if FIsUsing="N" then
			GetUsingStr = "<font color=#00FF00>x</font>"
		end if
	end function

	public Function GetSellStr()
		if FSellYn="N" then
			GetSellStr = "<font color=#FF0000>x</font>"
		end if
	end function

	public Function GetLimitStr()
		if (Fitemoption="0000") then
			if FLimityn="Y" then
				if FLimitNo-FLimitSold<1 then
					GetLimitStr = "0"
				else
					GetLimitStr = CStr(FLimitNo-FLimitSold)
				end if
			end if
		else
			if FOptLimityn="Y" then
				if Foptionlimitno-Foptionlimitsold<1 then
					GetLimitStr = "0"
				else
					GetLimitStr = CStr(Foptionlimitno-Foptionlimitsold)
				end if
			end if
		end if
	end function

	public Function GetBigoStr()
		dim reStr
		if IsNull(FDayForSellCount) then
			reStr = reStr + "기본값"
		end if

		GetBigoStr = reStr
	end function

	public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public function GetdeliverytypeName()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			GetdeliverytypeName = "업배"
		else
			GetdeliverytypeName = "텐배"
		end if
	end function

	public function IsInvalidOption()
		IsInvalidOption = (Fitemoption<>"0000") and (FitemoptionName="")
	end function

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fitemid) & Fitemoption
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	public function GetBarCodeBoldStr()
		GetBarCodeBoldStr = Fitemgubun & "-" & Format00(6,Fitemid) & "-" & Fitemoption
		if (Fitemid >= 1000000) then
    		GetBarCodeBoldStr = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	public function getLimitNo()
		getLimitNo = 0
		if (Flimityn = "Y") then
			getLimitNo = Flimitno-Flimitsold
		end if

		if getLimitNo < 1 then getLimitNo = 0
	end function

	public function getOptionLimitNo()
		getOptionLimitNo = 0
		if ((Flimityn = "Y") and (Foptionusing = "Y")) then
			getOptionLimitNo = Foptionlimitno - Foptionlimitsold
		end if

		if getOptionLimitNo < 1 then getOptionLimitNo = 0

		if (FOptionCnt < 1) then
		    getOptionLimitNo = getLimitNo
		end if
	end function

	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CShortageStock
	public FItemList()
	public FOneItem
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FSPageNo
	public FEPageNo

	public FRectOnlySell
	public FRectOnlyUsingItem
	public FRectOnlyUsingItemOption
	public FRectOnlyNotUpcheBeasong
	public FRectOnlynotDealItem
	public FRectShortage7days
	public FRectShortage14days
	public FRectShortageRealStock
	public FRectOnlyNotDanjong
	public FRectOnlyNotTempDanjong
	public FRectOnlyNotMDDanjong
	public FRectIncludePreOrder
	public FRectSkipLimitSoldOut
	public FRectOnlyStockMinus
	Public FRectMinusStockGubun
	public FRectPurchaseType
	public FRectOnlyNotInputDay
	public FRectMakerid
	public FRectItemGubun
	public FRectItemGubunExclude
	public FRectItemId
	public FRectItemOption
	public FRectItemName
	public FRectMWDiv
	public FRectExcMkr
	public FRectOnlyOn
	public FRectCenterMWDiv
	public FRectonlyrealstockexists
	public FRectCD1
	public FRectCD2
	public FRectCD3
	public FRectDayForSellCount
	public FRectAGVCheck
    Public FRectOnlyRealUp
    Public FRectOrderBy
    public FTotalMakeridCount
    public FTotalPieceCount

	'온라인상품
	'//admin/newstorage/popjumunitemNew.asp
	public Sub GetShortageItemListOnline
		dim i,sqlStr, withstr

        withstr = ""
        if (FRectMakerid="") then  ''쿼리방식 수정 2016/04/21
            withstr = withstr + " ;WITH TBL_STOCK_TARGET as ("
            withstr = withstr + " select s.*"
            withstr = withstr + " from  [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
            withstr = withstr + " where 1=1"

            if FRectonlyrealstockexists <> "" then
    			withstr = withstr + " and IsNull(s.realstock,0)>0"
    		end if
    		'======================================================================
    		''7일 초과부족수량
    		if FRectShortage7days <> "" then
    		    withstr = withstr + " and s.shortageno<0"
    		end if
    		'======================================================================
    		''14일 초과부족 수량
    		'======================================================================
    		''7일간 필요수량			C = (A+B)/P*D
    		''7일후 초과(부족)수량		(S-C-R)
    		''14일후 초과(부족)수량		(S-(C*2)-R) = (S-C-R) - C
    		'======================================================================
            if FRectShortage14days <> "" then
    		    withstr = withstr + " and (s.shortageno + s.requireno)<0"
    		end if
            ''현재고 N개 이하
            if FRectShortageRealStock <> "" then
    		    withstr = withstr + " and s.realstock <= 5"
    		end if
    		'기주문포함부족상품
    		if FRectIncludePreOrder <> "" then
    			withstr = withstr + " and (s.shortageno + s.preordernofix) < 0 "
    		end if
    		'실사재고마이너스상품
    		if FRectOnlyStockMinus <> "" then
    			withstr = withstr + " and s.realstock < 0 "
    		end if
    		If (FRectMinusStockGubun <> "") Then
    			Select Case FRectMinusStockGubun
    				Case "real"
    					''실사재고 마이너스
    					withstr = withstr + " and s.realstock < 0 "
    				Case "check"
    					''재고파악재고 마이너스
    					withstr = withstr + " and (s.realstock + s.ipkumdiv5 + s.offconfirmno) < 0 "
    				Case "may"
    					''예상재고 마이너스
    					withstr = withstr + " and ((s.realstock + s.ipkumdiv5 + s.offconfirmno) + s.ipkumdiv4 + s.ipkumdiv2 + s.offjupno) < 0 "
    				Case Else
    					''
    			End Select
    		End If
            withstr = withstr + " )"
        end if

		sqlStr = withstr & " select count(i.itemid) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
        sqlStr = sqlStr + " ,count(distinct i.makerid) as makeridCnt, sum((s.shortageno + s.preordernofix + s.sell7days)*-1) as pieceCnt "
		if FRectAGVCheck = "Y" then
		sqlStr = sqlStr + GetFromWhereOnLine(FRectMakerid="",true)
		else
		sqlStr = sqlStr + GetFromWhereOnLine(FRectMakerid="",false)
		end if

		'response.write sqlStr & "<br>"
		'response.end
		''rsget.Open sqlStr,dbget,1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
            FTotalMakeridCount = rsget("makeridCnt")
            FTotalPieceCount = rsget("pieceCnt")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			'exit sub
		end if

		sqlStr = withstr & " select top " + CStr(FPageSize*FCurrPage) + "  "
		sqlStr = sqlstr + " i.itemid, IsNull(o.itemoption, '0000') as itemoption, i.makerid, "
		sqlStr = sqlStr + " IsNull(s.ipgono,0) as ipgono, IsNULL(s.reipgono,0) as reipgono, "
		sqlStr = sqlStr + " IsNull(s.totipgono,0) as totipgono, IsNull(s.offchulgono,0) as offchulgono, "
        sqlStr = sqlStr + " IsNull(s.offrechulgono,0) as offrechulgono, IsNull(s.etcchulgono,0) as etcchulgono, "
        sqlStr = sqlStr + " IsNull(s.etcrechulgono,0) as etcrechulgono, IsNull(s.totchulgono,0) as totchulgono, "
        sqlStr = sqlStr + " IsNull(s.sellno,0) as sellno, IsNull(s.resellno,0) as resellno, "
   		sqlStr = sqlStr + " IsNull(s.totsellno,0) as totsellno, IsNull(s.errcsno,0) as errcsno, "
        sqlStr = sqlStr + " IsNull(s.errbaditemno,0) as errbaditemno, IsNull(s.errrealcheckno,0) as errrealcheckno, "
        sqlStr = sqlStr + " IsNull(s.erretcno,0) as erretcno, IsNull(s.toterrno,0) as toterrno, "
    	sqlStr = sqlStr + " IsNull(s.totsysstock,0) as totsysstock, IsNull(s.availsysstock,0) as availsysstock, "
        sqlStr = sqlStr + " IsNull(s.realstock,0) as realstock, IsNull(s.sell7days,0) as sell7days, "
     	sqlStr = sqlStr + " IsNull(s.offchulgo7days,0) as offchulgo7days, IsNull(s.ipkumdiv5,0) as ipkumdiv5, "
        sqlStr = sqlStr + " IsNull(s.ipkumdiv4,0) as ipkumdiv4, IsNull(s.ipkumdiv2,0) as ipkumdiv2, "
        sqlStr = sqlStr + " IsNull(s.offconfirmno,0) as offconfirmno, IsNull(s.offjupno,0) as offjupno, "
  		sqlStr = sqlStr + " IsNull(s.requireno,0) as requireno, IsNull(s.shortageno,0) as  shortageno, "
        sqlStr = sqlStr + " IsNull(s.preorderno,0) as preorderno, IsNull(s.preordernofix,0) as preordernofix, IsNull(s.offsellno,0) as  offsellno, "
        sqlStr = sqlStr + " IsNull(s.maxsellday,1) as maxsellday, s.regdate, s.lastupdate"
		sqlStr = sqlstr + " ,i.smallimage, i.itemname, i.sellyn, i.isusing, i.mwdiv, i.limityn, i.limitno, i.limitsold ,i.deliverytype, i.danjongyn"
		sqlStr = sqlstr + " , i.sellcash + IsNULL(o.optaddprice,0) as sellcash, i.orgprice + IsNULL(o.optaddprice,0) as orgprice, i.buycash + IsNULL(o.optaddbuyprice,0) as buycash "
		sqlStr = sqlstr + " ,IsNULL(o.optionname,'') as itemoptionname , IsNULL(o.isusing,'Y') as optionusing, o.optlimityn, o.optlimitno, o.optlimitsold, i.optioncnt, IsNULL(o.optdanjongyn,'N') as optdanjongyn "
		sqlStr = sqlstr + " ,T.StockReipgoDate as reipgoMayDate "
		sqlStr = sqlstr + " , IsNull(T.DayForSellCount, 7) as DayForSellCount, IsNull(T.DayForSafeStock, 5) as DayForSafeStock, IsNull(T.DayForLeadTime, 2) as DayForLeadTime, IsNull(T.DayForMaxStock, 12) as DayForMaxStock, IsNull(s.requireMaxno,s.requireno*2) as requireMaxno "
		sqlStr = sqlstr + " ,(case when i.mwdiv = 'U' then [db_shop].[dbo].[uf_GetCenterMWDivMargin](i.itemid) else '' end) as offmwmargin "
		sqlStr = sqlstr + " , i.sailyn, i.orgprice + IsNULL(o.optaddprice,0) as orgprice, (case when i.sailyn = 'Y' then [db_item].[dbo].[uf_GetOnItemSaleDateStr](i.itemid) else '' end) as saleStr "
		sqlStr = sqlstr + "	, (case "
		sqlStr = sqlstr + " 			when i.sellSTDate is NULL then 'ERR' "
		sqlStr = sqlstr + " 			when (DateDiff(day, i.sellSTDate, getdate()) + 1) >= 100 then '99+' "
		sqlStr = sqlstr + " 			when (DateDiff(day, i.sellSTDate, getdate()) + 1) < 100 then convert(varchar, (DateDiff(day, i.sellSTDate, getdate()) + 1)) "
		sqlStr = sqlstr + " 			else '' end) as sellSTDateStr "
		sqlStr = sqlstr + " 	,i.sellSTDate "
		sqlStr = sqlstr + " 	, db_summary.dbo.uf_GetThreeMonthSellNo(s.itemgubun, s.itemid, s.itemoption) as threeMonthSellNo "
		sqlStr = sqlstr & " , g.jungsan_gubun, IsNull(T.rackcodeByOption, T0.rackcodeByOption) as rackcodeByOption, IsNull(T.subRackcodeByOption, T0.subRackcodeByOption) as subRackcodeByOption"
        if (FRectMakerid <> "") then
            sqlStr = sqlstr & " , a.lastIpgoDate "
        else
            sqlStr = sqlstr & " , '' as lastIpgoDate"
        end if
		if FRectAGVCheck = "Y" then
		sqlStr = sqlstr & " , isnull(agv.agvstock,0) as agvstock"
		sqlStr = sqlStr + GetFromWhereOnLine(FRectMakerid="",true)
		else
		sqlStr = sqlStr + GetFromWhereOnLine(FRectMakerid="",false)
		end if

        if (FRectOrderBy = "subrackcode") then
            sqlStr = sqlStr + " order by T.subRackcodeByOption, i.itemid desc, IsNull(o.itemoption, '0000') "
        else
            sqlStr = sqlStr + " order by i.makerid, i.itemid desc, IsNull(o.itemoption, '0000') "
        end if

        'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		''rsget.Open sqlStr,dbget,1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShortageStockItem

				FItemList(i).Fjungsan_gubun   = rsget("jungsan_gubun")
				FItemList(i).Fitemgubun     = "10"
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FIsUsing       = rsget("isusing")
				FItemList(i).FSellYn        = rsget("sellyn")
				FItemList(i).FLimityn       = rsget("limityn")
				FItemList(i).FLimitNo       = rsget("limitno")
				FItemList(i).FLimitSold     = rsget("limitsold")
				FItemList(i).FOptionCnt     = rsget("optioncnt")
				FItemList(i).FimageSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + rsget("smallimage")
				FItemList(i).Ftotsellno     = rsget("totsellno")*-1
				FItemList(i).Ftotipgono     = rsget("totipgono")
				FItemList(i).Foffchulgono	= rsget("offchulgono")
				FItemList(i).Foffrechulgono	= rsget("offrechulgono")
				FItemList(i).Fetcchulgono	= rsget("etcchulgono")
				FItemList(i).Fetcrechulgono	= rsget("etcrechulgono")
				FItemList(i).Ftotchulgono   = rsget("totchulgono")
				FItemList(i).Foptionusing   = rsget("optionusing")
				FItemList(i).Fdeliverytype  = rsget("deliverytype")
				FItemList(i).Flastupdate    = rsget("lastupdate")
				FItemList(i).FMakerID       = rsget("makerid")
				FItemList(i).FMwDiv         = rsget("mwdiv")
				FItemList(i).FSellcash      = rsget("sellcash")
				FItemList(i).Forgprice      = rsget("orgprice")				'// 소비자가
				FItemList(i).Ftotsysstock	= rsget("totsysstock")
				FItemList(i).Ferrbaditemno	= rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
				FItemList(i).Ftoterrno		= rsget("toterrno")
				FItemList(i).Favailsysstock = rsget("availsysstock")
				FItemList(i).Frealstock 	= rsget("realstock")
				FItemList(i).Fsell7days     = rsget("sell7days")*-1
				FItemList(i).Foffchulgo7days= rsget("offchulgo7days")*-1
				FItemList(i).Foffconfirmno  = rsget("offconfirmno")
				FItemList(i).Foffjupno      = rsget("offjupno")
				FItemList(i).Frequireno     = rsget("requireno")
				FItemList(i).Fshortageno    = rsget("shortageno")
				FItemList(i).Fpreordernofix = rsget("preordernofix")
				FItemList(i).Fbuycash		= rsget("buycash")
				FItemList(i).Fpreorderno    = rsget("preorderno")
				FItemList(i).FIpkumdiv5		= rsget("ipkumdiv5")
				FItemList(i).FIpkumdiv4		= rsget("ipkumdiv4")
				FItemList(i).FIpkumdiv2		= rsget("ipkumdiv2")
				FItemList(i).Foptlimityn		= rsget("optlimityn")
				FItemList(i).Foptionlimitno		= rsget("optlimitno")
				FItemList(i).Foptionlimitsold	= rsget("optlimitsold")
				FItemList(i).Fdanjongyn			= rsget("danjongyn")
                FItemList(i).Foptdanjongyn		= rsget("optdanjongyn")
				FItemList(i).FreipgoMayDate		= rsget("reipgoMayDate")
				FItemList(i).Fmaxsellday		= rsget("maxsellday")
				FItemList(i).FDayForSellCount	= rsget("DayForSellCount")
				FItemList(i).FDayForSafeStock	= rsget("DayForSafeStock")
				FItemList(i).FDayForLeadTime	= rsget("DayForLeadTime")
				FItemList(i).FDayForMaxStock	= rsget("DayForMaxStock")
				FItemList(i).FOffMwMargin		= rsget("offmwmargin")

				FItemList(i).Fsailyn		= rsget("sailyn")
				FItemList(i).Forgprice		= rsget("orgprice")
				FItemList(i).FsaleStr		= rsget("saleStr")

				FItemList(i).FsellSTDateStr		= rsget("sellSTDateStr")
				FItemList(i).FsellSTDate		= rsget("sellSTDate")
				FItemList(i).FthreeMonthSellNo	= rsget("threeMonthSellNo")

                FItemList(i).FlastIpgoDate	= rsget("lastIpgoDate")
				FItemList(i).FrackcodeByOption	= rsget("rackcodeByOption")
				FItemList(i).FsubRackcodeByOption	= rsget("subRackcodeByOption")
				if FRectAGVCheck = "Y" then FItemList(i).FAGVStock = rsget("agvstock")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'오프라인상품
	'/admin/newstorage/popjumunitemNew.asp
	public Sub GetShortageItemListOffline
		dim i,sqlStr

		sqlStr = " select count(i.shopitemid) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		if FRectAGVCheck = "Y" then
		sqlStr = sqlStr + GetFromWhereOffLine(true)
		else
		sqlStr = sqlStr + GetFromWhereOffLine(false)
		end if

		'response.write sqlStr
		''rsget.Open sqlStr,dbget,1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + "  "
		sqlStr = sqlStr + " i.itemgubun, i.shopitemid as itemid, i.itemoption, i.offimgmain, i.offimglist, i.offimgsmall"
		sqlStr = sqlStr + " , i.makerid, i.shopitemname as itemname, i.shopitemoptionname as itemoptionname,"
		sqlStr = sqlStr + " i.shopitemprice as sellcash, i.shopsuplycash as suplycash, i.isusing, i.isusing as optionusing, i.regdate, i.extbarcode,"
		sqlStr = sqlStr + " d.defaultmargin, d.defaultsuplymargin, d.chargediv,"
		sqlStr = sqlStr + " 'Y' as sellyn, 'N' as limityn, 0 as limitno, 0 as limitsold, 'N' as optlimityn, 0 as optlimitno, 0 as optlimitsold, i.centermwdiv as mwdiv, s.preordernofix, "
		sqlStr = sqlStr + " IsNull(s.ipgono,0) as ipgono, IsNULL(s.reipgono,0) as reipgono, "
		sqlStr = sqlStr + " IsNull(s.totipgono,0) as totipgono, IsNull(s.offchulgono,0) as offchulgono, "
        sqlStr = sqlStr + " IsNull(s.offrechulgono,0) as offrechulgono, IsNull(s.etcchulgono,0) as etcchulgono, "
        sqlStr = sqlStr + " IsNull(s.etcrechulgono,0) as etcrechulgono, IsNull(s.totchulgono,0) as totchulgono, "
        sqlStr = sqlStr + " IsNull(s.sellno,0) as sellno, IsNull(s.resellno,0) as resellno, "
   		sqlStr = sqlStr + " IsNull(s.totsellno,0) as totsellno, IsNull(s.errcsno,0) as errcsno, "
        sqlStr = sqlStr + " IsNull(s.errbaditemno,0) as errbaditemno, IsNull(s.errrealcheckno,0) as errrealcheckno, "
        sqlStr = sqlStr + " IsNull(s.erretcno,0) as erretcno, IsNull(s.toterrno,0) as toterrno, "
    	sqlStr = sqlStr + " IsNull(s.totsysstock,0) as totsysstock, IsNull(s.availsysstock,0) as availsysstock, "
        sqlStr = sqlStr + " IsNull(s.realstock,0) as realstock, IsNull(s.sell7days,0) as sell7days, "
     	sqlStr = sqlStr + " IsNull(s.offchulgo7days,0) as offchulgo7days, IsNull(s.ipkumdiv5,0) as ipkumdiv5, "
        sqlStr = sqlStr + " IsNull(s.ipkumdiv4,0) as ipkumdiv4, IsNull(s.ipkumdiv2,0) as ipkumdiv2, "
        sqlStr = sqlStr + " IsNull(s.offconfirmno,0) as offconfirmno, IsNull(s.offjupno,0) as offjupno, "
  		sqlStr = sqlStr + " IsNull(s.requireno,0) as requireno, IsNull(s.shortageno,0) as  shortageno, "
        sqlStr = sqlStr + " IsNull(s.preorderno,0) as preorderno, IsNull(s.offsellno,0) as  offsellno, "
        sqlStr = sqlStr + " IsNull(s.maxsellday,1) as maxsellday, s.regdate, s.lastupdate"
		sqlStr = sqlstr + " ,T.StockReipgoDate as reipgoMayDate "
		sqlStr = sqlstr + " , IsNull(T.DayForSellCount, 7) as DayForSellCount, IsNull(T.DayForSafeStock, 5) as DayForSafeStock, IsNull(T.DayForLeadTime, 2) as DayForLeadTime, IsNull(T.DayForMaxStock, 12) as DayForMaxStock, IsNull(s.requireMaxno,s.requireno*2) as requireMaxno "
		sqlStr = sqlstr & " , g.jungsan_gubun, T.rackcodeByOption, T.subRackcodeByOption"
		if FRectAGVCheck = "Y" then
		sqlStr = sqlstr & " , isnull(agv.agvstock,0) as agvstock"
		sqlStr = sqlStr + GetFromWhereOffLine(true)
		else
		sqlStr = sqlStr + GetFromWhereOffLine(false)
		end if
		sqlStr = sqlStr + " order by i.makerid, i.itemgubun desc, i.shopitemid desc, i.itemoption "

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		''rsget.Open sqlStr,dbget,1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShortageStockItem

				FItemList(i).Fjungsan_gubun   = rsget("jungsan_gubun")
				FItemList(i).Fitemgubun     = rsget("itemgubun")
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FIsUsing       = rsget("isusing")
				FItemList(i).FSellYn        = rsget("sellyn")
				FItemList(i).FLimityn       = rsget("limityn")
				FItemList(i).FLimitNo       = rsget("limitno")
				FItemList(i).FLimitSold     = rsget("limitsold")
				'FItemList(i).FOptionCnt     = rsget("optioncnt")
				'FItemList(i).FimageSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + rsget("smallimage")
				FItemList(i).Ftotsellno     = rsget("totsellno")*-1
				FItemList(i).Ftotipgono     = rsget("totipgono")
				FItemList(i).Foffchulgono	= rsget("offchulgono")
				FItemList(i).Foffrechulgono	= rsget("offrechulgono")
				FItemList(i).Fetcchulgono	= rsget("etcchulgono")
				FItemList(i).Fetcrechulgono	= rsget("etcrechulgono")
				FItemList(i).Ftotchulgono   = rsget("totchulgono")
				FItemList(i).Foptionusing   = rsget("optionusing")
				'FItemList(i).Fdeliverytype  = rsget("deliverytype")
				FItemList(i).Flastupdate    = rsget("lastupdate")
				FItemList(i).FMakerID       = rsget("makerid")
				FItemList(i).FMwDiv         = rsget("mwdiv")
				FItemList(i).FSellcash      = rsget("sellcash")
				FItemList(i).Forgprice      = rsget("sellcash")			'// 임시로
				'==============================================================
				'오프라인상품은 샵별로 마진이 달라지므로 매입가가 입력되지 않은 경우 일단 디폴트 마진을 적용해서 매입가를 산정하고
				'매입분은 그대로 정산하고, 위탁분은 샵별 마진정보로 출고하고 그 내역을 정산한다.
				FItemList(i).Fbuycash      				= rsget("suplycash")
				FItemList(i).FOffLineDefaultMargin      = rsget("defaultmargin")
				FItemList(i).FOffLineDefaultSuplyMargin = rsget("defaultsuplymargin")
				if (FItemList(i).Fbuycash = 0) and (FItemList(i).FOffLineDefaultMargin <> 0) then
					FItemList(i).Fbuycash = CLng(FItemList(i).Fsellcash * (100 - FItemList(i).FOffLineDefaultMargin) / 100)
				end if
				'==============================================================
				FItemList(i).Ftotsysstock	= rsget("totsysstock")
				FItemList(i).Ferrbaditemno		= rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
				FItemList(i).Ftoterrno		= rsget("toterrno")
				FItemList(i).Favailsysstock = rsget("availsysstock")
				FItemList(i).Frealstock = rsget("realstock")
				FItemList(i).Fsell7days     = rsget("sell7days")*-1
				FItemList(i).Foffchulgo7days= rsget("offchulgo7days")*-1
				FItemList(i).Foffconfirmno  = rsget("offconfirmno")
				FItemList(i).Foffjupno      = rsget("offjupno")
				FItemList(i).Frequireno     = rsget("requireno")*-1
				FItemList(i).Fshortageno    = rsget("shortageno")
				FItemList(i).Fpreordernofix    = rsget("preordernofix")
				FItemList(i).Fpreorderno      = rsget("preorderno")
				FItemList(i).FIpkumdiv5		= rsget("ipkumdiv5")
				FItemList(i).FIpkumdiv4		= rsget("ipkumdiv4")
				FItemList(i).FIpkumdiv2		= rsget("ipkumdiv2")
				FItemList(i).Foptlimityn		= rsget("optlimityn")
				FItemList(i).Foptionlimitno		= rsget("optlimitno")
				FItemList(i).Foptionlimitsold	= rsget("optlimitsold")
				'FItemList(i).Fdanjongyn		= rsget("danjongyn")
				FItemList(i).FreipgoMayDate		= rsget("reipgoMayDate")
				FItemList(i).Fmaxsellday		= rsget("maxsellday")
				FItemList(i).FDayForSellCount	= rsget("DayForSellCount")
				FItemList(i).FDayForSafeStock	= rsget("DayForSafeStock")
				FItemList(i).FDayForLeadTime	= rsget("DayForLeadTime")
				FItemList(i).FDayForMaxStock	= rsget("DayForMaxStock")

				FItemList(i).FOffimgMain	= rsget("offimgmain")
					if isnull(FItemList(i).FOffimgMain) then FItemList(i).FOffimgMain=""
				FItemList(i).FOffimgList	= rsget("offimglist")
					if isnull(FItemList(i).FOffimgList) then FItemList(i).FOffimgList=""
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
					if isnull(FItemList(i).FOffimgSmall) then FItemList(i).FOffimgSmall=""

				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgSmall
				FItemList(i).FrackcodeByOption	= rsget("rackcodeByOption")
				FItemList(i).FsubRackcodeByOption	= rsget("subRackcodeByOption")
				if FRectAGVCheck = "Y" then FItemList(i).FAGVStock = rsget("agvstock")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'온라인상품		'/기계가 퍼감. 페이징 방식 변경	'/2016.10.10 한용민 수정
	'//admin/stock/stockbaseday_list.asp
	public Sub GetStockBaseDayItemListOnline
		dim i,sqlStr

		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		sqlStr = " select count(i.itemid) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr & GetFromWhereStockBaseDayOnLine

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FResultCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if clng(FCurrPage)>clng(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = "SELECT *"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlstr & " 	select ROW_NUMBER() OVER (order by i.makerid, i.itemid desc, IsNull(o.itemoption, '0000') asc) as RowNum"
		sqlStr = sqlstr & " 	,i.itemid, IsNull(o.itemoption, '0000') as itemoption, i.makerid "
		sqlStr = sqlstr & " 	,i.smallimage, i.itemname, i.sellyn, i.isusing, i.mwdiv, i.limityn, i.limitno, i.limitsold ,i.deliverytype, i.danjongyn"
		sqlStr = sqlstr & " 	,i.sellcash + IsNULL(o.optaddprice,0) as sellcash, i.buycash + IsNULL(o.optaddbuyprice,0) as buycash "
		sqlStr = sqlstr & " 	,IsNULL(o.optionname,'') as itemoptionname , IsNULL(o.isusing,'Y') as optionusing, o.optlimityn, o.optlimitno, o.optlimitsold, i.optioncnt "
		sqlStr = sqlstr & " 	,T.StockReipgoDate as reipgoMayDate "
		sqlStr = sqlstr & " 	, T.DayForSellCount, T.DayForSafeStock, T.DayForLeadTime, T.DayForMaxStock "
		sqlStr = sqlStr & GetFromWhereStockBaseDayOnLine
		sqlStr = sqlStr & " ) as t"
		sqlStr = sqlStr & " WHERE t.RowNum Between "& FSPageNo &" AND "& FEPageNo &""

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new CStockBaseDayItem

				FItemList(i).Fitemgubun     = "10"
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FMakerID       = rsget("makerid")
				FItemList(i).FMwDiv         = rsget("mwdiv")
				FItemList(i).FIsUsing       = rsget("isusing")
				FItemList(i).FSellYn        = rsget("sellyn")
				FItemList(i).FLimityn       = rsget("limityn")
				FItemList(i).FLimitNo       = rsget("limitno")
				FItemList(i).FLimitSold     = rsget("limitsold")
				FItemList(i).FOptionCnt     = rsget("optioncnt")
				FItemList(i).FimageSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + rsget("smallimage")
				FItemList(i).Fbuycash		= rsget("buycash")
				FItemList(i).Foptlimityn		= rsget("optlimityn")
				FItemList(i).Foptionlimitno		= rsget("optlimitno")
				FItemList(i).Foptionlimitsold	= rsget("optlimitsold")
				FItemList(i).Fdanjongyn			= rsget("danjongyn")
				FItemList(i).FDayForSellCount	= rsget("DayForSellCount")
				FItemList(i).FDayForSafeStock	= rsget("DayForSafeStock")
				FItemList(i).FDayForLeadTime	= rsget("DayForLeadTime")
				FItemList(i).FDayForMaxStock	= rsget("DayForMaxStock")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end Sub

	'오프라인상품	'/기계가 퍼감. 페이징 방식 변경	'/2016.10.10 한용민 수정
	'//admin/stock/stockbaseday_list.asp
	public Sub GetStockBaseDayItemListOffline
		dim i,sqlStr

		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		sqlStr = " select count(i.shopitemid) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr + GetFromWhereStockBaseDayOffLine

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FResultCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = "SELECT *"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & " 	select ROW_NUMBER() OVER (order by i.makerid, s.itemgubun desc, s.itemid desc, s.itemoption asc) as RowNum"
		sqlStr = sqlStr + " 	,i.itemgubun, i.shopitemid as itemid, i.itemoption,"
		sqlStr = sqlStr + " 	i.makerid, i.shopitemname as itemname, i.shopitemoptionname as itemoptionname,"
		sqlStr = sqlStr + " 	i.shopitemprice as sellcash, i.shopsuplycash as suplycash, i.isusing, i.isusing as optionusing, i.regdate, i.extbarcode,"
		sqlStr = sqlStr + " 	d.defaultmargin, d.defaultsuplymargin, d.chargediv"
		sqlStr = sqlStr + " 	, 'Y' as sellyn, 'N' as limityn, 0 as limitno, 0 as limitsold, 'N' as optlimityn, 0 as optlimitno, 0 as optlimitsold, i.centermwdiv as mwdiv, s.preordernofix "
		sqlStr = sqlstr + " 	,T.StockReipgoDate as reipgoMayDate "
		sqlStr = sqlstr + " 	, T.DayForSellCount, T.DayForSafeStock, T.DayForLeadTime, T.DayForMaxStock "
		sqlStr = sqlStr + GetFromWhereStockBaseDayOffLine
		sqlStr = sqlStr & " ) as t"
		sqlStr = sqlStr & " WHERE t.RowNum Between "& FSPageNo &" AND "& FEPageNo &""

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new CStockBaseDayItem

				FItemList(i).Fitemgubun     = rsget("itemgubun")
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FIsUsing       = rsget("isusing")
				FItemList(i).FSellYn        = rsget("sellyn")
				FItemList(i).FLimityn       = rsget("limityn")
				FItemList(i).FLimitNo       = rsget("limitno")
				FItemList(i).FLimitSold     = rsget("limitsold")
				'FItemList(i).FOptionCnt     = rsget("optioncnt")
				'FItemList(i).FimageSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + rsget("smallimage")
				FItemList(i).Foptionusing   = rsget("optionusing")
				'FItemList(i).Fdeliverytype  = rsget("deliverytype")
				FItemList(i).FMakerID       = rsget("makerid")
				FItemList(i).FMwDiv         = rsget("mwdiv")
				FItemList(i).FSellcash      = rsget("sellcash")
				FItemList(i).Foptlimityn		= rsget("optlimityn")
				FItemList(i).Foptionlimitno		= rsget("optlimitno")
				FItemList(i).Foptionlimitsold	= rsget("optlimitsold")
				'FItemList(i).Fdanjongyn		= rsget("danjongyn")
				FItemList(i).FreipgoMayDate		= rsget("reipgoMayDate")
				FItemList(i).FDayForSellCount	= rsget("DayForSellCount")
				FItemList(i).FDayForSafeStock	= rsget("DayForSafeStock")
				FItemList(i).FDayForLeadTime	= rsget("DayForLeadTime")
				FItemList(i).FDayForMaxStock	= rsget("DayForMaxStock")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end Sub

	public function GetFromWhereOnLine(isWith,isAGV)
		dim strtmp

		'신상품은 재고가 없다. 따라서 left join 한다.
		strtmp = " from "
		strtmp = strtmp + " 	[db_item].[dbo].tbl_item i with (nolock) "
		strtmp = strtmp + " 	left join [db_item].[dbo].tbl_item_option o with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		 i.itemid = o.itemid "
		if (isWith) then
		    strtmp = strtmp + " 	left join TBL_STOCK_TARGET s "
		else
		    strtmp = strtmp + " 	left join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock) "
	    end if
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and s.itemgubun='10' "
		strtmp = strtmp + " 		and s.itemid=i.itemid "
		strtmp = strtmp + " 		and s.itemoption=IsNull(o.itemoption,'0000') "
		strtmp = strtmp + " 	left join [db_item].[dbo].tbl_item_option_Stock T with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and T.itemgubun = '10' "
		strtmp = strtmp + " 		and s.itemgubun=T.itemgubun "
        strtmp = strtmp + " 		and s.itemid=T.itemid "
		strtmp = strtmp + " 		and s.itemoption=T.itemoption "
		strtmp = strtmp + " 	left join [db_item].[dbo].tbl_item_option_Stock T0 with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and T0.itemgubun = '10' "
		strtmp = strtmp + " 		and s.itemgubun=T0.itemgubun "
        strtmp = strtmp + " 		and s.itemid=T0.itemid "
        strtmp = strtmp + " 		and s.itemoption>='0000' "
		strtmp = strtmp + " 		and T0.itemoption='0000' "
		strtmp = strtmp + " 	left join db_partner.dbo.tbl_partner p with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		i.makerid = p.id "
		strtmp = strtmp + " 	left join [db_partner].[dbo].tbl_partner_group g with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		p.groupid = g.groupid "

        if Not isWith then
		    strtmp = strtmp + " 	left join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] a with (nolock) "
		    strtmp = strtmp + " 	on "
		    strtmp = strtmp + " 		1 = 1 "
            strtmp = strtmp + " 		and a.yyyymm = '" & Left(Now(), 7)& "' "
            strtmp = strtmp + " 		and a.itemgubun = s.itemgubun "
            strtmp = strtmp + " 		and a.itemid = s.itemid "
            strtmp = strtmp + " 		and a.itemoption = s.itemoption "
        end if

        if isAGV then
		    strtmp = strtmp + " 	join [db_summary].[dbo].[tbl_current_agvstock_summary] agv with (nolock) "
		    strtmp = strtmp + " 	on "
		    strtmp = strtmp + " 		1 = 1 "
            strtmp = strtmp + " 		and agv.itemgubun = s.itemgubun "
            strtmp = strtmp + " 		and agv.itemid = s.itemid "
            strtmp = strtmp + " 		and agv.itemoption = s.itemoption "
            strtmp = strtmp + " 		AND agv.warehouseCd = 'AGV' "
        end if

		strtmp = strtmp + " where "
		strtmp = strtmp + " 	1 = 1 "

		''if isAGV then
		''    strtmp = strtmp + " and (agv.agvstock + s.ipkumdiv5+s.ipkumdiv4+s.offconfirmno) < 0 "
        ''end if

        if (FRectCD1<>"") then
            strtmp = strtmp + " and i.cate_large='" & FRectCD1 & "'"
        end if

        if (FRectCD2<>"") then
            strtmp = strtmp + " and i.cate_mid='" & FRectCD2 & "'"
        end if

        if (FRectCD3<>"") then
            strtmp = strtmp + " and i.cate_small='" & FRectCD3 & "'"
        end if
		if FRectonlyrealstockexists <> "" then
			strtmp = strtmp + " and IsNull(s.realstock,0)>0"
		end if
		if FRectOnlySell <> "" then
			strtmp = strtmp + " and i.sellyn='Y'"
		end if
		if FRectOnlyUsingItem <> "" then
			strtmp = strtmp + " and i.isusing='Y'"
		end if
		if FRectOnlyUsingItemOption <> "" then
			strtmp = strtmp + " and ((IsNULL(o.itemoption,'0000')='0000') or (IsNULL(o.isusing,'N')='Y'))"
		end if
		if FRectOnlyNotUpcheBeasong <> "" then
			strtmp = strtmp + " and i.mwdiv<>'U'"
		end if
		if FRectOnlynotDealItem <> "" then
			strtmp = strtmp + " and i.itemdiv not in (21)"
		end if
		'======================================================================
		''7일 초과부족수량
		if FRectShortage7days <> "" then
		    strtmp = strtmp + " and s.shortageno<0"
		end if
		'======================================================================
		''14일 초과부족 수량
		'======================================================================
		''7일간 필요수량			C = (A+B)/P*D
		''7일후 초과(부족)수량		(S-C-R)
		''14일후 초과(부족)수량		(S-(C*2)-R) = (S-C-R) - C
		'======================================================================
        if FRectShortage14days <> "" then
		    strtmp = strtmp + " and (s.shortageno + s.requireno)<0"
		end if
        ''현재고 N개 이하
        if FRectShortageRealStock <> "" then
		    strtmp = strtmp + " and s.realstock <= 5"
		end if
		'단종
		if (FRectOnlyNotDanjong <> "") then
			strtmp = strtmp + " and i.danjongyn <> 'Y'"
            strtmp = strtmp + " and IsNull(o.optdanjongyn, 'N') <> 'Y' "
		end if
		'재고부족
		if (FRectOnlyNotTempDanjong <> "") then
			strtmp = strtmp + " and i.danjongyn <> 'S'"
            strtmp = strtmp + " and IsNull(o.optdanjongyn, 'N') <> 'S' "
		end if
        'MD품절
        if FRectOnlyNotMDDanjong <> "" then
		    strtmp = strtmp + " and i.danjongyn <> 'M'"
            strtmp = strtmp + " and IsNull(o.optdanjongyn, 'N') <> 'M' "
		end if
		'기주문포함부족상품
		if FRectIncludePreOrder <> "" then
			strtmp = strtmp + " and (s.shortageno + s.preordernofix) < 0 "
		end if
		'한정상품 and 한정잔여수량이 있는 상품
		if FRectSkipLimitSoldOut <> "" then
			strtmp = strtmp + " and ((i.optioncnt = 0 and ((i.limityn<>'Y') or ((i.limitno - i.limitsold) > 0))) or (i.optioncnt > 0 and ((o.optlimityn<>'Y') or ((o.optlimitno - o.optlimitsold) > 0)))) "
		end if
		'실사재고마이너스상품
		if FRectOnlyStockMinus <> "" then
			strtmp = strtmp + " and s.realstock < 0 "
		end if
		If (FRectMinusStockGubun <> "") Then
			Select Case FRectMinusStockGubun
				Case "real"
					''실사재고 마이너스
					strtmp = strtmp + " and s.realstock < 0 "
				Case "check"
					''재고파악재고 마이너스
					strtmp = strtmp + " and (s.realstock + s.ipkumdiv5 + s.offconfirmno) < 0 "
				Case "may"
					''예상재고 마이너스
					strtmp = strtmp + " and ((s.realstock + s.ipkumdiv5 + s.offconfirmno) + s.ipkumdiv4 + s.ipkumdiv2 + s.offjupno) < 0 "
                Case "agv"
					''AGV재고 0 이하
					strtmp = strtmp + " and (agv.agvstock + s.ipkumdiv5+s.ipkumdiv4+s.offconfirmno) <= 0 "
				Case Else
					''
			End Select
		End If

        if FRectOnlyRealUp <> "" then
            strtmp = strtmp + " and s.realstock > 0 "
        end if

		'구매유형
		if FRectPurchaseType <> "" then
			strtmp = strtmp + " and p.purchasetype = " & FRectPurchaseType & " "
		end if
		if FRectMakerid <> "" then
			strtmp = strtmp + " and i.makerid='" + FRectMakerid + "'"
		end if
		if FRectItemGubun <> "" then
			if (FRectItemGubun <> "XX") then
				strtmp = strtmp + " and i.itemgubun = '" + CStr(FRectItemGubun) + "'"
			else
				strtmp = strtmp + " and i.itemgubun not in ('10', '70', '80', '90') "
			end if
		end if
        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	strtmp = strtmp & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	strtmp = strtmp & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if
		if FRectItemName <> "" then
			strtmp = strtmp + " and i.itemname like '%" + html2db(FRectItemName) + "%' "
		end if
		if ((FRectItemOption <> "") and (FRectItemOption <> "0000")) then
			strtmp = strtmp + " and o.itemoption='" + CStr(FRectItemOption) + "'"
		end if

		if (FRectMWDiv <> "") then
			strtmp = strtmp + " 	and IsNull(i.mwdiv,'Z') = '" + CStr(FRectMWDiv) + "' "
		end if

		if (FRectExcMkr <> "") then
			strtmp = strtmp + " 	and i.makerid <> 'ithinkso' "
		end if

		if (FRectOnlyOn <> "") then
			strtmp = strtmp + " 	and s.sell7days < 0 "
		end if

		strtmp = strtmp + " and i.itemid<>0"


		GetFromWhereOnLine = strtmp
	end function

	public function GetFromWhereOffLine(isAGV)
		dim strtmp

		'신상품은 재고가 없다. 따라서 left join 한다.
		strtmp = " from "
		strtmp = strtmp + " 	[db_shop].[dbo].tbl_shop_item i with (nolock) "
		strtmp = strtmp + " 	left join [db_shop].[dbo].tbl_shop_designer d with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and i.makerid=d.makerid "
		strtmp = strtmp + " 		and d.shopid='streetshop000' "
		strtmp = strtmp + " 	left join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and i.itemgubun=s.itemgubun "
		strtmp = strtmp + " 		and i.shopitemid=s.itemid "
		strtmp = strtmp + " 		and i.itemoption=s.itemoption "
		strtmp = strtmp + " 	left join [db_item].[dbo].tbl_item_option_Stock T with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and i.itemgubun=T.itemgubun "
		strtmp = strtmp + " 		and i.shopitemid=T.itemid "
		strtmp = strtmp + " 		and i.itemoption=T.itemoption "
		strtmp = strtmp + " 	left join db_partner.dbo.tbl_partner p with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		i.makerid = p.id "
		strtmp = strtmp + " 	left join [db_partner].[dbo].tbl_partner_group g with (nolock) "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		p.groupid = g.groupid "
        if isAGV then
		    strtmp = strtmp + " 	left join [db_summary].[dbo].[tbl_current_agvstock_summary] agv with (nolock) "
		    strtmp = strtmp + " 	on "
		    strtmp = strtmp + " 		1 = 1 "
            strtmp = strtmp + " 		and agv.itemgubun = s.itemgubun "
            strtmp = strtmp + " 		and agv.itemid = s.itemid "
            strtmp = strtmp + " 		and agv.itemoption = s.itemoption "
        end if
		strtmp = strtmp + " 	where "
		strtmp = strtmp + " 		1 = 1 "
		''strtmp = strtmp + " 		and i.itemgubun<>'10' "
		''if isAGV then
		''    strtmp = strtmp + " and (agv.agvstock + s.ipkumdiv5+s.ipkumdiv4+s.offconfirmno) < 0 "
        ''end if

        if (FRectCD1<>"") then
            strtmp = strtmp + " and i.catecdl='" & FRectCD1 & "'"
        end if

        if (FRectCD2<>"") then
            strtmp = strtmp + " and i.catecdm='" & FRectCD2 & "'"
        end if

        if (FRectCD3<>"") then
            strtmp = strtmp + " and i.catecdn='" & FRectCD3 & "'"
        end if

		'if FRectOnlySell <> "" then
		'	strtmp = strtmp + " and i.sellyn='Y'"
		'end if

		if FRectOnlyUsingItem <> "" then
			strtmp = strtmp + " and i.isusing='Y'"
		end if

		if FRectOnlyUsingItemOption <> "" then
			strtmp = strtmp + " and ((i.itemoption='0000') or (i.isusing='Y'))"
		end if

		'if FRectOnlyNotUpcheBeasong <> "" then
		'	strtmp = strtmp + " and i.centermwdiv<>'U'"
		'end if

		''7일 초과부족수량
		if FRectShortage7days <> "" then
		    strtmp = strtmp + " and s.shortageno<0"
		end if

		''14일 초과부족 수량
		'======================================================================
		''7일간 필요수량			C = (A+B)/P*D
		''7일후 초과(부족)수량		(S-C-R)
		''14일후 초과(부족)수량		(S-(C*2)-R) = (S-C-R) - C
		'======================================================================
        if FRectShortage14days <> "" then
		    strtmp = strtmp + " and (s.shortageno + s.requireno)<0"
		end if

        ''현재고 N개 이하
        if FRectShortageRealStock <> "" then
		    strtmp = strtmp + " and s.realstock <= 5"
		end if

		'단종
		'if (FRectOnlyNotDanjong <> "") then
		'	strtmp = strtmp + " and i.danjongyn <> 'Y'"
		'end if

		'재고부족
		'if (FRectOnlyNotTempDanjong <> "") then
		'	strtmp = strtmp + " and i.danjongyn <> 'S'"
		'end if

        'MD품절
        'if FRectOnlyNotMDDanjong <> "" then
		'	strtmp = strtmp + " and i.danjongyn <> 'M'"
		'end if

		'기주문포함부족상품
		if FRectIncludePreOrder <> "" then
			strtmp = strtmp + " and (s.shortageno + s.preordernofix) < 0 "
		end if

		'한정상품이 아니거나 한정잔여수량이 있는 상품
		if FRectSkipLimitSoldOut <> "" then
			strtmp = strtmp + " and ((i.limityn<>'Y') or ((i.limitno - i.limitsold) > 0))"
		end if

		'실사재고마이너스상품
		if FRectOnlyStockMinus <> "" then
			strtmp = strtmp + " and s.realstock < 0 "
		end If

		If (FRectMinusStockGubun <> "") Then
			Select Case FRectMinusStockGubun
				Case "real"
					''실사재고 마이너스
					strtmp = strtmp + " and s.realstock < 0 "
				Case "check"
					''재고파악재고 마이너스
					strtmp = strtmp + " and (s.realstock + s.ipkumdiv5 + s.offconfirmno) < 0 "
				Case "may"
					''예상재고 마이너스
					strtmp = strtmp + " and ((s.realstock + s.ipkumdiv5 + s.offconfirmno) + s.ipkumdiv4 + s.ipkumdiv2 + s.offjupno) < 0 "
                Case "agv"
					''AGV재고 마이너스
					strtmp = strtmp + " and (agv.agvstock + s.ipkumdiv5+s.ipkumdiv4+s.offconfirmno) < 0 "
				Case Else
					''
			End Select
		End If

		'구매유형
		if FRectPurchaseType <> "" then
			strtmp = strtmp + " and p.purchasetype = " & FRectPurchaseType & " "
		end if

		if FRectMakerid <> "" then
			strtmp = strtmp + " and i.makerid='" + FRectMakerid + "'"
		end if

		if FRectItemGubun <> "" then
			if (FRectItemGubun <> "XX") then
				strtmp = strtmp + " and i.itemgubun = '" + CStr(FRectItemGubun) + "'"
			else
				strtmp = strtmp + " and i.itemgubun not in ('10', '70', '80', '90') "
			end if
		end if

		if FRectItemGubunExclude <> "" then
			strtmp = strtmp + " and i.itemgubun <> '" + CStr(FRectItemGubunExclude) + "'"
		end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	strtmp = strtmp & " and i.shopitemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	strtmp = strtmp & " and i.shopitemid in (" + FRectItemid + ")"
            end if
        end if

		if FRectItemName <> "" then
			strtmp = strtmp + " and i.shopitemname like '%" + html2db(FRectItemName) + "%' "
		end if

		if ((FRectItemOption <> "") and (FRectItemOption <> "0000")) then
			strtmp = strtmp + " and i.itemoption='" + CStr(FRectItemOption) + "'"
		end if

		if (FRectMWDiv <> "") then
			strtmp = strtmp + " 	and IsNull(i.centermwdiv,'Z') = '" + CStr(FRectMWDiv) + "' "
		end if

		if (FRectExcMkr <> "") then
			strtmp = strtmp + " 	and i.makerid <> 'ithinkso' "
		end if

		strtmp = strtmp + " and i.shopitemid<>0"

		GetFromWhereOffLine = strtmp
	end function

	public function GetFromWhereStockBaseDayOnLine()
		dim strtmp

		strtmp = " from "
		strtmp = strtmp + " 	[db_item].[dbo].tbl_item i "
		strtmp = strtmp + " 	left join [db_item].[dbo].tbl_item_option o "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		 i.itemid = o.itemid "
		strtmp = strtmp + " 	left join [db_item].[dbo].tbl_item_option_Stock T "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and T.itemgubun = '10' "
		strtmp = strtmp + " 		and i.itemid=T.itemid "
		strtmp = strtmp + " 		and IsNULL(o.itemoption,'0000')=T.itemoption "
		strtmp = strtmp + " 	left join db_partner.dbo.tbl_partner p "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		i.makerid = p.id "
		strtmp = strtmp + " where "
		strtmp = strtmp + " 	1 = 1 "

		if FRectOnlySell <> "" then
			strtmp = strtmp + " and i.sellyn='Y'"
		end if

		if FRectOnlyUsingItem <> "" then
			strtmp = strtmp + " and i.isusing='Y'"
		end if

		if FRectOnlyUsingItemOption <> "" then
			strtmp = strtmp + " and ((IsNULL(o.itemoption,'0000')='0000') or (IsNULL(o.isusing,'N')='Y'))"
		end if

		if FRectOnlyNotUpcheBeasong <> "" then
			strtmp = strtmp + " and i.mwdiv<>'U'"
		end if

		'단종
		if (FRectOnlyNotDanjong <> "") then
			strtmp = strtmp + " and i.danjongyn <> 'Y'"
		end if

		'재고부족
		if (FRectOnlyNotTempDanjong <> "") then
			strtmp = strtmp + " and i.danjongyn <> 'S'"
		end if

        'MD품절
        if FRectOnlyNotMDDanjong <> "" then
		    strtmp = strtmp + " and i.danjongyn <> 'M'"
		end if

		'한정상품 and 한정잔여수량이 있는 상품
		if FRectSkipLimitSoldOut <> "" then
			strtmp = strtmp + " and ((i.optioncnt = 0 and ((i.limityn<>'Y') or ((i.limitno - i.limitsold) > 0))) or (i.optioncnt > 0 and ((o.optlimityn<>'Y') or ((o.optlimitno - o.optlimitsold) > 0)))) "
		end if

		'구매유형
		if FRectPurchaseType <> "" then
			strtmp = strtmp + " and p.purchasetype = " & FRectPurchaseType & " "
		end if

		if FRectMakerid <> "" then
			strtmp = strtmp + " and i.makerid='" + FRectMakerid + "'"
		end if

		if FRectItemGubun <> "" then
			strtmp = strtmp + " and i.itemgubun = '" + CStr(FRectItemGubun) + "'"
		end if

		if FRectItemId <> "" then
			strtmp = strtmp + " and i.itemid=" + CStr(FRectItemId) + ""
		end if

		if FRectItemName <> "" then
			strtmp = strtmp + " and i.itemname like '%" + CStr(FRectItemName) + "%' "
		end if

		if FRectOnlyNotInputDay <> "" then
			strtmp = strtmp + " and T.DayForSellCount is NULL "
		end if

		if ((FRectItemOption <> "") and (FRectItemOption <> "0000")) then
			strtmp = strtmp + " and o.itemoption='" + CStr(FRectItemOption) + "'"
		end if

		if (FRectMWDiv <> "") then
			Select Case FRectMWDiv
				Case "MW"
					strtmp = strtmp + " and i.mwdiv in ('M', 'W') "
				Case "M"
					strtmp = strtmp + " and i.mwdiv = 'M' "
				Case "W"
					strtmp = strtmp + " and i.mwdiv = 'W' "
				Case "U"
					strtmp = strtmp + " and i.mwdiv = 'U' "
				Case Else
					''
			End Select
		end if

		if FRectDayForSellCount <> "" then
			strtmp = strtmp + " and t.DayForSellCount = " & FRectDayForSellCount & " "
		end if

		strtmp = strtmp + " and i.itemid<>0"

		GetFromWhereStockBaseDayOnLine = strtmp
	end function

	public function GetFromWhereStockBaseDayOffLine()
		dim strtmp

		'신상품은 재고가 없다. 따라서 left join 한다.
		strtmp = " from "
		strtmp = strtmp + " 	[db_shop].[dbo].tbl_shop_item i "
		strtmp = strtmp + " 	left join [db_shop].[dbo].tbl_shop_designer d "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and i.makerid=d.makerid "
		strtmp = strtmp + " 		and d.shopid='streetshop000' "
		strtmp = strtmp + " 	left join [db_summary].[dbo].tbl_current_logisstock_summary s "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and i.itemgubun=s.itemgubun "
		strtmp = strtmp + " 		and i.shopitemid=s.itemid "
		strtmp = strtmp + " 		and i.itemoption=s.itemoption "
		strtmp = strtmp + " 	left join [db_item].[dbo].tbl_item_option_Stock T "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and i.itemgubun=T.itemgubun "
		strtmp = strtmp + " 		and i.shopitemid=T.itemid "
		strtmp = strtmp + " 		and i.itemoption=T.itemoption "
		strtmp = strtmp + " 	left join db_partner.dbo.tbl_partner p "
		strtmp = strtmp + " 	on "
		strtmp = strtmp + " 		i.makerid = p.id "
		strtmp = strtmp + " 	where "
		strtmp = strtmp + " 		1 = 1 "
		strtmp = strtmp + " 		and i.itemgubun<>'10' "

		'if FRectOnlySell <> "" then
		'	strtmp = strtmp + " and i.sellyn='Y'"
		'end if

		if FRectOnlyUsingItem <> "" then
			strtmp = strtmp + " and i.isusing='Y'"
		end if

		if FRectOnlyUsingItemOption <> "" then
			strtmp = strtmp + " and ((i.itemoption='0000') or (i.isusing='Y'))"
		end if

		'if FRectOnlyNotUpcheBeasong <> "" then
		'	strtmp = strtmp + " and i.centermwdiv<>'U'"
		'end if

		''7일 초과부족수량
		if FRectShortage7days <> "" then
		    strtmp = strtmp + " and s.shortageno<0"
		end if

		''14일 초과부족 수량
		'======================================================================
		''7일간 필요수량			C = (A+B)/P*D
		''7일후 초과(부족)수량		(S-C-R)
		''14일후 초과(부족)수량		(S-(C*2)-R) = (S-C-R) - C
		'======================================================================
        if FRectShortage14days <> "" then
		    strtmp = strtmp + " and (s.shortageno + s.requireno)<0"
		end if

        ''현재고 N개 이하
        if FRectShortageRealStock <> "" then
		    strtmp = strtmp + " and s.realstock <= 5"
		end if

		'단종
		'if (FRectOnlyNotDanjong <> "") then
		'	strtmp = strtmp + " and i.danjongyn <> 'Y'"
		'end if

		'재고부족
		'if (FRectOnlyNotTempDanjong <> "") then
		'	strtmp = strtmp + " and i.danjongyn <> 'S'"
		'end if

        'MD품절
        'if FRectOnlyNotMDDanjong <> "" then
		'	strtmp = strtmp + " and i.danjongyn <> 'M'"
		'end if

		'기주문포함부족상품
		if FRectIncludePreOrder <> "" then
			strtmp = strtmp + " and (s.shortageno + s.preordernofix) < 0 "
		end if

		'한정상품이 아니거나 한정잔여수량이 있는 상품
		if FRectSkipLimitSoldOut <> "" then
			strtmp = strtmp + " and ((i.limityn<>'Y') or ((i.limitno - i.limitsold) > 0))"
		end if

		'실사재고마이너스상품
		if FRectOnlyStockMinus <> "" then
			strtmp = strtmp + " and s.realstock < 0 "
		end if

		'구매유형
		if FRectPurchaseType <> "" then
			strtmp = strtmp + " and p.purchasetype = " & FRectPurchaseType & " "
		end if

		if FRectMakerid <> "" then
			strtmp = strtmp + " and i.makerid='" + FRectMakerid + "'"
		end if

		if FRectItemGubun <> "" then
			strtmp = strtmp + " and i.itemgubun='" + CStr(FRectItemGubun) + "'"
		end if

		if FRectItemGubunExclude <> "" then
			strtmp = strtmp + " and i.itemgubun <> '" + CStr(FRectItemGubunExclude) + "'"
		end if

		if FRectItemId <> "" then
			strtmp = strtmp + " and i.shopitemid=" + CStr(FRectItemId) + ""
		end if

		if ((FRectItemOption <> "") and (FRectItemOption <> "0000")) then
			strtmp = strtmp + " and i.itemoption='" + CStr(FRectItemOption) + "'"
		end If

		if (FRectCenterMWDiv <> "") then
			Select Case FRectCenterMWDiv
				Case "MW"
					strtmp = strtmp + " and i.centermwdiv in ('M', 'W') "
				Case "M"
					strtmp = strtmp + " and i.centermwdiv = 'M' "
				Case "W"
					strtmp = strtmp + " and i.centermwdiv = 'W' "
				Case "U"
					strtmp = strtmp + " and i.centermwdiv = 'U' "
				Case "N"
					strtmp = strtmp + " and IsNull(i.centermwdiv, 'N') = 'N' "
				Case Else
					''
			End Select
		end if

		strtmp = strtmp + " and i.shopitemid<>0"

		GetFromWhereStockBaseDayOffLine = strtmp
	end function

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
%>
