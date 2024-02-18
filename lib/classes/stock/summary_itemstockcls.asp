<%
'###########################################################
' Description : 재고
' Hieditor : 2015.05.27 이상구 생성
'			 2017.09.27 한용민 수정
'###########################################################

Class CErritemBrandGroupItem
    public Fmakerid
    public FOncnt
    public FOffcnt

    Private Sub Class_Initialize()
		FOncnt =0
		FOffcnt=0
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CBadOrErritemBrandGroupItem
    public Fmakerid
    public Fuseyn
    public Fmakername
	public Fcompany_no
	public Fcompany_name
    public FOnCnt
    public FOffCnt
    public Fitem10M
    public Fitem10W
    public Fitem10U
	public Fitem10Z
    public Fitem70
    public Fitem80
    public Fitem90M
    public Fitem90W
    public Fitem90U
	public Fitem90Z
	public FitemetcM
	public FitemetcW
	public FitemetcZ

    Private Sub Class_Initialize()
		FOnCnt		= 0
		FOffCnt		= 0
		Fitem10M	= 0
		Fitem10W	= 0
		Fitem10U	= 0
		Fitem70		= 0
		Fitem80		= 0
		Fitem90M	= 0
		Fitem90W	= 0
		Fitem90U	= 0
		FitemetcM	= 0
		FitemetcW	= 0
		FitemetcZ	= 0
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CBadOrErrItemItem
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fregitemno
	public Frealstock
	public FItemName
	public Fmakerid
	public Fsellcash
	public FBuycash
	public flastmwdiv
	public Fmwdiv
	public Fcentermwdiv
	public Fdeliverytype
	public FItemOptionName
    public FimgSmall
    public Fsellyn
    public Fisusing
	public FlastIpgoDate

	public function GetdeliverytypeName()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			GetdeliverytypeName = "업배"
		else
			GetdeliverytypeName = "텐배"
		end if
	end function

	public function GetMwDivName()
		if Fmwdiv="M" then
			GetMwDivName = "매입"
		elseif Fmwdiv="W" then
			GetMwDivName = "위탁"
		elseif Fmwdiv="U" then
			GetMwDivName = "업체"
		else
			GetMwDivName = Fmwdiv
		end if
	end function

	public function GetMwDivColor()
		if Fmwdiv="M" then
			GetMwDivColor = "red"
		elseif Fmwdiv="W" then
			GetMwDivColor = "black"
		elseif Fmwdiv="U" then
			GetMwDivColor = "green"
		end if
	end function

	Private Sub Class_Initialize()
		Fregitemno      = 0
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CTurnOverBrand
    public Fmakerid
    public Fcnt
    public Frealstock
	public Fcurrrealstock
    public Fsellno
    public Foffchulgono
	public Freipgono

    Private Sub Class_Initialize()
		Fcnt =0
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CTempRegItem
	public FItemgubun
	public FItemId
	public FItemOption
	public Fitemno

	Private Sub Class_Initialize()
		Fitemno =0
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CErritemDailyItem
	public Fyyyymmdd
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Ferrcsno
	public Ferrbaditemno
	public Ferrrealcheckno
	public Ferretcno
	public Ftoterrno
	public Freguser
	public Fmodiuser
	public Fregdate
	public Flastupdate
	public FItemName
	public Fmakerid
	public Fsellcash
	public FBuycash
	public Fmwdiv
	public Fdeliverytype
	public FItemOptionName
    public FimgSmall

	public function GetdeliverytypeName()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			GetdeliverytypeName = "업배"
		else
			GetdeliverytypeName = "텐배"
		end if
	end function

	public function GetMwDivName()
		if Fmwdiv="M" then
			GetMwDivName = "매입"
		elseif Fmwdiv="W" then
			GetMwDivName = "위탁"
		elseif Fmwdiv="U" then
			GetMwDivName = "업체"
		end if
	end function

	Private Sub Class_Initialize()
		Fyyyymmdd      = 0
		Fitemgubun     = 0
		Fitemid        = 0
		Fitemoption    = 0
		Ferrcsno       = 0
		Ferrbaditemno  = 0
		Ferrrealcheckno= 0
		Ferretcno      = 0
		Ftoterrno      = 0
		Freguser       = 0
		Fmodiuser      = 0
		Fregdate       = 0
		Flastupdate    = 0
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CCurrentStockItem
	public Fitemgubun
	public Fitemid
	public Fitemname
	public Fitemoption
	public FitemoptionName
	public Fmakerid
	public Fdeliverytype
	public Fsellcash
	public Fbuycash
	public FOffLineDefaultMargin
	public FOffLineDefaultSuplyMargin
	public Fisusing
	public Flimityn
	public FLimitNo
	public FLimitSold
	public Flimitcount
	public Fsellyn
	public Fmwdiv
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
	public Favailsysstock           '''사용중지.
	public Frealstock               '''실사유효재고(불량반영된 값)
	public Fsell7days
	public Foffchulgo7days
	public Fipkumdiv5
	public Fipkumdiv4
	public Fipkumdiv2
	public Foffconfirmno
	public Foffjupno
	public Frequireno
	public FrequireMaxno		'// (FDayForMaxStock + FDayForLeadTime) 필요재고
	public Fshortageno
	public Fpreorderno
	public Fpreordernofix
	public Foffsellno
	public Fmaxsellday
	public Fimgsmall
	public Fregdate
	public Flastupdate
	public FOldSystemCurrno
	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Fdanjongyn
	public FItemrackcode
	public FItemsubrackcode
	public fprtidx
	public fsubitemrackcode
    public FOnlineCurrentSellcash
    public FOnlineCurrentBuycash
	public Forgprice				'// 소비자가
    public Fpre1RealStock
    public Fpre1chulgono
    public Fpre2RealStock
    public Fpre2chulgono
    public Faccumchulgo
    public FDayForSellCount
    public FDayForSafeStock
    public FDayForLeadTime
    public FDayForMaxStock
	public Fcurrrealstock
	Public FitemCnt
	Public FitemPlusCnt
    public Fstockreipgodate		''재입고 예정일
    public Foptioncnt
    public FCentermwdiv ''오프 센터매입구분 //2014/07/11
	public FlastIpgoDate	'// 마지막 입고월
	public FprevMonthSellCnt
	public FOffMwMargin
	public foptrackcode
	public FsellStdate	'판매시작일

	Public FpublicBarcode
    public Fitemgrade
    Public Fagvstock
    Public Fbulkstock
    Public Flastbulkstockdate
    public FwarehouseCd

    public Foptsellyn
    public Foptisusing

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

    ''(12+2) 일후 부족수량
	public function GetShortageMaxNo()
		'// shortageno		= realstock+requireno+ipkumdiv5+offconfirmno+ipkumdiv4+ipkumdiv2+offjupno
		'// shortageMaxno	= realstock+requireMaxno+ipkumdiv5+offconfirmno+ipkumdiv4+ipkumdiv2+offjupno
		'// 따라서
		'// shortageMaxno = shortageno + (requireMaxno - requireno)

		GetShortageMaxNo = Fshortageno + (FrequireMaxno - Frequireno)
	end function

    ''단종 이름
    public function getDanjongNameHTML()
        if IsNULL(Fdanjongyn) then Exit function

        if (Fdanjongyn="Y") then
            getDanjongNameHTML = "<font color='#33CC33'>단종</font>"
        elseif (Fdanjongyn="S") then
            getDanjongNameHTML = "<font color='#3333CC'>재고<br>부족</font>"
        elseif (Fdanjongyn="M") then
            getDanjongNameHTML = "<font color='#CC3333'>MD<br>품절</font>"
        elseif (Fdanjongyn="N") then
            getDanjongNameHTML = ""
        else
            getDanjongNameHTML = Fdanjongyn
        end if
    end function

    ''예상재고
	public function GetMaystock()
		GetMaystock = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2 + Foffjupno
	end function

    ''한정비교재고 2011-06-23 수정 : 오프 준비중 수량 제외. => 2011-06-24 다시 원상복귀
	public function GetLimitStockNo()
		GetLimitStockNo = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2
		''GetLimitStockNo = Frealstock + Fipkumdiv5 + Fipkumdiv4 + Fipkumdiv2
	end function

    ''재고파악재고
	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo '' 오프 준비중 포함
	end function

    ''금일 상품준비수량
	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function

    ''출고이전 필요수량(접수,결제완료..)
    public function GetReqNotChulgoNo()
		GetReqNotChulgoNo = Fipkumdiv5 + Foffconfirmno + Fipkumdiv4 + Fipkumdiv2 + Foffjupno
	end function

    ''' 실사재고 == 시스템재고+실사오차 == 실유효재고-불량재고 == Frealstock-Ferrbaditemno ''불량재고 따로관리
    public function getErrAssignStock()
        getErrAssignStock = (Ftotsysstock+Ferrrealcheckno)
    end function

	public function GetdeliverytypeName()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			GetdeliverytypeName = "업배"
		else
			GetdeliverytypeName = "텐배"
		end if
	end function

	public function GetMwDivName()
		if Fmwdiv="M" then
			GetMwDivName = "매입"
		elseif Fmwdiv="W" then
			GetMwDivName = "위탁"
		elseif Fmwdiv="U" then
			GetMwDivName = "업체"
		else
			GetMwDivName = Fmwdiv
		end if
	end function

	public function GetLimitNo()
		GetLimitNo = FLimitNo-FLimitSold

		if GetLimitNo<0 then GetLimitNo=0
	end function

	public function IsSoldOut()
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn<>"N") and (FLimitNo-FLimitSold<1))
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
				if FOptLimitNo-FOptLimitSold<1 then
					GetLimitStr = "0"
				else
					GetLimitStr = CStr(Foptlimitno-Foptlimitsold)
				end if
			end if
		end if
	end function

	Private Sub Class_Initialize()
		Fipgono		= 0
		Freipgono	= 0
		Ftotipgono	= 0
		Foffchulgono	= 0
		Foffrechulgono	= 0
		Fetcchulgono	= 0
		Fetcrechulgono	= 0
		Ftotchulgono	= 0
		Fsellno		= 0
		Fresellno	= 0
		Ftotsellno	= 0
		Ferrcsno	= 0
		Ferrbaditemno	= 0
		Ferrrealcheckno	= 0
		Ferretcno	= 0
		Ftoterrno	= 0
		Ftotsysstock	= 0
		Favailsysstock	= 0
		Frealstock	= 0
		Fsell7days	= 0
		Foffchulgo7days	= 0
		Fipkumdiv5		= 0
		Fipkumdiv4		= 0
		Fipkumdiv2		= 0
		Foffconfirmno	= 0
		Foffjupno		= 0
		Frequireno		= 0
		Fshortageno		= 0
		Fpreorderno		= 0
		Fpreordernofix		= 0
		Fmaxsellday		= 0

        FOnlineCurrentSellcash =0
        FOnlineCurrentBuycash =0
	end sub

	Private Sub Class_Terminate()
End Sub

end Class

Class CSummaryItemStockItem
	public Fyyyymm
	public Fyyyymmdd
	public Fitemgubun
	public Fitemid
	public Fitemoption
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
	public Foffsellno
	public Ftotsysstock
	public Favailsysstock
	public Frealstock
	public Fregdate
	public Flastupdate
    public Flastmwdiv
    public Flastbuyprice
    public Favgipgoprice
    public Flasttotsysstock

	public function GetDpartName()
		dim dpart
		dpart = DatePart("w",Fyyyymmdd)
		if dpart=1 then
			GetDpartName = "<font color=#FF0000>일</font>"
		elseif dpart=2 then
			GetDpartName = "월"
		elseif dpart=3 then
			GetDpartName = "화"
		elseif dpart=4 then
			GetDpartName = "수"
		elseif dpart=5 then
			GetDpartName = "목"
		elseif dpart=6 then
			GetDpartName = "금"
		elseif dpart=7 then
			GetDpartName = "<font color=#0000FF>토</font>"
		else
			GetDpartName = ""
		end if
	end function

	Private Sub Class_Initialize()
		Fipgono         = 0
		Freipgono       = 0
		Ftotipgono      = 0
		Foffchulgono    = 0
		Foffrechulgono  = 0
		Fetcchulgono    = 0
		Fetcrechulgono  = 0
		Ftotchulgono    = 0
		Fsellno         = 0
		Fresellno       = 0
		Ftotsellno      = 0
		Ferrcsno        = 0
		Ferrbaditemno   = 0
		Ferrrealcheckno = 0
		Ferretcno       = 0
		Ftoterrno       = 0
		Ftotsysstock    = 0
		Favailsysstock  = 0
		Frealstock      = 0
		Fregdate        = 0
		Flastupdate     = 0

	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CSummaryItemStock
	public FItemList()
	public FOneItem
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public farrlist
	public FRectMakerid
	public FRectNotUpBaeSong
	public FRectKindDisplay
	public FRectKindSort
	public FRectParameter
	public FRectDiffDiv
	public FRectMWDiv
    public FRectCenterMWDiv
	public FRectItemGubun
	public FRectItemID
	public FRectItemOption
	public FRectStartDate
	public FRectEndDate
	public FRectYYYYMM
	public FRectSellYN
	public FRectDispCate
	public FRectLimitSoldOut
	public FRectOnlyIsUsing
    public FRectOptIsUsing
    public frectsoldout_gubun
	public FRectOnlySellyn
    public FRectOptSellYN
	public FRectSearchMode
	public FRectDanjongyn
	public FRectLimityn
	public FRectUseYN
	public FRectrealstocknotzero
	public FRectCD1
	public FRectCD2
	public FRectCD3
	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Fdanjongyn
    public FRectOnlyOldItem
    public FRectOnlyOutItem
	Public FRectExcBaseRegItem
    public FTotalExistsItemCount
    public FRectChulgoNo
    public FRectTurnOverPro
    public FRectState
    public FRectReturnItemGubun
    public FRectOrderBy
	public FRectGroupBy
    public FRectItemIdArr
    public FRectItemName
	public FRectlimitrealstock
    public FRectSearchType
    public FRectPurchaseType
    public FRectDatetype
	public FRectMonthGubun
	public FRectMakerUseYN
	public FRectMonthDiff
	public FRectStockType
	Public FRectUseOffInfo
	Public FRectitemrackcode
	public FRectRackCode
	public FRectFromRackcode2
	public FRectToRackcode2
	Public FRectExcIts
    public FRectExcNoRack
    Public FRectItemGrade
    public FRectBulkStockGubun
    public FRectWarehouseCd
    public FRectAgvStockGubun
	public FRectlastmwdiv
	public FRectTplGubun
	public FRectIsSellStart
	public FRectStockMwDiv

	public Sub GetOneItemOptionStock()
		dim i,sqlstr

		sqlstr = " select T.itemid, T.itemoption, T.makerid,"
		sqlstr = sqlstr + " T.smallimage, T.listimage, "
		sqlstr = sqlstr + " T.itemname, v.optionname, "
		sqlstr = sqlstr + " T.isusing, T.sellyn, T.limityn, T.limitno, T.limitsold, T.optionusing,"
		sqlstr = sqlstr + " T.pojangok, T.mwdiv, T.sellcash, T.buycash, "
		sqlstr = sqlstr + " T.optioncnt, T.itemrackcode, T.brandname, T.deliverytype, "
		sqlstr = sqlstr + " IsNull(s.currno,0) as oldstockcurrno, s.regdate as oldstockupdate, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate"
		sqlstr = sqlstr + " from ("
		sqlstr = sqlstr + "  	select i.itemid, i.makerid, i.itemname, IsNULL(o.itemoption,'0000') as itemoption ,"
		sqlstr = sqlstr + "  	IsNULL(o.optionname,'') as optionname,  i.isusing,"
		sqlstr = sqlstr + "  	i.sellyn, i.limityn, i.limitno, i.limitsold, i.deliverytype, "
		sqlstr = sqlstr + "  	i.pojangok, i.sellcash, i.buycash, i.mwdiv, i.smallimage, i.listimage, "
		sqlstr = sqlstr + "  	i.optioncnt, i.itemrackcode, i.brandname, "
		sqlstr = sqlstr + "  	o.isusing as optionusing "
		sqlstr = sqlstr + "  	from [db_item].[dbo].tbl_item i "
		sqlstr = sqlstr + "  	left join [db_item].[dbo].tbl_item_option o "
		sqlstr = sqlstr + "  	on i.itemid=o.itemid"
		sqlstr = sqlstr + "  	where i.itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " ) as T"

		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " on T.itemid=sm.itemid and T.itemoption=sm.itemoption and sm.itemgubun='10'"
		sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_const_day_stock s"
		sqlstr = sqlstr + " on T.itemid=s.itemid and T.itemoption=s.itemoption"

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CRealJaeGoItem

				FItemList(i).FItemID          = rsget("itemid")
				FItemList(i).FItemOption      = rsget("itemoption")
				FItemList(i).FItemName        = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName  = db2html(rsget("optionname"))
				FItemList(i).Fmakerid		= rsget("makerid")
				FItemList(i).Fbrandname		= db2html(rsget("brandname"))
				FItemList(i).FIsUsing  		= rsget("isusing")
				FItemList(i).FSellYn   		= rsget("sellyn")
				FItemList(i).FLimityn  		= rsget("limityn")
				FItemList(i).FLimitNo  		= rsget("limitno")
				FItemList(i).FLimitSold		= rsget("limitsold")
				FItemList(i).Fsellcash		= rsget("sellcash")
				FItemList(i).Fbuycash		= rsget("buycash")
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")
				FItemList(i).FImageList	= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("listimage")
				FItemList(i).Foptionusing = rsget("optionusing")
				FItemList(i).Fpojangok = rsget("pojangok")
				FItemList(i).Fmwdiv = rsget("mwdiv")
				FItemList(i).Fdeliverytype = rsget("deliverytype")
				FItemList(i).Foptioncnt = rsget("optioncnt")
				FItemList(i).Fitemrackcode = rsget("itemrackcode")
				FItemList(i).FItemNo          = rsget("oldstockcurrno")
				FItemList(i).Fregdate = rsget("oldstockupdate")
				''아래 명으로 변경
				FItemList(i).Foldstockupdate = rsget("oldstockupdate")
				FItemList(i).Foldstockcurrno = rsget("oldstockcurrno")
				FItemList(i).Frealstock		 = rsget("realstock")
				FItemList(i).Fipkumdiv5		 = rsget("ipkumdiv5")
				FItemList(i).Foffconfirmno	 = rsget("offconfirmno")
				FItemList(i).FLastUpdate	 = rsget("lastupdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public sub GetTodayErrItem()
		dim sqlstr, i

		sqlstr = "select top 1 s.* "
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_erritem_daily_summary s"
		sqlstr = sqlstr + " where yyyymmdd=convert(varchar(10),getdate(),21)"
		sqlstr = sqlstr + " and s.itemgubun='" + CStr(FRectItemGubun) + "'"
		sqlstr = sqlstr + " and s.itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and s.itemoption='" + CStr(FRectItemOption) + "'"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		set FOneItem = new CErritemDailyItem

		i=0
		if  not rsget.EOF  then
			FOneItem.Fyyyymmdd       = rsget("yyyymmdd")
			FOneItem.Fitemgubun      = rsget("itemgubun")
			FOneItem.Fitemid         = rsget("itemid")
			FOneItem.Fitemoption     = rsget("itemoption")
			FOneItem.Ferrcsno        = rsget("errcsno")
			FOneItem.Ferrbaditemno   = rsget("errbaditemno")
			FOneItem.Ferrrealcheckno = rsget("errrealcheckno")
			FOneItem.Ferretcno       = rsget("erretcno")
			FOneItem.Ftoterrno       = rsget("toterrno")
			FOneItem.Freguser        = rsget("reguser")
			FOneItem.Fmodiuser       = rsget("modiuser")
			FOneItem.Fregdate        = rsget("regdate")
			FOneItem.Flastupdate     = rsget("lastupdate")
		end if
		rsget.close
	end sub

	public sub GetDailyErrItemList()
		dim sqlstr, i

		sqlstr = "select top 1000 s.*,"
		sqlstr = sqlstr + " isnull(i.makerid,ii.makerid) as makerid, isnull(i.itemname,ii.shopitemname) as itemname, i.deliverytype, isnull(i.sellcash,ii.shopitemprice) as sellcash, isnull(i.buycash,0) as buycash, i.mwdiv,"
		sqlstr = sqlstr + " IsNULL(v.optionname,'') as codeview "
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_erritem_daily_summary s"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.itemid=i.itemid"
		sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item ii on s.itemgubun <> '10' and s.itemgubun=ii.itemgubun and s.itemid=ii.shopitemid and s.itemoption=ii.itemoption "
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v on s.itemgubun='10' and s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where yyyymmdd>='" + FRectStartDate + "'"
		sqlstr = sqlstr + " and yyyymmdd<'" + FRectEndDate + "'"

		if FRectItemGubun<>"" then
			sqlstr = sqlstr + " and s.itemgubun='" + CStr(FRectItemGubun) + "'"
		end if

		if FRectItemID<>"" then
			sqlstr = sqlstr + " and s.itemid=" + CStr(FRectItemID)
		end if

		if FRectItemOption<>"" then
			sqlstr = sqlstr + " and s.itemoption='" + CStr(FRectItemOption) + "'"
		end if

		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and ((i.makerid='" + CStr(FRectMakerid) + "') or (ii.makerid='" + CStr(FRectMakerid) + "'))"
		end if

		if FRectKindDisplay="B" then
			sqlstr = sqlstr + " and errbaditemno<>0 "
		elseif FRectKindDisplay="D" then
			sqlstr = sqlstr + " and errrealcheckno<>0 "
		end if

		sqlstr = sqlstr + " order by yyyymmdd"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CErritemDailyItem

				FItemList(i).Fyyyymmdd       = rsget("yyyymmdd")
				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Ferrcsno        = rsget("errcsno")
				FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).Ferretcno       = rsget("erretcno")
				FItemList(i).Ftoterrno       = rsget("toterrno")
				FItemList(i).Freguser        = rsget("reguser")
				FItemList(i).Fmodiuser       = rsget("modiuser")
				FItemList(i).Fregdate        = rsget("regdate")
				FItemList(i).Flastupdate     = rsget("lastupdate")

				FItemList(i).FItemName		= db2html(rsget("itemname"))
				FItemList(i).FMakerid		= rsget("makerid")
				FItemList(i).Fsellcash		= rsget("sellcash")
				FItemList(i).Fbuycash		= rsget("buycash")
				FItemList(i).Fmwdiv			= rsget("mwdiv")
				FItemList(i).Fdeliverytype	= rsget("deliverytype")
				FItemList(i).FItemOptionName= db2html(rsget("codeview"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

    public sub GetDailyErrBadItemListByBrandGroup()
        dim sqlstr, i

        sqlstr = "select top 500 c.userid as Makerid"
        sqlstr = sqlstr + " , IsNULL(T1.Cnt,0) as OnCnt"
        sqlstr = sqlstr + " , IsNULL(T2.Cnt,0) as OffCnt"
        sqlstr = sqlstr + " from [db_user].[dbo].tbl_user_c c"
        sqlstr = sqlstr + " 	left join ("
        sqlstr = sqlstr + " 		select i.makerid, sum(s.errbaditemno)  as cnt "
        sqlstr = sqlstr + " 		from [db_summary].[dbo].tbl_current_logisstock_summary s"
        sqlstr = sqlstr + " 		 join [db_item].[dbo].tbl_item i"
        sqlstr = sqlstr + " 		on s.itemgubun='10' "
        sqlstr = sqlstr + " 		and s.itemid=i.itemid"
        sqlstr = sqlstr + " 		where s.errbaditemno<>0"
        sqlstr = sqlstr + " 		group by i.makerid"
        sqlstr = sqlstr + " 	) T1 on  c.userid=T1.makerid"
        sqlstr = sqlstr + " 	left join ("
        sqlstr = sqlstr + " 		select f.makerid, sum(s.errbaditemno)  as cnt "
        sqlstr = sqlstr + " 		from [db_summary].[dbo].tbl_current_logisstock_summary s"
        sqlstr = sqlstr + " 		 join [db_shop].[dbo].tbl_shop_item f "
        sqlstr = sqlstr + " 		on s.itemgubun<>'10' "
        sqlstr = sqlstr + " 		and s.itemgubun=f.itemgubun "
        sqlstr = sqlstr + " 		and s.itemid=f.shopitemid "
        sqlstr = sqlstr + " 		and s.itemoption=f.itemoption"
        sqlstr = sqlstr + " 		where s.errbaditemno<>0"
        sqlstr = sqlstr + " 		group by f.makerid"
        sqlstr = sqlstr + " 	) T2 on  c.userid=T2.makerid"
        sqlstr = sqlstr + " where T1.makerid is Not NULL"
        sqlstr = sqlstr + " or T2.makerid is Not NULL"
        sqlstr = sqlstr + " order by OnCnt, offCnt"

        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CErritemBrandGroupItem

				FItemList(i).FMakerid       = rsget("makerid")
				FItemList(i).FonCnt           = rsget("OnCnt")
				FItemList(i).FoffCnt          = rsget("OffCnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public sub GetDailyErrRealCheckItemListByBrandGroup()
        dim sqlstr, i

        sqlstr = "select top 500 c.userid as Makerid"
        sqlstr = sqlstr + " , IsNULL(T1.Cnt,0) as OnCnt"
        sqlstr = sqlstr + " , IsNULL(T2.Cnt,0) as OffCnt"
        sqlstr = sqlstr + " from [db_user].[dbo].tbl_user_c c"
        sqlstr = sqlstr + " 	left join ("
        sqlstr = sqlstr + " 		select i.makerid, sum(s.errrealcheckno)  as cnt "
        sqlstr = sqlstr + " 		from [db_summary].[dbo].tbl_current_logisstock_summary s"
        sqlstr = sqlstr + " 		 join [db_item].[dbo].tbl_item i"
        sqlstr = sqlstr + " 		on s.itemgubun='10' "
        sqlstr = sqlstr + " 		and s.itemid=i.itemid"
        sqlstr = sqlstr + " 		where s.errrealcheckno<>0"
        sqlstr = sqlstr + " 		group by i.makerid"
        sqlstr = sqlstr + " 	) T1 on  c.userid=T1.makerid"
        sqlstr = sqlstr + " 	left join ("
        sqlstr = sqlstr + " 		select f.makerid, sum(s.errrealcheckno)  as cnt "
        sqlstr = sqlstr + " 		from [db_summary].[dbo].tbl_current_logisstock_summary s"
        sqlstr = sqlstr + " 		 join [db_shop].[dbo].tbl_shop_item f "
        sqlstr = sqlstr + " 		on s.itemgubun<>'10' "
        sqlstr = sqlstr + " 		and s.itemgubun=f.itemgubun "
        sqlstr = sqlstr + " 		and s.itemid=f.shopitemid "
        sqlstr = sqlstr + " 		and s.itemoption=f.itemoption"
        sqlstr = sqlstr + " 		where s.errrealcheckno<>0"
        sqlstr = sqlstr + " 		group by f.makerid"
        sqlstr = sqlstr + " 	) T2 on  c.userid=T2.makerid"
        sqlstr = sqlstr + " where T1.makerid is Not NULL"
        sqlstr = sqlstr + " or T2.makerid is Not NULL"
        sqlstr = sqlstr + " order by OnCnt, offCnt"

        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CErritemBrandGroupItem

				FItemList(i).FMakerid       = rsget("makerid")
				FItemList(i).FonCnt           = rsget("OnCnt")
				FItemList(i).FoffCnt          = rsget("OffCnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	' /admin/stock/badorerritem_act_list.asp
    public sub GetBadOrErrItemListByBrandGroup()
        dim sqlstr, i, sqlMWDivON, sqlMWDivOFF

		if (FRectDatetype="yyyymm") then
			sqlMWDivON  = " IsNull(s.lastmwdiv, 'Z') "
			sqlMWDivOFF = " IsNull(s.lastmwdiv, 'Z') "
		else
			sqlMWDivON  = " IsNull(IsNull(ms.lastmwdiv,i.mwdiv), 'Z') "
			sqlMWDivOFF = " IsNull(i.centermwdiv, 'Z') "
		end if

		sqlstr = " SELECT TOP 500 "
        sqlstr = sqlstr + " 	c.userid as Makerid "
        sqlstr = sqlstr + " 	, c.isusing as useyn "
        sqlstr = sqlstr + " 	, c.socname as makername "
		sqlstr = sqlstr + " 	, p.company_no "
		sqlstr = sqlstr + " 	, g.company_name "
        sqlstr = sqlstr + " 	, IsNULL(T1.OnBadCnt,0) as OnBadCnt "
        sqlstr = sqlstr + " 	, IsNULL(T1.OnErrCnt,0) as OnErrCnt "
        sqlstr = sqlstr + " 	, IsNULL(T2.OffBadCnt,0) as OffBadCnt "
        sqlstr = sqlstr + " 	, IsNULL(T2.OffErrCnt,0) as OffErrCnt "
        sqlstr = sqlstr + " 	, IsNULL(T1.baditem10M,0) as baditem10M "
        sqlstr = sqlstr + " 	, IsNULL(T1.baditem10W,0) as baditem10W "
        sqlstr = sqlstr + " 	, IsNULL(T1.baditem10U,0) as baditem10U "
		sqlstr = sqlstr + " 	, IsNULL(T1.baditem10Z,0) as baditem10Z "
        sqlstr = sqlstr + " 	, IsNULL(T1.erritem10M,0) as erritem10M "
        sqlstr = sqlstr + " 	, IsNULL(T1.erritem10W,0) as erritem10W "
        sqlstr = sqlstr + " 	, IsNULL(T1.erritem10U,0) as erritem10U "
		sqlstr = sqlstr + " 	, IsNULL(T1.erritem10Z,0) as erritem10Z "
        'sqlstr = sqlstr + " 	, IsNULL(T2.baditem70,0) as baditem70 "
        'sqlstr = sqlstr + " 	, IsNULL(T2.baditem80,0) as baditem80 "
        sqlstr = sqlstr + " 	, IsNULL(T2.baditem90M,0) as baditem90M "
        sqlstr = sqlstr + " 	, IsNULL(T2.baditem90W,0) as baditem90W "
        sqlstr = sqlstr + " 	, IsNULL(T2.baditem90U,0) as baditem90U "
		sqlstr = sqlstr + " 	, IsNULL(T2.baditem90Z,0) as baditem90Z "
        sqlstr = sqlstr + " 	, IsNULL(T2.baditemetcM,0) as baditemetcM "
        sqlstr = sqlstr + " 	, IsNULL(T2.baditemetcW,0) as baditemetcW "
		sqlstr = sqlstr + " 	, IsNULL(T2.baditemetcZ,0) as baditemetcZ "
        'sqlstr = sqlstr + " 	, IsNULL(T2.erritem70,0) as erritem70 "
        'sqlstr = sqlstr + " 	, IsNULL(T2.erritem80,0) as erritem80 "
        sqlstr = sqlstr + " 	, IsNULL(T2.erritem90M,0) as erritem90M "
        sqlstr = sqlstr + " 	, IsNULL(T2.erritem90W,0) as erritem90W "
        sqlstr = sqlstr + " 	, IsNULL(T2.erritem90U,0) as erritem90U "
		sqlstr = sqlstr + " 	, IsNULL(T2.erritem90Z,0) as erritem90Z "
        sqlstr = sqlstr + " 	, IsNULL(T2.erritemetcM,0) as erritemetcM "
        sqlstr = sqlstr + " 	, IsNULL(T2.erritemetcW,0) as erritemetcW "
		sqlstr = sqlstr + " 	, IsNULL(T2.erritemetcZ,0) as erritemetcZ "
        sqlstr = sqlstr + " FROM [db_user].[dbo].tbl_user_c c with (nolock)"
        sqlstr = sqlstr + " 	LEFT JOIN ( "
        sqlstr = sqlstr + " 		SELECT "
        sqlstr = sqlstr + " 			i.makerid "
        sqlstr = sqlstr + " 			, sum(s.errbaditemno) as OnBadCnt "
        sqlstr = sqlstr + " 			, sum(s.errrealcheckno) as OnErrCnt "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='10' and " + CStr(sqlMWDivON) + " = 'M'  THEN s.errbaditemno ELSE 0 END) as baditem10M "
    	sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='10' and " + CStr(sqlMWDivON) + " = 'W'  THEN s.errbaditemno ELSE 0 END) as baditem10W "
		sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='10' and " + CStr(sqlMWDivON) + " = 'U'  THEN s.errbaditemno ELSE 0 END) as baditem10U "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='10' and " + CStr(sqlMWDivON) + " = 'Z'  THEN s.errbaditemno ELSE 0 END) as baditem10Z "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='10' and " + CStr(sqlMWDivON) + " = 'M'  THEN s.errrealcheckno ELSE 0 END) as erritem10M "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='10' and " + CStr(sqlMWDivON) + " = 'W'  THEN s.errrealcheckno ELSE 0 END) as erritem10W "
		sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='10' and " + CStr(sqlMWDivON) + " = 'U'  THEN s.errrealcheckno ELSE 0 END) as erritem10U "
		sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='10' and " + CStr(sqlMWDivON) + " = 'Z'  THEN s.errrealcheckno ELSE 0 END) as erritem10Z "
        sqlstr = sqlstr + " 		FROM "
        if (FRectDatetype="yyyymm") then
            sqlstr = sqlstr + " 			[db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s with (nolock)"
        else
            sqlstr = sqlstr + " 			[db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
			sqlstr = sqlstr + " 		LEFT JOIN [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary ms with (nolock)"
			sqlstr = sqlstr + " 			ON s.itemgubun=ms.itemgubun "
			sqlstr = sqlstr + " 				and s.itemid=ms.itemid "
			sqlstr = sqlstr + " 				and s.itemoption=ms.itemoption "
			sqlstr = sqlstr + " 				and ms.yyyymm='"&FRectYYYYMM&"' "
        end if
        sqlstr = sqlstr + " 			JOIN [db_item].[dbo].tbl_item i with (nolock)"
        sqlstr = sqlstr + " 			on "
        sqlstr = sqlstr + " 				1 = 1 "
        if (FRectDatetype="yyyymm") then
        sqlstr = sqlstr + " 				and s.yyyymm='"&FRectYYYYMM&"'"
        end if
        sqlstr = sqlstr + " 				and s.itemgubun='10' "
        sqlstr = sqlstr + " 				and s.itemid=i.itemid "
		if (FRectCenterMWDiv <> "") then
			sqlstr = sqlstr + " 			LEFT JOIN [db_shop].[dbo].tbl_shop_item si with (nolock)"
			sqlstr = sqlstr + " 			on "
			sqlstr = sqlstr + " 				1 = 1 "
			sqlstr = sqlstr + " 				and s.itemgubun=si.itemgubun "
			sqlstr = sqlstr + " 				and s.itemid=si.shopitemid "
			sqlstr = sqlstr + " 				and s.itemoption=si.itemoption "
		end if
        sqlstr = sqlstr + " 		WHERE "

		if (FRectSearchType = "bad") then
			sqlstr = sqlstr + " 			s.errbaditemno<>0 "
		else
			sqlstr = sqlstr + " 			s.errrealcheckno<>0 "
		end if

		if (FRectDatetype="yyyymm") then
			if (FRectMWDiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(i.mwdiv, 'Z') = '" + CStr(FRectMWDiv) + "' "
			end if
			if (FRectlastmwdiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(s.lastmwdiv, 'Z') = '" + CStr(FRectlastmwdiv) + "' "
			end if
		else
			if (FRectMWDiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(i.mwdiv, 'Z') = '" + CStr(FRectMWDiv) + "' "
			end if
			if (FRectlastmwdiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(IsNull(ms.lastmwdiv,i.mwdiv), 'Z') = '" + CStr(FRectlastmwdiv) + "' "
			end if
		end if

		if (FRectCenterMWDiv <> "") then
			sqlstr = sqlstr + " 		and IsNull(si.centermwdiv, 'Z') = '" & FRectCenterMWDiv & "' "
		end if

		if FRectItemGubun<>"" then
			if (FRectItemGubun = "OFF") then
					sqlstr = sqlstr + " and s.itemgubun <> '10' "
			else
					sqlstr = sqlstr + " and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
			end if
		end if

		if (FRectSellYN <> "") then
				sqlstr = sqlstr + " 	and IsNull(i.sellyn,'Y') = '" + CStr(FRectSellYN) + "' "
		end if

		if (FRectOnlyIsUsing <> "") then
				sqlstr = sqlstr + " 	and IsNull(i.isusing, 'Y') = '" + CStr(FRectOnlyIsUsing) + "' "
		end if

        sqlstr = sqlstr + " 		GROUP BY "
        sqlstr = sqlstr + " 			i.makerid "
        sqlstr = sqlstr + " 	) T1 "
        sqlstr = sqlstr + " 	on "
        sqlstr = sqlstr + " 		c.userid=T1.makerid "
        sqlstr = sqlstr + " 	LEFT JOIN ( "
        sqlstr = sqlstr + " 		SELECT "
        sqlstr = sqlstr + " 			i.makerid "
        sqlstr = sqlstr + " 			, sum(s.errbaditemno) as OffBadCnt "
        sqlstr = sqlstr + " 			, sum(s.errrealcheckno) as OffErrCnt "
        'sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='70' THEN s.errbaditemno ELSE 0 END) as baditem70 "
        'sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='80' THEN s.errbaditemno ELSE 0 END) as baditem80 "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='90' and " + CStr(sqlMWDivOFF) + " = 'M' THEN s.errbaditemno ELSE 0 END) as baditem90M "
    	sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='90' and " + CStr(sqlMWDivOFF) + " = 'W' THEN s.errbaditemno ELSE 0 END) as baditem90W "
		sqlstr = sqlstr + " 			, 0 as baditem90U "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='90' and " + CStr(sqlMWDivOFF) + " not in ('M', 'W') THEN s.errbaditemno ELSE 0 END) as baditem90Z "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun<>'90' and " + CStr(sqlMWDivOFF) + " = 'M' THEN s.errbaditemno ELSE 0 END) as baditemetcM "
    	sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun<>'90' and " + CStr(sqlMWDivOFF) + " = 'W' THEN s.errbaditemno ELSE 0 END) as baditemetcW "
		sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun<>'90' and " + CStr(sqlMWDivOFF) + " not in ('M', 'W') THEN s.errbaditemno ELSE 0 END) as baditemetcZ "
        'sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='70' THEN s.errrealcheckno ELSE 0 END) as erritem70 "
        'sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='80' THEN s.errrealcheckno ELSE 0 END) as erritem80 "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='90' and " + CStr(sqlMWDivOFF) + " = 'M' THEN s.errrealcheckno ELSE 0 END) as erritem90M "
    	sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='90' and " + CStr(sqlMWDivOFF) + " = 'W' THEN s.errrealcheckno ELSE 0 END) as erritem90W "
		sqlstr = sqlstr + " 			, 0 as erritem90U "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun='90' and " + CStr(sqlMWDivOFF) + " not in ('M', 'W') THEN s.errrealcheckno ELSE 0 END) as erritem90Z "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun<>'90' and " + CStr(sqlMWDivOFF) + " = 'M' THEN s.errrealcheckno ELSE 0 END) as erritemetcM "
    	sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun<>'90' and " + CStr(sqlMWDivOFF) + " = 'W' THEN s.errrealcheckno ELSE 0 END) as erritemetcW "
        sqlstr = sqlstr + " 			, sum(CASE WHEN s.itemgubun<>'90' and " + CStr(sqlMWDivOFF) + " not in ('M', 'W') THEN s.errrealcheckno ELSE 0 END) as erritemetcZ "
        sqlstr = sqlstr + " 		FROM "
        if (FRectDatetype="yyyymm") then
            sqlstr = sqlstr + " 			[db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s with (nolock)"
        else
            sqlstr = sqlstr + " 			[db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
        end if
        sqlstr = sqlstr + " 			JOIN [db_shop].[dbo].tbl_shop_item i with (nolock)"
        sqlstr = sqlstr + " 			on "
        sqlstr = sqlstr + " 				1 = 1 "
        if (FRectDatetype="yyyymm") then
        sqlstr = sqlstr + " 				and s.yyyymm='"&FRectYYYYMM&"'"
        end if
        sqlstr = sqlstr + " 				and s.itemgubun<>'10' "
        sqlstr = sqlstr + " 				and s.itemgubun=i.itemgubun "
        sqlstr = sqlstr + " 				and s.itemid=i.shopitemid "
        sqlstr = sqlstr + " 				and s.itemoption=i.itemoption "
		if (FRectCenterMWDiv <> "") then
			sqlstr = sqlstr + " 				and IsNull(i.centermwdiv, 'Z') = '" & FRectCenterMWDiv & "' "
		end if
        sqlstr = sqlstr + " 		WHERE "

		if (FRectSearchType = "bad") then
			sqlstr = sqlstr + " 			s.errbaditemno<>0 "
		else
			sqlstr = sqlstr + " 			s.errrealcheckno<>0 "
		end if

		if (FRectDatetype="yyyymm") then
			if (FRectMWDiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(i.centermwdiv, 'Z') = '" + CStr(FRectMWDiv) + "' "
			end if
			if (FRectlastmwdiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(s.lastmwdiv, 'Z') = '" + CStr(FRectlastmwdiv) + "' "
			end if
		else
			if (FRectMWDiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(i.centermwdiv, 'Z') = '" + CStr(FRectMWDiv) + "' "
			end if
			if (FRectlastmwdiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(i.centermwdiv, 'Z') = '" + CStr(FRectlastmwdiv) + "' "
			end if
		end if

		if FRectItemGubun<>"" then
			if (FRectItemGubun = "OFF") then
					sqlstr = sqlstr + " and s.itemgubun <> '10' "
			else
					sqlstr = sqlstr + " and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
			end if
		end if

		''if (FRectSellYN <> "") then
		''		sqlstr = sqlstr + " 	and IsNull(i.sellyn,'Y') = '" + CStr(FRectSellYN) + "' "
		''end if

		if (FRectOnlyIsUsing <> "") then
				sqlstr = sqlstr + " 	and IsNull(i.isusing, 'Y') = '" + CStr(FRectOnlyIsUsing) + "' "
		end if

        sqlstr = sqlstr + " 		GROUP BY "
        sqlstr = sqlstr + " 			i.makerid "
        sqlstr = sqlstr + " 	) T2 "
        sqlstr = sqlstr + " 	on "
        sqlstr = sqlstr + " 		c.userid=T2.makerid "
        sqlstr = sqlstr + " 	left join db_partner.dbo.tbl_partner p with (nolock)"
        sqlstr = sqlstr + " 	on "
        sqlstr = sqlstr + " 		c.userid = p.id "
        sqlstr = sqlstr + " 	left join db_partner.dbo.tbl_partner_group g with (nolock)"
        sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		p.groupid = g.groupid "

        sqlstr = sqlstr + " WHERE "
        sqlstr = sqlstr + " 	1 = 1 "
        sqlstr = sqlstr + " 	and (T1.makerid is Not NULL or T2.makerid is Not NULL) "

        if (FRectPurchaseType <> "") then
        	sqlstr = sqlstr + " 	and IsNull(p.purchasetype, 1) = " & FRectPurchaseType & " "
        end if

		if (FRectMakerUseYN <> "") then
			sqlstr = sqlstr + " 	and IsNull(c.isusing, 'Y') = '" + CStr(FRectMakerUseYN) + "' "
		end if

        sqlstr = sqlstr + " ORDER BY "

		if (FRectSearchType = "bad") then
			sqlstr = sqlstr + " 	OnBadCnt, offBadCnt, c.userid "
		else
			sqlstr = sqlstr + " 	OnErrCnt, offErrCnt, c.userid "
		end if

		'response.write sqlstr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CBadOrErritemBrandGroupItem

				FItemList(i).FMakerid       = rsget("makerid")
				FItemList(i).Fuseyn       	= rsget("useyn")
				FItemList(i).Fmakername     = db2html(rsget("makername"))
				FItemList(i).Fcompany_no     	= rsget("company_no")
				FItemList(i).Fcompany_name    = db2html(rsget("company_name"))

				if (FRectSearchType = "bad") then
					FItemList(i).FOnCnt      = rsget("OnBadCnt")
					FItemList(i).FOffCnt     = rsget("OffBadCnt")
					FItemList(i).Fitem10M    = rsget("baditem10M")
					FItemList(i).Fitem10W    = rsget("baditem10W")
					FItemList(i).Fitem10U    = rsget("baditem10U")
					FItemList(i).Fitem10Z    = rsget("baditem10Z")
					'FItemList(i).Fitem70     = rsget("baditem70")
					'FItemList(i).Fitem80     = rsget("baditem80")
					FItemList(i).Fitem90M    = rsget("baditem90M")
					FItemList(i).Fitem90W    = rsget("baditem90W")
					FItemList(i).Fitem90U    = rsget("baditem90U")
					FItemList(i).Fitem90Z    = rsget("baditem90Z")
					FItemList(i).FitemetcM    = rsget("baditemetcM")
					FItemList(i).FitemetcW    = rsget("baditemetcW")
					FItemList(i).FitemetcZ    = rsget("baditemetcZ")
				else
					FItemList(i).FOnCnt      = rsget("OnErrCnt")
					FItemList(i).FOffCnt     = rsget("OffErrCnt")
					FItemList(i).Fitem10M    = rsget("erritem10M")
					FItemList(i).Fitem10W    = rsget("erritem10W")
					FItemList(i).Fitem10U    = rsget("erritem10U")
					FItemList(i).Fitem10Z    = rsget("erritem10Z")
					'FItemList(i).Fitem70     = rsget("erritem70")
					'FItemList(i).Fitem80     = rsget("erritem80")
					FItemList(i).Fitem90M    = rsget("erritem90M")
					FItemList(i).Fitem90W    = rsget("erritem90W")
					FItemList(i).Fitem90U    = rsget("erritem90U")
					FItemList(i).Fitem90Z    = rsget("erritem90Z")
					FItemList(i).FitemetcM    = rsget("erritemetcM")
					FItemList(i).FitemetcW    = rsget("erritemetcW")
					FItemList(i).FitemetcZ    = rsget("erritemetcZ")
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	public sub GetDailyErrItemListByBrand()
		''온라인 오프라인 따로 구분 ==> 오프 상품은 거래구분이 다름..
		dim sqlstr, i

		if ((FRectMakerid = "") and (FRectItemID <> "")) then
    		sqlstr = " select top 1 i.makerid "
    		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i "
    		sqlstr = sqlstr + " where i.itemid = " + CStr(FRectItemID) + " "
    		rsget.Open sqlStr,dbget,1
    		if  not rsget.EOF  then
    		        FRectMakerid = rsget("makerid")
    		end if
    		rsget.close
		end if

		sqlstr = "select top 1000 s.itemgubun, s.itemid, s.itemoption, "
		sqlstr = sqlstr + " sum(s.errcsno) as errcsno, sum(s.errbaditemno) as errbaditemno, sum(s.errrealcheckno) as errrealcheckno, sum(s.erretcno) as erretcno, sum(s.toterrno) as toterrno,"
		sqlstr = sqlstr + " i.makerid, i.itemname, i.deliverytype, i.sellcash , i.buycash, i.mwdiv,"
		sqlstr = sqlstr + " f.makerid as shopmakerid, f.shopitemname, f.shopitemoptionname, f.shopitemprice , f.shopsuplycash, f.centermwdiv,"
		sqlstr = sqlstr + " IsNULL(v.optionname,'') as codeview, i.smallimage, f.offimgsmall "
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_erritem_daily_summary s"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.itemid=i.itemid"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item f on s.itemgubun<>'10' and s.itemgubun=f.itemgubun and s.itemid=f.shopitemid and s.itemoption=f.itemoption"
		sqlstr = sqlstr + " where 1 = 1 "
		sqlstr = sqlstr + " and (i.makerid='" + CStr(FRectMakerid) + "' or f.makerid='" + CStr(FRectMakerid) + "')"
'		sqlstr = sqlstr + " and ("
'		sqlstr = sqlstr + " 	 (s.itemgubun='10' and i.makerid='" + CStr(FRectMakerid) + "')"
'		sqlstr = sqlstr + " 	 or (s.itemgubun<>'10' and f.makerid='" + CStr(FRectMakerid) + "')"
'		sqlstr = sqlstr + " 	)"
        sqlstr = sqlstr + " group by s.itemgubun, s.itemid, s.itemoption, i.makerid, i.itemname, i.deliverytype, i.sellcash , i.buycash, i.mwdiv, IsNULL(v.optionname,'') "
		sqlstr = sqlstr + " ,f.makerid , f.shopitemname, f.shopitemoptionname, f.shopitemprice , f.shopsuplycash, f.centermwdiv, i.smallimage, f.offimgsmall"

		if (FRectSearchType = "bad") then
			sqlstr = sqlstr + " having sum(s.errbaditemno)<>0"
		elseif (FRectSearchType = "err") then
			sqlstr = sqlstr + " having sum(s.errrealcheckno)<>0"
		else
			sqlstr = sqlstr + " having sum(s.errbaditemno)<>0"
		end if

		sqlstr = sqlstr + " order by s.itemgubun, s.itemid, s.itemoption "

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CErritemDailyItem

				FItemList(i).Fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Ferrcsno        = rsget("errcsno")
				FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).Ferretcno       = rsget("erretcno")
				FItemList(i).Ftoterrno       = rsget("toterrno")

				if FItemList(i).Fitemgubun="10" then
					FItemList(i).FItemName		= db2html(rsget("itemname"))
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).Fsellcash		= rsget("sellcash")
					FItemList(i).Fbuycash		= rsget("buycash")
					FItemList(i).Fmwdiv		= rsget("mwdiv")
					FItemList(i).Fdeliverytype	= rsget("deliverytype")
					FItemList(i).FItemOptionName= db2html(rsget("codeview"))
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				else
					FItemList(i).FItemName		= db2html(rsget("shopitemname"))
					FItemList(i).FMakerid		= rsget("shopmakerid")
					FItemList(i).Fsellcash		= rsget("shopitemprice")
					FItemList(i).Fbuycash		= rsget("shopsuplycash")
					FItemList(i).Fmwdiv		= rsget("centermwdiv")
					''FItemList(i).Fdeliverytype	= rsget("deliverytype")
					FItemList(i).FItemOptionName= db2html(rsget("shopitemoptionname"))
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimgsmall")
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	' /admin/stock/badorerritem_act_list.asp
	public sub GetBadOrErrItemListByBrand()
		''온라인 오프라인 따로 구분 ==> 오프 상품은 거래구분이 다름..
		dim sqlstr, i, sqlMWDiv

		if (FPageSize = "") Then
			FPageSize = 100
		End If

		if (FRectDatetype="yyyymm") then
			sqlMWDiv  = " IsNull(s.lastmwdiv, 'Z') "
		else
			sqlMWDiv  = " IsNull(IsNull(IsNull(ms.lastmwdiv,i.mwdiv), f.centermwdiv), 'Z') "
		end if

		if ((FRectMakerid = "") and (FRectItemID <> "")) then
    		sqlstr = " select top 1 i.makerid "
    		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i with (nolock)"
    		sqlstr = sqlstr + " where i.itemid = " + CStr(FRectItemID) + " "

			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    		if  not rsget.EOF  then
    		        FRectMakerid = rsget("makerid")
    		end if
    		rsget.close
		end if

		sqlstr = " select top " & FPageSize & " "
		sqlstr = sqlstr + " 	s.itemgubun, s.itemid, s.itemoption, "
		sqlstr = sqlstr + " 	i.makerid, i.itemname, i.deliverytype, i.orgprice as sellcash , i.buycash, IsNull(IsNull(i.mwdiv, f.centermwdiv), 'Z') as mwdiv, i.sellyn, i.isusing, "
		sqlstr = sqlstr + " 	f.makerid as shopmakerid, f.shopitemname, f.shopitemoptionname, f.orgsellprice as shopitemprice , f.shopsuplycash"
		sqlstr = sqlstr + " 	, 'Y' as shopsellyn, f.isusing as shopisusing, "	' , f.centermwdiv
		sqlstr = sqlstr + " 	IsNULL(v.optionname,'') as codeview, "
		sqlstr = sqlstr + " 	i.smallimage, f.offimgsmall, "
		''sqlstr = sqlstr + " 	IsNull(s.lastbuyprice, 0) as lastbuyprice, "
		sqlstr = sqlstr + " 	0 as lastbuyprice, IsNull(u.centermwdiv, '') as centermwdiv"
		sqlstr = sqlstr + " 	, IsNull(s.errbaditemno, 0) as errbaditemno, "			'// 불량
		sqlstr = sqlstr + " 	IsNull(s.errrealcheckno, 0) as errrealcheckno, "		'// 오차
		sqlstr = sqlstr + " 	IsNull(s.realstock, 0) as realstock, lastIpgodate, " + CStr(sqlMWDiv) + " as lastmwdiv "
		sqlstr = sqlstr + " from "

		if (FRectDatetype="yyyymm") then
			sqlstr = sqlstr + " 	[db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s with (nolock)"
		else
			sqlstr = sqlstr + " 	[db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
			sqlstr = sqlstr + " 	left join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary  as ms with (nolock)"
			sqlstr = sqlstr + " 		on s.itemgubun=ms.itemgubun and s.itemid=ms.itemid and s.itemoption=ms.itemoption and ms.yyyymm = '" & left(date,7) & "'"
		end if

		sqlstr = sqlstr + " 	left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		1 = 1 "
		sqlstr = sqlstr + " 		and s.itemgubun='10' "
		sqlstr = sqlstr + " 		and s.itemid=i.itemid "
		sqlstr = sqlstr + " 	left join [db_item].[dbo].tbl_item_option v with (nolock)"
		sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		1 = 1 "
		sqlstr = sqlstr + " 		and s.itemid=v.itemid "
		sqlstr = sqlstr + " 		and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " 	left join [db_shop].[dbo].tbl_shop_item f with (nolock)"
		sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		1 = 1 "
		sqlstr = sqlstr + " 		and s.itemgubun<>'10' "
		sqlstr = sqlstr + " 		and s.itemgubun=f.itemgubun "
		sqlstr = sqlstr + " 		and s.itemid=f.shopitemid "
		sqlstr = sqlstr + " 		and s.itemoption=f.itemoption "
		sqlstr = sqlstr + " 	left join [db_shop].[dbo].tbl_shop_item u with (nolock)"		'// 10 업체배송, 90 상품 => 센터매입구분 이용
		sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		1 = 1 "
		sqlstr = sqlstr + " 		and s.itemgubun=u.itemgubun "
		sqlstr = sqlstr + " 		and s.itemid=u.shopitemid "
		sqlstr = sqlstr + " 		and s.itemoption=u.itemoption "
        sqlstr = sqlstr + " 	left join db_partner.dbo.tbl_partner p with (nolock)"
        sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		IsNull(i.makerid, f.makerid) = p.id "
        sqlstr = sqlstr + " 	left join [db_user].[dbo].tbl_user_c c with (nolock)"
        sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		IsNull(i.makerid, f.makerid) = c.userid "
        sqlstr = sqlstr + " 	left join db_partner.dbo.tbl_partner_group g with (nolock)"
        sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		p.groupid = g.groupid "
		sqlstr = sqlstr + " where "
		sqlstr = sqlstr + " 	1 = 1 "

		if (FRectDatetype="yyyymm") then
			sqlstr = sqlstr + "		and s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		end if

		If (FRectMakerid <> "all") then
			sqlstr = sqlstr + " 	and IsNull(i.makerid, f.makerid) = '" + CStr(FRectMakerid) + "' "
		End If

        if (FRectPurchaseType <> "") then
        	sqlstr = sqlstr + " 	and IsNull(p.purchasetype, 1) = " & FRectPurchaseType & " "
        end if

		if (FRectMakerUseYN <> "") then
			sqlstr = sqlstr + " 	and IsNull(c.isusing, 'Y') = '" + CStr(FRectMakerUseYN) + "' "
		end if

		if FRectItemGubun<>"" then
			if (FRectItemGubun = "OFF") then
				sqlstr = sqlstr + " and s.itemgubun <> '10' "
			else
				sqlstr = sqlstr + " and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
			end if
		end if

		if (FRectSearchType = "bad") then
			sqlstr = sqlstr + " 	and s.errbaditemno <> 0 "
		else
			sqlstr = sqlstr + " 	and s.errrealcheckno <> 0 "
		end if

		if (FRectDatetype="yyyymm") then
			if (FRectMWDiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(IsNull(i.mwdiv, f.centermwdiv), 'Z') = '" + CStr(FRectMWDiv) + "' "
			end if
			if (FRectlastmwdiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(s.lastmwdiv, 'Z') = '" + CStr(FRectlastmwdiv) + "' "
			end if
		else
			if (FRectMWDiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(IsNull(i.mwdiv, f.centermwdiv), 'Z') = '" + CStr(FRectMWDiv) + "' "
			end if
			if (FRectlastmwdiv <> "") then
				sqlstr = sqlstr + " 	and IsNull(IsNull(IsNull(ms.lastmwdiv,i.mwdiv), f.centermwdiv), 'Z') = '" + CStr(FRectlastmwdiv) + "' "
			end if
		end if

		if (FRectCenterMWDiv <> "") then
			sqlstr = sqlstr + " 	and IsNull(u.centermwdiv, 'Z') = '" + CStr(FRectCenterMWDiv) + "' "
		end if

		if (FRectSellYN <> "") and (FRectItemGubun = "10") then
			sqlstr = sqlstr + " 	and IsNull(i.sellyn,'Y') = '" + CStr(FRectSellYN) + "' "
		end if

		if (FRectOnlyIsUsing <> "") then
			sqlstr = sqlstr + " 	and IsNull(i.isusing, f.isusing) = '" + CStr(FRectOnlyIsUsing) + "' "
		end if

		sqlstr = sqlstr + " order by "
		sqlstr = sqlstr + " 	s.itemgubun, s.itemid, s.itemoption "

		'response.write sqlstr & "<br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CBadOrErrItemItem

				FItemList(i).Fitemgubun      	= rsget("itemgubun")
				FItemList(i).Fitemid         	= rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")

				if (FRectSearchType = "bad") then
					FItemList(i).Fregitemno		= rsget("errbaditemno")
				else
					FItemList(i).Fregitemno		= rsget("errrealcheckno")
				end if

				FItemList(i).Frealstock        	= rsget("realstock")
				FItemList(i).Fmwdiv				= rsget("mwdiv")
				FItemList(i).Fcentermwdiv		= rsget("centermwdiv")

				FItemList(i).FlastIpgoDate		= rsget("lastIpgoDate")
				FItemList(i).flastmwdiv		= rsget("lastmwdiv")

				if FItemList(i).Fitemgubun="10" then
					FItemList(i).FItemName			= db2html(rsget("itemname"))
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).Fsellcash			= rsget("sellcash")
					FItemList(i).Fbuycash			= rsget("buycash")
					''FItemList(i).Fmwdiv				= rsget("mwdiv")
					FItemList(i).Fdeliverytype		= rsget("deliverytype")
					FItemList(i).FItemOptionName	= db2html(rsget("codeview"))
					FItemList(i).Fimgsmall 			= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
					FItemList(i).Fsellyn			= rsget("sellyn")
					FItemList(i).Fisusing			= rsget("isusing")
				else
					FItemList(i).FItemName			= db2html(rsget("shopitemname"))
					FItemList(i).FMakerid			= rsget("shopmakerid")
					FItemList(i).Fsellcash			= rsget("shopitemprice")
					FItemList(i).Fbuycash			= rsget("shopsuplycash")

					if (FItemList(i).Fbuycash = 0) then
						FItemList(i).Fbuycash		= rsget("lastbuyprice")
					end if

					''FItemList(i).Fmwdiv				= rsget("centermwdiv")
					''FItemList(i).Fdeliverytype	= rsget("deliverytype")
					FItemList(i).FItemOptionName= db2html(rsget("shopitemoptionname"))
					FItemList(i).Fimgsmall 			= "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimgsmall")
					FItemList(i).Fsellyn			= "Y"
					FItemList(i).Fisusing			= rsget("shopisusing")
				end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public sub GetCurrentItemStock
		dim sqlstr

		sqlstr = " select top 1 "
		sqlstr = sqlstr + " 	c.* "
		sqlstr = sqlstr + " 	, IsNull(o.DayForSellCount, 7) as DayForSellCount "
		sqlstr = sqlstr + " 	, IsNull(o.DayForSafeStock, 3) as DayForSafeStock "
		sqlstr = sqlstr + " 	, IsNull(o.DayForLeadTime, 2) as DayForLeadTime "
		sqlstr = sqlstr + " 	, IsNull(o.DayForMaxStock, 10) as DayForMaxStock, IsNull(c.itemgrade, 'Z') as itemgrade "
		sqlstr = sqlstr + " from "
		sqlstr = sqlstr + " 	[db_summary].[dbo].tbl_current_logisstock_summary c "
		sqlstr = sqlstr + " 	left join db_item.dbo.tbl_item_option_stock o "
		sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		1 = 1 "
		sqlstr = sqlstr + " 		and c.itemgubun = o.itemgubun "
		sqlstr = sqlstr + " 		and c.itemid = o.itemid "
		sqlstr = sqlstr + " 		and c.itemoption = o.itemoption "
		sqlstr = sqlstr + " where "
		sqlstr = sqlstr + " 	1 = 1 "
		sqlstr = sqlstr + " 	and c.itemgubun = '" + FRectItemGubun + "' "
		sqlstr = sqlstr + " 	and c.itemid = " + CStr(FRectItemID) + " "
		sqlstr = sqlstr + " 	and c.itemoption = '" + CStr(FRectItemOption) + "' "

		rsget.Open sqlStr,dbget,1
        FResultCount = rsget.RecordCount

		set FOneItem = new CCurrentStockItem
		if Not rsget.Eof then

			FOneItem.Fitemgubun     = rsget("itemgubun")
			FOneItem.Fitemid        = rsget("itemid")
			FOneItem.Fitemoption    = rsget("itemoption")
			FOneItem.Fipgono        = rsget("ipgono")
			FOneItem.Freipgono      = rsget("reipgono")
			FOneItem.Ftotipgono     = rsget("totipgono")
			FOneItem.Foffchulgono   = rsget("offchulgono")
			FOneItem.Foffrechulgono = rsget("offrechulgono")
			FOneItem.Fetcchulgono   = rsget("etcchulgono")
			FOneItem.Fetcrechulgono = rsget("etcrechulgono")
			FOneItem.Ftotchulgono   = rsget("totchulgono")
			FOneItem.Fsellno        = rsget("sellno")
			FOneItem.Fresellno      = rsget("resellno")
			FOneItem.Ftotsellno     = rsget("totsellno")
			FOneItem.Ferrcsno       = rsget("errcsno")
			FOneItem.Ferrbaditemno  = rsget("errbaditemno")
			FOneItem.Ferrrealcheckno= rsget("errrealcheckno")
			FOneItem.Ferretcno      = rsget("erretcno")
			FOneItem.Ftoterrno      = rsget("toterrno")
			FOneItem.Ftotsysstock   = rsget("totsysstock")
			FOneItem.Favailsysstock = rsget("availsysstock")
			FOneItem.Frealstock     = rsget("realstock")
			FOneItem.Fsell7days     = rsget("sell7days")
			FOneItem.Foffchulgo7days= rsget("offchulgo7days")
			FOneItem.Fipkumdiv5     = rsget("ipkumdiv5")
			FOneItem.Fipkumdiv4     = rsget("ipkumdiv4")
			FOneItem.Fipkumdiv2     = rsget("ipkumdiv2")
			FOneItem.Foffconfirmno  = rsget("offconfirmno")
			FOneItem.Foffjupno      = rsget("offjupno")
			FOneItem.Frequireno     = rsget("requireno")
			FOneItem.FrequireMaxno  = rsget("requireMaxno")
			FOneItem.Fshortageno    = rsget("shortageno")
			FOneItem.Fpreorderno    = rsget("preorderno")
			FOneItem.Fpreordernofix = rsget("preordernofix")
			FOneItem.Foffsellno		= rsget("offsellno")
			FOneItem.Fmaxsellday    = rsget("maxsellday")
			FOneItem.Fimgsmall      = rsget("imgsmall")
			FOneItem.Fregdate       = rsget("regdate")
			FOneItem.Flastupdate    = rsget("lastupdate")
			FOneItem.FDayForSellCount  = rsget("DayForSellCount")
			FOneItem.FDayForSafeStock  = rsget("DayForSafeStock")
			FOneItem.FDayForLeadTime   = rsget("DayForLeadTime")
			FOneItem.FDayForMaxStock   = rsget("DayForMaxStock")

            FOneItem.Fitemgrade   = rsget("itemgrade")

			if IsNull(FOneItem.FrequireMaxno) then
				FOneItem.FrequireMaxno = FOneItem.Frequireno * 2
			end if

		end if
		rsget.Close
	end sub

	public sub GetCurrentAgvItemStock
		dim sqlstr

        '// , itemgubun, itemid, itemoption, skuCd, agvstock, regdate, lastupdate, warehouseCd, totsysstock, errrealcheckno, bulkstock, lastbulkstockdate
		sqlstr = " select top 1 a.* "
		sqlstr = sqlstr + " from "
		sqlstr = sqlstr + " 	[db_summary].[dbo].[tbl_current_agvstock_summary] a "
		sqlstr = sqlstr + " where "
		sqlstr = sqlstr + " 	1 = 1 "
		sqlstr = sqlstr + " 	and a.itemgubun = '" + FRectItemGubun + "' "
		sqlstr = sqlstr + " 	and a.itemid = " + CStr(FRectItemID) + " "
		sqlstr = sqlstr + " 	and a.itemoption = '" + CStr(FRectItemOption) + "' "
        ''response.write sqlstr

		rsget.Open sqlStr,dbget,1
        FResultCount = rsget.RecordCount

		set FOneItem = new CCurrentStockItem
		if Not rsget.Eof then

			FOneItem.Fitemgubun     = rsget("itemgubun")
			FOneItem.Fitemid        = rsget("itemid")
			FOneItem.Fitemoption    = rsget("itemoption")

            FOneItem.FwarehouseCd   = rsget("warehouseCd")
            FOneItem.Fagvstock      = rsget("agvstock")

			if IsNull(FOneItem.FwarehouseCd) then
				FOneItem.FwarehouseCd = "BLK"
			end if

		end if
		rsget.Close
	end sub

	'//admin/stock/brandcurrentstock.asp
	' 이 펑션을 수정할경우 GetCurrentStockByOfflineBrand_notpaping 도 반드시 같이 수정해 주세요.
	public sub GetCurrentStockByOfflineBrand
		dim sqlstr, i, sqlsearch, stockFieldName

		if (FRectItemGubun<>"") then
		    sqlsearch = sqlsearch + " and s.itemgubun='" + CStr(FRectItemGubun) + "'"
		end if

		if FRectMakerid<>"" then
		    sqlsearch = sqlsearch + " and i.makerid='" + CStr(FRectMakerid) + "'"
		end if
		if FRectOnlyIsUsing<>"" then
			sqlsearch = sqlsearch + " and i.isusing='" + FRectOnlyIsUsing + "'"
		end if

		if (FRectItemIdArr<>"") then
            sqlsearch = sqlsearch + " and i.shopitemid in (" & FRectItemIdArr & ")"
        end if

        if (FRectItemName<>"") then
            sqlsearch = sqlsearch + " and i.shopitemname like '%" & FRectItemName & "%'"
        end if

        if FRectCenterMWDiv="N" then
            sqlsearch = sqlsearch + " and i.centermwdiv is NULL"
        elseif FRectCenterMWDiv<>"" then
            sqlsearch = sqlsearch + " and i.centermwdiv='" + FRectCenterMWDiv + "'"
        end if

        if (FRectCD1<>"") then
            sqlsearch = sqlsearch + " and i.catecdl='" & FRectCD1 & "'"
        end if

        if (FRectCD2<>"") then
            sqlsearch = sqlsearch + " and i.catecdm='" & FRectCD2 & "'"
        end if

        if (FRectCD3<>"") then
            sqlsearch = sqlsearch + " and i.catecdn='" & FRectCD3 & "'"
        end if

'		if FRectlimitrealstock="1UP" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1"
'		elseif FRectlimitrealstock="0DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 0"
'		elseif FRectlimitrealstock="20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 20"
'		elseif FRectlimitrealstock="1UP20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1 and s.realstock <= 20"
'		end if

		''XXXX 시스템재고로 변경 2014/08/01
		''두가지 선택 2014-11-04
		stockFieldName = "totsysstock"
		if (FRectStockType = "real") then
			stockFieldName = "realstock"
		end if
		if FRectlimitrealstock="1UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1"
		elseif FRectlimitrealstock="0DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 0"
		elseif FRectlimitrealstock="20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock="1UP20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock = "20UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 20"
		end If

		If (FRectStartDate <> "") And IsNumeric(FRectStartDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) >= " & FRectStartDate & " "
		End If

		If (FRectEndDate <> "") And IsNumeric(FRectEndDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) <= " & FRectEndDate & " "
		End If

        if FRectRackCode <> "" then
            sqlsearch = sqlsearch & " and Left(os.rackcodeByOption, " & Len(FRectRackCode) & ") = '" & FRectRackCode & "' "
        end if

        if FRectBulkStockGubun <> "" then
            if (FRectBulkStockGubun = "nul") then
                '// 입력이전
                sqlsearch = sqlsearch & " and DateDiff(day, IsNull(agvs.lastbulkstockdate, DateAdd(day, -30, getdate())), getdate()) > 5 "
            elseif (FRectBulkStockGubun = "err") then
                '// 벌크오차 있음
                sqlsearch = sqlsearch & " and DateDiff(day, IsNull(agvs.lastbulkstockdate, DateAdd(day, -30, getdate())), getdate()) <= 5 "
                sqlsearch = sqlsearch & " and (s.realstock + s.ipkumdiv5 + s.offconfirmno) <> (IsNull(agvs.bulkstock, 0) + IsNull(agvs.agvstock, 0)) "
            end if
        end if

        sqlstr = "select count(s.itemid) as CNT"
        sqlstr = sqlstr + " from [db_shop].[dbo].tbl_shop_item i with (nolock)"
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
        sqlstr = sqlstr + " 	on s.itemid = i.shopitemid and s.itemgubun = i.itemgubun "
        sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_current_agvstock_summary] agvs with (nolock) "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and agvs.itemgubun = i.itemgubun "
        sqlStr = sqlStr & " 		and agvs.itemid = i.shopitemid "
        sqlStr = sqlStr & " 		and agvs.itemoption = i.itemoption "

		If (FRectUseOffInfo = "") Then
			sqlstr = sqlstr + " and s.itemgubun <> '10' "
		Else
			sqlstr = sqlstr + " and s.itemoption = i.itemoption "
		End If

		If FRectStartDate <> "" Or FRectEndDate <> "" then
			sqlstr = sqlstr + " 	LEFT JOIN [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p with (nolock) ON 1 = 1 "
			sqlstr = sqlstr + " 		AND p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
			sqlstr = sqlstr + " 		AND s.itemgubun = p.itemgubun "
			sqlstr = sqlstr + " 		AND s.itemid = p.itemid "
			sqlstr = sqlstr + " 		AND s.itemoption = p.itemoption "
		End if

        if FRectRackCode <> "" then
            sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] os with (nolock) "
            sqlStr = sqlStr & " 		on "
            sqlStr = sqlStr & " 			1 = 1 "
            sqlStr = sqlStr & " 			and os.itemgubun = i.itemoption "
            sqlStr = sqlStr & " 			and os.itemid = i.shopitemid "
            sqlStr = sqlStr & " 			and os.itemoption = i.itemoption "
        end if

		sqlstr = sqlstr + " where 1=1 " & sqlsearch

		''response.write sqlstr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlstr = "select top " + CStr(FPageSize*FCurrPage)
		sqlstr = sqlstr & " s.itemgubun,s.itemid,s.itemoption,s.ipgono,s.reipgono,s.totipgono,s.offchulgono,s.offrechulgono,s.etcchulgono"
		sqlstr = sqlstr & " ,s.etcrechulgono,s.totchulgono,s.sellno,s.resellno,s.totsellno,s.errcsno,s.errbaditemno,s.errrealcheckno,s.erretcno"
		sqlstr = sqlstr & " ,s.toterrno,s.totsysstock,s.availsysstock,s.realstock,s.sell7days,s.offchulgo7days,s.ipkumdiv5,s.ipkumdiv4"
		sqlstr = sqlstr & " ,s.ipkumdiv2,s.offconfirmno,s.offjupno,s.requireno,s.shortageno,s.preorderno,s.preordernofix,s.offsellno"
		sqlstr = sqlstr & " ,s.maxsellday,s.imgsmall,s.itemoptionname,s.regdate,s.lastupdate,s.requireMaxno,s.itemgrade"
		sqlstr = sqlstr + " , i.makerid"
		sqlstr = sqlstr + " , i.shopitemname as itemname"
		sqlstr = sqlstr & " , i.shopitemprice as sellcash"
		sqlstr = sqlstr & " , i.shopsuplycash as buycash"
		sqlstr = sqlstr & " , i.isusing"
		sqlstr = sqlstr & " , '' as deliverytype"
		sqlstr = sqlstr & " ,'' as sellyn"
		sqlstr = sqlstr & " ,'' as limityn"
		sqlstr = sqlstr & " ,'' as limitno"
		sqlstr = sqlstr & " ,'' as limitsold"
		sqlstr = sqlstr & " , i.centermwdiv as mwdiv"
		sqlstr = sqlstr & " ,'' as danjongyn"
		sqlstr = sqlstr & " , IsNull(os.rackcodeByOption, (case when i.itemgubun='10' then ii.itemrackcode else i.offitemrackcode end)) as itemrackcode"
		sqlstr = sqlstr & " , IsNull(os.subRackcodeByOption, ISNULL(a.subitemrackcode,ii.itemrackcode)) AS subitemrackcode"
		sqlstr = sqlstr & " , i.shopitemoptionname as codeview"
		sqlstr = sqlstr & " ,'' AS optionusing"
		sqlstr = sqlstr & " ,'' as optlimityn"
		sqlstr = sqlstr & " ,'' as optlimitno"
		sqlstr = sqlstr & " ,'' as optlimitsold"
		sqlstr = sqlstr & " ,i.orgsellprice as orgprice"
		sqlstr = sqlstr & " , i.centermwdiv as centermwdiv"
		sqlstr = sqlstr & " , p.lastIpgoDate"
		sqlstr = sqlstr & " ,'' as publicBarcode"
		sqlstr = sqlstr & " ,0 as prevMonthSellCnt"
		sqlstr = sqlstr & " ,'' as OffMwMargin"
		sqlstr = sqlstr & " , agvs.agvstock"
		sqlstr = sqlstr & " , agvs.bulkstock"
		sqlstr = sqlstr & " , agvs.lastbulkstockdate"
		sqlstr = sqlstr & " ,'' as warehouseCd"
		sqlstr = sqlstr & " , 0 IDX"
		sqlstr = sqlstr & " , c.prtidx"
		sqlstr = sqlstr + " , i.offimgsmall"
		sqlstr = sqlstr + " , i.shopsuplycash as suplycash, d.defaultmargin, d.defaultsuplymargin "
        sqlstr = sqlstr + " from [db_shop].[dbo].tbl_shop_item i with (nolock)"
		sqlstr = sqlstr + " 	left join [db_shop].[dbo].tbl_shop_designer d with (nolock) "
		sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		1 = 1 "
		sqlstr = sqlstr + " 		and i.makerid=d.makerid "
		sqlstr = sqlstr + " 		and d.shopid='streetshop000' "
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
        sqlstr = sqlstr + " 	on s.itemid = i.shopitemid and s.itemgubun = i.itemgubun "

		If (FRectUseOffInfo = "") Then
			sqlstr = sqlstr + " and s.itemgubun <> '10' "
		Else
			sqlstr = sqlstr + " and s.itemoption = i.itemoption "
		End If

		sqlstr = sqlstr + " 	LEFT JOIN [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p with (nolock) ON 1 = 1 "
		sqlstr = sqlstr + " 		AND p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
		sqlstr = sqlstr + " 		AND s.itemgubun = p.itemgubun "
		sqlstr = sqlstr + " 		AND s.itemid = p.itemid "
		sqlstr = sqlstr + " 		AND s.itemoption = p.itemoption "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item ii with (nolock) " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and i.itemgubun = '10' " + vbcrlf
		sqlStr = sqlStr + " 		and i.shopitemid = ii.itemid " + vbcrlf
        sqlstr = sqlstr & " left join [db_user].[dbo].tbl_user_c c with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on i.makerid=c.userid"	& vbcrlf
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on i.shopitemid = a.itemid"	& vbcrlf
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] os with (nolock) "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and os.itemgubun = i.itemgubun "
        sqlStr = sqlStr & " 		and os.itemid = i.shopitemid "
        sqlStr = sqlStr & " 		and os.itemoption = i.itemoption "
        sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_current_agvstock_summary] agvs with (nolock) "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and agvs.itemgubun = i.itemgubun "
        sqlStr = sqlStr & " 		and agvs.itemid = i.shopitemid "
        sqlStr = sqlStr & " 		and agvs.itemoption = i.itemoption "
		sqlstr = sqlstr + " where 1=1 " & sqlsearch
		sqlstr = sqlstr + " order by s.itemgubun, s.itemid desc, s.itemoption "

		''response.write sqlstr & "<br>"
		'response.End
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

					FItemList(i).Fmakerid	    = rsget("makerid")
        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
					FItemList(i).FMwDiv         = rsget("mwdiv")
					FItemList(i).FCenterMwdiv   = FItemList(i).FMwDiv
					FItemList(i).FSellcash      = rsget("sellcash")
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
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Foffsellno		= rsget("offsellno")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")
                    FItemList(i).Fimgsmall      = rsget("offimgsmall")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

					'FItemList(i).FOldSystemCurrno = rsget("currno")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).Fitemrackcode = rsget("itemrackcode")
					FItemList(i).fprtidx = rsget("prtidx")
					FItemList(i).fsubitemrackcode = rsget("subitemrackcode")
					FItemList(i).Forgprice = rsget("orgprice")		'// 소비자가
					FItemList(i).FlastIpgoDate = rsget("lastIpgoDate")

                    FItemList(i).Fagvstock = rsget("agvstock")
                    FItemList(i).Fbulkstock = rsget("bulkstock")
                    FItemList(i).Flastbulkstockdate = rsget("lastbulkstockdate")
                    if IsNull(FItemList(i).Flastbulkstockdate) then
                        FItemList(i).Fbulkstock = NULL
                    elseif (DateDiff("d", FItemList(i).Flastbulkstockdate, Now()) > 5) then
                        FItemList(i).Fbulkstock = NULL
                    end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'//admin/stock/brandcurrentstock.asp
	' 이 펑션을 수정할경우 GetCurrentStockByOfflineBrand 도 반드시 같이 수정해 주세요.
	public sub GetCurrentStockByOfflineBrand_notpaping
		dim sqlstr, i, sqlsearch, stockFieldName

		if (FRectItemGubun<>"") then
		    sqlsearch = sqlsearch + " and s.itemgubun='" + CStr(FRectItemGubun) + "'"
		end if

		if FRectMakerid<>"" then
		    sqlsearch = sqlsearch + " and i.makerid='" + CStr(FRectMakerid) + "'"
		end if
		if FRectOnlyIsUsing<>"" then
			sqlsearch = sqlsearch + " and i.isusing='" + FRectOnlyIsUsing + "'"
		end if

		if (FRectItemIdArr<>"") then
            sqlsearch = sqlsearch + " and i.shopitemid in (" & FRectItemIdArr & ")"
        end if

        if (FRectItemName<>"") then
            sqlsearch = sqlsearch + " and i.shopitemname like '%" & FRectItemName & "%'"
        end if

        if FRectCenterMWDiv="N" then
            sqlsearch = sqlsearch + " and i.centermwdiv is NULL"
        elseif FRectCenterMWDiv<>"" then
            sqlsearch = sqlsearch + " and i.centermwdiv='" + FRectCenterMWDiv + "'"
        end if

        if (FRectCD1<>"") then
            sqlsearch = sqlsearch + " and i.catecdl='" & FRectCD1 & "'"
        end if

        if (FRectCD2<>"") then
            sqlsearch = sqlsearch + " and i.catecdm='" & FRectCD2 & "'"
        end if

        if (FRectCD3<>"") then
            sqlsearch = sqlsearch + " and i.catecdn='" & FRectCD3 & "'"
        end if

'		if FRectlimitrealstock="1UP" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1"
'		elseif FRectlimitrealstock="0DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 0"
'		elseif FRectlimitrealstock="20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 20"
'		elseif FRectlimitrealstock="1UP20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1 and s.realstock <= 20"
'		end if

		''XXXX 시스템재고로 변경 2014/08/01
		''두가지 선택 2014-11-04
		stockFieldName = "totsysstock"
		if (FRectStockType = "real") then
			stockFieldName = "realstock"
		end if
		if FRectlimitrealstock="1UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1"
		elseif FRectlimitrealstock="0DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 0"
		elseif FRectlimitrealstock="20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock="1UP20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock = "20UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 20"
		end If

		If (FRectStartDate <> "") And IsNumeric(FRectStartDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) >= " & FRectStartDate & " "
		End If

		If (FRectEndDate <> "") And IsNumeric(FRectEndDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) <= " & FRectEndDate & " "
		End If

        if FRectRackCode <> "" then
            sqlsearch = sqlsearch & " and Left(os.rackcodeByOption, " & Len(FRectRackCode) & ") = '" & FRectRackCode & "' "
        end if

        if FRectBulkStockGubun <> "" then
            if (FRectBulkStockGubun = "nul") then
                '// 입력이전
                sqlsearch = sqlsearch & " and DateDiff(day, IsNull(agvs.lastbulkstockdate, DateAdd(day, -30, getdate())), getdate()) > 5 "
            elseif (FRectBulkStockGubun = "err") then
                '// 벌크오차 있음
                sqlsearch = sqlsearch & " and DateDiff(day, IsNull(agvs.lastbulkstockdate, DateAdd(day, -30, getdate())), getdate()) <= 5 "
                sqlsearch = sqlsearch & " and (s.realstock + s.ipkumdiv5 + s.offconfirmno) <> (IsNull(agvs.bulkstock, 0) + IsNull(agvs.agvstock, 0)) "
            end if
        end if

        sqlstr = "select count(s.itemid) as CNT"
        sqlstr = sqlstr + " from [db_shop].[dbo].tbl_shop_item i with (nolock)"
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
        sqlstr = sqlstr + " 	on s.itemid = i.shopitemid and s.itemgubun = i.itemgubun "
        sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_current_agvstock_summary] agvs with (nolock) "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and agvs.itemgubun = i.itemgubun "
        sqlStr = sqlStr & " 		and agvs.itemid = i.shopitemid "
        sqlStr = sqlStr & " 		and agvs.itemoption = i.itemoption "

		If (FRectUseOffInfo = "") Then
			sqlstr = sqlstr + " and s.itemgubun <> '10' "
		Else
			sqlstr = sqlstr + " and s.itemoption = i.itemoption "
		End If

		If FRectStartDate <> "" Or FRectEndDate <> "" then
			sqlstr = sqlstr + " 	LEFT JOIN [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p with (nolock) ON 1 = 1 "
			sqlstr = sqlstr + " 		AND p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
			sqlstr = sqlstr + " 		AND s.itemgubun = p.itemgubun "
			sqlstr = sqlstr + " 		AND s.itemid = p.itemid "
			sqlstr = sqlstr + " 		AND s.itemoption = p.itemoption "
		End if

        if FRectRackCode <> "" then
            sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] os with (nolock) "
            sqlStr = sqlStr & " 		on "
            sqlStr = sqlStr & " 			1 = 1 "
            sqlStr = sqlStr & " 			and os.itemgubun = i.itemoption "
            sqlStr = sqlStr & " 			and os.itemid = i.shopitemid "
            sqlStr = sqlStr & " 			and os.itemoption = i.itemoption "
        end if

		sqlstr = sqlstr + " where 1=1 " & sqlsearch

		''response.write sqlstr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlstr = "select top " + CStr(FPageSize*FCurrPage)
		sqlstr = sqlstr & " s.itemgubun,s.itemid,s.itemoption,s.ipgono,s.reipgono,s.totipgono,s.offchulgono,s.offrechulgono,s.etcchulgono"
		sqlstr = sqlstr & " ,s.etcrechulgono,s.totchulgono,s.sellno,s.resellno,s.totsellno,s.errcsno,s.errbaditemno,s.errrealcheckno,s.erretcno"
		sqlstr = sqlstr & " ,s.toterrno,s.totsysstock,s.availsysstock,s.realstock,s.sell7days,s.offchulgo7days,s.ipkumdiv5,s.ipkumdiv4"
		sqlstr = sqlstr & " ,s.ipkumdiv2,s.offconfirmno,s.offjupno,s.requireno,s.shortageno,s.preorderno,s.preordernofix,s.offsellno"
		sqlstr = sqlstr & " ,s.maxsellday,s.imgsmall,s.itemoptionname,s.regdate,s.lastupdate,s.requireMaxno,s.itemgrade"
		sqlstr = sqlstr + " , i.makerid"
		sqlstr = sqlstr + " , i.shopitemname as itemname"
		sqlstr = sqlstr & " , i.shopitemprice as sellcash"
		sqlstr = sqlstr & " , i.shopsuplycash as buycash"
		sqlstr = sqlstr & " , i.isusing"
		sqlstr = sqlstr & " , '' as deliverytype"
		sqlstr = sqlstr & " ,'' as sellyn"
		sqlstr = sqlstr & " ,'' as limityn"
		sqlstr = sqlstr & " ,'' as limitno"
		sqlstr = sqlstr & " ,'' as limitsold"
		sqlstr = sqlstr & " , i.centermwdiv as mwdiv"
		sqlstr = sqlstr & " ,'' as danjongyn"
		sqlstr = sqlstr & " , IsNull(os.rackcodeByOption, (case when i.itemgubun='10' then ii.itemrackcode else i.offitemrackcode end)) as itemrackcode"
		sqlstr = sqlstr & " , IsNull(os.subRackcodeByOption, ISNULL(a.subitemrackcode,ii.itemrackcode)) AS subitemrackcode"
		sqlstr = sqlstr & " , i.shopitemoptionname as codeview"
		sqlstr = sqlstr & " ,'' AS optionusing"
		sqlstr = sqlstr & " ,'' as optlimityn"
		sqlstr = sqlstr & " ,'' as optlimitno"
		sqlstr = sqlstr & " ,'' as optlimitsold"
		sqlstr = sqlstr & " ,i.orgsellprice as orgprice"
		sqlstr = sqlstr & " , i.centermwdiv as centermwdiv"
		sqlstr = sqlstr & " , p.lastIpgoDate"
		sqlstr = sqlstr & " ,'' as publicBarcode"
		sqlstr = sqlstr & " ,0 as prevMonthSellCnt"
		sqlstr = sqlstr & " ,'' as OffMwMargin"
		sqlstr = sqlstr & " , agvs.agvstock"
		sqlstr = sqlstr & " , agvs.bulkstock"
		sqlstr = sqlstr & " , agvs.lastbulkstockdate"
		sqlstr = sqlstr & " ,'' as warehouseCd"
		sqlstr = sqlstr & " , 0 IDX"
		sqlstr = sqlstr & " , c.prtidx"
		sqlstr = sqlstr + " , i.offimgsmall"
		sqlstr = sqlstr + " , i.shopsuplycash as suplycash, d.defaultmargin, d.defaultsuplymargin "
        sqlstr = sqlstr + " from [db_shop].[dbo].tbl_shop_item i with (nolock)"
		sqlstr = sqlstr + " 	left join [db_shop].[dbo].tbl_shop_designer d with (nolock) "
		sqlstr = sqlstr + " 	on "
		sqlstr = sqlstr + " 		1 = 1 "
		sqlstr = sqlstr + " 		and i.makerid=d.makerid "
		sqlstr = sqlstr + " 		and d.shopid='streetshop000' "
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
        sqlstr = sqlstr + " 	on s.itemid = i.shopitemid and s.itemgubun = i.itemgubun "

		If (FRectUseOffInfo = "") Then
			sqlstr = sqlstr + " and s.itemgubun <> '10' "
		Else
			sqlstr = sqlstr + " and s.itemoption = i.itemoption "
		End If

		sqlstr = sqlstr + " 	LEFT JOIN [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p with (nolock) ON 1 = 1 "
		sqlstr = sqlstr + " 		AND p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
		sqlstr = sqlstr + " 		AND s.itemgubun = p.itemgubun "
		sqlstr = sqlstr + " 		AND s.itemid = p.itemid "
		sqlstr = sqlstr + " 		AND s.itemoption = p.itemoption "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item ii with (nolock) " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and i.itemgubun = '10' " + vbcrlf
		sqlStr = sqlStr + " 		and i.shopitemid = ii.itemid " + vbcrlf
        sqlstr = sqlstr & " left join [db_user].[dbo].tbl_user_c c with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on i.makerid=c.userid"	& vbcrlf
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on i.shopitemid = a.itemid"	& vbcrlf
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] os with (nolock) "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and os.itemgubun = i.itemgubun "
        sqlStr = sqlStr & " 		and os.itemid = i.shopitemid "
        sqlStr = sqlStr & " 		and os.itemoption = i.itemoption "
        sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_current_agvstock_summary] agvs with (nolock) "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and agvs.itemgubun = i.itemgubun "
        sqlStr = sqlStr & " 		and agvs.itemid = i.shopitemid "
        sqlStr = sqlStr & " 		and agvs.itemoption = i.itemoption "
		sqlstr = sqlstr + " where 1=1 " & sqlsearch
		sqlstr = sqlstr + " order by s.itemgubun, s.itemid desc, s.itemoption "

		'response.write sqlstr & "<br>"
		'response.End
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        FResultCount  = rsget.RecordCount
		FtotalCount =  rsget.RecordCount
		if (FResultCount<1) then FResultCount=0

		i=0
		if not rsget.EOF then
			farrlist = rsget.getRows()
		end if
		rsget.Close
	end Sub

	public sub GetRealStockByOnlineBrand
		dim sqlstr, i, sqlsearch, sqlQuery, stockFieldName

		if (FRectMakerid<>"") then
		    sqlsearch = sqlsearch + " and IsNull(i.makerid, si.makerid)='" + CStr(FRectMakerid) + "'"
        end If

		if (FRectOnlyIsUsing <>"") then
            sqlsearch = sqlsearch + " and i.isusing = '" + FRectOnlyIsUsing + "' "
        end If

		if FRectMwDiv="MW" then
            sqlsearch = sqlsearch + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

        if FRectCenterMWDiv="N" then
            sqlsearch = sqlsearch + " and si.centermwdiv is NULL"
        elseif FRectCenterMWDiv<>"" then
            sqlsearch = sqlsearch + " and si.centermwdiv='" + FRectCenterMWDiv + "'"
        end If

		stockFieldName = "totsysstock"
		if (FRectStockType = "real") then
			stockFieldName = "realstock"
		end if
		if FRectlimitrealstock="1UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1"
		elseif FRectlimitrealstock="0DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 0"
		elseif FRectlimitrealstock="20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock="1UP20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock = "20UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 20"
		end If

		If (FRectStartDate <> "") And IsNumeric(FRectStartDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) >= " & FRectStartDate & " "
		End If

		If (FRectEndDate <> "") And IsNumeric(FRectEndDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) <= " & FRectEndDate & " "
		End If

        if (FRectPurchaseType <> "") then
        	sqlsearch = sqlsearch + " 	and IsNull(pt.purchasetype, 1) = " & FRectPurchaseType & " "
        end if

		sqlstr = " SELECT "
		sqlstr = sqlstr + " 	IsNull(i.makerid, si.makerid) as makerid "
		sqlstr = sqlstr + " 	,count(distinct i.itemid) as itemCnt "
		sqlstr = sqlstr + " 	,sum(case when s." + stockFieldName+ " > 0 then 1 else 0 end) as itemPlusCnt "
		sqlstr = sqlstr + " 	,s.itemgubun "
		sqlstr = sqlstr + " 	,IsNull(i.mwdiv,'Z') as mwdiv "
		sqlstr = sqlstr + " 	,IsNull(si.centermwdiv, 'N') as centermwdiv "
		sqlstr = sqlstr + " 	,sum(s.totsysstock) as totsysstock "
		sqlstr = sqlstr + " 	,sum(s.realstock) as realstock "
		sqlstr = sqlstr + " 	,sum(s.totsellno) as totsellno "
		sqlstr = sqlstr + " 	,sum(s.errrealcheckno) as errrealcheckno "
		sqlstr = sqlstr + " 	,sum(s.errbaditemno) as errbaditemno "
		sqlstr = sqlstr + " FROM "
		sqlstr = sqlstr + " 	[db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " 	LEFT JOIN [db_item].[dbo].tbl_item i "
		sqlstr = sqlstr + " 	ON "
		sqlstr = sqlstr + " 		1 = 1 "
		sqlstr = sqlstr + " 		and s.itemid = i.itemid "
		sqlstr = sqlstr + " 		AND s.itemgubun = '10' "
		sqlstr = sqlstr + " 	LEFT JOIN [db_shop].[dbo].tbl_shop_item si "
		sqlstr = sqlstr + " 	ON "
		sqlstr = sqlstr + " 		1 = 1 "
		''sqlstr = sqlstr + " 		AND s.itemgubun <> '10' "
		sqlstr = sqlstr + " 		and s.itemgubun = si.itemgubun "
		sqlstr = sqlstr + " 		AND s.itemid = si.shopitemid "
		sqlstr = sqlstr + " 		AND s.itemoption = si.itemoption "
		sqlstr = sqlstr + " 	LEFT JOIN [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p "
		sqlstr = sqlstr + " 	ON 1 = 1 "
		sqlstr = sqlstr + " 		AND p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
		sqlstr = sqlstr + " 		AND s.itemgubun = p.itemgubun "
		sqlstr = sqlstr + " 		AND s.itemid = p.itemid "
		sqlstr = sqlstr + " 		AND s.itemoption = p.itemoption "
		if (FRectPurchaseType <> "") then
			sqlstr = sqlstr + " 	left join db_partner.dbo.tbl_partner pt "
			sqlstr = sqlstr + " 	on "
			sqlstr = sqlstr + " 		i.makerid = pt.id "
		End if
		sqlstr = sqlstr + " WHERE 1=1 " & sqlsearch
		sqlstr = sqlstr + " GROUP BY "
		sqlstr = sqlstr + " 	IsNull(i.makerid, si.makerid) "
		sqlstr = sqlstr + " 	,s.itemgubun "
		sqlstr = sqlstr + " 	,i.mwdiv "
		sqlstr = sqlstr + " 	,si.centermwdiv "
		''response.write sqlstr & "<Br>"
		''response.end

		sqlQuery = " select count(T.makerid) as cnt "
		sqlQuery = sqlQuery + " from ("
		sqlQuery = sqlQuery + sqlstr
		sqlQuery = sqlQuery + " ) T "

		'response.write sqlQuery & "<Br>"
        rsget.Open sqlQuery,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlQuery = " select top " + CStr(FPageSize*FCurrPage) + " T.* "
		sqlQuery = sqlQuery + " from ("
		sqlQuery = sqlQuery + sqlstr
		sqlQuery = sqlQuery + " ) T "
		sqlQuery = sqlQuery + " ORDER BY "
		sqlQuery = sqlQuery + "  	T.makerid "
		sqlQuery = sqlQuery + "  	,T.itemgubun "
		sqlQuery = sqlQuery + "  	,T.mwdiv "
		sqlQuery = sqlQuery + "  	,T.centermwdiv "

		''response.write sqlQuery & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlQuery,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

				FItemList(i).Fmakerid	    = rsget("makerid")
				FItemList(i).Fitemgubun     = rsget("itemgubun")
				FItemList(i).Fmwdiv			= rsget("mwdiv")
				FItemList(i).FCentermwdiv 	= rsget("centermwdiv")
				FItemList(i).FitemCnt 		= rsget("itemCnt")
				FItemList(i).FitemPlusCnt 	= rsget("itemPlusCnt")
				FItemList(i).Ftotsellno     = rsget("totsellno")

				FItemList(i).Ftotsysstock   = rsget("totsysstock")
				FItemList(i).Frealstock     = rsget("realstock")
				FItemList(i).Ferrrealcheckno	= rsget("errrealcheckno")
				FItemList(i).Ferrbaditemno		= rsget("errbaditemno")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public sub GetRealStockByOfflineBrand
		dim sqlstr, i, sqlsearch, sqlQuery, stockFieldName

		if (FRectMakerid<>"") then
		    sqlsearch = sqlsearch + " and si.makerid='" + CStr(FRectMakerid) + "'"
        end If

		if (FRectOnlyIsUsing <>"") then
            sqlsearch = sqlsearch + " and si.isusing = '" + FRectOnlyIsUsing + "' "
        end If

		' if FRectMwDiv="MW" then
        '     sqlsearch = sqlsearch + " and i.mwdiv<>'U'"
        ' elseif FRectMwDiv<>"" then
        '     sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        ' end If

		If FRectItemGubun <> "" And FRectItemGubun <> "exc10" Then
			sqlsearch = sqlsearch + " and si.itemgubun = '" + FRectItemGubun + "' "
		End If

        if FRectCenterMWDiv="N" then
            sqlsearch = sqlsearch + " and si.centermwdiv is NULL"
        elseif FRectCenterMWDiv<>"" then
            sqlsearch = sqlsearch + " and si.centermwdiv='" + FRectCenterMWDiv + "'"
        end If

		stockFieldName = "totsysstock"
		if (FRectStockType = "real") then
			stockFieldName = "realstock"
		end if
		if FRectlimitrealstock="1UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1"
		elseif FRectlimitrealstock="0DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 0"
		elseif FRectlimitrealstock="20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock="1UP20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock = "20UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 20"
		end If

		If (FRectStartDate <> "") And IsNumeric(FRectStartDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) >= " & FRectStartDate & " "
		End If

		If (FRectEndDate <> "") And IsNumeric(FRectEndDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) <= " & FRectEndDate & " "
		End If

		sqlstr = " SELECT "
		sqlstr = sqlstr + " 	si.makerid "
		sqlstr = sqlstr + " 	,count(distinct si.shopitemid) as itemCnt "
		sqlstr = sqlstr + " 	,sum(case when s." + stockFieldName+ " > 0 then 1 else 0 end) as itemPlusCnt "
		sqlstr = sqlstr + " 	,s.itemgubun "
		sqlstr = sqlstr + " 	,IsNull(si.centermwdiv, 'N') as mwdiv "
		sqlstr = sqlstr + " 	,IsNull(si.centermwdiv, 'N') as centermwdiv "
		sqlstr = sqlstr + " 	,sum(s.totsysstock) as totsysstock "
		sqlstr = sqlstr + " 	,sum(s.realstock) as realstock "
		sqlstr = sqlstr + " 	,sum(s.totsellno) as totsellno "

		sqlstr = sqlstr + " 	,sum(s.errrealcheckno) as errrealcheckno "
		sqlstr = sqlstr + " 	,sum(s.errbaditemno) as errbaditemno "

		sqlstr = sqlstr + " FROM "
		sqlstr = sqlstr + " 	[db_shop].[dbo].tbl_shop_item si "
		sqlstr = sqlstr + " 	JOIN [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " 	ON "
		sqlstr = sqlstr + " 		1 = 1 "
		sqlstr = sqlstr + " 		AND s.itemgubun = si.itemgubun "
		sqlstr = sqlstr + " 		AND s.itemid = si.shopitemid "
		sqlstr = sqlstr + " 		AND s.itemoption = si.itemoption "
		sqlstr = sqlstr + " 		AND si.itemgubun <> '10' "
		sqlstr = sqlstr + " 	LEFT JOIN [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p ON 1 = 1 "
		sqlstr = sqlstr + " 		AND p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
		sqlstr = sqlstr + " 		AND s.itemgubun = p.itemgubun "
		sqlstr = sqlstr + " 		AND s.itemid = p.itemid "
		sqlstr = sqlstr + " 		AND s.itemoption = p.itemoption "
		sqlstr = sqlstr + " WHERE 1=1 " & sqlsearch
		sqlstr = sqlstr + " GROUP BY "
		sqlstr = sqlstr + " 	si.makerid "
		sqlstr = sqlstr + " 	,s.itemgubun "
		sqlstr = sqlstr + " 	,si.centermwdiv "
		''response.write sqlstr & "<Br>"

		sqlQuery = " select count(T.makerid) as cnt "
		sqlQuery = sqlQuery + " from ("
		sqlQuery = sqlQuery + sqlstr
		sqlQuery = sqlQuery + " ) T "

		'response.write sqlQuery & "<Br>"
        rsget.Open sqlQuery,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlQuery = " select top " + CStr(FPageSize*FCurrPage) + " T.* "
		sqlQuery = sqlQuery + " from ("
		sqlQuery = sqlQuery + sqlstr
		sqlQuery = sqlQuery + " ) T "
		sqlQuery = sqlQuery + " ORDER BY "
		sqlQuery = sqlQuery + "  	T.makerid "
		sqlQuery = sqlQuery + "  	,T.itemgubun "
		sqlQuery = sqlQuery + "  	,T.mwdiv "
		sqlQuery = sqlQuery + "  	,T.centermwdiv "

		''response.write sqlQuery & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlQuery,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

				FItemList(i).Fmakerid	    = rsget("makerid")
				FItemList(i).Fitemgubun     = rsget("itemgubun")
				FItemList(i).Fmwdiv			= rsget("mwdiv")
				FItemList(i).FCentermwdiv 	= rsget("centermwdiv")
				FItemList(i).FitemCnt 		= rsget("itemCnt")
				FItemList(i).FitemPlusCnt 	= rsget("itemPlusCnt")
				FItemList(i).Ftotsellno     = rsget("totsellno")

				FItemList(i).Ftotsysstock   = rsget("totsysstock")
				FItemList(i).Frealstock     = rsget("realstock")
				FItemList(i).Ferrrealcheckno	= rsget("errrealcheckno")
				FItemList(i).Ferrbaditemno		= rsget("errbaditemno")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

    ''//designer/storage/dailystorage.asp 2016/01/07 분리
    public sub GetCurrentStockByOnlineBrandByDesigner
		dim sqlstr, i, sqlsearch, stockFieldName

		sqlsearch = sqlsearch + " and i.makerid='" + CStr(FRectMakerid) + "'"

        if (FRectItemIdArr<>"") then
            sqlsearch = sqlsearch + " and i.itemid in (" & FRectItemIdArr & ")"
        end if

        if (FRectCD1<>"") then
            sqlsearch = sqlsearch + " and i.cate_large='" & FRectCD1 & "'"
        end if

        if (FRectCD2<>"") then
            sqlsearch = sqlsearch + " and i.cate_mid='" & FRectCD2 & "'"
        end if

        if (FRectCD3<>"") then
            sqlsearch = sqlsearch + " and i.cate_small='" & FRectCD3 & "'"
        end if

        ''반품 대상 상품(실사재고<>0)
        if (FRectReturnItemGubun="reton") then
            FRectOnlySellyn  = "N"
            ''FRectOnlyIsUsing = ""
            FRectDanjongyn   = "YM"
            FRectLimityn     = ""
            FRectMwDiv       = "MW"
            sqlsearch = sqlsearch + " and s.realstock <> 0"
        ''반품 완료 상품(실사재고=0)
        elseif (FRectReturnItemGubun="retfin") then
            FRectOnlySellyn  = "N"
            ''FRectOnlyIsUsing = ""
            FRectDanjongyn   = "YM"
            FRectLimityn     = ""
            FRectMwDiv       = "MW"
            sqlsearch = sqlsearch + " and s.realstock = 0"
        ''랙진열상품
        elseif (FRectReturnItemGubun="rackdisp") then
            sqlsearch = sqlsearch + " and ((i.sellyn <>'N') or (i.danjongyn in ('N','S')))"
            FRectOnlySellyn  = ""
            FRectOnlyIsUsing = ""
            FRectDanjongyn   = ""
            FRectLimityn     = ""
            'FRectMwDiv       = ""
        end if

        if FRectOnlySellyn="YS" then
            sqlsearch = sqlsearch + " and i.sellyn <>'N'"
        elseif FRectOnlySellyn<>"" then
            sqlsearch = sqlsearch + " and i.sellyn = '" + FRectOnlySellyn + "' "
        end if

        if (FRectOnlyIsUsing <>"") then
            sqlsearch = sqlsearch + " and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if

        if FRectDanjongyn="SN" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'Y'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'N'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlsearch = sqlsearch + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectLimityn="Y0" then
            sqlsearch = sqlsearch + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlsearch = sqlsearch + " and i.limityn='" + FRectLimityn + "'"
        end if

		if FRectMwDiv="MW" then
            sqlsearch = sqlsearch + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

        if FRectCenterMWDiv="N" then
            sqlsearch = sqlsearch + " and si.centermwdiv is NULL"
        elseif FRectCenterMWDiv<>"" then
            sqlsearch = sqlsearch + " and si.centermwdiv='" + FRectCenterMWDiv + "'"
        end if

        if frectsoldout_gubun = "Y" then
    		sqlsearch = sqlsearch + " and ((i.sellyn<>'Y') or ((i.limityn<>'N') and (i.limitno-i.limitsold<1)))"
    	elseif frectsoldout_gubun = "N" then
    		sqlsearch = sqlsearch + " and s.availsysstock > 0"
    	else
        end if

		''두가지 선택 2014-11-04
		stockFieldName = "totsysstock"
		if (FRectStockType = "real") then
			stockFieldName = "realstock"
		end if
		if FRectlimitrealstock="1UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1"
		elseif FRectlimitrealstock="0DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 0"
		elseif FRectlimitrealstock="20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock="1UP20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock = "20UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 20"
		end if

        sqlstr = " select count(i.itemid) as cnt"
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s"
        sqlstr = sqlstr + " 	on s.itemid = i.itemid"
        sqlstr = sqlstr + "		and s.itemgubun ='10'"

        if (FRectCenterMWDiv<>"") then
            sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item si "
    		sqlStr = sqlstr + "     on si.itemgubun='10' and s.itemgubun=si.itemgubun"
    		sqlStr = sqlstr + "     and s.itemid=si.shopitemid"
    		sqlStr = sqlstr + "     and s.itemoption=si.itemoption"
	    end if

        sqlstr = sqlstr + " where 1=1 " & sqlsearch

		'response.write sqlstr & "<Br>"
        rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlstr = " select top " + CStr(FPageSize*FCurrPage)
		sqlstr = sqlstr + " s.*, i.makerid, i.itemname, (i.sellcash + IsNULL(v.optaddprice,0)) as sellcash, (i.buycash + IsNULL(v.optaddbuyprice,0)) as buycash, i.isusing, i.deliverytype, i.sellyn"
		sqlstr = sqlstr + " , i.limityn, i.limitno, i.limitsold, i.mwdiv, i.danjongyn,i.itemrackcode, IsNULL(v.optionname,'') as codeview"
		sqlstr = sqlstr + " , IsNULL(v.isusing,'Y') as optionusing , v.optlimityn, v.optlimitno, v.optlimitsold, i.orgprice "
		sqlstr = sqlstr + " , si.centermwdiv"
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s"
        sqlstr = sqlstr + " 	on s.itemid = i.itemid"
        sqlstr = sqlstr + "		and s.itemgubun ='10'"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v "
		sqlStr = sqlstr + "     on s.itemgubun='10'"
		sqlStr = sqlstr + "     and s.itemid=v.itemid"
		sqlStr = sqlstr + "     and s.itemoption=v.itemoption"
		sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item si "
		sqlStr = sqlstr + "     on s.itemgubun=si.itemgubun"
		sqlStr = sqlstr + "     and s.itemid=si.shopitemid"
		sqlStr = sqlstr + "     and s.itemoption=si.itemoption"
        sqlstr = sqlstr + " where 1=1 " & sqlsearch
		sqlstr = sqlstr + " order by s.itemgubun asc, s.itemid desc, s.itemoption asc"

		''response.write sqlstr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

					FItemList(i).Fmakerid	    = rsget("makerid")
        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = db2html(rsget("itemoption"))
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
        			FItemList(i).Fdeliverytype  = rsget("deliverytype")
        			FItemList(i).Fmwdiv			= rsget("mwdiv")
					FItemList(i).Fsellyn		= rsget("sellyn")
					FItemList(i).Flimityn		= rsget("limityn")
					FItemList(i).Flimitno		= rsget("limitno")
					FItemList(i).Flimitsold		= rsget("limitsold")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

					''FItemList(i).FOldSystemCurrno = rsget("currno")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).Fdanjongyn  = rsget("danjongyn")
					FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")
					FItemList(i).FItemrackcode = rsget("itemrackcode")
					FItemList(i).FOnlineCurrentSellcash = rsget("sellcash")
                    FItemList(i).FOnlineCurrentBuycash = rsget("buycash")
					FItemList(i).Fsellcash = rsget("sellcash")
                    FItemList(i).Fbuycash = rsget("buycash")
					FItemList(i).Forgprice = rsget("orgprice")		'// 소비자가
					FItemList(i).FCentermwdiv = rsget("centermwdiv")
					''FItemList(i).FlastIpgoDate = rsget("lastIpgoDate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	'//admin/stock/brandcurrentstock.asp
	public sub GetCurrentStockByOnlineBrand
		dim sqlstr, i, sqlsearch, stockFieldName

		if (FRectMakerid<>"") then
		    sqlsearch = sqlsearch + " and i.makerid='" + CStr(FRectMakerid) + "'"
        end if

        if (FRectItemIdArr<>"") then
            sqlsearch = sqlsearch + " and i.itemid in (" & FRectItemIdArr & ")"
        end if

        'if (FRectItemName<>"") then
        '    sqlsearch = sqlsearch + " and i.itemname like '%" & FRectItemName & "%'"
        'end if

        if (FRectCD1<>"") then
            sqlsearch = sqlsearch + " and i.cate_large='" & FRectCD1 & "'"
        end if

        if (FRectCD2<>"") then
            sqlsearch = sqlsearch + " and i.cate_mid='" & FRectCD2 & "'"
        end if

        if (FRectCD3<>"") then
            sqlsearch = sqlsearch + " and i.cate_small='" & FRectCD3 & "'"
        end if

        ''반품 대상 상품(실사재고<>0)
        if (FRectReturnItemGubun="reton") then
            FRectOnlySellyn  = "N"
            ''FRectOnlyIsUsing = ""
            FRectDanjongyn   = "YM"
            FRectLimityn     = ""
            FRectMwDiv       = "MW"
            sqlsearch = sqlsearch + " and s.realstock <> 0"
        ''반품 완료 상품(실사재고=0)
        elseif (FRectReturnItemGubun="retfin") then
            FRectOnlySellyn  = "N"
            ''FRectOnlyIsUsing = ""
            FRectDanjongyn   = "YM"
            FRectLimityn     = ""
            FRectMwDiv       = "MW"
            sqlsearch = sqlsearch + " and s.realstock = 0"
        ''랙진열상품
        elseif (FRectReturnItemGubun="rackdisp") then
            sqlsearch = sqlsearch + " and ((i.sellyn <>'N') or (i.danjongyn in ('N','S')))"
            FRectOnlySellyn  = ""
            FRectOnlyIsUsing = ""
            FRectDanjongyn   = ""
            FRectLimityn     = ""
            'FRectMwDiv       = ""
        end if

        if FRectOnlySellyn="YS" then
            sqlsearch = sqlsearch + " and i.sellyn <>'N'"
        elseif FRectOnlySellyn<>"" then
            sqlsearch = sqlsearch + " and i.sellyn = '" + FRectOnlySellyn + "' "
        end if

        if (FRectOnlyIsUsing <>"") then
            sqlsearch = sqlsearch + " and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if

        if FRectDanjongyn="SN" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'Y'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'N'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlsearch = sqlsearch + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectLimityn="Y0" then
            sqlsearch = sqlsearch + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlsearch = sqlsearch + " and i.limityn='" + FRectLimityn + "'"
        end if

		if FRectMwDiv="MW" then
            sqlsearch = sqlsearch + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

        if FRectCenterMWDiv="N" then
            sqlsearch = sqlsearch + " and si.centermwdiv is NULL"
        elseif FRectCenterMWDiv<>"" then
            sqlsearch = sqlsearch + " and si.centermwdiv='" + FRectCenterMWDiv + "'"
        end if

        if frectsoldout_gubun = "Y" then
    		sqlsearch = sqlsearch + " and ((i.sellyn<>'Y') or ((i.limityn<>'N') and (i.limitno-i.limitsold<1)))"
    	elseif frectsoldout_gubun = "N" then
    		sqlsearch = sqlsearch + " and s.availsysstock > 0"
    	else
        end if

'		if FRectlimitrealstock="1UP" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1"
'		elseif FRectlimitrealstock="0DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 0"
'		elseif FRectlimitrealstock="20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 20"
'		elseif FRectlimitrealstock="1UP20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1 and s.realstock <= 20"
'		end if

		''두가지 선택 2014-11-04
		stockFieldName = "totsysstock"
		if (FRectStockType = "real") then
			stockFieldName = "realstock"
		end if
		if FRectlimitrealstock="1UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1"
		elseif FRectlimitrealstock="0DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 0"
		elseif FRectlimitrealstock="20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock="1UP20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock = "20UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 20"
		elseif FRectlimitrealstock="1UP3DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 3"
		elseif FRectlimitrealstock="MINUS" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " < 0"
		end If

		If (FRectStartDate <> "") And IsNumeric(FRectStartDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) >= " & FRectStartDate & " "
		End If

		If (FRectEndDate <> "") And IsNumeric(FRectEndDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) <= " & FRectEndDate & " "
		End If

		If FRectExcIts = "Y" Then
			sqlsearch = sqlsearch + " and i.makerid <> 'ithinkso' "
		End If

		if FRectDispCate<>"" Then
			sqlsearch = sqlsearch + " and s.itemgubun = '10' "
		    if LEN(FRectDispCate)>3 then
		         sqlsearch = sqlsearch + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
		    end if
			sqlsearch = sqlsearch + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        sqlstr = " select count(i.itemid) as cnt"
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s"
        sqlstr = sqlstr + " 	on s.itemid = i.itemid"
        sqlstr = sqlstr + "		and s.itemgubun ='10'"
        sqlstr = sqlstr + "		and s.itemoption >='0000' "

		If FRectStartDate <> "" Or FRectEndDate <> "" then
			sqlstr = sqlstr + " 	LEFT JOIN [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p ON 1 = 1 "
			sqlstr = sqlstr + " 		AND p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
			sqlstr = sqlstr + " 		AND s.itemgubun = p.itemgubun "
			sqlstr = sqlstr + " 		AND s.itemid = p.itemid "
			sqlstr = sqlstr + " 		AND s.itemoption = p.itemoption "
		End if

        if (FRectCenterMWDiv<>"") then
            sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item si "
    		sqlStr = sqlstr + "     on si.itemgubun='10' and s.itemgubun=si.itemgubun"
    		sqlStr = sqlstr + "     and s.itemid=si.shopitemid"
    		sqlStr = sqlstr + "     and s.itemoption=si.itemoption"
	    end if

        sqlstr = sqlstr + " where 1=1 " & sqlsearch

		''response.write sqlstr & "<Br>"
        rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlstr = " select top " + CStr(FPageSize*FCurrPage)
		sqlstr = sqlstr + " s.*, i.makerid, i.itemname, (i.sellcash + IsNULL(v.optaddprice,0)) as sellcash, (i.buycash + IsNULL(v.optaddbuyprice,0)) as buycash, i.isusing, i.deliverytype, i.sellyn"
		sqlstr = sqlstr + " , i.limityn, i.limitno, i.limitsold, i.mwdiv, i.danjongyn, IsNULL(v.optionname,'') as codeview"
		sqlstr = sqlstr + " , IsNULL(v.isusing,'Y') as optionusing , v.optlimityn, v.optlimitno, v.optlimitsold, i.orgprice "
		sqlstr = sqlstr + " , si.centermwdiv, p.lastIpgoDate, ot.barcode as publicBarcode"
		sqlstr = sqlstr + " , v.optrackcode"
        sqlstr = sqlstr + " , IsNull(ot.rackcodeByOption, i.itemrackcode) as itemrackcode "
        sqlstr = sqlstr + " , IsNull(ot.subRackcodeByOption, a.subitemrackcode) as subitemrackcode "
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s"
        sqlstr = sqlstr + " 	on s.itemid = i.itemid"
        sqlstr = sqlstr + "		and s.itemgubun ='10'"
        sqlstr = sqlstr + "		and s.itemoption >='0000' "
		sqlstr = sqlstr + " left join db_item.dbo.tbl_item_option_stock ot "
		sqlstr = sqlstr + " on "
		sqlstr = sqlstr + " 	1 = 1 "
		sqlstr = sqlstr + " 	and s.itemgubun = ot.itemgubun "
		sqlstr = sqlstr + " 	and s.itemid = ot.itemid "
		sqlstr = sqlstr + " 	and s.itemoption = ot.itemoption "
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v "
		sqlStr = sqlstr + "     on s.itemgubun='10'"
		sqlStr = sqlstr + "     and s.itemid=v.itemid"
		sqlStr = sqlstr + "     and s.itemoption=v.itemoption"
		sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item si "
		sqlStr = sqlstr + "		on s.itemgubun=si.itemgubun"
		sqlStr = sqlstr + "		and s.itemid=si.shopitemid"
		sqlStr = sqlstr + "		and s.itemoption=si.itemoption"
		sqlStr = sqlstr + " left join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p "
		sqlStr = sqlstr + " on "
		sqlStr = sqlstr + "		1 = 1 "
		sqlStr = sqlstr + "		and p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
		sqlStr = sqlstr + "		and s.itemgubun = p.itemgubun "
		sqlStr = sqlstr + "		and s.itemid = p.itemid "
		sqlStr = sqlstr + "		and s.itemoption = p.itemoption "
		sqlstr = sqlstr + " left join [db_item].[dbo].[tbl_item_logics_addinfo] a"
        sqlstr = sqlstr + " 	on i.itemid = a.itemid"
        sqlstr = sqlstr + " where 1=1 " & sqlsearch

		Select Case FRectOrderBy
			Case "itemid"
				sqlstr = sqlstr + " order by s.itemgubun asc, s.itemid desc, s.itemoption asc"
			Case "rackcode"
				sqlstr = sqlstr + " order by s.itemgubun asc, itemrackcode asc, subitemrackcode asc, s.itemid asc"
			Case Else
				sqlstr = sqlstr + " order by s.itemgubun asc, s.itemid desc, s.itemoption asc"
		End Select

		''response.write sqlstr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

					FItemList(i).Fmakerid	    = rsget("makerid")
        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = db2html(rsget("itemoption"))
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
        			FItemList(i).Fdeliverytype  = rsget("deliverytype")
        			FItemList(i).Fmwdiv			= rsget("mwdiv")
					FItemList(i).Fsellyn		= rsget("sellyn")
					FItemList(i).Flimityn		= rsget("limityn")
					FItemList(i).Flimitno		= rsget("limitno")
					FItemList(i).Flimitsold		= rsget("limitsold")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

					''FItemList(i).FOldSystemCurrno = rsget("currno")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).Fdanjongyn  = rsget("danjongyn")
					FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")
					FItemList(i).FOnlineCurrentSellcash = rsget("sellcash")
                    FItemList(i).FOnlineCurrentBuycash = rsget("buycash")
					FItemList(i).Fsellcash = rsget("sellcash")
                    FItemList(i).Fbuycash = rsget("buycash")
					FItemList(i).Forgprice = rsget("orgprice")		'/'// 소비자가
					FItemList(i).FCentermwdiv = rsget("centermwdiv")
					FItemList(i).FlastIpgoDate = rsget("lastIpgoDate")

					FItemList(i).FpublicBarcode = rsget("publicBarcode")
					FItemList(i).foptrackcode = rsget("optrackcode")
					FItemList(i).Fitemrackcode    = rsget("itemrackcode")
					FItemList(i).Fsubitemrackcode    = rsget("subitemrackcode")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

    public sub GetAgvStockDiffList
        dim sqlstr, i
        dim cmd

        set cmd = CreateObject("ADODB.Command")

        cmd.ActiveConnection = dbget
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "[db_summary].[dbo].[usp_AGV_GetDiff_GetList]"
        cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
        cmd.Parameters.Append cmd.CreateParameter("@PageSize", adInteger, adParamInput, , FPageSize)
        cmd.Parameters.Append cmd.CreateParameter("@CurrPage", adInteger, adParamInput, , FCurrPage)
        cmd.Parameters.Append cmd.CreateParameter("@makerid", adVarChar, adParamInput, 32, FRectMakerid)
        cmd.Parameters.Append cmd.CreateParameter("@mwdiv", adVarChar, adParamInput, 2, FRectMWDiv)
        cmd.Parameters.Append cmd.CreateParameter("@excTpl", adVarChar, adParamInput, 1, FRectExcIts)
        cmd.Parameters.Append cmd.CreateParameter("@excNoRack", adVarChar, adParamInput, 1, FRectExcNoRack)
        rsget.CursorLocation = adUseClient
        rsget.open cmd, , adOpenStatic, adLockReadOnly

        FTotalCount = cmd.Parameters("returnValue")

        FTotalPage =  CInt(FTotalCount\FPageSize)
        If (FTotalCount\FPageSize) <> (FTotalCount/FPageSize) Then
            FTotalPage = FTotalPage + 1
        End If

        FResultCount = rsget.RecordCount
        redim FItemList(FResultCount)

        if not rsget.eof then
            for i = 0 to FResultCount - 1

				set FItemList(i) = new CCurrentStockItem

					FItemList(i).Fmakerid	    = rsget("makerid")
        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = db2html(rsget("itemoption"))
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
        			FItemList(i).Fdeliverytype  = rsget("deliverytype")
        			FItemList(i).Fmwdiv			= rsget("mwdiv")
					FItemList(i).Fsellyn		= rsget("sellyn")
					FItemList(i).Flimityn		= rsget("limityn")
					FItemList(i).Flimitno		= rsget("limitno")
					FItemList(i).Flimitsold		= rsget("limitsold")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

					''FItemList(i).FOldSystemCurrno = rsget("currno")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).Fdanjongyn  = rsget("danjongyn")
					FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")
					FItemList(i).FItemrackcode = rsget("itemrackcode")
					''FItemList(i).fprtidx = rsget("prtidx")
					FItemList(i).fsubitemrackcode = rsget("subitemrackcode")
					FItemList(i).FOnlineCurrentSellcash = rsget("sellcash")
                    FItemList(i).FOnlineCurrentBuycash = rsget("buycash")
					FItemList(i).Fsellcash = rsget("sellcash")
                    FItemList(i).Fbuycash = rsget("buycash")
					FItemList(i).Forgprice = rsget("orgprice")		'/'// 소비자가
					FItemList(i).FCentermwdiv = rsget("centermwdiv")
					FItemList(i).FlastIpgoDate = rsget("lastIpgoDate")

					FItemList(i).FpublicBarcode = rsget("publicBarcode")
					FItemList(i).FprevMonthSellCnt = rsget("prevMonthSellCnt")

					FItemList(i).FOffMwMargin = rsget("OffMwMargin")
					if IsNull(FItemList(i).FOffMwMargin) then
						FItemList(i).FOffMwMargin = ""
					end if

                    FItemList(i).Fitemgrade = rsget("itemgrade")
                    FItemList(i).Fagvstock = rsget("agvstock")

                rsget.movenext
            next
        end if
        rsget.close
        set cmd = Nothing
    end sub

    public sub GetItemSellStateList
        dim sqlstr, i, sqlsearch

        sqlstr = " select top 500 "
        sqlstr = sqlstr & "		i.makerid, s.itemgubun, i.itemid, IsNull(o.itemoption, '0000') itemoption "
        sqlstr = sqlstr & "		, i.sellyn, i.isusing, i.itemname, o.optionname "
        sqlstr = sqlstr & "		, IsNull(o.isusing, 'Y') optisusing, IsNull(o.optsellyn, 'Y') optsellyn "
        sqlstr = sqlstr & "		, s.totsysstock, s.realstock "
        ''sqlstr = sqlstr & "		, convert(varchar(10), i.sellSTDate, 121) as sellStartDate "
        sqlstr = sqlstr & "	from "
        sqlstr = sqlstr & "		[db_item].[dbo].[tbl_item] i "
        sqlstr = sqlstr & "		join [db_partner].[dbo].[tbl_partner] p on i.makerid = p.id "
        sqlstr = sqlstr & "		left join [db_item].[dbo].tbl_item_option o on i.itemid = o.itemid "
        sqlstr = sqlstr & "		left join [db_summary].[dbo].[tbl_current_logisstock_summary] s "
        sqlstr = sqlstr & "		on "
        sqlstr = sqlstr & "			1 = 1 "
        sqlstr = sqlstr & "			and s.itemgubun = '10' "
        sqlstr = sqlstr & "			and s.itemid = i.itemid "
        sqlstr = sqlstr & "			and s.itemoption = IsNull(o.itemoption, '0000') "
        sqlstr = sqlstr & "	where "
        sqlstr = sqlstr & "		1 = 1 "

		if (FRectMakerid<>"") then
		    sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "'"
        end if

        if (FRectItemIdArr<>"") then
            sqlstr = sqlstr + " and i.itemid in (" & FRectItemIdArr & ")"
        end if

        if (FRectMWDiv = "MW") then
            sqlstr = sqlstr + " and i.mwdiv in ('M', 'W')"
        elseif (FRectMWDiv <> "") then
            sqlstr = sqlstr + " and i.mwdiv = '" & FRectMWDiv & "'"
        end if

        sqlstr = sqlstr & "		and ( "
        sqlstr = sqlstr & "			(i.sellyn in ('N', 'S') or i.isusing = 'N') "
        sqlstr = sqlstr & "			or "
        sqlstr = sqlstr & "			(IsNull(o.isusing, 'Y') <> 'Y') "
        sqlstr = sqlstr & "		) "
        sqlstr = sqlstr & "		and s.realstock >= " & FRectlimitrealstock
        sqlstr = sqlstr & "		and p.tplcompanyid is NULL "
        sqlstr = sqlstr & " order by s.itemgubun, i.itemid, IsNull(o.itemoption, '0000') "

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.Open sqlStr, dbget

        FResultCount  = rsget.RecordCount
		FTotalCount =  FResultCount
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

					FItemList(i).Fmakerid	    = rsget("makerid")
        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemoption    = db2html(rsget("itemoption"))
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
                    FItemList(i).FitemoptionName      = db2html(rsget("optionname"))
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")

                    FItemList(i).Fsellyn		= rsget("sellyn")
					FItemList(i).Fisusing 		= rsget("isusing")
                    FItemList(i).Foptsellyn		= rsget("optsellyn")
					FItemList(i).Foptisusing 	= rsget("optisusing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close



    end sub

	'//admin/stock/brandcurrentstock.asp
	' 이 펑션을 수정할경우 GetCurrentStockByOnlineBrandNEW_notpaping 도 반드시 같이 수정해 주세요.
	public sub GetCurrentStockByOnlineBrandNEW
		dim sqlstr, i, sqlsearch, stockFieldName

		if (FRectMakerid<>"") then
		    sqlsearch = sqlsearch + " and i.makerid='" + CStr(FRectMakerid) + "'"
        end if

        if (FRectItemIdArr<>"") then
            sqlsearch = sqlsearch + " and i.itemid in (" & FRectItemIdArr & ")"
        end if

        if (FRectItemName<>"") then
            sqlsearch = sqlsearch + " and i.itemname like '%" & FRectItemName & "%'"
        end if

        if (FRectCD1<>"") then
            sqlsearch = sqlsearch + " and i.cate_large='" & FRectCD1 & "'"
        end if

        if (FRectCD2<>"") then
            sqlsearch = sqlsearch + " and i.cate_mid='" & FRectCD2 & "'"
        end if

        if (FRectCD3<>"") then
            sqlsearch = sqlsearch + " and i.cate_small='" & FRectCD3 & "'"
        end if

        ''반품 대상 상품(실사재고<>0)
        if (FRectReturnItemGubun="reton") then
            FRectOnlySellyn  = "N"
            ''FRectOnlyIsUsing = ""
            FRectDanjongyn   = "YM"
            FRectLimityn     = ""
            FRectMwDiv       = "MW"
            sqlsearch = sqlsearch + " and s.realstock <> 0"
        ''반품 완료 상품(실사재고=0)
        elseif (FRectReturnItemGubun="retfin") then
            FRectOnlySellyn  = "N"
            ''FRectOnlyIsUsing = ""
            FRectDanjongyn   = "YM"
            FRectLimityn     = ""
            FRectMwDiv       = "MW"
            sqlsearch = sqlsearch + " and s.realstock = 0"
        ''랙진열상품
        elseif (FRectReturnItemGubun="rackdisp") then
            sqlsearch = sqlsearch + " and ((i.sellyn <>'N') or (i.danjongyn in ('N','S')))"
            FRectOnlySellyn  = ""
            FRectOnlyIsUsing = ""
            FRectDanjongyn   = ""
            FRectLimityn     = ""
            'FRectMwDiv       = ""
        end if

        if FRectOnlySellyn="YS" then
            sqlsearch = sqlsearch + " and i.sellyn <>'N'"
        elseif FRectOnlySellyn="NN" then
            sqlsearch = sqlsearch + " and (i.sellyn = 'N' or IsNull(o.isusing, 'Y') = 'N') "
        elseif FRectOnlySellyn="NS" then
            sqlsearch = sqlsearch + " and (i.sellyn <> 'Y' or IsNull(o.isusing, 'Y') <> 'Y') "
        elseif FRectOnlySellyn<>"" then
            sqlsearch = sqlsearch + " and i.sellyn = '" + FRectOnlySellyn + "' "
        end if

        if (FRectOnlyIsUsing = "Y") then
            sqlsearch = sqlsearch + " and (i.isusing = 'Y' and IsNull(o.isusing, 'Y') = 'Y') "
        elseif (FRectOnlyIsUsing = "N") then
            sqlsearch = sqlsearch + " and not (i.isusing = 'Y' and IsNull(o.isusing, 'Y') = 'Y') "
        end if

        if FRectDanjongyn="SN" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'Y'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'N'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlsearch = sqlsearch + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectLimityn="Y0" then
            sqlsearch = sqlsearch + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlsearch = sqlsearch + " and i.limityn='" + FRectLimityn + "'"
        end if

		if FRectMwDiv="MW" then
            sqlsearch = sqlsearch + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

        if FRectCenterMWDiv="N" then
            sqlsearch = sqlsearch + " and si.centermwdiv is NULL"
        elseif FRectCenterMWDiv<>"" then
            sqlsearch = sqlsearch + " and si.centermwdiv='" + FRectCenterMWDiv + "'"
        end if

        if frectsoldout_gubun = "Y" then
    		sqlsearch = sqlsearch + " and ((i.sellyn<>'Y') or ((i.limityn<>'N') and (i.limitno-i.limitsold<1)))"
    	elseif frectsoldout_gubun = "N" then
    		sqlsearch = sqlsearch + " and s.availsysstock > 0"
    	else
			'
		end if

'		if FRectlimitrealstock="1UP" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1"
'		elseif FRectlimitrealstock="0DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 0"
'		elseif FRectlimitrealstock="20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 20"
'		elseif FRectlimitrealstock="1UP20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1 and s.realstock <= 20"
'		end if

		''두가지 선택 2014-11-04
		stockFieldName = "totsysstock"
		if (FRectStockType = "real") then
			stockFieldName = "realstock"
		end if
		if FRectlimitrealstock="1UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1"
		elseif FRectlimitrealstock="0DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 0"
		elseif FRectlimitrealstock="20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock="1UP20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock = "20UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 20"
		elseif FRectlimitrealstock="1UP3DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 3"
		elseif FRectlimitrealstock="MINUS" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " < 0"
		end If

		If (FRectStartDate <> "") And IsNumeric(FRectStartDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) >= " & FRectStartDate & " "
		End If

		If (FRectEndDate <> "") And IsNumeric(FRectEndDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) <= " & FRectEndDate & " "
		End If

		If FRectExcIts = "Y" Then
			''sqlsearch = sqlsearch + " and i.makerid not in ('ithinkso', 'conitale') "
            sqlsearch = sqlsearch + " and pp.tplcompanyid is NULL "
		End If

		if FRectDispCate<>"" Then
			sqlsearch = sqlsearch + " and s.itemgubun = '10' "
		    if LEN(FRectDispCate)>3 then
		         sqlsearch = sqlsearch + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
		    end if
			sqlsearch = sqlsearch + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        If FRectPurchasetype <> "" Then
            Select Case FRectPurchasetype
                Case "101"
                    sqlsearch = sqlsearch & " and pp.purchasetype in (4, 5, 6, 7, 8) "
                Case Else
                    sqlsearch = sqlsearch & " and pp.purchasetype = "& FRectPurchasetype &""
            End Select
        End If

        If FRectItemGrade <> "" Then
            select case FRectItemGrade
                case "AB":
                    sqlsearch = sqlsearch & " and s.itemgrade in ('A', 'B') "
                case "ABC":
                    sqlsearch = sqlsearch & " and s.itemgrade in ('A', 'B', 'C') "
                case else
                    sqlsearch = sqlsearch & " and s.itemgrade = '" & FRectItemGrade & "' "
            end select
        End If

        if FRectRackCode <> "" then
            sqlsearch = sqlsearch & " and Left(os.rackcodeByOption, " & Len(FRectRackCode) & ") = '" & FRectRackCode & "' "
        end if

        if FRectBulkStockGubun <> "" then
            if (FRectBulkStockGubun = "nul") then
                '// 입력이전
                sqlsearch = sqlsearch & " and DateDiff(day, IsNull(agvs.lastbulkstockdate, DateAdd(day, -30, getdate())), getdate()) > 5 "
            elseif (FRectBulkStockGubun = "err") then
                '// 벌크오차 있음
                sqlsearch = sqlsearch & " and DateDiff(day, IsNull(agvs.lastbulkstockdate, DateAdd(day, -30, getdate())), getdate()) <= 5 "
                sqlsearch = sqlsearch & " and (s.realstock + s.ipkumdiv5 + s.offconfirmno) <> (IsNull(agvs.bulkstock, 0) + IsNull(agvs.agvstock, 0)) "
            end if
        end if

        if FRectWarehouseCd <> "" then
            sqlsearch = sqlsearch & " and IsNull(agvs.warehouseCd, 'BLK') = '" & FRectWarehouseCd & "' "
        end if

        if FRectAgvStockGubun <> "" then
            if (FRectAgvStockGubun = "availdiff") then
                sqlsearch = sqlsearch & " and agvs.agvstock <> s.realstock "
            end if
            if (FRectAgvStockGubun = "ipkum5diff") then
                sqlsearch = sqlsearch & " and agvs.agvstock <> (s.realstock + s.ipkumdiv5 + s.offconfirmno) "
            end if
            if (FRectAgvStockGubun = "oneup") then
                sqlsearch = sqlsearch & " and agvs.agvstock > 0 "
            end if
            if (FRectAgvStockGubun = "zero") then
                sqlsearch = sqlsearch & " and agvs.agvstock = 0 "
            end if
            if (FRectAgvStockGubun = "minus") then
                sqlsearch = sqlsearch & " and agvs.agvstock < 0 "
            end if
        end if


        sqlstr = " select count(i.itemid) as cnt"
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option] o with (nolock) on i.itemid = o.itemid "
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
        sqlstr = sqlstr + " 	on s.itemid = i.itemid"
        sqlstr = sqlstr + "		and s.itemgubun ='10'"
        sqlStr = sqlStr & " 	and s.itemoption = IsNull(o.itemoption, '0000') "

		If FRectStartDate <> "" Or FRectEndDate <> "" then
			sqlstr = sqlstr + " 	LEFT JOIN [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p with (nolock) ON 1 = 1 "
			sqlstr = sqlstr + " 		AND p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
			sqlstr = sqlstr + " 		AND s.itemgubun = p.itemgubun "
			sqlstr = sqlstr + " 		AND s.itemid = p.itemid "
			sqlstr = sqlstr + " 		AND s.itemoption = p.itemoption "
		End if

        if (FRectCenterMWDiv<>"") then
            sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item si with (nolock) "
    		sqlStr = sqlstr + "     on si.itemgubun='10' and s.itemgubun=si.itemgubun"
    		sqlStr = sqlstr + "     and s.itemid=si.shopitemid"
    		sqlStr = sqlstr + "     and s.itemoption=si.itemoption"
	    end if

		If FRectPurchasetype <> "" or FRectExcIts <> "" Then
            sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner as pp with (nolock) on i.makerid = pp.id"
        End If

        if FRectRackCode <> "" then
            sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] os with (nolock) "
            sqlStr = sqlStr & " 		on "
            sqlStr = sqlStr & " 			1 = 1 "
            sqlStr = sqlStr & " 			and os.itemgubun = '10' "
            sqlStr = sqlStr & " 			and os.itemid = i.itemid "
            sqlStr = sqlStr & " 			and os.itemoption = IsNull(o.itemoption, '0000') "
        end if

        if FRectBulkStockGubun <> "" or FRectWarehouseCd <> "" or FRectAgvStockGubun <> "" then
            sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_current_agvstock_summary] agvs with (nolock) "
            sqlStr = sqlStr & " 		on "
            sqlStr = sqlStr & " 			1 = 1 "
            sqlStr = sqlStr & " 			and agvs.itemgubun = '10' "
            sqlStr = sqlStr & " 			and agvs.itemid = i.itemid "
            sqlStr = sqlStr & " 			and agvs.itemoption = IsNull(o.itemoption, '0000') "
        end if

        sqlstr = sqlstr + " where 1=1 " & sqlsearch

		'response.write sqlstr & "<Br>"
        ''response.end
        rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlstr = " exec [db_summary].[dbo].[usp_Ten_GetCurrentStockByOnlineBrand_LIST] " & FCurrPage & ", "& FPageSize & ", '" & FRectMakerid & "', '" & FRectItemIdArr & "' "
		sqlstr = sqlstr & ", '" & FRectCD1 & "'"
		sqlstr = sqlstr & ", '" & FRectCD2 & "'"
		sqlstr = sqlstr & ", '" & FRectCD3 & "'"
		sqlstr = sqlstr & ", '" & FRectOnlySellyn & "'"
		sqlstr = sqlstr & ", '" & FRectOnlyIsUsing & "'"
		sqlstr = sqlstr & ", '" & FRectDanjongyn & "'"
		sqlstr = sqlstr & ", '" & FRectLimityn & "'"
		sqlstr = sqlstr & ", '" & FRectMwDiv & "'"
		sqlstr = sqlstr & ", '" & FRectSoldOut_Gubun & "'"
		sqlstr = sqlstr & ", '" & FRectStockType & "'"
		sqlstr = sqlstr & ", '" & FRectlimitrealstock & "'"
		sqlstr = sqlstr & ", '" & FRectExcIts & "'"
		sqlstr = sqlstr & ", '" & FRectStartDate & "'"
		sqlstr = sqlstr & ", '" & FRectEndDate & "'"
		sqlstr = sqlstr & ", '" & FRectCenterMWDiv & "'"
		sqlstr = sqlstr & ", '" & FRectOrderBy & "'"
		sqlstr = sqlstr & ", '" & FRectReturnItemGubun & "'"
		sqlstr = sqlstr & ", '" & FRectDispCate & "'"
		sqlstr = sqlstr & ", '" & FRectPurchasetype & "'"
        sqlstr = sqlstr & ", '" & FRectItemGrade & "'"
        sqlstr = sqlstr & ", '" & FRectRackCode & "'"
        sqlstr = sqlstr & ", '" & FRectBulkStockGubun & "'"
        sqlstr = sqlstr & ", '" & FRectWarehouseCd & "'"
        sqlstr = sqlstr & ", '" & FRectAgvStockGubun & "'"
        sqlstr = sqlstr & ", '"& FRectItemName &"'"

		'response.write sqlstr & "<Br>"
		'response.end
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.Open sqlStr, dbget

        FResultCount  = rsget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

					FItemList(i).Fmakerid	    = rsget("makerid")
        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = db2html(rsget("itemoption"))
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
        			FItemList(i).Fdeliverytype  = rsget("deliverytype")
        			FItemList(i).Fmwdiv			= rsget("mwdiv")
					FItemList(i).Fsellyn		= rsget("sellyn")
					FItemList(i).Flimityn		= rsget("limityn")
					FItemList(i).Flimitno		= rsget("limitno")
					FItemList(i).Flimitsold		= rsget("limitsold")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

					''FItemList(i).FOldSystemCurrno = rsget("currno")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).Fdanjongyn  = rsget("danjongyn")
					FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")
					FItemList(i).FItemrackcode = rsget("itemrackcode")
					FItemList(i).fprtidx = rsget("prtidx")
					FItemList(i).fsubitemrackcode = rsget("subitemrackcode")
					FItemList(i).FOnlineCurrentSellcash = rsget("sellcash")
                    FItemList(i).FOnlineCurrentBuycash = rsget("buycash")
					FItemList(i).Fsellcash = rsget("sellcash")
                    FItemList(i).Fbuycash = rsget("buycash")
					FItemList(i).Forgprice = rsget("orgprice")		'/'// 소비자가
					FItemList(i).FCentermwdiv = rsget("centermwdiv")
					FItemList(i).FlastIpgoDate = rsget("lastIpgoDate")

					FItemList(i).FpublicBarcode = rsget("publicBarcode")
					FItemList(i).FprevMonthSellCnt = rsget("prevMonthSellCnt")

					FItemList(i).FOffMwMargin = rsget("OffMwMargin")
					if IsNull(FItemList(i).FOffMwMargin) then
						FItemList(i).FOffMwMargin = ""
					end if

                    FItemList(i).Fitemgrade = rsget("itemgrade")
                    FItemList(i).Fagvstock = rsget("agvstock")
                    FItemList(i).Fbulkstock = rsget("bulkstock")
                    FItemList(i).Flastbulkstockdate = rsget("lastbulkstockdate")
                    if IsNull(FItemList(i).Flastbulkstockdate) then
                        FItemList(i).Fbulkstock = NULL
                    elseif (DateDiff("d", FItemList(i).Flastbulkstockdate, Now()) > 5) then
                        FItemList(i).Fbulkstock = NULL
                    end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	'//admin/stock/brandcurrentstock.asp
	' 이 펑션을 수정할경우 GetCurrentStockByOnlineBrandNEW 도 반드시 같이 수정해 주세요.
	public sub GetCurrentStockByOnlineBrandNEW_notpaping
		dim sqlstr, i, sqlsearch, stockFieldName

		if (FRectMakerid<>"") then
		    sqlsearch = sqlsearch + " and i.makerid='" + CStr(FRectMakerid) + "'"
        end if

        if (FRectItemIdArr<>"") then
            sqlsearch = sqlsearch + " and i.itemid in (" & FRectItemIdArr & ")"
        end if

        if (FRectItemName<>"") then
            sqlsearch = sqlsearch + " and i.itemname like '%" & FRectItemName & "%'"
        end if

        if (FRectCD1<>"") then
            sqlsearch = sqlsearch + " and i.cate_large='" & FRectCD1 & "'"
        end if

        if (FRectCD2<>"") then
            sqlsearch = sqlsearch + " and i.cate_mid='" & FRectCD2 & "'"
        end if

        if (FRectCD3<>"") then
            sqlsearch = sqlsearch + " and i.cate_small='" & FRectCD3 & "'"
        end if

        ''반품 대상 상품(실사재고<>0)
        if (FRectReturnItemGubun="reton") then
            FRectOnlySellyn  = "N"
            ''FRectOnlyIsUsing = ""
            FRectDanjongyn   = "YM"
            FRectLimityn     = ""
            FRectMwDiv       = "MW"
            sqlsearch = sqlsearch + " and s.realstock <> 0"
        ''반품 완료 상품(실사재고=0)
        elseif (FRectReturnItemGubun="retfin") then
            FRectOnlySellyn  = "N"
            ''FRectOnlyIsUsing = ""
            FRectDanjongyn   = "YM"
            FRectLimityn     = ""
            FRectMwDiv       = "MW"
            sqlsearch = sqlsearch + " and s.realstock = 0"
        ''랙진열상품
        elseif (FRectReturnItemGubun="rackdisp") then
            sqlsearch = sqlsearch + " and ((i.sellyn <>'N') or (i.danjongyn in ('N','S')))"
            FRectOnlySellyn  = ""
            FRectOnlyIsUsing = ""
            FRectDanjongyn   = ""
            FRectLimityn     = ""
            'FRectMwDiv       = ""
        end if

        if FRectOnlySellyn="YS" then
            sqlsearch = sqlsearch + " and i.sellyn <>'N'"
        elseif FRectOnlySellyn="NN" then
            sqlsearch = sqlsearch + " and (i.sellyn = 'N' or IsNull(o.isusing, 'Y') = 'N') "
        elseif FRectOnlySellyn="NS" then
            sqlsearch = sqlsearch + " and (i.sellyn <> 'Y' or IsNull(o.isusing, 'Y') <> 'Y') "
        elseif FRectOnlySellyn<>"" then
            sqlsearch = sqlsearch + " and i.sellyn = '" + FRectOnlySellyn + "' "
        end if

        if (FRectOnlyIsUsing = "Y") then
            sqlsearch = sqlsearch + " and (i.isusing = 'Y' and IsNull(o.isusing, 'Y') = 'Y') "
        elseif (FRectOnlyIsUsing = "N") then
            sqlsearch = sqlsearch + " and not (i.isusing = 'Y' and IsNull(o.isusing, 'Y') = 'Y') "
        end if

        if FRectDanjongyn="SN" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'Y'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'N'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlsearch = sqlsearch + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectLimityn="Y0" then
            sqlsearch = sqlsearch + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlsearch = sqlsearch + " and i.limityn='" + FRectLimityn + "'"
        end if

		if FRectMwDiv="MW" then
            sqlsearch = sqlsearch + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

        if FRectCenterMWDiv="N" then
            sqlsearch = sqlsearch + " and si.centermwdiv is NULL"
        elseif FRectCenterMWDiv<>"" then
            sqlsearch = sqlsearch + " and si.centermwdiv='" + FRectCenterMWDiv + "'"
        end if

        if frectsoldout_gubun = "Y" then
    		sqlsearch = sqlsearch + " and ((i.sellyn<>'Y') or ((i.limityn<>'N') and (i.limitno-i.limitsold<1)))"
    	elseif frectsoldout_gubun = "N" then
    		sqlsearch = sqlsearch + " and s.availsysstock > 0"
    	else
			'
		end if

'		if FRectlimitrealstock="1UP" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1"
'		elseif FRectlimitrealstock="0DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 0"
'		elseif FRectlimitrealstock="20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock <= 20"
'		elseif FRectlimitrealstock="1UP20DOWN" then
'			sqlsearch = sqlsearch + " and s.realstock >= 1 and s.realstock <= 20"
'		end if

		''두가지 선택 2014-11-04
		stockFieldName = "totsysstock"
		if (FRectStockType = "real") then
			stockFieldName = "realstock"
		end if
		if FRectlimitrealstock="1UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1"
		elseif FRectlimitrealstock="0DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 0"
		elseif FRectlimitrealstock="20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock="1UP20DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 20"
		elseif FRectlimitrealstock = "20UP" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 20"
		elseif FRectlimitrealstock="1UP3DOWN" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " >= 1 and s." + CStr(stockFieldName) + " <= 3"
		elseif FRectlimitrealstock="MINUS" then
			sqlsearch = sqlsearch + " and s." + CStr(stockFieldName) + " < 0"
		end If

		If (FRectStartDate <> "") And IsNumeric(FRectStartDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) >= " & FRectStartDate & " "
		End If

		If (FRectEndDate <> "") And IsNumeric(FRectEndDate) Then
			sqlsearch = sqlsearch + " and DateDiff(m, p.lastIpgoDate + '-01', getdate()) <= " & FRectEndDate & " "
		End If

		If FRectExcIts = "Y" Then
			'' sqlsearch = sqlsearch + " and i.makerid not in ('ithinkso', 'conitale') "
            sqlsearch = sqlsearch + " and pp.tplcompanyid is NULL "
		End If

		if FRectDispCate<>"" Then
			sqlsearch = sqlsearch + " and s.itemgubun = '10' "
		    if LEN(FRectDispCate)>3 then
		         sqlsearch = sqlsearch + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
		    end if
			sqlsearch = sqlsearch + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        If FRectPurchasetype <> "" Then
            Select Case FRectPurchasetype
                Case "101"
                    sqlsearch = sqlsearch & " and pp.purchasetype in (4, 5, 6, 7, 8) "
                Case Else
                    sqlsearch = sqlsearch & " and pp.purchasetype = "& FRectPurchasetype &""
            End Select
        End If

        If FRectItemGrade <> "" Then
            select case FRectItemGrade
                case "AB":
                    sqlsearch = sqlsearch & " and s.itemgrade in ('A', 'B') "
                case "ABC":
                    sqlsearch = sqlsearch & " and s.itemgrade in ('A', 'B', 'C') "
                case else
                    sqlsearch = sqlsearch & " and s.itemgrade = '" & FRectItemGrade & "' "
            end select
        End If

        if FRectRackCode <> "" then
            sqlsearch = sqlsearch & " and Left(os.rackcodeByOption, " & Len(FRectRackCode) & ") = '" & FRectRackCode & "' "
        end if

        if FRectBulkStockGubun <> "" then
            if (FRectBulkStockGubun = "nul") then
                '// 입력이전
                sqlsearch = sqlsearch & " and DateDiff(day, IsNull(agvs.lastbulkstockdate, DateAdd(day, -30, getdate())), getdate()) > 5 "
            elseif (FRectBulkStockGubun = "err") then
                '// 벌크오차 있음
                sqlsearch = sqlsearch & " and DateDiff(day, IsNull(agvs.lastbulkstockdate, DateAdd(day, -30, getdate())), getdate()) <= 5 "
                sqlsearch = sqlsearch & " and (s.realstock + s.ipkumdiv5 + s.offconfirmno) <> (IsNull(agvs.bulkstock, 0) + IsNull(agvs.agvstock, 0)) "
            end if
        end if

        if FRectWarehouseCd <> "" then
            sqlsearch = sqlsearch & " and IsNull(agvs.warehouseCd, 'BLK') = '" & FRectWarehouseCd & "' "
        end if

        if FRectAgvStockGubun <> "" then
            if (FRectAgvStockGubun = "availdiff") then
                sqlsearch = sqlsearch & " and agvs.agvstock <> s.realstock "
            end if
            if (FRectAgvStockGubun = "ipkum5diff") then
                sqlsearch = sqlsearch & " and agvs.agvstock <> (s.realstock + s.ipkumdiv5 + s.offconfirmno) "
            end if
            if (FRectAgvStockGubun = "oneup") then
                sqlsearch = sqlsearch & " and agvs.agvstock > 0 "
            end if
            if (FRectAgvStockGubun = "zero") then
                sqlsearch = sqlsearch & " and agvs.agvstock = 0 "
            end if
            if (FRectAgvStockGubun = "minus") then
                sqlsearch = sqlsearch & " and agvs.agvstock < 0 "
            end if
        end if


        sqlstr = " select count(i.itemid) as cnt"
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option] o with (nolock) on i.itemid = o.itemid "
        sqlstr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
        sqlstr = sqlstr + " 	on s.itemid = i.itemid"
        sqlstr = sqlstr + "		and s.itemgubun ='10'"
        sqlStr = sqlStr & " 	and s.itemoption = IsNull(o.itemoption, '0000') "

		If FRectStartDate <> "" Or FRectEndDate <> "" then
			sqlstr = sqlstr + " 	LEFT JOIN [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] p with (nolock) ON 1 = 1 "
			sqlstr = sqlstr + " 		AND p.yyyymm = '" + CStr(Left(DateAdd("m",-1,Now()), 7)) + "' "
			sqlstr = sqlstr + " 		AND s.itemgubun = p.itemgubun "
			sqlstr = sqlstr + " 		AND s.itemid = p.itemid "
			sqlstr = sqlstr + " 		AND s.itemoption = p.itemoption "
		End if

        if (FRectCenterMWDiv<>"") then
            sqlstr = sqlstr + " left join [db_shop].[dbo].tbl_shop_item si with (nolock) "
    		sqlStr = sqlstr + "     on si.itemgubun='10' and s.itemgubun=si.itemgubun"
    		sqlStr = sqlstr + "     and s.itemid=si.shopitemid"
    		sqlStr = sqlstr + "     and s.itemoption=si.itemoption"
	    end if

		If FRectPurchasetype <> "" or FRectExcIts <> "" Then
            sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner as pp with (nolock) on i.makerid = pp.id"
        End If

        if FRectRackCode <> "" then
            sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] os with (nolock) "
            sqlStr = sqlStr & " 		on "
            sqlStr = sqlStr & " 			1 = 1 "
            sqlStr = sqlStr & " 			and os.itemgubun = '10' "
            sqlStr = sqlStr & " 			and os.itemid = i.itemid "
            sqlStr = sqlStr & " 			and os.itemoption = IsNull(o.itemoption, '0000') "
        end if

        if FRectBulkStockGubun <> "" or FRectWarehouseCd <> "" or FRectAgvStockGubun <> "" then
            sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_current_agvstock_summary] agvs with (nolock) "
            sqlStr = sqlStr & " 		on "
            sqlStr = sqlStr & " 			1 = 1 "
            sqlStr = sqlStr & " 			and agvs.itemgubun = '10' "
            sqlStr = sqlStr & " 			and agvs.itemid = i.itemid "
            sqlStr = sqlStr & " 			and agvs.itemoption = IsNull(o.itemoption, '0000') "
        end if

        sqlstr = sqlstr + " where 1=1 " & sqlsearch

		''response.write sqlstr & "<Br>"
        rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlstr = " exec [db_summary].[dbo].[usp_Ten_GetCurrentStockByOnlineBrand_LIST] " & FCurrPage & ", "& FPageSize & ", '" & FRectMakerid & "', '" & FRectItemIdArr & "' "
		sqlstr = sqlstr & ", '" & FRectCD1 & "'"
		sqlstr = sqlstr & ", '" & FRectCD2 & "'"
		sqlstr = sqlstr & ", '" & FRectCD3 & "'"
		sqlstr = sqlstr & ", '" & FRectOnlySellyn & "'"
		sqlstr = sqlstr & ", '" & FRectOnlyIsUsing & "'"
		sqlstr = sqlstr & ", '" & FRectDanjongyn & "'"
		sqlstr = sqlstr & ", '" & FRectLimityn & "'"
		sqlstr = sqlstr & ", '" & FRectMwDiv & "'"
		sqlstr = sqlstr & ", '" & FRectSoldOut_Gubun & "'"
		sqlstr = sqlstr & ", '" & FRectStockType & "'"
		sqlstr = sqlstr & ", '" & FRectlimitrealstock & "'"
		sqlstr = sqlstr & ", '" & FRectExcIts & "'"
		sqlstr = sqlstr & ", '" & FRectStartDate & "'"
		sqlstr = sqlstr & ", '" & FRectEndDate & "'"
		sqlstr = sqlstr & ", '" & FRectCenterMWDiv & "'"
		sqlstr = sqlstr & ", '" & FRectOrderBy & "'"
		sqlstr = sqlstr & ", '" & FRectReturnItemGubun & "'"
		sqlstr = sqlstr & ", '" & FRectDispCate & "'"
		sqlstr = sqlstr & ", '" & FRectPurchasetype & "'"
        sqlstr = sqlstr & ", '" & FRectItemGrade & "'"
        sqlstr = sqlstr & ", '" & FRectRackCode & "'"
        sqlstr = sqlstr & ", '" & FRectBulkStockGubun & "'"
        sqlstr = sqlstr & ", '" & FRectWarehouseCd & "'"
        sqlstr = sqlstr & ", '" & FRectAgvStockGubun & "'"

		'response.write sqlstr & "<Br>"
		'response.end
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.Open sqlStr, dbget

        FResultCount  = rsget.RecordCount
		FtotalCount =  rsget.RecordCount
		if (FResultCount<1) then FResultCount=0

		if not rsget.EOF then
			farrlist = rsget.getRows()
		end if
		rsget.Close
	end sub

	'마이너스재고상품 검색
	public sub GetCurrentStockByOnlineBrandMinus
		dim sqlstr, i

		sqlstr = "select top 1000 s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.isusing, i.mwdiv, "
		sqlstr = sqlstr + " IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing , v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		''''sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_const_day_stock c on s.itemgubun='10' and s.itemid=c.itemid and s.itemoption=c.itemoption"
		sqlstr = sqlstr + " where s.itemid = i.itemid and s.itemgubun ='10' "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

        if (FRectMWDiv <> "") then
                sqlstr = sqlstr + " and IsNull(i.mwdiv, 'Z') = '" + CStr(FRectMWDiv) + "' "
        end if

        if (FRectUseYN <> "") then
                sqlstr = sqlstr + " and IsNull(i.isusing, 'Y') = '" + CStr(FRectUseYN) + "' "
        end if

		if (FRectMakerid <> "") then
		        sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectKindDisplay <> "") then
		        if (FRectKindDisplay = "totsysstock") then
		                sqlstr = sqlstr + " and totsysstock<=" + CStr(FRectParameter) + " "
		        elseif (FRectKindDisplay = "availsysstock") then
		                sqlstr = sqlstr + " and availsysstock<=" + CStr(FRectParameter) + " "
		        elseif (FRectKindDisplay = "realstock") then
		                sqlstr = sqlstr + " and realstock<=" + CStr(FRectParameter) + " "
		        elseif (FRectKindDisplay = "diff") then
		                sqlstr = sqlstr + " and abs(availsysstock - realstock) >= abs(" + CStr(FRectParameter) + ") "
		        end if
		end if

		if (FRectKindSort <> "") then
		        if (FRectKindSort = "makerid") then
		                sqlstr = sqlstr + " order by i.makerid, s.itemgubun, s.itemid, s.itemoption "
		        elseif (FRectKindSort = "itemid") then
		                sqlstr = sqlstr + " order by s.itemgubun, s.itemid desc, s.itemoption "
		        elseif (FRectKindSort = "diff") then
		                sqlstr = sqlstr + " order by abs(availsysstock - realstock) desc "
		        end if
		end if

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
					FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

					''FItemList(i).FOldSystemCurrno = rsget("currno")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	'한정판매 검색
	'//admin/shopmaster/stock_limit.asp
	public sub GetCurrentStockByOnlineBrandLimit
		dim sqlstr, i, sqlsearch

        if FRectOnlySellyn="YS" then
            sqlsearch = sqlsearch + " and i.sellyn <>'N'"
        elseif FRectOnlySellyn<>"" then
            sqlsearch = sqlsearch + " and i.sellyn = '" + FRectOnlySellyn + "' "
            if (FRectOnlySellyn="Y") then
                sqlsearch = sqlsearch + " and ( ((s.itemoption='0000') and (i.limitno-i.limitsold>0)) or ((s.itemoption<>'0000') and (v.optlimitno-v.optlimitsold>0)) )"
            end if
        end if
		if (FRectLimitSoldOut = "on") then
			sqlsearch = sqlsearch + " and (i.LimitNo - i.LimitSold) = 0 "
		end if
        if (FRectDiffDiv = "number") then
			sqlsearch = sqlsearch + " and " + CStr(FRectParameter) + " < abs(s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2) -  (IsNULL(v.optlimitno - v.optlimitsold,i.LimitNo - i.LimitSold)) "
        elseif (FRectDiffDiv = "percent") then
			sqlsearch = sqlsearch + " and (s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2)<>0"
			sqlsearch = sqlsearch + " and " + CStr(FRectParameter) + " < (round((100 - abs( (IsNULL(v.optlimitno - v.optlimitsold,i.LimitNo - i.LimitSold))*100/(s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2))),0)) "
        elseif (FRectDiffDiv = "over") then
			sqlsearch = sqlsearch + " and IsNULL(v.optlimitno - v.optlimitsold,i.LimitNo - i.LimitSold)>(s.realstock + s.ipkumdiv5 + s.offconfirmno + s.ipkumdiv4 + s.ipkumdiv2)"
        end if
        if (FRectOnlyIsUsing<>"") then
			sqlsearch = sqlsearch + " and i.isusing = '" & FRectOnlyIsUsing & "' "
        end if
		if (FRectMakerid <> "") then
			sqlsearch = sqlsearch + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if
        if FRectMwDiv="MW" then
            sqlsearch = sqlsearch + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
		if (FRectitemid <> "") then
			sqlsearch = sqlsearch + " and i.itemid=" & FRectitemid & ""
		end if
		if (FRectSearchType = "R") then
			sqlsearch = sqlsearch + " and left(i.itemrackcode,2) >= '" & FRectFromRackcode2 &"' "
			sqlsearch = sqlsearch + " and left(i.itemrackcode,2) <= '" & FRectToRackcode2 &"' "
		else
			if Len(FRectRackCode)=2 then
				sqlsearch = sqlsearch + " and left(i.itemrackcode,2) = '" & FRectRackCode &"'"
			elseif (Len(FRectRackCode)>2) then
				sqlsearch = sqlsearch + " and i.itemrackcode = = '" & FRectRackCode &"'"		'// ???
			end if
		end If

		If FRectExcIts = "Y" Then
			sqlsearch = sqlsearch + " and i.makerid <> 'ithinkso' "
		End If


		sqlstr = "select count(*) as Totalcnt ,CEILING(CAST(Count(*) AS FLOAT)/"& Cstr(FPageSize) &")  as TotalPage "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlStr = sqlstr + " 	on s.itemid = i.itemid and s.itemgubun='10'" + " and NOT (i.optionCnt>0 and s.itemoption='0000')"  ''2014/06/25 추가 (문재)
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v "
		sqlStr = sqlstr + "     on s.itemid=v.itemid "
		sqlStr = sqlstr + "     and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where 1=1"
        sqlstr = sqlstr + " and i.limityn = 'Y' " & sqlsearch

		'response.write sqlstr & "<Br>"
		rsget.Open sqlStr,dbget,1
		IF not rsget.Eof Then
			FTotalCount 	= rsget("Totalcnt")
			FTotalPage 		= rsget("TotalPage")
		End IF
		rsget.Close

		sqlstr = "select top " + CStr(FCurrPage*FPageSize)
		sqlStr = sqlstr + " s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, i.itemrackcode, i.optionCnt "
		sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing,"
		sqlStr = sqlstr + " v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlstr + " join [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlStr = sqlstr + " 	on s.itemid = i.itemid and s.itemgubun ='10' " + " and NOT (i.optionCnt>0 and s.itemoption='0000')"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v "
		sqlStr = sqlstr + "     on s.itemid=v.itemid "
		sqlStr = sqlstr + "     and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where 1=1 "
        sqlstr = sqlstr + " and i.limityn = 'Y' " & sqlsearch

        if (FRectOrderBy="makerid") then
            sqlstr = sqlstr + " order by i.makerid"
        elseif (FRectOrderBy="itemid") then
            sqlstr = sqlstr + " order by i.itemid desc"
        elseif (FRectOrderBy="itemrackcode") then
            sqlstr = sqlstr + " order by i.itemrackcode"
        else
		    sqlstr = sqlstr + " order by s.realstock desc  "
		end if

		'response.write sqlstr & "<Br>"
		rsget.pagesize=FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount - FPageSize*(FCurrPage-1)
		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
                    FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")
                    FItemList(i).Fitemrackcode  = rsget("itemrackcode")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")
        			FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")
                    FItemList(i).Foptioncnt = rsget("optioncnt")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	'한정솔드아웃 검색
	public sub GetCurrentStockByOnlineBrandLimitSoldout
		dim sqlstr, i

		sqlstr = "select top " + CStr(FPageSize) + " s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, i.danjongyn "
		sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where s.itemid = i.itemid and s.itemgubun ='10' "
        sqlstr = sqlstr + " and i.limityn = 'Y' "

        if (FRectSearchMode="S1") then
            sqlstr = sqlstr + " and i.sellyn = 'S' and (i.LimitNo - i.LimitSold) > 0 "
        else
            sqlstr = sqlstr + " and i.sellyn = 'Y' and (i.LimitNo - i.LimitSold) <1 "
        end if
'        sqlstr = sqlstr + " and ("
'        sqlstr = sqlstr + "     (i.sellyn = 'Y' and (i.LimitNo - i.LimitSold) <1) "
'        sqlstr = sqlstr + " or"
'        sqlstr = sqlstr + "     (i.sellyn = 'S' and (i.LimitNo - i.LimitSold) >0) "
'        sqlstr = sqlstr + " )"

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		        sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		sqlstr = sqlstr + " order by s.realstock desc  "

		'response.write sqlstr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
                    FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")
                    FItemList(i).Fdanjongyn     = rsget("danjongyn")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")
        			FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	'단종상품검색(브랜드합계)
	public sub GetCurrentStockByOnlineBrandDanjong_GroupBrand
	    dim sqlstr, i

	    sqlstr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlstr + " i.makerid, count(s.itemid) as cnt"
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item i on s.itemid=i.itemid"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where s.itemgubun='10' "
		sqlstr = sqlstr + " and i.sellyn = 'N' "
        sqlstr = sqlstr + " and i.danjongyn in ('Y','M') "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

        if (FRectrealstocknotzero = "on") then
                sqlstr = sqlstr + " and s.realstock <> 0 "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

	    sqlstr = sqlstr + " group by i.makerid"
	    sqlstr = sqlstr + " order by cnt desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
        FTotalCount  = FResultCount
		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTurnOverBrand

                FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).FCnt       = rsget("cnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	'단종상품검색(반품대상상품)
	public sub GetCurrentStockByOnlineBrandDanjong
		dim sqlstr, i

        sqlstr = "select count(s.itemgubun) as cnt"
        sqlstr = sqlstr + " from [db_summary].[dbo].tbl_current_logisstock_summary s "
        sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item i on s.itemid=i.itemid"
        sqlstr = sqlstr + " where s.itemgubun='10' "
		sqlstr = sqlstr + " and i.sellyn = 'N' "
        sqlstr = sqlstr + " and i.danjongyn in ('Y','M') "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

        if (FRectrealstocknotzero = "on") then
                sqlstr = sqlstr + " and s.realstock <> 0 "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		    sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlstr = "select top " + CStr(FPageSize*FCurrPage) + " s.*,"
		sqlStr = sqlstr + " i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, i.danjongyn "
		sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlStr = sqlstr + " left join [db_item].[dbo].tbl_item i on s.itemid=i.itemid"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " where s.itemgubun='10' "
		sqlstr = sqlstr + " and i.sellyn = 'N' "
        sqlstr = sqlstr + " and i.danjongyn in ('Y','M') "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

        if (FRectrealstocknotzero = "on") then
                sqlstr = sqlstr + " and s.realstock <> 0 "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		    sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		    sqlstr = sqlstr + " order by s.itemid desc "
		else
			sqlstr = sqlstr + " order by s.realstock desc "
		end if

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

    			FItemList(i).Fitemgubun     = rsget("itemgubun")
    			FItemList(i).Fitemid        = rsget("itemid")
    			FItemList(i).Fitemname      = db2html(rsget("itemname"))
    			FItemList(i).Fitemoption    = rsget("itemoption")
    			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
                FItemList(i).Fisusing       = rsget("isusing")
    			FItemList(i).Flimityn       = rsget("limityn")
    			FItemList(i).FLimitNo       = rsget("LimitNo")
    			FItemList(i).FLimitSold     = rsget("LimitSold")
    			FItemList(i).Fsellyn        = rsget("sellyn")
    			FItemList(i).Fmakerid       = rsget("makerid")
    			FItemList(i).Fmwdiv         = rsget("mwdiv")
                FItemList(i).Fdanjongyn     = rsget("danjongyn")
    			FItemList(i).Fipgono        = rsget("ipgono")
    			FItemList(i).Freipgono      = rsget("reipgono")
    			FItemList(i).Ftotipgono     = rsget("totipgono")
    			FItemList(i).Foffchulgono   = rsget("offchulgono")
    			FItemList(i).Foffrechulgono = rsget("offrechulgono")
    			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
    			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
    			FItemList(i).Ftotchulgono   = rsget("totchulgono")
    			FItemList(i).Fsellno        = rsget("sellno")
    			FItemList(i).Fresellno      = rsget("resellno")
    			FItemList(i).Ftotsellno     = rsget("totsellno")
    			FItemList(i).Ferrcsno       = rsget("errcsno")
    			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
    			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
    			FItemList(i).Ferretcno      = rsget("erretcno")
    			FItemList(i).Ftoterrno      = rsget("toterrno")
    			FItemList(i).Ftotsysstock   = rsget("totsysstock")
    			FItemList(i).Favailsysstock = rsget("availsysstock")
    			FItemList(i).Frealstock     = rsget("realstock")
    			FItemList(i).Fsell7days     = rsget("sell7days")
    			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
    			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
    			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
    			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
    			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
    			FItemList(i).Foffjupno      = rsget("offjupno")
    			FItemList(i).Frequireno     = rsget("requireno")
    			FItemList(i).Fshortageno    = rsget("shortageno")
    			FItemList(i).Fpreorderno    = rsget("preorderno")
    			FItemList(i).Fpreordernofix    = rsget("preordernofix")
    			FItemList(i).Fmaxsellday    = rsget("maxsellday")
    			FItemList(i).Fimgsmall      = rsget("imgsmall")
    			FItemList(i).Fregdate       = rsget("regdate")
    			FItemList(i).Flastupdate    = rsget("lastupdate")
    			FItemList(i).Foptlimityn = rsget("optlimityn")
    			FItemList(i).Foptlimitno = rsget("optlimitno")
    			FItemList(i).Foptlimitsold = rsget("optlimitsold")

                if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

    ''일시 품절 상품 리스트
    public Sub GetImsiSoldOutList
        dim sqlstr, i

        sqlstr = "select count(i.itemid) as cnt"
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
		sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid "
		sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option_stock ot "
		sqlstr = sqlstr + "         on ot.itemgubun='10'  "
		sqlstr = sqlstr + "         and ot.itemid=i.itemid "
		sqlstr = sqlstr + "         and ot.itemoption=IsNULL(v.itemoption,'0000') "
		sqlstr = sqlstr + " where i.sellyn='S'"
        sqlstr = sqlstr + " and i.danjongyn in ('S','N')"

        if (FRectCd1<>"") then
            sqlstr = sqlstr + " and i.cate_large='" + FRectCd1 + "'"
        end if

        if (FRectOnlyIsUsing = "on") then
            sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		    sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectMWDiv = "MW") then
			sqlstr = sqlstr + " and (i.mwdiv='M' or i.mwdiv='W') "
		elseif (FRectMWDiv <> "") then
		    sqlstr = sqlstr + " and i.mwdiv='" + CStr(FRectMWDiv) + "' "
		end if

		if (FRectState<>"") then
		    sqlstr = sqlstr + " and ot.stockreipgodate is NULL"
		end if

		'rw sqlstr
		rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

        sqlstr = "select top " + CStr(FPageSize*FCurrPage) + " s.itemgubun, i.itemid, IsNULL(v.itemoption,'0000') as itemoption"
        sqlStr = sqlstr + " ,s.ipgono ,s.reipgono ,s.totipgono ,s.offchulgono ,s.offrechulgono ,s.etcchulgono ,s.etcrechulgono"
        sqlStr = sqlstr + " ,s.totchulgono ,s.sellno ,s.resellno ,s.totsellno ,s.errcsno ,s.errbaditemno ,s.errrealcheckno"
        sqlStr = sqlstr + " ,s.erretcno ,s.toterrno ,s.totsysstock ,s.availsysstock ,s.realstock ,s.sell7days ,s.offchulgo7days"
        sqlStr = sqlstr + " ,s.ipkumdiv5  ,s.ipkumdiv4  ,s.ipkumdiv2 ,s.offconfirmno ,s.offjupno ,s.requireno ,s.shortageno ,s.preorderno ,s.preordernofix "
        sqlStr = sqlstr + " ,s.maxsellday ,s.imgsmall ,s.regdate ,s.lastupdate "
        sqlStr = sqlstr + " ,i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, i.danjongyn "
		sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlStr = sqlstr + " ,ot.stockreipgodate"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
		sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid "
		sqlstr = sqlstr + "     left join [db_summary].[dbo].tbl_current_logisstock_summary s "
        sqlstr = sqlstr + "         on s.itemgubun ='10' "
        sqlstr = sqlstr + "         and s.itemid=i.itemid"
        sqlstr = sqlstr + "         and s.itemoption=IsNULL(v.itemoption,'0000') "
		sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option_stock ot "
		sqlstr = sqlstr + "         on ot.itemgubun='10'  "
		sqlstr = sqlstr + "         and ot.itemid=i.itemid "
		sqlstr = sqlstr + "         and ot.itemoption=IsNULL(v.itemoption,'0000') "
		sqlstr = sqlstr + " where i.sellyn='S'"
        sqlstr = sqlstr + " and i.danjongyn in ('S','N')"

        if (FRectCd1<>"") then
            sqlstr = sqlstr + " and i.cate_large='" + FRectCd1 + "'"
        end if

        if (FRectOnlyIsUsing = "on") then
            sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

        if (FRectItemID<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		    sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectMWDiv = "MW") then
			sqlstr = sqlstr + " and (i.mwdiv='M' or i.mwdiv='W') "
		elseif (FRectMWDiv <> "") then
		    sqlstr = sqlstr + " and i.mwdiv='" + CStr(FRectMWDiv) + "' "
		end if

		if (FRectState<>"") then
		    sqlstr = sqlstr + " and ot.stockreipgodate is NULL"
		end if

		sqlstr = sqlstr + " order by i.itemid desc"

		rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
                    FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fdanjongyn     = rsget("danjongyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")
        			FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")
                    FItemList(i).Fstockreipgodate = rsget("stockreipgodate")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

    public Sub GetQuickDlvItemList(isQuickValid)
        dim sqlstr, addSql, i

        if (FRectCd1<>"") then
            addSql = addSql + " and i.cate_large='" + FRectCd1 + "'"
        end if

        if (FRectOnlyIsUsing="on") then
            addSql = addSql + " and i.isusing = 'Y' "
        end if

        if (FRectItemID<>"") then
			addSql = addSql + " and i.itemid in (" + CStr(FRectItemID) + ")"
		end if

		if (FRectMakerid <> "") then
		    addSql = addSql + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectMWDiv = "MW") then
			addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W') "
		elseif (FRectMWDiv <> "") then
		    addSql = addSql + " and i.mwdiv='" + CStr(FRectMWDiv) + "' "
		end if

        sqlstr = "select count(*) as cnt"
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
        if (isQuickValid) then
    		sqlstr = sqlstr + "     Join db_item.dbo.tbl_item_quickdlv q"
    		sqlstr = sqlstr + "     on i.itemid=q.itemid"
	    ELSE
	        sqlstr = sqlstr + "     Join db_item.dbo.tbl_item_quickdlv_invalid q"
    		sqlstr = sqlstr + "     on i.itemid=q.itemid"
        end if

		sqlstr = sqlstr + " where 1=1" & addSql

		rsget.CursorLocation = adUseClient
        rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
		    FTotalCount = rsget("cnt")
		rsget.Close

        sqlstr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid"
        sqlStr = sqlstr + " ,i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, i.danjongyn , i.smallimage "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
		if (isQuickValid) then
    		sqlstr = sqlstr + "     Join db_item.dbo.tbl_item_quickdlv q"
    		sqlstr = sqlstr + "     on i.itemid=q.itemid"
	    ELSE
	        sqlstr = sqlstr + "     Join db_item.dbo.tbl_item_quickdlv_invalid q"
    		sqlstr = sqlstr + "     on i.itemid=q.itemid"
        end if

		sqlstr = sqlstr + " where 1=1" & addSql
		sqlstr = sqlstr + " order by i.itemid desc"

		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemoption    = "0000"
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
                    FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fdanjongyn     = rsget("danjongyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")

        			FItemList(i).Fimgsmall      = rsget("smallimage")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

    public Sub GetQuickDlvItemOptList(isQuickValid)
        dim sqlstr, addSql, i

        if (FRectCd1<>"") then
            addSql = addSql + " and i.cate_large='" + FRectCd1 + "'"
        end if

        if (FRectOnlyIsUsing = "on") then
            addSql = addSql + " and i.isusing = 'Y' "
        end if

        if (FRectItemID<>"") then
			addSql = addSql + " and i.itemid in (" + CStr(FRectItemID) + ")"
		end if

		if (FRectMakerid <> "") then
		    addSql = addSql + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectMWDiv = "MW") then
			addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W') "
		elseif (FRectMWDiv <> "") then
		    addSql = addSql + " and i.mwdiv='" + CStr(FRectMWDiv) + "' "
		end if

		if (FRectState<>"") then
		    addSql = addSql + " and ot.stockreipgodate is NULL"
		end if

        sqlstr = "select count(*) as cnt"
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
        if (isQuickValid) then
    		sqlstr = sqlstr + "     Join db_item.dbo.tbl_item_quickdlv q"
    		sqlstr = sqlstr + "     on i.itemid=q.itemid"
	    ELSE
	        sqlstr = sqlstr + "     Join db_item.dbo.tbl_item_quickdlv_invalid q"
    		sqlstr = sqlstr + "     on i.itemid=q.itemid"
        end if
        sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid and v.optsellyn='Y'"
		'sqlstr = sqlstr + "     left join [db_summary].[dbo].tbl_current_logisstock_summary s "
        'sqlstr = sqlstr + "         on s.itemgubun ='10' "
        'sqlstr = sqlstr + "         and s.itemid=i.itemid"
        'sqlstr = sqlstr + "         and s.itemoption=IsNULL(v.itemoption,'0000') "

		sqlstr = sqlstr + " where 1=1" & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
		    FTotalCount = rsget("cnt")
		rsget.Close

        sqlstr = "select top " + CStr(FPageSize*FCurrPage) + " s.itemgubun, i.itemid, IsNULL(v.itemoption,'0000') as itemoption"
        sqlStr = sqlstr + " ,s.ipgono ,s.reipgono ,s.totipgono ,s.offchulgono ,s.offrechulgono ,s.etcchulgono ,s.etcrechulgono"
        sqlStr = sqlstr + " ,s.totchulgono ,s.sellno ,s.resellno ,s.totsellno ,s.errcsno ,s.errbaditemno ,s.errrealcheckno"
        sqlStr = sqlstr + " ,s.erretcno ,s.toterrno ,s.totsysstock ,s.availsysstock ,s.realstock ,s.sell7days ,s.offchulgo7days"
        sqlStr = sqlstr + " ,s.ipkumdiv5  ,s.ipkumdiv4  ,s.ipkumdiv2 ,s.offconfirmno ,s.offjupno ,s.requireno ,s.shortageno ,s.preorderno ,s.preordernofix "
        sqlStr = sqlstr + " ,s.maxsellday ,s.imgsmall ,s.regdate ,s.lastupdate "
        sqlStr = sqlstr + " ,i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, i.danjongyn "
		sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
		if (isQuickValid) then
    		sqlstr = sqlstr + "     Join db_item.dbo.tbl_item_quickdlv q"
    		sqlstr = sqlstr + "     on i.itemid=q.itemid"
	    ELSE
	        sqlstr = sqlstr + "     Join db_item.dbo.tbl_item_quickdlv_invalid q"
    		sqlstr = sqlstr + "     on i.itemid=q.itemid"
        end if
		sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid and v.optsellyn='Y'"
		sqlstr = sqlstr + "     left join [db_summary].[dbo].tbl_current_logisstock_summary s "
        sqlstr = sqlstr + "         on s.itemgubun ='10' "
        sqlstr = sqlstr + "         and s.itemid=i.itemid"
        sqlstr = sqlstr + "         and s.itemoption=IsNULL(v.itemoption,'0000') "

		sqlstr = sqlstr + " where 1=1" & addSql
		sqlstr = sqlstr + " order by i.itemid desc"

		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
                    FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fdanjongyn     = rsget("danjongyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

       			    FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	'전시판매설정 검색(물류센터) '신상(90일) 표시안함 추가
	public sub GetCurrentStockByOnlineBrandDispSell
		dim sqlstr, i, sqlsearch

		If FRectExcIts = "Y" Then
			sqlsearch = sqlsearch & " and i.makerid <> 'ithinkso' "
		End If
		If FRectTplGubun<>"" then
			if (FRectTplGubun = "3X") then
				sqlsearch = sqlsearch + " 	and IsNull(pp.tplcompanyid, '') = '' "
			else
				sqlsearch = sqlsearch + " 	and IsNull(pp.tplcompanyid, '') = '" + CStr(FRectTplGubun) + "' "
			end if
		end If
		if (FRectDiffDiv = "sellSlimit1") then
        	sqlsearch = sqlsearch & " and i.sellyn = 'S' "
        	sqlsearch = sqlsearch & " and (i.LimitNo - i.LimitSold) > 0 "
        elseif (FRectDiffDiv = "sellSlimit2") then
        	sqlsearch = sqlsearch & " and i.sellyn = 'S' "
        	sqlsearch = sqlsearch & " and ((s.realstock + (s.ipkumdiv5 + s.offconfirmno)) + s.ipkumdiv4 + s.ipkumdiv2) > 0 "
        elseif (FRectDiffDiv = "sellN") then
        	sqlsearch = sqlsearch & " and i.sellyn = 'N' "
			sqlsearch = sqlsearch & " and ((s.realstock + (s.ipkumdiv5 + s.offconfirmno)) + s.ipkumdiv4 + s.ipkumdiv2) > 0 "
        	sqlsearch = sqlsearch & " and datediff(d,i.regdate,getdate())>=90"
        elseif (FRectDiffDiv = "sellY0") then
            sqlsearch = sqlsearch & " and i.sellyn = 'Y' and i.limityn='Y' and (i.LimitNo - i.LimitSold) <1 "
        end if

        if (FRectOnlyIsUsing = "Y") then
            sqlsearch = sqlsearch & " and i.isusing = 'Y' "
        end if
        if (FRectIsSellStart = "Y") then
            sqlsearch = sqlsearch & " and i.SellStDate is not null "
        elseif (FRectIsSellStart = "N") then
            sqlsearch = sqlsearch & " and i.SellStDate is null "
        end if

        if (FRectItemID<>"") then
			sqlsearch = sqlsearch & " and i.itemid=" + CStr(FRectItemID)
		end if

		if (FRectMakerid <> "") then
		    sqlsearch = sqlsearch & " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectMWDiv = "MW") then
			sqlsearch = sqlsearch & " and (i.mwdiv='M' or i.mwdiv='W') "
		elseif (FRectMWDiv <> "") then
		    sqlsearch = sqlsearch & " and i.mwdiv='" + CStr(FRectMWDiv) + "' "
		end if
        If FRectPurchasetype <> "" Then
            Select Case FRectPurchasetype
                Case "101"
                    sqlsearch = sqlsearch & " and pp.purchasetype in (4, 5, 6, 7, 8) "
                Case Else
                    sqlsearch = sqlsearch & " and pp.purchasetype = "& FRectPurchasetype &""
            End Select
        End If

		if FRectDispCate<>"" then
		    if LEN(FRectDispCate)>3 then
		         sqlsearch = sqlsearch + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'" ''2015/03/27추가
		    end if
			sqlsearch = sqlsearch + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item with (nolock) where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if
		if (FRectStockMwDiv <> "") then
			if (FRectStockMwDiv = "X") then
				sqlsearch = sqlsearch + " and IsNull(L.LastMwdiv, '') not in ('M', 'W') "
			else
				sqlsearch = sqlsearch + " and isnull(L.LastMwdiv,'') = '" & CStr(FRectStockMwDiv) & "' "
			end if
		end if

		sqlstr = "select top 1000"
		sqlStr = sqlstr & " s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.mwdiv, i.sellSTDate "
		sqlStr = sqlstr & " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
		sqlStr = sqlstr & " , Case "
		sqlStr = sqlstr & " 	When i.sellyn='Y' and isNull(v.optsellyn,'Y')='Y' and isNull(v.isusing,'Y')='Y' then 'Y' "
		sqlStr = sqlstr & " 	When i.sellyn='Y' and (isNull(v.optsellyn,'Y')='N' or isNull(v.isusing,'Y')='N') then 'N' "
		sqlStr = sqlstr & " 	Else i.sellyn "
		sqlStr = sqlstr & " end as sellyn "
		sqlStr = sqlstr & " , isnull(os.RackcodeByOption,i.itemrackcode) as itemrackcode"
		sqlStr = sqlstr & " , isnull(os.subRackcodeByOption,'') as subRackcodeByOption"
		sqlstr = sqlstr & " from [db_item].[dbo].tbl_item i with (nolock)"
		sqlstr = sqlstr & " join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
		sqlstr = sqlstr & " 	on i.itemid = s.itemid"
		sqlstr = sqlstr & " 	and s.itemgubun ='10'"
		sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item_option v with (nolock)"
		sqlstr = sqlstr & " 	on s.itemid=v.itemid"
		sqlstr = sqlstr & " 	and s.itemoption=v.itemoption "
		sqlstr = sqlstr & " left join db_item.dbo.tbl_item_option_stock os with(noLock)"
		sqlstr = sqlstr & " 	on os.itemgubun='10'"
		sqlstr = sqlstr & " 	and s.itemgubun = os.itemgubun"
		sqlstr = sqlstr & " 	and s.itemid = os.itemid"
		sqlstr = sqlstr & " 	and s.itemoption = os.itemoption"
		sqlstr = sqlstr & " LEFT JOIN db_partner.dbo.tbl_partner as pp with (nolock)"
		sqlstr = sqlstr & " 	on i.makerid = pp.id"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary L with (nolock)"
		sqlStr = sqlStr + "     on L.yyyymm=convert(Varchar(7),getdate(),121)"
		sqlStr = sqlStr + "     and  L.itemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and L.itemid=s.itemid"
		sqlStr = sqlStr + "     and L.itemoption=s.itemoption"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by s.realstock desc"

		'response.write sqlstr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
                    FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")
        			FItemList(i).Foptlimityn = rsget("optlimityn")
					FItemList(i).Foptlimitno = rsget("optlimitno")
					FItemList(i).Foptlimitsold = rsget("optlimitsold")
					FItemList(i).FItemrackcode = rsget("itemrackcode")
					FItemList(i).FItemsubrackcode = rsget("subRackcodeByOption")
					FItemList(i).FsellStdate = rsget("sellStdate")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	'미판매재고상품 검색
	public sub GetCurrentStockByOnlineBrandNoSellWithStock
		dim sqlstr, i

		sqlstr = "select top 500 s.*, i.itemname, i.isusing, i.makerid, i.limityn, i.LimitNo, i.LimitSold, i.sellyn, i.mwdiv, IsNULL(v.optionname,'') as codeview, c.currno "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v on s.itemid=v.itemid and s.itemoption=v.itemoption "
		sqlstr = sqlstr + " left join [db_storage].[dbo].tbl_const_day_stock c on s.itemgubun='10' and s.itemid=c.itemid and s.itemoption=c.itemoption"
		sqlstr = sqlstr + " where s.itemid = i.itemid and s.itemgubun ='10' "
        sqlstr = sqlstr + " and i.limityn = 'Y' "

        if (FRectSellYN <> "") then
                sqlstr = sqlstr + " and i.sellyn = '" + FRectSellYN + "' "
        end if

        sqlstr = sqlstr + " and (s.realstock > " + CStr(FRectParameter) + ") "

        if (FRectOnlyIsUsing = "on") then
                sqlstr = sqlstr + " and i.isusing = 'Y' "
        end if

		if (FRectMakerid <> "") then
		        sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if (FRectKindSort <> "") then
		        if (FRectKindSort = "makerid") then
		                sqlstr = sqlstr + " order by i.makerid, s.itemgubun, s.itemid, s.itemoption "
		        elseif (FRectKindSort = "itemid") then
		                sqlstr = sqlstr + " order by s.itemgubun, s.itemid desc, s.itemoption "
		        end if
		end if

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

        			FItemList(i).Fitemgubun     = rsget("itemgubun")
        			FItemList(i).Fitemid        = rsget("itemid")
        			FItemList(i).Fitemname      = db2html(rsget("itemname"))
        			FItemList(i).Fitemoption    = rsget("itemoption")
        			FItemList(i).FitemoptionName= db2html(rsget("codeview"))
					FItemList(i).Fisusing       = rsget("isusing")
        			FItemList(i).Flimityn       = rsget("limityn")
        			FItemList(i).FLimitNo       = rsget("LimitNo")
        			FItemList(i).FLimitSold     = rsget("LimitSold")
        			FItemList(i).Fsellyn        = rsget("sellyn")
        			FItemList(i).Fmakerid       = rsget("makerid")
        			FItemList(i).Fmwdiv         = rsget("mwdiv")
        			FItemList(i).Fipgono        = rsget("ipgono")
        			FItemList(i).Freipgono      = rsget("reipgono")
        			FItemList(i).Ftotipgono     = rsget("totipgono")
        			FItemList(i).Foffchulgono   = rsget("offchulgono")
        			FItemList(i).Foffrechulgono = rsget("offrechulgono")
        			FItemList(i).Fetcchulgono   = rsget("etcchulgono")
        			FItemList(i).Fetcrechulgono = rsget("etcrechulgono")
        			FItemList(i).Ftotchulgono   = rsget("totchulgono")
        			FItemList(i).Fsellno        = rsget("sellno")
        			FItemList(i).Fresellno      = rsget("resellno")
        			FItemList(i).Ftotsellno     = rsget("totsellno")
        			FItemList(i).Ferrcsno       = rsget("errcsno")
        			FItemList(i).Ferrbaditemno  = rsget("errbaditemno")
        			FItemList(i).Ferrrealcheckno= rsget("errrealcheckno")
        			FItemList(i).Ferretcno      = rsget("erretcno")
        			FItemList(i).Ftoterrno      = rsget("toterrno")
        			FItemList(i).Ftotsysstock   = rsget("totsysstock")
        			FItemList(i).Favailsysstock = rsget("availsysstock")
        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fsell7days     = rsget("sell7days")
        			FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Foffjupno      = rsget("offjupno")
        			FItemList(i).Frequireno     = rsget("requireno")
        			FItemList(i).Fshortageno    = rsget("shortageno")
        			FItemList(i).Fpreorderno    = rsget("preorderno")
        			FItemList(i).Fpreordernofix    = rsget("preordernofix")
        			FItemList(i).Fmaxsellday    = rsget("maxsellday")
        			FItemList(i).Fimgsmall      = rsget("imgsmall")
        			FItemList(i).Fregdate       = rsget("regdate")
        			FItemList(i).Flastupdate    = rsget("lastupdate")

                    if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                    if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

					FItemList(i).FOldSystemCurrno = rsget("currno")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	public sub GetMonthly_Logisstock_Summary()
		dim sqlstr, i

		sqlstr = "select top 300 s.yyyymm"
		sqlstr = sqlstr + " ,s.itemgubun, s.itemid, s.itemoption, s.ipgono, s.reipgono, s.totipgono, s.offchulgono, s.offrechulgono, s.etcchulgono, s.etcrechulgono, s.totchulgono"
		sqlstr = sqlstr + " ,s.sellno, s.resellno, s.totsellno,s.errcsno ,s.errbaditemno, s.errrealcheckno, s.erretcno, s.toterrno"
		sqlstr = sqlstr + " ,s.offsellno, s.totsysstock,s.availsysstock, s.realstock, convert(varchar(19),s.regdate,21) as regdate, convert(varchar(19),s.lastupdate,21) as lastupdate"
		sqlstr = sqlstr + " , a.lastmwdiv, a.lastbuyprice, a.avgipgoprice, a.totsysstock as lasttotsysstock"
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_monthly_logisstock_summary s"
		sqlstr = sqlstr + " left join  [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary a"
    	sqlstr = sqlstr + "     on s.yyyymm=a.yyyymm"
    	sqlstr = sqlstr + "     and s.itemgubun=a.itemgubun"
    	sqlstr = sqlstr + "     and s.itemid=a.itemid"
    	sqlstr = sqlstr + "     and s.itemoption=a.itemoption"
		sqlstr = sqlstr + " where s.yyyymm<'" + FRectYYYYMM + "'"
		sqlstr = sqlstr + " and s.itemgubun='" + FRectItemGubun + "'"
		sqlstr = sqlstr + " and s.itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and s.itemoption='" + CStr(FRectItemOption) + "'"
		sqlstr = sqlstr + " order by s.yyyymm"

		rsget.CursorLocation = adUseClient
        rsget.Open sqlstr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSummaryItemStockItem

				FItemList(i).Fyyyymm        = rsget("yyyymm")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fipgono          = rsget("ipgono")
				FItemList(i).Freipgono        = rsget("reipgono")
				FItemList(i).Ftotipgono       = rsget("totipgono")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				FItemList(i).Foffrechulgono   = rsget("offrechulgono")
				FItemList(i).Fetcchulgono     = rsget("etcchulgono")
				FItemList(i).Fetcrechulgono   = rsget("etcrechulgono")
				FItemList(i).Ftotchulgono     = rsget("totchulgono")
				FItemList(i).Fsellno          = -1*rsget("sellno")
				FItemList(i).Fresellno        = -1*rsget("resellno")
				FItemList(i).Ftotsellno       = -1*rsget("totsellno")
				FItemList(i).Ferrcsno         = rsget("errcsno")
				FItemList(i).Ferrbaditemno    = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno  = rsget("errrealcheckno")
				FItemList(i).Ferretcno        = rsget("erretcno")
				FItemList(i).Ftoterrno        = rsget("toterrno")
				FItemList(i).Foffsellno		  = rsget("offsellno")
				FItemList(i).Ftotsysstock     = rsget("totsysstock")
				FItemList(i).Favailsysstock   = rsget("availsysstock")
				FItemList(i).Frealstock       = rsget("realstock")
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Flastupdate      = rsget("lastupdate")
				FItemList(i).Flastmwdiv         = rsget("lastmwdiv")
                FItemList(i).Flastbuyprice      = rsget("lastbuyprice")
                FItemList(i).Favgipgoprice      = rsget("avgipgoprice")
                FItemList(i).Flasttotsysstock   = rsget("lasttotsysstock")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

    public sub getLastMonthStock()
		dim sqlstr, i

		sqlstr = "select top 1 * from [db_summary].[dbo].tbl_Last_monthly_logisstock"
		sqlstr = sqlstr + " where itemgubun='" + FRectItemGubun + "'"
		sqlstr = sqlstr + " and itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and itemoption='" + CStr(FRectItemOption) + "'"

		rsget.CursorLocation = adUseClient
        rsget.Open sqlstr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		i=0
		if  not rsget.EOF  then
			set FOneItem = new CSummaryItemStockItem

			FOneItem.Fyyyymm        = rsget("lastyyyymm")
			FOneItem.Fitemgubun       = rsget("itemgubun")
			FOneItem.Fitemid          = rsget("itemid")
			FOneItem.Fitemoption      = rsget("itemoption")
			FOneItem.Fipgono          = rsget("ipgono")
			FOneItem.Freipgono        = rsget("reipgono")
			FOneItem.Ftotipgono       = rsget("totipgono")
			FOneItem.Foffchulgono     = rsget("offchulgono")
			FOneItem.Foffrechulgono   = rsget("offrechulgono")
			FOneItem.Fetcchulgono     = rsget("etcchulgono")
			FOneItem.Fetcrechulgono   = rsget("etcrechulgono")
			FOneItem.Ftotchulgono     = rsget("totchulgono")
			FOneItem.Fsellno          = -1*rsget("sellno")
			FOneItem.Fresellno        = -1*rsget("resellno")
			FOneItem.Ftotsellno       = -1*rsget("totsellno")
			FOneItem.Ferrcsno         = rsget("errcsno")
			FOneItem.Ferrbaditemno    = rsget("errbaditemno")
			FOneItem.Ferrrealcheckno  = rsget("errrealcheckno")
			FOneItem.Ferretcno        = rsget("erretcno")
			FOneItem.Ftoterrno        = rsget("toterrno")
			FOneItem.Foffsellno		  = rsget("offsellno")
			FOneItem.Ftotsysstock     = rsget("totsysstock")
			FOneItem.Favailsysstock   = rsget("availsysstock")
			FOneItem.Frealstock       = rsget("realstock")
			FOneItem.Fregdate         = rsget("regdate")
			FOneItem.Flastupdate      = rsget("lastupdate")
		else
		    set FOneItem = new CSummaryItemStockItem
		    FOneItem.Fitemgubun       = FRectItemGubun
			FOneItem.Fitemid          = FRectItemID
			FOneItem.Fitemoption      = FRectItemOption
		    FOneItem.Fipgono          = 0
			FOneItem.Freipgono        = 0
			FOneItem.Ftotipgono       = 0
			FOneItem.Foffchulgono     = 0
			FOneItem.Foffrechulgono   = 0
			FOneItem.Fetcchulgono     = 0
			FOneItem.Fetcrechulgono   = 0
			FOneItem.Ftotchulgono     = 0
			FOneItem.Fsellno          = 0
			FOneItem.Fresellno        = 0
			FOneItem.Ftotsellno       = 0
			FOneItem.Ferrcsno         = 0
			FOneItem.Ferrbaditemno    = 0
			FOneItem.Ferrrealcheckno  = 0
			FOneItem.Ferretcno        = 0
			FOneItem.Ftoterrno        = 0
			FOneItem.Foffsellno		  = 0
			FOneItem.Ftotsysstock     = 0
			FOneItem.Favailsysstock   = 0
			FOneItem.Frealstock       = 0
			FOneItem.Fregdate         = 0
			FOneItem.Flastupdate      = 0
		end if
		rsget.close
	end sub

    '정리대상 상품
	public sub GetItemListForOut()
		dim sqlstr, i, yyyymm1, yyyymm2, pre3yyyymmdd
        dim whereDetail

		if (FRectYYYYMM = "") then
        	yyyymm1 = CStr(dateadd("m" ,-1, now()))
        	yyyymm1 = Left(yyyymm1,7)

        	yyyymm2 = CStr(dateadd("m" ,-2, now()))
        	yyyymm2 = Left(yyyymm2,7)

        	pre3yyyymmdd = Left(CStr(dateadd("m" ,-4, now())),7) + "-01"
		else
	        yyyymm1 = CStr(dateadd("m" ,-1, (FRectYYYYMM + "-01")))
	        yyyymm1 = Left(yyyymm1,7)

	        yyyymm2 = CStr(dateadd("m" ,-2, (FRectYYYYMM + "-01")))
	        yyyymm2 = Left(yyyymm2,7)

	        pre3yyyymmdd = Left(CStr(dateadd("m" ,-4, (FRectYYYYMM + "-01"))),7) + "-01"
		end if

		if FRectSearchMode="out2" then
			yyyymm2 = "2001-01-01"
		end if

		whereDetail = " and i.itemid<>0"

		if FRectMakerid<>"" then
			whereDetail = whereDetail + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		whereDetail = whereDetail + " and i.mwdiv<>'U' "

        if FRectCd1<>"" then
        	whereDetail = whereDetail + " 	and i.cate_large='" + FRectCd1 + "'"
        end if
        if FRectCd2<>"" then
        	whereDetail = whereDetail + " 	and i.cate_mid='" + FRectCd2 + "'"
        end if
        if FRectCd3<>"" then
        	whereDetail = whereDetail + " 	and i.cate_small='" + FRectCd3 + "'"
        end if

        if FRectOnlySellyn="YS" then
            whereDetail = whereDetail+ " and i.sellyn <>'N'"
        elseif FRectOnlySellyn<>"" then
        	whereDetail = whereDetail + " 	and i.sellyn = '" & FRectOnlySellyn & "' "
        end if

        if (FRectOnlyIsUsing <>"") then
        	whereDetail = whereDetail + " 	and i.isusing = '" & FRectOnlyIsUsing & "' "
        end if

        if FRectDanjongyn="SN" then
            whereDetail = whereDetail + " and i.danjongyn<>'Y'"
            whereDetail = whereDetail + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            whereDetail = whereDetail + " and i.danjongyn<>'N'"
            whereDetail = whereDetail + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            whereDetail = whereDetail + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectLimityn="Y0" then
            whereDetail = whereDetail + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            whereDetail = whereDetail + " and i.limityn='" + FRectLimityn + "'"
        end if

		whereDetail = whereDetail + " and i.regdate < '" + pre3yyyymmdd + "' "

        sqlstr = " select count(*) as cnt "
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i, "
        sqlstr = sqlstr + " [db_summary].[dbo].tbl_current_logisstock_summary c "

        if FRectSearchMode="out2" then
            sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary "
            sqlstr = sqlstr + " s on s.yyyymm='" & yyyymm1 & "' and c.itemgubun=s.itemgubun and c.itemid=s.itemid and c.itemoption=s.itemoption "
        else
            sqlstr = sqlstr + " left join ( "
            sqlstr = sqlstr + " 	select s.itemgubun, s.itemid, s.itemoption "
            sqlstr = sqlstr + " 	,sum(s.sellno) as sellno, sum(s.offchulgono) as offchulgono "
            sqlstr = sqlstr + " 	from [db_summary].[dbo].tbl_monthly_logisstock_summary s "
            sqlstr = sqlstr + " 	where s.yyyymm <= '" + yyyymm1 + "' "
            sqlstr = sqlstr + " 	and s.yyyymm >= '" + yyyymm2 + "' "
            sqlstr = sqlstr + " 	and s.itemgubun='10' "
            sqlstr = sqlstr + " 	group by s.itemgubun, s.itemid, s.itemoption "
            sqlstr = sqlstr + " ) s on c.itemgubun=s.itemgubun and c.itemid=s.itemid and c.itemoption=s.itemoption "
        end if

        sqlstr = sqlstr + " where c.itemgubun='10' "
        sqlstr = sqlstr + " and c.itemid=i.itemid"

        sqlstr = sqlstr +  whereDetail

        sqlstr = sqlstr + " and IsNULL(s.sellno,0) <= 0"
        sqlstr = sqlstr + " and IsNULL(s.offchulgono,0) >= 0 "

		'response.write sqlstr
		rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

        sqlstr = " select top " + CStr(FPageSize*FCurrPage) + " c.itemgubun,c.itemid, c.itemoption,IsNULL(s.sellno,0) as sellno, IsNULL(s.offchulgono,0) as offchulgono"
        sqlStr = sqlstr + " , i.makerid, i.itemname, i.smallimage, i.deliverytype, i.mwdiv, i.limityn, i.limitno, i.limitsold, i.sellyn, i.isusing, i.danjongyn"
        sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as itemoptionname, c.errbaditemno, c.realstock, c.toterrno "
     	sqlStr = sqlstr + " ,IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i, "
        sqlstr = sqlstr + " [db_summary].[dbo].tbl_current_logisstock_summary c"

        if FRectSearchMode="out2" then
            sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary "
            sqlstr = sqlstr + " s on s.yyyymm='" & yyyymm1 & "' and c.itemgubun=s.itemgubun and c.itemid=s.itemid and c.itemoption=s.itemoption "
        else
            sqlstr = sqlstr + " left join ( "
            sqlstr = sqlstr + " 	select s.itemgubun, s.itemid, s.itemoption, sum(s.sellno) as sellno, sum(s.offchulgono) as offchulgono "
            sqlstr = sqlstr + " 	from [db_summary].[dbo].tbl_monthly_logisstock_summary s "
            sqlstr = sqlstr + " 	where s.yyyymm <= '" + yyyymm1 + "' "
            sqlstr = sqlstr + " 	and s.yyyymm >= '" + yyyymm2 + "' "
            sqlstr = sqlstr + " 	and s.itemgubun='10' "
            sqlstr = sqlstr + " 	group by s.itemgubun, s.itemid, s.itemoption "
            sqlstr = sqlstr + " ) s on c.itemgubun=s.itemgubun and c.itemid=s.itemid and c.itemoption=s.itemoption "
        end if
        sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v on c.itemgubun='10' and c.itemid=v.itemid and c.itemoption=v.itemoption "
        sqlstr = sqlstr + " where c.itemgubun='10' "
        sqlstr = sqlstr + " and c.itemid=i.itemid"

        sqlstr = sqlstr +  whereDetail

        sqlstr = sqlstr + " and IsNULL(s.sellno,0) <= 0 "
        sqlstr = sqlstr + " and IsNULL(s.offchulgono,0) >= 0 "
        sqlstr = sqlstr + " order by i.makerid, c.itemid desc, c.itemoption "

		'response.write sqlstr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemname        = db2html(rsget("itemname"))
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).FitemoptionName  = db2html(rsget("itemoptionname"))
				FItemList(i).Fdeliverytype    = rsget("deliverytype")
				FItemList(i).Fmwdiv           = rsget("mwdiv")
				FItemList(i).Fimgsmall        = rsget("smallimage")
				FItemList(i).Fsellyn          = rsget("sellyn")
				FItemList(i).Fisusing         = rsget("isusing")
				FItemList(i).Flimityn         = rsget("limityn")
				FItemList(i).Flimitno         = rsget("limitno")
				FItemList(i).Flimitsold       = rsget("limitsold")
				FItemList(i).Fsellno          = rsget("sellno")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				''FItemList(i).Fetcchulgono     = rsget("etcchulgono")
				FItemList(i).Ferrbaditemno    = rsget("errbaditemno")
				FItemList(i).Frealstock       = rsget("realstock")
				FItemList(i).Ftoterrno        = rsget("toterrno")
				FItemList(i).Foptlimityn = rsget("optlimityn")
				FItemList(i).Foptlimitno = rsget("optlimitno")
				FItemList(i).Foptlimitsold = rsget("optlimitsold")
                FItemList(i).Fdanjongyn = rsget("danjongyn")

                if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	'적정재고초과상품
	public sub GetItemListOverStore()
		dim sqlstr, i

        ''신상품기준일
		dim pre3yyyymmdd
        pre3yyyymmdd = Left(CStr(dateadd("m" ,-3, CDate(FRectEndDate + "-01"))),10)

        sqlstr = " select count(*) as cnt "
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i, [db_summary].[dbo].tbl_current_logisstock_summary c"
        sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_logisstock_summary m "
        sqlstr = sqlstr + " on c.itemgubun = m.itemgubun and c.itemid = m.itemid and c.itemoption = m.itemoption and m.yyyymm = '" + CStr(FRectEndDate) + "' "
        sqlstr = sqlstr + " where c.itemgubun='10'"
        sqlstr = sqlstr + " and  c.itemid = i.itemid "

        if FRectCd1<>"" then
        	sqlstr = sqlstr + " 	and i.cate_large='" + FRectCd1 + "'"
        end if
        if FRectCd2<>"" then
        	sqlstr = sqlstr + " 	and i.cate_mid='" + FRectCd2 + "'"
        end if
        if FRectCd3<>"" then
        	sqlstr = sqlstr + " 	and i.cate_small='" + FRectCd3 + "'"
        end if

        if FRectOnlySellyn="YS" then
            sqlstr = sqlstr + " and i.sellyn <>'N'"
        elseif FRectOnlySellyn<>"" then
            sqlstr = sqlstr + " and i.sellyn = '" + FRectOnlySellyn + "' "
        end if

        if (FRectOnlyIsUsing <>"") then
            sqlstr = sqlstr + " and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if

        if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlstr = sqlstr + " and i.danjongyn<>'N'"
            sqlstr = sqlstr + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if

        '' 신상품만
        if FRectOnlyOldItem<>"" then
            sqlstr = sqlstr + " 	and i.regdate<'" + pre3yyyymmdd + "'"
        end if

		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

        sqlstr = sqlstr + " and ((IsNULL(m.sellno,0) - IsNULL(m.offchulgono,0))*2) < c.realstock "

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

        sqlstr = " select top " + CStr(FPageSize*FCurrPage) + " c.itemgubun, c.itemid, c.itemoption, IsNULL(v.optionname,'') as itemoptionname, i.makerid, i.itemname, i.smallimage, i.deliverytype, i.mwdiv, i.limityn, i.limitno, i.limitsold, i.sellyn, i.isusing, i.danjongyn, IsNULL(m.reipgono,0) as sellno, IsNULL(m.reipgono,0) as reipgono, IsNULL(m.offchulgono,0) as offchulgono, IsNULL(m.etcchulgono,0) as etcchulgono, c.errbaditemno, c.realstock as currrealstock, c.toterrno "
        sqlStr = sqlstr + " ,IsNULL(v.optionname,'') as codeview , IsNULL(v.isusing,'Y') as optionusing, v.optlimityn, v.optlimitno, v.optlimitsold "
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i, "
        sqlstr = sqlstr + " [db_summary].[dbo].tbl_current_logisstock_summary c "
        sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v "
        sqlstr = sqlstr + "     on c.itemgubun='10' and c.itemid=v.itemid and c.itemoption=v.itemoption "
        sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_logisstock_summary m "
        sqlstr = sqlstr + " on c.itemgubun = m.itemgubun and c.itemid = m.itemid and c.itemoption = m.itemoption and m.yyyymm = '" + CStr(FRectEndDate) + "' "
        sqlstr = sqlstr + " where c.itemgubun='10'"
        sqlstr = sqlstr + " and  c.itemid = i.itemid "

        if FRectCd1<>"" then
        	sqlstr = sqlstr + " 	and i.cate_large='" + FRectCd1 + "'"
        end if
        if FRectCd2<>"" then
        	sqlstr = sqlstr + " 	and i.cate_mid='" + FRectCd2 + "'"
        end if
        if FRectCd3<>"" then
        	sqlstr = sqlstr + " 	and i.cate_small='" + FRectCd3 + "'"
        end if

        if FRectOnlySellyn="YS" then
            sqlstr = sqlstr + " and i.sellyn <>'N'"
        elseif FRectOnlySellyn<>"" then
            sqlstr = sqlstr + " and i.sellyn = '" + FRectOnlySellyn + "' "
        end if

        if (FRectOnlyIsUsing <>"") then
            sqlstr = sqlstr + " and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if

        if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlstr = sqlstr + " and i.danjongyn<>'N'"
            sqlstr = sqlstr + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if

        '' 신상품만
        if FRectOnlyOldItem<>"" then
            sqlstr = sqlstr + " 	and i.regdate<'" + pre3yyyymmdd + "'"
        end if

		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if
        sqlstr = sqlstr + " and ((IsNULL(m.sellno,0) - IsNULL(m.offchulgono,0))*2) < c.realstock "
        sqlstr = sqlstr + " order by i.makerid, c.itemid desc, c.itemoption "

		rsget.pagesize = FPageSize

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemname        = db2html(rsget("itemname"))
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).FitemoptionName  = db2html(rsget("itemoptionname"))
				FItemList(i).Fdeliverytype    = rsget("deliverytype")
				FItemList(i).Fmwdiv           = rsget("mwdiv")
				FItemList(i).Fimgsmall        = rsget("smallimage")
				FItemList(i).Flimityn         = rsget("limityn")
				FItemList(i).Flimitno         = rsget("limitno")
				FItemList(i).Flimitsold       = rsget("limitsold")
				FItemList(i).Fsellyn          = rsget("sellyn")
				FItemList(i).Fisusing         = rsget("isusing")
				FItemList(i).Fdanjongyn       = rsget("danjongyn")
				FItemList(i).Freipgono     	  = rsget("reipgono")
				FItemList(i).Fsellno          = -1 * rsget("sellno")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				FItemList(i).Fetcchulgono     = rsget("etcchulgono")
				FItemList(i).Ferrbaditemno    = rsget("errbaditemno")
				FItemList(i).Frealstock   	  = ""
				FItemList(i).Fcurrrealstock   = rsget("currrealstock")
				FItemList(i).Ftoterrno        = rsget("toterrno")
				FItemList(i).Foptlimityn = rsget("optlimityn")
				FItemList(i).Foptlimitno = rsget("optlimitno")
				FItemList(i).Foptlimitsold = rsget("optlimitsold")

                if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

    '상품회전율/정리대상상품 쿼리
    '//admin/stock/turnover_item.asp
	public sub GetItemListTurnOver()
		dim sqlstr, i, PreYYYYMM, pre3yyyymmdd

		if FRectMakerid="" then
		    response.write "브랜드 ID 필수. - 관리자 문의 요망"
		    response.end
		end if

		if (FRectMonthDiff = "") then
			FRectMonthDiff = "1"
		end if

	    PreYYYYMM = Left(CStr(dateadd("m" ,(-1 * FRectMonthDiff), CDate(FRectYYYYMM + "-01"))),7)
        pre3yyyymmdd = Left(CStr(dateadd("m" ,-3, CDate(FRectEndDate + "-01"))),10)		'/신상품기준일

		if (FRectGroupBy = "itemid") then
			sqlstr = " select top " + CStr(FPageSize*FCurrPage) + " s.itemgubun, s.itemid, '0000' as itemoption, sum(s.realstock) as realstock, sum(IsNull(c.realstock, 0)) as currrealstock, "
			sqlstr = sqlstr + " sum(IsNULL(s.sellno,0)-IsNULL(p.sellno,0)) as sellno, sum(IsNULL(s.offchulgono,0)-IsNULL(p.offchulgono,0)) as offchulgono, sum((IsNULL(s.reipgono,0)-IsNULL(p.reipgono,0))) as reipgono, "
			sqlstr = sqlstr + " i.makerid, i.smallimage, i.itemname, '' as itemoptionname, i.mwdiv, "
			sqlstr = sqlstr + " i.sellyn, i.isusing, NULL as limityn, i.limitno, i.limitsold, i.danjongyn,"
			sqlstr = sqlstr + " NULL as optlimityn, NULL as optlimitno, NULL as optlimitsold"
		else
			sqlstr = " select top " + CStr(FPageSize*FCurrPage) + " s.itemgubun, s.itemid, s.itemoption, s.realstock, IsNull(c.realstock, 0) as currrealstock, "
			sqlstr = sqlstr + " IsNULL(s.sellno,0)-IsNULL(p.sellno,0) as sellno, IsNULL(s.offchulgono,0)-IsNULL(p.offchulgono,0) as offchulgono, (IsNULL(s.reipgono,0)-IsNULL(p.reipgono,0)) as reipgono, "
			sqlstr = sqlstr + " i.makerid, i.smallimage, i.itemname, IsNULL(o.optionname,'') as itemoptionname, i.mwdiv, "
			sqlstr = sqlstr + " i.sellyn, i.isusing, i.limityn, i.limitno, i.limitsold, i.danjongyn,"
			sqlstr = sqlstr + " o.optlimityn, o.optlimitno, o.optlimitsold"
			''sqlstr = sqlstr + " ,s.realstock as thisRealStock "
		end if

        ''누적재고.
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
        sqlstr = sqlstr + " Join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
        sqlstr = sqlstr + " 	on s.yyyymm='" & FRectEndDate & "'"
        sqlstr = sqlstr + " 	and s.itemgubun='10'"
        sqlstr = sqlstr + " 	and s.itemid=i.itemid"

        sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option o"
        sqlstr = sqlstr + " 	on s.itemgubun='10'"
        sqlstr = sqlstr + " 	and s.itemid=o.itemid"
        sqlstr = sqlstr + " 	and s.itemoption=o.itemoption"

        ''월 기간출고
'        sqlstr = sqlstr + " left join ("
'        sqlstr = sqlstr + "     select itemgubun, itemid, itemoption, sum(sellno) as sellno, sum(offchulgono) as offchulgono"
'        sqlstr = sqlstr + "     from [db_summary].[dbo].tbl_monthly_logisstock_summary "
'        sqlstr = sqlstr + " 	where yyyymm>='" + FRectYYYYMM + "'"
'        sqlstr = sqlstr + " 	and yyyymm<='" + FRectEndDate + "'"
'        sqlstr = sqlstr + " 	and itemgubun='10'"
'        sqlstr = sqlstr + " 	and (sellno<>0 or offchulgono<>0)"
'        sqlstr = sqlstr + " 	group by itemgubun, itemid, itemoption"
'        sqlstr = sqlstr + " ) m "
'        sqlstr = sqlstr + " 	on s.itemgubun='10'"
'        sqlstr = sqlstr + " 	and m.itemgubun=s.itemgubun"
'        sqlstr = sqlstr + " 	and m.itemid=s.itemid"
'        sqlstr = sqlstr + " 	and m.itemoption=s.itemoption"
        sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary p"
        sqlstr = sqlstr + "     on p.yyyymm='" & PreYYYYMM & "'"
        sqlstr = sqlstr + "     and p.itemgubun='10'"
        sqlstr = sqlstr + " 	and p.itemgubun=s.itemgubun"
        sqlstr = sqlstr + " 	and p.itemid=s.itemid"
        sqlstr = sqlstr + " 	and p.itemoption=s.itemoption"
        sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary c"
		sqlstr = sqlstr + " on c.itemgubun='10'"
        sqlstr = sqlstr + " 	and c.itemgubun=s.itemgubun"
        sqlstr = sqlstr + " 	and c.itemid=s.itemid"
        sqlstr = sqlstr + " 	and c.itemoption=s.itemoption"
        sqlstr = sqlstr + " where i.itemid<>0"

		if FRectDispCate<>"" Then
			sqlStr = sqlStr + " and s.itemgubun = '10' "
		    if LEN(FRectDispCate)>3 then
		         sqlStr = sqlStr + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
		    end if
			sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            '' MD품절 추가
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlstr = sqlstr + " and i.danjongyn<>'N'"
            sqlstr = sqlstr + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if FRectCd1<>"" then
        	sqlstr = sqlstr + " and i.cate_large='" + FRectCd1 + "'"
        end if

        if FRectCd2<>"" then
        	sqlstr = sqlstr + " and i.cate_mid='" + FRectCd2 + "'"
        end if

        if FRectCd3<>"" then
        	sqlstr = sqlstr + " and i.cate_small='" + FRectCd3 + "'"
        end if

        if FRectOnlySellyn="YS" then
            sqlstr = sqlstr + " and i.sellyn <>'N'"
        elseif FRectOnlySellyn<>"" then
        	sqlstr = sqlstr + " 	and i.sellyn = '" + FRectOnlySellyn + "' "
        end if
        if FRectOnlyIsUsing<>"" then
        	sqlstr = sqlstr + " 	and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if

        '' 신상품만
        if FRectOnlyOldItem<>"" then
            sqlstr = sqlstr + " 	and i.regdate<'" + pre3yyyymmdd + "'"
        end if

        '' 정리 대상 상품만
        if FRectOnlyOutItem<>"" then
			if (FRectGroupBy <> "itemid") then
            	sqlstr = sqlstr + " and (IsNULL(s.sellno,0)-IsNULL(p.sellno,0))-(IsNULL(s.offchulgono,0)-IsNULL(p.offchulgono,0)) - (IsNULL(s.reipgono,0)-IsNULL(p.reipgono,0)) <" & FRectChulgoNo

				If (FRectExcBaseRegItem <> "") Then
					sqlstr = sqlstr + " and i.regdate<'" + PreYYYYMM + "-01' "
				End If
			end if
            if (FRectTurnOverPro<>"") then
                sqlstr = sqlstr + " and ((s.realstock=0) or (s.realstock<>0 and ((-1*(IsNULL(s.sellno,0)-IsNULL(p.sellno,0)) + (IsNULL(s.offchulgono,0)-IsNULL(p.offchulgono,0)))*-1.0/(CASE when s.realstock=0 then 1 else s.realstock end))<" & FRectTurnOverPro & "))"
            end if
        end if

		if (FRectMonthGubun <> "") then
			Select Case FRectMonthGubun
				Case "2"
					sqlstr = sqlstr + " 	and DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') <= 2 "
				Case "5"
					sqlstr = sqlstr + " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') > 2) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') <= 5) "
				Case "11"
					sqlstr = sqlstr + " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') > 5) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') <= 11) "
				Case "23"
					sqlstr = sqlstr + " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') <= 23) "
				Case "24"
					sqlstr = sqlstr + " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') > 23) "
				Case Else
					''
			End Select
		end if

		if (FRectGroupBy = "itemid") then
			sqlstr = sqlstr + " group by s.itemgubun, s.itemid, i.makerid, i.smallimage, i.itemname, i.mwdiv, i.sellyn, i.isusing, i.limitno, i.limitsold, i.danjongyn "
			sqlstr = sqlstr + " having sum((IsNULL(s.sellno,0)-IsNULL(p.sellno,0))-(IsNULL(s.offchulgono,0)-IsNULL(p.offchulgono,0)) - (IsNULL(s.reipgono,0)-IsNULL(p.reipgono,0))) < " & FRectChulgoNo
			sqlstr = sqlstr + " order by i.makerid, s.itemid desc"
		else
			sqlstr = sqlstr + " order by i.makerid, s.itemid desc, s.itemoption"
		end if

		'response.write sqlstr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalCount = rsget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCurrentStockItem

				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemname        = db2html(rsget("itemname"))
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).FitemoptionName  = db2html(rsget("itemoptionname"))
				'FItemList(i).Fdeliverytype    = rsget("deliverytype")
				FItemList(i).Fmwdiv           = rsget("mwdiv")
				FItemList(i).Fimgsmall        = rsget("smallimage")
				FItemList(i).Flimityn         = rsget("limityn")
				FItemList(i).Flimitno         = rsget("limitno")
				FItemList(i).Flimitsold       = rsget("limitsold")
				FItemList(i).Fsellyn          = rsget("sellyn")
				FItemList(i).Fisusing         = rsget("isusing")
				FItemList(i).Fdanjongyn		  = rsget("danjongyn")
				FItemList(i).Fsellno          = -1 * rsget("sellno")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				FItemList(i).Freipgono     = rsget("reipgono")
				FItemList(i).Frealstock       = rsget("realstock")
				FItemList(i).Fcurrrealstock       = rsget("currrealstock")
				FItemList(i).Foptlimityn      = rsget("optlimityn")
				FItemList(i).Foptlimitno      = rsget("optlimitno")
				FItemList(i).Foptlimitsold    = rsget("optlimitsold")

                if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall

'                FItemList(i).Fpre1RealStock     = rsget("pre1RealStock")
'                FItemList(i).Fpre1chulgono      = rsget("pre1chulgono")
'                FItemList(i).Fpre2RealStock     = rsget("pre2RealStock")
'                FItemList(i).Fpre2chulgono      = rsget("pre2chulgono")

                '' 누적 수량으로 수정요망
                FItemList(i).Faccumchulgo       = FItemList(i).Fsellno + FItemList(i).Foffchulgono
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	'상품회전율/정리대상상품 포함 브랜드 쿼리
	'//admin/stock/turnover_item.asp
	public sub GetBrandListTurnOver()
	    dim sqlStr, i, PreYYYYMM, pre3yyyymmdd

		if (FRectMonthDiff = "") then
			FRectMonthDiff = "1"
		end if

	    PreYYYYMM = Left(CStr(dateadd("m" ,(-1 * FRectMonthDiff), CDate(FRectYYYYMM + "-01"))),7)
        pre3yyyymmdd = Left(CStr(dateadd("m" ,-3, CDate(FRectEndDate + "-01"))),10)		'/신상품기준일

	    sqlstr = " select top " + CStr(FPageSize*FCurrPage) + " i.makerid, count(i.itemid) as cnt,"
        sqlstr = sqlstr + " sum(s.realstock) as realstock"
		sqlstr = sqlstr + " , sum(IsNull(c.realstock, 0)) as currrealstock"
        sqlstr = sqlstr + " , sum(IsNULL(s.sellno,0)-IsNULL(p.sellno,0)) as sellno"
		sqlstr = sqlstr + " , sum(IsNULL(s.reipgono,0)-IsNULL(p.reipgono,0)) as reipgono"
        sqlstr = sqlstr + " , sum(IsNULL(s.offchulgono,0)-IsNULL(p.offchulgono,0)) as offchulgono "

        ''누적재고.
        sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i"
        sqlstr = sqlstr + " Join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
        sqlstr = sqlstr + "     on s.yyyymm='" & FRectEndDate & "'"
        sqlstr = sqlstr + "     and s.itemgubun='10'"
        sqlstr = sqlstr + "     and s.itemid=i.itemid"
        sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option o"
        sqlstr = sqlstr + " 	on s.itemgubun='10'"
        sqlstr = sqlstr + " 	and s.itemid=o.itemid"
        sqlstr = sqlstr + " 	and s.itemoption=o.itemoption"

        ''월 기간출고 - : 누적재고로
'        sqlstr = sqlstr + " left join ("
'        sqlstr = sqlstr + "     select itemgubun, itemid, itemoption, sum(sellno) as sellno, sum(offchulgono) as offchulgono"
'        sqlstr = sqlstr + "     from [db_summary].[dbo].tbl_monthly_logisstock_summary "
'        sqlstr = sqlstr + " 	where yyyymm>='" + FRectYYYYMM + "'"
'        sqlstr = sqlstr + " 	and yyyymm<='" + FRectEndDate + "'"
'        sqlstr = sqlstr + " 	and itemgubun='10'"
'        sqlstr = sqlstr + " 	and (sellno<>0 or offchulgono<>0)"
'        sqlstr = sqlstr + " 	group by itemgubun, itemid, itemoption"
'        sqlstr = sqlstr + " ) m "
'        sqlstr = sqlstr + " 	on s.itemgubun='10'"
'        sqlstr = sqlstr + " 	and m.itemgubun=s.itemgubun"
'        sqlstr = sqlstr + " 	and m.itemid=s.itemid"
'        sqlstr = sqlstr + " 	and m.itemoption=s.itemoption"

        sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary p"
        sqlstr = sqlstr + "     on p.yyyymm='" & PreYYYYMM & "'"
        sqlstr = sqlstr + "     and p.itemgubun='10'"
        sqlstr = sqlstr + " 	and p.itemgubun=s.itemgubun"
        sqlstr = sqlstr + " 	and p.itemid=s.itemid"
        sqlstr = sqlstr + " 	and p.itemoption=s.itemoption"
        sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary c"
		sqlstr = sqlstr + " 	on c.itemgubun='10'"
        sqlstr = sqlstr + " 	and c.itemgubun=s.itemgubun"
        sqlstr = sqlstr + " 	and c.itemid=s.itemid"
        sqlstr = sqlstr + " 	and c.itemoption=s.itemoption"
        sqlstr = sqlstr + " where i.itemid<>0"

		if FRectDispCate<>"" Then
			sqlStr = sqlStr + " and s.itemgubun = '10' "
		    if LEN(FRectDispCate)>3 then
		         sqlStr = sqlStr + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
		    end if
			sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        if FRectMwDiv="MW" then
            sqlstr = sqlstr + " and i.mwdiv<>'U'"
        elseif FRectMwDiv<>"" then
            sqlstr = sqlstr + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

        if FRectDanjongyn="SN" then
            sqlstr = sqlstr + " and i.danjongyn<>'Y'"
            '' MD품절 추가
            sqlstr = sqlstr + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlstr = sqlstr + " and i.danjongyn<>'N'"
            sqlstr = sqlstr + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlstr = sqlstr + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectLimityn="Y0" then
            sqlstr = sqlstr + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlstr = sqlstr + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "' "
		end if

		if FRectCd1<>"" then
        	sqlstr = sqlstr + " and i.cate_large='" + FRectCd1 + "'"
        end if

        if FRectCd2<>"" then
        	sqlstr = sqlstr + " and i.cate_mid='" + FRectCd2 + "'"
        end if

        if FRectCd3<>"" then
        	sqlstr = sqlstr + " and i.cate_small='" + FRectCd3 + "'"
        end if

        if FRectOnlySellyn="YS" then
            sqlstr = sqlstr + " and i.sellyn <>'N'"
        elseif FRectOnlySellyn<>"" then
        	sqlstr = sqlstr + " 	and i.sellyn = '" + FRectOnlySellyn + "' "
        end if
        if FRectOnlyIsUsing<>"" then
        	sqlstr = sqlstr + " 	and i.isusing = '" + FRectOnlyIsUsing + "' "
        end if

        '' 구상품만
        if FRectOnlyOldItem<>"" then
            sqlstr = sqlstr + " 	and i.regdate<'" + pre3yyyymmdd + "'"
        end if

        '' 정리 대상 상품만
        if FRectOnlyOutItem<>"" then
            sqlstr = sqlstr + " and (IsNULL(s.sellno,0)-IsNULL(p.sellno,0)) - (IsNULL(s.offchulgono,0)-IsNULL(p.offchulgono,0)) - (IsNULL(s.reipgono,0)-IsNULL(p.reipgono,0)) <" & FRectChulgoNo
			sqlstr = sqlstr + " and i.regdate<'" + PreYYYYMM + "-01' "
            if (FRectTurnOverPro<>"") then
                sqlstr = sqlstr + " and ((s.realstock=0) or (s.realstock<>0 and ((-1*(IsNULL(s.sellno,0)-IsNULL(p.sellno,0)) + (IsNULL(s.offchulgono,0)-IsNULL(p.offchulgono,0)))*-1.0/(CASE when s.realstock=0 then 1 else s.realstock end))<" & FRectTurnOverPro & "))"
            end if
        end if

		if (FRectMonthGubun <> "") then
			Select Case FRectMonthGubun
				Case "2"
					sqlstr = sqlstr + " 	and DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') <= 2 "
				Case "5"
					sqlstr = sqlstr + " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') > 2) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') <= 5) "
				Case "11"
					sqlstr = sqlstr + " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') > 5) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') <= 11) "
				Case "23"
					sqlstr = sqlstr + " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') <= 23) "
				Case "24"
					sqlstr = sqlstr + " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectEndDate) + "' + '-01') > 23) "
				Case Else
					''
			End Select
		end if

        sqlstr = sqlstr + " group by i.makerid "
        sqlstr = sqlstr + " order by cnt desc, realstock desc "

		'response.write sqlstr & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTurnOverBrand

					FItemList(i).Fmakerid   	= rsget("makerid")
	                FItemList(i).Fcnt       	= rsget("cnt")
	                FItemList(i).Frealstock 	= rsget("realstock")
					FItemList(i).Fcurrrealstock = rsget("currrealstock")
	                FItemList(i).Fsellno    	= rsget("sellno")
	                FItemList(i).Foffchulgono 	= rsget("offchulgono")
					FItemList(i).Freipgono 		= rsget("reipgono")

	                FTotalcount = FTotalcount + FItemList(i).Fcnt

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public sub GetDaily_Logisstock_Summary()
		dim sqlstr, i

		sqlstr = "select top 300 * from [db_summary].[dbo].tbl_daily_logisstock_summary"
		sqlstr = sqlstr + " where yyyymmdd>='" + FRectStartDate + "'"
		sqlstr = sqlstr + " and itemgubun='" + FRectItemGubun + "'"
		sqlstr = sqlstr + " and itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " and itemoption='" + CStr(FRectItemOption) + "'"
		sqlstr = sqlstr + " order by yyyymmdd"

		rsget.CursorLocation = adUseClient
        rsget.Open sqlstr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSummaryItemStockItem

				FItemList(i).Fyyyymmdd        = rsget("yyyymmdd")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fipgono          = rsget("ipgono")
				FItemList(i).Freipgono        = rsget("reipgono")
				FItemList(i).Ftotipgono       = rsget("totipgono")
				FItemList(i).Foffchulgono     = rsget("offchulgono")
				FItemList(i).Foffrechulgono   = rsget("offrechulgono")
				FItemList(i).Fetcchulgono     = rsget("etcchulgono")
				FItemList(i).Fetcrechulgono   = rsget("etcrechulgono")
				FItemList(i).Ftotchulgono     = rsget("totchulgono")
				FItemList(i).Fsellno          = -1*rsget("sellno")
				FItemList(i).Fresellno        = -1*rsget("resellno")
				FItemList(i).Ftotsellno       = -1*rsget("totsellno")
				FItemList(i).Ferrcsno         = rsget("errcsno")
				FItemList(i).Ferrbaditemno    = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno  = rsget("errrealcheckno")
				FItemList(i).Ferretcno        = rsget("erretcno")
				FItemList(i).Ftoterrno        = rsget("toterrno")
				FItemList(i).Foffsellno        = rsget("offsellno")
				FItemList(i).Ftotsysstock     = rsget("totsysstock")
				FItemList(i).Favailsysstock   = rsget("availsysstock")
				FItemList(i).Frealstock       = rsget("realstock")
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Flastupdate      = rsget("lastupdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

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

function IsOffContractExist(FOffMwMargin)
    IsOffContractExist = Not (FOffMwMargin = "")
end function

function GetOffContractBuycash(FBuycash,Fsellcash,FOffMwMargin)
	dim tmpArr

	GetOffContractBuycash = FBuycash

	'// M_45_0
	if FOffMwMargin<>"" and not(isnull(FOffMwMargin)) then
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
	end if
end function

function GetOffContractCenterMW(FOffMwMargin)
	dim tmpArr

	GetOffContractCenterMW = "U"

	'// M_45_0
	if FOffMwMargin<>"" and not(isnull(FOffMwMargin)) then
		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			GetOffContractCenterMW = tmpArr(0)
		end if
	end if
end function

function GetOffContractMargin(FOffMwMargin)
	dim tmpArr

	GetOffContractMargin = ""

	'// M_45_0
	if FOffMwMargin<>"" and not(isnull(FOffMwMargin)) then
		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			GetOffContractMargin = tmpArr(1)
		end if
	end if
end function

''' 실사재고 == 시스템재고+실사오차 == 실유효재고-불량재고 == Frealstock-Ferrbaditemno ''불량재고 따로관리
function getErrAssignStock(Ftotsysstock,Ferrrealcheckno)
	getErrAssignStock = (Ftotsysstock+Ferrrealcheckno)
end function

Function GetLimitStr(Fitemoption,FLimityn,FLimitNo,FLimitSold,FOptLimityn,FOptLimitNo,Foptlimitsold)
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
			if FOptLimitNo-FOptLimitSold<1 then
				GetLimitStr = "0"
			else
				GetLimitStr = CStr(Foptlimitno-Foptlimitsold)
			end if
		end if
	end if
end function

public function GetMwDivColorCd(mwdiv)
	if mwdiv="M" then
		GetMwDivColorCd = "red"
	elseif mwdiv="W" then
		GetMwDivColorCd = "black"
	elseif mwdiv="U" then
		GetMwDivColorCd = "green"
	else
		GetMwDivColorCd = "gray"
	end if
end function
%>
