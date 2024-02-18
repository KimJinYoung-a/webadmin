<%

Class CDanpumBaljuItem
    public FItemId
    public FItemName
    public FMakerid
    public FImageSmall
    public FDivCD
	public FItemOption
	public FItemOptionName
	Public Fmwdiv
	Public Freguserid

	public function GetDivCDString()
		if FDivCD="O" then
			'단품
			GetDivCDString = "단품상품"
		elseif FDivCD="E" then
			'제외
			GetDivCDString = "제외상품"
		elseif FDivCD="I" then
			'포함
			GetDivCDString = "포함상품"
		end if
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CDanpumBaljuBrand
    public Fbrandid
    public Fsocname
    public Fsocname_kor
    public Fcompany_name
    public FDivCD

	public function GetDivCDString()
		if FDivCD="O" then
			'단품
			GetDivCDString = "단품브랜드"
		elseif FDivCD="E" then
			'제외
			GetDivCDString = "제외브랜드"
		elseif FDivCD="I" then
			'포함
			GetDivCDString = "포함브랜드"
		end if
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CTenBaljuItem
	public Ftplcompanyid
	public Forderserial
	public Fjumundiv
	public Fuserid
	public Faccountname
	public Faccountdiv
	public Faccountno
	public Ftotalmileage
	public Ftotalsum
	public Fipkumdiv
	public Fipkumdate
	public Fregdate
	public Fbeadaldiv
	public Fbeadaldate
	public Fcancelyn
	public Fbuyname
	public Fbuyphone
	public Fbuyhp
	public Fbuyemail
	public Freqname
	public Freqzipcode
	public Freqaddress
	public Freqphone
	public Freqhp
	public Fdeliverno
	public Fsitename
	public Fpaygatetid
	public Fdiscountrate
	public Fsubtotalprice
	public Fresultmsg
	public Frduserid
	public Fmilelogid
	public Fmiletotalprice
	public Fjungsanflag
	public Freqzipaddr
	public Fauthcode


	public Ftenbeaexists
    public FDlvcountryCode
	public FTenbeaItemKindCnt

	public function IpkumDivColor()
		if FjumunDiv="9" then
			IpkumDivColor = "#FF0000"
		else
			if Fipkumdiv="0" then
				IpkumDivColor="#FF0000"
			elseif Fipkumdiv="1" then
				IpkumDivColor="#FF0000"
			elseif Fipkumdiv="2" then
				IpkumDivColor="#000000"
			elseif Fipkumdiv="3" then
				IpkumDivColor="#000000"
			elseif Fipkumdiv="4" then
				IpkumDivColor="#0000FF"
			elseif Fipkumdiv="5" then
				IpkumDivColor="#444400"
			elseif Fipkumdiv="6" then
				IpkumDivColor="#FFFF00"
			elseif Fipkumdiv="7" then
				IpkumDivColor="#004444"
			elseif Fipkumdiv="8" then
				IpkumDivColor="#FF00FF"
			end if
		end if
	end function

	public function SiteNameColor()
		if Fsitename="uto" then
			SiteNameColor = "#55AA22"
		elseif Fsitename="cara" then
			SiteNameColor = "#225555"
		elseif Fsitename="emoden" then
			SiteNameColor = "#992255"
		elseif Fsitename="netian" then
			SiteNameColor = "#AA22AA"
		elseif Fsitename="miclub" then
			SiteNameColor = "#22AA22"
		else
			SiteNameColor = "#000000"
		end if
	end function

	public function SubTotalColor()
		if FSubtotalPrice<0 then
			SubTotalColor = "#DD3333"
		''elseif FSubtotalPrice>50000 then
		''	SubTotalColor = "#33AAAA"
		else
			SubTotalColor = "#000000"
		end if
	end function

	public function CancelYnColor()
		if FCancelYn="D" then
			CancelYnColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelYnColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelYnColor = "#000000"
		end if
	end function

	public function CancelYnName()
		if FCancelYn="D" then
			CancelYnName = "삭제"
		elseif UCase(FCancelYn)="Y" then
			CancelYnName = "취소"
		elseif FCancelYn="N" then
			CancelYnName = "정상"
		end if
	end function

	Public function JumunMethodName()
		if Faccountdiv="7" then
			JumunMethodName="무통장"
		elseif Faccountdiv="14" then
			JumunMethodName="편의점결제"
		elseif Faccountdiv="100" then
			JumunMethodName="신용카드"
		elseif Faccountdiv="20" then
			JumunMethodName="실시간이체"
		elseif Faccountdiv="30" then
			JumunMethodName="포인트"
		elseif Faccountdiv="50" then
			JumunMethodName="외부몰"
		elseif Faccountdiv="80" then
			JumunMethodName="All@카드"
		elseif Faccountdiv="90" then
			JumunMethodName="상품권결제"
		elseif Faccountdiv="110" then
			JumunMethodName="OK+신용"
		elseif Faccountdiv="400" then
			JumunMethodName="휴대폰"
		end if
	end function

	Public function IpkumDivName()
		if FjumunDiv="9" then
			IpkumDivName = "마이너스"
		else
			if Fipkumdiv="0" then
				IpkumDivName="주문대기"
			elseif Fipkumdiv="1" then
				IpkumDivName="주문실패"
			elseif Fipkumdiv="2" then
				IpkumDivName="주문접수"
			elseif Fipkumdiv="3" then
				IpkumDivName="주문접수"
			elseif Fipkumdiv="4" then
				IpkumDivName="결제완료"
			elseif Fipkumdiv="5" then
				IpkumDivName="주문통보"
			elseif Fipkumdiv="6" then
				IpkumDivName="상품준비"
			elseif Fipkumdiv="7" then
				IpkumDivName="일부출고"
			elseif Fipkumdiv="8" then
				IpkumDivName="상품출고"
			end if
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CTenBalju
	public FItemList()

	public FLastQuery

	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FTotalCount

	public FRectRegStart
	public FRectNotitemlist
	public FRectItemlist

	public FRectNotIncludeItem
	public FRectIncludeItem

	public FRectNotIncludebrand
	public FRectIncludebrand

	public FRectTenbeaMakeOnOrder		'// 텐배(주문제작)

    public FRectUpbeaInclude
    public FRectTenbeaOnly
    public FRectDeliveryArea
    public FRectOnlyManyItem
	public FRectOnlyFewItem

    public FRectOnlyOneJumun
    public FRectOnlyOneJumunType
    public FRectOnlyOneJumunCount
    public FRectOnlyOneJumunCompare

	public FRectItemDivCD
	public FRectBrandDivCD

	public FRectSiteGubun

	Public FRectStockLocationGubun
	Public FRectExcMinusStock
	Public FRectPresentOnly
	public FRectRepeatOrderCnt
    public FRectExcZipcode

    public FRectOnlySagawaDeliverArea

	public FRectdcnt

	public FSubTotalsum
	public FAvgTotalsum

    public Sub GetBaljuItemListProc()

		dim sqlStr,i,tmp

		'======================================================================
		''총 갯수. 총금액
		sqlStr = "select count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal " & vbcrlf
		sqlStr = sqlStr & "from [db_threepl].[dbo].[tbl_tpl_orderMaster] m WITH(READUNCOMMITTED) " & vbcrlf
		sqlStr = sqlStr & "where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & "	and m.cancelyn ='N' " & vbcrlf
		sqlStr = sqlStr & "	and m.ipkumdiv > 3 " & vbcrlf
		sqlStr = sqlStr & "	and m.baljudate is NULL " & vbcrlf
		sqlStr = sqlStr & "	and m.jumundiv <> 9 " & vbcrlf					'마이너스 주문 제외
		sqlStr = sqlStr & "	and m.jumundiv <> 4 " & vbcrlf					'티켓 주문 제외
		sqlStr = sqlStr & "	and m.jumundiv <> 6 " & vbcrlf					'교환 주문 제외

		sqlStr = sqlStr & "	and m.regdate>'" & FRectRegStart & "' " & vbcrlf

		rsget_TPL.Open sqlStr,dbget_TPL,1
			FTotalCount = rsget_TPL("cnt")

			FSubtotalsum = rsget_TPL("subtotal")
			FAvgTotalsum = rsget_TPL("avgtotal")

			if IsNull(FSubtotalsum) then FSubtotalsum=0
			if IsNull(FAvgTotalsum) then FAvgTotalsum=0
		rsget_TPL.Close


		'======================================================================
		'데이타
		sqlStr = "exec [db_threepl].[dbo].[usp_TPL_MakeBaljuList] " + CStr(FPageSize) + ", '" + CStr(FRectRegStart) + "' "

		sqlStr = sqlStr + " , '" + CStr(FRectNotIncludeItem) + "', '" + CStr(FRectIncludeItem) + "', '" + CStr(FRectNotIncludebrand) + "', '" + CStr(FRectIncludebrand) + "' "
		sqlStr = sqlStr + " , '" + CStr(FRectTenbeaMakeOnOrder) + "' "
		sqlStr = sqlStr + " , '" + CStr(FRectDeliveryArea) + "' "
		sqlStr = sqlStr + " , '" + CStr(FRectOnlyFewItem) + "', '" + CStr(FRectUpbeaInclude) + "', '" + CStr(FRectTenbeaOnly) + "' "
		sqlStr = sqlStr + " , '" + CStr(FRectOnlyOneJumun) + "', '" + CStr(FRectOnlyOneJumunCompare) + "', '" + CStr(FRectOnlyOneJumunCount) + "', '" + CStr(FRectOnlyOneJumunType) + "' "
		sqlStr = sqlStr + " , '" + CStr(FRectSiteGubun) + "', '" & FRectStockLocationGubun & "', '" & FRectExcMinusStock & "', '" & FRectPresentOnly & "', " & FRectRepeatOrderCnt & ", '" & FRectExcZipcode & "' "

		response.write sqlStr & "<br>"
		''response.end

		FLastQuery = sqlStr


		rsget_TPL.CursorLocation = 3
		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr, dbget_TPL, 3, 1
		'rsget_TPL.Open sqlStr, dbget_TPL, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

        FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		do until rsget_TPL.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Ftplcompanyid 	= rsget_TPL("tplcompanyid")
			FItemList(i).Forderserial 	= rsget_TPL("orderserial")
			FItemList(i).Fjumundiv	  	= rsget_TPL("jumundiv")
			FItemList(i).Fuserid		= rsget_TPL("userid")
			FItemList(i).Faccountdiv	= trim(rsget_TPL("accountdiv"))
			FItemList(i).Ftotalsum		= rsget_TPL("totalsum")
			FItemList(i).Fipkumdiv		= rsget_TPL("ipkumdiv")
			FItemList(i).Fregdate		= rsget_TPL("regdate")
			FItemList(i).Fcancelyn		= rsget_TPL("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsget_TPL("buyname"))
			FItemList(i).Freqname		= db2Html(rsget_TPL("reqname"))
			FItemList(i).Fsitename		= rsget_TPL("sitename")
			FItemList(i).Fsubtotalprice	= rsget_TPL("subtotalprice")

            if (rsget_TPL("TenbeaCnt")=0) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

            FItemList(i).FDlvcountryCode = rsget_TPL("DlvcountryCode")
			FItemList(i).FTenbeaItemKindCnt = rsget_TPL("TenbeaItemKindCnt")

			rsget_TPL.movenext
			i=i+1
		loop
		rsget_TPL.Close

    end Sub

    public Sub GetBaljuItemListNew()

		dim sqlStr,i,tmp

		response.write "시스템팀 문의"
		response.end

		'======================================================================
		''총 갯수. 총금액
		sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal " & vbcrlf
		sqlStr = sqlStr & "from [db_order].[dbo].tbl_order_master m " & vbcrlf
		sqlStr = sqlStr & "where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & "	and m.cancelyn ='N' " & vbcrlf
		sqlStr = sqlStr & "	and m.ipkumdiv > 3 " & vbcrlf
		sqlStr = sqlStr & "	and m.baljudate is NULL " & vbcrlf
		sqlStr = sqlStr & "	and m.jumundiv <> 9 " & vbcrlf					'마이너스 주문 제외
		sqlStr = sqlStr & "	and m.jumundiv <> 4 " & vbcrlf					'티켓 주문 제외
		sqlStr = sqlStr & "	and m.jumundiv <> 6 " & vbcrlf					'교환 주문 제외

		sqlStr = sqlStr & "	and m.regdate>'" & FRectRegStart & "' " & vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")

			FSubtotalsum = rsget("subtotal")
			FAvgTotalsum = rsget("avgtotal")

			if IsNull(FSubtotalsum) then FSubtotalsum=0
			if IsNull(FAvgTotalsum) then FAvgTotalsum=0
		rsget.Close



		'======================================================================
		'데이타
		sqlStr = "select top " & CStr(FPageSize) & " " & vbcrlf
		sqlStr = sqlStr & " 	m.orderserial " & vbcrlf
		sqlStr = sqlStr & "	, m.idx " & vbcrlf
		sqlStr = sqlStr & "	, m.sitename " & vbcrlf
		sqlStr = sqlStr & "	, m.jumundiv " & vbcrlf
		sqlStr = sqlStr & "	, m.DlvcountryCode " & vbcrlf
		sqlStr = sqlStr & "	, m.userid " & vbcrlf
		sqlStr = sqlStr & "	, m.buyname " & vbcrlf
		sqlStr = sqlStr & "	, m.reqname " & vbcrlf
		sqlStr = sqlStr & "	, m.totalsum " & vbcrlf
		sqlStr = sqlStr & "	, m.cancelyn " & vbcrlf
		sqlStr = sqlStr & "	, m.subtotalprice " & vbcrlf
		sqlStr = sqlStr & "	, m.accountdiv " & vbcrlf
		sqlStr = sqlStr & "	, m.ipkumdiv " & vbcrlf
		sqlStr = sqlStr & "	, convert(varchar,m.regdate,20) as regdate " & vbcrlf
		sqlStr = sqlStr & "	, (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid<>0 and d.cancelyn<>'Y' and d.isupchebeasong<>'Y') as TenbeaCnt " & vbcrlf
		sqlStr = sqlStr & "	, (select sum(d.itemno) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid<>0 and d.cancelyn<>'Y' and d.isupchebeasong<>'Y') as TenbeaItemNoCnt " & vbcrlf
		sqlStr = sqlStr & "	, (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid<>0 and d.cancelyn<>'Y' and d.isupchebeasong='Y') as UpcheBeaCnt " & vbcrlf
		sqlStr = sqlStr & "from [db_order].[dbo].tbl_order_master m " & vbcrlf
		sqlStr = sqlStr & "where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & "	and m.cancelyn ='N' " & vbcrlf
		sqlStr = sqlStr & "	and m.ipkumdiv > 3 " & vbcrlf
		sqlStr = sqlStr & "	and m.baljudate is NULL " & vbcrlf
		sqlStr = sqlStr & "	and m.jumundiv <> 9 " & vbcrlf					'마이너스 주문 제외
		sqlStr = sqlStr & "	and m.jumundiv <> 4 " & vbcrlf					'티켓 주문 제외
		sqlStr = sqlStr & "	and m.jumundiv <> 6 " & vbcrlf					'교환 주문 제외

		sqlStr = sqlStr & "	and m.regdate>'" & FRectRegStart & "' " & vbcrlf

		if (FRectNotIncludeItem <> "") then
			sqlStr = sqlStr & "	and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.itemid in (select itemid from [db_item].[dbo].tbl_baljureg_item where IsNull(divcd, 'O') = 'E') and d.cancelyn<>'Y') < 1 " & vbcrlf
		end if

		if (FRectIncludeItem <> "") then
			sqlStr = sqlStr & "	and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.itemid in (select itemid from [db_item].[dbo].tbl_baljureg_item where IsNull(divcd, 'O') = 'I') and d.itemoption = '0012' and d.cancelyn<>'Y') > 0 " & vbcrlf
		end if

		if (FRectNotIncludebrand <> "") then
			sqlStr = sqlStr & "	and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.makerid in (select brandid from [db_item].[dbo].tbl_baljureg_brand where IsNull(divcd, 'O') = 'E') and d.cancelyn<>'Y') < 1 " & vbcrlf
		end if

		if (FRectIncludebrand <> "") then
			sqlStr = sqlStr & "	and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.makerid in (select brandid from [db_item].[dbo].tbl_baljureg_brand where IsNull(divcd, 'O') = 'I') and d.cancelyn<>'Y') > 0 " & vbcrlf
		end if


		'서브쿼리
		tmp = sqlStr
		sqlStr = "select m.* " & vbcrlf
		sqlStr = sqlStr & "from " & vbcrlf
		sqlStr = sqlStr & "( " & vbcrlf
		sqlStr = sqlStr & " " & tmp & " " & vbcrlf
		sqlStr = sqlStr & ") as m " & vbcrlf
		sqlStr = sqlStr & "where 1 = 1 " & vbcrlf



		if (FRectDeliveryArea <> "") then
			if (FRectDeliveryArea = "ZZ") then
				'군부대배송 ==> 군부대 및 제주도.
				sqlStr = sqlStr & " and (IsNULL(m.DlvCountryCode,'KR')='ZZ') " ''' 제주  left(reqzipcode,2)<>'69'
			elseif (FRectDeliveryArea = "EMS") then
				'해외배송
				sqlStr = sqlStr & " and ((IsNULL(m.DlvCountryCode,'KR')<>'KR') and (IsNULL(m.DlvCountryCode,'KR')<>'ZZ')) "
			else
				'국내배송
				sqlStr = sqlStr & " and (IsNULL(m.DlvCountryCode,'KR')='KR') "  ''' 제주  left(reqzipcode,2)<>'69'
			end if
		end if

		if (FRectOnlyManyItem <> "") then
			sqlStr = sqlStr & " and m.ipkumdiv='4' "
			sqlStr = sqlStr & " and TenbeaCnt >= 20 "
			sqlStr = sqlStr & " and m.subtotalPrice >= 500000 "
		end if

		if (FRectUpbeaInclude <> "") then
			sqlStr = sqlStr & " and m.ipkumdiv = 4 "
			sqlStr = sqlStr & " and UpcheBeaCnt > 0 "
		end if

		if (FRectTenbeaOnly <> "") then
			sqlStr = sqlStr & " and m.ipkumdiv = 4 "
			sqlStr = sqlStr & " and UpcheBeaCnt < 1 "
			sqlStr = sqlStr & " and TenbeaCnt > 0 "
		end if

		if (FRectOnlyOneJumun <> "") then
			sqlStr = sqlStr & " and m.ipkumdiv = 4 "
			sqlStr = sqlStr & " and UpcheBeaCnt < 1 "
			sqlStr = sqlStr & " and TenbeaCnt = 1 "

			if (FRectOnlyOneJumunCompare = "equal") then
				sqlStr = sqlStr & " and TenbeaItemNoCnt = " & CStr(FRectOnlyOneJumunCount) & " "
			elseif (FRectOnlyOneJumunCompare = "less") then
				sqlStr = sqlStr & " and TenbeaItemNoCnt <= " & CStr(FRectOnlyOneJumunCount) & " "
			else
				sqlStr = sqlStr & " and TenbeaItemNoCnt >= " & CStr(FRectOnlyOneJumunCount) & " "
			end if

			if (FRectOnlyOneJumunType = "reg") then
				sqlStr = sqlStr & "	and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.itemid in (select itemid from [db_item].[dbo].tbl_baljureg_item where IsNull(divcd, 'O') = 'O') and d.cancelyn<>'Y') > 0 " & vbcrlf
			end if
		end if

		''if (FRectNotIncludeItem <> "") then
		''	sqlStr = sqlStr & "	and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.itemid in (select itemid from [db_item].[dbo].tbl_baljureg_item where IsNull(divcd, 'O') = 'E') and d.cancelyn<>'Y') < 1 " & vbcrlf
		''end if

		''if (FRectIncludeItem <> "") then
		''	sqlStr = sqlStr & "	and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.itemid in (select itemid from [db_item].[dbo].tbl_baljureg_item where IsNull(divcd, 'O') = 'I') and d.cancelyn<>'Y') > 0 " & vbcrlf
		''end if

		''if (FRectNotIncludebrand <> "") then
		''	sqlStr = sqlStr & "	and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.makerid in (select brandid from [db_item].[dbo].tbl_baljureg_brand where IsNull(divcd, 'O') = 'E') and d.cancelyn<>'Y') < 1 " & vbcrlf
		''end if

		''if (FRectIncludebrand <> "") then
		''	sqlStr = sqlStr & "	and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.makerid in (select brandid from [db_item].[dbo].tbl_baljureg_brand where IsNull(divcd, 'O') = 'I') and d.cancelyn<>'Y') > 0 " & vbcrlf
		''end if

		sqlStr = sqlStr & " and m.orderserial <> '10042631803' " '임시.... aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa

		sqlStr = sqlStr & "order by m.idx " & vbcrlf



		''response.write sqlStr

		FLastQuery = sqlStr


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Forderserial 	= rsget("orderserial")
			FItemList(i).Fjumundiv	  	= rsget("jumundiv")
			FItemList(i).Fuserid		= rsget("userid")
			FItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FItemList(i).Ftotalsum		= rsget("totalsum")
			FItemList(i).Fipkumdiv		= rsget("ipkumdiv")
			FItemList(i).Fregdate		= rsget("regdate")
			FItemList(i).Fcancelyn		= rsget("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FItemList(i).Freqname		= db2Html(rsget("reqname"))
			FItemList(i).Fsitename		= rsget("sitename")
			FItemList(i).Fsubtotalprice	= rsget("subtotalprice")

            if (rsget("TenbeaCnt")=0) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

            FItemList(i).FDlvcountryCode = rsget("DlvcountryCode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close

    end Sub

    public Sub GetDanpumBaljuItemList()
        dim sqlStr,i

        sqlStr = "select count(b.itemid) as cnt from [db_threepl].[dbo].tbl_baljureg_item b, [db_threepl].[dbo].tbl_tpl_item i"
        sqlStr = sqlStr + " where b.itemid=i.itemid"

        rsget_TPL.Open sqlStr,dbget_TPL,1
			FTotalCount = rsget_TPL("cnt")
		rsget_TPL.Close

        sqlStr = "select top " + CStr(FPageSize*FCurrpage) + " i.itemid, i.itemname, i.makerid, i.smallimage, IsNull(b.divcd, 'O') as divcd, i.itemoption, i.itemoptionname, b.reguserid "
        sqlStr = sqlStr + " from [db_threepl].[dbo].tbl_baljureg_item b "
        sqlStr = sqlStr + " join [db_threepl].[dbo].tbl_tpl_item i on b.itemid=i.itemid "
        sqlStr = sqlStr + " where 1=1 "

        if (FRectItemDivCD <> "") then
        	sqlStr = sqlStr + " and b.divcd= '" & FRectItemDivCD & "' "
        end if

        sqlStr = sqlStr + " order by b.regdate desc"

        rsget_TPL.pagesize = FPageSize
        rsget_TPL.Open sqlStr,dbget_TPL,1
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
        if  not rsget_TPL.EOF  then
			i = 0
			rsget_TPL.absolutepage = FCurrPage
    		do until rsget_TPL.eof

    			set FItemList(i) = new CDanpumBaljuItem
    			FItemList(i).FItemId    = rsget_TPL("itemid")
                FItemList(i).FItemName  = db2html(rsget_TPL("itemname"))
                FItemList(i).FMakerid   = rsget_TPL("makerid")
                FItemList(i).FImageSmall= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget_TPL("smallimage")

				FItemList(i).FItemOption   		= rsget_TPL("itemoption")
				FItemList(i).FItemOptionName   	= db2html(rsget_TPL("optionname"))

                FItemList(i).FDivCD     = rsget_TPL("divcd")
				FItemList(i).Freguserid	= rsget_TPL("reguserid")

    			rsget_TPL.movenext
    			i=i+1
    		loop
    	end if
		rsget_TPL.Close

    end Sub


    public Sub GetDanpumBaljuBrandList()
        dim sqlStr,i

        sqlStr = " select count(b.brandid) as cnt "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_item].[dbo].tbl_baljureg_brand b "
        sqlStr = sqlStr + " 	left join [db_user].[dbo].tbl_user_c c "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		b.brandid = c.userid "
        sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner p "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		c.userid=p.id "

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

        sqlStr = "select top " + CStr(FPageSize*FCurrpage) + " b.brandid, IsNull(b.divcd, 'O') as divcd, c.socname, c.socname_kor, p.company_name "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_item].[dbo].tbl_baljureg_brand b "
        sqlStr = sqlStr + " 	left join [db_user].[dbo].tbl_user_c c "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		b.brandid = c.userid "
        sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner p "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		c.userid=p.id "

        if (FRectBrandDivCD <> "") then
        	sqlStr = sqlStr + " and b.divcd= '" & FRectBrandDivCD & "' "
        end if

        sqlStr = sqlStr + " order by b.regdate desc"

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
        if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
    		do until rsget.eof

    			set FItemList(i) = new CDanpumBaljuBrand

    			FItemList(i).Fbrandid    		= rsget("brandid")
    			FItemList(i).Fsocname    		= rsget("socname")
    			FItemList(i).Fsocname_kor    	= rsget("socname_kor")
    			FItemList(i).Fcompany_name    	= rsget("company_name")
    			FItemList(i).FDivCD    			= rsget("DivCD")

    			rsget.movenext
    			i=i+1
    		loop
    	end if
		rsget.Close

    end Sub


    ''단품주문건. (텐배송 상품 1개 - 특수 박스에 패킹.)
	public Sub SearchDanpumChulgoJumunList()
		dim sqlStr,i

		''#################################################
		''총 갯수. 총금액
		''#################################################

		response.write "시스템팀 문의"
		response.end

		sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where m.idx<>0"
		sqlStr = sqlStr + " and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " and m.cancelyn ='N'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and m.baljudate is NULL"


		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")

			FSubtotalsum = rsget("subtotal")
			FAvgTotalsum = rsget("avgtotal")

			if IsNull(FSubtotalsum) then FSubtotalsum=0
			if IsNull(FAvgTotalsum) then FAvgTotalsum=0
		rsget.Close


		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize) + "T.*, T2.orderserial as tenbeaexists, T2.tenitemscount "
		sqlStr = sqlStr + " from ("
		sqlStr = sqlStr + " 	select  m.idx, m.orderserial, m.jumundiv, m.DlvcountryCode,"
		sqlStr = sqlStr + " 	m.userid, m.accountname, m.accountdiv, m.totalsum, m.ipkumdiv, "
		sqlStr = sqlStr + " 	m.ipkumdate, m.cancelyn, m.buyname, "
		sqlStr = sqlStr + " 	m.reqname, m.sitename, m.subtotalprice, "
		sqlStr = sqlStr + " 	convert(varchar,m.regdate,20) as cvreg, "
		sqlStr = sqlStr + " 	sum(d.itemno) as dcnt "
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m ,"
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d, "
		sqlStr = sqlStr + " 	[db_item].[dbo].tbl_baljureg_item r "
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial "
		sqlStr = sqlStr + " 	and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " 	and m.cancelyn ='N'"
		sqlStr = sqlStr + "     and m.ipkumdiv>3"
		sqlStr = sqlStr + "     and m.jumundiv<>9"
		sqlStr = sqlStr + "     and m.baljudate is NULL"
		sqlStr = sqlStr + " 	and d.itemid<>0 "
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y' "
		sqlStr = sqlStr + " 	and d.itemid=r.itemid "
		sqlStr = sqlStr + " 	group by m.idx, m.orderserial, m.jumundiv, m.DlvcountryCode,"
		sqlStr = sqlStr + " 	m.userid, m.accountname, m.accountdiv, m.totalsum, m.ipkumdiv, "
		sqlStr = sqlStr + " 	m.ipkumdate, m.cancelyn, m.buyname, "
		sqlStr = sqlStr + " 	m.reqname, m.sitename, m.subtotalprice, "
		sqlStr = sqlStr + " 	convert(varchar,m.regdate,20)"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.orderserial, sum(d.itemno) as tenitemscount "
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " 	and m.cancelyn ='N'"
		sqlStr = sqlStr + "     and m.ipkumdiv>3"
		sqlStr = sqlStr + "     and m.jumundiv<>9"
		sqlStr = sqlStr + "     and m.baljudate is NULL"
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " 	and d.itemid<>0"
		sqlStr = sqlStr + " 	and d.isupchebeasong<>'Y'"
		sqlStr = sqlStr + " 	group by m.orderserial"
		sqlStr = sqlStr + " ) T2 on T.orderserial=T2.orderserial"

		sqlStr = sqlStr + " where T.dcnt=T2.tenitemscount"

		if (FRectdcnt=11) then
		    sqlStr = sqlStr + " and T.dcnt>=11"
		elseif (FRectdcnt=0) then
		    sqlStr = sqlStr + " and T.dcnt>0"
		else
		    sqlStr = sqlStr + " and T.dcnt=" & FRectdcnt
		end if

		sqlStr = sqlStr + " order by T.idx"
'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Forderserial = rsget("orderserial")
			FItemList(i).Fjumundiv	= rsget("jumundiv")
			FItemList(i).Fuserid		= rsget("userid")
			FItemList(i).Faccountname	= db2Html(rsget("accountname"))
			FItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FItemList(i).Ftotalsum	= rsget("totalsum")
			FItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FItemList(i).Fipkumdate	= rsget("ipkumdate")
			FItemList(i).Fregdate		= rsget("cvreg")
			FItemList(i).Fcancelyn	= rsget("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FItemList(i).Freqname		= db2Html(rsget("reqname"))
			FItemList(i).Fsitename	= rsget("sitename")
			FItemList(i).Fsubtotalprice	= rsget("subtotalprice")

			if IsNULL(rsget("tenbeaexists")) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

            FItemList(i).FDlvcountryCode = rsget("DlvcountryCode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close
	end sub


    ''단품주문건. (텐배송 1건 갯수무관)
	public Sub SearchOnlyOnJumunList()
		dim sqlStr,i


		''#################################################
		''총 갯수. 총금액
		''#################################################

		response.write "시스템팀 문의"
		response.end

		sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where m.idx<>0"
		sqlStr = sqlStr + " and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " and m.cancelyn ='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='4'"


		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")

			FSubtotalsum = rsget("subtotal")
			FAvgTotalsum = rsget("avgtotal")

			if IsNull(FSubtotalsum) then FSubtotalsum=0
			if IsNull(FAvgTotalsum) then FAvgTotalsum=0
		rsget.Close


		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize) + "T.*, T2.orderserial as tenbeaexists, IsNULL(T2.tenchulcnt,0) as tenchulcnt "
		sqlStr = sqlStr + " from ("
		sqlStr = sqlStr + " 	select  m.idx, m.orderserial, m.jumundiv, m.DlvcountryCode,"
		sqlStr = sqlStr + " 	m.userid, m.accountname, m.accountdiv, m.totalsum, m.ipkumdiv, "
		sqlStr = sqlStr + " 	m.ipkumdate, m.cancelyn, m.buyname, "
		sqlStr = sqlStr + " 	m.reqname, m.sitename, m.subtotalprice, "
		sqlStr = sqlStr + " 	convert(varchar,m.regdate,20) as cvreg, "
		sqlStr = sqlStr + " 	count(d.idx) as dcnt "
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m ,"
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial "
		sqlStr = sqlStr + " 	and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " 	and m.cancelyn ='N'"
		sqlStr = sqlStr + " 	and m.ipkumdiv='4'"
		sqlStr = sqlStr + " 	and d.itemid<>0 "
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y' "
		sqlStr = sqlStr + " 	group by m.idx, m.orderserial, m.jumundiv, m.DlvcountryCode,"
		sqlStr = sqlStr + " 	m.userid, m.accountname, m.accountdiv, m.totalsum, m.ipkumdiv, "
		sqlStr = sqlStr + " 	m.ipkumdate, m.cancelyn, m.buyname, "
		sqlStr = sqlStr + " 	m.reqname, m.sitename, m.subtotalprice, "
		sqlStr = sqlStr + " 	convert(varchar,m.regdate,20)"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.orderserial, count(d.idx) as tenchulcnt "
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " 	and m.cancelyn ='N'"
		sqlStr = sqlStr + " 	and m.ipkumdiv='4'"
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " 	and d.itemid<>0"
		sqlStr = sqlStr + " 	and d.isupchebeasong<>'Y'"
		sqlStr = sqlStr + " 	group by m.orderserial"
		sqlStr = sqlStr + " ) T2 on T.orderserial=T2.orderserial"

		sqlStr = sqlStr + " where tenchulcnt <='" + FRectdcnt + "'"

		sqlStr = sqlStr + " order by T.idx"
'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Forderserial = rsget("orderserial")
			FItemList(i).Fjumundiv	= rsget("jumundiv")
			FItemList(i).Fuserid		= rsget("userid")
			FItemList(i).Faccountname	= db2Html(rsget("accountname"))
			FItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FItemList(i).Ftotalsum	= rsget("totalsum")
			FItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FItemList(i).Fipkumdate	= rsget("ipkumdate")
			FItemList(i).Fregdate		= rsget("cvreg")
			FItemList(i).Fcancelyn	= rsget("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FItemList(i).Freqname		= db2Html(rsget("reqname"))
			FItemList(i).Fsitename	= rsget("sitename")
			FItemList(i).Fsubtotalprice	= rsget("subtotalprice")

			if IsNULL(rsget("tenbeaexists")) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

            FItemList(i).FDlvcountryCode = rsget("DlvcountryCode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close
	end sub

    public Sub SearchBaljuBigItem()
        dim sqlStr
		dim i

		sqlStr = "select distinct top " + CStr(FPageSize) + " "
		sqlStr = sqlStr + " m.idx, m.orderserial, m.jumundiv, m.userid, m.accountname, m.accountdiv,"
		sqlStr = sqlStr + " m.accountno, m.totalmileage, m.totalsum,"
		sqlStr = sqlStr + " m.ipkumdiv,m.ipkumdate,m.beadaldiv,m.beadaldate,m.cancelyn,m.DlvcountryCode,"
		sqlStr = sqlStr + " m.buyname,m.buyphone,m.buyhp,"
		sqlStr = sqlStr + " m.buyemail,m.reqname,m.reqzipcode,m.reqaddress,m.reqphone,"
		sqlStr = sqlStr + " m.reqhp,m.deliverno,m.sitename,m.paygatetid,"
		sqlStr = sqlStr + " m.discountrate,m.subtotalprice,m.resultmsg,m.rduserid,"
		sqlStr = sqlStr + " m.miletotalprice,m.jungsanflag,m.reqzipaddr,m.authcode,"
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as cvreg, T.orderserial  as tenbeaexists "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.orderserial, "
        sqlStr = sqlStr + " 	sum(case when d.isupchebeasong='N' then 1 else 0 end ) tenBeaCount "
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " 	and m.cancelyn ='N'"
		sqlStr = sqlStr + " 	and m.ipkumdiv='4'"
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " 	and d.itemid<>0"
		''sqlStr = sqlStr + "     and IsNULL(m.DlvcountryCode,'KR')<>'KR'"
		sqlStr = sqlStr + " 	group by m.orderserial"
		sqlStr = sqlStr + " ) T on m.orderserial=T.orderserial"
		sqlStr = sqlStr + " where  m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " and m.cancelyn ='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='4'"
		''sqlStr = sqlStr + " and IsNULL(m.DlvcountryCode,'KR')<>'KR'"
		sqlStr = sqlStr + " and IsNULL(T.tenBeaCount,0)>=20"
		sqlStr = sqlStr + " and m.subtotalPrice>=500000"
		sqlStr = sqlStr + " order by m.idx "

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Forderserial = rsget("orderserial")
			FItemList(i).Fjumundiv	= rsget("jumundiv")
			FItemList(i).Fuserid		= rsget("userid")
			FItemList(i).Faccountname	= db2Html(rsget("accountname"))
			FItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FItemList(i).Faccountno	= rsget("accountno")
			FItemList(i).Ftotalmileage= rsget("totalmileage")

			FItemList(i).Ftotalsum	= rsget("totalsum")
			FItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FItemList(i).Fipkumdate	= rsget("ipkumdate")
			FItemList(i).Fregdate		= rsget("cvreg")
			FItemList(i).Fcancelyn	= rsget("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FItemList(i).Fbuyphone	= rsget("buyphone")
			FItemList(i).Fbuyhp		= rsget("buyhp")
			FItemList(i).Fbuyemail	= rsget("buyemail")
			FItemList(i).Freqname		= db2Html(rsget("reqname"))
			FItemList(i).Freqzipcode	= rsget("reqzipcode")
			FItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
			FItemList(i).Freqzipaddr	= db2Html(rsget("reqzipaddr"))
			FItemList(i).Freqphone	= rsget("reqphone")
			FItemList(i).Freqhp		= rsget("reqhp")
			FItemList(i).Fdeliverno	= rsget("deliverno")
			FItemList(i).Fsitename	= rsget("sitename")
			FItemList(i).Fpaygatetid	= rsget("paygatetid")
			FItemList(i).Fdiscountrate	= rsget("discountrate")
			FItemList(i).Fsubtotalprice	= rsget("subtotalprice")
			FItemList(i).Fresultmsg		= rsget("resultmsg")
			FItemList(i).Fmiletotalprice	= rsget("miletotalprice")

			FItemList(i).Fauthcode		= rsget("authcode")

			if IsNULL(rsget("tenbeaexists")) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

            FItemList(i).FDlvcountryCode = rsget("DlvcountryCode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close
    end Sub

    public Sub SearchBaljuEMS()
        dim sqlStr
		dim i

		response.write "시스템팀 문의"
		response.end

		sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal  "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " and m.cancelyn ='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='4'"
		sqlStr = sqlStr + " and ((IsNULL(m.DlvCountryCode,'KR')<>'KR') and (IsNULL(m.DlvCountryCode,'KR')<>'ZZ')) "

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")

			FSubtotalsum = rsget("subtotal")
			FAvgTotalsum = rsget("avgtotal")

			if IsNull(FSubtotalsum) then FSubtotalsum=0
			if IsNull(FAvgTotalsum) then FAvgTotalsum=0
		rsget.Close

		''특정상품 포함 복합 주문건
		sqlStr = "select distinct top " + CStr(FPageSize) + " "
		sqlStr = sqlStr + " m.idx, m.orderserial, m.jumundiv, m.userid, m.accountname, m.accountdiv,"
		sqlStr = sqlStr + " m.accountno, m.totalmileage, m.totalsum,"
		sqlStr = sqlStr + " m.ipkumdiv,m.ipkumdate,m.beadaldiv,m.beadaldate,m.cancelyn,m.DlvCountryCode,"
		sqlStr = sqlStr + " m.buyname,m.buyphone,m.buyhp,"
		sqlStr = sqlStr + " m.buyemail,m.reqname,m.reqzipcode,m.reqaddress,m.reqphone,"
		sqlStr = sqlStr + " m.reqhp,m.deliverno,m.sitename,m.paygatetid,"
		sqlStr = sqlStr + " m.discountrate,m.subtotalprice,m.resultmsg,m.rduserid,"
		sqlStr = sqlStr + " m.miletotalprice,m.jungsanflag,m.reqzipaddr,m.authcode,"
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as cvreg, T.orderserial  as tenbeaexists "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.orderserial, "
        sqlStr = sqlStr + " 	sum(case when d.isupchebeasong='N' then 1 else 0 end ) tenBeaCount "
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " 	and m.cancelyn ='N'"
		sqlStr = sqlStr + " 	and m.ipkumdiv='4'"
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " 	and d.itemid<>0"
		sqlStr = sqlStr + "     and ((IsNULL(m.DlvCountryCode,'KR')<>'KR') and (IsNULL(m.DlvCountryCode,'KR')<>'ZZ')) "
		sqlStr = sqlStr + " 	group by m.orderserial"
		sqlStr = sqlStr + " ) T on m.orderserial=T.orderserial"
		sqlStr = sqlStr + " where  m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " and m.cancelyn ='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='4'"
		sqlStr = sqlStr + " and ((IsNULL(m.DlvCountryCode,'KR')<>'KR') and (IsNULL(m.DlvCountryCode,'KR')<>'ZZ')) "
		sqlStr = sqlStr + " order by m.idx "

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Forderserial = rsget("orderserial")
			FItemList(i).Fjumundiv	= rsget("jumundiv")
			FItemList(i).Fuserid		= rsget("userid")
			FItemList(i).Faccountname	= db2Html(rsget("accountname"))
			FItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FItemList(i).Faccountno	= rsget("accountno")
			FItemList(i).Ftotalmileage= rsget("totalmileage")

			FItemList(i).Ftotalsum	= rsget("totalsum")
			FItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FItemList(i).Fipkumdate	= rsget("ipkumdate")
			FItemList(i).Fregdate		= rsget("cvreg")
			FItemList(i).Fcancelyn	= rsget("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FItemList(i).Fbuyphone	= rsget("buyphone")
			FItemList(i).Fbuyhp		= rsget("buyhp")
			FItemList(i).Fbuyemail	= rsget("buyemail")
			FItemList(i).Freqname		= db2Html(rsget("reqname"))
			FItemList(i).Freqzipcode	= rsget("reqzipcode")
			FItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
			FItemList(i).Freqzipaddr	= db2Html(rsget("reqzipaddr"))
			FItemList(i).Freqphone	= rsget("reqphone")
			FItemList(i).Freqhp		= rsget("reqhp")
			FItemList(i).Fdeliverno	= rsget("deliverno")
			FItemList(i).Fsitename	= rsget("sitename")
			FItemList(i).Fpaygatetid	= rsget("paygatetid")
			FItemList(i).Fdiscountrate	= rsget("discountrate")
			FItemList(i).Fsubtotalprice	= rsget("subtotalprice")
			FItemList(i).Fresultmsg		= rsget("resultmsg")
			FItemList(i).Fmiletotalprice	= rsget("miletotalprice")

			FItemList(i).Fauthcode		= rsget("authcode")

			if IsNULL(rsget("tenbeaexists")) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

            FItemList(i).FDlvCountryCode = rsget("DlvCountryCode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close

    end Sub

	'군부대배송
    public Sub SearchBaljuMilitary()
        dim sqlStr
		dim i

		response.write "시스템팀 문의"
		response.end

		sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal  "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " and m.cancelyn ='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='4'"
		sqlStr = sqlStr + " and (IsNULL(m.DlvCountryCode,'KR') = 'ZZ') "

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")

			FSubtotalsum = rsget("subtotal")
			FAvgTotalsum = rsget("avgtotal")

			if IsNull(FSubtotalsum) then FSubtotalsum=0
			if IsNull(FAvgTotalsum) then FAvgTotalsum=0
		rsget.Close

		''특정상품 포함 복합 주문건
		sqlStr = "select distinct top " + CStr(FPageSize) + " "
		sqlStr = sqlStr + " m.idx, m.orderserial, m.jumundiv, m.userid, m.accountname, m.accountdiv,"
		sqlStr = sqlStr + " m.accountno, m.totalmileage, m.totalsum,"
		sqlStr = sqlStr + " m.ipkumdiv,m.ipkumdate,m.beadaldiv,m.beadaldate,m.cancelyn,m.DlvcountryCode,"
		sqlStr = sqlStr + " m.buyname,m.buyphone,m.buyhp,"
		sqlStr = sqlStr + " m.buyemail,m.reqname,m.reqzipcode,m.reqaddress,m.reqphone,"
		sqlStr = sqlStr + " m.reqhp,m.deliverno,m.sitename,m.paygatetid,"
		sqlStr = sqlStr + " m.discountrate,m.subtotalprice,m.resultmsg,m.rduserid,"
		sqlStr = sqlStr + " m.miletotalprice,m.jungsanflag,m.reqzipaddr,m.authcode,"
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as cvreg, T.orderserial  as tenbeaexists "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.orderserial, "
        sqlStr = sqlStr + " 	sum(case when d.isupchebeasong='N' then 1 else 0 end ) tenBeaCount "
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " 	and m.cancelyn ='N'"
		sqlStr = sqlStr + " 	and m.ipkumdiv='4'"
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " 	and d.itemid<>0"
		sqlStr = sqlStr + "     and (IsNULL(m.DlvCountryCode,'KR') = 'ZZ') "
		sqlStr = sqlStr + " 	group by m.orderserial"
		sqlStr = sqlStr + " ) T on m.orderserial=T.orderserial"
		sqlStr = sqlStr + " where  m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " and m.cancelyn ='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='4'"
		sqlStr = sqlStr + " and (IsNULL(m.DlvCountryCode,'KR') = 'ZZ') "
		sqlStr = sqlStr + " order by m.idx "

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Forderserial = rsget("orderserial")
			FItemList(i).Fjumundiv	= rsget("jumundiv")
			FItemList(i).Fuserid		= rsget("userid")
			FItemList(i).Faccountname	= db2Html(rsget("accountname"))
			FItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FItemList(i).Faccountno	= rsget("accountno")
			FItemList(i).Ftotalmileage= rsget("totalmileage")

			FItemList(i).Ftotalsum	= rsget("totalsum")
			FItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FItemList(i).Fipkumdate	= rsget("ipkumdate")
			FItemList(i).Fregdate		= rsget("cvreg")
			FItemList(i).Fcancelyn	= rsget("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FItemList(i).Fbuyphone	= rsget("buyphone")
			FItemList(i).Fbuyhp		= rsget("buyhp")
			FItemList(i).Fbuyemail	= rsget("buyemail")
			FItemList(i).Freqname		= db2Html(rsget("reqname"))
			FItemList(i).Freqzipcode	= rsget("reqzipcode")
			FItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
			FItemList(i).Freqzipaddr	= db2Html(rsget("reqzipaddr"))
			FItemList(i).Freqphone	= rsget("reqphone")
			FItemList(i).Freqhp		= rsget("reqhp")
			FItemList(i).Fdeliverno	= rsget("deliverno")
			FItemList(i).Fsitename	= rsget("sitename")
			FItemList(i).Fpaygatetid	= rsget("paygatetid")
			FItemList(i).Fdiscountrate	= rsget("discountrate")
			FItemList(i).Fsubtotalprice	= rsget("subtotalprice")
			FItemList(i).Fresultmsg		= rsget("resultmsg")
			FItemList(i).Fmiletotalprice	= rsget("miletotalprice")

			FItemList(i).Fauthcode		= rsget("authcode")

			if IsNULL(rsget("tenbeaexists")) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

            FItemList(i).FDlvcountryCode = rsget("DlvcountryCode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close

    end Sub

    public Sub SearchBaljuImsi()
        dim sqlStr
		dim i


		''특정상품 포함 복합 주문건
		sqlStr = "select distinct top " + CStr(FPageSize) + " "
		sqlStr = sqlStr + " m.idx, m.orderserial, m.jumundiv, m.userid, m.accountname, m.accountdiv,"
		sqlStr = sqlStr + " m.accountno, m.totalmileage, m.totalsum,"
		sqlStr = sqlStr + " m.ipkumdiv,m.ipkumdate,m.beadaldiv,m.beadaldate,m.cancelyn,m.DlvcountryCode,"
		sqlStr = sqlStr + " m.buyname,m.buyphone,m.buyhp,"
		sqlStr = sqlStr + " m.buyemail,m.reqname,m.reqzipcode,m.reqaddress,m.reqphone,"
		sqlStr = sqlStr + " m.reqhp,m.deliverno,m.sitename,m.paygatetid,"
		sqlStr = sqlStr + " m.discountrate,m.subtotalprice,m.resultmsg,m.rduserid,"
		sqlStr = sqlStr + " m.miletotalprice,m.jungsanflag,m.reqzipaddr,m.authcode,"
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as cvreg, T.orderserial  as tenbeaexists "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.orderserial, "
		sqlStr = sqlStr + " 	sum(case when d.itemid=131267 then 1 else 0 end ) mooExists, "
        sqlStr = sqlStr + " 	sum(case when d.itemid<>131267 then 1 else 0 end ) NotmooExists, "
        sqlStr = sqlStr + " 	sum(case when d.isupchebeasong='N' then 1 else 0 end ) tenBeaCount "
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " 	and m.cancelyn ='N'"
		sqlStr = sqlStr + " 	and m.ipkumdiv='4'"
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " 	and d.itemid<>0"
		sqlStr = sqlStr + " 	group by m.orderserial"
		sqlStr = sqlStr + " ) T on m.orderserial=T.orderserial"
		sqlStr = sqlStr + " where  m.regdate>'" + FRectRegStart + "'"
		sqlStr = sqlStr + " and m.cancelyn ='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='4'"
		sqlStr = sqlStr + " and T.mooExists>0"
		sqlStr = sqlStr + " and T.NotmooExists>0"
		sqlStr = sqlStr + " order by m.idx "

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Forderserial = rsget("orderserial")
			FItemList(i).Fjumundiv	= rsget("jumundiv")
			FItemList(i).Fuserid		= rsget("userid")
			FItemList(i).Faccountname	= db2Html(rsget("accountname"))
			FItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FItemList(i).Faccountno	= rsget("accountno")
			FItemList(i).Ftotalmileage= rsget("totalmileage")

			FItemList(i).Ftotalsum	= rsget("totalsum")
			FItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FItemList(i).Fipkumdate	= rsget("ipkumdate")
			FItemList(i).Fregdate		= rsget("cvreg")
			FItemList(i).Fcancelyn	= rsget("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FItemList(i).Fbuyphone	= rsget("buyphone")
			FItemList(i).Fbuyhp		= rsget("buyhp")
			FItemList(i).Fbuyemail	= rsget("buyemail")
			FItemList(i).Freqname		= db2Html(rsget("reqname"))
			FItemList(i).Freqzipcode	= rsget("reqzipcode")
			FItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
			FItemList(i).Freqzipaddr	= db2Html(rsget("reqzipaddr"))
			FItemList(i).Freqphone	= rsget("reqphone")
			FItemList(i).Freqhp		= rsget("reqhp")
			FItemList(i).Fdeliverno	= rsget("deliverno")
			FItemList(i).Fsitename	= rsget("sitename")
			FItemList(i).Fpaygatetid	= rsget("paygatetid")
			FItemList(i).Fdiscountrate	= rsget("discountrate")
			FItemList(i).Fsubtotalprice	= rsget("subtotalprice")
			FItemList(i).Fresultmsg		= rsget("resultmsg")
			FItemList(i).Fmiletotalprice	= rsget("miletotalprice")

			FItemList(i).Fauthcode		= rsget("authcode")

			if IsNULL(rsget("tenbeaexists")) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

            FItemList(i).FDlvcountryCode = rsget("DlvcountryCode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close

    end sub

	public Sub SearchBaljuJumunList()
		dim sqlStr
		dim i

        ''FRectOnlySagawaDeliverArea

		response.write "시스템팀 문의"
		response.end

		''#################################################
		''총 갯수. 총금액
		''#################################################
		if FRectnotitemlist<>"" then
			sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal  "
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
			sqlStr = sqlStr + " where m.regdate>'" + FRectRegStart + "'"
			sqlStr = sqlStr + " and m.cancelyn ='N'"
			sqlStr = sqlStr + " and m.ipkumdiv>3"
			sqlStr = sqlStr + " and m.jumundiv<>9"
			sqlStr = sqlStr + " and m.baljudate is NULL"
			sqlStr = sqlStr + " and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.itemid in (" + FRectnotitemlist + ") and d.cancelyn<>'Y')<1"

'			sqlStr = sqlStr + " and orderserial not in ("
'			sqlStr = sqlStr + " 	select distinct m.orderserial from "
'			sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m"
'			sqlStr = sqlStr + " 	Join [db_order].[dbo].tbl_order_detail d "
'			sqlStr = sqlStr + " 	on m.orderserial=d.orderserial"
'			sqlStr = sqlStr + " 	where m.regdate>'" + FRectRegStart + "'"
'			sqlStr = sqlStr + " 	and m.cancelyn ='N'"
'			sqlStr = sqlStr + " 	and m.ipkumdiv>3"
'			sqlStr = sqlStr + "     and m.baljudate is NULL"
'			sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
'			sqlStr = sqlStr + " 	and d.itemid<>0"
'			sqlStr = sqlStr + " 	and d.itemid in (" + FRectnotitemlist + ")"
'			sqlStr = sqlStr + " )"

		elseif FRectitemlist<>"" then
			sqlStr = "select count(distinct m.orderserial) as cnt, sum(distinct m.subtotalprice) as subtotal , avg(distinct m.subtotalprice) as avgtotal "
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
			sqlStr = sqlStr + "  Join [db_order].[dbo].tbl_order_detail d "
			sqlStr = sqlStr + " on m.orderserial=d.orderserial"
			sqlStr = sqlStr + " where m.regdate>'" + FRectRegStart + "'"
			sqlStr = sqlStr + " and m.cancelyn ='N'"
			sqlStr = sqlStr + " and m.ipkumdiv>3"
			sqlStr = sqlStr + " and m.jumundiv<>9"
			sqlStr = sqlStr + " and m.baljudate is NULL"
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.itemid in (" + FRectitemlist + ")"
			'sqlStr = sqlStr + " and d.itemoption='0011'"
			''sqlStr = sqlStr + " and (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid=311341 and (d.itemoption='0013' or d.itemoption='0014') and d.cancelyn<>'Y' and d.isupchebeasong<>'Y')<1"
			''sqlStr = sqlStr + " and (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid=311341 and (d.itemoption='0012' or d.itemoption='0014' or d.itemoption='0016') and d.cancelyn<>'Y' and d.isupchebeasong<>'Y')<1"
			'sqlStr = sqlStr + " and (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid=311341 and (d.itemoption='0012' or d.itemoption='0013' or  d.itemoption='0014' or  d.itemoption='0015' or d.itemoption='0016') and d.cancelyn<>'Y' and d.isupchebeasong<>'Y')<1"

		else
			sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal "
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
			if FRectOnlySagawaDeliverArea<>"" then
			    sqlStr = sqlStr + " Join db_temp.dbo.tbl_sagawa_deliver_area S"
			    sqlStr = sqlStr + " on m.reqzipcode=S.ZIP_NO"
			end if
			sqlStr = sqlStr + " where m.regdate>'" + FRectRegStart + "'"
			sqlStr = sqlStr + " and m.ipkumdiv>3"
            ''sqlStr = sqlStr + " and ipkumdiv<8"   '출고처리 된것도 발주..
            sqlStr = sqlStr + " and m.jumundiv<>9"
            sqlStr = sqlStr + " and cancelyn='N'"
            sqlStr = sqlStr + " and m.baljudate is NULL"

		end if



		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")

			FSubtotalsum = rsget("subtotal")
			FAvgTotalsum = rsget("avgtotal")

			if IsNull(FSubtotalsum) then FSubtotalsum=0
			if IsNull(FAvgTotalsum) then FAvgTotalsum=0
		rsget.Close

		''#################################################
		''데이타.
		''#################################################
		if FRectUpbeaInclude<>"" then
            sqlStr = "select top " + CStr(FPageSize)
            sqlStr = sqlStr + " m.orderserial, m.sitename, m.jumundiv"
			sqlStr = sqlStr + " , m.DlvcountryCode, m.userid, m.buyname, m.reqname, m.totalsum"
            sqlStr = sqlStr + " , m.cancelyn, m.subtotalprice, m.accountdiv, m.ipkumdiv, convert(varchar,m.regdate,20) as regdate"
            '', T.orderserial as tenbeaexists "
			sqlStr = sqlStr + " , (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid<>0 and d.cancelyn<>'Y' and d.isupchebeasong<>'Y') as TenbeaCnt"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
'			sqlStr = sqlStr + " left join ("
'			sqlStr = sqlStr + " 	select distinct m.orderserial"
'			sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
'			sqlStr = sqlStr + " 	where m.orderserial=d.orderserial"
'			sqlStr = sqlStr + " 	and m.regdate>'" + FRectRegStart + "'"
'			sqlStr = sqlStr + " 	and m.cancelyn ='N'"
'			sqlStr = sqlStr + " 	and m.ipkumdiv=4"
'            sqlStr = sqlStr + "     and m.jumundiv<>9"
'			sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
'			sqlStr = sqlStr + " 	and d.itemid<>0"
'			sqlStr = sqlStr + " 	and d.isupchebeasong='Y'"
'			sqlStr = sqlStr + "     and m.baljudate is NULL"
'			sqlStr = sqlStr + " ) T on m.orderserial=T.orderserial"
			sqlStr = sqlStr + " where m.regdate>'" + FRectRegStart + "'"
			sqlStr = sqlStr + " and m.ipkumdiv=4"
            sqlStr = sqlStr + " and m.ipkumdiv<8"
            sqlStr = sqlStr + " and m.jumundiv<>9"
            sqlStr = sqlStr + " and m.cancelyn='N'"
            sqlStr = sqlStr + " and m.baljudate is NULL"
'            sqlStr = sqlStr + " and T.orderserial is Not NULL"
			sqlStr = sqlStr + " order by idx "

		elseif FRectnotitemlist<>"" then
		    ''특정상품 제외 주문건
			sqlStr = "select top " + CStr(FPageSize)
			sqlStr = sqlStr + " m.orderserial, m.sitename, m.jumundiv"
			sqlStr = sqlStr + " , m.DlvcountryCode, m.userid, m.buyname, m.reqname, m.totalsum"
            sqlStr = sqlStr + " , m.cancelyn, m.subtotalprice, m.accountdiv, m.ipkumdiv, convert(varchar,m.regdate,20) as regdate"
            sqlStr = sqlStr + " , (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid<>0 and d.cancelyn<>'Y' and d.isupchebeasong<>'Y') as TenbeaCnt"
            ''sqlStr = sqlStr + ", T.orderserial as tenbeaexists "
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
			sqlStr = sqlStr + " where m.regdate>'" + FRectRegStart + "'"
			sqlStr = sqlStr + " and m.cancelyn ='N'"
			sqlStr = sqlStr + " and m.ipkumdiv>3"
            sqlStr = sqlStr + " and m.jumundiv<>9"
            sqlStr = sqlStr + " and m.baljudate is NULL"
            sqlStr = sqlStr + " and (select count(*) from db_order.dbo.tbl_order_detail d where d.orderserial=m.orderserial and d.itemid in (" + FRectnotitemlist + ") and d.cancelyn<>'Y')<1"

'			sqlStr = sqlStr + " and m.orderserial not in ("
'			sqlStr = sqlStr + " 	select distinct m.orderserial from [db_order].[dbo].tbl_order_master m "
'			sqlStr = sqlStr + " 	    Join [db_order].[dbo].tbl_order_detail d "
'			sqlStr = sqlStr + " 	    on m.orderserial=d.orderserial"
'			sqlStr = sqlStr + " 	where m.regdate>'" + FRectRegStart + "'"
'			sqlStr = sqlStr + " 	and m.cancelyn ='N'"
'			sqlStr = sqlStr + "     and m.ipkumdiv>3"
'			sqlStr = sqlStr + "     and m.baljudate is NULL"
'			sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
'			sqlStr = sqlStr + " 	and d.itemid<>0"
'			sqlStr = sqlStr + " 	and d.itemid in (" + FRectnotitemlist + ")"
'			sqlStr = sqlStr + " )"
			sqlStr = sqlStr + " order by m.idx "
		elseif FRectitemlist<>"" then
		    ''특정상품 포함 주문건
			sqlStr = "select distinct top " + CStr(FPageSize) + " "
			sqlStr = sqlStr + " m.idx, m.ipkumdate, m.orderserial, m.sitename, m.jumundiv"
			sqlStr = sqlStr + " , m.DlvcountryCode, m.userid, m.buyname, m.reqname, m.totalsum"
            sqlStr = sqlStr + " , m.cancelyn, m.subtotalprice, m.accountdiv, m.ipkumdiv, convert(varchar,m.regdate,20) as regdate"
            sqlStr = sqlStr + " , (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid<>0 and d.cancelyn<>'Y' and d.isupchebeasong<>'Y') as TenbeaCnt"
			'', T.orderserial as tenbeaexists "
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
			sqlStr = sqlStr + " Join [db_order].[dbo].tbl_order_detail d"
			sqlStr = sqlStr + " on m.orderserial=d.orderserial"
			sqlStr = sqlStr + " where m.regdate>'" + FRectRegStart + "'"
			sqlStr = sqlStr + " and m.cancelyn ='N'"
			sqlStr = sqlStr + " and m.ipkumdiv>3"
			sqlStr = sqlStr + " and m.jumundiv<>9"
			sqlStr = sqlStr + " and m.baljudate is NULL"
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.itemid in (" + FRectitemlist + ")"
			'sqlStr = sqlStr + " and d.itemoption='0011'"
			''sqlStr = sqlStr + " and (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid=311341 and (d.itemoption='0013' or d.itemoption='0014') and d.cancelyn<>'Y' and d.isupchebeasong<>'Y')<1"
			''sqlStr = sqlStr + " and (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid=311341 and (d.itemoption='0012' or d.itemoption='0014' or d.itemoption='0016') and d.cancelyn<>'Y' and d.isupchebeasong<>'Y')<1"
			'sqlStr = sqlStr + " and (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid=311341 and (d.itemoption='0012' or d.itemoption='0013' or  d.itemoption='0014' or  d.itemoption='0015' or d.itemoption='0016') and d.cancelyn<>'Y' and d.isupchebeasong<>'Y')<1"
            sqlStr = sqlStr + " order by m.idx "

		else
		    ''일반발주
			sqlStr = "select top " + CStr(FPageSize)
			sqlStr = sqlStr + " m.orderserial, m.sitename, m.jumundiv"
			sqlStr = sqlStr + " , m.DlvcountryCode, m.userid, m.buyname, m.reqname, m.totalsum"
            sqlStr = sqlStr + " , m.cancelyn, m.subtotalprice, m.accountdiv, m.ipkumdiv, convert(varchar,m.regdate,20) as regdate"
            '', T.orderserial as tenbeaexists "
            sqlStr = sqlStr + " , (select count(*) from [db_order].[dbo].tbl_order_detail d where d.orderserial=m.orderserial and d.itemid<>0 and d.cancelyn<>'Y' and d.isupchebeasong<>'Y') as TenbeaCnt"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
			if FRectOnlySagawaDeliverArea<>"" then
			    sqlStr = sqlStr + " Join db_temp.dbo.tbl_sagawa_deliver_area S"
			    sqlStr = sqlStr + " on m.reqzipcode=S.ZIP_NO"
			end if
			sqlStr = sqlStr + " where m.regdate>'" + FRectRegStart + "'"
			sqlStr = sqlStr + " and m.ipkumdiv>3" ''입금이후 상태에서도 발주 안하는 것 이 있을 수 있음.
            '''sqlStr = sqlStr + " and m.ipkumdiv<8"  ''출고처리된 것도 발주
            sqlStr = sqlStr + " and m.jumundiv<>9"
            sqlStr = sqlStr + " and m.cancelyn='N'"
            sqlStr = sqlStr + " and m.baljudate is NULL"
			sqlStr = sqlStr + " order by idx "
		end if



		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Forderserial = rsget("orderserial")
			FItemList(i).Fjumundiv	  = rsget("jumundiv")
			FItemList(i).Fuserid		= rsget("userid")
			FItemList(i).Faccountdiv	= trim(rsget("accountdiv"))
			FItemList(i).Ftotalsum	= rsget("totalsum")
			FItemList(i).Fipkumdiv	= rsget("ipkumdiv")
			FItemList(i).Fregdate		= rsget("regdate")
			FItemList(i).Fcancelyn	= rsget("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
			FItemList(i).Freqname		= db2Html(rsget("reqname"))
			FItemList(i).Fsitename	= rsget("sitename")
			FItemList(i).Fsubtotalprice	= rsget("subtotalprice")

			'FItemList(i).Faccountname	= db2Html(rsget("accountname"))
			'FItemList(i).Faccountno	= rsget("accountno")
			'FItemList(i).Ftotalmileage= rsget("totalmileage")
			'FItemList(i).Fipkumdate	= rsget("ipkumdate")
			'FItemList(i).Fbuyphone	= rsget("buyphone")
			'FItemList(i).Fbuyhp		= rsget("buyhp")
			'FItemList(i).Fbuyemail	= rsget("buyemail")
			'FItemList(i).Freqzipcode	= rsget("reqzipcode")
			'FItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
			'FItemList(i).Freqzipaddr	= db2Html(rsget("reqzipaddr"))
			'FItemList(i).Freqphone	= rsget("reqphone")
			'FItemList(i).Freqhp		= rsget("reqhp")
			'FItemList(i).Fdeliverno	= rsget("deliverno")
			'FItemList(i).Fpaygatetid	= rsget("paygatetid")
			'FItemList(i).Fdiscountrate	= rsget("discountrate")
			'FItemList(i).Fresultmsg		= rsget("resultmsg")
			'FItemList(i).Fmiletotalprice	= rsget("miletotalprice")
			'FItemList(i).Fauthcode		= rsget("authcode")


            if (rsget("TenbeaCnt")=0) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

'			if IsNULL(rsget("tenbeaexists")) then
'				FItemList(i).Ftenbeaexists = false
'			else
'				FItemList(i).Ftenbeaexists = true
'			end if

            FItemList(i).FDlvcountryCode = rsget("DlvcountryCode")

			rsget.movenext
			i=i+1
		loop
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
