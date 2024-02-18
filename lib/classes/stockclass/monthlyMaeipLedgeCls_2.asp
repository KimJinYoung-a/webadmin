<%
'###########################################################
' Description : 월별 재고월령 클래스
' History : 이상구 생성
'###########################################################

Class CMonthlyMaeipLedgeItem
    public FisJungsan
    public Fyyyymm
	public FstockPlace
	public Fshopid
	public FtargetGbn
	public Fitemgubun
    public Flastmwdiv
	public Flastvatinclude
    public FMakerid

	public Fitemid
	public Fitemoption

	public FprevSysStockNo
	public FprevSysStockSum

	public FIpgoNo
	public FIpgoSum
	public FMoveNo
	public FMoveSum
	public FSellNo
	public FSellSum
	public FOffChulNo
	public FOffChulSum
	public FEtcChulNo
	public FEtcChulSum
	public FCsNo
	public FCsSum
	public FLossChulNo
	public FLossChulSum
	public FcurSysStockNo
	public FcurSysStockSum
	public FcurErrRealCheckNo
	public FcurErrRealCheckSum

'    public function IsMoveItem()
'        IsMoveItem = false
'
'        if (FisJungsan) then
'            if (Fyyyymm>="2012-01") and (Fyyyymm<"2012-10") and (LCASE(Fmakerid)="ithinkso") then
'                IsMoveItem = true
'            end if
'        else
'            if (Fyyyymm>="2012-01") and (Fyyyymm<"2012-10") and (LCASE(Fmakerid)="ithinkso") then
'                IsMoveItem = true
'            end if
'        end if
'    end function

    public function getTotErrNo()
        getTotErrNo = getDiffNo*-1
    end function

    public function getTotErrSum()
        getTotErrSum = getDiffSum*-1
    end function

    public function getDiffNo()
        getDiffNo = FprevSysStockNo + getIpgoNo + getMoveNo + FSellNo + FOffChulNo + FEtcChulNo + FCsNo + FLossChulNo - FcurSysStockNo
    end function

    public function getDiffSum()
        getDiffSum = FprevSysStockSum + getIpgoSum + getMoveSum + FSellSum + FOffChulSum + FEtcChulSum + FCsSum + FLossChulSum - FcurSysStockSum
    end function

    public function getIpgoNo()
        getIpgoNo = FIpgoNo
    end function

    public function getIpgoSum()
        if isNULL(FIpgoSum) then
            getIpgoSum = 0
        else
            getIpgoSum = FIpgoSum
        end if
    end function

    public function getMoveNo()
        getMoveNo = FMoveNo
    end function

    public function getMoveSum()
        getMoveSum = FMoveSum
    end function

    public function getStockPlaceOrShopid
        if (Fshopid<>"") then
            getStockPlaceOrShopid = Fshopid
        else
            getStockPlaceOrShopid = FstockPlace
        end if
    end function

    public function getBusiName()
        getBusiName=""
        Exit function

        if (FtargetGbn="ON") then
		    getBusiName      = "온라인"
		elseif (FtargetGbn="OF") then
		    getBusiName      = "오프라인"
		elseif (FtargetGbn="AC") then
		    getBusiName      = "아카데미"
		elseif (FtargetGbn="IT") then
		    getBusiName      = "아이띵소(구)"
		elseif (FtargetGbn="ET") then
		    getBusiName      = "띵소"
	    elseif (FtargetGbn="EG") then
		    getBusiName      = "EG"
		else
		    getBusiName      = "-"
	    end if
    end function

    public function getItemGubunName()
        if Fitemgubun="10" then
			getITemGubunName = "일반"
		elseif Fitemgubun="90" then
			getITemGubunName = "오프전용"
		elseif Fitemgubun="60" then
			getITemGubunName = "기타"
		elseif Fitemgubun="70" then
			getITemGubunName = "소모품"
		elseif Fitemgubun="75" then
			getITemGubunName = "저장품"
		elseif Fitemgubun="80" then
			getITemGubunName = "사은품"
		elseif Fitemgubun="85" then
			getITemGubunName = "사은품"
		elseif Fitemgubun="97" then
			getITemGubunName = "강좌"
		elseif Fitemgubun="98" then
			getITemGubunName = "DIY"
		elseif Fitemgubun="99" then
			getITemGubunName = "일반"
		elseif Fitemgubun="95" then
			getITemGubunName = "기타"
		else
			getITemGubunName = "기타" ''Fitemgubun
		end if
    end function

    public function getMeaipTypeName()

        if Flastmwdiv="M" then
			getMeaipTypeName = "입고분매입"
		elseif Flastmwdiv="S" then
			getMeaipTypeName = "판매분매입"
		elseif Flastmwdiv="C" then
			getMeaipTypeName = "출고분매입"
		elseif Flastmwdiv="E" then
			getMeaipTypeName = "기타매입"
		elseif Flastmwdiv="W" then
			getMeaipTypeName = "입고분매입(W)"
		elseif Flastmwdiv="U" then
			getMeaipTypeName = "업체<br>(U)"
		elseif Flastmwdiv="Z" then
			getMeaipTypeName = "-<br>(Z)"
		elseif Flastmwdiv="J" then
			getMeaipTypeName = "판매(출고)분매입"
		elseif Flastmwdiv="B011" then
			getMeaipTypeName = "위탁판매<br>(B011)"
		elseif Flastmwdiv="B012" then
			getMeaipTypeName = "업체위탁<br>(B012)"
		elseif Flastmwdiv="B013" then
			getMeaipTypeName = "출고위탁"
		elseif Flastmwdiv="B021" then
			getMeaipTypeName = "오프매입"
		elseif Flastmwdiv="B022" then
			getMeaipTypeName = "매장매입"
		elseif Flastmwdiv="B023" then
			getMeaipTypeName = "가맹점매입"
		elseif Flastmwdiv="B031" then
			getMeaipTypeName = "출고매입"
		elseif Flastmwdiv="B032" then
			getMeaipTypeName = "센터매입"
		else
			getMeaipTypeName = Flastmwdiv
		end if

    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class

Class CStockOverValueSum
    public FtargetGbn
	public Fitemgubun
	public FMaeIpGubun

	public FpurchasetypeStr

	public FTotBuySum1
	public FTotBuySum2
	public FTotBuySum3
	public FTotBuySum4
	public FTotBuySum5
	public FTotBuySum6
	public FTotBuySum7
	public FTotBuySum8

	public FTotBuySum11
	public FTotBuySum12
	public FTotBuySum13
	public FTotBuySum14

	public FTotBuySum

	Public Fshopid
	public Fmakerid
	public FlastIpgoDate
	public FtotStockNo

	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public FbuyPrice

	public FLstComm_cd


	public function getOverValueStockPrice
		' * 최종입고월을 기준으로 재고월령을 산정합니다.
		' * 재고월령이 1년을 넘는 상품의 경우 재고평가충당금(재고평가손실)을 적용합니다.
		' * 재고월령이 1년-2년 사이인 경우 매입가 대비 50% 의 평가충당금을 산정합니다.
		' * 재고월령이 2년을 넘는 경우 매입가 대비 100% 의 평가충당금을 산정합니다.
		getOverValueStockPrice = (FTotBuySum4 / 2) + FTotBuySum5 + FTotBuySum6
	end function

	public function getOverValueStockPriceYear
		' * 최종입고월을 기준으로 재고월령을 산정합니다.
		' * 재고월령이 1년을 넘는 상품의 경우 재고평가충당금(재고평가손실)을 적용합니다.
		' * 재고월령이 1년-2년 사이인 경우 매입가 대비 50% 의 평가충당금을 산정합니다.
		' * 재고월령이 2년을 넘는 경우 매입가 대비 100% 의 평가충당금을 산정합니다.
		getOverValueStockPriceYear = (FTotBuySum12 / 2) + FTotBuySum13 + FTotBuySum14 + FTotBuySum6
	end function

	public function getLastCommCD
		''B011 : 위탁판매, B012 : 업체위탁, B013 : 출고위탁, B021 : 오프매입, B022 : 매장매입, B023 : 가맹점매입, B031 : 출고매입, B032 : 센터매입

		select case FLstComm_cd
			case "B011"
				getLastCommCD = "위탁판매"
			case "B012"
				getLastCommCD = "업체위탁"
			case "B013"
				getLastCommCD = "출고위탁"
			case "B021"
				getLastCommCD = "오프매입"
			case "B022"
				getLastCommCD = "매장매입"
			case "B023"
				getLastCommCD = "가맹점매입"
			case "B031"
				getLastCommCD = "출고매입"
			case "B032"
				getLastCommCD = "센터매입"
			case else
				getLastCommCD = FLstComm_cd
		end select

	end function

	public function GetlastIpgoDate
		if (FlastIpgoDate = "1") then
			GetlastIpgoDate = "1개월~3개월"
		elseif (FlastIpgoDate = "2") then
			GetlastIpgoDate = "4개월~6개월"
		elseif (FlastIpgoDate = "3") then
			GetlastIpgoDate = "7개월~12개월"
		elseif (FlastIpgoDate = "4") then
			GetlastIpgoDate = "1년~2년"
		elseif (FlastIpgoDate = "5") then
			GetlastIpgoDate = "2년초과"
		elseif (FlastIpgoDate = "6") then
			GetlastIpgoDate = "NULL"
		elseif (FlastIpgoDate = "7") then
			GetlastIpgoDate = "13개월~18개월"
		elseif (FlastIpgoDate = "8") then
			GetlastIpgoDate = "19개월~24개월"
		else
			GetlastIpgoDate = FlastIpgoDate
		end if
	end function

    public function getBusiName
        if (FtargetGbn="ON") then
		    getBusiName      = "온라인"
		elseif (FtargetGbn="OF") then
		    getBusiName      = "오프라인"
		elseif (FtargetGbn="IT") then
		    getBusiName      = "아이띵소(구)"
		elseif (FtargetGbn="ET") then
		    getBusiName      = "3PL(아이띵소)"
		elseif (FtargetGbn="EG") then
		    getBusiName      = "3PL(유그레잇)"
		else
		    getBusiName      = "-"
	    end if
	end function

	public function getMaeipGubunName()
		if FMaeIpGubun="M" then
			getMaeipGubunName = "매입"
		elseif FMaeIpGubun="W" then
			getMaeipGubunName = "위탁"
		elseif FMaeIpGubun="U" then
			getMaeipGubunName = "업체"
		elseif FMaeIpGubun="Z" then
			getMaeipGubunName = "-"
		elseif FMaeIpGubun="B011" then
			getMaeipGubunName = "위탁판매"
		elseif FMaeIpGubun="B012" then
			getMaeipGubunName = "업체위탁"
		elseif FMaeIpGubun="B013" then
			getMaeipGubunName = "출고위탁"
		elseif FMaeIpGubun="B021" then
			getMaeipGubunName = "오프매입"
		elseif FMaeIpGubun="B022" then
			getMaeipGubunName = "매장매입"
		elseif FMaeIpGubun="B023" then
			getMaeipGubunName = "가맹점매입"
		elseif FMaeIpGubun="B031" then
			getMaeipGubunName = "출고매입"
		elseif FMaeIpGubun="B032" then
			getMaeipGubunName = "센터매입"
		else
			getMaeipGubunName = FMaeIpGubun
		end if

		IF isNULL(FMaeIpGubun) then
		    getMaeipGubunName ="-"
		end if
	end function

	public function getITemGubunName()
		if Fitemgubun="10" then
			getITemGubunName = "일반"
		elseif Fitemgubun="90" then
			getITemGubunName = "오프전용"
		elseif Fitemgubun="70" then
			getITemGubunName = "소모품"
		elseif Fitemgubun="75" then
			getITemGubunName = "저장품"
		elseif Fitemgubun="80" then
			getITemGubunName = "사은품"
		elseif Fitemgubun="85" then
			getITemGubunName = "사은품"
		else
			getITemGubunName = Fitemgubun
		end if
	end function

	public function getITemMaeipGubunName()
		if FItemMaeIpGubun = "M" then
			getITemMaeipGubunName = "<font color=red>매입</font>"
		elseif FItemMaeIpGubun = "W" then
			getITemMaeipGubunName = "위탁"
		else
			getITemMaeipGubunName = FItemMaeIpGubun
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMonthlyMaeipLedge
	public FItemList()
	public FOneItem
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

    public FRectYYYY
	public FRectYYYYMM
	public FRectStockPlace
	public FRectShopid
	public FRectMakerid
    public FRectBySuplyPrice
    public FRectMeaipTp
    public FRectItemgubun
    public FRectTargetGbn

    public FRectSubGrpType
	public FRectShowShopid
    public FRectOnlyIpgoMeaip

	public FRectShowDiff
	public FRectPriceGubun

	public FRectGubun
	public FRectMwDiv
	public FRectVatYn
	public FRectShopSuplyPrice
	public FRectetcjungsantype
	public FRectLastIpgoGBN

	public FRectMonthGubun
	public FRectOrdTp
	public FRectShowItemList
	public frectreqYYYYMM
	public frectreqStrplace
	public fArrLIst
	public frectIsUsingV2
	public frectver
	public frectplaceGubun
	public frectPriceGbn

    '' 재고 위치 T 인경우 당기매입은 (부자재 등이 포함되어 있음) ==> 들어왔다 바로 나감 (계산서 발행금액<>재고매입가)
    function getCaseStrNo(iyyyymm,ifieldNm)
        dim AddCASEStr
        ''입고 이동 분리 ==>디비 플래그 생성.
        if (ifieldNm="stIpgoNo") or (ifieldNm="totItemNo") then
            ''AddCASEStr = " and NOT ((isNULL(m.isMove,0)<>0) or (m.stockPlace<>'T' and m.makerid in ('ithinkso','grandmintfestival','beautifulmintlife') and m.yyyymm>='2012-01' and m.yyyymm<'2012-10') or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))" ''
            AddCASEStr = " and NOT ((isNULL(m.isMove,0)<>0) or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))" ''
        elseif (ifieldNm="stIpgoMoveNo") then
            ''AddCASEStr = " and ((isNULL(m.isMove,0)<>0) or (m.stockPlace<>'T' and m.makerid in ('ithinkso','grandmintfestival','beautifulmintlife') and m.yyyymm>='2012-01' and  m.yyyymm<'2012-10') or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))" ''
            AddCASEStr = " and ((isNULL(m.isMove,0)<>0) or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))" ''
            ifieldNm = "stIpgoNo"
        end if

        if (ifieldNm<>"curSysStockNo") and (FRectYYYY<>"") then ''년도별 합계.
            getCaseStrNo = "case when LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then IsNull("+ifieldNm+",0) else 0 end"
        else
            getCaseStrNo = "case when m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then IsNull("+ifieldNm+",0) else 0 end"
        end if
    end function

    function getCaseStrPrice(iyyyymm,ifieldNmNo,ifieldNmPrc)
        dim AddCASEStr, AddCASENoStr
        ''입고 이동 분리
        if (ifieldNmNo="stIpgoNo") then
            ''AddCASEStr = " and NOT ((isNULL(m.isMove,0)<>0) or (m.stockPlace<>'T' and m.makerid in ('ithinkso','grandmintfestival','beautifulmintlife') and m.yyyymm>='2012-01' and  m.yyyymm<'2012-10') or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
            AddCASEStr = " and NOT ((isNULL(m.isMove,0)<>0) or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
        elseif (ifieldNmNo="stIpgoMoveNo") then
            ''AddCASEStr = " and ((isNULL(m.isMove,0)<>0) or (m.stockPlace<>'T' and m.makerid in ('ithinkso','grandmintfestival','beautifulmintlife') and m.yyyymm>='2012-01' and  m.yyyymm<'2012-10') or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
            AddCASEStr = " and ((isNULL(m.isMove,0)<>0) or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
            ifieldNmNo = "stIpgoNo"
        end if

		AddCASENoStr = ifieldNmNo
		if (FRectPriceGubun = "V") and (ifieldNmNo = "stIpgoNo" and ifieldNmPrc = "totBuyCash") then
			AddCASENoStr = "1"  ''2014/10/13

			'AddCASENoStr = "(CASE WHEN i.totitemno=0 THEN 1 ELSE i.totitemno END)" ''2016/06/02 재수정 수량이 0 이나 매입가가 있을수 있음..
			'ifieldNmPrc = "(CASE WHEN i.totitemno=0 THEN "&ifieldNmPrc&" ELSE "&ifieldNmPrc&"/"&AddCASENoStr&" END)" ''2016/06/02 재수정
		end if

        if (ifieldNmNo<>"curSysStockNo") and (FRectYYYY<>"") then
            if (FRectBySuplyPrice=1) then ''Round 관련 오차 있음
				if (FRectPriceGubun = "V") then
					'평균매입가는 단가에 대해 평균매입가 산정
					getCaseStrPrice = "case when LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN Round((IsNull("+ifieldNmPrc+",0)*10/11), 0) ELSE IsNull("+ifieldNmPrc+",0) END) else 0 end"
				else
					getCaseStrPrice = "case when LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN (IsNull("+ifieldNmPrc+",0)*10/11) ELSE IsNull("+ifieldNmPrc+",0) END) else 0 end"
				end if
            else
               getCaseStrPrice = "case when LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then "+AddCASENoStr+"*IsNull("+ifieldNmPrc+",0) else 0 end"
            end if
        else
            if (FRectBySuplyPrice=1) then ''Round 관련 오차 있음
				if (FRectPriceGubun = "V") then
					'평균매입가는 단가에 대해 평균매입가 산정
					getCaseStrPrice = "case when m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN Round((IsNull("+ifieldNmPrc+",0)*10/11), 0) ELSE IsNull("+ifieldNmPrc+",0) END) else 0 end"
				else
					getCaseStrPrice = "case when m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN (IsNull("+ifieldNmPrc+",0)*10/11) ELSE IsNull("+ifieldNmPrc+",0) END) else 0 end"
				end if
            else
               getCaseStrPrice = "case when m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then "+AddCASENoStr+"*IsNull("+ifieldNmPrc+",0) else 0 end"
            end if
        end if
    end function

    public Function GetMaeipJungsanSumSubDetail
        FRectSubGrpType = "makerid"

		if (FRectMakerid <> "") then
			FRectSubGrpType = "itemid"
		end if

        call GetMaeipJungsanSum
    end Function

    public Function GetMaeipJungsanSum
        dim sqlStr, addSql, i
		dim prevYYYYMM

        IF (FRectYYYY>="2014") then FRectYYYY="1998"

        IF (FRectYYYY<>"") then ''년도별 조회인경우
            FRectYYYYMM=FRectYYYY+"-12"
            prevYYYYMM = Left(dateAdd("m",-1,FRectYYYY+"-01-01"),7)
        else
		    prevYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)
	    end if

		''prevYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)

        addSql = " 	from db_summary.dbo.tbl_monthly_jungsanSum"
        addSql = addSql + " 	where 1=1"
        addSql = addSql + " 	and jGubun<>'CC'" '' 수수료는 표시안함
        addSql = addSql + " 	and jTaxType<>'03'" '' 원천징수 제외  ''addSql = addSql + " 	and itemgubun<>'97'" '' 강좌는 매입 아님
        addSql = addSql + " 	and NOT (jmakerid in ('ithinkso','grandmintfestival','beautifulmintlife') and (yyyymm>='2012-01') and (yyyymm<'2013-11') and jMwdiv<>'M') " ''일단 아이띵소 제외 출고분 매입 존재.. 2012-01~2012-09 까지는 정산내역에서 빠져야함.


        if (FRectOnlyIpgoMeaip="on") then
            'addSql = addSql + " 	and jMwdiv='M'"  '' 입고분 매입은 표시안함
            addSql = addSql + " 	and ((jMwdiv='M') or (dtlGubunCD='B022'))"  '' 입고분 매입은 표시안함
        else
            addSql = addSql + " 	and jMwdiv<>'M'"  '' 입고분 매입은 표시안함
            addSql = addSql + " 	and dtlGubunCD<>'B022'" ''매장매입 제외
        end if

        addSql = addSql + " 	and yyyymm>'"&prevYYYYMM&"' and yyyymm<='"&FRectYYYYMM&"'"  ''정산은 = 가 없음
        if (FRectMakerid<>"") then
            addSql = addSql + " 	and jmakerid='"&FRectMakerid&"'"
        end if

        if (FRectItemGubun<>"") then
            addSql = addSql + " 	and itemgubun='"&FRectItemGubun&"'"
        end if

		if (FRectShopid <> "") then
			addSql = addSql + " 	and jShopid='"&FRectShopid&"'"
		end if

        if (FRectTargetGbn<>"") then
            addSql = addSql + " 	and jtargetGbn='"&FRectTargetGbn&"'"
        end if

        if (FRectMeaipTp<>"") then
            addSql = addSql + " 	and jMwdiv='"&FRectMeaipTp&"'"
        end if

        ''if (FRectStockPlace="L") then
        ''    addSql = addSql + " and ((jtargetGbn<>'OF') or ( (jtargetGbn='OF') and (jMwdiv='M' and dtlgubuncd in ('B021')) ) )"
        ''elseif (FRectStockPlace="S") then
        ''    addSql = addSql + " and NOT ((jtargetGbn<>'OF') or ((jtargetGbn='OF') and (jMwdiv='M' and dtlgubuncd in ('B021'))) )"
        ''end if

        addSql = addSql + " 	group by itemgubun,jMwdiv"  '''jtargetGbn,
		if (FRectShowShopid <> "") then
			addSql = addSql + " 	,jshopid "
		end if
        IF (FRectSubGrpType="makerid") then
		    addSql = addSql + " 	, jmakerid "
		elseif (FRectSubGrpType = "itemid") then
			addSql = addSql + " 	, jmakerid, itemid, itemoption "
		end if


        if (FRectSubGrpType<>"") then
    		sqlStr = " SELECT top 1 (COUNT(*) OVER ()) as CNT, CEILING(CAST((COUNT(*) OVER ()) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
    		sqlStr = sqlStr + addSql

        	rsget.Open sqlStr,dbget,1
    		if  not rsget.EOF  then
    			FTotalCount = rsget("cnt")
    			FTotalPage = rsget("totPg")
    		else
    			FTotalCount = 0
    			FTotalPage = 0
    		end if
        	rsget.Close

        	'지정페이지가 전체 페이지보다 클 때 함수종료
        	if CLng(FCurrPage)>CLng(FTotalPage) then
        		FResultCount = 0
        		exit function
        	end if
        end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	'" + CStr(FRectYYYYMM) + "' as yyyymm "
		sqlStr = sqlStr + " 	,'"&FRectStockPlace&"' as stockPlace "

		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,jshopid as shopid "
		end if

        '''sqlStr = sqlStr + " 	,jtargetGbn as targetGbn"
        sqlStr = sqlStr + " 	,itemgubun"
        IF (FRectSubGrpType="makerid") then
		    sqlStr = sqlStr + " 	,jmakerid "
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	, jmakerid, itemid, itemoption "
		end if
        sqlStr = sqlStr + " 	,jMwdiv as lastmwdiv"
        sqlStr = sqlStr + " 	,0 as prevSysStockNo"
        sqlStr = sqlStr + " 	,0 as prevSysStockSum"
        if (FRectBySuplyPrice="1") then
            sqlStr = sqlStr + " 	,sum(jtotItemno) as IpgoNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END) as IpgoSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='S' THEN jtotItemno*-1 ELSE 0 END) as SellNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='S' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as SellSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='C' THEN jtotItemno*-1 ELSE 0 END) as OffChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='C' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as OffChulSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='E' THEN jtotItemno*-1 ELSE 0 END) as EtcChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='E' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as EtcChulSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='T' THEN jtotItemno*-1 ELSE 0 END) as CsNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='T' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as CsSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='L' THEN jtotItemno*-1 ELSE 0 END) as LossChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='L' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as LossChulSum"
        else
            sqlStr = sqlStr + " 	,sum(jtotItemno) as IpgoNo"
            sqlStr = sqlStr + " 	,sum(jtotBuycash) as IpgoSum"

            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='S' THEN jtotItemno*-1 ELSE 0 END) as SellNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='S' THEN jtotBuycash*-1 ELSE 0 END) as SellSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='C' THEN jtotItemno*-1 ELSE 0 END) as OffChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='C' THEN jtotBuycash*-1 ELSE 0 END) as OffChulSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='E' THEN jtotItemno*-1 ELSE 0 END) as EtcChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='E' THEN jtotBuycash*-1 ELSE 0 END) as EtcChulSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='T' THEN jtotItemno*-1 ELSE 0 END) as CsNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='T' THEN jtotBuycash*-1 ELSE 0 END) as CsSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='L' THEN jtotItemno*-1 ELSE 0 END) as LossChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='L' THEN jtotBuycash*-1 ELSE 0 END) as LossChulSum"
        end if
        sqlStr = sqlStr + " 	,0 as curSysStockNo"
        sqlStr = sqlStr + " 	,0 as curSysStockSum"
        sqlStr = sqlStr + " 	,0 as curErrRealCheckNo"
        sqlStr = sqlStr + " 	,0 as curErrRealCheckSum"

		sqlStr = sqlStr + addSql

        sqlStr = sqlStr + " 	order by (CASE WHEN itemgubun='00' THEN '999' ELSE itemgubun END) asc ,jMwdiv desc"  '''(CASE WHEN jtargetGbn='TT' THEN 'AA' ELSE jtargetGbn END)  desc,
		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,jshopid "
		end if
        IF (FRectSubGrpType="makerid") then
		    sqlStr = sqlStr + " 	,jmakerid "
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	, jmakerid, itemid, itemoption "
		end if

		''rw 	sqlStr

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
				set FItemList(i) = new CMonthlyMaeipLedgeItem
                FItemList(i).FisJungsan         = true
				FItemList(i).Fyyyymm     		= rsget("yyyymm")
				FItemList(i).FstockPlace    	= rsget("stockPlace")
				''FItemList(i).FtargetGbn     	= rsget("targetGbn")
				FItemList(i).Fitemgubun     	= rsget("itemgubun")
                FItemList(i).Flastmwdiv         = rsget("lastmwdiv")

				if (FRectShowShopid <> "") then
					FItemList(i).Fshopid         = rsget("shopid")
				end if

				FItemList(i).FprevSysStockNo    = rsget("prevSysStockNo")
				FItemList(i).FprevSysStockSum   = rsget("prevSysStockSum")

				FItemList(i).FIpgoNo     		= rsget("IpgoNo")
				FItemList(i).FIpgoSum     		= rsget("IpgoSum")
				'FItemList(i).FMoveNo     		= rsget("MoveNo")
				'FItemList(i).FMoveSum     		= rsget("MoveSum")
				FItemList(i).FSellNo     		= rsget("SellNo")
				FItemList(i).FSellSum     		= rsget("SellSum")
				FItemList(i).FOffChulNo     	= rsget("OffChulNo")
				FItemList(i).FOffChulSum     	= rsget("OffChulSum")
				FItemList(i).FEtcChulNo     	= rsget("EtcChulNo")
				FItemList(i).FEtcChulSum     	= rsget("EtcChulSum")
				FItemList(i).FCsNo     			= rsget("CsNo")
				FItemList(i).FCsSum     		= rsget("CsSum")
				FItemList(i).FLossChulNo     	= rsget("LossChulNo")
				FItemList(i).FLossChulSum     	= rsget("LossChulSum")

				FItemList(i).FcurSysStockNo     = rsget("curSysStockNo")
				FItemList(i).FcurSysStockSum    = rsget("curSysStockSum")

				FItemList(i).FcurErrRealCheckNo = rsget("curErrRealCheckNo")
				FItemList(i).FcurErrRealCheckSum	= rsget("curErrRealCheckSum")

                IF (FRectSubGrpType="makerid") then
                    FItemList(i).FMakerid = rsget("jmakerid")
				elseif (FRectSubGrpType = "itemid") then
					FItemList(i).FMakerid 		= rsget("jmakerid")
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
                end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    public Function GetMaeipLedgeSUMSubDetail
        FRectSubGrpType = "makerid"

		if (FRectMakerid <> "") then
			FRectSubGrpType = "itemid"
		end if

        call GetMaeipLedgeSUM
    end Function

	'/admin/newreport/monthlyMaeipLedge_excel_download.asp
	public Sub GetMaeipLedgeListNotPaging()
		dim sqlStr,i

		if (frectver = "V2") then
		    '' sp_Ten_monthlyMaeipLedge_MakeEXL_List_V2 => sp_Ten_monthlyMaeipLedge_MakeEXL_List_V2_1 ''임시변경 2015/01/12
			sqlStr ="exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_List_V2_1] '" & frectyyyymm & "','" & frectplaceGubun & "'," & FCurrPage & "," & FPageSize & ",'" + CStr(frectPriceGbn) + "'"  ''위치변경 2015/04/13
		elseif (frectver = "DW") then
			sqlStr ="exec [db_datamart].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_List_DW] '" & frectyyyymm & "','" & frectplaceGubun & "'," & FCurrPage & "," & FPageSize & ",'" + CStr(frectPriceGbn) + "'"
		else
			sqlStr ="exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_List] '" & frectyyyymm & "','" & frectplaceGubun & "'," & FCurrPage & "," & FPageSize & ",'" + CStr(frectPriceGbn) + "'"
		end if

		'response.write sqlStr & "<br>"
		if (ver = "DW") then
    		db3_rsget.CursorLocation = adUseClient
			db3_dbget.CommandTimeout = 60*5   ' 5분
			db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly', adCmdStoredProc

			FTotalCount = db3_rsget.RecordCount
			FResultCount = db3_rsget.RecordCount

			IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
				fArrLIst = db3_rsget.getRows()
			END IF
			db3_rsget.close
		else
    		rsget.CursorLocation = adUseClient
			dbget.CommandTimeout = 60*5   ' 5분
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly', adCmdStoredProc

			FTotalCount = rsget.RecordCount
			FResultCount = rsget.RecordCount

			IF Not (rsget.EOF OR rsget.BOF) THEN
				fArrLIst = rsget.getRows()
			END IF
			rsget.close
		end if
	end sub

	public Function GetMaeipLedgeSUM_PROC
		dim sqlStr, i

		sqlStr = " EXEC [db_summary].[dbo].[sp_Ten_monthly_Maeip_Stockledger_SUM] '" & FRectYYYYMM & "', '" & FRectStockPlace & "', '" & FRectMakerid & "', '" & FRectShopid & "', '" & CHKIIF(FRectBySuplyPrice=1, "Y", "") & "', '" & CHKIIF(FRectShowShopid <> "", "Y", "") & "' "
''rw  sqlStr
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMonthlyMaeipLedgeItem

                FItemList(i).FisJungsan         = false
				FItemList(i).Fyyyymm     		= rsget("yyyymm")
				FItemList(i).FstockPlace    	= rsget("stockPlace")
				''FItemList(i).FtargetGbn     	= rsget("targetGbn")
				FItemList(i).Fitemgubun     	= rsget("itemgubun")
                FItemList(i).Flastmwdiv         = rsget("lastmwdiv")

				if (FRectShowShopid <> "") then
					FItemList(i).Fshopid         = rsget("shopid")
				end if

				FItemList(i).FprevSysStockNo    = rsget("prevSysStockNo")
				FItemList(i).FprevSysStockSum   = rsget("prevSysStockSum")

				FItemList(i).FIpgoNo     		= rsget("IpgoNo")
				FItemList(i).FIpgoSum     		= rsget("IpgoSum")
				FItemList(i).FMoveNo     		= rsget("MoveNo")
				FItemList(i).FMoveSum     		= rsget("MoveSum")
				FItemList(i).FSellNo     		= rsget("SellNo")
				FItemList(i).FSellSum     		= rsget("SellSum")
				FItemList(i).FOffChulNo     	= rsget("OffChulNo")
				FItemList(i).FOffChulSum     	= rsget("OffChulSum")
				FItemList(i).FEtcChulNo     	= rsget("EtcChulNo")
				FItemList(i).FEtcChulSum     	= rsget("EtcChulSum")
				FItemList(i).FCsNo     			= rsget("CsNo")
				FItemList(i).FCsSum     		= rsget("CsSum")
				FItemList(i).FLossChulNo     	= rsget("LossChulNo")
				FItemList(i).FLossChulSum     	= rsget("LossChulSum")

				FItemList(i).FcurSysStockNo     = rsget("curSysStockNo")
				FItemList(i).FcurSysStockSum    = rsget("curSysStockSum")

				FItemList(i).FcurErrRealCheckNo = rsget("curErrRealCheckNo")
				FItemList(i).FcurErrRealCheckSum	= rsget("curErrRealCheckSum")

                IF (FRectSubGrpType="makerid") then
                    FItemList(i).FMakerid = rsget("makerid")
				elseif (FRectSubGrpType = "itemid") then
					FItemList(i).FMakerid 		= rsget("makerid")
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).Fitemoption    	= rsget("itemoption")
					FItemList(i).Flastvatinclude    = rsget("lastvatinclude")
                end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	End Function


	public Function GetMaeipLedgeSUM
		dim sqlStr, addSql, i
		dim prevYYYYMM

		IF (FRectYYYY<>"") and (FRectYYYY>="2014") then FRectYYYY="1998"

        if (FRectYYYY<>"") then ''년도별 조회인경우
            FRectYYYYMM=FRectYYYY+"-12"
            prevYYYYMM = Left(dateAdd("m",-1,FRectYYYY+"-01-01"),7)
        else
		    prevYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)
	    end if

		addSql = " from "
		addSql = addSql + " 	db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail_V2 m "
		addSql = addSql + " 	left join db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum_V2 i "
		addSql = addSql + " 	on "
		addSql = addSql + " 		1 = 1 "
		addSql = addSql + " 		and i.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		addSql = addSql + " 		and m.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		addSql = addSql + " 		and m.yyyymm = i.yyyymm "
		addSql = addSql + " 		and m.stockPlace = i.stockPlace "
		addSql = addSql + " 		and IsNull(i.targetGbn, 'ON') in ('ON', 'OF', 'AC') "
		''addSql = addSql + " 		and m.targetGbn = i.targetGbn "
		addSql = addSql + " 		and m.shopid = i.shopid "
		addSql = addSql + " 		and m.itemgubun = i.itemgubun "
		addSql = addSql + " 		and m.itemid = i.itemid "
		addSql = addSql + " 		and m.itemoption = i.itemoption "
		addSql = addSql + " 		and i.ipgoMWDIV = 'M' "
		addSql = addSql + " where "
		addSql = addSql + " 	1 = 1 "
		addSql = addSql + " 	and m.yyyymm>='"&prevYYYYMM&"' and m.yyyymm<='"&FRectYYYYMM&"'"
		addSql = addSql + " 	and m.targetGbn not in ('ET','EG')"
		addSql = addSql + " 	and m.etcjungsantype in (1,4)"
		''addSql = addSql + " 	and m.lastmwdiv not in ('B012')"

        addSql = addSql + " 	and NOT (m.lastmwdiv='B013' and m.targetGbn<>'IT')"               ''출고위탁은 IT만
        addSql = addSql + " 	and m.lastmwdiv not in ('W','B012','B011')"                     ''업체위탁 제외 //재고자산이 매입이 아닌 CASE (재고자산 형태로 뿌릴경우) W제외

		if (FRectStockPlace <> "") then
			addSql = addSql + " 	and m.stockPlace = '" + CStr(FRectStockPlace) + "' "
		end if

		''addSql = addSql + " 	and m.stockPlace = 'L' "

        if (FRectMakerid<>"") then
            addSql = addSql + " 	and m.makerid='"&FRectMakerid&"'"
        end if

        if (FRectItemGubun<>"") then
            addSql = addSql + " 	and m.itemgubun='"&FRectItemGubun&"'"
        end if

		if (FRectShopid <> "") then
			addSql = addSql + " 	and m.shopid='"&FRectShopid&"'"
		end if

        if (FRectTargetGbn<>"") then
            addSql = addSql + " 	and m.targetGbn='"&FRectTargetGbn&"'"
        end if

        if (FRectMeaipTp<>"") then
            ''addSql = addSql + " 	and ((isNULL(m.lastmwdiv,'unknown')='"&FRectMeaipTp&"') or (isNULL(m.lastmwdiv,'unknown')='W' and '" + CStr(FRectMeaipTp) + "' = 'M')) "
        end if



		addSql = addSql + " group by "
		addSql = addSql + " 	m.stockPlace "
		''addSql = addSql + " 	,m.targetGbn "
		addSql = addSql + " 	,m.itemgubun "
		if (FRectShowShopid <> "") then
			addSql = addSql + " 	,m.shopid "
		end if

		''addSql = addSql + " 	,isNULL(lastmwdiv,'unknown') "
		IF (FRectSubGrpType="makerid") then
		    addSql = addSql + " 	,m.makerid "
		elseif (FRectSubGrpType = "itemid") then
			addSql = addSql + " 	,m.makerid, m.itemid, m.itemoption, m.lastvatinclude "
		end if

		'// having
		if (FRectShowDiff <> "") then
			'addSql = addSql + " 	having (sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") + sum(case when m.yyyymm = '" + CStr(FRectYYYYMM) + "' then IpgoNo else 0 end) + sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&")) <> 0 "
			if (FRectYYYY<>"") then
			    addSql = addSql + " 	having (sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") + sum(case when LEFT(m.yyyymm,4) = '" + CStr(FRectYYYY) + "' then stIpgoNo else 0 end) + sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&")) <> 0 "
			else
			    addSql = addSql + " 	having (sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") + sum(case when m.yyyymm = '" + CStr(FRectYYYYMM) + "' then stIpgoNo else 0 end) + sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&")) <> 0 "
		    end if
		end if


		sqlStr = " SELECT top 1 (COUNT(*) OVER ()) as CNT, CEILING(CAST((COUNT(*) OVER ()) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
		sqlStr = sqlStr + addSql
		''response.write sqlStr
		''response.end

        if (FRectSubGrpType<>"") then
        	rsget.Open sqlStr,dbget,1
    		if  not rsget.EOF  then
    			FTotalCount = rsget("cnt")
    			FTotalPage = rsget("totPg")
    		else
    			FTotalCount = 0
    			FTotalPage = 0
    		end if
        	rsget.Close


        	'지정페이지가 전체 페이지보다 클 때 함수종료
        	if CLng(FCurrPage)>CLng(FTotalPage) then
        		FResultCount = 0
        		exit function
        	end if
        end if

		sqlStr = " select "
		sqlStr = sqlStr + " 	'" + CStr(FRectYYYYMM) + "' as yyyymm "
		sqlStr = sqlStr + " 	,m.stockPlace "
		''sqlStr = sqlStr + " 	,m.targetGbn "
		sqlStr = sqlStr + " 	,m.itemgubun "
		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,m.shopid "
		end if
		IF (FRectSubGrpType="makerid") then
			'// 브랜드 바뀌는걸 한줄로 표시해야 하는데 쿼리가 잘 안나옴. skyer9, 2017-10-11
		    sqlStr = sqlStr + " 	,m.makerid as makerid " ''다른 브랜드 하나로 표시
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	,m.makerid, m.itemid, m.itemoption, m.lastvatinclude "
		end if
		''2015-03-17, skyer9
		''sqlStr = sqlStr + " 	,'M' as lastmwdiv"
		sqlStr = sqlStr + " 	,(case when m.stockPlace in ('L', 'S') then 'M' else 'J' end) as lastmwdiv "

		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") as prevSysStockNo "
		''sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"IpgoNo")&") as IpgoNo "
		''sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"stIpgoNo")&") as IpgoNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"stIpgoMoveNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulMoveNo")&") as MoveNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") as SellNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"OffChulMoveNo")&") as OffChulNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") as EtcChulNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") as CsNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") as LossChulNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&") as curSysStockNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"curErrRealCheckNo")&") as curErrRealCheckNo "

		if (FRectPriceGubun = "V") then
			sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"totItemNo")&") as IpgoNo "

			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(prevYYYYMM,"curSysStockNo","avgIpgoPrice")&") as prevSysStockSum "
			''sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"1","IpgoSum")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoNo","totBuyCash")&") as IpgoSum "
			''sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoMoveNo","totBuyCash")&") + IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","avgIpgoPrice")&"),0) as MoveSum " '// totBuyCash : 물류-매장 매입구분 동일한 경우 물류평균매입가
			'// 2015-02-05, totBuyCash -> avgIpgoPrice
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoMoveNo","avgIpgoPrice")&") + IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","avgIpgoPrice")&"),0) as MoveSum " '// totBuyCash : 물류-매장 매입구분 동일한 경우 물류평균매입가
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"SellNo","avgIpgoPrice")&") as SellSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"OffChulNo","avgIpgoPrice")&") - IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","avgIpgoPrice")&"),0) as OffChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"EtcChulNo","avgIpgoPrice")&") as EtcChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"CsNo","avgIpgoPrice")&") as CsSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"LossChulNo","avgIpgoPrice")&") as LossChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curSysStockNo","avgIpgoPrice")&") as curSysStockSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curErrRealCheckNo","avgIpgoPrice")&") as curErrRealCheckSum "
		else
			sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"stIpgoNo")&") as IpgoNo "

			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(prevYYYYMM,"curSysStockNo","lastbuyPrice")&") as prevSysStockSum "
			''sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"1","IpgoSum")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoNo","lastbuyPrice")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoMoveNo","lastbuyPrice")&") + IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","lastbuyPrice")&"),0) as MoveSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"SellNo","lastbuyPrice")&") as SellSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"OffChulNo","lastbuyPrice")&") - IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","lastbuyPrice")&"),0) as OffChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"EtcChulNo","lastbuyPrice")&") as EtcChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"CsNo","lastbuyPrice")&") as CsSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"LossChulNo","lastbuyPrice")&") as LossChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curSysStockNo","lastbuyPrice")&") as curSysStockSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curErrRealCheckNo","lastbuyPrice")&") as curErrRealCheckSum "
		end if

		sqlStr = sqlStr + addSql

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	(case when (case when m.stockPlace in ('L', 'S') then 'M' else 'J' end) = 'M' then 1 else 100 end) "
		sqlStr = sqlStr + " 	,m.stockPlace "
		sqlStr = sqlStr + " 	,m.itemgubun "
		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,m.shopid "
		end if
		''sqlStr = sqlStr + " 	,lastmwdiv "
		IF (FRectSubGrpType="makerid") then
		    sqlStr = sqlStr + " 	,m.makerid "
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	,m.makerid, m.itemid, m.itemoption, m.lastvatinclude "
		end if

		''response.write "테스트중<br><br>"
		''response.write sqlStr
		''dbget.close()
		''response.end
''rw sqlStr
'response.end

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
				set FItemList(i) = new CMonthlyMaeipLedgeItem

                FItemList(i).FisJungsan         = false
				FItemList(i).Fyyyymm     		= rsget("yyyymm")
				FItemList(i).FstockPlace    	= rsget("stockPlace")
				''FItemList(i).FtargetGbn     	= rsget("targetGbn")
				FItemList(i).Fitemgubun     	= rsget("itemgubun")
                FItemList(i).Flastmwdiv         = rsget("lastmwdiv")

				if (FRectShowShopid <> "") then
					FItemList(i).Fshopid         = rsget("shopid")
				end if

				FItemList(i).FprevSysStockNo    = rsget("prevSysStockNo")
				FItemList(i).FprevSysStockSum   = rsget("prevSysStockSum")

				FItemList(i).FIpgoNo     		= rsget("IpgoNo")
				FItemList(i).FIpgoSum     		= rsget("IpgoSum")
				FItemList(i).FMoveNo     		= rsget("MoveNo")
				FItemList(i).FMoveSum     		= rsget("MoveSum")
				FItemList(i).FSellNo     		= rsget("SellNo")
				FItemList(i).FSellSum     		= rsget("SellSum")
				FItemList(i).FOffChulNo     	= rsget("OffChulNo")
				FItemList(i).FOffChulSum     	= rsget("OffChulSum")
				FItemList(i).FEtcChulNo     	= rsget("EtcChulNo")
				FItemList(i).FEtcChulSum     	= rsget("EtcChulSum")
				FItemList(i).FCsNo     			= rsget("CsNo")
				FItemList(i).FCsSum     		= rsget("CsSum")
				FItemList(i).FLossChulNo     	= rsget("LossChulNo")
				FItemList(i).FLossChulSum     	= rsget("LossChulSum")

				FItemList(i).FcurSysStockNo     = rsget("curSysStockNo")
				FItemList(i).FcurSysStockSum    = rsget("curSysStockSum")

				FItemList(i).FcurErrRealCheckNo = rsget("curErrRealCheckNo")
				FItemList(i).FcurErrRealCheckSum	= rsget("curErrRealCheckSum")

                IF (FRectSubGrpType="makerid") then
                    FItemList(i).FMakerid = rsget("makerid")
				elseif (FRectSubGrpType = "itemid") then
					FItemList(i).FMakerid 		= rsget("makerid")
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).Fitemoption    	= rsget("itemoption")
					FItemList(i).Flastvatinclude    = rsget("lastvatinclude")
                end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	' /admin/newreport/monthlystock_overValue_excel.asp
	public Sub GetJeagoOverValueListNotPaging()
		dim sqlStr,i

		if (frectIsUsingV2 = "Y") then
			sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthlystock_overValue_MakeEXL_List_V2] '"&frectreqYYYYMM&"','"&frectreqStrplace&"',"&FCurrPage&","&FPageSize&""
		else
			sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthlystock_overValue_MakeEXL_List] '"&frectreqYYYYMM&"','"&frectreqStrplace&"',"&FCurrPage&","&FPageSize&""
		end if

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		dbget.CommandTimeout = 60*5   ' 5분
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/06

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		i=0
		if  not rsget.EOF  then
			fArrLIst = rsget.getrows()
		end if

		rsget.Close
	end sub

	'// 재고평가충당금 : 물류
    public Sub GetJeagoOverValueSum()
    	dim sqlStr
		dim colName, valPrice, valPriceFieldName

		valPriceFieldName = " s.lastbuyPrice "
		if (FRectPriceGubun = "V") then
			valPriceFieldName = " IsNull(s.avgipgoPrice, 0) "
		end if

		colName = "s.curSysStockNo"

		if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(s.lastvatinclude, 'Y') = 'Y' then (" + CStr(valPriceFieldName) + "*10/11) else " + CStr(valPriceFieldName) + " end)"
			if (FRectPriceGubun = "V") then
				valPrice = "(case when IsNull(s.lastvatinclude, 'Y') = 'Y' then Round((" + CStr(valPriceFieldName) + "*10/11), 0) else " + CStr(valPriceFieldName) + " end)"
			end if
		else
			valPrice = valPriceFieldName
		end if

		sqlStr = " SELECT "
		sqlStr = sqlStr & " 	s.targetGbn, s.itemgubun, IsNULL(s.lastmwdiv,'Z') as mwdiv, "
		sqlStr = sqlStr & " 	SUM("+CStr(colName)+") as totStockNo,"
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun1, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun2, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun3, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun4, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun5, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN IsNull(s.lastIpgoDate, '') = '' "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun6, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun7, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun8, "

		'// 연도별
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun11, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun12, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun13, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun14, "

		sqlStr = sqlStr & " 	SUM(" + CStr(colName) + " * " + CStr(valPrice) + ") as MonthGubunSUM "
		sqlStr = sqlStr & " FROM db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail_V2 s with (nolock)"

		sqlStr = sqlStr & " WHERE "
		sqlStr = sqlStr & " 	1 = 1 "
		sqlStr = sqlStr & " 	and s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		sqlStr = sqlStr & "		and not (s.itemgubun='10' and s.itemid  in (0,11406,6400))"
		sqlStr = sqlStr & "		and s.stockPlace = 'L' "

		if (FRectMwDiv <> "") then
			sqlStr = sqlStr & " 	and IsNULL(s.lastmwdiv,'Z') = '" + CStr(FRectMwDiv) + "' "
		end if

		if (FRectItemGubun <> "") then
			sqlStr = sqlStr & " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end if

        if (FRectVatYn<>"") then
		    sqlStr = sqlStr + " and IsNull(s.lastvatinclude, 'Y')='" + FRectVatYn + "'"
		end if

        if (FRectTargetGbn<>"") then
			if (FRectTargetGbn = "3X") then
				sqlStr = sqlStr + " and s.targetGbn not in ('ET', 'EG') "
			else
				sqlStr = sqlStr + " and s.targetGbn='" + FRectTargetGbn + "'"
			end if
		end if

		sqlStr = sqlStr & " 	and " + CStr(colName) + " <> 0 "
		sqlStr = sqlStr & " GROUP BY "
		sqlStr = sqlStr & " 	s.targetGbn, s.itemgubun, s.lastMWDiv "
		'sqlStr = sqlStr & " ORDER BY s.targetGbn DESC, s.itemgubun asc, IsNULL(s.lastmwdiv,'Z') asc"
		sqlStr = sqlStr & " ORDER BY s.itemgubun asc, IsNULL(s.lastmwdiv,'Z') asc"

		''rw sqlStr
		''response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CStockOverValueSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).FTotBuySum1 	= rsget("MonthGubun1")
				FItemList(i).FTotBuySum2 	= rsget("MonthGubun2")
				FItemList(i).FTotBuySum3 	= rsget("MonthGubun3")
				FItemList(i).FTotBuySum4 	= rsget("MonthGubun4")
				FItemList(i).FTotBuySum5 	= rsget("MonthGubun5")
				FItemList(i).FTotBuySum6 	= rsget("MonthGubun6")
				FItemList(i).FTotBuySum7 	= rsget("MonthGubun7")
				FItemList(i).FTotBuySum8 	= rsget("MonthGubun8")

				FItemList(i).FTotBuySum11 	= rsget("MonthGubun11")
				FItemList(i).FTotBuySum12 	= rsget("MonthGubun12")
				FItemList(i).FTotBuySum13 	= rsget("MonthGubun13")
				FItemList(i).FTotBuySum14 	= rsget("MonthGubun14")

				FItemList(i).FTotBuySum 	= rsget("MonthGubunSUM")
                FItemList(i).FtotStockNo    = rsget("totStockNo")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

	'// 재고평가충당금 : 물류 : 상세
    public Sub GetJeagoOverValueDetailSum()
    	dim sqlStr, sqlAdd
		dim colName, valPrice, valPriceFieldName
        Dim isPagingReq : isPagingReq=False

		valPriceFieldName = " s.lastbuyPrice "
		if (FRectPriceGubun = "V") then
			valPriceFieldName = " IsNull(s.avgipgoPrice, 0) "
		end if

        if (FRectMakerid <> "")or(FRectShowItemList<>"") then
            isPagingReq = TRUE
        end if

		colName = "s.curSysStockNo"

		if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(s.lastvatinclude, 'Y') = 'Y' then (" + CStr(valPriceFieldName) + "*10/11) else " + CStr(valPriceFieldName) + " end)"
			if (FRectPriceGubun = "V") then
				valPrice = "(case when IsNull(s.lastvatinclude, 'Y') = 'Y' then Round((" + CStr(valPriceFieldName) + "*10/11), 0) else " + CStr(valPriceFieldName) + " end)"
			end if
		else
			valPrice = valPriceFieldName
		end if

        sqlAdd = ""

        if (FRectMonthGubun = "1") then
			'// 1-3개월
			sqlAdd = sqlAdd & " 	and DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 "
		elseif (FRectMonthGubun = "2") then
			'// 4-6개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) "
		elseif (FRectMonthGubun = "3") then
			'// 7-12개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) "
		elseif (FRectMonthGubun = "4") then
			'// 13-24개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		elseif (FRectMonthGubun = "5") then
			'// 25개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) "
		elseif (FRectMonthGubun = "6") then
			'// 미지정
			sqlAdd = sqlAdd & " 	and IsNull(s.lastIpgoDate, '') = '' "
		elseif (FRectMonthGubun = "7") then
			'// 13-18개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) "
		elseif (FRectMonthGubun = "8") then
			'// 19-24개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		elseif (FRectMonthGubun = "11") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) "
		elseif (FRectMonthGubun = "12") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) "
		elseif (FRectMonthGubun = "13") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) "
		elseif (FRectMonthGubun = "14") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) "
		end if

		if (FRectMwDiv <> "") then
			sqlAdd = sqlAdd & " 	and IsNULL(s.lastmwdiv,'Z') = '" + CStr(FRectMwDiv) + "' "
		end if

		if (FRectItemGubun <> "") then
			sqlAdd = sqlAdd & " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end if

        if (FRectVatYn<>"") then
		    sqlAdd = sqlAdd + " and IsNull(s.lastvatinclude, 'Y')='" + FRectVatYn + "'"
		end if

        if (FRectTargetGbn<>"") then
			if (FRectTargetGbn = "3X") then
				sqlAdd = sqlAdd + " and s.targetGbn not in ('ET', 'EG') "
			else
				sqlAdd = sqlAdd + " and s.targetGbn='" + FRectTargetGbn + "'"
			end if
		end if

        if (FRectMakerid<>"") then
		    sqlAdd = sqlAdd + " and s.makerid='" + FRectMakerid + "'"
		end if


        if (isPagingReq) then
            sqlStr = " SELECT count(*) as CNT, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
            sqlStr = sqlStr & " FROM "
    		sqlStr = sqlStr & " 	db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail_V2 s "
            sqlStr = sqlStr & " WHERE "
    		sqlStr = sqlStr & " 	1 = 1 "
    		sqlStr = sqlStr & " 	and s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
    		sqlStr = sqlStr & "		and not (s.itemgubun='10' and s.itemid  in (0,11406,6400))"
    		sqlStr = sqlStr & "		and s.stockPlace = 'L' "

    		sqlStr = sqlStr & sqlAdd
            sqlStr = sqlStr & " 	and " + CStr(colName) + " <> 0 "

    		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
    		rsget.Close

    		'지정페이지가 전체 페이지보다 클 때 함수종료
    		if Cint(FCurrPage)>Cint(FTotalPage) then
    			FResultCount = 0
    			exit sub
    		end if
        end if

		sqlStr = " SELECT top "&FPageSize*FCurrPage&" "
		sqlStr = sqlStr & " 	s.targetGbn, s.itemgubun, IsNULL(s.lastmwdiv,'Z') as mwdiv, s.makerid, '' as purchasetypeStr, "
		sqlStr = sqlStr & " 	SUM(" + CStr(colName) + ") as totStockNo, "

		if (FRectMakerid <> "") or (FRectShowItemList<>"") then
			sqlStr = sqlStr & " s.itemgubun,s.itemid,s.itemoption, isNULL(i.itemname,si.shopitemname) as itemname, IsNULL(o.optionname,'') as itemoptionname, IsNull(s.lastIpgoDate, '') as lastIpgoDate, " + CStr(valPrice) + " as buyPrice, "
		else
			sqlStr = sqlStr & " '" + CStr(FRectMonthGubun) + "' as lastIpgoDate, "
		end if

		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun1, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun2, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun3, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun4, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun5, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN IsNull(s.lastIpgoDate, '') = '' "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun6, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun7, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun8, "

		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun11, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun12, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun13, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun14, "

		sqlStr = sqlStr & " 	SUM(" + CStr(colName) + " * " + CStr(valPrice) + ") as MonthGubunSUM "

		sqlStr = sqlStr & " FROM "
		sqlStr = sqlStr & " 	db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail_V2 s "

        if (FRectVatYn<>"") or (FRectMakerid <> "") or (FRectShowItemList<>"") or (FRectShopSuplyPrice = "Y") then
            sqlStr = sqlStr & "		left join [db_item].[dbo].tbl_item i "
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun='10' "
            sqlStr = sqlStr & "		and s.itemid=i.itemid "

			sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on s.yyyymm='"&FRectYYYYMM&"' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"

            sqlStr = sqlStr & "		left join [db_shop].[dbo].tbl_shop_item si "
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun<>'10' "
            sqlStr = sqlStr & "		and s.itemgubun=si.itemgubun"
            sqlStr = sqlStr & "		and s.itemid=si.shopitemid "
            sqlStr = sqlStr & "		and s.itemoption=si.itemoption "
        end if

		sqlStr = sqlStr & " WHERE "
		sqlStr = sqlStr & " 	1 = 1 "
		sqlStr = sqlStr & " 	and s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		sqlStr = sqlStr & "		and not (s.itemgubun='10' and s.itemid  in (0,11406,6400))"
		sqlStr = sqlStr & "		and s.stockPlace = 'L' "

		sqlStr = sqlStr & sqlAdd

		sqlStr = sqlStr & " 	and " + CStr(colName) + " <> 0 "
		sqlStr = sqlStr & " GROUP BY "
		sqlStr = sqlStr & " 	s.targetGbn, s.itemgubun, s.lastMWDiv, s.makerid "

		if (FRectMakerid <> "")or(FRectShowItemList<>"") then
			sqlStr = sqlStr & " ,s.itemgubun,s.itemid,s.itemoption, isNULL(i.itemname,si.shopitemname), IsNULL(o.optionname,''), IsNull(s.lastIpgoDate, ''), " + CStr(valPrice) + " "
		end if

        if ((FRectMakerid <> "")or(FRectShowItemList<>"")) then
            if (FRectOrdTp="S") then
       		    sqlStr = sqlStr & " ORDER BY "
        		sqlStr = sqlStr & " 	totStockNo desc, s.targetGbn DESC, s.itemgubun, IsNULL(s.lastmwdiv,'Z'),  s.makerid "
        		sqlStr = sqlStr & " ,s.itemid,s.itemoption "
            else
        		sqlStr = sqlStr & " ORDER BY "
        		sqlStr = sqlStr & " 	s.itemgubun "
        		sqlStr = sqlStr & " ,s.itemid,s.itemoption "
            end if
        else
            sqlStr = sqlStr & " ORDER BY "
        	sqlStr = sqlStr & " 	totStockNo desc, s.targetGbn DESC, s.itemgubun"
        end if
		''response.write sqlStr
		''response.end

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
				set FItemList(i) = new CStockOverValueSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).FpurchasetypeStr     = rsget("purchasetypeStr")



				FItemList(i).Fmakerid     	= rsget("makerid")
				FItemList(i).FlastIpgoDate  = rsget("lastIpgoDate")
				FItemList(i).FtotStockNo    = rsget("totStockNo")

				if (FRectMakerid <> "")or(FRectShowItemList<>"") then
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname     	= rsget("itemname")
					FItemList(i).Fitemoptionname = rsget("itemoptionname")
					FItemList(i).FbuyPrice 		= rsget("buyPrice")
				end if

				FItemList(i).FTotBuySum1 	= rsget("MonthGubun1")
				FItemList(i).FTotBuySum2 	= rsget("MonthGubun2")
				FItemList(i).FTotBuySum3 	= rsget("MonthGubun3")
				FItemList(i).FTotBuySum4 	= rsget("MonthGubun4")
				FItemList(i).FTotBuySum5 	= rsget("MonthGubun5")
				FItemList(i).FTotBuySum6 	= rsget("MonthGubun6")
				FItemList(i).FTotBuySum7 	= rsget("MonthGubun7")
				FItemList(i).FTotBuySum8 	= rsget("MonthGubun8")

				FItemList(i).FTotBuySum11 	= rsget("MonthGubun11")
				FItemList(i).FTotBuySum12 	= rsget("MonthGubun12")
				FItemList(i).FTotBuySum13 	= rsget("MonthGubun13")
				FItemList(i).FTotBuySum14 	= rsget("MonthGubun14")

				FItemList(i).FTotBuySum 	= rsget("MonthGubunSUM")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

	'// 재고평가충당금 : 매장
    public Sub GetJeagoOverValueSum_Shop()
    	dim sqlStr
		dim colName, valPrice, colNameIpgoDate, valPriceFieldName

		'valPriceFieldName = " s.LstBuyCash "
		valPriceFieldName = " s.lastbuyprice "
		if (FRectPriceGubun = "V") then
			valPriceFieldName = " IsNull(s.avgipgoPrice, 0) "
		end if

		''G02799 : 아이띵소 그룹코드
		''B011 : 위탁판매, B012 : 업체위탁, B013 : 출고위탁, B021 : 오프매입, B022 : 매장매입, B023 : 가맹점매입, B031 : 출고매입, B032 : 센터매입

		colName = "s.curSysStockNo"

        if (FRectLastIpgoGBN="S") then
            colNameIpgoDate = "s.lastIpgoDate"
		elseif (FRectLastIpgoGBN="M") then
			colNameIpgoDate = "(case when IsNull(s1.LstCenterMwDiv, 'W') = 'M' then s.lastipgoDateByMW else s1.lastIpgoDate end)"
        else
            colNameIpgoDate = "s1.lastIpgodateLogics"
        end if

		if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(s.lastvatinclude, 'Y') = 'Y' then (" + CStr(valPriceFieldName) + "*10/11) else " + CStr(valPriceFieldName) + " end)"
			if (FRectPriceGubun = "V") then
				valPrice = "(case when IsNull(s.lastvatinclude, 'Y') = 'Y' then Round((" + CStr(valPriceFieldName) + "*10/11), 0) else " + CStr(valPriceFieldName) + " end)"
			end if
		else
			valPrice = valPriceFieldName
		end if

		sqlStr = " SELECT "
        sqlStr = sqlStr & " 	isNULL(s.targetGbn,'OF') as targetGbn, "
		sqlStr = sqlStr & " 	s.itemgubun "
		sqlStr = sqlStr & " 	, IsNULL(s.lastmwdiv,'Z') as mwdiv "
		sqlStr = sqlStr & " 	, SUM("+CStr(colName)+") as totStockNo"
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun1 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun2 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun3 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun4 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun5 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN IsNull("&colNameIpgoDate&", '') = '' THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun6 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun7 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun8 "

		'// 연도별
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE WHEN DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") ELSE 0 END), 0) AS MonthGubun11 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE WHEN DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") ELSE 0 END), 0) AS MonthGubun12 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE WHEN DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") ELSE 0 END), 0) AS MonthGubun13 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE WHEN DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") ELSE 0 END), 0) AS MonthGubun14 "

		sqlStr = sqlStr & " 	,IsNull(SUM(" + CStr(colName) + " * " + CStr(valPrice) + "), 0) AS MonthGubunSUM "
		sqlStr = sqlStr & " FROM db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail_V2 s with (nolock)"

		if (FRectLastIpgoGBN <> "S") then
			'// 물류입고일은 누적재고 테이블에서 가져온다.
			sqlStr = sqlStr & " join db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s1 with (nolock)"
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		1 = 1 "
			sqlStr = sqlStr & " 		and s.yyyymm = s1.yyyymm "
			sqlStr = sqlStr & " 		and s.shopid = s1.shopid "
			sqlStr = sqlStr & " 		and s.itemgubun = s1.itemgubun "
			sqlStr = sqlStr & " 		and s.itemid = s1.itemid "
			sqlStr = sqlStr & " 		and s.itemoption = s1.itemoption "
		end if

		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " 	AND s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		sqlStr = sqlStr & " 	AND NOT (s.itemgubun = '10' AND s.itemid IN (0,11406,6400)) "
		sqlStr = sqlStr & " 	and s.itemid <> 0 "
		sqlStr = sqlStr & " 	AND " + CStr(colName) + " <> 0 "
		sqlStr = sqlStr & "		and s.stockPlace = 'S' "

		if (FRectMwDiv <> "") then
			if (FRectMwDiv = "M") then
				sqlStr = sqlStr & " 	and IsNULL(s.lastmwdiv,'Z') in ('B021','B022','B031','B032','B013') "
				sqlStr = sqlStr & " 	and (isNULL(s.targetGbn,'OF') not in ('ET','EG'))" ''NOT ((s.targetGbn='ET') and (IsNULL(s.lastmwdiv,'Z')='B013'))"
			elseif (FRectMwDiv = "W") then
				sqlStr = sqlStr & " 	and IsNULL(s.lastmwdiv,'Z') in ('B011','B012') "
			elseif (FRectMwDiv = "Z") then
				sqlStr = sqlStr & " 	and IsNULL(s.lastmwdiv,'Z') not in ('B021','B022','B031','B032','B011','B012','B013') "
			else
				sqlStr = sqlStr & " 	and IsNULL(s.lastmwdiv,'Z') = '" + CStr(FRectMwDiv) + "' "
			end if
		end if

		if (FRectItemGubun <> "") then
			sqlStr = sqlStr & " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end if

        if (FRectVatYn<>"") then
		    sqlStr = sqlStr + " and IsNull(s.lastvatinclude, 'Y')='" + FRectVatYn + "'"
		end if

        if (FRectTargetGbn<>"") then
			if (FRectTargetGbn = "3X") then
				sqlStr = sqlStr + " and isNULL(s.targetGbn,'OF') not in ('ET', 'EG') "
			else
				sqlStr = sqlStr + " and isNULL(s.targetGbn,'OF')='" + FRectTargetGbn + "'"
			end if
		end if

		if (FRectShopID <> "") then
			sqlStr = sqlStr & " 	and s.shopid = '" + CStr(FRectShopID) + "' "
		end if

        IF (FRectetcjungsantype<>"") then
    	    if (FRectetcjungsantype="41") then
    	        sqlStr = sqlStr + " and s.etcjungsantype in ('4','1')"
    	    else
    	        sqlStr = sqlStr + " and s.etcjungsantype='"&FRectetcjungsantype&"'"
    	    end if
    	end if

		sqlStr = sqlStr & " GROUP BY "
        sqlStr = sqlStr & " isNULL(s.targetGbn,'OF'), "
		sqlStr = sqlStr & " 	s.itemgubun "
		sqlStr = sqlStr & " 	, IsNULL(s.lastmwdiv,'Z') "
		'sqlStr = sqlStr & " ORDER BY targetGbn DESC, s.itemgubun asc, IsNULL(s.lastmwdiv,'Z') asc"
		sqlStr = sqlStr & " ORDER BY s.itemgubun asc, IsNULL(s.lastmwdiv,'Z') asc"

		''response.write sqlStr
		''response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CStockOverValueSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).FTotBuySum1 	= rsget("MonthGubun1")
				FItemList(i).FTotBuySum2 	= rsget("MonthGubun2")
				FItemList(i).FTotBuySum3 	= rsget("MonthGubun3")
				FItemList(i).FTotBuySum4 	= rsget("MonthGubun4")
				FItemList(i).FTotBuySum5 	= rsget("MonthGubun5")
				FItemList(i).FTotBuySum6 	= rsget("MonthGubun6")
				FItemList(i).FTotBuySum7 	= rsget("MonthGubun7")
				FItemList(i).FTotBuySum8 	= rsget("MonthGubun8")

				FItemList(i).FTotBuySum11 	= rsget("MonthGubun11")
				FItemList(i).FTotBuySum12 	= rsget("MonthGubun12")
				FItemList(i).FTotBuySum13 	= rsget("MonthGubun13")
				FItemList(i).FTotBuySum14 	= rsget("MonthGubun14")

				FItemList(i).FTotBuySum 	= rsget("MonthGubunSUM")
                FItemList(i).FtotStockNo    = rsget("totStockNo")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

	'// 재고평가충당금 : 매장 : 상세
    public Sub GetJeagoOverValueDetailSum_Shop()
    	dim sqlStr, sqlAdd
		dim colName, valPrice, colNameIpgoDate, valPriceFieldName
        Dim isPagingReq : isPagingReq=False

		valPriceFieldName = " s.LstBuyCash "
		if (FRectPriceGubun = "V") then
			valPriceFieldName = " IsNull(s.avgipgoPrice, 0) "
		end if

        if (FRectMakerid <> "")or(FRectShowItemList<>"") then
            isPagingReq = TRUE
        end if

		''G02799 : 아이띵소 그룹코드
		''B011 : 위탁판매, B012 : 업체위탁, B013 : 출고위탁, B021 : 오프매입, B022 : 매장매입, B023 : 가맹점매입, B031 : 출고매입, B032 : 센터매입

		colName = "s.curSysStockNo"

        if (FRectLastIpgoGBN="S") then
            colNameIpgoDate = "s1.lastIpgoDate"
		elseif (FRectLastIpgoGBN="M") then
			colNameIpgoDate = "(case when IsNull(s1.LstCenterMwDiv, 'W') = 'M' then s.lastipgoDateByMW else s1.lastIpgoDate end)"
        else
            colNameIpgoDate = "s1.lastIpgodateLogics"
        end if

		if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(s.lastvatinclude, 'Y') = 'Y' then (" + CStr(valPriceFieldName) + "*10/11) else " + CStr(valPriceFieldName) + " end)"
			if (FRectPriceGubun = "V") then
				valPrice = "(case when IsNull(s.lastvatinclude, 'Y') = 'Y' then Round((" + CStr(valPriceFieldName) + "*10/11), 0) else " + CStr(valPriceFieldName) + " end)"
			end if
		else
			valPrice = valPriceFieldName
		end if

        sqlAdd = ""
        if (FRectMwDiv <> "") then
			if (FRectMwDiv = "M") then
				sqlAdd = sqlAdd & " 	and IsNULL(s.lastmwdiv,'Z') in ('B021','B022','B031','B032','B013') "
				sqlAdd = sqlAdd & " 	and (isNULL(s.targetGbn,'OF') not in ('ET','EG')) " ''NOT ((s.targetGbn='ET') and (IsNULL(s.lastmwdiv,'Z')='B013'))" ''
			elseif (FRectMwDiv = "W") then
				sqlAdd = sqlAdd & " 	and IsNULL(s.lastmwdiv,'Z') in ('B011','B012') "
			elseif (FRectMwDiv = "Z") then
				sqlAdd = sqlAdd & " 	and IsNULL(s.lastmwdiv,'Z') not in ('B021','B022','B031','B032','B011','B012','B013') "
			else
				sqlAdd = sqlAdd & " 	and IsNULL(s.lastmwdiv,'Z') = '" + CStr(FRectMwDiv) + "' "
			end if
		end if

		if (FRectItemGubun <> "") then
			sqlAdd = sqlAdd & " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end if

        if (FRectVatYn<>"") then
		    sqlAdd = sqlAdd + " and IsNull(s.lastvatinclude, 'Y')='" + FRectVatYn + "'"
		end if

        if (FRectTargetGbn<>"") then
			if (FRectTargetGbn = "3X") then
				sqlAdd = sqlAdd + " and isNULL(s.targetGbn,'OF') not in ('ET', 'EG') "
			else
				sqlAdd = sqlAdd + " and isNULL(s.targetGbn,'OF')='" + FRectTargetGbn + "'"
			end if
		end if

		if (FRectShopID <> "") then
			sqlAdd = sqlAdd & " 	and s.shopid = '" + CStr(FRectShopID) + "' "
		end if

		if (FRectMonthGubun = "1") then
			'// 1-3개월
			sqlAdd = sqlAdd & " 	and DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 "
		elseif (FRectMonthGubun = "2") then
			'// 4-6개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) "
		elseif (FRectMonthGubun = "3") then
			'// 7-12개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) "
		elseif (FRectMonthGubun = "4") then
			'// 13-24개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		elseif (FRectMonthGubun = "5") then
			'// 25개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) "
		elseif (FRectMonthGubun = "6") then
			'// 미지정
			sqlAdd = sqlAdd & " 	and IsNull("&colNameIpgoDate&", '') = '' "
		elseif (FRectMonthGubun = "7") then
			'// 13-18개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) "
		elseif (FRectMonthGubun = "8") then
			'// 19-24개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		elseif (FRectMonthGubun = "11") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) "
		elseif (FRectMonthGubun = "12") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) "
		elseif (FRectMonthGubun = "13") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) "
		elseif (FRectMonthGubun = "14") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) "
		end if

        if (FRectMakerid<>"") then
		    sqlAdd = sqlAdd + " and s.Makerid='" + FRectMakerid + "'"
		end if

        IF (FRectetcjungsantype<>"") then
    	    if (FRectetcjungsantype="41") then
    	        sqlAdd = sqlAdd + " and s.etcjungsantype in ('4','1')"
    	    else
    	        sqlAdd = sqlAdd + " and s.etcjungsantype='"&FRectetcjungsantype&"'"
    	    end if
    	end if

        if (isPagingReq) then
            sqlStr = " SELECT count(*) as CNT, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
            sqlStr = sqlStr & " FROM "
            sqlStr = sqlStr & " 	db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail_V2 s "

			''if (FRectLastIpgoGBN <> "S") then
				sqlStr = sqlStr & " 	join db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s1 "
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		1 = 1 "
				sqlStr = sqlStr & " 		and s.yyyymm = s1.yyyymm "
				sqlStr = sqlStr & " 		and s.shopid = s1.shopid "
				sqlStr = sqlStr & " 		and s.itemgubun = s1.itemgubun "
				sqlStr = sqlStr & " 		and s.itemid = s1.itemid "
				sqlStr = sqlStr & " 		and s.itemoption = s1.itemoption "
			''end if

    		sqlStr = sqlStr & " WHERE 1 = 1 "
    		sqlStr = sqlStr & " 	AND s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
    		sqlStr = sqlStr & " 	AND NOT (s.itemgubun = '10' AND s.itemid IN (0,11406,6400)) "
    		sqlStr = sqlStr & " 	AND s.itemid <> 0 "
    		sqlStr = sqlStr & " 	AND " + CStr(colName) + " <> 0 "
			sqlStr = sqlStr & "		and s.stockPlace = 'S' "
    		sqlStr = sqlStr & sqlAdd
            rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
    		rsget.Close

    		'지정페이지가 전체 페이지보다 클 때 함수종료
    		if Cint(FCurrPage)>Cint(FTotalPage) then
    			FResultCount = 0
    			exit sub
    		end if
        end if

		sqlStr = " SELECT top "&FPageSize*FCurrPage&" "
		sqlStr = sqlStr & " 	isNULL(s.targetGbn,'OF') as targetGbn "
		sqlStr = sqlStr & " 	,s.itemgubun "
		sqlStr = sqlStr & " 	,IsNULL(s.lastmwdiv,'Z') AS mwdiv "
		sqlStr = sqlStr & " 	,s.makerid "
		sqlStr = sqlStr & " 	,SUM(" + CStr(colName) + ") as totStockNo "

		if (FRectMakerid <> "")or(FRectShowItemList<>"") then
			sqlStr = sqlStr & " ,s.shopid,s.itemgubun,s.itemid,s.itemoption, isNULL(i.itemname,si.shopitemname) as itemname, IsNULL(o.optionname,'') as itemoptionname, IsNull("&colNameIpgoDate&", '') as lastIpgoDate, " + CStr(valPrice) + " as buyPrice, IsNULL(s.lastmwdiv,'Z') as LstComm_cd "
		else
			sqlStr = sqlStr & " ,'" + CStr(FRectMonthGubun) + "' as lastIpgoDate "
		end if

		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun1 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun2 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun3 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun4 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun5 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun6 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun7 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun8 "

		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun11 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun12 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun13 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun14 "

		sqlStr = sqlStr & " 	,IsNull(SUM(" + CStr(colName) + " * " + CStr(valPrice) + "), 0) AS MonthGubunSUM "
		sqlStr = sqlStr & " FROM "
		sqlStr = sqlStr & " 	db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail_V2 s "

        if (FRectVatYn<>"") or (FRectMakerid <> "") or (FRectShowItemList<>"") or (FRectShopSuplyPrice = "Y") then
            sqlStr = sqlStr & "		left join [db_item].[dbo].tbl_item i "
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun='10' "
            sqlStr = sqlStr & "		and s.itemid=i.itemid "

			sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on s.yyyymm='"&FRectYYYYMM&"' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"

            sqlStr = sqlStr & "		left join [db_shop].[dbo].tbl_shop_item si "
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun<>'10' "
            sqlStr = sqlStr & "		and s.itemgubun=si.itemgubun"
            sqlStr = sqlStr & "		and s.itemid=si.shopitemid "
            sqlStr = sqlStr & "		and s.itemoption=si.itemoption "
        end if

		''if (FRectLastIpgoGBN <> "S") then
			'// 물류입고일은 누적재고 테이블에서 가져온다.
			sqlStr = sqlStr & " 	join db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s1 "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		1 = 1 "
			sqlStr = sqlStr & " 		and s.yyyymm = s1.yyyymm "
			sqlStr = sqlStr & " 		and s.shopid = s1.shopid "
			sqlStr = sqlStr & " 		and s.itemgubun = s1.itemgubun "
			sqlStr = sqlStr & " 		and s.itemid = s1.itemid "
			sqlStr = sqlStr & " 		and s.itemoption = s1.itemoption "
		''end if

		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " 	AND s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		sqlStr = sqlStr & " 	AND NOT (s.itemgubun = '10' AND s.itemid IN (0,11406,6400)) "
		sqlStr = sqlStr & " 	AND s.itemid <> 0 "
		sqlStr = sqlStr & " 	AND " + CStr(colName) + " <> 0 "
		sqlStr = sqlStr & "		and s.stockPlace = 'S' "

		sqlStr = sqlStr & sqlAdd

		sqlStr = sqlStr & " GROUP BY isNULL(s.targetGbn,'OF') "
		sqlStr = sqlStr & " 	,s.itemgubun "
		sqlStr = sqlStr & " 	,IsNULL(s.lastmwdiv,'Z') "
		sqlStr = sqlStr & " 	,s.Makerid "

		if (FRectMakerid <> "")or(FRectShowItemList<>"") then
			sqlStr = sqlStr & " ,s.shopid,s.itemgubun,s.itemid,s.itemoption, isNULL(i.itemname,si.shopitemname), IsNULL(o.optionname,''), IsNull("&colNameIpgoDate&", ''), " + CStr(valPrice) + ", IsNULL(s.lastmwdiv,'Z') "
		end if

        if ((FRectMakerid <> "")or(FRectShowItemList<>"")) then
            if (FRectOrdTp="S") then
                sqlStr = sqlStr & " ORDER BY totStockNo desc, targetGbn DESC "
        		sqlStr = sqlStr & " 	,s.itemgubun "
        		sqlStr = sqlStr & " ,s.shopid,s.itemid,s.itemoption "
            else
        		sqlStr = sqlStr & " ORDER BY s.shopid,s.itemgubun "
        		sqlStr = sqlStr & " ,s.itemid,s.itemoption "
        	end if
        else
            sqlStr = sqlStr & " ORDER BY targetGbn DESC "
    		sqlStr = sqlStr & " 	,s.itemgubun "
    		sqlStr = sqlStr & " 	,IsNULL(s.lastmwdiv,'Z') "
    		sqlStr = sqlStr & " 	,IsNull(SUM(" + CStr(colName) + " * " + CStr(valPrice) + "), 0) desc, s.Makerid "
        end if

		''response.write sqlStr
		''response.end
        rsget.pagesize = FPageSize
       rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
'   if (FResultCount>=3000) then  ''서동석 임시
'        rw sqlStr
'        dbget.close() : response.end
'   end if
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CStockOverValueSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).Fmakerid     	= rsget("makerid")
				FItemList(i).FlastIpgoDate  = rsget("lastIpgoDate")
				FItemList(i).FtotStockNo    = rsget("totStockNo")

				if (FRectMakerid <> "")or(FRectShowItemList<>"") then
					FItemList(i).Fshopid     	= rsget("shopid")
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname     	= rsget("itemname")
					FItemList(i).Fitemoptionname = rsget("itemoptionname")
					FItemList(i).FbuyPrice 		= rsget("buyPrice")
					FItemList(i).FLstComm_cd	= rsget("LstComm_cd")
				end if

				FItemList(i).FTotBuySum1 	= rsget("MonthGubun1")
				FItemList(i).FTotBuySum2 	= rsget("MonthGubun2")
				FItemList(i).FTotBuySum3 	= rsget("MonthGubun3")
				FItemList(i).FTotBuySum4 	= rsget("MonthGubun4")
				FItemList(i).FTotBuySum5 	= rsget("MonthGubun5")
				FItemList(i).FTotBuySum6 	= rsget("MonthGubun6")
				FItemList(i).FTotBuySum7 	= rsget("MonthGubun7")
				FItemList(i).FTotBuySum8 	= rsget("MonthGubun8")

				FItemList(i).FTotBuySum11 	= rsget("MonthGubun11")
				FItemList(i).FTotBuySum12 	= rsget("MonthGubun12")
				FItemList(i).FTotBuySum13 	= rsget("MonthGubun13")
				FItemList(i).FTotBuySum14 	= rsget("MonthGubun14")

				FItemList(i).FTotBuySum 	= rsget("MonthGubunSUM")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)

		FCurrPage = 1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end class

Sub drawSelectBoxAccShop(yyyymm, makerid, selectBoxName, selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value="" <%if selectedId="" then response.write " selected"%>>ALL</option><%
   query1 = " select distinct shopid from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary "
   query1 = query1 + " where yyyymm = '" + CStr(yyyymm) + "' "

   if (makerid <> "") then
	   query1 = query1 + " and IsNull(LstMakerid, 'Z') = '" + CStr(makerid) + "' "
   end if

   query1 = query1 + " order by shopid "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("shopid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("shopid")&"' "&tmp_str&">" + rsget("shopid") + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxBuseoGubunWith3PL(selectBoxName, selectedId)
%>
	<select class="select" name="<%=selectBoxName%>" class="select">
		<option value="">선택</option>
		<option value="3X" <% if (selectedId="3X") then response.write "selected" %> >텐바이텐(3PL제외)</option>
		<option value="ON" <% if (selectedId="ON") then response.write "selected" %> >온라인</option>
		<option value="OF" <% if (selectedId="OF") then response.write "selected" %> >오프라인</option>
		<option value="IT" <% if (selectedId="IT") then response.write "selected" %> >아이띵소(구)</option>
		<option value="ET" <% if (selectedId="ET") then response.write "selected" %> >3PL(아이띵소)</option>
		<option value="EG" <% if (selectedId="EG") then response.write "selected" %> >3PL(유그레잇)</option>
	</select>
<%
End Sub

%>
