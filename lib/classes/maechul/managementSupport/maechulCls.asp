<%
'#######################################################
' Description : 매출클래스 파일
' History	:  서동석 생성
'              2022.09.19 한용민 수정(보안 취약부분 수정, 쿼리 클래스로 분리)
'#######################################################

class cManagementSupportMaechul_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fonoff							'온, 오프 구분
	public fitemdiv							'상품구분(ON/OF/IT/AC)
	public fbaesongdate						'배송일(출고일)
	public ftot_itemno						'총건수
	public ftot_reducedPrice				'반품/취소시 환불액
	public ftot_reducedPrice_d				'반품/취소시 환불액 배송비
	public ftot_buycash						'
	public ftot_buycashCouponNotApplied		'상품쿠폰적용전 실매입가
	public ftot_orgitemcost					'소비자가
	public ftot_orgitemcost_d				'소비자가 배송비
	public ftot_itemcostCouponNotApplied	'상품쿠폰적용전 실판매가
	public ftot_itemcostCouponNotApplied_d	'상품쿠폰적용전 실판매가 배송비
	public ftot_itemcost					'실판매가
	public ftot_itemcost_d					'실판매가 배송비
	public ftot_DivSpendCouponSum			'정액쿠폰 안분
	public ftot_DivSpendCouponSum_d			'정액쿠폰 안분 배송비

	public ftot_DivSpendMileSum				'마일리지 안분
	public ftot_DivSpendMileSum_d			'마일리지 안분 배송비

	public fsellType						'기본 매출계정 코드
	public fsellTypeName					'기본 매출계정 이름
	public fsitename						'
	public fsellBizCdName					'기본 매출부서

	public fjPrice                          ''정산액
    public fjPriceEtc                       ''기타정산(반품배송비등)
    public fjPriceEtcChulgo                 ''기타출고정산

    public FHanDlePriceNoVat                ''취급액 Vat 제외
    public ftot_buycashNoVat                ''매입가 Vat 제외
	public fomwdiv
	public fsellbizcd                       ''기본매출부서

	public function getHanDlePrice() ''취급액
	    getHanDlePrice = ftot_reducedPrice-ftot_DivSpendCouponSum
    end function

    public function getCalcuMeachul() ''매출액
        getCalcuMeachul = -1
        IF (fonoff="ON") and (fitemdiv<>"IT") then       '' 온라인
            if (fomwdiv="M") or (fomwdiv="Z") or (fomwdiv = "C") or (fomwdiv = "E") then
                getCalcuMeachul = getHanDlePrice
            ELSEIF (fomwdiv="U") or (fomwdiv="W") then
                getCalcuMeachul = (getHanDlePrice-ftot_buycash)
            ELSEIF (fomwdiv="Y") then
                getCalcuMeachul = 0
            END IF
        ELSEIF (fonoff="ON") and (fitemdiv="IT") then   '' 아이띵소_온라인
            getCalcuMeachul = getHanDlePrice
        ELSEIF (fonoff="AC") then   '' 아카데미
            IF (fomwdiv="A") or (fomwdiv="D")  then
                getCalcuMeachul = (getHanDlePrice-ftot_buycash)
            ELSEIF (fomwdiv="Y") then
                getCalcuMeachul = 0
            End If
        ELSEIF (fonoff="OF") and (fitemdiv<>"IT") then   '' 오프
            if (fomwdiv="B012") then
                getCalcuMeachul = (getHanDlePrice-ftot_buycash)
            else
                getCalcuMeachul = getHanDlePrice
            end if
        ELSEIF (fonoff="OF") and (fitemdiv="IT") then   '' 아이띵소_오프라인
            getCalcuMeachul = getHanDlePrice
        END IF

    end function

    public function getCalcuMeachulNoVat() ''매출액(Vat 제외) ''사용안함
        'getCalcuMeachulNoVat = 0
        'exit function
        getCalcuMeachulNoVat = -1
        IF (fonoff="ON") and (fitemdiv<>"IT") then       '' 온라인
            if (fomwdiv="M") or (fomwdiv="Z") or (fomwdiv = "C") or (fomwdiv = "E") then
                getCalcuMeachulNoVat = FHanDlePriceNoVat
            ELSEIF (fomwdiv="U") or (fomwdiv="W") then
                getCalcuMeachulNoVat = (FHanDlePriceNoVat-ftot_buycashNoVat)
            ELSEIF (fomwdiv="Y") then
                getCalcuMeachulNoVat = 0
            END IF
        ELSEIF (fonoff="ON") and (fitemdiv="IT") then   '' 아이띵소_온라인
            getCalcuMeachulNoVat = FHanDlePriceNoVat
        ELSEIF (fonoff="AC") then   '' 아카데미
            IF (fomwdiv="A") or (fomwdiv="D") then
                getCalcuMeachulNoVat = (FHanDlePriceNoVat-ftot_buycashNoVat)
            ELSEIF (fomwdiv="Y") then
                getCalcuMeachulNoVat = 0
            End If
        ELSEIF (fonoff="OF") and (fitemdiv<>"IT") then   '' 오프
            if (fomwdiv="B012") then
                getCalcuMeachulNoVat = (FHanDlePriceNoVat-ftot_buycashNoVat)
            else
                getCalcuMeachulNoVat = FHanDlePriceNoVat
            end if
        ELSEIF (fonoff="OF") and (fitemdiv="IT") then   '' 아이띵소_오프라인
            getCalcuMeachulNoVat = FHanDlePriceNoVat

        END IF

    end function

    public function getErrJungsan() ''정산 매입가 오차
        getErrJungsan = 0
        if (fomwdiv="U") or (fomwdiv="W") or (fomwdiv="Y") or (fomwdiv="A") or (fomwdiv="D") then
            getErrJungsan = ftot_buycash-fjPrice
        end if

        if (fonoff="OF") and (fomwdiv="B012") then
            getErrJungsan = ftot_buycash-fjPrice
        end if
    end function

    public function getOnOffGubunName()
        getOnOffGubunName = fonoff
    end function

    public function getItemGubunName()
        getItemGubunName = fitemdiv
    end function

	public function getMwGubunName()
	    getMwGubunName =""
	    if IsNULL(fomwdiv) then Exit function

	    if (fomwdiv="M") then
	        getMwGubunName = "매입"
	    elseif (fomwdiv="W") then
	        getMwGubunName = "위탁"
	    elseif (fomwdiv="U") then
	        getMwGubunName = "업체"
	    elseif (fomwdiv="Y") then
	        getMwGubunName = "업배"
	    elseif (fomwdiv="Z") then
	        getMwGubunName = "텐배"
	    elseif (fomwdiv="A") then
	        getMwGubunName = "강좌"
	    elseif (fomwdiv="D") then
	        getMwGubunName = "DIY"
	    elseif (fomwdiv="C") then
	        getMwGubunName = "수수료"
	    elseif (fomwdiv="E") then
	        getMwGubunName = "오차"
	    elseif (fomwdiv="P") then
	        getMwGubunName = "포장비"
	    elseif (fomwdiv="B000") then
	        getMwGubunName = "미지정"
	    elseif (fomwdiv="B011") then
	        getMwGubunName = "위탁판매"
	    elseif (fomwdiv="B012") then
	        getMwGubunName = "업체위탁"
	    elseif (fomwdiv="B013") then
	        getMwGubunName = "출고위탁"
	    elseif (fomwdiv="B021") then
	        getMwGubunName = "오프매입"
	    elseif (fomwdiv="B022") then
	        getMwGubunName = "매장매입"
	    elseif (fomwdiv="B023") then
	        getMwGubunName = "가맹점매입"
	    elseif (fomwdiv="B031") then
	        getMwGubunName = "출고매입"
	    elseif (fomwdiv="B032") then
	        getMwGubunName = "센터매입"
	    elseif (fomwdiv="B999") then
	        getMwGubunName = "기타보정"


	    else
	        getMwGubunName = fomwdiv
	    end if
    end function

end class

class cManagementSupportMaechul_list
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public flist

	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	public FRectOnOff
	public FRectStartdate
	public FRectEndDate
	public frectdatecancle
	public frectbancancle
	public frectaccountdiv
	public frectsitename
	public frectipkumdatesucc
	public frectpurchasetype
	public frectvatinclude
	public FRectDLVdiv
	public frectGroupByMwDiv
	public frectGroupByMonth
	public frectGroupBySitename
	public FRectBizSectionCd
    public FRectSupptype        '' 공급가/ 합계
	public frectdatetype
	public frectinccancel
	public frectitemoption
	public frectitemstate
	public frectw10102
	public frectm10102
	public frecta10102
	public fArrLIst

	public function fmaechul_list			'일별매출통계
	dim i , sql

		sql = "SELECT "
		sql = sql & "	MST.onoff, "
		if (frectGroupByMonth="m") then
		    sql = sql & "	convert(varchar(7),MST.beasongdate,21) as beasongdate, "
		else
    		sql = sql & "	MST.beasongdate, "
    	end if

    	sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_itemno else 0 END) as tot_itemno, "
    	sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_buycash else 0 END) as tot_buycash, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_buycashCouponNotApplied else 0 END) as tot_buycashCouponNotApplied, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_reducedPrice else 0 END) as tot_reducedPrice, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_reducedPrice else 0 END) as tot_reducedPrice_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_orgitemcost else 0 END) as tot_orgitemcost, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_orgitemcost else 0 END) as tot_orgitemcost_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_itemcostCouponNotApplied else 0 END) as tot_itemcostCouponNotApplied, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_itemcostCouponNotApplied else 0 END) as tot_itemcostCouponNotApplied_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_itemcost else 0 END) as tot_itemcost, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_itemcost else 0 END) as tot_itemcost_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_DivSpendCouponSum else 0 END) as tot_DivSpendCouponSum, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_DivSpendCouponSum else 0 END) as tot_DivSpendCouponSum_d,  "
	    sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_DivSpendMileSum else 0 END) as tot_DivSpendMileSum, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_DivSpendMileSum else 0 END) as tot_DivSpendMileSum_d,  "

		sql = sql & "	count(MST.beasongdate) as cnt "
		sql = sql & "FROM [db_datamart].[dbo].[tbl_ManagementSupportTeam_Daily_totalsale] AS MST "

		If frectpurchasetype <> "" Then
			sql = sql & " INNER JOIN [TENDB].[db_partner].[dbo].[tbl_partner] AS P ON MST.sitename = P.id AND P.purchasetype = '" & frectpurchasetype & "' "
		End IF

		sql = sql & " WHERE 1=1 "
		sql = sql & " and MST.onoff in ('ON')" '' 오프라인 별도
        sql = sql & " and MST.itemdiv not in ('OC','OE')" '' 2013/04/08 추가

		if FRectOnOff <> "" then
			sql = sql & " AND MST.onoff = '" & FRectOnOff & "' "
		end if
		if frectsitename <> "" then
			sql = sql & " AND MST.sitename = '" & frectsitename & "' "
		end if
		if frectaccountdiv <> "" then
			sql = sql & " AND MST.accountdiv = '" & frectaccountdiv & "' "
		end if

		sql = sql & " AND MST.beasongdate BETWEEN '"& FRectStartdate& "' AND '" &FRectEndDate & "' "

		if frectbancancle = "1" then
		elseif frectbancancle = "2" then
			sql = sql & " AND MST.jumundiv = '9' "
		else
			sql = sql & " AND MST.jumundiv <> '9' "
		end if

		if (frectvatinclude<>"") then
		    sql = sql & " AND MST.vatinclude = '" & vatinclude & "' "
		end if

		sql = sql & " GROUP BY MST.onoff"
		if (frectGroupByMonth="m") then
		    sql = sql & "	,convert(varchar(7),MST.beasongdate,21) "
		else
    		sql = sql & "	,MST.beasongdate "
    	end if

		if (frectGroupByMonth="m") then
			sql = sql & " ORDER BY convert(varchar(7),MST.beasongdate,21) DESC "
		else
			sql = sql & " ORDER BY MST.beasongdate DESC "
		end if

		''response.write sql&"<br>"
		''dbget.close() : response.end
	db3_rsget.open sql,db3_dbget,1

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new cManagementSupportMaechul_oneitem
				flist(i).fonoff								= db3_rsget("onoff")
				flist(i).fbaesongdate 						= db3_rsget("beasongdate")
				flist(i).ftot_itemno           				= db3_rsget("tot_itemno")
				flist(i).ftot_reducedPrice              	= db3_rsget("tot_reducedPrice")
				flist(i).ftot_reducedPrice_d				= db3_rsget("tot_reducedPrice_d")
				flist(i).ftot_buycash 						= db3_rsget("tot_buycash")
				flist(i).ftot_buycashCouponNotApplied   	= db3_rsget("tot_buycashCouponNotApplied")
				flist(i).ftot_orgitemcost              		= db3_rsget("tot_orgitemcost")
				flist(i).ftot_orgitemcost_d             	= db3_rsget("tot_orgitemcost_d")
				flist(i).ftot_itemcostCouponNotApplied 		= db3_rsget("tot_itemcostCouponNotApplied")
				flist(i).ftot_itemcostCouponNotApplied_d    = db3_rsget("tot_itemcostCouponNotApplied_d")
				flist(i).ftot_itemcost 						= db3_rsget("tot_itemcost")
				flist(i).ftot_itemcost_d 					= db3_rsget("tot_itemcost_d")
				flist(i).ftot_DivSpendCouponSum				= db3_rsget("tot_DivSpendCouponSum")
                flist(i).ftot_DivSpendCouponSum_d			= db3_rsget("tot_DivSpendCouponSum_d")

                flist(i).ftot_DivSpendMileSum               = db3_rsget("tot_DivSpendMileSum")
                flist(i).ftot_DivSpendMileSum_d             = db3_rsget("tot_DivSpendMileSum_d")

		db3_rsget.movenext
		i = i + 1
		loop
	end if

	db3_rsget.close
	end function

	public function fmaechul_listByGbn			'매입구분별 매출통계
	dim i , sql

		sql = "SELECT "
		sql = sql & "	MST.onoff, MST.itemdiv, isNULL(p.sellbizcd,'0000000000') as sellbizcd,"
		if (frectGroupByMonth="m") then
		    sql = sql & "	convert(varchar(7),MST.beasongdate,21) as beasongdate, "
		else
    		sql = sql & "	MST.beasongdate, "
    	end if
    	if (frectGroupBySitename<>"") then
    	    sql = sql & "	MST.sitename, "
			sql = sql & "	P.sellType, isNull((select pcomm_name FROM db_partner.dbo.tbl_partner_comm_code WHERE pcomm_group = 'sellacccd' and pcomm_cd = P.sellType),'') AS sellTypeName, " ''[TENDB].
    	end if
    	sql = sql & "	MST.omwdiv, "
    	sql = sql & "	isNull((select BIZSECTION_NM FROM db_partner.dbo.tbl_TMS_BA_BIZSECTION WHERE BIZSECTION_CD = isNULL(p.sellbizcd,'0000000000')),'') AS sellBizCdName, " ''[TENDB].
    	sql = sql & "	sum(MST.tot_itemno) as tot_itemno, "
        IF (FRectSupptype="S") then
            sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_buycash*10/11 ELSE MST.tot_buycash END) as tot_buycash, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_buycashCouponNotApplied*10/11 else MST.tot_buycashCouponNotApplied END) as tot_buycashCouponNotApplied, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_reducedPrice*10/11 ELSE MST.tot_reducedPrice END) as tot_reducedPrice, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_orgitemcost*10/11 ELSE MST.tot_orgitemcost END) as tot_orgitemcost, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_itemcostCouponNotApplied*10/11 ELSE MST.tot_itemcostCouponNotApplied END) as tot_itemcostCouponNotApplied, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_itemcost*10/11 ELSE MST.tot_itemcost END) as tot_itemcost, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN isNULL(MST.tot_DivSpendCouponSum,0)*10/11 ELSE isNULL(MST.tot_DivSpendCouponSum,0) END) as tot_DivSpendCouponSum, "
    	    sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN isNULL(MST.tot_DivSpendMileSum,0)*10/11 ELSE isNULL(MST.tot_DivSpendMileSum,0) END) as tot_DivSpendMileSum, "
        ELSE
        	sql = sql & "	sum(MST.tot_buycash) as tot_buycash, "
    		sql = sql & "	sum(MST.tot_buycashCouponNotApplied) as tot_buycashCouponNotApplied, "
    		sql = sql & "	sum(MST.tot_reducedPrice) as tot_reducedPrice, "
    		sql = sql & "	sum(MST.tot_orgitemcost) as tot_orgitemcost, "
    		sql = sql & "	sum(MST.tot_itemcostCouponNotApplied) as tot_itemcostCouponNotApplied, "
    		sql = sql & "	sum(MST.tot_itemcost) as tot_itemcost, "
    		sql = sql & "	sum(isNULL(MST.tot_DivSpendCouponSum,0)) as tot_DivSpendCouponSum, "
    	    sql = sql & "	sum(isNULL(MST.tot_DivSpendMileSum,0)) as tot_DivSpendMileSum, "
	    END IF
	    sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' then (MST.tot_reducedPrice-isNULL(MST.tot_DivSpendCouponSum,0))*10/11 ELSE (MST.tot_reducedPrice-isNULL(MST.tot_DivSpendCouponSum,0)) END) as HanDlePriceNoVat,"
		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' then (MST.tot_buycash)*10/11 ELSE (MST.tot_buycash) END) as tot_buycashNoVat,"
		sql = sql & "	count(MST.beasongdate) as cnt "
		sql = sql & "	,isNULL(j.jPrice,0) as jPrice"
		sql = sql & "	,isNULL(j.jPriceEtc,0) as jPriceEtc"
		sql = sql & "	,isNULL(j.jPriceEtcChulgo,0) as jPriceEtcChulgo"

		sql = sql & " FROM [db_datamart].[dbo].[tbl_ManagementSupportTeam_Daily_totalsale] AS MST "

		If (frectpurchasetype <> "") or (FRectBizSectionCd<>"") Then
		    sql = sql & "  JOIN [TENDB].[db_partner].[dbo].[tbl_partner] AS P ON MST.sitename = P.id " ''[TENDB].
		    if (frectpurchasetype <> "") then
    		    sql = sql & " AND P.purchasetype = '" & frectpurchasetype & "' "
    	    end if

    	    if (FRectBizSectionCd<>"") then
    			sql = sql & " AND isNULL(p.sellbizcd,'0000000000')='"&FRectBizSectionCd&"'"
    		end if

		else
		    sql = sql & "  LEFT JOIN [TENDB].[db_partner].[dbo].[tbl_partner] AS P ON MST.sitename = P.id " ''[TENDB].
		End IF

		sql = sql & " left join (select "
        sql = sql & " j.yyyymm, j.targetGbn, j.itemGbn"

        sql = sql & " ,(CASE WHEN j.mwgbn='witakchulgo' and subflag=0 then 'Y'"
        sql = sql & " 		WHEN j.mwgbn='upche' and subflag=0 then 'Y'"
        sql = sql & " 		WHEN j.mwgbn='witakchulgo' and subflag<>0 then 'W'"
        sql = sql & " 		WHEN j.mwgbn='maeip' then 'M'"
        sql = sql & " 		WHEN j.mwgbn='upche' then 'U'"
        sql = sql & " 		WHEN j.mwgbn='witaksell' then 'W'"
        sql = sql & " 		WHEN j.mwgbn='D' and subflag=0 then 'Y'"
        sql = sql & " 		ELSE j.mwgbn END) as mwgbn"
        IF (FRectSupptype="S") then
            sql = sql & " ,sum(CASE WHEN j.mwgbn<>'witakchulgo' THEN (CASE WHEN j.taxtype='01' THEN totSuplySum*10/11 ELSE totSuplySum END) ELSE 0 END) as jPrice"
            sql = sql & " ,sum(CASE WHEN j.subflag=0 and j.mwgbn='witakchulgo' THEN (CASE WHEN j.taxtype='01' THEN totSuplySum*10/11 ELSE totSuplySum end) ELSE 0 END) as jPriceEtc"
            sql = sql & " ,sum(CASE WHEN j.subflag<>0 and j.mwgbn='witakchulgo' THEN (CASE WHEN j.taxtype='01' THEN totSuplySum*10/11 ELSE totSuplySum end) ELSE 0 END) as jPriceEtcChulgo"
        ELSE
            sql = sql & " ,sum(CASE WHEN j.mwgbn<>'witakchulgo' THEN (totSuplySum) ELSE 0 END) as jPrice"
            sql = sql & " ,sum(CASE WHEN j.subflag=0 and j.mwgbn='witakchulgo' THEN (totSuplySum) ELSE 0 END) as jPriceEtc"
            sql = sql & " ,sum(CASE WHEN j.subflag<>0 and j.mwgbn='witakchulgo' THEN (totSuplySum) ELSE 0 END) as jPriceEtcChulgo"
        END IF
        sql = sql & " from db_datamart.dbo.tbl_monthly_jungsan_sum j"
        sql = sql & " where j.yyyymm>='"&Left(FRectStartdate,7)&"'"
        sql = sql & " and j.yyyymm<='"&Left(FRectEndDate,7)&"'"
        sql = sql & " and j.mwgbn <> 'maeipchulgo'"
        if FRectOnOff <> "" then
            if (FRectOnOff="NOAC") then
                sql = sql & " and j.itemGbn<>'AC'"
            else
                sql = sql & " and j.itemGbn='"&FRectOnOff&"'"
            end if
        end if

        if (frectvatinclude<>"") then
            if frectvatinclude="Y" then
                sql = sql & " and j.taxtype='01'"
            elseif frectvatinclude="N" then
                sql = sql & " and j.taxtype<>'01'"
            end if
        end if
        sql = sql & " group by j.yyyymm,j.targetGbn,j.itemGbn,"

        sql = sql & " (CASE  "
        sql = sql & " 		WHEN j.mwgbn='witakchulgo' and subflag=0 then 'Y'"
        sql = sql & " 		WHEN j.mwgbn='upche' and subflag=0 then 'Y'"
        sql = sql & " 		WHEN j.mwgbn='witakchulgo' and subflag<>0 then 'W'"
        sql = sql & " 		WHEN j.mwgbn='maeip' then 'M'"
        sql = sql & " 		WHEN j.mwgbn='upche' then 'U'"
        sql = sql & " 		WHEN j.mwgbn='witaksell' then 'W'"
        sql = sql & " 		WHEN j.mwgbn='D' and subflag=0 then 'Y'"
        sql = sql & " 		ELSE j.mwgbn END)"
        sql = sql & " ) J on convert(varchar(7),MST.beasongdate,21)=j.yyyymm"
        sql = sql & " and MST.omwdiv=J.mwgbn"
        sql = sql & " and MST.onoff=J.targetGbn"
		sql = sql & " and MST.itemdiv=J.itemGbn"

		sql = sql & " WHERE 1=1 "
		sql = sql & " and MST.onoff in ('ON', 'AC') " '' 오프라인 별도
		if FRectOnOff <> "" then
		    if (FRectOnOff="NOAC") then
                sql = sql & " and MST.itemdiv<>'AC'"
            else
			    sql = sql & " AND MST.itemdiv = '" & FRectOnOff & "' "
			end if
		end if

		if FRectDLVdiv <> "" then
		    if (FRectDLVdiv="s") then
		        sql = sql & " AND MST.omwdiv not in ('Y','Z')"
		    elseif (FRectDLVdiv="d") then
    			sql = sql & " AND MST.omwdiv in ('Y','Z')"
    	    else
    	        sql = sql & " AND MST.omwdiv='"&FRectDLVdiv&"'"
    		end if
		end if

		if frectsitename <> "" then
			sql = sql & " AND MST.sitename = '" & frectsitename & "' "
		end if
		if frectaccountdiv <> "" then
			sql = sql & " AND MST.accountdiv = '" & frectaccountdiv & "' "
		end if

		sql = sql & " AND MST.beasongdate BETWEEN '"& FRectStartdate& "' AND '" &FRectEndDate & "' "

		if frectbancancle = "1" then
		elseif frectbancancle = "2" then
			sql = sql & " AND MST.jumundiv = '9' "
		else
			sql = sql & " AND MST.jumundiv <> '9' "
		end if

		if (frectvatinclude<>"") then
		    sql = sql & " AND MST.vatinclude = '" & vatinclude & "' "
		end if

		sql = sql & " GROUP BY MST.onoff,MST.itemdiv,isNULL(p.sellbizcd,'0000000000') ,isNULL(j.jPrice,0),isNULL(j.jPriceEtc,0),isNULL(j.jPriceEtcChulgo,0)"
		if (frectGroupByMonth="m") then
		    sql = sql & "	,convert(varchar(7),MST.beasongdate,21) "
		else
    		sql = sql & "	,MST.beasongdate "
    	end if
    	if (frectGroupBySitename<>"") then
    	    sql = sql & "	,MST.sitename "
			sql = sql & "	,P.sellType "
    	end if
    	sql = sql & "	,MST.omwdiv "
		sql = sql & " ORDER BY beasongdate DESC, sellbizcd, MST.onoff desc "
		sql = sql & "	,MST.omwdiv "
		if (frectGroupBySitename<>"") then
    	    sql = sql & "	,MST.sitename "
			sql = sql & "	,P.sellType "
    	end if
'rw sql
	db3_rsget.open sql,db3_dbget,1

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new cManagementSupportMaechul_oneitem
				flist(i).fonoff								= db3_rsget("onoff")
				flist(i).fitemdiv							= db3_rsget("itemdiv")
				flist(i).fbaesongdate 						= db3_rsget("beasongdate")
				flist(i).ftot_itemno           				= db3_rsget("tot_itemno")
				flist(i).ftot_reducedPrice              	= db3_rsget("tot_reducedPrice")
				flist(i).ftot_buycash 						= db3_rsget("tot_buycash")
				flist(i).ftot_buycashCouponNotApplied   	= db3_rsget("tot_buycashCouponNotApplied")
				flist(i).ftot_orgitemcost              		= db3_rsget("tot_orgitemcost")
				flist(i).ftot_itemcostCouponNotApplied 		= db3_rsget("tot_itemcostCouponNotApplied")
				flist(i).ftot_itemcost 						= db3_rsget("tot_itemcost")
				flist(i).ftot_DivSpendCouponSum				= db3_rsget("tot_DivSpendCouponSum")

                flist(i).ftot_DivSpendMileSum               = db3_rsget("tot_DivSpendMileSum")

                flist(i).fjPrice                    = db3_rsget("jPrice")
                flist(i).fjPriceEtc                 = db3_rsget("jPriceEtc")
                flist(i).fjPriceEtcChulgo           = db3_rsget("jPriceEtcChulgo")

                flist(i).FHanDlePriceNoVat          = db3_rsget("HanDlePriceNoVat")
                flist(i).ftot_buycashNoVat          = db3_rsget("tot_buycashNoVat")
                flist(i).fomwdiv								= db3_rsget("omwdiv")
                if (frectGroupBySitename<>"") then
                    flist(i).fsitename				= db3_rsget("sitename")
					flist(i).fsellTypeName			= db3_rsget("sellTypeName")
                end if

                flist(i).fsellbizcd= db3_rsget("sellbizcd")
                flist(i).fsellBizCdName	= db3_rsget("sellBizCdName")
		db3_rsget.movenext
		i = i + 1
		loop
	end if

	db3_rsget.close
	end function

	public function fmaechul_listOFByGbn			'매입구분별 매출통계(오프)
	dim i , sql

		sql = "SELECT "
		sql = sql & "	MST.onoff, MST.itemdiv, isNULL(p.sellbizcd,'0000000000') as sellbizcd,"
		if (frectGroupByMonth="m") then
		    sql = sql & "	convert(varchar(7),MST.beasongdate,21) as beasongdate, "
		else
    		sql = sql & "	MST.beasongdate, "
    	end if
    	sql = sql & "	isNull((select BIZSECTION_NM FROM db_partner.dbo.tbl_TMS_BA_BIZSECTION WHERE BIZSECTION_CD = isNULL(p.sellbizcd,'0000000000')),'') AS sellBizCdName, " ''[TENDB].
    	sql = sql & "	MST.omwdiv,MST.sitename, "
		sql = sql & "	P.sellType, isNull((select pcomm_name FROM db_partner.dbo.tbl_partner_comm_code WHERE pcomm_group = 'sellacccd' and pcomm_cd = P.sellType),'') AS sellTypeName, " ''[TENDB].
    	sql = sql & "	sum(MST.tot_itemno) as tot_itemno, "
    	IF (FRectSupptype="S") then
    	    sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_buycash*10/11 ELSE MST.tot_buycash END) as tot_buycash, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_buycashCouponNotApplied*10/11 else MST.tot_buycashCouponNotApplied END) as tot_buycashCouponNotApplied, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_reducedPrice*10/11 ELSE MST.tot_reducedPrice END) as tot_reducedPrice, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_orgitemcost*10/11 ELSE MST.tot_orgitemcost END) as tot_orgitemcost, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_itemcostCouponNotApplied*10/11 ELSE MST.tot_itemcostCouponNotApplied END) as tot_itemcostCouponNotApplied, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN MST.tot_itemcost*10/11 ELSE MST.tot_itemcost END) as tot_itemcost, "
    		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN isNULL(MST.tot_DivSpendCouponSum,0)*10/11 ELSE isNULL(MST.tot_DivSpendCouponSum,0) END) as tot_DivSpendCouponSum, "
    	    sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' THEN isNULL(MST.tot_DivSpendMileSum,0)*10/11 ELSE isNULL(MST.tot_DivSpendMileSum,0) END) as tot_DivSpendMileSum, "
    	ELSE
        	sql = sql & "	sum(MST.tot_buycash) as tot_buycash, "
    		sql = sql & "	sum(MST.tot_buycashCouponNotApplied) as tot_buycashCouponNotApplied, "
    		sql = sql & "	sum(MST.tot_reducedPrice) as tot_reducedPrice, "
    		sql = sql & "	sum(MST.tot_orgitemcost) as tot_orgitemcost, "
    		sql = sql & "	sum(MST.tot_itemcostCouponNotApplied) as tot_itemcostCouponNotApplied, "
    		sql = sql & "	sum(MST.tot_itemcost) as tot_itemcost, "
    		sql = sql & "	sum(isNULL(MST.tot_DivSpendCouponSum,0)) as tot_DivSpendCouponSum, "
    	    sql = sql & "	sum(isNULL(MST.tot_DivSpendMileSum,0)) as tot_DivSpendMileSum, "
    	END IF
	    sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' then (MST.tot_reducedPrice-isNULL(MST.tot_DivSpendCouponSum,0))*10/11 ELSE (MST.tot_reducedPrice-isNULL(MST.tot_DivSpendCouponSum,0)) END) as HanDlePriceNoVat,"
		sql = sql & "	sum(CASE WHEN MST.vatinclude='Y' then (MST.tot_buycash)*10/11 ELSE (MST.tot_buycash) END) as tot_buycashNoVat,"
		sql = sql & "	count(MST.beasongdate) as cnt "
		sql = sql & "	,isNULL(j.jPrice,0) as jPrice"
		sql = sql & "	,isNULL(j.jPriceEtc,0) as jPriceEtc"
		sql = sql & "	,isNULL(j.jPriceEtcChulgo,0) as jPriceEtcChulgo"

		sql = sql & " FROM [db_datamart].[dbo].[tbl_ManagementSupportTeam_Daily_totalsale] AS MST "

		If (frectpurchasetype <> "") or (FRectBizSectionCd<>"") Then
		    sql = sql & "  JOIN [TENDB].[db_partner].[dbo].[tbl_partner] AS P ON MST.sitename = P.id " ''[TENDB].
		    if (frectpurchasetype <> "") then
    		    sql = sql & " AND P.purchasetype = '" & frectpurchasetype & "' "
    	    end if

    	    if (FRectBizSectionCd<>"") then
    			sql = sql & " AND Left(isNULL(p.sellbizcd,'0000000000'),8)=Left('"&FRectBizSectionCd&"',8)"
    		end if

		else
		    sql = sql & "  LEFT JOIN [TENDB].[db_partner].[dbo].[tbl_partner] AS P ON MST.sitename = P.id " ''[TENDB].
		End IF

		sql = sql & " left join (select "
        sql = sql & " j.yyyymm, j.targetGbn, j.itemGbn"
        sql = sql & " ,j.mwgbn,j.sitename"
        IF (FRectSupptype="S") then
            sql = sql & " ,sum(CASE WHEN j.mwgbn<>'B999' THEN (CASE WHEN j.taxtype='01' THEN totSuplySum*10/11 ELSE totSuplySum END) ELSE 0 END) as jPrice"
            sql = sql & " ,sum(CASE WHEN j.subflag=0 and j.mwgbn='B999' THEN (CASE WHEN j.taxtype='01' THEN totSuplySum*10/11 ELSE totSuplySum end) ELSE 0 END) as jPriceEtc"
            sql = sql & " ,sum(CASE WHEN j.subflag<>0 and j.mwgbn='B999' THEN (CASE WHEN j.taxtype='01' THEN totSuplySum*10/11 ELSE totSuplySum end) ELSE 0 END) as jPriceEtcChulgo"
        ELSE
            sql = sql & " ,sum(CASE WHEN j.mwgbn<>'B999' THEN (totSuplySum) ELSE 0 END) as jPrice"
            sql = sql & " ,sum(CASE WHEN j.subflag=0 and j.mwgbn='B999' THEN (totSuplySum) ELSE 0 END) as jPriceEtc"
            sql = sql & " ,sum(CASE WHEN j.subflag<>0 and j.mwgbn='B999' THEN (totSuplySum) ELSE 0 END) as jPriceEtcChulgo"
        END IF
        sql = sql & " from db_datamart.dbo.tbl_monthly_jungsan_sum j"
        sql = sql & " where j.yyyymm>='"&Left(FRectStartdate,7)&"'"
        sql = sql & " and j.yyyymm<='"&Left(FRectEndDate,7)&"'"
        sql = sql & " and j.mwgbn <> 'maeipchulgo'"
        if FRectOnOff <> "" then
            sql = sql & " and j.itemGbn='"&FRectOnOff&"'"
        end if

        if (frectvatinclude<>"") then
            if frectvatinclude="Y" then
                sql = sql & " and j.taxtype='01'"
            elseif frectvatinclude="N" then
                sql = sql & " and j.taxtype<>'01'"
            end if
        end if
        sql = sql & " group by j.yyyymm,j.targetGbn,j.itemGbn,"
        sql = sql & " j.mwgbn,j.sitename "
        sql = sql & " ) J on convert(varchar(7),MST.beasongdate,21)=j.yyyymm"
        sql = sql & " and MST.omwdiv=J.mwgbn"
        sql = sql & " and MST.onoff=J.targetGbn"
		sql = sql & " and MST.itemdiv=J.itemGbn"
        sql = sql & " and MST.sitename=J.sitename"

		sql = sql & " WHERE 1=1 "
		sql = sql & " and MST.onoff = 'OF' " '' 오프라인 별도
		if FRectOnOff <> "" then
			sql = sql & " AND MST.itemdiv = '" & FRectOnOff & "' "
		end if

		if FRectDLVdiv <> "" then
		    if (FRectDLVdiv="s") then
		        sql = sql & " AND MST.omwdiv not in ('B012')"
		    elseif (FRectDLVdiv="d") then
    			sql = sql & " AND MST.omwdiv in ('B012')"
    	    else
    	        sql = sql & " AND MST.omwdiv='"&FRectDLVdiv&"'"
    		end if
		end if

		if frectsitename <> "" then
			sql = sql & " AND MST.sitename = '" & frectsitename & "' "
		end if
		if frectaccountdiv <> "" then
			sql = sql & " AND MST.accountdiv = '" & frectaccountdiv & "' "
		end if

		sql = sql & " AND MST.beasongdate BETWEEN '"& FRectStartdate& "' AND '" &FRectEndDate & "' "

		if frectbancancle = "1" then
		elseif frectbancancle = "2" then
			sql = sql & " AND MST.jumundiv = '1' "
		else
			sql = sql & " AND MST.jumundiv <> '1' "
		end if

		if (frectvatinclude<>"") then
		    sql = sql & " AND MST.vatinclude = '" & vatinclude & "' "
		end if

		sql = sql & " GROUP BY MST.onoff,MST.itemdiv,isNULL(p.sellbizcd,'0000000000'),isNULL(j.jPrice,0),isNULL(j.jPriceEtc,0),isNULL(j.jPriceEtcChulgo,0)"
		if (frectGroupByMonth="m") then
		    sql = sql & "	,convert(varchar(7),MST.beasongdate,21) "
		else
    		sql = sql & "	,MST.beasongdate "
    	end if

    	sql = sql & "	,MST.omwdiv,MST.sitename "
		sql = sql & "	,P.sellType "
		sql = sql & " ORDER BY MST.beasongdate DESC,sellbizcd, MST.onoff desc "
		sql = sql & "	,MST.omwdiv "
'rw sql
	db3_rsget.open sql,db3_dbget,1

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new cManagementSupportMaechul_oneitem
				flist(i).fonoff								= db3_rsget("onoff")
				flist(i).fitemdiv							= db3_rsget("itemdiv")
				flist(i).fbaesongdate 						= db3_rsget("beasongdate")
				flist(i).ftot_itemno           				= db3_rsget("tot_itemno")
				flist(i).ftot_reducedPrice              	= db3_rsget("tot_reducedPrice")
				flist(i).ftot_buycash 						= db3_rsget("tot_buycash")
				flist(i).ftot_buycashCouponNotApplied   	= db3_rsget("tot_buycashCouponNotApplied")
				flist(i).ftot_orgitemcost              		= db3_rsget("tot_orgitemcost")
				flist(i).ftot_itemcostCouponNotApplied 		= db3_rsget("tot_itemcostCouponNotApplied")
				flist(i).ftot_itemcost 						= db3_rsget("tot_itemcost")
				flist(i).ftot_DivSpendCouponSum				= db3_rsget("tot_DivSpendCouponSum")

                flist(i).ftot_DivSpendMileSum               = db3_rsget("tot_DivSpendMileSum")

                flist(i).fjPrice                    = db3_rsget("jPrice")
                flist(i).fjPriceEtc                 = db3_rsget("jPriceEtc")
                flist(i).fjPriceEtcChulgo           = db3_rsget("jPriceEtcChulgo")

                flist(i).FHanDlePriceNoVat          = db3_rsget("HanDlePriceNoVat")
                flist(i).ftot_buycashNoVat          = db3_rsget("tot_buycashNoVat")
                flist(i).fomwdiv					= db3_rsget("omwdiv")
                flist(i).fsitename					= db3_rsget("sitename")
				flist(i).fsellTypeName				= db3_rsget("sellTypeName")

                flist(i).fsellbizcd= db3_rsget("sellbizcd")
                flist(i).fsellBizCdName	= db3_rsget("sellBizCdName")
		db3_rsget.movenext
		i = i + 1
		loop
	end if

	db3_rsget.close
	end function

	public function fconsumer_list_selltype			'소비자매출[ON] - 윗부분 계정과목별
	dim i , sql

		sql = "SELECT "
		sql = sql & "	isNull(P.sellType,0) AS sellType, Convert(varchar(7),MST.beasongdate,120) AS beasongdate, "
		sql = sql & "	isNull((select pcomm_name FROM db_partner.dbo.tbl_partner_comm_code WHERE pcomm_group = 'sellacccd' and pcomm_cd = P.sellType),'') AS sellTypeName, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_itemno else 0 END) as tot_itemno, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_buycash else 0 END) as tot_buycash, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_buycashCouponNotApplied else 0 END) as tot_buycashCouponNotApplied, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_reducedPrice else 0 END) as tot_reducedPrice, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_reducedPrice else 0 END) as tot_reducedPrice_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_orgitemcost else 0 END) as tot_orgitemcost, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_orgitemcost else 0 END) as tot_orgitemcost_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_itemcostCouponNotApplied else 0 END) as tot_itemcostCouponNotApplied, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_itemcostCouponNotApplied else 0 END) as tot_itemcostCouponNotApplied_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_itemcost else 0 END) as tot_itemcost, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_itemcost else 0 END) as tot_itemcost_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_DivSpendCouponSum else 0 END) as tot_DivSpendCouponSum ,"
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_DivSpendCouponSum else 0 END) as tot_DivSpendCouponSum_d,  "
		sql = sql & "	count(MST.beasongdate) as cnt "
		sql = sql & "FROM [db_datamart].[dbo].[tbl_ManagementSupportTeam_Daily_totalsale] AS MST "
		sql = sql & "	INNER JOIN [db_partner].[dbo].[tbl_partner] AS P ON MST.sitename = P.id "
		sql = sql & " WHERE 1=1 "
        sql = sql & " and MST.itemdiv not in ('OC','OE')" '' 2013/04/08 추가

		if frectpurchasetype <> "" then
			sql = sql & " AND P.purchasetype = '" & frectpurchasetype & "' "
		end if
		if FRectOnOff <> "" then
			sql = sql & " AND MST.onoff = '" & FRectOnOff & "' "
		end if
		if frectsitename <> "" then
			sql = sql & " AND MST.sitename = '" & frectsitename & "' "
		end if
		if frectaccountdiv <> "" then
			sql = sql & " AND MST.accountdiv = '" & frectaccountdiv & "' "
		end if

		sql = sql & " AND MST.beasongdate BETWEEN '"& FRectStartdate& "' AND '" &FRectEndDate & "' "

		if frectbancancle = "1" then
		elseif frectbancancle = "2" then
			sql = sql & " AND MST.jumundiv = '9' "
		else
			sql = sql & " AND MST.jumundiv <> '9' "
		end if

		if (frectvatinclude<>"") then
		    sql = sql & " AND MST.vatinclude = '" & vatinclude & "' "
		end if

		sql = sql & " GROUP BY P.sellType, Convert(varchar(7),MST.beasongdate,120) "
		sql = sql & " ORDER BY beasongdate DESC, sellTypeName DESC "
		''response.write sql&"<br>"
	db3_rsget.open sql,db3_dbget,1

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new cManagementSupportMaechul_oneitem
				flist(i).fbaesongdate 						= db3_rsget("beasongdate")
				flist(i).fsellTypeName						= db3_rsget("sellTypeName")
				flist(i).ftot_itemno           				= db3_rsget("tot_itemno")
				flist(i).ftot_reducedPrice              	= db3_rsget("tot_reducedPrice")
				flist(i).ftot_reducedPrice_d				= db3_rsget("tot_reducedPrice_d")
				flist(i).ftot_buycash 						= db3_rsget("tot_buycash")
				flist(i).ftot_buycashCouponNotApplied   	= db3_rsget("tot_buycashCouponNotApplied")
				flist(i).ftot_orgitemcost              		= db3_rsget("tot_orgitemcost")
				flist(i).ftot_orgitemcost_d             	= db3_rsget("tot_orgitemcost_d")
				flist(i).ftot_itemcostCouponNotApplied 		= db3_rsget("tot_itemcostCouponNotApplied")
				flist(i).ftot_itemcostCouponNotApplied_d    = db3_rsget("tot_itemcostCouponNotApplied_d")
				flist(i).ftot_itemcost 						= db3_rsget("tot_itemcost")
				flist(i).ftot_itemcost_d 					= db3_rsget("tot_itemcost_d")
				flist(i).ftot_DivSpendCouponSum				= db3_rsget("tot_DivSpendCouponSum")
				flist(i).ftot_DivSpendCouponSum_d			= db3_rsget("tot_DivSpendCouponSum_d")

		db3_rsget.movenext
		i = i + 1
		loop
	end if

	db3_rsget.close
	end function

	public function fconsumer_list_sitename			'소비자매출[ON] - 아래부분 출고처ID별
	dim i , sql

		sql = "SELECT "
		sql = sql & "	MST.sitename, isNull(P.sellType,0) AS sellType, Convert(varchar(7),MST.beasongdate,120) AS beasongdate, "
		sql = sql & "	isNull((select pcomm_name FROM db_partner.dbo.tbl_partner_comm_code WHERE pcomm_group = 'sellacccd' and pcomm_cd = P.sellType),'') AS sellTypeName, "
		sql = sql & "	isNull((select BIZSECTION_NM FROM db_partner.dbo.tbl_TMS_BA_BIZSECTION WHERE BIZSECTION_CD = P.sellBizCd),'') AS sellBizCdName, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_itemno else 0 END) as tot_itemno, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_buycash else 0 END) as tot_buycash, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_buycashCouponNotApplied else 0 END) as tot_buycashCouponNotApplied, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_reducedPrice else 0 END) as tot_reducedPrice, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_reducedPrice else 0 END) as tot_reducedPrice_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_orgitemcost else 0 END) as tot_orgitemcost, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_orgitemcost else 0 END) as tot_orgitemcost_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_itemcostCouponNotApplied else 0 END) as tot_itemcostCouponNotApplied, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_itemcostCouponNotApplied else 0 END) as tot_itemcostCouponNotApplied_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN MST.tot_itemcost else 0 END) as tot_itemcost, "
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_itemcost else 0 END) as tot_itemcost_d, "
		sql = sql & "	sum(CASE WHEN omwdiv Not in ('Y','Z') THEN isNULL(MST.tot_DivSpendCouponSum,0) else 0 END) as tot_DivSpendCouponSum ,"
		sql = sql & "	sum(CASE WHEN omwdiv in ('Y','Z') THEN MST.tot_DivSpendCouponSum else 0 END) as tot_DivSpendCouponSum_d,  "
		sql = sql & "	count(MST.beasongdate) as cnt "
		sql = sql & "FROM [db_datamart].[dbo].[tbl_ManagementSupportTeam_Daily_totalsale] AS MST "
		sql = sql & "	INNER JOIN [db_partner].[dbo].[tbl_partner] AS P ON MST.sitename = P.id "
		sql = sql & " WHERE 1=1 "
		sql = sql & " and MST.itemdiv not in ('OC','OE')" '' 2013/04/08 추가

		if frectpurchasetype <> "" then
			sql = sql & " AND P.purchasetype = '" & frectpurchasetype & "' "
		end if
		if FRectOnOff <> "" then
			sql = sql & " AND MST.onoff = '" & FRectOnOff & "' "
		end if
		if frectsitename <> "" then
			sql = sql & " AND MST.sitename = '" & frectsitename & "' "
		end if
		if frectaccountdiv <> "" then
			sql = sql & " AND MST.accountdiv = '" & frectaccountdiv & "' "
		end if

		sql = sql & " AND MST.beasongdate BETWEEN '"& FRectStartdate& "' AND '" &FRectEndDate & "' "

		if frectbancancle = "1" then
		elseif frectbancancle = "2" then
			sql = sql & " AND MST.jumundiv = '9' "
		else
			sql = sql & " AND MST.jumundiv <> '9' "
		end if

		if (frectvatinclude<>"") then
		    sql = sql & " AND MST.vatinclude = '" & vatinclude & "' "
		end if

		sql = sql & " GROUP BY MST.sitename, P.sellType, P.sellBizCd, Convert(varchar(7),MST.beasongdate,120) "
		sql = sql & " ORDER BY beasongdate DESC, MST.sitename ASC " ''MST.beasongdate => beasongdate
	''rw sql
		''response.write sql&"<br>"
	db3_rsget.open sql,db3_dbget,1

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new cManagementSupportMaechul_oneitem
				flist(i).fsitename							= db3_rsget("sitename")
				flist(i).fbaesongdate 						= db3_rsget("beasongdate")
				flist(i).fsellTypeName						= db3_rsget("sellTypeName")
				flist(i).fsellBizCdName						= db3_rsget("sellBizCdName")
				flist(i).ftot_itemno           				= db3_rsget("tot_itemno")
				flist(i).ftot_reducedPrice              	= db3_rsget("tot_reducedPrice")
				flist(i).ftot_reducedPrice_d				= db3_rsget("tot_reducedPrice_d")
				flist(i).ftot_buycash 						= db3_rsget("tot_buycash")
				flist(i).ftot_buycashCouponNotApplied   	= db3_rsget("tot_buycashCouponNotApplied")
				flist(i).ftot_orgitemcost              		= db3_rsget("tot_orgitemcost")
				flist(i).ftot_orgitemcost_d             	= db3_rsget("tot_orgitemcost_d")
				flist(i).ftot_itemcostCouponNotApplied 		= db3_rsget("tot_itemcostCouponNotApplied")
				flist(i).ftot_itemcostCouponNotApplied_d    = db3_rsget("tot_itemcostCouponNotApplied_d")
				flist(i).ftot_itemcost 						= db3_rsget("tot_itemcost")
				flist(i).ftot_itemcost_d 					= db3_rsget("tot_itemcost_d")
				flist(i).ftot_DivSpendCouponSum				= db3_rsget("tot_DivSpendCouponSum")
				flist(i).ftot_DivSpendCouponSum_d			= db3_rsget("tot_DivSpendCouponSum_d")

		db3_rsget.movenext
		i = i + 1
		loop
	end if

	db3_rsget.close
	end function

	' /admin/ordermaster/oneitembuylist.asp
	public Sub GetOneItemOrderListNotPaging()
		dim sqlStr,i, AddSql

		if itemid="" or isnull(itemid) then exit Sub

		AddSql=""
		if FRectStartDate<>"" and FRectEndDate<>"" then
			if (frectdatetype="ipkum") then
				if FRectStartDate<>"" then
					AddSql = AddSql & " and m.ipkumdate>='"& FRectStartDate &"'"
				end if
				if FRectEndDate<>"" then
					AddSql = AddSql & " and m.ipkumdate<'"& FRectEndDate &"'"
				end if
			elseif (frectdatetype="beasong") then
				if FRectStartDate<>"" then
					AddSql = AddSql & " and d.beasongdate>='"& FRectStartDate &"'"
				end if
				if FRectEndDate<>"" then
					AddSql = AddSql & " and d.beasongdate<'"& FRectEndDate &"'"
				end if
			else
				if FRectStartDate<>"" then
					AddSql = AddSql & " and m.regdate>='"& FRectStartDate &"'"
				end if
				if FRectEndDate<>"" then
					AddSql = AddSql & " and m.regdate<'"& FRectEndDate &"'"
				end if
			end if
		end if
		if (frectinccancel <> "Y") then
			AddSql = AddSql & " and m.cancelyn='N'"
			AddSql = AddSql & " and d.cancelyn<>'Y'"
		end if
		if frectitemoption<>"" then
			AddSql = AddSql & " and d.itemoption='" + CStr(frectitemoption) + "'"
		end if
		if frectitemstate<>"" and not(isnull(frectitemstate)) then
			if frectitemstate="2" then   '주문접수
				AddSql = AddSql & " and m.ipkumdiv=2"
			elseif frectitemstate="4" then	'결제완료
				AddSql = AddSql & " and m.ipkumdiv>=4 and m.ipkumdiv<8 and IsNULL(d.currstate,0)=0"
			elseif frectitemstate="6" then	'상품준비/주문통보
				AddSql = AddSql & " and (d.currstate=2 or d.currstate=3)"
			elseif frectitemstate="8" then	'출고완료
				AddSql = AddSql & " and d.currstate=7"
			elseif frectitemstate="9" then	'마이너스
				AddSql = AddSql & " and d.itemno<0"
			elseif frectitemstate="ipkumfinishall" then	'결제완료이상
				AddSql = AddSql & " and m.ipkumdiv>=4"
			end if
		end if
		if frectsitename <> "" then
			AddSql = AddSql & " and m.sitename = '" & CStr(frectsitename) & "' "
		end if
		if frectw10102 <> "" or frectm10102 <> "" or frecta10102 <> "" then
			AddSql = AddSql & " and isnull(m.rdsite,'') in ('" & frectw10102 & "','" & frectm10102 & "','" & frecta10102 & "')"
		end if

		sqlStr = " select top "&FPageSize*FCurrPage
		sqlStr = sqlStr & " m.orderserial, m.ipkumdiv, d.itemno as sm,m.buyname,m.buyemail,m.buyhp,m.buyphone, m.reqname,m.reqhp"
		sqlStr = sqlStr & " ,m.reqphone,d.itemoptionname, IsNULL(d.currstate,0) as currstate, m.sitename, d.beasongdate, m.userid"
		sqlStr = sqlStr & " ,m.jumundiv,d.omwdiv,d.itemcostCouponNotApplied,d.reducedPrice,d.buycash,d.idx,m.regdate,m.rdsite"
		sqlStr = sqlStr & " ,d.itemoption,d.vatinclude, m.userlevel, m.accountdiv"
		sqlStr = sqlStr & " , (case when m.cancelyn='N' and d.cancelyn<>'Y' then 'N' else 'Y' end) as cancelyn"
		sqlStr = sqlStr & " ,d.dlvfinishdt, d.jungsanfixdate "

		if oldlist="on" then
			sqlStr = sqlStr & " from [db_log].[dbo].tbl_old_order_detail_2003 d WITH(NOLOCK) "
			sqlStr = sqlStr & " Join [db_log].[dbo].tbl_old_order_master_2003 m WITH(NOLOCK)"
		else
			sqlStr = sqlStr & " from [db_order].[dbo].tbl_order_detail d WITH(NOLOCK) "
			sqlStr = sqlStr & " Join [db_order].[dbo].tbl_order_master m WITH(NOLOCK)"
		end if

		sqlStr = sqlStr & " 	on m.orderserial=d.orderserial"
		sqlStr = sqlStr & " where m.ipkumdiv>1"
		sqlStr = sqlStr & " and d.itemid="& itemid &" " & AddSql

		''(oa:주문정, od주문역, ra:유입정, rd:유입역, la: 등급정, ld: 등급역, ca:수량정, cd:수량역)
		Select Case sortType
			Case "oa"
				sqlStr = sqlStr & " order by m.orderserial asc "
			Case "od"
				sqlStr = sqlStr & " order by m.orderserial desc "
			Case "ra"
				sqlStr = sqlStr & " order by m.rdsite asc "
			Case "rd"
				sqlStr = sqlStr & " order by m.rdsite desc "
			Case "la"
				sqlStr = sqlStr & " order by m.userlevel asc "
			Case "ld"
				sqlStr = sqlStr & " order by m.userlevel desc "
			Case "ca"
				sqlStr = sqlStr & " order by d.itemno asc "
			Case "cd"
				sqlStr = sqlStr & " order by d.itemno desc "
			Case else
				sqlStr = sqlStr & " order by m.orderserial desc "
		end Select

		'response.write sqlStr & "<br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/06

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		i=0
		if  not rsget.EOF  then
			fArrLIst = rsget.getrows()
		end if

		rsget.Close
	end sub

end class

''사이트구분  //'' toDo rdsite 관련 수정필요
Sub Drawsitename(selectboxname, sitename)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str

	response.write "<select name='" & selectboxname & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if sitename ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">전체</option>"								'선택이란 단어가 나오도록.

	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = " select id from [db_partner].[dbo].tbl_partner"
	userquery = userquery + " where 1=1"
	userquery = userquery + " and id <> ''"
	userquery = userquery + " and id is not null"
	userquery = userquery + " and userdiv= '999'"
	userquery = userquery + " group by id"

	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(sitename) = Lcase(rsget("id")) then 	'검색될 이름과 db에 저장된 이름을 비교해서 맞다면, //
				tem_str = " selected"								'// 검색어로 선택
			else
				tem_str = ""
			end if

			response.write "<option value='" & rsget("id") & "' " & tem_str & ">" & rsget("id") & "</option>"
			rsget.movenext
		loop
	end if
	rsget.close

	tem_str = ""
	if Lcase(sitename) = Lcase("mobileAll") then
	    tem_str = " selected"
	end if
	response.write "<option value='mobileAll' " & tem_str & ">모바일</option>"

	'if (sitename<>"") and (tem_str="") then ''2014/06/23 추가????? 뭐에 필요한것인지 확인 필요(2014-06-30; 허진원)
	'    response.write "<option value='"&sitename&"' selected >"&sitename&"</option>"
	'end if

	response.write "</select>"
End Sub


function NullOrCurrFormat(oval)
    If IsNULL(oval) then
        NullOrCurrFormat = " "
    else
        NullOrCurrFormat = FormatNumber(oval,0)
    end if
end function


Function DefaultSettingWeek()
	Dim vDate
	vDate = DateAdd("ww",-12,now())
	If DatePart("w",vDate) = "1" Then
		DefaultSettingWeek = vDate
	Else
		DefaultSettingWeek = DateAdd("d",((CInt(DatePart("w",vDate))-1)*-1),vDate)
	End If
End Function


Function DateColorSetting(d)
	If DatePart("w",d) = "1" Then
		DateColorSetting = "<font color=""red"">" & d & "</font>"
	ElseIf DatePart("w",d) = "7" Then
		DateColorSetting = "<font color=""blue"">" & d & "</font>"
	Else
		DateColorSetting = d
	End IF
End Function

Function fnChannelDiv(a)
	Dim vBody
	SELECT CASE a
		CASE "web"
			vBody = "'10x10','criteo','naver','naver.','naverM'"
		CASE "jaehu"
			vBody = "'gifticon_web','okcashbag','tworld','cjmall','interpark','lotteCom','lotteimall','giftting'," & _
					"'nvshop_boxA1','nvshop_boxA2','nvshop_boxlogo','nvshop_cast1','nvshop_cast2','nvshop_castleft','nvshop_castright','nvshop_exhibition'," & _
					"'nvshop_logo','nvshop_logo2','nvshop_luckmain','nvshop_lucksub','nvshop_mainb','nvshop_mens','nvshop_pb','nvshop_sp','nvshop_sticb'"
		CASE "mjaehu"
			vBody = "'gifticon_mob','giftting_mob'," & _
					"'mobile_nvshop_boxA1','mobile_nvshop_boxA2','mobile_nvshop_boxlogo','mobile_nvshop_cast1','mobile_nvshop_cast2','mobile_nvshop_castleft'," & _
					"'mobile_nvshop_castright','mobile_nvshop_exhibition','mobile_nvshop_logo','mobile_nvshop_logo2','mobile_nvshop_luckmain','mobile_nvshop_lucksub'," & _
					"'mobile_nvshop_mainb','mobile_nvshop_mens','mobile_nvshop_pb','mobile_nvshop_sp','mobile_nvshop_sticb'"
		CASE "mobile"
			vBody = "'mobile','mobile_adam','mobile_between','mobile_kakaotalk','mobile_kakaotms','mobile_naverM'"
		CASE "ipjum"
			vBody = "'dnshop','gabangpop','gseshop','itsCjmall','privia','shinsegae','wconcept','wizwid'"
		CASE "etc"
			vBody = "'empas','kbcard','KGinicis','mobile_criteo','mobile_nate','mobile_naver','nate','yahoo','11stITS','29cm','bandinlunis','byulshopITS','cjmallITS','cn10x10'," & _
					"'coupang','fashionplus','GVG','hiphoper','hottracks','its29cm','itsByulshop','itsDnshop','itsFashionplus','itsGabangpop','itsGsshop','itsHiphoper','itsHottracks'," & _
					"'itsMusinsa','itsPlayer1','itsShinsegae','itsWconcept','itsWizwid','musinsaITS','NJOYNY','player','suhaITS'"
	END SELECT
	fnChannelDiv = vBody
End Function
%>
