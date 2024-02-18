<%

Class CDelayTaxItem
	public Fyyyymm

	public Fttl
	public FttlCnt

	public FarrTrPrice()
	public FarrTrCnt()

	public FtrNullPrice
	public FtrNullCnt

	public FtrErrPrice
	public FtrErrCnt

	Public Function RedimArrItem(newsize)
		redim preserve FarrTrPrice(newsize)
		redim preserve FarrTrCnt(newsize)
	End Function


	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end Class

Class CDelayTaxDetailItem
	public Fyyyymm
	public Fmakerid
	public Feserotaxkey
	public Fipkumdate
	public Ftaxregdate
	public FjungsanPrice

	public Fpurchasetype
	public FpurchasetypeName

	public FbizsectionName
	public FselltypeName

	public Ffinishflag
	public fjungsan_hp
	public Fcompany_name
	public Fgroupid
	public Ferpcust_cd
	public Fjungsan_gubun
	public Ftaxinputdate
	public FitemvatYn
	public Fjgubun
	public FtaxType

	public function GetFinishFlagName

		if (Ffinishflag = "1") then
			GetFinishFlagName = "업체확인중"
		elseif (Ffinishflag = "3") then
			GetFinishFlagName = "업체확인완료"
		elseif (Ffinishflag = "7") then
			GetFinishFlagName = "완료"
		else
			GetFinishFlagName = Ffinishflag
		end if
	end function

    public function getTaxTypeName
        if (IsCommissionTax) then
            if isNULL(Fitemvatyn) then Exit function

            if (Fitemvatyn="Y") then
                getTaxTypeName = "과세"
            elseif (Fitemvatyn="N") then
                getTaxTypeName = "<font color=red>면세<font>"
            else
                getTaxTypeName = Fitemvatyn
            end if
        else
            if FtaxType="02" then
                getTaxTypeName = "<font color=red>면세<font>"
            elseif FtaxType="01" then
                getTaxTypeName = "과세"
            else
                getTaxTypeName = FtaxType
            end if
        end if
    end function

    public function IsCommissionTax()  ''수수료 매출 세금 계산서 인지.
        IsCommissionTax = false
        if isNULL(Fjgubun) then Exit function

        IsCommissionTax = (Fjgubun="CC") or (Fjgubun="CE")
    end function

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end Class

Class CDelayTax
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
  public FTot_jungsanPrice

	public FRectGubun
	public FRectIssueGubun

	public FRectStartYYYYMM
	public FRectEndYYYYMM
	public FRectMakerid

	public FRectYYYYMM
	public FRectIssueYYYYMM

	public FRectdesigner
	public FRectGroupid
	public FRectPurchaseType
	public FRectErpCustCD
    public FRectJGubun
    public FRectCompanynoYN

	public FRectJacctcdExists

	public sub GetDelayTaxList()
		dim i, j, sqlStr, innerSql
		dim monthCnt, tmpYYYYMM

		dim nextYYYYMM
		nextYYYYMM = Left(CStr(dateserial(Left(FRectEndYYYYMM,4),Right(FRectEndYYYYMM,2)+1,1)), 7)

		if (FRectGubun = "") then
			FRectGubun = "ON"
		end if

		if (FRectGubun = "OFF") then
			innerSql = " SELECT "
			innerSql = innerSql + "	jm.yyyymm "
			innerSql = innerSql + "	, jm.groupid "
			innerSql = innerSql + "	, jm.makerid "
			innerSql = innerSql + "	, jm.eseroevalseq "
			innerSql = innerSql + "	, jm.taxtype "
			innerSql = innerSql + "	, jm.finishflag "
			innerSql = innerSql + "	, convert(varchar(10),jm.ipkumdate,21) as ipkumdate "
			innerSql = innerSql + "	, convert(varchar(10),jm.taxregdate,21) as taxregdate "
			innerSql = innerSql + "	, jm.tot_jungsanPrice "
			innerSql = innerSql + "	, CASE WHEN (jm.taxregdate is NULL) THEN 'X' "
			innerSql = innerSql + "		WHEN jm.yyyymm=convert(varchar(7),jm.taxregdate,21) THEN '' ELSE 'Y' END as DIFFKey "
			innerSql = innerSql + " FROM db_jungsan.dbo.tbl_off_jungsan_master jm "
			innerSql = innerSql + " WHERE "
			innerSql = innerSql + " 1 = 1 "
			innerSql = innerSql + " and jm.yyyymm>='" + CStr(FRectStartYYYYMM) + "' "
			innerSql = innerSql + " and jm.yyyymm<'" + CStr(nextYYYYMM) + "' "
			innerSql = innerSql + " and makerid<>'' "
			innerSql = innerSql + " and Not (finishFlag=0 and jm.tot_jungsanPrice=0) "
			innerSql = innerSql + " and jm.taxtype<>'03' "								'// 영세 제외
		elseif (FRectGubun = "ETC") then
			innerSql = " SELECT "
			innerSql = innerSql + "		jm.yyyymm "
			innerSql = innerSql + "		, jm.shopid "
			innerSql = innerSql + "		, jm.eserotaxkey "
			innerSql = innerSql + "		, jm.papertype "
			innerSql = innerSql + "		, jm.statecd "
			innerSql = innerSql + "		, convert(varchar(10),jm.ipkumdate,21) as ipkumdate "
			innerSql = innerSql + "		, convert(varchar(10),jm.taxdate,21) as taxregdate "
			innerSql = innerSql + "		, jm.totalsum as tot_jungsanPrice "
			innerSql = innerSql + "		, CASE WHEN (jm.taxdate is NULL) THEN 'X' "
			innerSql = innerSql + "			WHEN jm.yyyymm=convert(varchar(7),jm.taxdate,21) THEN '' ELSE 'Y' END as DIFFKey "
			innerSql = innerSql + "	FROM "
			innerSql = innerSql + "		[db_shop].[dbo].tbl_fran_meachuljungsan_master jm "
			innerSql = innerSql + "	WHERE "
			innerSql = innerSql + "		1 = 1 "
			innerSql = innerSql + "		and jm.yyyymm>='" + CStr(FRectStartYYYYMM) + "' "
			innerSql = innerSql + "		and jm.yyyymm<'" + CStr(nextYYYYMM) + "' "
			innerSql = innerSql + "		and Not (statecd = 0) "
			innerSql = innerSql + "		and jm.papertype in ('100', '101') "			'// 영세 수출신고필증 제외
		else
			'// 나머지 = ON
			innerSql = " SELECT "
			innerSql = innerSql + "		jm.yyyymm "
			innerSql = innerSql + "			, jm.groupid "
			innerSql = innerSql + "			, jm.designerid "
			innerSql = innerSql + "			, jm.eseroevalseq "
			innerSql = innerSql + "			, jm.taxtype "
			innerSql = innerSql + "			, jm.finishflag "
			innerSql = innerSql + "			, convert(varchar(10),jm.ipkumdate,21) as ipkumdate "
			innerSql = innerSql + "			, convert(varchar(10),jm.taxregdate,21) as taxregdate "
			innerSql = innerSql + "			, jm.ub_totalsuplycash+jm.me_totalsuplycash+jm.wi_totalsuplycash+jm.et_totalsuplycash+jm.dlv_totalsuplycash as tot_jungsanPrice "
			innerSql = innerSql + "			, CASE WHEN (jm.taxregdate is NULL) THEN 'X' "
			innerSql = innerSql + "				WHEN jm.yyyymm=convert(varchar(7),jm.taxregdate,21) THEN '' ELSE 'Y' END as DIFFKey "
			innerSql = innerSql + "		FROM "
			innerSql = innerSql + "			db_jungsan.dbo.tbl_designer_jungsan_master jm "
			innerSql = innerSql + "		WHERE "
			innerSql = innerSql + "			1 = 1 "
			innerSql = innerSql + "			and jm.yyyymm>='" + CStr(FRectStartYYYYMM) + "' "
			innerSql = innerSql + "			and jm.yyyymm<'" + CStr(nextYYYYMM) + "' "
			innerSql = innerSql + "			and designerid<>'' "
			innerSql = innerSql + "			and Not (finishFlag=0 and jm.ub_totalsuplycash+jm.me_totalsuplycash+jm.wi_totalsuplycash+jm.et_totalsuplycash+jm.dlv_totalsuplycash=0) "
			innerSql = innerSql + "			and jm.taxtype<>'03' "						'// 영세 제외
		end if

		monthCnt = DateDiff("m", FRectStartYYYYMM + "-01", nextYYYYMM + "-01")

		sqlStr = " SELECT "
		sqlStr = sqlStr + " 	yyyymm "
		sqlStr = sqlStr + " 	, sum(tot_jungsanPrice) as ttl "
		sqlStr = sqlStr + " 	, count(*) as ttlCnt "

		tmpYYYYMM = FRectStartYYYYMM
		for i = 0 to monthCnt - 1
			sqlStr = sqlStr + " 	, sum(CASE WHEN convert(varchar(7),taxregdate,21)='" + CStr(tmpYYYYMM) + "' THEN tot_jungsanPrice else 0 end) as [TrPrice " + CStr(tmpYYYYMM) + "] "
			sqlStr = sqlStr + " 	, sum(CASE WHEN convert(varchar(7),taxregdate,21)='" + CStr(tmpYYYYMM) + "' THEN 1 else 0 end) as [TrCnt " + CStr(tmpYYYYMM) + "] "

			tmpYYYYMM = Left(CStr(dateserial(Left(tmpYYYYMM,4),Right(tmpYYYYMM,2)+1,1)), 7)
		next

		sqlStr = sqlStr + " 	, sum(CASE WHEN taxregdate is NULL THEN tot_jungsanPrice else 0 end) as trNullPrice "
		sqlStr = sqlStr + " 	, sum(CASE WHEN taxregdate is NULL THEN 1 else 0 end) as trNullCnt "
		sqlStr = sqlStr + " 	, sum(CASE WHEN (convert(varchar(7),taxregdate,21)<yyyymm) THEN tot_jungsanPrice else 0 end) as trErrPrice "
		sqlStr = sqlStr + " 	, sum(CASE WHEN (convert(varchar(7),taxregdate,21)<yyyymm) THEN 1 else 0 end) as trErrCnt "
		sqlStr = sqlStr + " FROM ( "

		sqlStr = sqlStr + innerSql

		sqlStr = sqlStr + " ) T "
		sqlStr = sqlStr + " GROUP BY yyyymm "
		sqlStr = sqlStr + " ORDER BY yyyymm "

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				''rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CDelayTaxItem

					FItemList(i).RedimArrItem(monthCnt)

					FItemList(i).Fyyyymm 		= rsget("yyyymm")
					FItemList(i).Fttl 			= rsget("ttl")
					FItemList(i).FttlCnt 		= rsget("ttlCnt")

					tmpYYYYMM = FRectStartYYYYMM
					for j = 0 to monthCnt - 1
						FItemList(i).FarrTrPrice(j) 	= rsget("TrPrice " + CStr(tmpYYYYMM))
						FItemList(i).FarrTrCnt(j) 		= rsget("TrCnt " + CStr(tmpYYYYMM))

						tmpYYYYMM = Left(CStr(dateserial(Left(tmpYYYYMM,4),Right(tmpYYYYMM,2)+1,1)), 7)
					next

					FItemList(i).FtrNullPrice 	= rsget("trNullPrice")
					FItemList(i).FtrNullCnt 	= rsget("trNullCnt")
					FItemList(i).FtrErrPrice 	= rsget("trErrPrice")
					FItemList(i).FtrErrCnt 		= rsget("trErrCnt")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub


	public sub GetDelayTaxDetailList()
		dim i, j, sqlStr, innerSql
		dim maxCount
		maxCount = 3000
		dim nextYYYYMM
		nextYYYYMM = Left(CStr(dateserial(Left(FRectEndYYYYMM,4),Right(FRectEndYYYYMM,2)+1,1)), 7)

		if (FRectGubun = "") then
			FRectGubun = "ON"
		end if

		if (FRectGubun = "OFF") then
			innerSql = " SELECT TOP " + Cstr(maxCount)
			innerSql = innerSql + "		jm.yyyymm "
			innerSql = innerSql + "		, jm.groupid "
			innerSql = innerSql + "		, jm.makerid "
			innerSql = innerSql + "		, jm.eseroevalseq as eserotaxkey "
			innerSql = innerSql + "		, jm.taxtype "
			innerSql = innerSql + "		, jm.finishflag "
			innerSql = innerSql + "		, convert(varchar(10),jm.ipkumdate,21) as ipkumdate "
			innerSql = innerSql + "		, convert(varchar(10),jm.taxregdate,21) as taxregdate "
			innerSql = innerSql + "		, jm.tot_jungsanPrice "
			innerSql = innerSql + "		, CASE WHEN (jm.taxregdate is NULL) THEN 'X' "
			innerSql = innerSql + "			WHEN jm.yyyymm=convert(varchar(7),jm.taxregdate,21) THEN '' ELSE 'Y' END as DIFFKey "
			innerSql = innerSql + "		, p.purchasetype, p.jungsan_hp "
			innerSql = innerSql + "		, pc.pcomm_name as purchasetypeName "
			innerSql = innerSql + "		, '' as bizsectionName "
			innerSql = innerSql + "		, '' as selltypeName "
			innerSql = innerSql + "		, pg.company_name "
			''innerSql = innerSql + "		, isNull(pg.erpcust_cd,jm.groupid) as erpcust_cd "
			innerSql = innerSql + "		, isNull(b.CUST_USE_CD,jm.groupid) as erpcust_cd "
			innerSql = innerSql + "		, pg.jungsan_gubun "
			innerSql = innerSql + "		, jm.taxinputdate "
			innerSql = innerSql + "		, jm.taxtype,  jm.itemvatYn, jm.jgubun "
			innerSql = innerSql + " FROM db_jungsan.dbo.tbl_off_jungsan_master jm "
			innerSql = innerSql + "	left join [db_partner].[dbo].tbl_partner p "
			innerSql = innerSql + "	on "
			innerSql = innerSql + "		jm.makerid=p.id "
			innerSql = innerSql + "	left join [db_partner].[dbo].tbl_partner_comm_code pc "
			innerSql = innerSql + "	on "
			innerSql = innerSql + "		1 = 1 "
			innerSql = innerSql + "		and pc.pcomm_group = 'purchasetype' "
			innerSql = innerSql + "		and pc.pcomm_cd=p.purchasetype "
			innerSql = innerSql + " left join  [db_partner].[dbo].tbl_partner_group as pg "
			innerSql = innerSql + "	on "
			innerSql = innerSql + "	jm.groupid = pg.groupid "
			innerSql = innerSql + " left join db_partner.dbo.tbl_TMS_BA_CUST b"
	        innerSql = innerSql + " on pg.erpCust_cd=b.CUST_CD"
			innerSql = innerSql + " WHERE "
			innerSql = innerSql + " 	 jm.yyyymm>='" + CStr(FRectStartYYYYMM) + "' "
			innerSql = innerSql + " and jm.yyyymm<'" + CStr(nextYYYYMM) + "' "

			innerSql = innerSql + " 	and makerid<>'' "
			innerSql = innerSql + " 	and Not (finishFlag=0 and jm.tot_jungsanPrice=0) "
			innerSql = innerSql + " 	and jm.taxtype<>'03' "								'// 영세 제외

			if (FRectIssueGubun = "1") then
				'// 정상발행
				innerSql = innerSql + " 	and jm.taxregdate is not NULL "
				innerSql = innerSql + " 	and convert(varchar(7),jm.taxregdate,21) = '" + CStr(FRectIssueYYYYMM) + "' "
			elseif (FRectIssueGubun = "2") then
				'// 발행이전
				innerSql = innerSql + " 	and IsNull(jm.taxregdate, '') = '' "
				'innerSql = innerSql + " 	and IsNull(jm.finishflag, '') <> 7 "		'2017-07-28 김진영 추가
			else
				'// 기타발행(선발행)
				innerSql = innerSql + " 	and jm.taxregdate is not NULL "
				innerSql = innerSql + " 	and convert(varchar(7),jm.taxregdate,21) < jm.yyyymm "
			end if

			if FRectdesigner <> "" THEN
					innerSql = innerSql + " 	and jm.makerid='"+FRectdesigner+"' "
			END IF
			IF FRectGroupid <> "" THEN
					innerSql = innerSql + " 	and  jm.groupid ='"+FRectGroupid+"' "
			END IF
 			IF FRectPurchaseType <> "" THEN
 					innerSql = innerSql + " 	and   p.purchasetype = "+Cstr(FRectPurchaseType)
 			END IF
			IF FRectErpCustCD <> "" THEN
					innerSql = innerSql + " 	and   b.CUST_USE_CD = '"+Cstr(FRectErpCustCD)+"'"
			END IF
            IF FRectJGubun <> "" THEN
					innerSql = innerSql + " 	and   jm.jgubun = '"+Cstr(FRectJGubun)+"'"
			END IF
			If FRectCompanynoYN <> "" Then
				Select Case FRectCompanynoYN
					Case "Y"		innerSql = innerSql & " and isnull(replace(p.company_no,'-',''), '') = '2118700620'  "
					Case "N"		innerSql = innerSql & " and isnull(replace(p.company_no,'-',''), '') <> '2118700620'  "
				End Select
			End If
			
			if (FRectJacctcdExists<>"") then
				innerSql = innerSql + " and isNULL(jm.jacctcd,'')<>''"
			end if

			innerSql = innerSql + "	order by "
			innerSql = innerSql + "		jm.tot_jungsanPrice desc "
		elseif (FRectGubun = "ETC") then
			innerSql = " SELECT TOP  " + Cstr(maxCount)
			innerSql = innerSql + "		jm.yyyymm "
			innerSql = innerSql + "		,'' as groupid "
			innerSql = innerSql + "		, jm.shopid as makerid "
			innerSql = innerSql + "		, jm.eserotaxkey "
			innerSql = innerSql + "		, jm.papertype "
			innerSql = innerSql + "		, jm.statecd as finishflag "
			innerSql = innerSql + "		, convert(varchar(10),jm.ipkumdate,21) as ipkumdate "
			innerSql = innerSql + "		, convert(varchar(10),jm.taxdate,21) as taxregdate "
			innerSql = innerSql + "		, jm.totalsum as tot_jungsanPrice "
			innerSql = innerSql + "		, CASE WHEN (jm.taxdate is NULL) THEN 'X' "
			innerSql = innerSql + "			WHEN jm.yyyymm=convert(varchar(7),jm.taxdate,21) THEN '' ELSE 'Y' END as DIFFKey "
			innerSql = innerSql + "		, '' as purchasetype, '' as jungsan_hp "
			innerSql = innerSql + "		, '' as purchasetypeName "
			innerSql = innerSql + "		, m.bizsection_nm as bizsectionName "
			innerSql = innerSql + "		, c.pcomm_name as selltypeName "
			innerSql = innerSql + "		, '' as company_name "
			innerSql = innerSql + "		, '' as erpcust_cd "
			innerSql = innerSql + "		, '' as jungsan_gubun "
			innerSql = innerSql + "		, jm.taxregdate as taxinputdate "
			innerSql = innerSql + "		, '' as taxtype "
			innerSql = innerSql + "		, '' as itemvatYn "
			innerSql = innerSql + "		, '' as jgubun "
			innerSql = innerSql + "	FROM "
			innerSql = innerSql + "		[db_shop].[dbo].tbl_fran_meachuljungsan_master jm "
			innerSql = innerSql + "		left join db_partner.dbo.tbl_TMS_BA_BIZSECTION m "
			innerSql = innerSql + "		on "
			innerSql = innerSql + "			m.bizsection_cd = jm.bizsection_cd "
			innerSql = innerSql + "		left join [db_partner].[dbo].tbl_partner_comm_code c "
			innerSql = innerSql + "		on "
			innerSql = innerSql + "			jm.selltype = c.pcomm_cd and c.pcomm_group = 'sellacccd' "
			innerSql = innerSql + "	WHERE "
			innerSql = innerSql + " 	 jm.yyyymm>='" + CStr(FRectStartYYYYMM) + "' "
			innerSql = innerSql + " 	and jm.yyyymm<'" + CStr(nextYYYYMM) + "' "
			innerSql = innerSql + "		and Not (statecd = 0) "
			innerSql = innerSql + "		and jm.papertype in ('100', '101') "			'// 영세 수출신고필증 제외

			if (FRectIssueGubun = "1") then
				'// 정상발행
				innerSql = innerSql + " 	and jm.taxdate is not NULL "
				innerSql = innerSql + " 	and convert(varchar(7),jm.taxdate,21) = '" + CStr(FRectIssueYYYYMM) + "' "
			elseif (FRectIssueGubun = "2") then
				'// 발행이전
				innerSql = innerSql + " 	and IsNull(jm.taxdate, '') = '' "
				'innerSql = innerSql + " 	and IsNull(jm.finishflag, '') <> 7 "		'2017-07-28 김진영 추가
			else
				'// 기타발행(선발행)
				innerSql = innerSql + " 	and jm.taxdate is not NULL "
				innerSql = innerSql + " 	and convert(varchar(7),jm.taxdate,21) < jm.yyyymm "
			end if
			if FRectdesigner <> "" THEN
					innerSql = innerSql + " 	and jm.shopid='"+FRectdesigner+"' "
			END IF
			IF FRectJGubun <> "" THEN
					innerSql = innerSql + " 	and   jm.jgubun = '"+Cstr(FRectJGubun)+"'"
			END IF

			If FRectCompanynoYN <> "" Then
				Select Case FRectCompanynoYN
					Case "Y"		innerSql = innerSql & " and (Left(jm.shopid, 10) = 'streetshop' OR Left(jm.shopid, 9) = 'wholesale')  "
					Case "N"		innerSql = innerSql & " and (Left(jm.shopid, 10) <> 'streetshop' AND Left(jm.shopid, 9) <> 'wholesale')  "
				End Select
			End If


			innerSql = innerSql + "	order by "
			innerSql = innerSql + "		jm.totalsum desc "
	'rw innerSql
	'response.end
		else
			'// 나머지 = ON
			innerSql = " SELECT TOP  " + Cstr(maxCount)
			innerSql = innerSql + "		jm.yyyymm "
			innerSql = innerSql + "			, jm.groupid "
			innerSql = innerSql + "			, jm.designerid as makerid "
			innerSql = innerSql + "			, jm.eseroevalseq as eserotaxkey "
			innerSql = innerSql + "			, jm.taxtype "
			innerSql = innerSql + "			, jm.finishflag "
			innerSql = innerSql + "			, convert(varchar(10),jm.ipkumdate,21) as ipkumdate "
			innerSql = innerSql + "			, convert(varchar(10),jm.taxregdate,21) as taxregdate "
			innerSql = innerSql + "			, jm.ub_totalsuplycash+jm.me_totalsuplycash+jm.wi_totalsuplycash+jm.et_totalsuplycash+jm.dlv_totalsuplycash as tot_jungsanPrice "
			innerSql = innerSql + "			, CASE WHEN (jm.taxregdate is NULL) THEN 'X' "
			innerSql = innerSql + "				WHEN jm.yyyymm=convert(varchar(7),jm.taxregdate,21) THEN '' ELSE 'Y' END as DIFFKey "
			innerSql = innerSql + "			, p.purchasetype, p.jungsan_hp "
			innerSql = innerSql + "			, pc.pcomm_name as purchasetypeName "
			innerSql = innerSql + "			, '' as bizsectionName "
			innerSql = innerSql + "			, '' as selltypeName "
			innerSql = innerSql + "		, pg.company_name "
			''innerSql = innerSql + "		, isNull(pg.erpcust_cd,jm.groupid) as erpcust_cd  "
			innerSql = innerSql + "		, isNull(b.CUST_USE_CD,jm.groupid) as erpcust_cd "
			innerSql = innerSql + "		, pg.jungsan_gubun "
			innerSql = innerSql + "		, jm.taxinputdate "
			innerSql = innerSql + "		, jm.taxtype,  jm.itemvatYn, jm.jgubun "
			innerSql = innerSql + "		FROM "
			innerSql = innerSql + "			db_jungsan.dbo.tbl_designer_jungsan_master jm "
			innerSql = innerSql + "		left join [db_partner].[dbo].tbl_partner p "
			innerSql = innerSql + "		on "
			innerSql = innerSql + "			jm.designerid=p.id "
			innerSql = innerSql + "		left join [db_partner].[dbo].tbl_partner_comm_code pc "
			innerSql = innerSql + "		on "
			innerSql = innerSql + "			1 = 1 "
			innerSql = innerSql + "			and pc.pcomm_group = 'purchasetype' "
			innerSql = innerSql + "			and pc.pcomm_cd=p.purchasetype "
			innerSql = innerSql + " 	left join  [db_partner].[dbo].tbl_partner_group as pg "
			innerSql = innerSql + "		on "
			innerSql = innerSql + "		jm.groupid = pg.groupid "
			innerSql = innerSql + " left join db_partner.dbo.tbl_TMS_BA_CUST b"
	        innerSql = innerSql + " on pg.erpCust_cd=b.CUST_CD"
			innerSql = innerSql + "		WHERE "
			innerSql = innerSql + " 	 jm.yyyymm>='" + CStr(FRectStartYYYYMM) + "' "
			innerSql = innerSql + " 		and jm.yyyymm<'" + CStr(nextYYYYMM) + "' "
			innerSql = innerSql + "			and designerid<>'' "
			innerSql = innerSql + "			and Not (finishFlag=0 and jm.ub_totalsuplycash+jm.me_totalsuplycash+jm.wi_totalsuplycash+jm.et_totalsuplycash+jm.dlv_totalsuplycash=0) "
			innerSql = innerSql + "			and jm.taxtype<>'03' "						'// 영세 제외

			if (FRectIssueGubun = "1") then
				'// 정상발행
				innerSql = innerSql + " 	and jm.taxregdate is not NULL "
				innerSql = innerSql + " 	and convert(varchar(7),jm.taxregdate,21) = '" + CStr(FRectIssueYYYYMM) + "' "
			elseif (FRectIssueGubun = "2") then
				'// 발행이전
				innerSql = innerSql + " 	and IsNull(jm.taxregdate, '') = '' "
				'innerSql = innerSql + " 	and IsNull(jm.finishflag, '') <> 7 "		'2017-07-28 김진영 추가
			else
				'// 기타발행(선발행)
				innerSql = innerSql + " 	and jm.taxregdate is not NULL "
				innerSql = innerSql + " 	and convert(varchar(7),jm.taxregdate,21) < jm.yyyymm "
			end if
			if FRectdesigner <> "" THEN
					innerSql = innerSql + " 	and jm.designerid='"+FRectdesigner+"' "
			END IF
			IF FRectGroupid <> "" THEN
					innerSql = innerSql + " 	and  jm.groupid ='"+FRectGroupid+"' "
			END IF
 			IF FRectPurchaseType <> "" THEN
 					innerSql = innerSql + " 	and   p.purchasetype = "+Cstr(FRectPurchaseType)
 			END IF
			IF FRectErpCustCD <> "" THEN
					innerSql = innerSql + " 	and   b.CUST_USE_CD = '"+Cstr(FRectErpCustCD)+"'"
			END IF
			IF FRectJGubun <> "" THEN
					innerSql = innerSql + " 	and   jm.jgubun = '"+Cstr(FRectJGubun)+"'"
			END IF

			If FRectCompanynoYN <> "" Then
				Select Case FRectCompanynoYN
					Case "Y"		innerSql = innerSql & " and isnull(replace(p.company_no,'-',''), '') = '2118700620'  "
					Case "N"		innerSql = innerSql & " and isnull(replace(p.company_no,'-',''), '') <> '2118700620'  "
				End Select
			End If

			if (FRectJacctcdExists<>"") then
				innerSql = innerSql + " and isNULL(jm.jacctcd,'')<>''"
			end if
			
			innerSql = innerSql + "	order by "
			innerSql = innerSql + "		(jm.ub_totalsuplycash+jm.me_totalsuplycash+jm.wi_totalsuplycash+jm.et_totalsuplycash+jm.dlv_totalsuplycash) desc "
		end if
'rw innerSql
		sqlStr = innerSql
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				''rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CDelayTaxDetailItem

					FItemList(i).Fyyyymm 		= rsget("yyyymm")
					FItemList(i).Fmakerid 		= rsget("makerid")
					FItemList(i).Feserotaxkey 	= rsget("eserotaxkey")
					FItemList(i).Fipkumdate 	= rsget("ipkumdate")
					FItemList(i).Ftaxregdate 	= rsget("taxregdate")
					FItemList(i).FjungsanPrice 	= rsget("tot_jungsanPrice")

					FItemList(i).Fpurchasetype 		= rsget("purchasetype")
					FItemList(i).FpurchasetypeName 	= rsget("purchasetypeName")

					FItemList(i).FbizsectionName 	= rsget("bizsectionName")
					FItemList(i).FselltypeName 		= rsget("selltypeName")

					FItemList(i).Ffinishflag 	= rsget("finishflag")

					FItemList(i).Fgroupid					= rsget("groupid")
					FItemList(i).Fcompany_name		= rsget("company_name")
					FItemList(i).FerpCust_CD			= rsget("erpCust_CD")
					FItemList(i).Fjungsan_gubun		= rsget("jungsan_gubun")
					FItemList(i).Ftaxinputdate  = rsget("taxinputdate")

				    FTot_jungsanPrice =  FTot_jungsanPrice + FItemList(i).FjungsanPrice
				    FItemList(i).Ftaxtype  = rsget("taxtype")
				    FItemList(i).FitemvatYn  = rsget("itemvatYn")
				    FItemList(i).Fjgubun        = rsget("jgubun")
					FItemList(i).fjungsan_hp        = rsget("jungsan_hp")

					i=i+1
					rsget.moveNext
				loop

			end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end Class

%>
