<%
Class CUpcheJungsanSummaryItem
    public Fjgubun
    public FtargetGbn
    public Fsitename
    public FitemvatYn
    public Fgubuncd
    ''public FcommissionSum           ''기존 삭제
    public FsellcashSum
    public FreducedpriceSum
    public FsuplycashSum
    public FsitenameName
    public Fcomm_name

    public FttlcommissionSum        ''총수수료
    public FitemcommissionSum       ''상품수수료
    public FpgcommissionSum         ''PG수수료

    public function getJGubunName
        if (FjGubun="MM") then
            getJGubunName = "매입"
        elseif (FjGubun="CC") then
            getJGubunName = "<font color=blue>수수료</font>"
        elseif (FjGubun="CE") then
            getJGubunName = "기타"
        else
            getJGubunName = FjGubun
        end if
    end function

    public function getItemVatTypeName()
        if FitemvatYn="N" then
            getItemVatTypeName = "<font color=red>면세<font>"
        elseif FitemvatYn="Y" then
            getItemVatTypeName = "과세"
        else
            getItemVatTypeName = FitemvatYn
        end if
    end function

    Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CUpcheJungsanTaxItem
    public Fid
    public FtargetGbn
    public Fjgubun
    public Fmakerid
    public Fgroupid
    public Fyyyymm
    public Ftitle

    public Fregdate
    public Ffinishflag
    public Fipkumdate
    public Ftaxregdate

    public Fdifferencekey
    public Ftaxtype
    public FTaxLinkidx
    public Fneotaxno
    public FBillsiteCode


    public FeseroEvalSeq
    public FbillSiteName

    public Fcompany_name
    public Fjungsan_gubun
    public Ftaxinputdate
    public Fcompany_no

    public Fbankingupflag

    public FtotalJungsanSum
    public Ftotalcommission
    public FDlvPgCommission

    public FPrdMeachulsum
    public FPrdCommissionSum
    public FdlvMeachulsum
    public FetMeachulsum

    public FprdJungsanSum
    public FdlvJungsanSum
    public FetJungsanSum

    public Fjungsan_date
    public Fjungsan_date_off

    public FitemvatYn

    public FerpCust_cd
    public FerpUsing

    public FSSuply
    public FCSuply
    public FMSuply
    public FDSuply
    public FESuply

    public Fjacctcd
    public Facc_nm

    public FPgCommissionSum

	Public FDesignerid
	Public FMastercode
	Public FSitename
	Public FBuyname
	Public FReqname
	Public FItemid
	Public FItemoption
	Public FItemname
	Public FItemoptionname
	Public FItemno
	Public FSellcash
	Public FCouponPlusCommi
	Public FCoupoonDiscount
	Public FReducedprice
	Public FCommission
	Public FPgcommission
	Public FSuplycash
	Public FSumsuplycash
	Public FPaymethod
    Public Fauthcode


    ''배송비(지급예정액) 추가배송비 제외, 반품배송비 제외, 기타프로모션 제외
    public function getOriginDlvJungsanSum()
        ''getOriginDlvJungsanSum = FdlvMeachulsum  ''2016/09/29 수정
        getOriginDlvJungsanSum = (FdlvMeachulsum-FDlvPgCommission)
    end function

    ''추가 배송비 : 판매액<>정산액(판매액 0 배송비 2500, 판매액 2500 정산액 3000)
    public function getAddDlvJungsanSum()
        ''getAddDlvJungsanSum = (FdlvJungsanSum-FdlvMeachulsum)-getEtcDlvJungsanSum-getPromotionJungsanSum
        getAddDlvJungsanSum = (FdlvJungsanSum-(FdlvMeachulsum-FDlvPgCommission))-getEtcDlvJungsanSum-getPromotionJungsanSum
    end function

    ''반품 배송비 등(DT)
    public function getEtcDlvJungsanSum()
        getEtcDlvJungsanSum = FCSuply
    end function

    ''기타 프로모션 비용 등 (DP)
    public function getPromotionJungsanSum()
        getPromotionJungsanSum = FESuply
    end function

    public function getCalcuToTalJungsanSum()
        getCalcuToTalJungsanSum = FSSuply+FCSuply+FMSuply+FDSuply+FESuply
    end function

    public function getDbDate()
		dim sqlstr
		sqlstr = " select convert(varchar(10),getdate(),21) as nowdate "
		rsget.Open sqlStr,dbget,1
		getDbDate = CDate(rsget("nowdate"))
		rsget.Close
	end function

   public function GetNormalTaxDate()
		''if Not(IsNULL(FpreFixedTaxDate)) and (FpreFixedTaxDate<>"") then
		''	GetNormalTaxDate = FpreFixedTaxDate
		''else
				GetNormalTaxDate = dateserial(Left(Fyyyymm,4),Right(Fyyyymm,2)+1,1-1)
		''end if
	end function

    public function GetPreFixSegumil()
        dim thisdate, maytaxdate
		dim ithis1day , ithis21day, premonth1day, premonth21day

		thisdate = getDbDate()
		maytaxdate = GetNormalTaxDate()

        '' 12일까지 마감할 경우 13으로 세팅
        '' 10일까지 마감할 경우 11로 쎄팅
		premonth1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"01")
		premonth21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"12") ''11
		ithis1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"01")
		ithis21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"12") ''11

		''######################################## 2017-09-21 김진영 추가 ########################################
		Dim strSql, taxdate
		strSql = ""
		strSql = strSql & " SELECT TOP 1 isnull(taxdate,'') as taxdate FROM "
		strSql = strSql & " [db_sitemaster].[dbo].[tbl_taxdate_manage] with (nolock)"
		strSql = strSql & " WHERE yyyymm = '"& Left(thisdate, 7) &"' "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
            if rsget("taxdate")<>"" then
			    taxdate = CDate(rsget("taxdate"))
            end if
		End If
		rsget.Close
		''######################################## 2017-09-21 김진영 추가 끝 #######################################

    	'if (thisdate>=ithis21day) then		2017-09-21 김진영 주석, 아래 taxdate 변경 및 부등호 제거
	    if (thisdate > taxdate) then
			GetPreFixSegumil = ithis1day
		elseif (maytaxdate<premonth21day)  then
			GetPreFixSegumil = premonth1day
		else
			GetPreFixSegumil = maytaxdate
		end if

        if (Fid=201998) then
            GetPreFixSegumil = "2014-07-15"
        end if

        if (Fid=208998) then
            GetPreFixSegumil = "2014-09-20"
        end if

        if (Fid=208394) then
            GetPreFixSegumil = "2014-09-20"
        end if

        if (Fid=218357) then
            GetPreFixSegumil = "2014-11-04"
        end if

        if (Fid=225812) or (Fid=227439) or (Fid=227288) or (Fid=227120)  then
            GetPreFixSegumil = "2015-01-01"
        end if

        if (Fid=232953) then
            GetPreFixSegumil = "2015-03-18"
        end if

        if (Fid=234270) then
            GetPreFixSegumil = "2015-03-23"
        end if

        if (Fid=238873) or (Fid=236212) or (Fid=236213) or (Fid=236214)   then
            GetPreFixSegumil = "2015-04-05"
        end if

        if (Fid=238418) then
            GetPreFixSegumil = "2015-04-21"
        end if

        if (Fid=241190) then
            GetPreFixSegumil = "2015-05-21"
        end if

        if (Fid=241921) then
            GetPreFixSegumil = "2015-05-28"
        end if

        if (Fid=244318) then
            GetPreFixSegumil = "2015-06-01"
        end if

        if (Fid=244974) then
            GetPreFixSegumil = "2015-06-22"
        end if

        if (Fid=278483) then
            GetPreFixSegumil = "2016-02-01"
        end if
        if (Fid=281205) then
            GetPreFixSegumil = "2016-02-01"
        end if

        if (Fid=288213) or (Fid=288185) or (Fid=286564)    then
            GetPreFixSegumil = "2016-04-01"
        end if

        if (Fid=88388)    then
            GetPreFixSegumil = "2016-06-30"
        end if

'        ''' 2012-09-03 추가.
'        if Not(IsNULL(FpreFixedTaxDate)) and (FpreFixedTaxDate<>"") then
'            if (thisdate>=ithis21day) and (CStr(FpreFixedTaxDate)<ithis1day) then
'               ''기본 계산값
'            elseif (Left(FpreFixedTaxDate,10)<CStr(premonth1day)) then
'                ''기본 계산값
'            ELSE
'                GetPreFixSegumil = Left(FpreFixedTaxDate,10)
'            end if
'        end if
    end function

    public function getBill_SELL_DAM_DEPT()
        if (FtargetGbn="ON") THEN
            getBill_SELL_DAM_DEPT = "온라인"
        elseif (FtargetGbn="OF") THEN
            getBill_SELL_DAM_DEPT = "오프라인"
        elseif (FtargetGbn="AC") THEN
            getBill_SELL_DAM_DEPT = "더핑거스"
        end if
    end function


    public function getBill_FG_VAT() ''// 1과세,2영세,3면세
        if (IsCommissionTax) then
            getBill_FG_VAT = "1"
        else
            if (Ftaxtype="02") then
                getBill_FG_VAT = "3"
            elseif (Ftaxtype="01") then
                getBill_FG_VAT = "1"
            ''영세 일단 없음.
            end if
        end if
    end function

    public function getBill_NO_SENDER_PK()
        getBill_NO_SENDER_PK = "DZ_TEN_"&ojungsanTaxCC.FOneItem.FtargetGbn&"_"& ojungsanTaxCC.FOneItem.Fid&"_"&ojungsanTaxCC.FOneItem.Fdifferencekey&"_"&ojungsanTaxCC.FOneItem.getJungsanTaxSum
    end function

    public function getBill_NM_ITEM()
        getBill_NM_ITEM = ojungsanTaxCC.FOneItem.getBill_SELL_DAM_DEPT &" "& ojungsanTaxCC.FOneItem.Fmakerid &" "& ojungsanTaxCC.FOneItem.Ftitle
    end function

    public function getBill_FG_BILL ''//청구1 영수2
        if (IsCommissionTax) then
            getBill_FG_BILL = "2"
        else
            getBill_FG_BILL = "1"
        end if
    end function

    public function getToTalJungsanSum()
        getToTalJungsanSum = FprdJungsanSum+FdlvJungsanSum+FetJungsanSum
    end function

    public function getMayIpkumdateStr()
        dim ret
        if (FtargetGbn="OF") then
            ret = Fjungsan_date_off
        else
            ret = Fjungsan_date
        end if
        if isNULL(ret) then ret=""
        if ret="" then ret="말일"
        getMayIpkumdateStr = ret
    end function


    public function getItemVatTypeName()
        if FitemvatYn="N" then
            getItemVatTypeName = "<font color=red>면세<font>"
        elseif FitemvatYn="Y" then
            getItemVatTypeName = "과세"
        else
            getItemVatTypeName = FitemvatYn
        end if
    end function

    public function getTaxTypeName()
        if Ftaxtype="02" then
            getTaxTypeName = "<font color=red>면세<font>"
        elseif Ftaxtype="01" then
            getTaxTypeName = "과세"
        else
            getTaxTypeName = Ftaxtype
        end if
    end function

    public function getTaxEvalStyleStr
        if (IsCommissionTax) then
            getTaxEvalStyleStr = "텐바이텐"
        else
            getTaxEvalStyleStr = "협력사"
        end if
    end function

    ''공급가액
    public function getJungsanTaxSuply()
        getJungsanTaxSuply = getJungsanTaxSum-getJungsanTaxVat
    end function

    ''부과세
    public function getJungsanTaxVat()
        if (IsCommissionTax) or (Ftaxtype="01") then
            getJungsanTaxVat = CLng(getJungsanTaxSum / 11)
        else
            getJungsanTaxVat = 0
        end if
    end function

    ''합계
    public function getJungsanTaxSum()
        if (IsCommissionTax) then
            getJungsanTaxSum = CLNG(Ftotalcommission)
        else
            getJungsanTaxSum = CLNG(FtotalJungsanSum)
        end if
    end function


    public function GetTaxEvalStateName()
		if Ffinishflag="0" then
			GetTaxEvalStateName = "<font color='#000000'>수정중</font>"
		elseif Ffinishflag="1" then
		    if IsCommissionTax then
		        GetTaxEvalStateName = "<font color='#448888'>텐바이텐에서<br>발행예정</font>"
		    else
    		    GetTaxEvalStateName = "<font color='#448888'>업체확인대기</font>"
    		end if
		elseif Ffinishflag="2" then
		    GetTaxEvalStateName = "<font color='#0000FF'>업체확인완료</font>"
		elseif Ffinishflag="3" then
			GetTaxEvalStateName = "<font color='#0000FF'>정산확정</font>"
		elseif Ffinishflag="7" then
			GetTaxEvalStateName = "<font color='#FF0000'>입금완료</font>"
		else

		end if
    end function

    public function IsCommissionTax()  ''수수료 매출 세금 계산서 인지.
        IsCommissionTax = false
        if isNULL(Fjgubun) then Exit function

        IsCommissionTax = (Fjgubun="CC") or (Fjgubun="CE")
    end function

    public function IsCommissionETCTax()  ''기타 매출 세금 계산서 인지.
        IsCommissionETCTax = false
        if isNULL(Fjgubun) then Exit function

        IsCommissionETCTax = (Fjgubun="CE")
    end function

    public function getTaxJungsanGubun
        dim retVal
        if (IsCommissionTax) then
            if (IsCommissionETCTax) then
                retVal = "<font color=red>기타 정산</font>"
            else
                retVal = "<font color=blue>수수료 정산</font>"
            end if
        else
            retVal = "매입 정산"
        end if
        getTaxJungsanGubun = retVal
    end function

    ''업체 기준 계산서 구분
    public function getTaxTypeStrUpcheView()
        if (Ftaxtype="02") then
            getTaxTypeStrUpcheView= "계산서"
        elseif (Ftaxtype="01") then
            getTaxTypeStrUpcheView = "세금계산서"
        elseif (Ftaxtype="03") then
            getTaxTypeStrUpcheView = "간이"
        else
            getTaxTypeStrUpcheView = Ftaxtype
        end if

    end function


    public function getTargetNm()
        if isNULL(FtargetGbn) then Exit function

        if (FtargetGbn="ON") then
            if (NOT IsCommissionTax) and (LEFT(Fyyyymm,4)>="2016") then
                getTargetNm = "전사"           ''2016/01 추가.
            else
                getTargetNm = "온라인"
            end if
        elseif (FtargetGbn="OF") then
            getTargetNm = "오프라인"
        elseif (FtargetGbn="AC") then
            getTargetNm = "더핑거스"
        else
            getTargetNm = FtargetGbn
        end if
    end function

    public function IsElecTaxExists()
		IsElecTaxExists = Not(IsNULL(FTaxLinkidx) or (FTaxLinkidx="")) and (Ffinishflag>=3)
	end function


	''//세금계산서
	public function IsElecTaxCase()
		IsElecTaxCase = (Ftaxtype="01") and (Fjungsan_gubun="일반과세") and (Ffinishflag<3)
	end function


	''//계산서
	public function IsElecFreeTaxCase()
		IsElecFreeTaxCase = (Ftaxtype="02") and (Ffinishflag<3) 'and (Fjungsan_gubun="면세")
	end function

    public function IsEvaledTax()
        IsEvaledTax = (Ffinishflag>=3)
    end function

	''//간이, 원천, 기타
	public function IsElecSimpleBillCase()
		IsElecSimpleBillCase = (Ftaxtype="03") and (Ffinishflag<3)
	end function


    Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
End Class

class CUpcheJungsanTax
    public FItemList()
    public FOneItem
    public FSumaryOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

    public FRectMakerid
    public FRectYYYYMM
    public FRectJGubun  '' MM, CC
    public FRectJjungsanIdx
    public FRectTargetGbn
    public FRectFinishFlag
    public FRectGroupid

    public FRectTaxType
    public FRectItemVatYn
    public FRectJungsanDate
    public FRectjacctcd
    public FRectNotIncTen

	public FRectSearchType
	public FRectSearchText
    public FRectcomm_cd
    public FRectJungsanException

    public Sub getOneUpcheJungsanTax
        Dim sqlStr
        sqlStr = "[db_jungsan].[dbo].[sp_Ten_getOneUpcheJungsanTaxByKey]("&FRectJjungsanIdx&",'"&FRectTargetGbn&"','"&FRectMakerid&"')"
        rsget.CursorLocation = adUseClient
	    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

		FResultCount = rsget.RecordCount
		FTotalCount  = FResultCount

		if  not rsget.EOF  then
		    set FOneItem = new CUpcheJungsanTaxItem
            FOneItem.Fid                = rsget("id")
            FOneItem.Fjgubun            = rsget("jgubun")
			FOneItem.Fmakerid           = rsget("makerid")
			FOneItem.Fgroupid		    = rsget("groupid")
			FOneItem.Fyyyymm            = rsget("yyyymm")
			FOneItem.FtargetGbn         = rsget("targetGbn")
			FOneItem.Ftitle             = rsget("title")

			FOneItem.Fregdate          = rsget("jungsanregdate")
			FOneItem.Ffinishflag       = rsget("finishflag")
			FOneItem.Fipkumdate        = rsget("ipkumdate")
			FOneItem.Ftaxregdate       = rsget("taxregdate")

			FOneItem.Fdifferencekey = rsget("differencekey")
			FOneItem.Ftaxtype      = rsget("taxtype")
			FOneItem.FTaxLinkidx   = rsget("taxlinkidx")
			FOneItem.Fneotaxno     = rsget("neotaxno")

            FOneItem.FeseroEvalSeq = rsget("eseroEvalSeq")

			FOneItem.Fcompany_name	= db2html(rsget("company_name"))
			FOneItem.Fjungsan_gubun	= db2html(rsget("jungsan_gubun"))
			FOneItem.Ftaxinputdate	= rsget("taxinputdate")
			FOneItem.Fcompany_no	= db2html(rsget("company_no"))
			FOneItem.Fbankingupflag = rsget("bankingupflag")

            FOneItem.FPrdMeachulsum     = rsget("PrdMeachulsum")
            FOneItem.FPrdCommissionSum  = rsget("PrdCommissionSum")
            FOneItem.FdlvMeachulsum     = rsget("dlvMeachulsum")
            FOneItem.FetMeachulsum      = rsget("etMeachulsum")
            FOneItem.FprdJungsanSum     = rsget("prdJungsanSum")
            FOneItem.FdlvJungsanSum     = rsget("dlvJungsanSum")
            FOneItem.FetJungsanSum      = rsget("etJungsanSum")
            FOneItem.FitemvatYn         = rsget("itemvatYn")

            FOneItem.Ftotalcommission   = FOneItem.FPrdCommissionSum
		end if
		rsget.close
    end Sub

    public Sub getMonthUpcheJungsanList
        Dim sqlStr, i
        sqlStr = "[db_jungsan].[dbo].[sp_Ten_getMonthJungsanListByJGubun]('"&FRectYYYYMM&"','"&FRectMakerid&"','"&FRectJGubun&"')"
        rsget.CursorLocation = adUseClient
	    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

		FResultCount = rsget.RecordCount
		FTotalCount  = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CUpcheJungsanTaxItem
                FItemList(i).Fid                = rsget("id")
                FItemList(i).Fjgubun            = rsget("jgubun")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fgroupid		    = rsget("groupid")
				FItemList(i).Fyyyymm            = rsget("yyyymm")
				FItemList(i).FtargetGbn         = rsget("targetGbn")
				FItemList(i).Ftitle             = rsget("title")

				FItemList(i).Fregdate          = rsget("jungsanregdate")
				FItemList(i).Ffinishflag       = rsget("finishflag")
				FItemList(i).Fipkumdate        = rsget("ipkumdate")
				FItemList(i).Ftaxregdate       = rsget("taxregdate")

				FItemList(i).Fdifferencekey = rsget("differencekey")
				FItemList(i).Ftaxtype      = rsget("taxtype")
				FItemList(i).FTaxLinkidx   = rsget("taxlinkidx")
				FItemList(i).Fneotaxno     = rsget("neotaxno")

                FItemList(i).FeseroEvalSeq = rsget("eseroEvalSeq")

				FItemList(i).Fcompany_name	= db2html(rsget("company_name"))
				FItemList(i).Fjungsan_gubun	= db2html(rsget("jungsan_gubun"))
				FItemList(i).Ftaxinputdate	= rsget("taxinputdate")
				FItemList(i).Fcompany_no	= db2html(rsget("company_no"))
				FItemList(i).Fbankingupflag = rsget("bankingupflag")

                FItemList(i).FPrdMeachulsum     = rsget("PrdMeachulsum")
                FItemList(i).FPrdCommissionSum  = rsget("PrdCommissionSum") ''상품판매수수료
                FItemList(i).FPgCommissionSum   = rsget("PgCommissionSum")  ''PG수수료
                FItemList(i).FdlvMeachulsum     = rsget("dlvMeachulsum")
                FItemList(i).FetMeachulsum      = rsget("etMeachulsum")
                FItemList(i).FprdJungsanSum     = rsget("prdJungsanSum")
                FItemList(i).FdlvJungsanSum     = rsget("dlvJungsanSum")
                FItemList(i).FetJungsanSum      = rsget("etJungsanSum")
                FItemList(i).FitemvatYn         = rsget("itemvatYn")

                FItemList(i).Ftotalcommission   = FItemList(i).FPrdCommissionSum + FItemList(i).FPgCommissionSum
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
    end sub

    public Sub getMonthUpcheJungsanListAdmAll
        Dim sqlStr, i
		sqlStr ="[db_jungsan].[dbo].[sp_Ten_getMonthJungsanAdmDtlGbnCnt]('"&FRectYYYYMM&"','"&FRectMakerid&"','"&FRectJGubun&"','"&FRecttargetGbn&"','"&FRectgroupid&"','"&FRectTaxType&"',"&FRectFinishFlag&",'"&FRectJungsanDate&"','"&FRectjacctcd&"','"&FRectNotIncTen&"','"&FRectSearchType&"','"&FRectSearchText&"')"
'rw sqlStr
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotalCount = rsget("CNT")
			set FSumaryOneItem = new CUpcheJungsanTaxItem
			FSumaryOneItem.FPrdMeachulsum     = rsget("PrdMeachulsum")
            FSumaryOneItem.FPrdCommissionSum  = rsget("PrdCommissionSum")
            FSumaryOneItem.FdlvMeachulsum     = rsget("dlvMeachulsum")
            FSumaryOneItem.FetMeachulsum      = rsget("etMeachulsum")
            FSumaryOneItem.FprdJungsanSum     = rsget("prdJungsanSum")
            FSumaryOneItem.FdlvJungsanSum     = rsget("dlvJungsanSum")
            FSumaryOneItem.FetJungsanSum      = rsget("etJungsanSum")

            FSumaryOneItem.FpgcommissionSum   = rsget("pgcommissionSum")
            FSumaryOneItem.Ftotalcommission   = FSumaryOneItem.FPrdCommissionSum+FSumaryOneItem.FpgcommissionSum

            FSumaryOneItem.FSSuply = rsget("SSuply")
            FSumaryOneItem.FCSuply = rsget("CSuply")
            FSumaryOneItem.FMSuply = rsget("MSuply")
            FSumaryOneItem.FDSuply = rsget("DSuply")
            FSumaryOneItem.FESuply = rsget("ESuply")

            FSumaryOneItem.FDlvPgCommission = rsget("DlvPgCommission")
		END IF
		rsget.close

		IF FTotalCount > 0 THEN
            sqlStr = "[db_jungsan].[dbo].[sp_Ten_getMonthJungsanAdmDtlGbnList]('"&FRectYYYYMM&"','"&FRectMakerid&"','"&FRectJGubun&"','"&FRecttargetGbn&"','"&FRectgroupid&"','"&FRectTaxType&"',"&FRectFinishFlag&",'"&FRectJungsanDate&"','"&FRectjacctcd&"','"&FRectNotIncTen&"',"&FPageSize&","&FCurrPage&",'"&FRectSearchType&"','"&FRectSearchText&"')"
'rw sqlStr
            rsget.CursorLocation = adUseClient
    	    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

    		FResultCount = rsget.RecordCount
    		FtotalPage =  CInt(FTotalCount\FPageSize)
    		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
    			FtotalPage = FtotalPage +1
    		end if

    		if (FResultCount<1) then FResultCount=0

    		redim preserve FItemList(FResultCount)
    		i=0
    		if  not rsget.EOF  then
    			do until rsget.eof
    				set FItemList(i) = new CUpcheJungsanTaxItem
                    FItemList(i).Fid                = rsget("id")
                    FItemList(i).Fjgubun            = rsget("jgubun")
    				FItemList(i).Fmakerid           = rsget("makerid")
    				FItemList(i).Fgroupid		    = rsget("groupid")
    				FItemList(i).Fyyyymm            = rsget("yyyymm")
    				FItemList(i).FtargetGbn         = rsget("targetGbn")
    				FItemList(i).Ftitle             = rsget("title")

    				FItemList(i).Fregdate          = rsget("jungsanregdate")
    				FItemList(i).Ffinishflag       = rsget("finishflag")
    				FItemList(i).Fipkumdate        = rsget("ipkumdate")
    				FItemList(i).Ftaxregdate       = rsget("taxregdate")

    				FItemList(i).Fdifferencekey = rsget("differencekey")
    				FItemList(i).Ftaxtype      = rsget("taxtype")
    				FItemList(i).FTaxLinkidx   = rsget("taxlinkidx")
    				FItemList(i).Fneotaxno     = rsget("neotaxno")

                    FItemList(i).FeseroEvalSeq = rsget("eseroEvalSeq")

    				FItemList(i).Fcompany_name	= db2html(rsget("company_name"))
    				FItemList(i).Fjungsan_gubun	= db2html(rsget("jungsan_gubun"))
    				FItemList(i).Ftaxinputdate	= rsget("taxinputdate")
    				FItemList(i).Fcompany_no	= db2html(rsget("company_no"))
    				FItemList(i).Fbankingupflag = rsget("bankingupflag")

                    FItemList(i).FPrdMeachulsum     = rsget("PrdMeachulsum")
                    FItemList(i).FPrdCommissionSum   = rsget("PrdCommissionSum")
                    FItemList(i).FdlvMeachulsum     = rsget("dlvMeachulsum")
                    FItemList(i).FetMeachulsum      = rsget("etMeachulsum")
                    FItemList(i).FprdJungsanSum     = rsget("prdJungsanSum")
                    FItemList(i).FdlvJungsanSum     = rsget("dlvJungsanSum")
                    FItemList(i).FetJungsanSum      = rsget("etJungsanSum")
                    FItemList(i).FitemvatYn         = rsget("itemvatYn")

                    FItemList(i).FerpCust_cd    = rsget("erpCust_cd")
                    FItemList(i).FerpUsing      = rsget("erpUsing")
                    FItemList(i).Fcompany_name  = rsget("company_name")

                    FItemList(i).FpgcommissionSum   = rsget("pgcommissionSum")
                    FItemList(i).Ftotalcommission   = FItemList(i).FPrdCommissionSum+FItemList(i).FpgcommissionSum
                    FItemList(i).FDlvPgCommission = rsget("DlvPgCommission")

                    FItemList(i).FSSuply = rsget("SSuply")
                    FItemList(i).FCSuply = rsget("CSuply")
                    FItemList(i).FMSuply = rsget("MSuply")
                    FItemList(i).FDSuply = rsget("DSuply")
                    FItemList(i).FESuply = rsget("ESuply")

                    if isNULL(FItemList(i).FSSuply) then FItemList(i).FSSuply=0
                    if isNULL(FItemList(i).FCSuply) then FItemList(i).FCSuply=0
                    if isNULL(FItemList(i).FMSuply) then FItemList(i).FMSuply=0
                    if isNULL(FItemList(i).FDSuply) then FItemList(i).FDSuply=0
                    if isNULL(FItemList(i).FESuply) then FItemList(i).FESuply=0

                    FItemList(i).Fjacctcd = rsget("jacctcd")
                    FItemList(i).Facc_nm  = rsget("acc_nm")

    				i=i+1
    				rsget.moveNext
    			loop
    		end if

    		rsget.close
    	END IF
    end sub

    '' admin2009scm\admin\upchejungsan\monthjungsanSummaryAdm.asp 에서 사용
    public Sub getMonthUpcheJungsanSummaryAdm()
        Dim sqlStr, i

        sqlStr = "[db_jungsan].[dbo].[sp_Ten_getMonthJungsanSummaryAdm]('"&FRectYYYYMM&"','"&FRecttargetGbn&"','"&FRectJGubun&"','"&FRectMakerid&"','"&FRectgroupid&"','"&FRectTaxType&"','"&FRectItemVatYn&"',"&FRectFinishFlag&",'"& FRectcomm_cd &"')"

        'rw sqlStr & "<Br>"
        rsget.CursorLocation = adUseClient
	    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		FResultCount = rsget.RecordCount
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CUpcheJungsanSummaryItem
                FItemList(i).Fjgubun            = rsget("jgubun")
                FItemList(i).FtargetGbn         = FRecttargetGbn
                FItemList(i).Fsitename          = rsget("sitename")
                FItemList(i).FitemvatYn         = rsget("itemvatYn")
                FItemList(i).Fgubuncd           = rsget("gubuncd")
                ''FItemList(i).FcommissionSum     = rsget("commissionSum")
                FItemList(i).FsellcashSum       = rsget("sellcashSum")
                FItemList(i).FreducedpriceSum   = rsget("reducedpriceSum")
                FItemList(i).FsuplycashSum      = rsget("suplycashSum")
                FItemList(i).FsitenameName      = rsget("sitenameName")
                FItemList(i).Fcomm_name         = rsget("comm_name")

                ''2016/09/29 수정
                FItemList(i).FitemcommissionSum  = rsget("itemcommissionSum")
                FItemList(i).FpgcommissionSum    = rsget("pgcommissionSum")

                FItemList(i).FttlcommissionSum   = FItemList(i).FitemcommissionSum+FItemList(i).FpgcommissionSum

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
    end Sub

    public Sub getJungsanTaxListAdm()
        Dim sqlStr, i
		sqlStr ="[db_jungsan].[dbo].[sp_Ten_JungsanTaxAdmCnt]('"&FRectMakerid&"','"&FRectYYYYMM&"','"&FRectJGubun&"','"&FRecttargetGbn&"','"&FRectgroupid&"',"&FRectFinishFlag&",'"&FRectJungsanDate&"','"&FRectJungsanException&"')"
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotalCount = rsget("CNT")
		END IF
		rsget.close
		IF FTotalCount > 0 THEN

    		sqlStr ="[db_jungsan].[dbo].sp_Ten_JungsanTaxAdmList('"&FRectMakerid&"','"&FRectYYYYMM&"','"&FRectJGubun&"','"&FRecttargetGbn&"','"&FRectgroupid&"',"&FRectFinishFlag&",'"&FRectJungsanDate&"',"&FPageSize&","&FCurrPage&",'"&FRectJungsanException&"')"
    		rsget.CursorLocation = adUseClient
		    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

    		FResultCount = rsget.RecordCount
    		FtotalPage =  CInt(FTotalCount\FPageSize)
    		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
    			FtotalPage = FtotalPage +1
    		end if

    		if (FResultCount<1) then FResultCount=0

    		redim preserve FItemList(FResultCount)
    		i=0
    		if  not rsget.EOF  then
    			do until rsget.eof
    				set FItemList(i) = new CUpcheJungsanTaxItem

    				FItemList(i).Fid                = rsget("id")
    				FItemList(i).Fmakerid           = rsget("makerid")
    				FItemList(i).Fgroupid		    = rsget("groupid")
    				FItemList(i).Fyyyymm            = rsget("yyyymm")
    				FItemList(i).FtargetGbn         = rsget("targetGbn")
    				FItemList(i).Ftitle             = rsget("title")

    				FItemList(i).Fregdate          = rsget("regdate")
    				FItemList(i).Ffinishflag       = rsget("finishflag")
    				FItemList(i).Fipkumdate        = rsget("ipkumdate")
    				FItemList(i).Ftaxregdate       = rsget("taxregdate")

    				FItemList(i).Fdifferencekey = rsget("differencekey")
    				FItemList(i).Ftaxtype      = rsget("taxtype")
    				FItemList(i).FTaxLinkidx   = rsget("taxlinkidx")
    				FItemList(i).Fneotaxno     = rsget("neotaxno")

                    FItemList(i).FeseroEvalSeq = rsget("eseroEvalSeq")

    				FItemList(i).Fcompany_name	= db2html(rsget("company_name"))
    				FItemList(i).Fjungsan_gubun	= db2html(rsget("jungsan_gubun"))
    				FItemList(i).Ftaxinputdate	= rsget("taxinputdate")
    				FItemList(i).Fcompany_no	= db2html(rsget("company_no"))
    				FItemList(i).Fbankingupflag = rsget("bankingupflag")

                    FItemList(i).Fjgubun                = rsget("jgubun")

                    FItemList(i).FtotalJungsanSum   = rsget("totalJungsanSum")
                    FItemList(i).Ftotalcommission   = rsget("totalcommission")

                    FItemList(i).Fjungsan_date      = rsget("jungsan_date")
                    FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")


    				i=i+1
    				rsget.moveNext
    			loop
    		end if


    		rsget.close
		END IF
    end sub

    public Sub getEvalJungsanTaxTargetListAdm()
        Dim sqlStr, i
		

        sqlStr ="[db_jungsan].[dbo].[usp_TEN_JungsanBatchEvalTarget]('"&FRectYYYYMM&"','"&FRecttargetGbn&"','"&FRectJungsanDate&"',"&FPageSize&",'"&FRectJungsanException&"')"

        'response.write sqlStr & "<Br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

        FResultCount = rsget.RecordCount

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)
        i=0
        if  not rsget.EOF  then
            do until rsget.eof
                set FItemList(i) = new CUpcheJungsanTaxItem

                FItemList(i).Fid                = rsget("id")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fgroupid		    = rsget("groupid")
                FItemList(i).Fyyyymm            = rsget("yyyymm")
                FItemList(i).FtargetGbn         = rsget("targetGbn")
                FItemList(i).Ftitle             = rsget("title")

                FItemList(i).Fregdate          = rsget("regdate")
                FItemList(i).Ffinishflag       = rsget("finishflag")
                FItemList(i).Fipkumdate        = rsget("ipkumdate")
                FItemList(i).Ftaxregdate       = rsget("taxregdate")

                FItemList(i).Fdifferencekey = rsget("differencekey")
                FItemList(i).Ftaxtype      = rsget("taxtype")
                FItemList(i).FTaxLinkidx   = rsget("taxlinkidx")
                FItemList(i).Fneotaxno     = rsget("neotaxno")

                FItemList(i).FeseroEvalSeq = rsget("eseroEvalSeq")

                FItemList(i).Fcompany_name	= db2html(rsget("company_name"))
                FItemList(i).Fjungsan_gubun	= db2html(rsget("jungsan_gubun"))
                FItemList(i).Ftaxinputdate	= rsget("taxinputdate")
                FItemList(i).Fcompany_no	= db2html(rsget("company_no"))
                FItemList(i).Fbankingupflag = rsget("bankingupflag")

                FItemList(i).Fjgubun                = rsget("jgubun")

                FItemList(i).FtotalJungsanSum   = rsget("totalJungsanSum")
                FItemList(i).Ftotalcommission   = rsget("totalcommission")

                FItemList(i).Fjungsan_date      = rsget("jungsan_date")
                FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")


                i=i+1
                rsget.moveNext
            loop
        end if


        rsget.close
		
    end sub

    public sub getJungsanTaxListByMakerid2GroupID
        Dim sqlStr, i
		sqlStr ="[db_jungsan].[dbo].[sp_Ten_JungsanTaxCnt]('"&FRectMakerid&"')"
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotalCount = rsget("CNT")
		END IF
		rsget.close

		IF FTotalCount > 0 THEN

    		sqlStr ="[db_jungsan].[dbo].sp_Ten_JungsanTaxList('"&FRectMakerid&"',"&FPageSize&","&FCurrPage&")"
    		rsget.CursorLocation = adUseClient
		    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

    		FResultCount = rsget.RecordCount
    		FtotalPage =  CInt(FTotalCount\FPageSize)
    		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
    			FtotalPage = FtotalPage +1
    		end if

    		if (FResultCount<1) then FResultCount=0

    		redim preserve FItemList(FResultCount)
    		i=0
    		if  not rsget.EOF  then
    			do until rsget.eof
    				set FItemList(i) = new CUpcheJungsanTaxItem

    				FItemList(i).Fid                = rsget("id")
    				FItemList(i).Fmakerid           = rsget("makerid")
    				FItemList(i).Fgroupid		    = rsget("groupid")
    				FItemList(i).Fyyyymm            = rsget("yyyymm")
    				FItemList(i).FtargetGbn         = rsget("targetGbn")
    				FItemList(i).Ftitle             = rsget("title")

    				FItemList(i).Fregdate          = rsget("regdate")
    				FItemList(i).Ffinishflag       = rsget("finishflag")
    				FItemList(i).Fipkumdate        = rsget("ipkumdate")
    				FItemList(i).Ftaxregdate       = rsget("taxregdate")

    				FItemList(i).Fdifferencekey = rsget("differencekey")
    				FItemList(i).Ftaxtype      = rsget("taxtype")
    				FItemList(i).FTaxLinkidx   = rsget("taxlinkidx")
    				FItemList(i).Fneotaxno     = rsget("neotaxno")

                    FItemList(i).FeseroEvalSeq = rsget("eseroEvalSeq")

    				FItemList(i).Fcompany_name	= db2html(rsget("company_name"))
    				FItemList(i).Fjungsan_gubun	= db2html(rsget("jungsan_gubun"))
    				FItemList(i).Ftaxinputdate	= rsget("taxinputdate")
    				FItemList(i).Fcompany_no	= db2html(rsget("company_no"))
    				FItemList(i).Fbankingupflag = rsget("bankingupflag")

                    FItemList(i).Fjgubun                = rsget("jgubun")

                    FItemList(i).FtotalJungsanSum   = rsget("totalJungsanSum")
                    FItemList(i).Ftotalcommission   = rsget("totalcommission")

                    FItemList(i).Fjungsan_date      = rsget("jungsan_date")
                    FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")

'    				FItemList(i).FBillsiteCode = rsget("BillsiteCode")
'    				FItemList(i).Fbigo			   = db2html(rsget("bigo"))
'    				FItemList(i).FDesignerEmail	   = rsget("jungsan_email")
'    				FItemList(i).Fjungsan_bank		= rsget("jungsan_bank")
'    				FItemList(i).Fjungsan_date		= rsget("jungsan_date")
'    				FItemList(i).Fjungsan_acctno	= rsget("jungsan_acctno")
'    				FItemList(i).Fjungsan_acctname	= rsget("jungsan_acctname")
'                   FItemList(i).Fjungsan_hp   = rsget("jungsan_hp")
'                   FItemList(i).FbillSiteName = rsget("billSiteName")


'                   FItemList(i).FpreFixedTaxDate   = rsget("preFixedTaxDate")          ''2012-09-03 추가
'    				FItemList(i).Fub_cnt           = rsget("ub_cnt")
'    				FItemList(i).Fub_totalsellcash = rsget("ub_totalsellcash")
'    				FItemList(i).Fub_totalsuplycash= rsget("ub_totalsuplycash")
'    				FItemList(i).Fub_comment       = db2html(rsget("ub_comment"))
'    				FItemList(i).Fme_cnt           = rsget("me_cnt")
'    				FItemList(i).Fme_totalsellcash = rsget("me_totalsellcash")
'    				FItemList(i).Fme_totalsuplycash= rsget("me_totalsuplycash")
'    				FItemList(i).Fme_comment       = db2html(rsget("me_comment"))
'    				FItemList(i).Fwi_cnt           = rsget("wi_cnt")
'    				FItemList(i).Fwi_totalsellcash = rsget("wi_totalsellcash")
'    				FItemList(i).Fwi_totalsuplycash= rsget("wi_totalsuplycash")
'    				FItemList(i).Fwi_comment       = db2html(rsget("wi_comment"))
'
'    				FItemList(i).Fet_cnt           = rsget("et_cnt")
'    				FItemList(i).Fet_totalsellcash = rsget("et_totalsellcash")
'    				FItemList(i).Fet_totalsuplycash= rsget("et_totalsuplycash")
'    				FItemList(i).Fet_comment       = db2html(rsget("et_comment"))
'    				FItemList(i).Fsh_cnt           = rsget("sh_cnt")
'    				FItemList(i).Fsh_totalsellcash = rsget("sh_totalsellcash")
'    				FItemList(i).Fsh_totalsuplycash= rsget("sh_totalsuplycash")
'    				FItemList(i).Fsh_comment       = db2html(rsget("sh_comment"))

'                    ''2014/01/27 추가

'                    FItemList(i).Fwi_totalreducedprice  = rsget("wi_totalreducedprice")
'                    FItemList(i).Fub_totalreducedprice  = rsget("ub_totalreducedprice")
'                    FItemList(i).Fet_totalreducedprice  = rsget("et_totalreducedprice")
'                    FItemList(i).Fdlv_totalreducedprice = rsget("dlv_totalreducedprice")
'                    FItemList(i).Fdlv_totalsuplycash    = rsget("dlv_totalsuplycash")
'                    FItemList(i).Ftotalcommission       = rsget("totalcommission")
    				i=i+1
    				rsget.moveNext
    			loop
    		end if


    		rsget.close
		END IF
    end Sub

    Public Sub getMonthCsjungsanList
		Dim sqlStr, i, addSql

		If FRectYYYYMM <> "" then
			addSql = addSql & " and m.yyyymm = '"& FRectYYYYMM &"' "
		End If

		If FRectJGubun <> "" Then
			addSql = addSql & " and m.jgubun = '"& FRectJGubun &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.designerid, d.mastercode, d.sitename, d.buyname, d.reqname, d.itemid, d.itemoption, d.itemname, isnull(d.itemoptionname, '') as itemoptionname, d.itemno "
		sqlStr = sqlStr & " ,d.sellcash, ((d.sellcash - d.reducedprice) + d.commission) as couponPlusCommi, d.sellcash - d.reducedprice as coupoonDiscount, d.reducedprice "
		sqlStr = sqlStr & " ,d.commission, d.pgcommission, d.suplycash, (d.itemno * d.suplycash) as sumsuplycash "
		sqlStr = sqlStr & " ,Case WHEN d.paymethod = '100' Then '신용카드' "
		sqlStr = sqlStr & " 	  WHEN d.paymethod = '110' Then 'OK+신용카드' "
		sqlStr = sqlStr & " 	  WHEN d.paymethod = '7' Then '무통장' "
		sqlStr = sqlStr & " 	  WHEN d.paymethod = '400' Then '휴대폰' "
		sqlStr = sqlStr & " 	  WHEN d.paymethod = '20' Then '실시간이체' "
		sqlStr = sqlStr & " 	  WHEN d.paymethod = '50' Then '외부' else isnull(d.paymethod, '') end as paymethod "
        sqlStr = sqlStr & " , (select isnull(authcode,'') from [db_order].[dbo].tbl_order_master where orderserial=d.mastercode) as authcode"
		sqlStr = sqlStr & " FROM db_jungsan.dbo.tbl_designer_jungsan_master m WITH (INDEX (IX_tbl_designer_jungsan_master_yyyymm)) "
		sqlStr = sqlStr & " JOIN db_jungsan.dbo.tbl_designer_jungsan_detail d WITH (INDEX (IX_tbl_designer_jungsan_detail_masteridx)) on d.masteridx=m.id "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and d.gubuncd in ('DT') "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY m.designerid, d.mastercode "
rw sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordcount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		FPageCount = FCurrPage - 1

		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				set FItemList(i) = new CUpcheJungsanTaxItem
					FItemList(i).FDesignerid		= rsget("designerid")
					FItemList(i).FMastercode 		= rsget("mastercode")
					FItemList(i).FSitename 			= rsget("sitename")
					FItemList(i).FBuyname 			= rsget("buyname")
					FItemList(i).FReqname 			= rsget("reqname")
					FItemList(i).FItemid 			= rsget("itemid")
					FItemList(i).FItemoption 		= rsget("itemoption")
					FItemList(i).FItemname 			= rsget("itemname")
					FItemList(i).FItemoptionname 	= rsget("itemoptionname")
					FItemList(i).FItemno 			= rsget("itemno")
					FItemList(i).FSellcash 			= rsget("sellcash")
					FItemList(i).FCouponPlusCommi 	= rsget("couponPlusCommi")
					FItemList(i).FCoupoonDiscount 	= rsget("coupoonDiscount")
					FItemList(i).FReducedprice 		= rsget("reducedprice")
					FItemList(i).FCommission 		= rsget("commission")
					FItemList(i).FPgcommission 		= rsget("pgcommission")
					FItemList(i).FSuplycash 		= rsget("suplycash")
					FItemList(i).FSumsuplycash 		= rsget("sumsuplycash")
					FItemList(i).FPaymethod 		= rsget("paymethod")
                    FItemList(i).Fauthcode 		    = rsget("authcode")
				rsget.Movenext
				i = i + 1
			Loop
		End If
		rsget.Close

	End Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 300
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

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
