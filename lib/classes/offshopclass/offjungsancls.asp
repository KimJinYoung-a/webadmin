<%
'####################################################
' Description :  오프라인 정산 클래스
' History : 2010.05.13 한용민 수정
'####################################################

function getPartnerId2GroupID(ipartnerid)
    dim sqlStr
	sqlStr = "select groupid from db_partner.dbo.tbl_partner where id='"&ipartnerid&"'"

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    getPartnerId2GroupID = rsget("groupid")
    end if
    rsget.Close
end function

function chkAvailViewJungsanOF(jid,makerid,groupid)
    dim sqlStr

    chkAvailViewJungsanOF= false

    sqlStr = "select top 1 idx"&VBCRLF
    sqlStr = sqlStr & " from db_jungsan.dbo.tbl_off_jungsan_master"&VBCRLF
    sqlStr = sqlStr & " where idx="&jid&VBCRLF
    sqlStr = sqlStr & " and finishflag>0"&VBCRLF
    sqlStr = sqlStr & " and (makerid='"&makerid&"' or groupid='"&groupid&"')"&VBCRLF
    rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    chkAvailViewJungsanOF = true
    end if
    rsget.Close

end function

Class COffJungsanSummaryByTaxDateItem
    public Ftaxregdate
    public Fjungsansum_susi
    public Fjungsansum_31date
    public Fjungsansum_15date
    public Fjungsansum_etcdate
    public Fewol_jungsansum
    public Fnext_jungsansum
    public Ffixedsum
    public Fipkumsum
    public Ftot_jungsanprice

    Private Sub Class_Initialize()
        Ftaxregdate        = 0
        Fjungsansum_susi   = 0
        Fjungsansum_31date = 0
        Fjungsansum_15date = 0
        Fjungsansum_etcdate= 0
        Fewol_jungsansum   = 0
        Fnext_jungsansum   = 0
        Ffixedsum          = 0
        Fipkumsum          = 0
        Ftot_jungsanprice  = 0
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class

Class COffJungsanSummaryItem
    public Fyyyymm
    public Fjungsan_date_off
    public FTW_price
    public FUW_price
    public FCM_price
    public FOM_price
    public FSM_price
    public FET_price
    public Fipkumsum
    public Ffixedsum
    public Ffixedthissum
    public Ffixednextsum
    public Fwaitsum
    public Ftot_jungsanprice

    Private Sub Class_Initialize()
        FTW_price         = 0
        FUW_price         = 0
        FCM_price         = 0
        FOM_price         = 0
        FSM_price         = 0
        FET_price         = 0
        Fipkumsum         = 0
        Ffixedsum         = 0
        Ffixedthissum     = 0
        Ffixednextsum     = 0
        Fwaitsum          = 0
        Ftot_jungsanprice = 0
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class

Class COffJungsanDetailSummaryItem
    public Fshopid
    public Fgubuncd
    public Fcomm_name
    public Fshopname
    public Ftot_itemno
    public Ftot_orgsellprice
    public Ftot_realsellprice
    public Ftot_jungsanprice

    ''기본 정산조건.
    public Fchargediv
    public Fdefaultmargin
    public Fdefaultsuplymargin
    public Fautojungsan
    public Fautojungsandiv


    public Fjgubun
    public FitemvatYn
    public Ftot_commission

    public function GetItemVatTypeName()
        if isNULL(FitemvatYn) then Exit function

        if (FitemvatYn="Y") then
            GetItemVatTypeName = "과세"
        elseif (FitemvatYn="N") then
            GetItemVatTypeName = "<font color=red>면세</font>"
        else
            GetItemVatTypeName = FitemvatYn
        end if
    end function

    public function getJSummaryGugunName
        if IsCommissionTax then
            getJSummaryGugunName = "수수료정산"
        else
            if Fgubuncd="B021" or Fgubuncd="B022" or Fgubuncd="B023" or Fgubuncd="B032" then
                getJSummaryGugunName = "입고분매입"
            elseif Fgubuncd="B011" or Fgubuncd="B012" or Fgubuncd="B013" then
                getJSummaryGugunName = "판매분매입"
            elseif Fgubuncd="B031" then
                getJSummaryGugunName = "출고분매입"
            elseif Fgubuncd="B999" then
                getJSummaryGugunName = "기타출고매입"
            else
                getJSummaryGugunName = Fgubuncd
            end if
        end if
    end function

    public function IsCommissionTax()  ''수수료 매출 세금 계산서 인지.
        IsCommissionTax = false
        if isNULL(Fjgubun) then Exit function

        IsCommissionTax = (Fjgubun="CC")
    end function

    public function GetChargeDivName()
        select case Fchargediv
            case "2"
                : GetChargeDivName = "텐위"
            case "6"
                : GetChargeDivName = "업위"
            case "4"
                : GetChargeDivName = "텐매"
            case "5"
                : GetChargeDivName = "출고"
            case "8"
                : GetChargeDivName = "업매"
            case else
                : GetChargeDivName = Fchargediv
        end select

    end function

    Private Sub Class_Initialize()
        Ftot_itemno =0
        Ftot_orgsellprice =0
        Ftot_realsellprice =0
        Ftot_jungsanprice =0
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class

class COffJungsanDetailItem
    public Fdetailidx
    public Fmasteridx
    '' 추가
    public Fshopid      ''  오프라인용 매입인경우 streetshop800 (가맹점 대표코드)
    public Fgubuncd     ''  정산구분. //위탁판매, 업체위탁판매, 매입, 업체매입, 출고매입  ([db_jungsan].[dbo].tbl_jungsan_comm_code)
                        ''              B011,      B012,         B021,     B022,     B031
    public Forderno
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public Fitemname
    public Fitemoptionname
    public Forgsellprice
    public Frealsellprice
    public Fsuplyprice
    public Fitemno
    public Fmakerid
    public Flinkidx
    public Fcentermwdiv
    public Fvatinclude ''상품

    public Fcommission
    public Fiszerotax
    public Fpaymethod
    public Fvatyn       '' 정산 디테일

    public function GetBarCode()
        GetBarCode = Fitemgubun + Format00(6,Fitemid) + Fitemoption
        if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
    end function

    Private Sub Class_Initialize()
        Forgsellprice =0
        Frealsellprice =0
        Fsuplyprice =0
        Fitemno =0
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class

class COffJungsanMasterItem
    public Fidx
    public Fyyyymm
    public Fdifferencekey
    public Ftaxtype
    public Fmakerid
    public Ftitle
    public Ftot_itemno
    public Ftot_orgsellprice
    public Ftot_realsellprice
    public Ftot_jungsanprice
    public FTW_price  '' B011 위탁판매
    public FUW_price  '' B012 업체위탁판매
    public FCM_price  '' B031 출고매입
    public FOM_price  '' B021 오프매입
    public FSM_price  '' B022 매장매입
    public FET_price  '' B999 기타보정
    public Fcomment
    public Ffinishflag
    public Fipkumdate
    public Ftaxregdate
    public Ftaxinputdate
    public Ftaxlinkidx
    public Fneotaxno
    public Fbankingupflag
    public Fregdate
    ''수기정산 존재
    public Fautojungsan
    public Fjungsan_email
    public Fjungsan_bank
    public Fjungsan_date_off
    public Fjungsan_acctno
    public Fjungsan_acctname
    public Fcompany_name
    public Fjungsan_gubun
    public Fcompany_no
    public FGroupid
    public Favailneo

    public Fipkum_bank
    public Fipkum_acctno

    public FBillsiteCode
    public FISSU_SEQNO
    public FeseroEvalSeq
    public FbillSiteName
    public FipFileNo
    public FtargetGbn
    public FholdGroupid
    public Fholdcause
    public FrefPayreqIdx
    public FpreFixedTaxDate

    ''2014 추가
    public FjGubun
    public Ftotalcommission
    public FitemvatYn
    public Fjacctcd
    public Fjacc_nm

    public function getBill_SELL_DAM_DEPT()
        if (FtargetGbn="ON") THEN
            getBill_SELL_DAM_DEPT = "온라인"
        elseif (FtargetGbn="OF") THEN
            getBill_SELL_DAM_DEPT = "오프라인"
        elseif (FtargetGbn="AC") THEN
            getBill_SELL_DAM_DEPT = "더핑거스"
        end if
    end function

    public function getBill_NM_ITEM()
        getBill_NM_ITEM = getBill_SELL_DAM_DEPT &" "& Fmakerid &" "& Ftitle

		if (FtargetGbn="OF") and InStr(Ftitle, "오프샵") > 0 then
			'// "2019년 08월 오프라인 earpearp 매입 정산" 으로 변경요청
			getBill_NM_ITEM = Replace(Ftitle, "오프샵", getBill_SELL_DAM_DEPT &" "& Fmakerid)
		end if
    end function

    public function IsCommissionTax()  ''수수료 매출 세금 계산서 인지.
        IsCommissionTax = false
        if isNULL(Fjgubun) then Exit function

        IsCommissionTax = (Fjgubun="CC")
    end function

    public function getJGubunName
        if (FjGubun="MM") then
            getJGubunName = "매입"
        elseif (FjGubun="CC") then
            getJGubunName = "<font color=blue>수수료</font>"
        elseif Fjgubun="CE" then
            getJGubunName = "기타"
        else
            getJGubunName = FjGubun
        end if
    end function

    public function GetItemVatTypeName()
        if isNULL(FitemvatYn) then Exit function

        if (FitemvatYn="Y") then
            GetItemVatTypeName = "과세"
        elseif (FitemvatYn="N") then
            GetItemVatTypeName = "<font color=red>면세</font>"
        else
            GetItemVatTypeName = FitemvatYn
        end if
    end function

    ''전자 세금계산서 관련
    public function GetTotalTaxSuply()
		if Ftaxtype="01" then
			GetTotalTaxSuply = CLng(Ftot_jungsanprice / 1.1)
		else
			GetTotalTaxSuply = CLNG(Ftot_jungsanprice)
		end if
	end function

	public function GetTotalTaxVat()
		GetTotalTaxVat = CLNG(Ftot_jungsanprice) - GetTotalTaxSuply
	end function

	public function getDbDate()
		dim sqlstr
		sqlstr = " select convert(varchar(10),getdate(),21) as nowdate "
		rsget.Open sqlStr,dbget,1
		getDbDate = CDate(rsget("nowdate"))
		rsget.Close
	end function

	public function GetNormalTaxDate()
	    '' 이미 지정되 있는경우 지정일로 그외에는 정산일 말일이 기본값.
		if Not(IsNULL(FpreFixedTaxDate)) and (FpreFixedTaxDate<>"") then
			GetNormalTaxDate = FpreFixedTaxDate
		else
		    GetNormalTaxDate = dateserial(Left(Fyyyymm,4),Right(Fyyyymm,2)+1,1-1) ''': 정산월 말일
		end if
	end function

	public function GetPreFixSegumil()
		dim thisdate, maytaxdate
		dim ithis1day , ithis21day, premonth1day, premonth21day

		thisdate = getDbDate()
		maytaxdate = GetNormalTaxDate()

        '' 12일까지 마감할 경우 13으로 세팅 13일까지일경우 14
        '' 10일까지 마감할 경우 11로 쎄팅
		premonth1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"01")
		premonth21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"12") ''11
		ithis1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"01")
		ithis21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"12") ''11

        ''(매달 12일 까지 발행시 : 정산월 말일)<br>
		''(매달 13일 이후 발행 : 발행월 1일)<br>
		''(이월 내역발행시 12일까지 발행: 발행전월 1일)<br>
		''(이월 내역발행시 13일 이후 발행: 발행월 1일)
		''그외 : 발행일=Today

		''######################################## 2017-09-21 김진영 추가 ########################################
		Dim strSql, taxdate
		strSql = ""
		strSql = strSql & " SELECT TOP 1 taxdate FROM "
		strSql = strSql & " [db_sitemaster].[dbo].[tbl_taxdate_manage] "
		strSql = strSql & " WHERE yyyymm = '"& Left(thisdate, 7) &"' "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			taxdate = CDate(rsget("taxdate"))
		End If
		rsget.Close
		''######################################## 2017-09-21 김진영 추가 끝 #######################################
		if (CStr(FYYYYMM) = Left(CStr(premonth1day),7)) then
		''정상 발행의 경우
	    	'if (thisdate>=ithis21day) then		2017-09-21 김진영 주석, 아래 taxdate 변경 및 부등호 제거
		    if (thisdate > taxdate) then
		    ''13일 이후 발행건은 이월됨 클릭일 1일
		        GetPreFixSegumil = ithis1day
		    else
		        GetPreFixSegumil = maytaxdate
		    end if
		elseif (CStr(FYYYYMM) < Left(CStr(premonth1day),7)) then
		''이월 발행의 경우
		    'if (thisdate>=ithis21day) then		2017-09-21 김진영 주석, 아래 taxdate 변경 및 부등호 제거
		    if (thisdate > taxdate) then
		    ''13일 이후 발행건은 이월됨 클릭일 1일
		        GetPreFixSegumil = ithis1day
		    else
		        GetPreFixSegumil = premonth1day
		    end if
		else
		    GetPreFixSegumil = Left(CStr(thisdate),10)
		end if

		if (Fidx=43993) then
            GetPreFixSegumil = "2009-02-28"
        end if

        ''' 2012-09-03 추가.
        if Not(IsNULL(FpreFixedTaxDate)) and (FpreFixedTaxDate<>"") then
            'if (thisdate>=ithis21day) and (CStr(FpreFixedTaxDate)<ithis1day) then
            if (thisdate > taxdate) and (CStr(FpreFixedTaxDate)<ithis1day) then
                ''기본 계산값
            elseif (Left(FpreFixedTaxDate,10)<CStr(premonth1day)) then
                ''기본 계산값
            ELSE
                GetPreFixSegumil = Left(FpreFixedTaxDate,10)
            end if
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
		IsElecFreeTaxCase = (Ftaxtype="02") 'and (Fjungsan_gubun="면세")
	end function

    public function IsEditenable()
        IsEditenable = (Ffinishflag="0")
    end function

    public function GetSimpleTaxtypeName()
		if Ftaxtype="01" then
			GetSimpleTaxtypeName = "과세"
		elseif Ftaxtype="02" then
			GetSimpleTaxtypeName = "면세"
		elseif Ftaxtype="03" then
			GetSimpleTaxtypeName = "간이"
		end if
	end function

	public function GetTaxtypeNameColor()
		if Ftaxtype="01" then
			GetTaxtypeNameColor = "#000000"
		elseif Ftaxtype="02" then
			GetTaxtypeNameColor = "#FF3333"
		elseif Ftaxtype="03" then
			GetTaxtypeNameColor = "#3333FF"
		end if
	end function

	public function GetStateName()
		if Ffinishflag="0" then
			GetStateName = "수정중"
		elseif Ffinishflag="1" then
			GetStateName = "업체확인중"
		elseif Ffinishflag="2" then
			GetStateName = "업체확인완료"
		elseif Ffinishflag="3" then
			GetStateName = "정산확정"
		elseif Ffinishflag="7" then
			GetStateName = "입금완료"
		elseif Ffinishflag="8" then
			GetStateName = "정산안함"
		elseif Ffinishflag="9" then
			GetStateName = "통합정산"
		else
            GetStateName = Ffinishflag
		end if
	end function

	public function GetStateColor()
		if Ffinishflag="0" then
			GetStateColor = "#000000"
		elseif Ffinishflag="1" then
			GetStateColor = "#448888"
		elseif Ffinishflag="2" then
			GetStateColor = "#0000FF"
		elseif Ffinishflag="3" then
			GetStateColor = "#0000FF"
		elseif Ffinishflag="7" then
			GetStateColor = "#FF0000"
		elseif Ffinishflag="8" then
			GetStateColor = "#CCCCCC"
		elseif Ffinishflag="8" then
			GetStateColor = "#BBBBBB"
		else

		end if
	end function

    Private Sub Class_Initialize()
		Ftot_itemno = 0
		Ftot_orgsellprice = 0
		Ftot_realsellprice  = 0
        Ftot_jungsanprice = 0
        FTW_price = 0
        FUW_price = 0
        FCM_price = 0
        FOM_price = 0
        FSM_price = 0
        FET_price = 0
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class

class COffJungsan
	public FItemList()
	public FOneItem
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FTotalSum
	public FRectYYYYMM
	public FRectMakerid
	public FRectIdx
	public FRectGubunCd
	public FRectShopid
	public FRectfinishflag
	public FRectTaxtype
	public FRectAutojungsan
    public FRectJungsanDate
    public FRectBankingUpFlag
    public FRectGroupid
    public FRectSOCNO
    '' FRectStartYYYYMM<= RECT <=FRectEndYYYYMM
    public FRectStartYYYYMM
    public FRectEndYYYYMM
    '' FRectStartYYYYMMDD<= RECT <FRectEndYYYYMMDD
    public FRectStartYYYYMMDD
    public FRectEndYYYYMMDD
    public FRectFixStateExiste
    public FRectNotIncludeWonChon
    public FRectOnlyIncludeWonChon
    public FRectOnlyIncludeKani
    public FRectNotYYYYMM
    public FRectTaxRegDate
    public FRectIpkumDate
    public FRectOffgubun
    public FRectMinusGubnu
    public FRectbankingupFile
    public FRectPurchaseType
    public FRectJungsanGubunCD
    public FRectJungsanGubun

    public FRectJGubun
    public FRectjacctcd
    public FRectdifferencekey

	public FRectSearchType
	public FRectSearchText

    public function JungsanFixedList()
		dim sqlStr,i, sqlAdd, bufSqlStr

		sqlStr = "select m.*, "
		sqlStr = sqlStr + " p.jungsan_email,p.jungsan_bank,p.jungsan_date_off,p.jungsan_acctno,"
		sqlStr = sqlStr + " p.jungsan_acctname,p.company_name, p.jungsan_gubun,p.company_no"

		IF (FRectbankingupFile<>"") then
		    sqlStr = sqlStr + " ,ip.ipFileNo, ip.targetGbn"
		ELSE
		    sqlStr = sqlStr + " ,NULL as ipFileNo, NULL as targetGbn"
		End IF
		sqlStr = sqlStr + " ,HH.groupid as holdGroupid, HH.holdcause"

		sqlAdd = ""
		sqlAdd = sqlAdd + " from [db_jungsan].[dbo].tbl_off_jungsan_master m"
		sqlAdd = sqlAdd + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"

		sqlAdd = sqlAdd + "     left join (select g.groupid, H.holdcause from db_partner.dbo.tbl_partner_group g "
        sqlAdd = sqlAdd + "                 Join db_jungsan.dbo.tbl_jungsan_hold H"
        sqlAdd = sqlAdd + "                 on replace(g.company_no,'-','')=H.holdsocno) as HH"
        sqlAdd = sqlAdd + "     on m.groupid=HH.groupid"

		if FRectbankingupFile<>"" then
		    sqlAdd = sqlAdd + "     left join db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail IP"
		    sqlAdd = sqlAdd + "     on Ip.targetGbn='OF' and Ip.targetIdx=m.idx"
	    end if

		if FRectfinishflag="ALL" then
		    sqlAdd = sqlAdd + " where m.finishflag>=3"
		elseif FRectfinishflag<>"" then
		    sqlAdd = sqlAdd + " where m.finishflag='" + FRectfinishflag + "'"
		else
		    sqlAdd = sqlAdd + " where m.finishflag='3'"
        end if

        if (FRectIpkumDate<>"") then
            sqlAdd = sqlAdd + " and m.ipkumdate='" + FRectIpkumDate + "'"
        end if

        if (FRectTaxRegDate<>"") then
            sqlAdd = sqlAdd + " and m.taxregdate='" + FRectTaxRegDate + "'"
        end if

        if (FRectJGubun<>"") then
            sqlAdd = sqlAdd + " and m.jgubun='"&FRectJGubun&"'"
        end if
        '' AA 전월 정산내역 중 발행일이 전월 & 정산일 수시/15일
        '' BB 전월 정산내역 중 발행일이 전월 & 정산일 말일
        '' CC 전전월 이하 정산내역 중 발행일이 전월
        '' DD 발행일이 현재월 이상
        '' EE 정상발행 전체
        '' FF 이월발행 전체 (비정상발행)
        '' ZZ 발행일이 빈값이거나, 그 외 날짜
        if FRectGubunCd="ZZ" then
            sqlAdd = sqlAdd + " and m.taxregdate is NULL"
        elseif FRectGubunCd="SS" then
            sqlAdd = sqlAdd + " and (p.jungsan_date_off='수시')"           ''수시
        elseif FRectGubunCd="AA" then
            sqlAdd = sqlAdd + " and ((p.jungsan_date_off='수시')"           ''수시추가
            sqlAdd = sqlAdd + " or ((IsNULL(p.jungsan_date_off,'')='' or p.jungsan_date_off<>'말일')"
            ''sqlAdd = sqlAdd + "     and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
            sqlAdd = sqlAdd + "))"
        elseif FRectGubunCd="BB" then
            sqlAdd = sqlAdd + " and ((p.jungsan_date_off='수시')"
            sqlAdd = sqlAdd + " or ((p.jungsan_date_off='말일')"
            sqlAdd = sqlAdd + "     and m.yyyymm=convert(varchar(7),m.taxregdate,21)))"
        elseif FRectGubunCd="CC" then
            sqlAdd = sqlAdd + " and m.yyyymm<convert(varchar(7),m.taxregdate,21)"
            sqlAdd = sqlAdd + " and convert(varchar(7),getdate(),21)>convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubunCd="DD" then
            sqlAdd = sqlAdd + " and convert(varchar(7),getdate(),21)<=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubunCd="EE" then
            sqlAdd = sqlAdd + " and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubunCd="FF" then
            sqlAdd = sqlAdd + " and m.yyyymm<>convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubunCd="NN" then
            sqlAdd = sqlAdd + " and m.yyyymm='"&LEFT(now(),7)&"'"
        end if

        if FRectJungsanDate="NULL" then
            sqlAdd = sqlAdd + " and IsNULL(p.jungsan_date_off,'')=''"
        elseif FRectJungsanDate<>"" then
            sqlAdd = sqlAdd + " and p.jungsan_date_off='" + FRectJungsanDate + "'"
        end if

        if FRectNotIncludeWonChon<>"" then
			''sqlAdd = sqlAdd + " and p.jungsan_gubun<>'원천징수'"   ''2017/11/01 아래로 수정
			sqlAdd = sqlAdd + " and m.taxtype<>'03'"
			''sqlAdd = sqlAdd + " and p.jungsan_gubun<>'간이과세'"  ''주석처리 2018/02/13
		end if

        if (FRectOnlyIncludeKani="on") then
            sqlAdd = sqlAdd + " and p.jungsan_gubun='간이과세'"
        end if

		if FRectOnlyIncludeWonChon<>"" then
			'sqlAdd = sqlAdd + " and p.jungsan_gubun='원천징수'"
			sqlAdd = sqlAdd + " and m.taxtype='03'"
		end if

		if FRectbankingupflag<>"" then
		    sqlAdd = sqlAdd + " and m.bankingupflag='" + FRectBankingupflag + "'"
		end if

		if FRectYYYYMM<>"" then
			sqlAdd = sqlAdd + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectNotYYYYMM<>"" then
			sqlAdd = sqlAdd + " and m.yyyymm<>'" + FRectNotYYYYMM + "'"
		end if

        if FRectMakerid<>"" then
			sqlAdd = sqlAdd + " and m.makerid='" + FRectMakerid + "'"
		end if

		if FRectGroupid<>"" then
			sqlAdd = sqlAdd + " and p.groupid='" + FRectGroupid + "'"
		end if

		bufSqlStr = sqlAdd

		if FRectMinusGubnu="MI" then
            sqlAdd = sqlAdd + " and m.groupid in ("
            sqlAdd = sqlAdd + "     select m.groupid " + bufSqlStr
            sqlAdd = sqlAdd + "     and m.tot_jungsanprice<1"
            sqlAdd = sqlAdd + " )"
            sqlAdd = sqlAdd + " order by (CASE WHEN IsNULL(m.billsitecode,'')='' THEN 'ZZZ' ELSE m.billsitecode END) desc, m.groupid, m.taxinputdate"
        ELSEif FRectMinusGubnu="MJ" then ''마이너스 제외
            sqlAdd = sqlAdd + "     and m.tot_jungsanprice>0"
            sqlAdd = sqlAdd + " order by (CASE WHEN IsNULL(m.billsitecode,'')='' THEN 'ZZZ' ELSE m.billsitecode END) desc, m.taxinputdate"
        ELSEif (FRectMinusGubnu="CX") or (FRectMinusGubnu="CX1") then
            sqlAdd = sqlAdd + " and m.groupid in ("
            sqlAdd = sqlAdd + "     select groupid from ("
            sqlAdd = sqlAdd + "         select m2.groupid"
            sqlAdd = sqlAdd + "         from  [db_jungsan].[dbo].tbl_off_jungsan_master m2"
            sqlAdd = sqlAdd + "         where m2.tot_jungsanprice<1 and m2.finishflag='3'"
            if (FRectJGubun<>"") then
                sqlAdd = sqlAdd + "     and m2.jgubun='"&FRectJGubun&"'"
            end if
            sqlAdd = sqlAdd + "         Union ALL "
            sqlAdd = sqlAdd + "         select m.groupid from  [db_jungsan].[dbo].tbl_designer_jungsan_master m"
            sqlAdd = sqlAdd + "         where m.ub_totalsuplycash+m.me_totalsuplycash+m.wi_totalsuplycash+m.et_totalsuplycash+m.sh_totalsuplycash+m.dlv_totalsuplycash<1"
            sqlAdd = sqlAdd + "         and m.finishflag='3'"
            if (FRectJGubun<>"") then
                sqlAdd = sqlAdd + "     and m.jgubun='"&FRectJGubun&"'"
            end if
            sqlAdd = sqlAdd + "     ) T0"
            sqlAdd = sqlAdd + "     group by T0.groupid"
            sqlAdd = sqlAdd + " )"
            sqlAdd = sqlAdd + " and m.groupid in ("
            sqlAdd = sqlAdd + "     select groupid from ("
            sqlAdd = sqlAdd + "         select m.groupid, 0 as jSum1, m.tot_jungsanprice as jSum2 "
            sqlAdd = sqlAdd + "         from  [db_jungsan].[dbo].tbl_off_jungsan_master m"
            sqlAdd = sqlAdd + "         where m.finishflag='3'"
            if (FRectJGubun<>"") then
                sqlAdd = sqlAdd + "     and m.jgubun='"&FRectJGubun&"'"
            end if
            sqlAdd = sqlAdd + "         Union ALL"
            sqlAdd = sqlAdd + "         select m2.groupid, m2.ub_totalsuplycash+m2.me_totalsuplycash+m2.wi_totalsuplycash+m2.et_totalsuplycash+m2.sh_totalsuplycash+m2.dlv_totalsuplycash as jSum1, 0 as jSum2 "
            sqlAdd = sqlAdd + "         from  [db_jungsan].[dbo].tbl_designer_jungsan_master m2"
            sqlAdd = sqlAdd + "         where m2.finishflag='3'"
            if (FRectJGubun<>"") then
                sqlAdd = sqlAdd + "     and m2.jgubun='"&FRectJGubun&"'"
            end if
            sqlAdd = sqlAdd + "     ) T"
            sqlAdd = sqlAdd + "     group by T.groupid"
            if (FRectMinusGubnu="CX") then
                sqlAdd = sqlAdd + "     having sum(T.jSum1+T.jSum2)>0 and sum(T.jSum2)>0 and (sum(T.jSum1)<1 or sum(CASE WHEN T.jSum1<0 then 1 ELSE 0 END)=0)"
            elseif (FRectMinusGubnu="CX1") then
                sqlAdd = sqlAdd + "     having sum(T.jSum1+T.jSum2)>0 and sum(T.jSum1)>0 and sum(T.jSum2)<1"
            end if
            sqlAdd = sqlAdd + "     "
            sqlAdd = sqlAdd + " )"
            sqlAdd = sqlAdd + " order by (CASE WHEN IsNULL(m.billsitecode,'')='' THEN 'ZZZ' ELSE m.billsitecode END) desc, m.groupid, m.taxinputdate"
        ELSE
           sqlAdd = sqlAdd + " order by (CASE WHEN IsNULL(m.billsitecode,'')='' THEN 'ZZZ' ELSE m.billsitecode END) desc, m.taxinputdate"
        end if

        sqlStr = sqlStr + sqlAdd

		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffJungsanMasterItem

				FItemList(i).Fidx               = rsget("idx")
                FItemList(i).Fyyyymm            = rsget("yyyymm")
                FItemList(i).Fdifferencekey     = rsget("differencekey")
                FItemList(i).Ftaxtype           = rsget("taxtype")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Ftitle             = db2html(rsget("title"))
                FItemList(i).Ftot_itemno        = rsget("tot_itemno")
                FItemList(i).Ftot_orgsellprice  = rsget("tot_orgsellprice")
                FItemList(i).Ftot_realsellprice = rsget("tot_realsellprice")
                FItemList(i).Ftot_jungsanprice  = rsget("tot_jungsanprice")
                FItemList(i).Fcomment           = db2html(rsget("comment"))
                FItemList(i).Ffinishflag        = rsget("finishflag")
                FItemList(i).Fipkumdate         = rsget("ipkumdate")
                FItemList(i).Ftaxregdate        = rsget("taxregdate")
                FItemList(i).Ftaxinputdate      = rsget("taxinputdate")
                FItemList(i).Ftaxlinkidx        = rsget("taxlinkidx")
                FItemList(i).Fneotaxno          = rsget("neotaxno")
                FItemList(i).Fbankingupflag     = rsget("bankingupflag")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).FTW_price          = rsget("TW_price")
                FItemList(i).FUW_price          = rsget("UW_price")
                FItemList(i).FCM_price          = rsget("CM_price")
                FItemList(i).FOM_price          = rsget("OM_price")
                FItemList(i).FSM_price          = rsget("SM_price")
                FItemList(i).FET_price          = rsget("ET_price")
                FItemList(i).Fjungsan_email     = db2html(rsget("jungsan_email"))
                FItemList(i).Fjungsan_bank      = rsget("jungsan_bank")
                FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")
                FItemList(i).Fjungsan_acctno    = rsget("jungsan_acctno")
                FItemList(i).Fjungsan_acctname  = db2html(rsget("jungsan_acctname"))
                FItemList(i).Fcompany_name      = db2html(rsget("company_name"))
                FItemList(i).Fjungsan_gubun     = rsget("jungsan_gubun")
                FItemList(i).Fcompany_no        = rsget("company_no")

                FItemList(i).Fipkum_bank      = rsget("ipkum_bank")
                FItemList(i).Fipkum_acctno    = rsget("ipkum_acctno")

                FItemList(i).FBillsiteCode  = rsget("BillsiteCode")

                FItemList(i).FipFileNo = rsget("ipFileNo")
                FItemList(i).FtargetGbn= rsget("targetGbn")

                FItemList(i).FholdGroupid   = rsget("holdGroupid")
                FItemList(i).Fholdcause     = rsget("holdcause")

                FItemList(i).Fjgubun            = rsget("jgubun")
                FItemList(i).Ftotalcommission   = rsget("totalcommission")
                FItemList(i).FitemvatYn         = rsget("itemvatYn")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

    end function

    public Sub GetOffJungsanSummaryBySegumDate()
        dim sqlStr, i

        sqlStr = " select m.taxregdate," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and (g.jungsan_date_off='수시') then tot_jungsanprice else 0 end) as jungsansum_susi," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and (g.jungsan_date_off='말일') then tot_jungsanprice else 0 end) as jungsansum_31date," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and (g.jungsan_date_off='15일') then tot_jungsanprice else 0 end) as jungsansum_15date," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and ((g.jungsan_date_off is NULL) or (g.jungsan_date_off not in('수시','말일','15일'))) then tot_jungsanprice else 0 end) as jungsansum_etcdate," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm<>convert(varchar(7),m.taxregdate,21))  then tot_jungsanprice else 0 end) as ewol_jungsansum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') then tot_jungsanprice else 0 end) as fixedsum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='7') then tot_jungsanprice else 0 end) as ipkumsum," + VbCrlf
        sqlStr = sqlStr + " sum(tot_jungsanprice) as tot_jungsanprice" + VbCrlf
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m" + VbCrlf
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group g " + VbCrlf
        sqlStr = sqlStr + "     on m.groupid=g.groupid" + VbCrlf
        sqlStr = sqlStr + " where m.finishflag >=3" + VbCrlf

        if (FRectStartYYYYMMDD<>"") then
            sqlStr = sqlStr + " and m.taxregdate>='" + FRectStartYYYYMMDD + "'" + VbCrlf
        end if

        if (FRectEndYYYYMMDD<>"") then
            sqlStr = sqlStr + " and m.taxregdate<'" + FRectEndYYYYMMDD + "'" + VbCrlf
        end if

        if (FRectTaxType<>"") then
            sqlStr = sqlStr + " and m.taxtype='" & FRectTaxType & "'" + VbCrlf
        end if

        sqlStr = sqlStr + " group by m.taxregdate" + VbCrlf
        sqlStr = sqlStr + " order by m.taxregdate desc " + VbCrlf

        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount
        FTotalCount = FResultCount

        if FResultCount<1 then FResultCount=0
        redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
		    i = 0
		    rsget.absolutepage = FCurrPage
		    do until rsget.eof

			set FItemList(i) = new COffJungsanSummaryByTaxDateItem

            FItemList(i).Ftaxregdate         = rsget("taxregdate")
            FItemList(i).Fjungsansum_susi    = rsget("jungsansum_susi")
            FItemList(i).Fjungsansum_31date  = rsget("jungsansum_31date")
            FItemList(i).Fjungsansum_15date  = rsget("jungsansum_15date")
            FItemList(i).Fjungsansum_etcdate = rsget("jungsansum_etcdate")
            FItemList(i).Fewol_jungsansum    = rsget("ewol_jungsansum")
            FItemList(i).Ffixedsum          = rsget("fixedsum")
            FItemList(i).Fipkumsum          = rsget("ipkumsum")
            FItemList(i).Ftot_jungsanprice  = rsget("tot_jungsanprice")

			rsget.MoveNext
			i = i + 1
		loop
	    end if
        rsget.Close
    end Sub

    '//admin/offupchejungsan/off_jungsansumury.asp
    public Sub GetOffJungsanSummary()
        dim sqlStr, i

        sqlStr = " select  m.yyyymm, g.jungsan_date_off," + VbCrlf
        sqlStr = sqlStr + " sum(TW_price) as TW_price," + VbCrlf
        sqlStr = sqlStr + " sum(UW_price) as UW_price," + VbCrlf
        sqlStr = sqlStr + " sum(CM_price) as CM_price," + VbCrlf
        sqlStr = sqlStr + " sum(OM_price) as OM_price," + VbCrlf
        sqlStr = sqlStr + " sum(SM_price) as SM_price," + VbCrlf
        sqlStr = sqlStr + " sum(ET_price) as ET_price," + VbCrlf
        sqlStr = sqlStr + " sum(case when m.finishflag='7' then tot_jungsanprice else 0 end) as ipkumsum," + VbCrlf
        sqlStr = sqlStr + " sum(case when m.finishflag='3' then tot_jungsanprice else 0 end) as fixedsum," + VbCrlf

        ''금월 기준으로 입금예정금액 산출.
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (convert(varchar(7),getdate(),21)>convert(varchar(7),taxregdate,21))  then tot_jungsanprice else 0 end) as fixedthissum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (convert(varchar(7),getdate(),21)<=convert(varchar(7),taxregdate,21))  then tot_jungsanprice else 0 end) as fixednextsum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag <'3') then tot_jungsanprice else 0 end) as waitsum," + VbCrlf
        sqlStr = sqlStr + " sum(tot_jungsanprice) as tot_jungsanprice " + VbCrlf
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m" + VbCrlf
        sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_group g " + VbCrlf
        sqlStr = sqlStr + " on m.groupid=g.groupid" + VbCrlf
        sqlStr = sqlStr + " where 1=1" + VbCrlf

        if (FRectStartYYYYMM<>"") then
            sqlStr = sqlStr + " and m.yyyymm>='" + FRectStartYYYYMM + "'" + VbCrlf
        end if

        if (FRectEndYYYYMM<>"") then
            sqlStr = sqlStr + " and m.yyyymm<='" + FRectEndYYYYMM + "'" + VbCrlf
        end if

        sqlStr = sqlStr + " group by m.yyyymm, g.jungsan_date_off" + VbCrlf

        if (FRectFixStateExiste<>"") then
            ''미처리 내역이 있는것..
            sqlStr = sqlStr + " having sum(case when m.finishflag<=3 then tot_jungsanprice else 0 end)<>0"
        end if
        sqlStr = sqlStr + " order by m.yyyymm desc, g.jungsan_date_off " + VbCrlf

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount
        FTotalCount = FResultCount

        if FResultCount<1 then FResultCount=0

        redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
		    i = 0
		    rsget.absolutepage = FCurrPage
		    do until rsget.eof

			set FItemList(i) = new COffJungsanSummaryItem

            FItemList(i).Fyyyymm            = rsget("yyyymm")
            FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")
            FItemList(i).FTW_price          = rsget("TW_price")
            FItemList(i).FUW_price          = rsget("UW_price")
            FItemList(i).FCM_price          = rsget("CM_price")
            FItemList(i).FOM_price          = rsget("OM_price")
            FItemList(i).FSM_price          = rsget("SM_price")
            FItemList(i).FET_price          = rsget("ET_price")
            FItemList(i).Fipkumsum          = rsget("ipkumsum")
            FItemList(i).Ffixedsum          = rsget("fixedsum")
            FItemList(i).Ffixedthissum      = rsget("fixedthissum")
            FItemList(i).Ffixednextsum      = rsget("fixednextsum")
            FItemList(i).Fwaitsum           = rsget("waitsum")
            FItemList(i).Ftot_jungsanprice  = rsget("tot_jungsanprice")

			rsget.MoveNext
			i = i + 1
		loop

		end if

        rsget.Close

    end Sub

    public Sub GetOneOffJungsanMaster()
        dim sqlStr

        sqlStr = "select top 1 m.*, "
        sqlStr = sqlStr + " p.jungsan_email,p.jungsan_bank,p.jungsan_date_off,p.jungsan_acctno"
		sqlStr = sqlStr + " ,p.jungsan_acctname,p.company_name, p.jungsan_gubun,p.company_no, b.billSiteName "
		sqlStr = sqlStr + " ,c.acc_nm"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlStr = sqlStr + "     left join db_jungsan.dbo.tbl_tax_asp_Info b on m.billsiteCode=b.BillSiteCode"
		sqlStr = sqlStr + "     left join db_partner.dbo.tbl_TMS_SL_ACC_CD c on m.jacctcd=c.acc_Use_cd"
        sqlStr = sqlStr + " where m.idx=" + CStr(FRectIdx)
        if FRectMakerid<>"" then
            sqlStr = sqlStr + " and m.makerid='" + FRectMakerid + "'"
        end if

        rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        FTotalCount = FResultCount

        if FResultCount<1 then FResultCount=0

		if  not rsget.EOF  then
			set FOneItem = new COffJungsanMasterItem

			FOneItem.Fidx               = rsget("idx")
            FOneItem.Fyyyymm            = rsget("yyyymm")
			FOneItem.FtargetGbn			= "OF"
            FOneItem.Fdifferencekey     = rsget("differencekey")
            FOneItem.Ftaxtype           = rsget("taxtype")
            FOneItem.Fmakerid           = rsget("makerid")
            FOneItem.Ftitle             = db2html(rsget("title"))
            FOneItem.Ftot_itemno        = rsget("tot_itemno")
            FOneItem.Ftot_orgsellprice  = rsget("tot_orgsellprice")
            FOneItem.Ftot_realsellprice = rsget("tot_realsellprice")
            FOneItem.Ftot_jungsanprice  = rsget("tot_jungsanprice")
            FOneItem.Fcomment           = db2html(rsget("comment"))
            FOneItem.Ffinishflag        = rsget("finishflag")
            FOneItem.Fipkumdate         = rsget("ipkumdate")
            FOneItem.Ftaxregdate        = rsget("taxregdate")
            FOneItem.Ftaxinputdate      = rsget("taxinputdate")
            FOneItem.Ftaxlinkidx        = rsget("taxlinkidx")
            FOneItem.Fneotaxno          = rsget("neotaxno")
            FOneItem.Fbankingupflag     = rsget("bankingupflag")
            FOneItem.Fregdate           = rsget("regdate")
            FOneItem.FTW_price          = rsget("TW_price")
            FOneItem.FUW_price          = rsget("UW_price")
            FOneItem.FCM_price          = rsget("CM_price")
            FOneItem.FOM_price          = rsget("OM_price")
            FOneItem.FSM_price          = rsget("SM_price")
            FOneItem.FET_price          = rsget("ET_price")
            FOneItem.Fjungsan_email     = db2html(rsget("jungsan_email"))
            FOneItem.Fjungsan_bank      = rsget("jungsan_bank")
            FOneItem.Fjungsan_date_off  = rsget("jungsan_date_off")
            FOneItem.Fjungsan_acctno    = rsget("jungsan_acctno")
            FOneItem.Fjungsan_acctname  = db2html(rsget("jungsan_acctname"))
            FOneItem.Fcompany_name      = db2html(rsget("company_name"))
            FOneItem.Fjungsan_gubun     = rsget("jungsan_gubun")
            FOneItem.Fcompany_no        = rsget("company_no")
            FOneItem.FGroupid           = rsget("groupid")
            FOneItem.Favailneo           = rsget("availneo")

            FOneItem.FBillsiteCode      = rsget("BillsiteCode")
            FOneItem.FISSU_SEQNO        = rsget("ISSU_SEQNO")
            FOneItem.FeseroEvalSeq      = rsget("eseroEvalSeq")
            FOneItem.FbillSiteName      = rsget("billSiteName")

            FOneItem.FrefPayreqIdx      = rsget("refPayreqIdx")
            FOneItem.FpreFixedTaxDate   = rsget("preFixedTaxDate")          ''2012-09-03 추가

            FOneItem.Fjgubun            = rsget("jgubun")
            FOneItem.Ftotalcommission   = rsget("totalcommission")
            FOneItem.FitemvatYn         = rsget("itemvatYn")
            FOneItem.Fjacctcd           = rsget("jacctcd")
            FOneItem.Fjacc_nm           = rsget("acc_nm")
		end if
		rsget.close

    end Sub

    public Sub GetOffJungsanMasterListBrandView()
        dim sqlStr, i

        sqlStr = "select count(m.idx) as cnt, IsNULL(sum(m.tot_jungsanprice),0) as totsum "
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m "
        sqlStr = sqlStr + " where makerid='" + FRectMakerid + "'"
        sqlStr = sqlStr + " and m.finishflag>0"
        sqlStr = sqlStr + " and m.finishflag<8"

        if FRectIdx<>"" then
            sqlStr = sqlStr + " and m.idx=" + CStr(FRectIdx)
        end if

        if FRectYYYYMM<>"" then
            sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
        end if

        if FRectfinishflag<>"" then
            sqlStr = sqlStr + " and m.finishflag='" + FRectfinishflag + "'"
        end if

        if FRectTaxtype<>"" then
            sqlStr = sqlStr + " and m.taxtype='" + FRectTaxtype + "'"
        end if

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalSum   = rsget("totsum")
		rsget.close

        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.*, "
        sqlStr = sqlStr + " p.jungsan_email,p.jungsan_bank,p.jungsan_date_off,p.jungsan_acctno,"
		sqlStr = sqlStr + " p.jungsan_acctname,p.company_name, p.jungsan_gubun,p.company_no "
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m"
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p "
        sqlStr = sqlStr + "     on m.groupid=p.groupid"
        sqlStr = sqlStr + " where makerid='" + FRectMakerid + "'"
        sqlStr = sqlStr + " and m.finishflag>0"

        if FRectIdx<>"" then
            sqlStr = sqlStr + " and m.idx=" + CStr(FRectIdx)
        end if

        if FRectYYYYMM<>"" then
            sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
        end if

        if FRectfinishflag<>"" then
            sqlStr = sqlStr + " and m.finishflag='" + FRectfinishflag + "'"
        end if

        if FRectTaxtype<>"" then
            sqlStr = sqlStr + " and m.taxtype='" + FRectTaxtype + "'"
        end if

        sqlStr = sqlStr + " order by m.yyyymm desc,m.makerid, m.idx desc"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffJungsanMasterItem

				FItemList(i).Fidx               = rsget("idx")
                FItemList(i).Fyyyymm            = rsget("yyyymm")
                FItemList(i).Fdifferencekey     = rsget("differencekey")
                FItemList(i).Ftaxtype           = rsget("taxtype")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Ftitle             = db2html(rsget("title"))
                FItemList(i).Ftot_itemno        = rsget("tot_itemno")
                FItemList(i).Ftot_orgsellprice  = rsget("tot_orgsellprice")
                FItemList(i).Ftot_realsellprice = rsget("tot_realsellprice")
                FItemList(i).Ftot_jungsanprice  = rsget("tot_jungsanprice")
                FItemList(i).Fcomment           = db2html(rsget("comment"))
                FItemList(i).Ffinishflag        = rsget("finishflag")
                FItemList(i).Fipkumdate         = rsget("ipkumdate")
                FItemList(i).Ftaxregdate        = rsget("taxregdate")
                FItemList(i).Ftaxinputdate      = rsget("taxinputdate")
                FItemList(i).Ftaxlinkidx        = rsget("taxlinkidx")
                FItemList(i).Fneotaxno          = rsget("neotaxno")
                FItemList(i).Fbankingupflag     = rsget("bankingupflag")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).FTW_price          = rsget("TW_price")
                FItemList(i).FUW_price          = rsget("UW_price")
                FItemList(i).FCM_price          = rsget("CM_price")
                FItemList(i).FOM_price          = rsget("OM_price")
                FItemList(i).FSM_price          = rsget("SM_price")
                FItemList(i).FET_price          = rsget("ET_price")
                FItemList(i).Fjungsan_email     = db2html(rsget("jungsan_email"))
                FItemList(i).Fjungsan_bank      = rsget("jungsan_bank")
                FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")
                FItemList(i).Fjungsan_acctno    = rsget("jungsan_acctno")
                FItemList(i).Fjungsan_acctname  = db2html(rsget("jungsan_acctname"))
                FItemList(i).Fcompany_name      = db2html(rsget("company_name"))
                FItemList(i).Fjungsan_gubun     = rsget("jungsan_gubun")
                FItemList(i).Fcompany_no        = rsget("company_no")
                FItemList(i).FGroupid           = rsget("groupid")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

    end Sub

    public Sub GetOffJungsanMasterList()
        dim sqlStr, i

        sqlStr = "select count(m.idx) as cnt, IsNULL(sum(m.tot_jungsanprice),0) as totsum "

        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m "
        sqlStr = sqlStr + " 	inner join [db_partner].[dbo].tbl_partner pp on m.makerid = pp.id"
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group g "
        sqlStr = sqlStr + "     on m.groupid=g.groupid"
        sqlStr = sqlStr + " where 1=1"

        if FRectIdx<>"" then
            sqlStr = sqlStr + " and m.idx=" + CStr(FRectIdx)
        end if

        if (FRectJGubun<>"") then
            sqlStr = sqlStr + " and m.jgubun='" + FRectJGubun + "'"
        end if

		if FRectJungsanGubun<>"" then
			sqlStr = sqlStr + " and g.jungsan_gubun='" & FRectJungsanGubun & "'"
		end if

        if (FRectMakerid<>"") or (FRectGroupID<>"") or (FRectSOCNO<>"") then
            IF (FRectMakerid<>"") then
                sqlStr = sqlStr + " and m.makerid='" + FRectMakerid + "'"
            ELSEIF (FRectGroupID<>"") then
                sqlStr = sqlStr + " and m.groupid='" + FRectGroupID + "'"
            ELSE
                sqlStr = sqlStr + " and g.company_no='" + FRectSOCNO + "'"
            END IF

        else
            if FRectYYYYMM<>"" then
                sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
            end if

            if FRectfinishflag<>"" then
                sqlStr = sqlStr + " and m.finishflag='" + FRectfinishflag + "'"
            end if

            if FRectTaxtype<>"" then
                sqlStr = sqlStr + " and m.taxtype='" + FRectTaxtype + "'"
            end if


            if FRectJungsanDate<>"" then
                if FRectJungsanDate="NULL" then
                    sqlStr = sqlStr + " and g.jungsan_date_off is NULL"
                else
                    sqlStr = sqlStr + " and g.jungsan_date_off='" + FRectJungsanDate + "'"
                end if
            end if
        end if

		if FRectPurchaseType<>"" then
			sqlStr = sqlStr + " and pp.PurchaseType = '" + FRectPurchaseType + "'"
		end if

        if (FRectJungsanGubunCD<>"") then
            sqlStr = sqlStr + " and m.idx in ("
        	sqlStr = sqlStr + "     select distinct d.masteridx"
        	sqlStr = sqlStr + "     from [db_jungsan].[dbo].tbl_off_jungsan_detail d "
        	sqlStr = sqlStr + "     where m.idx=d.masteridx "
        	sqlStr = sqlStr + "     and d.gubuncd ='"&FRectJungsanGubunCD&"'"
            sqlStr = sqlStr + " )"
        end if

        if FRectjacctcd<>"" then
			sqlStr = sqlStr + " and m.jacctcd = '" + FRectjacctcd + "'"
		end if

		if (FRectdifferencekey<>"") then
		     sqlStr = sqlStr + " and m.differencekey = '" + FRectdifferencekey + "'"
		end if

		if (FRectSearchType <> "") and (FRectSearchText <> "") then
			Select Case FRectSearchType
				Case "socname"
					sqlStr = sqlStr + " and m.makerid in ( "
					sqlStr = sqlStr + " 	select distinct p1.id "
					sqlStr = sqlStr + " 	from "
					sqlStr = sqlStr + " 		db_partner.dbo.tbl_partner p1 "
					sqlStr = sqlStr + " 		join [db_partner].[dbo].tbl_partner_group g1 "
					sqlStr = sqlStr + " 		on "
					sqlStr = sqlStr + " 			p1.groupid = g1.groupid "
					sqlStr = sqlStr + " 	where "
					sqlStr = sqlStr + " 		1 = 1 "
					sqlStr = sqlStr + " 		and g1.company_name like '%" + CStr(FRectSearchText) + "%' "
					sqlStr = sqlStr + " 		and p1.isusing = 'Y' "
					sqlStr = sqlStr + " ) "
				Case "socno"
					sqlStr = sqlStr + " and m.makerid in ( "
					sqlStr = sqlStr + " 	select distinct p1.id "
					sqlStr = sqlStr + " 	from "
					sqlStr = sqlStr + " 		db_partner.dbo.tbl_partner p1 "
					sqlStr = sqlStr + " 		join [db_partner].[dbo].tbl_partner_group g1 "
					sqlStr = sqlStr + " 		on "
					sqlStr = sqlStr + " 			p1.groupid = g1.groupid "
					sqlStr = sqlStr + " 	where "
					sqlStr = sqlStr + " 		1 = 1 "
					sqlStr = sqlStr + " 		and g1.company_no = '" + CStr(FRectSearchText) + "' "
					sqlStr = sqlStr + " 		and p1.isusing = 'Y' "
					sqlStr = sqlStr + " ) "
				Case Else
					''
			End Select
		end if

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalSum   = rsget("totsum")
		rsget.close

        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.* "
        sqlStr = sqlStr + " , 'Y' as autojungsan,"
        sqlStr = sqlStr + " g.jungsan_email,g.jungsan_bank,g.jungsan_date_off,g.jungsan_acctno,"
		sqlStr = sqlStr + " g.jungsan_acctname,g.company_name, g.jungsan_gubun,g.company_no "
		sqlStr = sqlStr + " , c.acc_nm"
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m"
        sqlStr = sqlStr + " 	inner join [db_partner].[dbo].tbl_partner pp on m.makerid = pp.id"
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group g "
        sqlStr = sqlStr + "     on m.groupid=g.groupid"
        sqlStr = sqlStr + "     left join db_partner.dbo.tbl_TMS_SL_ACC_CD c on m.jacctcd=c.acc_Use_cd"
        sqlStr = sqlStr + " where 1=1"

        if FRectIdx<>"" then
            sqlStr = sqlStr + " and m.idx=" + CStr(FRectIdx)
        end if

        if (FRectJGubun<>"") then
            sqlStr = sqlStr + " and m.jgubun='" + FRectJGubun + "'"
        end if

		if FRectJungsanGubun<>"" then
			sqlStr = sqlStr + " and g.jungsan_gubun='" & FRectJungsanGubun & "'"
		end if

        if (FRectMakerid<>"") or (FRectGroupID<>"") or (FRectSOCNO<>"") then
            IF (FRectMakerid<>"") then
                sqlStr = sqlStr + " and m.makerid='" + FRectMakerid + "'"
            ELSEIF (FRectGroupID<>"") then
                sqlStr = sqlStr + " and m.groupid='" + FRectGroupID + "'"
            ELSE
                sqlStr = sqlStr + " and g.company_no='" + FRectSOCNO + "'"
            END IF

        else
            if FRectYYYYMM<>"" then
                sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
            end if

            if FRectfinishflag<>"" then
                sqlStr = sqlStr + " and m.finishflag='" + FRectfinishflag + "'"
            end if

            if FRectTaxtype<>"" then
                sqlStr = sqlStr + " and m.taxtype='" + FRectTaxtype + "'"
            end if

            if FRectJungsanDate<>"" then
                if FRectJungsanDate="NULL" then
                    sqlStr = sqlStr + " and g.jungsan_date_off is NULL"
                else
                    sqlStr = sqlStr + " and g.jungsan_date_off='" + FRectJungsanDate + "'"
                end if
            end if
        end if

		if FRectPurchaseType<>"" then
			sqlStr = sqlStr + " and pp.PurchaseType = '" + FRectPurchaseType + "'"
		end if

        if (FRectJungsanGubunCD<>"") then
            sqlStr = sqlStr + " and m.idx in ("
        	sqlStr = sqlStr + "     select distinct d.masteridx"
        	sqlStr = sqlStr + "     from [db_jungsan].[dbo].tbl_off_jungsan_detail d "
        	sqlStr = sqlStr + "     where m.idx=d.masteridx "
        	sqlStr = sqlStr + "     and d.gubuncd ='"&FRectJungsanGubunCD&"'"
            sqlStr = sqlStr + " )"
        end if

        if FRectjacctcd<>"" then
			sqlStr = sqlStr + " and m.jacctcd = '" + FRectjacctcd + "'"
		end if

		if (FRectdifferencekey<>"") then
		     sqlStr = sqlStr + " and m.differencekey = '" + FRectdifferencekey + "'"
		end if

		if (FRectSearchType <> "") and (FRectSearchText <> "") then
			Select Case FRectSearchType
				Case "socname"
					sqlStr = sqlStr + " and m.makerid in ( "
					sqlStr = sqlStr + " 	select distinct p1.id "
					sqlStr = sqlStr + " 	from "
					sqlStr = sqlStr + " 		db_partner.dbo.tbl_partner p1 "
					sqlStr = sqlStr + " 		join [db_partner].[dbo].tbl_partner_group g1 "
					sqlStr = sqlStr + " 		on "
					sqlStr = sqlStr + " 			p1.groupid = g1.groupid "
					sqlStr = sqlStr + " 	where "
					sqlStr = sqlStr + " 		1 = 1 "
					sqlStr = sqlStr + " 		and g1.company_name like '%" + CStr(FRectSearchText) + "%' "
					sqlStr = sqlStr + " 		and p1.isusing = 'Y' "
					sqlStr = sqlStr + " ) "
				Case "socno"
					sqlStr = sqlStr + " and m.makerid in ( "
					sqlStr = sqlStr + " 	select distinct p1.id "
					sqlStr = sqlStr + " 	from "
					sqlStr = sqlStr + " 		db_partner.dbo.tbl_partner p1 "
					sqlStr = sqlStr + " 		join [db_partner].[dbo].tbl_partner_group g1 "
					sqlStr = sqlStr + " 		on "
					sqlStr = sqlStr + " 			p1.groupid = g1.groupid "
					sqlStr = sqlStr + " 	where "
					sqlStr = sqlStr + " 		1 = 1 "
					sqlStr = sqlStr + " 		and g1.company_no = '" + CStr(FRectSearchText) + "' "
					sqlStr = sqlStr + " 		and p1.isusing = 'Y' "
					sqlStr = sqlStr + " ) "
				Case Else
					''
			End Select
		end if

        sqlStr = sqlStr + " order by m.yyyymm desc,m.makerid, m.idx desc"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffJungsanMasterItem

				FItemList(i).Fidx               = rsget("idx")
                FItemList(i).Fyyyymm            = rsget("yyyymm")
                FItemList(i).Fdifferencekey     = rsget("differencekey")
                FItemList(i).Ftaxtype           = rsget("taxtype")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Ftitle             = db2html(rsget("title"))
                FItemList(i).Ftot_itemno        = rsget("tot_itemno")
                FItemList(i).Ftot_orgsellprice  = rsget("tot_orgsellprice")
                FItemList(i).Ftot_realsellprice = rsget("tot_realsellprice")
                FItemList(i).Ftot_jungsanprice  = rsget("tot_jungsanprice")
                FItemList(i).Fcomment           = db2html(rsget("comment"))
                FItemList(i).Ffinishflag        = rsget("finishflag")
                FItemList(i).Fipkumdate         = rsget("ipkumdate")
                FItemList(i).Ftaxregdate        = rsget("taxregdate")
                FItemList(i).Ftaxinputdate      = rsget("taxinputdate")
                FItemList(i).Ftaxlinkidx        = rsget("taxlinkidx")
                FItemList(i).Fneotaxno          = rsget("neotaxno")
                FItemList(i).Fbankingupflag     = rsget("bankingupflag")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).FTW_price          = rsget("TW_price")
                FItemList(i).FUW_price          = rsget("UW_price")
                FItemList(i).FCM_price          = rsget("CM_price")
                FItemList(i).FOM_price          = rsget("OM_price")
                FItemList(i).FSM_price          = rsget("SM_price")
                FItemList(i).FET_price          = rsget("ET_price")
                FItemList(i).Fautojungsan       = rsget("autojungsan")
                FItemList(i).Fjungsan_email     = db2html(rsget("jungsan_email"))
                FItemList(i).Fjungsan_bank      = rsget("jungsan_bank")
                FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")
                FItemList(i).Fjungsan_acctno    = rsget("jungsan_acctno")
                FItemList(i).Fjungsan_acctname  = db2html(rsget("jungsan_acctname"))
                FItemList(i).Fcompany_name      = db2html(rsget("company_name"))
                FItemList(i).Fjungsan_gubun     = rsget("jungsan_gubun")
                FItemList(i).Fcompany_no        = rsget("company_no")
                FItemList(i).FGroupid           = rsget("groupid")

                FItemList(i).FBillsiteCode      = rsget("BillsiteCode")
                FItemList(i).FISSU_SEQNO        = rsget("ISSU_SEQNO")
                FItemList(i).FeseroEvalSeq      = rsget("eseroEvalSeq")
                '''FItemList(i).FbillSiteName      = rsget("billSiteName")

                ''2014 추가
                FItemList(i).FjGubun            = rsget("jGubun")  ''정산방식
                FItemList(i).Ftotalcommission   = rsget("totalcommission")  ''총 수수료
                FItemList(i).FitemvatYn         = rsget("itemvatYn")  ''상품 과세구분
                FItemList(i).Fjacctcd           = rsget("jacctcd")
                FItemList(i).Fjacc_nm           = rsget("acc_nm")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

    end Sub

    public Sub GetOneOffJungsanDetailSummary()
        dim sqlStr, i

        sqlStr = "select T.*, "
        sqlStr = sqlStr + " c.comm_name, u.shopname, " + VbCrlf
        sqlStr = sqlStr + " s.chargediv, s.defaultmargin, s.defaultsuplymargin, s.autojungsan, s.autojungsandiv" + VbCrlf
        sqlStr = sqlStr + " from ( select d.shopid, d.gubuncd,"
        sqlStr = sqlStr + " sum(d.itemno) as tot_itemno, " + VbCrlf
        sqlStr = sqlStr + " sum(d.sellprice*d.itemno) as tot_orgsellprice, " + VbCrlf
        sqlStr = sqlStr + " sum(d.realsellprice*d.itemno) as tot_realsellprice, " + VbCrlf
        sqlStr = sqlStr + " sum(d.suplyprice*d.itemno) as tot_jungsanprice " + VbCrlf
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_detail d" + VbCrlf
        sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectIdx)
        sqlStr = sqlStr + " and d.gubuncd='" + FRectGubunCd + "'"
        sqlStr = sqlStr + " and d.shopid='" + FRectShopId + "'"
        sqlStr = sqlStr + " group by d.shopid, d.gubuncd "
        sqlStr = sqlStr + " ) T"
        sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_jungsan_comm_code c " + VbCrlf
        sqlStr = sqlStr + "     on c.comm_group='Z002' and T.gubuncd=c.comm_cd " + VbCrlf
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_designer s " + VbCrlf
        sqlStr = sqlStr + "     on T.shopid=s.shopid and s.makerid='" + FRectMakerid + "'" + VbCrlf
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_user u " + VbCrlf
        sqlStr = sqlStr + "     on T.shopid=u.userid"

        sqlStr = sqlStr + " order by T.shopid, T.gubuncd"

        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount

		if  not rsget.EOF  then
			set FOneItem = new COffJungsanDetailSummaryItem

            FOneItem.Fshopid            = rsget("shopid")
            FOneItem.Fgubuncd           = rsget("gubuncd")
            FOneItem.Fcomm_name         = db2html(rsget("comm_name"))
            FOneItem.Fshopname          = db2html(rsget("shopname"))
            FOneItem.Ftot_itemno        = rsget("tot_itemno")
            FOneItem.Ftot_orgsellprice  = rsget("tot_orgsellprice")
            FOneItem.Ftot_realsellprice = rsget("tot_realsellprice")
            FOneItem.Ftot_jungsanprice  = rsget("tot_jungsanprice")
            FOneItem.Fchargediv         = rsget("chargediv")
            FOneItem.Fdefaultmargin     = rsget("defaultmargin")
            FOneItem.Fdefaultsuplymargin= rsget("defaultsuplymargin")
            FOneItem.Fautojungsan       = rsget("autojungsan")
            FOneItem.Fautojungsandiv    = rsget("autojungsandiv")

		end if
        rsget.Close
    end Sub

    public Sub GetOffJungsanDetailSummaryListByMonth()
        dim sqlStr, i
        sqlStr = "select T.*, "
        sqlStr = sqlStr + " c.comm_name, u.shopname, " + VbCrlf
        sqlStr = sqlStr + " s.chargediv, s.defaultmargin, s.defaultsuplymargin, s.autojungsan, s.autojungsandiv" + VbCrlf
        sqlStr = sqlStr + " from (" + VbCrlf
        sqlStr = sqlStr + "     select d.shopid, d.gubuncd, " + VbCrlf
        sqlStr = sqlStr + "     sum(d.itemno) as tot_itemno, " + VbCrlf
        sqlStr = sqlStr + "     sum(d.sellprice*d.itemno) as tot_orgsellprice, " + VbCrlf
        sqlStr = sqlStr + "     sum(d.realsellprice*d.itemno) as tot_realsellprice, " + VbCrlf
        sqlStr = sqlStr + "     sum(d.suplyprice*d.itemno) as tot_jungsanprice " + VbCrlf
        sqlStr = sqlStr + "     from [db_jungsan].[dbo].tbl_off_jungsan_detail d" + VbCrlf
        sqlStr = sqlStr + "         Join [db_jungsan].[dbo].tbl_off_jungsan_master m" + VbCrlf
        sqlStr = sqlStr + "         on m.yyyymm='"&FRectYYYYMM&"'" + VbCrlf
        sqlStr = sqlStr + "         and m.idx=d.masteridx"
        sqlStr = sqlStr + "         and m.makerid='"&FRectMakerid&"'"
        sqlStr = sqlStr + "     where 1=1"
        if (FRectShopid<>"") then
            sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
        end if
        if (FRectGubunCD<>"") then
            sqlStr = sqlStr + " and d.gubuncd='" + FRectGubunCD + "'"
        end if
        sqlStr = sqlStr + "     group by d.shopid, d.gubuncd "
        sqlStr = sqlStr + " ) T"
        sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_jungsan_comm_code c " + VbCrlf
        sqlStr = sqlStr + "     on c.comm_group='Z002' and T.gubuncd=c.comm_cd " + VbCrlf
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_designer s " + VbCrlf
        sqlStr = sqlStr + "     on T.shopid=s.shopid and s.makerid='" + FRectMakerid + "'" + VbCrlf
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_user u " + VbCrlf
        sqlStr = sqlStr + "     on T.shopid=u.userid"
        sqlStr = sqlStr + " order by T.shopid, T.gubuncd"

        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COffJungsanDetailSummaryItem

                FItemList(i).Fshopid            = rsget("shopid")
                FItemList(i).Fgubuncd           = rsget("gubuncd")
                FItemList(i).Fcomm_name         = db2html(rsget("comm_name"))
                FItemList(i).Fshopname          = db2html(rsget("shopname"))
                FItemList(i).Ftot_itemno        = rsget("tot_itemno")
                FItemList(i).Ftot_orgsellprice  = rsget("tot_orgsellprice")
                FItemList(i).Ftot_realsellprice = rsget("tot_realsellprice")
                FItemList(i).Ftot_jungsanprice  = rsget("tot_jungsanprice")
                FItemList(i).Fchargediv         = rsget("chargediv")
                FItemList(i).Fdefaultmargin     = rsget("defaultmargin")
                FItemList(i).Fdefaultsuplymargin= rsget("defaultsuplymargin")
                FItemList(i).Fautojungsan       = rsget("autojungsan")
                FItemList(i).Fautojungsandiv    = rsget("autojungsandiv")

				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.Close
    end Sub

    public Sub GetOffJungsanDetailSummaryList()
        dim sqlStr, i
        sqlStr = "select T.*, "
        sqlStr = sqlStr + " c.comm_name, u.shopname, " + VbCrlf
        sqlStr = sqlStr + " s.chargediv, s.defaultmargin, s.defaultsuplymargin, s.autojungsan, s.autojungsandiv" + VbCrlf
        sqlStr = sqlStr + " from (" + VbCrlf
        sqlStr = sqlStr + " select m.jgubun, d.shopid, d.gubuncd, m.itemvatYn" + VbCrlf
        sqlStr = sqlStr + " ,sum(d.itemno) as tot_itemno " + VbCrlf
        sqlStr = sqlStr + " ,sum(d.sellprice*d.itemno) as tot_orgsellprice " + VbCrlf
        sqlStr = sqlStr + " ,sum(d.realsellprice*d.itemno) as tot_realsellprice " + VbCrlf
        sqlStr = sqlStr + " ,sum(d.suplyprice*d.itemno) as tot_jungsanprice " + VbCrlf
        sqlStr = sqlStr + " ,sum(d.commission*d.itemno) as tot_commission " + VbCrlf
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_detail d" + VbCrlf
        sqlStr = sqlStr + "     Join [db_jungsan].[dbo].tbl_off_jungsan_master m" + VbCrlf
        sqlStr = sqlStr + "     on m.idx=d.masteridx" + VbCrlf
        sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectIdx)
        sqlStr = sqlStr + " group by m.jgubun, d.shopid, d.gubuncd, m.itemvatYn "
        sqlStr = sqlStr + " ) T"
        sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_jungsan_comm_code c " + VbCrlf
        sqlStr = sqlStr + "     on c.comm_group='Z002' and T.gubuncd=c.comm_cd " + VbCrlf
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_designer s " + VbCrlf
        sqlStr = sqlStr + "     on T.shopid=s.shopid and s.makerid='" + FRectMakerid + "'" + VbCrlf
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_user u " + VbCrlf
        sqlStr = sqlStr + "     on T.shopid=u.userid"
        sqlStr = sqlStr + " order by T.shopid, T.gubuncd"

        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COffJungsanDetailSummaryItem

                FItemList(i).Fshopid            = rsget("shopid")
                FItemList(i).Fgubuncd           = rsget("gubuncd")
                FItemList(i).Fcomm_name         = db2html(rsget("comm_name"))
                FItemList(i).Fshopname          = db2html(rsget("shopname"))
                FItemList(i).Ftot_itemno        = rsget("tot_itemno")
                FItemList(i).Ftot_orgsellprice  = rsget("tot_orgsellprice")
                FItemList(i).Ftot_realsellprice = rsget("tot_realsellprice")
                FItemList(i).Ftot_jungsanprice  = rsget("tot_jungsanprice")
                FItemList(i).Fchargediv         = rsget("chargediv")
                FItemList(i).Fdefaultmargin     = rsget("defaultmargin")
                FItemList(i).Fdefaultsuplymargin= rsget("defaultsuplymargin")
                FItemList(i).Fautojungsan       = rsget("autojungsan")
                FItemList(i).Fautojungsandiv    = rsget("autojungsandiv")
                FItemList(i).Ftot_commission    = rsget("tot_commission")
                FItemList(i).Fjgubun            = rsget("jgubun")
                FItemList(i).FitemvatYn         = rsget("itemvatYn")
				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.Close
    end Sub

    public Sub GetOffJungsanDetailListByMonth()
        dim sqlStr, i

        sqlStr = "select count(d.detailidx) as cnt from [db_jungsan].[dbo].tbl_off_jungsan_detail d"
        sqlStr = sqlStr + "         Join [db_jungsan].[dbo].tbl_off_jungsan_master m" + VbCrlf
        sqlStr = sqlStr + "         on m.yyyymm='"&FRectYYYYMM&"'" + VbCrlf
        sqlStr = sqlStr + "         and m.idx=d.masteridx"
        sqlStr = sqlStr + "         and m.makerid='"&FRectMakerid&"'"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and d.gubuncd='" + FRectGubunCd + "'"
        if (FRectShopid<>"") then
            sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
        end if

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        sqlStr = "select Top " + CStr(FPageSize*FCurrPage) + " d.* "
        sqlStr = sqlStr + " , s.centermwdiv, s.vatinclude"
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_detail d" + VbCrlf
        sqlStr = sqlStr + "         Join [db_jungsan].[dbo].tbl_off_jungsan_master m" + VbCrlf
        sqlStr = sqlStr + "         on m.yyyymm='"&FRectYYYYMM&"'" + VbCrlf
        sqlStr = sqlStr + "         and m.idx=d.masteridx"
        sqlStr = sqlStr + "         and m.makerid='"&FRectMakerid&"'"
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_item s"
        sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun"
        sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
        sqlStr = sqlStr + "     and d.itemoption=s.itemoption"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and d.gubuncd='" + FRectGubunCd + "'"
        if (FRectShopid<>"") then
            sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
        end if
        sqlStr = sqlStr + " order by d.shopid, d.orderno, d.detailidx"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffJungsanDetailItem

                FItemList(i).Fdetailidx     = rsget("detailidx")
                FItemList(i).Fmasteridx     = rsget("masteridx")
                FItemList(i).Fshopid        = rsget("shopid")
                FItemList(i).Fgubuncd       = rsget("gubuncd")
                FItemList(i).Forderno       = rsget("orderno")
                FItemList(i).Fitemgubun     = rsget("itemgubun")
                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fitemoption    = rsget("itemoption")
                FItemList(i).Fitemname      = db2html(rsget("itemname"))
                FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
                FItemList(i).Forgsellprice  = rsget("sellprice")
                FItemList(i).Frealsellprice = rsget("realsellprice")
                FItemList(i).Fsuplyprice    = rsget("suplyprice")
                FItemList(i).Fitemno        = rsget("itemno")
                FItemList(i).Fmakerid       = rsget("makerid")
                FItemList(i).Flinkidx       = rsget("linkidx")
                FItemList(i).Fcentermwdiv   = rsget("centermwdiv")
                FItemList(i).Fvatinclude    = rsget("vatinclude")

				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.Close
    end Sub

    public Sub GetOffJungsanDetailList()
        dim sqlStr, i

        sqlStr = "select count(d.detailidx) as cnt from [db_jungsan].[dbo].tbl_off_jungsan_detail d"
        sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectIdx)
        sqlStr = sqlStr + " and d.gubuncd='" + FRectGubunCd + "'"
        if (FRectShopid<>"") then
            sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
        end if

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        sqlStr = "select Top " + CStr(FPageSize*FCurrPage) + " d.* "
        sqlStr = sqlStr + " , s.centermwdiv, s.vatinclude"
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_detail d" + VbCrlf
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_item s"
        sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun"
        sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
        sqlStr = sqlStr + "     and d.itemoption=s.itemoption"
        sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectIdx)
        sqlStr = sqlStr + " and d.gubuncd='" + FRectGubunCd + "'"
        if (FRectShopid<>"") then
            sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
        end if
        sqlStr = sqlStr + " order by d.shopid, d.orderno, d.detailidx"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffJungsanDetailItem

                FItemList(i).Fdetailidx     = rsget("detailidx")
                FItemList(i).Fmasteridx     = rsget("masteridx")
                FItemList(i).Fshopid        = rsget("shopid")
                FItemList(i).Fgubuncd       = rsget("gubuncd")
                FItemList(i).Forderno       = rsget("orderno")
                FItemList(i).Fitemgubun     = rsget("itemgubun")
                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fitemoption    = rsget("itemoption")
                FItemList(i).Fitemname      = db2html(rsget("itemname"))
                FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
                FItemList(i).Forgsellprice  = rsget("sellprice")
                FItemList(i).Frealsellprice = rsget("realsellprice")
                FItemList(i).Fsuplyprice    = rsget("suplyprice")
                FItemList(i).Fitemno        = rsget("itemno")
                FItemList(i).Fmakerid       = rsget("makerid")
                FItemList(i).Flinkidx       = rsget("linkidx")
                FItemList(i).Fcentermwdiv   = rsget("centermwdiv")
                FItemList(i).Fvatinclude    = rsget("vatinclude")  '' item

                FItemList(i).Fcommission    = rsget("commission")
                FItemList(i).Fiszerotax     = rsget("iszerotax")
                FItemList(i).Fpaymethod     = rsget("paymethod")
                FItemList(i).Fvatyn         = rsget("vatyn")

				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.Close
    end Sub

    public Sub GetOffJungsanDetailSumList()
        dim sqlStr, i

        sqlStr = "select Top " + CStr(FPageSize*FCurrPage) + " d.itemgubun, d.itemid, d.itemoption, itemname, itemoptionname, sellprice, realsellprice, commission, suplyprice ,sum(itemno)  as itemno"
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_detail d" + VbCrlf
        sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectIdx)
        if (FRectGubunCd<>"") then
            sqlStr = sqlStr + " and d.gubuncd='" + FRectGubunCd + "'"
        end if
        if (FRectShopid<>"") then
            sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
        end if
        sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption, itemname, itemoptionname, sellprice, realsellprice, commission, suplyprice"
        sqlStr = sqlStr + " order by d.itemgubun"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffJungsanDetailItem

                FItemList(i).Fitemgubun     = rsget("itemgubun")
                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fitemoption    = rsget("itemoption")
                FItemList(i).Fitemname      = db2html(rsget("itemname"))
                FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
                FItemList(i).Forgsellprice  = rsget("sellprice")
                FItemList(i).Frealsellprice = rsget("realsellprice")
                FItemList(i).Fcommission    = rsget("commission")
                FItemList(i).Fsuplyprice    = rsget("suplyprice")
                FItemList(i).Fitemno        = rsget("itemno")

				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.Close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 300
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		FTotalSum =0
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
end Class

function DrawOffJungsanStateCombo(selectBoxName,selectedId)
%>
    <select name="<%= selectBoxName %>" >
     <option value='' <%if selectedId="" then response.write " selected" %> >선택</option>
     <option value='0' <%if selectedId="0" then response.write " selected" %> >수정중</option>
	 <option value='1' <%if selectedId="1" then response.write " selected" %> >업체확인중</option>
	 <option value='2' <%if selectedId="2" then response.write " selected" %> >업체확인완료</option>
     <option value='3' <%if selectedId="3" then response.write " selected" %> >정산확정</option>
     <option value='7' <%if selectedId="7" then response.write " selected" %> >입금완료</option>
     <option value='8' <%if selectedId="8" then response.write " selected" %> >정산안함</option>
     <option value='9' <%if selectedId="9" then response.write " selected" %> >통합정산내역</option>
   </select>
<%
end function

function drawSelectBoxJungsanCommCombo(selectBoxName,selectedId,groupCode)
   dim tmp_str,sqlStr
   %>
     <select name="<%=selectBoxName%>" >
     <option value='' <%if selectedId="" then response.write " selected" %> >선택</option>
   <%
       sqlStr = " select comm_cd,comm_name "
       sqlStr = sqlStr + " from  "
       sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_jungsan_comm_code "
       sqlStr = sqlStr + " where comm_group='" + groupCode + "' "
       sqlStr = sqlStr + " and comm_isDel='N' "
       sqlStr = sqlStr + " order by comm_cd "

       rsget.Open sqlStr,dbget,1

       if  not rsget.EOF  then
           do until rsget.EOF
               if LCase(selectedId) = LCase(rsget("comm_cd")) then
                   tmp_str = " selected"
               end if
               response.write("<option value='" & rsget("comm_cd") & "' " & tmp_str & ">" + db2html(rsget("comm_name")) + " </option>")
               tmp_str = ""
               rsget.MoveNext
           loop
       end if
       rsget.close
   %>
       </select>
   <%
End function
%>
