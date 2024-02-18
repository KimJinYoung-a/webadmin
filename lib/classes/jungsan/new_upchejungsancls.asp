<%
function getDecSOCNobyGroupID(igroupid,chkcompNo)
    getDecSOCNobyGroupID = chkcompNo
    if (InStr(chkcompNo,"*")<1) then Exit function
    if LEN(replace(chkcompNo,"-",""))=10 then Exit function

    dim sqlStr
    ' sqlStr = "select [db_partner].[dbo].[uf_DecSOCNoPH1](encCompNo) as decCompNo from [db_partner].[dbo].tbl_partner_group g where groupid='"&igroupid&"'"
	' rsget.CursorLocation = adUseClient
    ' rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    ' if Not rsget.Eof then
	'     getDecSOCNobyGroupID = rsget("decCompNo")
    ' end if
    ' rsget.Close

	sqlStr = "select db_cs.[dbo].[uf_DecCompanyNoAES256](encCompNo64) as decCompNo64 "
	sqlStr = sqlStr & " from [db_partner].[dbo].tbl_partner_group_adddata g where groupid='"&igroupid&"'"
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
	    getDecSOCNobyGroupID = rsget("decCompNo64")
    end if
    rsget.Close
end function


''2014 추가
function getPartnerId2GroupID(ipartnerid)
    dim sqlStr
	sqlStr = "select groupid from db_partner.dbo.tbl_partner where id='"&ipartnerid&"'"

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    getPartnerId2GroupID = rsget("groupid")
    end if
    rsget.Close
end function

function chkAvailViewJungsanON(jid,makerid,groupid)
    dim sqlStr

    chkAvailViewJungsanON= false

    sqlStr = "select top 1 id"&VBCRLF
    sqlStr = sqlStr & " from db_jungsan.dbo.tbl_designer_jungsan_master"&VBCRLF
    sqlStr = sqlStr & " where id="&jid&VBCRLF
    sqlStr = sqlStr & " and finishflag>0"&VBCRLF
    sqlStr = sqlStr & " and (designerid='"&makerid&"' or groupid='"&groupid&"')"&VBCRLF
    rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    chkAvailViewJungsanON = true
    end if
    rsget.Close

end function

class CWitakJungSanItem
	public Fitemid
	public Fitemoption
	public Fitemname
	public FitemoptionName
	public FSellCash
	public FSuplycash
	public FSellcash_ipgo
	public FSuplycash_ipgo
	public FSellcash_sell
	public FSuplycash_sell
	public FIpGoNo
	public FChulgoNo
	public Fsellno
	public Fprejaego
	public Fprejaego2
	public Frealjaego
	public Frealjaego2
	public FoCha
	public FjungsanNo
	public FsysJaeGo

	public FIsUsing
	public FIsDelete
	public FDetailidx

	public FPrejaego_tmp
	public FIpGoNo_tmp
	public FChulGoNo_tmp
	public Frealjaego_tmp

	public FPreIdx
	public FCurrIdx
	public FPreMasterCode
	public FCurrMasterCode

	public Foffsellno

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end Class

''2014 추가
class CJungsanSubSummaryItem
    public Fjgubun
    public Fgubuncd
    public FgubuncdName
    public FtaxType
    public Fitemvatyn
    public FitemCNT
    public FsellcashSum
    public FsuplycashSum
    public FreducedpriceSum
    public FcommissionSum
    public FPgCommissionSum
    public FCpnNotAppliedPriceSum

    public function getSellcashSum
        getSellcashSum = FsellcashSum
    end function

    public function getCouponDiscountSum
        getCouponDiscountSum = (FsellcashSum-FreducedpriceSum)
    end function

    public function getReducedpriceSum
        getReducedpriceSum = FreducedpriceSum
    end function

    public function getCommissionSum
        getCommissionSum = FcommissionSum
    end function

    ''PG 수수료 2016/09/27
    public function getPGCommissionSum
        getPGCommissionSum = FPgCommissionSum
    end function

    public function getsuplycashSum
        getsuplycashSum = FsuplycashSum
    end function


    public function getJGubuncd2Name
        getJGubuncd2Name = FgubuncdName
    end function

    public function getJSummaryGugunName

        if IsCommissionTax then
            if (IsCommissionETCTax) then
                getJSummaryGugunName = "기타정산"
            else
                getJSummaryGugunName = "수수료정산"
            end if
        else
            if Fgubuncd="maeip" then
                getJSummaryGugunName = "입고분매입"
            elseif Fgubuncd="witaksell" then
                getJSummaryGugunName = "판매분매입"
            elseif Fgubuncd="upche" then
                getJSummaryGugunName = "판매분매입"
            elseif Fgubuncd="witakchulgo" then
                getJSummaryGugunName = "기타출고매입"
            else
                getJSummaryGugunName = "기타매입"
            end if
        end if
    end function

    public function getJGubunName
        if isNULL(Fjgubun) then
            getJGubunName = "매입정산"
        elseif Fjgubun="CC" then
            getJGubunName = "수수료정산"
        elseif Fjgubun="MM" then
            getJGubunName = "매입정산"
        elseif Fjgubun="CE" then
            getJGubunName = "기타정산"
        else
            getJGubunName = Fjgubun
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

    public function IsCommissionETCTax()  ''기타 매출 세금 계산서 인지.
        IsCommissionETCTax = false
        if isNULL(Fjgubun) then Exit function

        IsCommissionETCTax = (Fjgubun="CE")
    end function

    Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end Class


class CJungsanMasterItem
	public Fid
	public Fdesignerid
	public Fgroupid
	public Fyyyymm
	public Ftitle
	public Fub_cnt
	public Fub_totalsellcash
	public Fub_totalsuplycash
	public Fub_comment
	public Fme_cnt
	public Fme_totalsellcash
	public Fme_totalsuplycash
	public Fme_comment
	public Fwi_cnt
	public Fwi_totalsellcash
	public Fwi_totalsuplycash
	public Fwi_comment
	public Fet_cnt
	public Fet_totalsellcash
	public Fet_totalsuplycash
	public Fet_comment
	public Fsh_cnt
	public Fsh_totalsellcash
	public Fsh_totalsuplycash
	public Fsh_comment

	public Fregdate
	public Fcancelyn
	public Ffinishflag
	public Fipkumdate
	public Ftaxregdate
	public Fbigo

	public FDesignerEmail

	public Fjungsan_bank
	public Fjungsan_date
	public Fjungsan_acctno
	public Fjungsan_acctname
	public Fcompany_name
	public Fjungsan_gubun

	public Ftaxinputdate
	public Fcompany_no

	public Fdifferencekey
	public Ftaxtype
	public FTaxLinkidx
	public Fneotaxno

	public Fbankingupflag
    public Favailneo

    public Fipkum_bank
    public Fipkum_acctno
    public Fjungsan_hp
    public Fceoname
    public Fcompany_address
    public Fcompany_address2

    public FBillsiteCode
    public FISSU_SEQNO
    public FeseroEvalSeq
    public FbillSiteName
    public FipFileNo
    public FtargetGbn
    public FholdGroupid
    public Fholdcause
    public FpreFixedTaxDate

    ''2014/01/27 추가  수수료 매출 관련 =================================================
    public Fjgubun
    public Fwi_totalreducedprice
    public Fub_totalreducedprice
    public Fet_totalreducedprice
    public Fdlv_totalreducedprice
    public Fdlv_totalsuplycash
    public Ftotalcommission
    public FitemvatYn
    public Fjacctcd
    public Fjacc_nm

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

    public function getJGubunName
        if (FjGubun="MM") then
            getJGubunName = "매입"
        elseif (FjGubun="CC") then
            getJGubunName = "<font color=blue>수수료</font>"
        elseif (FjGubun="CE") then
            getJGubunName = "<font color=red>기타</font>"
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

    public function getTaxTypeName
        if (IsCommissionTax) then
            getTaxTypeName = "" ''수수료는 구분 없음.
        else
            if Ftaxtype="02" then
                getTaxTypeName = "<font color=red>면세<font>"
            elseif Ftaxtype="01" then
                getTaxTypeName = "과세"
            else
                getTaxTypeName = Ftaxtype
            end if
        end if
    end function

    public function getPrdMeachulSum() ''상품매출
        if (IsCommissionTax) then
            getPrdMeachulSum = CLNG(Fwi_totalreducedprice + Fub_totalreducedprice)
        else
            getPrdMeachulSum = CLNG(Fub_totalsuplycash + Fme_totalsuplycash + Fwi_totalsuplycash )
        end if
    end function

    public function getPrdCommissionSum() ''상품수수료
        if (IsCommissionTax) then
            getPrdCommissionSum = Ftotalcommission
        else
            getPrdCommissionSum = 0 ''
        end if
    end function

    public function getDlvMeachulSum() ''배송비매출
        if (IsCommissionTax) then
            getDlvMeachulSum = Fdlv_totalreducedprice
        else
            getDlvMeachulSum = Fdlv_totalsuplycash
        end if
    end function

    public function getEtcMeachulSum() ''기타매출
        if (IsCommissionTax) then
            getEtcMeachulSum = Fet_totalreducedprice
        else
            getEtcMeachulSum = Fet_totalsuplycash
        end if
    end function

    public function getPrdJungsanSum() ''상품정산액(지급예정액)
        getPrdJungsanSum = CLNG(Fub_totalsuplycash + Fme_totalsuplycash + Fwi_totalsuplycash )
    end function

    public function getDlvJungsanSum() ''배송비정산액(지급예정액)
        getDlvJungsanSum = CLNG(Fdlv_totalsuplycash)
    end function

    public function getEtcJungsanSum() ''기타정산액(지급예정액)
        getEtcJungsanSum = CLNG(Fet_totalsuplycash)
    end function

    public function getTotalJungsanSum() ''정산총액(지급예정액)
        getTotalJungsanSum = getPrdJungsanSum+getDlvJungsanSum+getEtcJungsanSum
    end function

    ''2014/01/27 추가  수수료 매출 관련 =================================================

	public function getDbDate()
		dim sqlstr
		sqlstr = " select convert(varchar(10),getdate(),21) as nowdate "
		rsget.Open sqlStr,dbget,1
		getDbDate = CDate(rsget("nowdate"))
		rsget.Close
	end function

    public function CheckTaxBillStatus(ijungsanid,ijungsangubun)
        dim sqlstr
		sqlstr = " [db_jungsan].[dbo].sp_Ten_JungsanStatusUpdateByIDX "&ijungsanid&",'"&ijungsangubun&"'"
'rw 	sqlstr
		dbget.Execute sqlstr
    end function

	public function GetNormalTaxDate()
		if Not(IsNULL(FpreFixedTaxDate)) and (FpreFixedTaxDate<>"") then
			GetNormalTaxDate = FpreFixedTaxDate
		else
			''if Fjungsan_date="말일" then
				GetNormalTaxDate = dateserial(Left(Fyyyymm,4),Right(Fyyyymm,2)+1,1-1)
			''else
			''	GetNormalTaxDate = dateserial(Left(Fyyyymm,4),Right(Fyyyymm,2)+1,1-2)
			''end if
		end if
	end function

	public function GetPreFixSegumil()
		dim thisdate, maytaxdate
		dim ithis1day , ithis21day, premonth1day, premonth21day

		thisdate = getDbDate()
		maytaxdate = GetNormalTaxDate()

        '' 12일까지 마감할 경우 13으로 세팅
        '' 10일까지 마감할 경우 11로 쎄팅
		premonth1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"01")
		premonth21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"11") ''11
		ithis1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"01")
		ithis21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"11") ''11

		''######################################## 2017-09-21 김진영 추가 ########################################
		Dim strSql, taxdate
		strSql = ""
		strSql = strSql & " SELECT TOP 1 isnull(taxdate,'') as taxdate FROM "
		strSql = strSql & " [db_sitemaster].[dbo].[tbl_taxdate_manage] with (nolock) "
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
		elseif (maytaxdate<premonth21day)  then		'maytaxdate가 저번달 12일보다 작다면 저번달 1일을 담는다
			GetPreFixSegumil = premonth1day
		else										'그 외는 maytaxdate를 담는다
			GetPreFixSegumil = maytaxdate
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



        if (Fid=278483) then
            GetPreFixSegumil = "2016-02-01"
        end if
        if (Fid=281205) then
            GetPreFixSegumil = "2016-02-01"
        end if

'
'        if (Fid=55328) then
'            GetPreFixSegumil = "2009-12-31"
'        end if
        'if (Fcompany_no="126-20-13576") then
        '    GetPreFixSegumil = "2012-02-01"
        'end if
		'if (Fdesignerid="pop_plan")  then
		'    GetPreFixSegumil = "2008-05-31"
		'end if
	end function

	'public function GetPreFixSegumil()
	'	dim thisdate, maytaxdate
	'	dim i0116, i0416, i0716, i1016
	'	dim i0101, i0401, i0701, i1001

	'	thisdate = getDbDate()
	'	maytaxdate = GetNormalTaxDate()

	'	i0116 = dateserial(Left(thisdate,4),"01","16")
	'	i0416 = dateserial(Left(thisdate,4),"04","16")
	'	i0716 = dateserial(Left(thisdate,4),"07","16")
	'	i1016 = dateserial(Left(thisdate,4),"10","16")

	'	i0101 = dateserial(Left(thisdate,4),"01","01")
	'	i0401 = dateserial(Left(thisdate,4),"04","01")
	'	i0701 = dateserial(Left(thisdate,4),"07","01")
	'	i1001 = dateserial(Left(thisdate,4),"10","01")

	'	if ((thisdate>=i1016) and (maytaxdate<i1001)) then
	'		GetPreFixSegumil = i1001
	'	elseif ((thisdate>=i0716) and (maytaxdate<i0701)) then
	'		GetPreFixSegumil = i0701
	'	elseif ((thisdate>=i0416) and (maytaxdate<i0401)) then
	'		GetPreFixSegumil = i0401
	'	elseif ((thisdate>=i0116) and (maytaxdate<i0101)) then
	'		GetPreFixSegumil = i0101
	'	else
	'		GetPreFixSegumil = maytaxdate
	'	end if
	'end function

    public function IsJungsanFixed()
        IsJungsanFixed = (Ffinishflag>=3)
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


	''//간이, 원천, 기타
	public function IsElecSimpleBillCase()
		IsElecSimpleBillCase = (Ftaxtype="03") and (Ffinishflag<3)
	end function

	public function GetSimpleTaxtypeName()
		if Ftaxtype="01" then
			GetSimpleTaxtypeName = "과세"
		elseif Ftaxtype="02" then
			GetSimpleTaxtypeName = "면세"
		elseif Ftaxtype="03" then
			GetSimpleTaxtypeName = "원천" '''"간이"
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

	public function GetTotalSellcash()
		GetTotalSellcash = Fub_totalsellcash + Fme_totalsellcash + Fwi_totalsellcash + Fet_totalsellcash + Fsh_totalsellcash + Fdlv_totalreducedprice
	end function

	public function GetTotalSuplycash()
		GetTotalSuplycash = CLNG(Fub_totalsuplycash + Fme_totalsuplycash + Fwi_totalsuplycash + Fet_totalsuplycash + Fsh_totalsuplycash + Fdlv_totalsuplycash) ''Fdlv_totalsuplycash 추가
	end function

	''원천징수대상자 정산금액
    public function GetTotalWithHoldingJungSanSum()
        dim ototalsum
        dim TreePercentTax
        ototalsum = GetTotalSuplycash
        TreePercentTax = Fix(Fix(ototalsum*0.03)/10)*10

        GetTotalWithHoldingJungSanSum = ototalsum

        ''3%세금이 1000원 이하이면 세금없음.=>미만(2018/06/08)
        ''if (TreePercentTax<=1000) then Exit function
        ''2018/06/05 수정
        if (TreePercentTax<1000) and (TreePercentTax>-1000) then Exit function

        GetTotalWithHoldingJungSanSum = ototalsum - TreePercentTax - Fix(Fix(TreePercentTax*0.1)/10)*10

	end function


	public function GetTotalTaxSuply()
		if Ftaxtype="01" then
			GetTotalTaxSuply = CLng(GetTotalSuplycash / 1.1)
		else
			GetTotalTaxSuply = GetTotalSuplycash
		end if
	end function

	public function GetTotalTaxVat()
		GetTotalTaxVat = GetTotalSuplycash - GetTotalTaxSuply
	end function

    ''역발행 중-----------------------
    public function IsInverseReqState()
        IsInverseReqState = Ffinishflag="1" and Not (isNULL(FISSU_SEQNO) or (FISSU_SEQNO=""))
    end function

    ''업체 역발행 요청중--------------
    public function IsInverseReqFinishAndEvalWaitState()
        IsInverseReqFinishAndEvalWaitState = (Ffinishflag="2") and Not isNULL(FISSU_SEQNO)
    end function

    public function IsInverseStateExists()
         IsInverseStateExists = Not (isNULL(FISSU_SEQNO) or (FISSU_SEQNO=""))
    end function

	public function GetStateName()
		if Ffinishflag="0" then
			GetStateName = "수정중"
		elseif Ffinishflag="1" then
		    IF isNULL(FISSU_SEQNO) or (FISSU_SEQNO="") then
		        GetStateName = "업체확인대기"
		    ELSE
			    GetStateName = "역발행등록중"
			ENd IF
		elseif Ffinishflag="2" then
		    IF isNULL(FISSU_SEQNO) or (FISSU_SEQNO="") then
		        GetStateName = "업체확인완료"
		    ELSE
			    GetStateName = "업체역발행대기"
			ENd IF
		elseif Ffinishflag="3" then
			GetStateName = "정산확정"
		elseif Ffinishflag="7" then
			GetStateName = "입금완료"
		else

		end if
	end function

	public function GetStateColor()
		if Ffinishflag="0" then
			GetStateColor = "#000000"
		elseif Ffinishflag="1" then
		    IF isNULL(FISSU_SEQNO) or (FISSU_SEQNO="") then
		        GetStateColor = "#448888"
		    ELSE
			    GetStateColor = "#884488"
		    ENd IF
		elseif Ffinishflag="2" then
		    IF isNULL(FISSU_SEQNO) or (FISSU_SEQNO="") then
			    GetStateColor = "#0000FF"
			ELSE
			    GetStateColor = "#AA44AA"
		    ENd IF
		elseif Ffinishflag="3" then
			GetStateColor = "#0000FF"
		elseif Ffinishflag="7" then
			GetStateColor = "#FF0000"
		else

		end if
	end function

    ''2016/02/01추가
    public function getBillItemName()
        '' 기존 : 온라인  makerid  판매대금
        getBillItemName = Ftitle &"-"& Fdesignerid
    end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CJungsanDetailItem
	public Fid
	public Fmasteridx
	public Fgubuncd
	public Fdetailidx
	public Fmastercode
	public Fbuyname
	public Freqname
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fitemno
	public Fsellcash
	public Fsuplycash

	public FOrgSellCash
	public FOrgSuplyCash

    public FOrgOptaddprice
    public FOrgOptaddbuyprice

	public FExecDate
	public Fcomment

    public Fvatinclude



    public Fmakerid
    public Fsitename
    public Freducedprice
    public Fcommission
    public Fiszerotax
    public Fpaymethod
    public FPgcommission
    public FCpnNotAppliedPrice

    public function getpaymethodName
        if isNULL(Fpaymethod) then Exit function

        if (Fpaymethod="100") then
            getpaymethodName = "신용카드"
        elseif (Fpaymethod="110") then
            getpaymethodName = "OK+신용카드"
        elseif (Fpaymethod="7") then
            getpaymethodName = "무통장"
        elseif (Fpaymethod="400") then
            getpaymethodName = "휴대폰"
        elseif (Fpaymethod="20") then
            getpaymethodName = "실시간이체"
        elseif (Fpaymethod="50") then
            getpaymethodName = "외부"
        else
            getpaymethodName = Fpaymethod
        end if
    end function

    public function getCouponDiscount
        getCouponDiscount = Fsellcash-Freducedprice
    end function

    public function getReducedprice
        getReducedprice = Freducedprice
    end function

    public function getCommission
        getCommission = Fcommission
    end function

    public function getPgCommission
        getPgCommission = FPgcommission
    end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CUpcheJungsanItem
	public FIdx
	public FOrderSerial
	public FItemId
	public FItemOption
	public FItemName
	public FItemOptionName
	public FItemNo
	public FBuyCash
	public FSellCash

	public FCurrState
	public FBeasongDate
	public FUpcheSongjangNo
	public FRegDate
	public FIpkumDate
	public FBuyName
	public FJumunDiv

	public FIpkumDiv
	public FMWDiv

	public Flec_date
    public Fvatinclude

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CMaeIpJungsanItem
	public FID
	public FMasterID
	public FDesignerID
	public FCode
	public FDivCode
	public FIpGoDate
	public FRegDate
	public FChargeId
	public Fchargename
	public FTotalsellcash
	public FTotalsuplycash
	public FTotalbuycash
	public FVatCode

	public FScheduleDate
	public FExecuteDate
	public Fsellcash
	public Fsuplycash
	public Fsuplycash2

	public FItemName
	public FItemGubun
	public FItemOptionName
	public FItemId
	public FItemOption
	public FItemNo

	public FComment
	public FMwDiv
	public FMakerid

	public Fvatinclude

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CMaeIpJungsanDetailItem
	public FID
	public FMasterCode
	public FDesignerID
	public FDivCode

	public FItemName
	public FItemOptionName
	public FItemId
	public FItemOption
	public FItemNo
	public Fsellcash
	public Fsuplycash

	public FSocID
	public FSocName
	public Fscheduledt
	public FExecuteDate
	public FRegDate
	public FScheduleDate
	public FCode
	public FTotalsuplycash

	public FBuyCash
	public Fmwgubun

	public Fcomment
	public Fitemgubun

    public Fvatinclude

	public function GetChulgoMwName()
		if Fmwgubun="C" then
			if FItemNo<0 then
				GetChulgoMwName = "[위탁재고->매입출고]"
			else
				GetChulgoMwName = "[출고반품->위탁재고]"
			end if
		elseif Fmwgubun="S" then
			if FItemNo<0 then
				GetChulgoMwName = "[매입재고->출고]"
			else
				GetChulgoMwName = "[출고반품->매입재고]"
			end if
		end if
	end function

	public function GetDivCodeColor()
		if Fdivcode="002" then
			GetDivCodeColor = "#000000"
		elseif Fdivcode="001" then
			GetDivCodeColor = "#DD5555"
		elseif Fdivcode="801" then
			GetDivCodeColor = "#DD5555"
		elseif Fdivcode="802" then
			GetDivCodeColor = "#5555DD"
		end if
	end function

	public function GetDivCodeName()
		if Fdivcode="002" then
			GetDivCodeName = "위탁"
		elseif Fdivcode="001" then
			GetDivCodeName = "매입"
		elseif Fdivcode="003" then
			GetDivCodeName = "판촉"
		elseif Fdivcode="004" then
			GetDivCodeName = "외부"
		elseif Fdivcode="005" then
			GetDivCodeName = "협찬"
		elseif Fdivcode="006" then
			GetDivCodeName = "B2B"
		elseif Fdivcode="007" then
			GetDivCodeName = "기타"
		elseif Fdivcode="801" then
			GetDivCodeName = "Off매입"
		elseif Fdivcode="802" then
			GetDivCodeName = "Off위탁"
		end if
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CJungsanSummaryByTaxDateItem
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

class CJungsanSumaryItem
	public Fyyyymm
	public Ftot

	''매입가
	public Fuptot
	public Fmetot
	public Fwitot
	public Fshtot
	public Fettot
    public Fdlvtot

	''판매가
	public Fupselltot
	public Fmeselltot
	public Fwiselltot
	public Fshselltot
	public Fetselltot
    public Fdlvselltot

	public Ffinishflag
	public Fjungsan_date

	public Ftotflag_notconfirmsum
	public Ftotflag_confirmsum
	public Ftotflag_ipkumsum

    public Ffixedthissum
    public Ffixednextsum

	public function GetStateName()
		if Ffinishflag="0" then
			GetStateName = "수정중"
		elseif Ffinishflag="1" then
			GetStateName = "업체확인대기"
		elseif Ffinishflag="2" then
			GetStateName = "업체확인완료"
		elseif Ffinishflag="3" then
			GetStateName = "정산확정"
		elseif Ffinishflag="7" then
			GetStateName = "입금완료"
		else

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
		else

		end if
	end function

	public function getTotSum()
		getTotSum = Fuptot + Fmetot + Fwitot + Fshtot + Fettot + Fdlvtot
	end function

	public function getTotSellcashSum()
		getTotSellcashSum = Fupselltot + Fmeselltot + Fwiselltot + Fshselltot + Fetselltot + Fdlvselltot
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CUpcheJungsan
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectid
	public FRectGubun
	public FRectDesigner
	public FRectMastercodes
	public FRectOrder
	public FRectYYYYMM
	public FRectPreYYYYMM
	public FRectStartDay
	public FRectEndDay

	public FWitakInsserted
	public FRectDesignerViewOnly

	public FCurrmastercode
	public FPremastercode
	public FRectState
	public FRectIpkumilNot

	public FRectNotIncludeWonChon
	public FRectOnlyIncludeWonChon
	public FRectOnlyIncludeNoTax
	public FRectOnlyIncludeSimpleTax
    public FRectOnlyIncludeKani

	public FRectDifferencekey

	public FRectSearchType
	public FRectSearchText
	public FRectCompanynoYN

	public FRectOnlyElecTax
	public FRectOnlyNotElecTax

	public FRectBankingupflag
	public FRectNotYYYYMM

    public FRectStartYYYYMM
    public FRectEndYYYYMM

    public FRectFixStateExiste
    public FRectfinishflag
    public FRectTaxRegDate
    public FRectJungsanDate
	public FRectJungsanGubun
    public FRectIpkumDate

    public FRectTaxType
    public FRectGroupID
    public FRectTaxDate
    public FRectTaxYYYYMM
    public FRectUnderMargin
    public FRectMinusGubnu
    public FRectbankingupFile
    public FRectTopN
    public FRectPurchaseType

    public FRectJGubun

    public FRectOnlyYYYYMM
    public FRectOnlyCommissionType
    public FRectItemVatYn

    public FREctSitename
    public FRecttargetGbn
    public FRectNotIncDivcode999
    public FRectjacctcd '' 계정과목코드

	public sub SearchJungsanList()
		if FRectGubun="upche" then
			SearchUpCheBeasongJungsanList
		elseif FRectGubun="maeip" then
			SearchMaeIpJungsanDetailList
		elseif FRectGubun="maeipchulgo" then
			SearchMaeipChulGoJungsanList
			'SearchChulGoJungsanList
		elseif FRectGubun="witak" then
			SearchWitakJungsanList
		elseif FRectGubun="witakchulgo" then
			Searchwitakchulgojungsanlist
		elseif FRectGubun="witaksell" then
			SearchWitakSellungsanList
		elseif FRectGubun="lecture" then
			SearchLectureJungsanList
		end if
	end sub


	public function JungsanDetailListSum()
		dim sqlStr,i
		sqlStr = "select T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.itemno, T.sellcash, T.suplycash,"
		sqlStr = sqlStr + " i.sellcash as orgsellcash, i.buycash as orgsuplycash"
		sqlStr = sqlStr + " ,T.reducedprice, T.commission"
		sqlStr = sqlStr + " from ("
			sqlStr = sqlStr + "select d.itemid,d.itemoption,d.itemname,d.itemoptionname,sum(d.itemno) as itemno,d.sellcash,d.suplycash"
			sqlStr = sqlStr + " , d.reducedprice, d.commission"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d "
			sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectid)

			if FRectgubun<>"" then
				sqlStr = sqlStr + " and gubuncd='" + FRectgubun + "'"
			end if

            if (FRectItemVatYn<>"") then
                sqlStr = sqlStr + " and d.vatyn='" + FRectItemVatYn + "'"
            end if

			sqlStr = sqlStr + " group by d.itemid,d.itemoption,d.itemname,d.itemoptionname,d.sellcash,d.suplycash"
			sqlStr = sqlStr + " , d.reducedprice, d.commission"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on T.itemid=i.itemid"
'rw sqlStr
	    rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanDetailItem

				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fitemno        = rsget("itemno")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fsuplycash  	= rsget("suplycash")

				FItemList(i).FOrgsellcash   = rsget("orgsellcash")
				FItemList(i).FOrgsuplycash  = rsget("orgsuplycash")

				FItemList(i).Freducedprice  = rsget("reducedprice")
				FItemList(i).Fcommission    = rsget("commission")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

    public function JungsanDetailListLectureSum()
        dim sqlStr,i
		sqlStr = "select T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.itemno, T.sellcash, T.suplycash,"
		''sqlStr = sqlStr + " isNULL(i.lec_cost,di.sellcash) as orgsellcash, isNULL(i.buying_cost,di.buycash) as orgsuplycash"
		sqlStr = sqlStr + " NULL as orgsellcash, NULL as orgsuplycash"
		sqlStr = sqlStr + " ,T.reducedprice,T.commission,T.pgcommission"
		sqlStr = sqlStr + " from ("
			sqlStr = sqlStr + "select d.jitemgubun,d.itemid,d.itemoption,d.itemname,d.itemoptionname,sum(d.itemno) as itemno,d.sellcash,"
			sqlStr = sqlStr + " d.suplycash, d.reducedprice, d.commission,isNULL(d.pgcommission,0) as pgcommission"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d "
			sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectid)

			if FRectgubun<>"" then
				sqlStr = sqlStr + " and gubuncd='" + FRectgubun + "'"
			end if
			sqlStr = sqlStr + " group by d.jitemgubun,d.itemid,d.itemoption,d.itemname,d.itemoptionname,d.sellcash,d.suplycash"
			sqlStr = sqlStr + " , d.reducedprice, d.commission, isNULL(d.pgcommission,0)"
		sqlStr = sqlStr + " ) as T"
		''sqlStr = sqlStr + " left join [ACADEMYDB].[db_academy].[dbo].tbl_lec_item i on T.jitemgubun='97' and T.itemid=i.idx"
		''sqlStr = sqlStr + " left join [ACADEMYDB].[db_academy].[dbo].tbl_diy_item di on T.jitemgubun='98' and T.itemid=di.itemid"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanDetailItem

				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fitemno        = rsget("itemno")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fsuplycash  	= rsget("suplycash")

				FItemList(i).FOrgsellcash      = rsget("orgsellcash")
				FItemList(i).FOrgsuplycash  	= rsget("orgsuplycash")

				FItemList(i).Freducedprice      = rsget("reducedprice")
				FItemList(i).Fcommission        = rsget("commission")

				FItemList(i).FPgcommission        = rsget("pgcommission")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end function

	public function JungsanDetailListWitakSum()
		dim sqlStr,i
		sqlStr = "select T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.itemno, T.sellcash, T.suplycash,"
		sqlStr = sqlStr + " IsNULL(i.sellcash,0) as orgsellcash, IsNULL(i.buycash,0) as orgsuplycash"
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice,IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr + " ,T.reducedprice,T.commission, T.pgcommission"
		sqlStr = sqlStr + " from ("
			sqlStr = sqlStr + "select d.itemid,d.itemoption,d.itemname,d.itemoptionname,sum(d.itemno) as itemno,d.sellcash,"
			sqlStr = sqlStr + " d.suplycash, d.reducedprice, d.commission, isNULL(d.pgcommission,0) as pgcommission"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d "
			sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectid)

			if FRectgubun<>"" then
			    'if (FRectgubun="DL") or (FRectgubun="DT") then
			    '    sqlStr = sqlStr + " and gubuncd in ('DL','DT','DP')"
			    'else
				'    sqlStr = sqlStr + " and gubuncd='" + FRectgubun + "'"
			    'end if
			    sqlStr = sqlStr + " and gubuncd='" + FRectgubun + "'"
			end if

			if (FREctSitename<>"") then
			    if (FREctSitename="N10x10") then
			        sqlStr = sqlStr + " and sitename<>'10x10'"
			    else
    			    sqlStr = sqlStr + " and sitename='" + FREctSitename + "'"
    			end if
			end if
			sqlStr = sqlStr + " group by d.itemid,d.itemoption,d.itemname,d.itemoptionname,d.sellcash,d.suplycash"
			sqlStr = sqlStr + " , d.reducedprice, d.commission, isNULL(d.pgcommission,0)"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on T.itemid=i.itemid"
        sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on T.itemid=o.itemid and T.itemoption=o.itemoption"

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanDetailItem

				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fitemno        = rsget("itemno")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fsuplycash  	= rsget("suplycash")

				FItemList(i).FOrgsellcash      = rsget("orgsellcash")
				FItemList(i).FOrgsuplycash  	= rsget("orgsuplycash")

				FItemList(i).FOrgOptaddprice    = rsget("optaddprice")
				FItemList(i).FOrgOptaddbuyprice = rsget("optaddbuyprice")
				FItemList(i).Freducedprice      = rsget("reducedprice")
				FItemList(i).Fcommission        = rsget("commission")

				FItemList(i).FPgcommission      = rsget("pgcommission")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

	'//admin/upchejungsan/jungsandetailsumONAdm.asp
	public function JungsanDetailList()
		dim sqlStr,i


		if (FRectgubun="upche") or (FRectgubun="witaksell") or (FRectgubun="DL") or (FRectgubun="DT")  then
			sqlStr = "select top 10000 j.id,j.masteridx,j.gubuncd,j.detailidx,j.mastercode,j.buyname,j.reqname,"
			sqlStr = sqlStr + "j.itemid,j.itemoption,j.itemname,j.itemoptionname,j.itemno,j.sellcash,"
			sqlStr = sqlStr + "j.suplycash"
			sqlStr = sqlStr + ", convert(varchar(10),j.beasongdate,21) as execdate"
			sqlStr = sqlStr + ", j.sitename, j.reducedprice, j.commission, j.iszerotax, j.paymethod"
			sqlStr = sqlStr + ", isNULL(j.pgcommission,0) as pgcommission"
			sqlStr = sqlStr + ", isNULL(j.CpnNotAppliedPrice,0) as CpnNotAppliedPrice"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail j"

			''sqlStr = sqlStr + " left join "
			''sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d on d.idx=j.detailidx"
		elseif FRectgubun="maeip" then
			sqlStr = "select distinct top 10000 j.id,j.masteridx,j.gubuncd,j.detailidx,j.mastercode,j.buyname,j.reqname,"
			sqlStr = sqlStr + "j.itemid,j.itemoption,j.itemname,j.itemoptionname,j.itemno,j.sellcash,"
			sqlStr = sqlStr + "j.suplycash, convert(varchar(10),d.executedt,21) as execdate "
			sqlStr = sqlStr + ", j.sitename, j.reducedprice, j.commission, j.iszerotax, j.paymethod"
			sqlStr = sqlStr + ", isNULL(j.pgcommission,0) as pgcommission"
			sqlStr = sqlStr + ", isNULL(j.CpnNotAppliedPrice,0) as CpnNotAppliedPrice"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail j"

			sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master d on d.code=j.mastercode"
		else
			sqlStr = "select distinct top 10000 j.id,j.masteridx,j.gubuncd,j.detailidx,j.mastercode,j.buyname,j.reqname,"
			sqlStr = sqlStr + "j.itemid,j.itemoption,j.itemname,j.itemoptionname,j.itemno,j.sellcash,"
			sqlStr = sqlStr + "j.suplycash, convert(varchar(10),d.executedt,21) as execdate "
			sqlStr = sqlStr + ", j.sitename, j.reducedprice, j.commission, j.iszerotax, j.paymethod"
			sqlStr = sqlStr + ", isNULL(j.pgcommission,0) as pgcommission"
			sqlStr = sqlStr + ", isNULL(j.CpnNotAppliedPrice,0) as CpnNotAppliedPrice"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail j"

			sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master d on d.code=j.mastercode"
		end if
		sqlStr = sqlStr + " where j.masteridx=" + CStr(FRectid)

		if FRectgubun<>"" then
	        'if (FRectgubun="DL") or (FRectgubun="DT") then
		    '    sqlStr = sqlStr + " and j.gubuncd in ('DL','DT','DP')"
		    'else
			'    sqlStr = sqlStr + " and j.gubuncd='" + FRectgubun + "'"
		    'end if
			sqlStr = sqlStr + " and j.gubuncd='" + FRectgubun + "'"
		end if
        if (FRectItemVatYn<>"") then
            sqlStr = sqlStr + " and j.vatyn='" + FRectItemVatYn + "'"
        end if

        if (FREctSitename<>"") then
            if (FREctSitename="N10x10") then
		        sqlStr = sqlStr + " and sitename<>'10x10'"
		    else
			    sqlStr = sqlStr + " and sitename='" + FREctSitename + "'"
			end if
        end if
		if FRectOrder="itemid" then
			sqlStr = sqlStr + " order by j.itemid, j.itemoption, j.mastercode desc"
		else
			sqlStr = sqlStr + " order by j.mastercode"
		end if

		''response.write sqlStr & "<br>"
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
			do until rsget.eof
				set FItemList(i) = new CJungsanDetailItem

				FItemList(i).Fid            = rsget("id")
				FItemList(i).Fmasteridx     = rsget("masteridx")
				FItemList(i).Fgubuncd       = rsget("gubuncd")
				FItemList(i).Fdetailidx     = rsget("detailidx")
				FItemList(i).Fmastercode    = rsget("mastercode")
				FItemList(i).Fbuyname       = db2html(rsget("buyname"))
				FItemList(i).Freqname       = db2html(rsget("reqname"))
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fitemno        = rsget("itemno")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fsuplycash  	= rsget("suplycash")

				FItemList(i).FExecDate      = rsget("execdate")

				''2014
				'FItemList(i).Fmakerid       = rsget("makerid")
                FItemList(i).Fsitename      = rsget("sitename")
                FItemList(i).Freducedprice  = rsget("reducedprice")
                FItemList(i).Fcommission    = rsget("commission")
                FItemList(i).Fiszerotax     = rsget("iszerotax")
                FItemList(i).Fpaymethod     = rsget("paymethod")

                FItemList(i).Fpgcommission  = rsget("pgcommission")
                ''2018/07/02
                FItemList(i).FCpnNotAppliedPrice = rsget("CpnNotAppliedPrice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

	public function JungsanDetailListByYYYYMM()
		dim sqlStr,i

		sqlStr = "select d.id,d.masteridx,d.gubuncd,d.detailidx,d.mastercode,d.buyname,d.reqname,"
		sqlStr = sqlStr + "d.itemid,d.itemoption,d.itemname,d.itemoptionname,d.itemno,d.sellcash,"
		sqlStr = sqlStr + "d.suplycash, i.vatinclude "
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m,"
		sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i on d.itemid=i.itemid"
		sqlStr = sqlStr + " where m.designerid='" + CStr(FRectdesigner) + "'"
		sqlStr = sqlStr + " and m.yyyymm='" + CStr(FRectYYYYMM) + "'"
		sqlStr = sqlStr + " and m.id=d.masteridx"
		if FRectdifferencekey<>"" then
			sqlStr = sqlStr + " and m.differencekey=" + CStr(FRectdifferencekey)
		end if
		sqlStr = sqlStr + " and d.gubuncd='" + FRectgubun + "'"
		sqlStr = sqlStr + " order by d.mastercode desc, d.detailidx"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanDetailItem

				FItemList(i).Fid            = rsget("id")
				FItemList(i).Fmasteridx     = rsget("masteridx")
				FItemList(i).Fgubuncd       = rsget("gubuncd")
				FItemList(i).Fdetailidx     = rsget("detailidx")
				FItemList(i).Fmastercode    = rsget("mastercode")
				FItemList(i).Fbuyname       = rsget("buyname")
				FItemList(i).Freqname       = rsget("reqname")
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = rsget("itemname")
				FItemList(i).Fitemoptionname= rsget("itemoptionname")
				FItemList(i).Fitemno        = rsget("itemno")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fsuplycash  	= rsget("suplycash")

                FItemList(i).Fvatinclude = rsget("vatinclude")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

    public function JungsanSummaryBySegumDate()
        dim sqlStr,i
        ''taxregdate IsNULL = 원천징수 등.

        sqlStr = " select m.taxregdate," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and (g.jungsan_date='수시') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as jungsansum_susi," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and (g.jungsan_date='말일') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as jungsansum_31date," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and (g.jungsan_date='15일') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as jungsansum_15date," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and ((g.jungsan_date is NULL) or (g.jungsan_date not in('수시','말일','15일'))) then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as jungsansum_etcdate," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm<>convert(varchar(7),m.taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as ewol_jungsansum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as fixedsum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='7') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as ipkumsum," + VbCrlf
        sqlStr = sqlStr + " sum(m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) as tot_jungsanprice" + VbCrlf
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m" + VbCrlf
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group g " + VbCrlf
        sqlStr = sqlStr + "     on m.groupid=g.groupid" + VbCrlf
        sqlStr = sqlStr + " where m.finishflag >=3" + VbCrlf

        if (FRectStartDay<>"") then
            sqlStr = sqlStr + " and m.taxregdate>='" + FRectStartDay + "'" + VbCrlf
        end if

        if (FRectEndDay<>"") then
            sqlStr = sqlStr + " and m.taxregdate<'" + FRectEndDay + "'" + VbCrlf
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

			set FItemList(i) = new CJungsanSummaryByTaxDateItem


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

    end function

	public function JungsanSummary0()
		dim sqlStr,i
		sqlStr = "select m.yyyymm,"
		sqlStr = sqlStr + " IsNull(p.jungsan_date,'') as jungsan_date, "
		sqlStr = sqlStr + " Sum(m.ub_totalsuplycash) as uptot, "
		sqlStr = sqlStr + " Sum(m.me_totalsuplycash) as metot, "
		sqlStr = sqlStr + " Sum(m.wi_totalsuplycash) as witot, "
		sqlStr = sqlStr + " Sum(m.sh_totalsuplycash) as shtot, "
		sqlStr = sqlStr + " Sum(m.et_totalsuplycash) as ettot, "
		sqlStr = sqlStr + " Sum(m.dlv_totalsuplycash) as dlvtot, "

        sqlStr = sqlStr + " Sum(m.ub_totalsellcash) as upselltot, "
		sqlStr = sqlStr + " Sum(m.me_totalsellcash) as meselltot, "
		sqlStr = sqlStr + " Sum(m.wi_totalsellcash) as wiselltot, "
		sqlStr = sqlStr + " Sum(m.sh_totalsellcash) as shselltot, "
		sqlStr = sqlStr + " Sum(m.et_totalsellcash) as etselltot, "
		sqlStr = sqlStr + " Sum(m.dlv_totalsellcash) as dlvselltot, "

		sqlStr = sqlStr + " Sum(CASE "
        sqlStr = sqlStr + "     WHEN (m.finishflag='7') THEN (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash+ m.dlv_totalsuplycash)"
        sqlStr = sqlStr + "     ELSE 0"
      	sqlStr = sqlStr + "     END ) as totflag_ipkumsum,"
      	sqlStr = sqlStr + " Sum(CASE "
        sqlStr = sqlStr + "     WHEN (m.finishflag='3') THEN (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash+ m.dlv_totalsuplycash)"
        sqlStr = sqlStr + "     ELSE 0"
      	sqlStr = sqlStr + "     END ) as totflag_confirmsum,"
      	sqlStr = sqlStr + " Sum(CASE "
        sqlStr = sqlStr + "     WHEN (m.finishflag <'3') THEN (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash+ m.dlv_totalsuplycash)"
        sqlStr = sqlStr + "     ELSE 0"
      	sqlStr = sqlStr + "     END ) as totflag_notconfirmsum,"

      	''정산일 기준으로 입금예정금액 산출.
        ''sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (m.yyyymm=convert(varchar(7),taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as fixedthissum," + VbCrlf
        ''sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (m.yyyymm<>convert(varchar(7),taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as fixednextsum " + VbCrlf

        ''금월 기준으로 입금예정금액 산출. taxregdate IsNULL = 원천징수 등.
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (convert(varchar(7),getdate(),21)>convert(varchar(7),m.taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as fixedthissum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (convert(varchar(7),getdate(),21)<=convert(varchar(7),m.taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash + m.dlv_totalsuplycash) else 0 end) as fixednextsum" + VbCrlf

		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlStr = sqlStr + " where 1=1"

		if (FRectStartYYYYMM<>"") then
            sqlStr = sqlStr + " and m.yyyymm>='" + FRectStartYYYYMM + "'" + VbCrlf
        end if

        if (FRectEndYYYYMM<>"") then
            sqlStr = sqlStr + " and m.yyyymm<='" + FRectEndYYYYMM + "'" + VbCrlf
        end if

		sqlStr = sqlStr + " group by m.yyyymm, p.jungsan_date"
		if FRectFixStateExiste<>"" then
		    ''미처리 내역이 있는것..
            sqlStr = sqlStr + " having sum(case when (m.finishflag<=3) then (m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash + m.dlv_totalsellcash) else 0 end)<>0"
		end if
		sqlStr = sqlStr + " order by m.yyyymm desc, p.jungsan_date"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanSumaryItem

				FItemList(i).Fyyyymm           = rsget("yyyymm")
				FItemList(i).Fuptot       	   = rsget("uptot")
				FItemList(i).Fmetot            = rsget("metot")
				FItemList(i).Fwitot            = rsget("witot")
				FItemList(i).Fshtot            = rsget("shtot")
				FItemList(i).Fettot            = rsget("ettot")
				FItemList(i).Fdlvtot           = rsget("dlvtot")

				FItemList(i).Fupselltot       	   = rsget("upselltot")
				FItemList(i).Fmeselltot            = rsget("meselltot")
				FItemList(i).Fwiselltot            = rsget("wiselltot")
				FItemList(i).Fshselltot            = rsget("shselltot")
				FItemList(i).Fetselltot            = rsget("etselltot")
				FItemList(i).Fdlvselltot           = rsget("dlvselltot")

				FItemList(i).Ftotflag_notconfirmsum  = rsget("totflag_notconfirmsum")
				FItemList(i).Ftotflag_confirmsum     = rsget("totflag_confirmsum")
				FItemList(i).Ftotflag_ipkumsum       = rsget("totflag_ipkumsum")

                FItemList(i).Ffixedthissum      = rsget("fixedthissum")
                FItemList(i).Ffixednextsum      = rsget("fixednextsum")

				FItemList(i).Fjungsan_date     = rsget("jungsan_date")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public function JungsanSummary()
		dim sqlStr,i
		sqlStr = "select m.yyyymm, sum(IsNull(m.ub_totalsuplycash,0)) as uptot,"
		sqlStr = sqlStr + " sum(IsNull(m.me_totalsuplycash,0)) as metot,"
		sqlStr = sqlStr + " sum(IsNull(m.wi_totalsuplycash,0)) as witot, "
		sqlStr = sqlStr + " sum(IsNull(m.sh_totalsuplycash,0)) as shtot, "
		sqlStr = sqlStr + " sum(IsNull(m.et_totalsuplycash,0)) as ettot, m.finishflag,"
		sqlStr = sqlStr + " sum(IsNull(m.dlv_totalsuplycash,0)) as dlvtot,"
		sqlStr = sqlStr + " IsNull(p.jungsan_date,'') as jungsan_date"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlStr = sqlStr + " where m.cancelyn='N'"
		sqlStr = sqlStr + " and m.finishflag<7"
		sqlStr = sqlStr + " group by m.yyyymm, m.finishflag, p.jungsan_date"
		sqlStr = sqlStr + " order by m.yyyymm desc, p.jungsan_date, m.finishflag"
		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanSumaryItem

				FItemList(i).Fyyyymm           = rsget("yyyymm")
				FItemList(i).Fuptot       	   = rsget("uptot")
				FItemList(i).Fmetot            = rsget("metot")
				FItemList(i).Fwitot            = rsget("witot")
				FItemList(i).Fshtot            = rsget("shtot")
				FItemList(i).Fettot            = rsget("ettot")
				FItemList(i).Fdlvtot           = rsget("dlvtot")

				FItemList(i).Ffinishflag       = rsget("finishflag")
				FItemList(i).Fjungsan_date     = rsget("jungsan_date")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

    public function JungsanFixedList()
		dim sqlStr,i, sqlAdd, bufSqlStr

		sqlStr = "select "
		if FRectTopN<>"" then
		    sqlStr = sqlStr + " top "&Cstr(FRectTopN)
		end if
		sqlStr = sqlStr + " m.id, m.designerid, m.groupid, m.yyyymm, m.title, m.ub_cnt,"
		sqlStr = sqlStr + " m.ub_totalsellcash,"
		sqlStr = sqlStr + " m.ub_totalsuplycash,"
		sqlStr = sqlStr + " m.ub_comment,"
		sqlStr = sqlStr + " m.me_cnt, m.me_totalsellcash,"
		sqlStr = sqlStr + " m.me_totalsuplycash, m.me_comment,"
		sqlStr = sqlStr + " m.wi_cnt, m.wi_totalsellcash,"
		sqlStr = sqlStr + " m.wi_totalsuplycash, m.wi_comment,"
		sqlStr = sqlStr + " m.et_cnt, m.et_totalsellcash,"
		sqlStr = sqlStr + " m.et_totalsuplycash, m.et_comment,"
		sqlStr = sqlStr + " m.sh_cnt, m.sh_totalsellcash,"
		sqlStr = sqlStr + " m.sh_totalsuplycash, m.sh_comment,"
		sqlStr = sqlStr + " m.regdate,m.cancelyn,m.finishflag,convert(varchar(10),m.ipkumdate,20) as ipkumdate,"
		sqlStr = sqlStr + " convert(varchar(10),m.taxregdate,20) as taxregdate, m.bigo"
		sqlStr = sqlStr + " , p.jungsan_email,isnull(pf.jungsan_date,p.jungsan_date) as jungsan_date,isnull(pf.jungsan_bank,p.jungsan_bank) as jungsan_bank"
		sqlStr = sqlStr + " , isnull(pf.jungsan_acctno,p.jungsan_acctno) as jungsan_acctno,m.ipkum_bank,m.ipkum_acctno"
		sqlStr = sqlStr + " , isnull(pf.jungsan_acctname,p.jungsan_acctname) as jungsan_acctname,p.company_name"
		sqlStr = sqlStr + " , p.jungsan_gubun,p.company_no,p.ceoname,p.company_address,p.company_address2"
		sqlStr = sqlStr + " , m.taxinputdate, m.differencekey, m.taxtype, m.taxlinkidx, m.neotaxno, m.bankingupflag, m.billsitecode"
        sqlStr = sqlStr + " ,m.jgubun,m.itemvatYn,m.wi_totalreducedprice,m.ub_totalreducedprice,m.et_totalreducedprice,m.dlv_totalreducedprice,m.dlv_totalsuplycash,m.totalcommission"

		IF (FRectbankingupFile<>"") then
		    sqlStr = sqlStr + " ,ip.ipFileNo, ip.targetGbn"
		ELSE
		    sqlStr = sqlStr + " ,NULL as ipFileNo, NULL as targetGbn"
		End IF

		sqlStr = sqlStr + " ,HH.groupid as holdGroupid, HH.holdcause"

		sqlAdd = ""
		sqlAdd = sqlAdd + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlAdd = sqlAdd + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlAdd = sqlAdd + " left join db_partner.dbo.tbl_partner_addJungsanInfo pf with (readuncommitted)"
		sqlAdd = sqlAdd + " 	on m.designerid = pf.partnerid"

		if FRectbankingupFile<>"" then
		    sqlAdd = sqlAdd + "     left join db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail IP"
		    sqlAdd = sqlAdd + "     on Ip.targetGbn='ON' and Ip.targetIdx=m.id"
	    end if

		sqlAdd = sqlAdd + "     left join (select g.groupid, H.holdcause from db_partner.dbo.tbl_partner_group g "
        sqlAdd = sqlAdd + "                 Join db_jungsan.dbo.tbl_jungsan_hold H"
        sqlAdd = sqlAdd + "                 on replace(g.company_no,'-','')=H.holdsocno) as HH"
        sqlAdd = sqlAdd + "     on m.groupid=HH.groupid"

		if FRectfinishflag="ALL" then
		    sqlAdd = sqlAdd + " where m.finishflag>=3"
		elseif FRectfinishflag="NFixInclude" then
		    sqlAdd = sqlAdd + " where 1=1"
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
        if FRectGubun="ZZ" then
            sqlAdd = sqlAdd + " and m.taxregdate is NULL"
        elseif FRectGubun="SS" then
            sqlAdd = sqlAdd + " and (p.jungsan_date='수시')"           ''수시
        elseif FRectGubun="AA" then
            sqlAdd = sqlAdd + " and ((p.jungsan_date='수시')"           ''수시추가
            sqlAdd = sqlAdd + " or ((IsNULL(p.jungsan_date,'')='' or p.jungsan_date<>'말일')"
           '' sqlAdd = sqlAdd + "     and m.yyyymm=convert(varchar(7),m.taxregdate,21)"  ''2018/07/13 제거
            sqlAdd = sqlAdd + " ))"
        elseif FRectGubun="BB" then
            sqlAdd = sqlAdd + " and ((p.jungsan_date='수시')"           ''수시추가
            sqlAdd = sqlAdd + " or ((p.jungsan_date='말일')"
            sqlAdd = sqlAdd + "     and m.yyyymm=convert(varchar(7),m.taxregdate,21)))"

            'sqlAdd = sqlAdd + " and p.jungsan_date in ('말일','수시')"
            'sqlAdd = sqlAdd + " and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
            ''rw sqlAdd
        elseif FRectGubun="CC" then
            sqlAdd = sqlAdd + " and m.yyyymm<convert(varchar(7),m.taxregdate,21)"
            sqlAdd = sqlAdd + " and convert(varchar(7),getdate(),21)>convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubun="DD" then
            sqlAdd = sqlAdd + " and convert(varchar(7),getdate(),21)<=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubun="EE" then
            sqlAdd = sqlAdd + " and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubun="FF" then
            sqlAdd = sqlAdd + " and m.yyyymm<>convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubun="NN" then
            sqlAdd = sqlAdd + " and m.yyyymm='"&LEFT(now(),7)&"'"
        end if

        if FRectJungsanDate="NULL" then
            sqlAdd = sqlAdd + " and IsNULL(p.jungsan_date,'')=''"
        elseif FRectJungsanDate<>"" then
            sqlAdd = sqlAdd + " and p.jungsan_date='" + FRectJungsanDate + "'"
        end if

        if FRectNotIncludeWonChon<>"" then
			''''sqlAdd = sqlAdd + " and p.jungsan_gubun<>'원천징수'"
			sqlAdd = sqlAdd + " and m.taxtype<>'03'"
			''sqlAdd = sqlAdd + " and p.jungsan_gubun<>'간이과세'"  ''주석처리 2018/02/13
		end if

        if (FRectOnlyIncludeKani="on") then
            sqlAdd = sqlAdd + " and p.jungsan_gubun='간이과세'"
        end if

		if FRectOnlyIncludeWonChon<>"" then
			'''sqlAdd = sqlAdd + " and p.jungsan_gubun in ('원천징수','간이과세')"
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

		if FRectDesigner<>"" then
			sqlAdd = sqlAdd + " and m.designerid='" + FRectDesigner + "'"
		end if

		if FRectGroupid<>"" then
			sqlAdd = sqlAdd + " and p.groupid='" + FRectGroupid + "'"
		end if

		if (FRectTaxDate<>"") then
		    sqlAdd = sqlAdd + " and m.taxregdate='" + FRectTaxDate + "'"
		end if

		if (FRectTaxYYYYMM<>"") then
		    sqlAdd = sqlAdd + " and convert(varchar(7),m.taxregdate,21)='" + FRectTaxYYYYMM + "'"
		end if

		bufSqlStr = sqlAdd

		if FRectMinusGubnu="MI" then

            sqlAdd = sqlAdd + " and m.groupid in ("
            sqlAdd = sqlAdd + "     select m.groupid " + bufSqlStr
            sqlAdd = sqlAdd + "     and m.ub_totalsuplycash+m.me_totalsuplycash+m.wi_totalsuplycash+m.et_totalsuplycash+m.sh_totalsuplycash+m.dlv_totalsuplycash<1"
            sqlAdd = sqlAdd + " )"
            sqlAdd = sqlAdd + " order by (CASE WHEN IsNULL(m.billsitecode,'')='' THEN 'ZZZ' ELSE m.billsitecode END) desc, m.groupid, m.taxinputdate"
        ELSEif FRectMinusGubnu="MJ" then ''마이너스 제외
             bufSqlStr = replace(replace(replace(bufSqlStr,"and m.jgubun='CC'"," "),"and m.jgubun='MM'"," "),"and m.jgubun='CE'"," ")
             bufSqlStr = replace(bufSqlStr,"and m.yyyymm='" + FRectYYYYMM + "'"," ")

            sqlAdd = sqlAdd + "     and m.ub_totalsuplycash+m.me_totalsuplycash+m.wi_totalsuplycash+m.et_totalsuplycash+m.sh_totalsuplycash+m.dlv_totalsuplycash>0"
            ''추가..2017/02/22
            sqlAdd = sqlAdd + " and m.groupid not in ("
            sqlAdd = sqlAdd + "     select m.groupid " + bufSqlStr
            sqlAdd = sqlAdd + "     and m.ub_totalsuplycash+m.me_totalsuplycash+m.wi_totalsuplycash+m.et_totalsuplycash+m.sh_totalsuplycash+m.dlv_totalsuplycash<1"
            sqlAdd = sqlAdd + " )"
            sqlAdd = sqlAdd + " order by (CASE WHEN IsNULL(m.billsitecode,'')='' THEN 'ZZZ' ELSE m.billsitecode END) desc, m.taxinputdate"
        ELSEif (FRectMinusGubnu="CX") or (FRectMinusGubnu="CX1") then
            ''' 온라인 자체 상계처리 또는 오프라인 온라인 금액합>마이너스건
            sqlAdd = sqlAdd + " and m.groupid in ("
            sqlAdd = sqlAdd + "     select groupid from ("
            sqlAdd = sqlAdd + "         select m.groupid from  [db_jungsan].[dbo].tbl_designer_jungsan_master m"
            sqlAdd = sqlAdd + "         where m.ub_totalsuplycash+m.me_totalsuplycash+m.wi_totalsuplycash+m.et_totalsuplycash+m.sh_totalsuplycash+m.dlv_totalsuplycash<1"
            sqlAdd = sqlAdd + "         and m.finishflag='3'"
            if (FRectJGubun<>"") then
                sqlAdd = sqlAdd + " and m.jgubun='"&FRectJGubun&"'"
            end if
            sqlAdd = sqlAdd + "         Union ALL "
            sqlAdd = sqlAdd + "         select m2.groupid"
            sqlAdd = sqlAdd + "         from  [db_jungsan].[dbo].tbl_off_jungsan_master m2"
            sqlAdd = sqlAdd + "         where m2.tot_jungsanprice<1 and m2.finishflag='3'"
            if (FRectJGubun<>"") then
                sqlAdd = sqlAdd + " and m2.jgubun='"&FRectJGubun&"'"
            end if
            sqlAdd = sqlAdd + "     ) T0"
            sqlAdd = sqlAdd + "     group by T0.groupid"
            sqlAdd = sqlAdd + " )"
            sqlAdd = sqlAdd + " and m.groupid in ("
            sqlAdd = sqlAdd + "     select groupid from ("
            sqlAdd = sqlAdd + "         select m.groupid, m.ub_totalsuplycash+m.me_totalsuplycash+m.wi_totalsuplycash+m.et_totalsuplycash+m.sh_totalsuplycash+m.dlv_totalsuplycash as jSum1, 0 as jSum2"
            sqlAdd = sqlAdd + "         from  [db_jungsan].[dbo].tbl_designer_jungsan_master m"
            sqlAdd = sqlAdd + "         where m.finishflag='3'"
            if (FRectJGubun<>"") then
                sqlAdd = sqlAdd + " and m.jgubun='"&FRectJGubun&"'"
            end if
            sqlAdd = sqlAdd + "         Union ALL "
            sqlAdd = sqlAdd + "         select m2.groupid,0 as jSum1, m2.tot_jungsanprice as jSum2 "
            sqlAdd = sqlAdd + "         from  [db_jungsan].[dbo].tbl_off_jungsan_master m2"
            sqlAdd = sqlAdd + "         where m2.finishflag='3'"
            if (FRectJGubun<>"") then
                sqlAdd = sqlAdd + " and m2.jgubun='"&FRectJGubun&"'"
            end if
            sqlAdd = sqlAdd + "     ) T"
            sqlAdd = sqlAdd + "     group by T.groupid"
            if (FRectMinusGubnu="CX") then
                sqlAdd = sqlAdd + "     having sum(T.jSum1+T.jSum2)>=0 and sum(T.jSum1)>=0 and (sum(T.jSum2)<1 or sum(CASE WHEN T.jSum2<0 then 1 ELSE 0 END)=0)"
            elseif (FRectMinusGubnu="CX1") then
                sqlAdd = sqlAdd + "     having sum(T.jSum1+T.jSum2)>=0 and sum(T.jSum1)<1 and sum(T.jSum2)>0"
            end if
            sqlAdd = sqlAdd + "     "
           sqlAdd = sqlAdd + " )"
            sqlAdd = sqlAdd + " order by (CASE WHEN IsNULL(m.billsitecode,'')='' THEN 'ZZZ' ELSE m.billsitecode END) desc, m.groupid, m.taxinputdate"
        ELSE
            sqlAdd = sqlAdd + " order by (CASE WHEN IsNULL(m.billsitecode,'')='' THEN 'ZZZ' ELSE m.billsitecode END) desc, m.taxinputdate"
        end if

        sqlStr = sqlStr + sqlAdd
		'rw sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanMasterItem

				FItemList(i).Fid               = rsget("id")
				FItemList(i).Fdesignerid       = rsget("designerid")
				FItemList(i).Fgroupid		   = rsget("groupid")
				FItemList(i).Fyyyymm           = rsget("yyyymm")
				FItemList(i).Ftitle            = rsget("title")
				FItemList(i).Fub_cnt           = rsget("ub_cnt")
				FItemList(i).Fub_totalsellcash = rsget("ub_totalsellcash")
				FItemList(i).Fub_totalsuplycash= rsget("ub_totalsuplycash")
				FItemList(i).Fub_comment       = db2html(rsget("ub_comment"))
				FItemList(i).Fme_cnt           = rsget("me_cnt")
				FItemList(i).Fme_totalsellcash = rsget("me_totalsellcash")
				FItemList(i).Fme_totalsuplycash= rsget("me_totalsuplycash")
				FItemList(i).Fme_comment       = db2html(rsget("me_comment"))
				FItemList(i).Fwi_cnt           = rsget("wi_cnt")
				FItemList(i).Fwi_totalsellcash = rsget("wi_totalsellcash")
				FItemList(i).Fwi_totalsuplycash= rsget("wi_totalsuplycash")
				FItemList(i).Fwi_comment       = db2html(rsget("wi_comment"))

				FItemList(i).Fet_cnt           = rsget("et_cnt")
				FItemList(i).Fet_totalsellcash = rsget("et_totalsellcash")
				FItemList(i).Fet_totalsuplycash= rsget("et_totalsuplycash")
				FItemList(i).Fet_comment       = db2html(rsget("et_comment"))
				FItemList(i).Fsh_cnt           = rsget("sh_cnt")
				FItemList(i).Fsh_totalsellcash = rsget("sh_totalsellcash")
				FItemList(i).Fsh_totalsuplycash= rsget("sh_totalsuplycash")
				FItemList(i).Fsh_comment       = db2html(rsget("sh_comment"))

                FItemList(i).Fjgubun        = rsget("jgubun")
                FItemList(i).FitemvatYn     = rsget("itemvatYn")
                FItemList(i).Fwi_totalreducedprice = rsget("wi_totalreducedprice")
                FItemList(i).Fub_totalreducedprice = rsget("ub_totalreducedprice")
                FItemList(i).Fet_totalreducedprice = rsget("et_totalreducedprice")
                FItemList(i).Fdlv_totalreducedprice= rsget("dlv_totalreducedprice")
                FItemList(i).Fdlv_totalsuplycash   = rsget("dlv_totalsuplycash")
                FItemList(i).Ftotalcommission      = rsget("totalcommission")

				FItemList(i).Fregdate          = rsget("regdate")
				FItemList(i).Fcancelyn         = rsget("cancelyn")
				FItemList(i).Ffinishflag       = rsget("finishflag")
				FItemList(i).Fipkumdate        = rsget("ipkumdate")
				FItemList(i).Ftaxregdate       = rsget("taxregdate")
				FItemList(i).Fbigo			   = db2html(rsget("bigo"))
				FItemList(i).FDesignerEmail		= rsget("jungsan_email")

				FItemList(i).Fjungsan_bank		= rsget("jungsan_bank")
				FItemList(i).Fjungsan_date		= rsget("jungsan_date")
				FItemList(i).Fjungsan_acctno		= rsget("jungsan_acctno")
				FItemList(i).Fjungsan_acctname		= rsget("jungsan_acctname")
				FItemList(i).Fcompany_name		= db2html(rsget("company_name"))

				FItemList(i).Fjungsan_gubun		= db2html(rsget("jungsan_gubun"))
				FItemList(i).Ftaxinputdate	= rsget("taxinputdate")
				FItemList(i).Fcompany_no	= db2html(rsget("company_no"))

				FItemList(i).Fdifferencekey = rsget("differencekey")
				FItemList(i).Ftaxtype = rsget("taxtype")
				FItemList(i).FTaxLinkidx = rsget("taxlinkidx")
				FItemList(i).Fneotaxno = rsget("neotaxno")
                FItemList(i).Fbillsitecode = rsget("billsitecode")
				FItemList(i).Fbankingupflag = rsget("bankingupflag")

                FItemList(i).Fipkum_bank  = rsget("ipkum_bank")
                FItemList(i).Fipkum_acctno  = rsget("ipkum_acctno")

                FItemList(i).Fceoname = db2html(rsget("ceoname"))
                FItemList(i).Fcompany_address = db2html(rsget("company_address"))
                FItemList(i).Fcompany_address2 = db2html(rsget("company_address2"))

                FItemList(i).FipFileNo = rsget("ipFileNo")
                FItemList(i).FtargetGbn= rsget("targetGbn")

                FItemList(i).FholdGroupid   = rsget("holdGroupid")
                FItemList(i).Fholdcause     = rsget("holdcause")


				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

    end function

    ''2014 추가
    public function getJungsanSubSummary()
        dim sqlStr,i

        sqlStr = " select top 1000"
        sqlStr = sqlStr + " m.jgubun"
        sqlStr = sqlStr + " ,D.gubuncd"
        sqlStr = sqlStr + " ,Jc.comm_name as gubuncdName"
        sqlStr = sqlStr + " ,m.taxtype"
        sqlStr = sqlStr + " ,d.vatyn"
        sqlStr = sqlStr + " ,sum(D.itemno) as itemCNT"
        sqlStr = sqlStr + " ,sum(D.sellcash*D.itemno) as sellcashSum"
        sqlStr = sqlStr + " ,sum(D.suplycash*D.itemno) as suplycashSum"
        sqlStr = sqlStr + " ,sum(isNULL(D.reducedprice,0)*D.itemno) as reducedpriceSum"
        sqlStr = sqlStr + " ,sum(isNULL(D.commission,0)*D.itemno) as commissionSum"
        sqlStr = sqlStr + " ,sum(isNULL(D.pgcommission,0)*D.itemno) as PgcommissionSum"
        sqlStr = sqlStr + " ,sum(isNULL(D.CpnNotAppliedPrice,0)*D.itemno) as CpnNotAppliedPriceSum"
        sqlStr = sqlStr + " from db_jungsan.dbo.tbl_designer_jungsan_detail d with (nolock)"
        sqlStr = sqlStr + " 	Join db_jungsan.dbo.tbl_designer_jungsan_master m with (nolock)"
        sqlStr = sqlStr + " 	on m.id=D.masteridx"
        sqlStr = sqlStr + " 	left join db_jungsan.dbo.tbl_jungsan_comm_code jc with (nolock)"
        sqlStr = sqlStr + " 	on D.gubuncd=jc.comm_cd"
        sqlStr = sqlStr + "     and jc.comm_group in ('Z003')"
        sqlStr = sqlStr + " where D.masteridx="&FRectId
        if (FRectDesigner="") then
			sqlStr = sqlStr + " and m.designerid='" + FRectDesigner + "'"
		end if
        sqlStr = sqlStr + " and gubuncd<>'maeipchulgo'"
        sqlStr = sqlStr + " group by m.jgubun, Jc.comm_name, D.gubuncd,m.taxtype,d.vatyn"
        sqlStr = sqlStr + " order by (CASE WHEN D.gubuncd='witakchulgo' THEN 'A' ELSE D.gubuncd END) desc"

		'response.write sqlStr & "<br>"
        rsget.PageSize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly


		FResultCount = rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    rsget.absolutepage = FCurrPage
			do until rsget.eof

				set FItemList(i) = new CJungsanSubSummaryItem
                FItemList(i).Fjgubun            = rsget("jgubun")
                FItemList(i).Fgubuncd           = rsget("gubuncd")
                FItemList(i).FtaxType           = rsget("taxType")
                FItemList(i).FitemVatyn         = rsget("vatyn")
                FItemList(i).FitemCNT           = rsget("itemCNT")
                FItemList(i).FsellcashSum       = rsget("sellcashSum")
                FItemList(i).FsuplycashSum      = rsget("suplycashSum")
                FItemList(i).FreducedpriceSum   = rsget("reducedpriceSum")
                FItemList(i).FcommissionSum     = rsget("commissionSum")
                FItemList(i).FgubuncdName       = rsget("gubuncdName")

                ''2016/09/27추가
                FItemList(i).FPgCommissionSum     = rsget("pgcommissionSum")

                ''2018/07/02추가
                FItemList(i).FCpnNotAppliedPriceSum       = rsget("CpnNotAppliedPriceSum")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

    end function

    '' 수수료 정산 관련 변경 있음
	public function JungsanMasterList()
		dim sqlStr,i
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + " inner join [db_partner].[dbo].tbl_partner pp on m.designerid = pp.id"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlStr = sqlStr + " where m.id<>0"

        if (FRectJGubun<>"") then
            sqlStr = sqlStr + " and m.jgubun='" + FRectJGubun + "'"
        end if

        if (FRecttargetGbn<>"") then
            sqlStr = sqlStr + " and m.targetGbn='" + FRecttargetGbn + "'"
        end if

        if (FRectOnlyYYYYMM<>"") then
            sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
        end if

        if (FRectOnlyCommissionType<>"") then
            sqlStr = sqlStr + " and m.jgubun='CC'"
        end if

		if (FRectTaxType<>"") then
		    sqlStr = sqlStr + " and m.taxtype='" + FRectTaxType + "'"
		end if

		if (Frectfinishflag<>"") then
            sqlStr = sqlStr + " and m.finishflag='" + Frectfinishflag + "'"
        end if

		if (FRectDesigner="") and (FRectGroupID="") and (FRectYYYYMM<>"") then
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectDesignerViewOnly=true then
			sqlStr = sqlStr + " and m.finishflag>0"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and m.designerid='" + FRectDesigner + "'"
		end if

        if (FRectGroupID<>"") then
            sqlStr = sqlStr + " and m.groupid='" + FRectGroupID + "'"
        end if

		if FRectID<>"" then
			sqlStr = sqlStr + " and m.id=" + CStr(FRectID)
		end if

		if FRectState<>"" then
			sqlStr = sqlStr + " and m.finishflag='" + FRectState + "'"
		end if

		if FRectIpkumilNot<>"" then
			sqlStr = sqlStr + " and p.jungsan_date<>'" + FRectIpkumilNot + "'"
		end if

		if FRectNotIncludeWonChon<>"" then
			''sqlStr = sqlStr + " and p.jungsan_gubun<>'원천징수'"
			sqlAdd = sqlAdd + " and m.taxtype<>'03'"
			sqlStr = sqlStr + " and p.jungsan_gubun<>'간이과세'"
			''''sqlStr = sqlStr + " and p.jungsan_gubun<>'면세'"
		end if

		if FRectOnlyIncludeWonChon<>"" then
			''sqlStr = sqlStr + " and p.jungsan_gubun='원천징수'"
			sqlStr = sqlStr + " and m.taxtype='03'"
		end if

		if FRectOnlyIncludeNoTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='면세'"
		end if

		if FRectOnlyIncludeSimpleTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='간이과세'"
		end if

		if FRectJungsanGubun<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='" & FRectJungsanGubun & "'"
		end if

		if FRectOnlyElecTax<>"" then
			sqlStr = sqlStr + " and m.neotaxno is Not NULL"
		end if

		if FRectOnlyNotElecTax<>"" then
			sqlStr = sqlStr + " and m.neotaxno is NULL"
		end if

		if FRectBankingupflag<>"" then
			sqlStr = sqlStr + " and m.bankingupflag='" + FRectBankingupflag + "'"
		end if

		if FRectNotYYYYMM<>"" then
			sqlStr = sqlStr + " and m.yyyymm<>'" + FRectNotYYYYMM + "'"
		end if

		IF (FRectUnderMargin<>"") Then
		    sqlStr = sqlStr + " and (m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash +  m.et_totalsellcash + m.sh_totalsellcash)<>0"
		    sqlStr = sqlStr + " and 100-(m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.et_totalsuplycash + m.sh_totalsuplycash + m.dlv_totalsuplycash)/(m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash +  m.et_totalsellcash + m.sh_totalsellcash+ m.dlv_totalsellcash)*100<"&FRectUnderMargin
		ENd IF

		if FRectPurchaseType<>"" then
			sqlStr = sqlStr + " and pp.PurchaseType = '" + FRectPurchaseType + "'"
		end if

		if FRectjacctcd<>"" then
			sqlStr = sqlStr + " and m.jacctcd = '" + FRectjacctcd + "'"
		end if

		if (FRectdifferencekey<>"") then
		    sqlStr = sqlStr + " and m.differencekey = '" + FRectdifferencekey + "'"
		end if

		If FRectCompanynoYN <> "" Then
			Select Case FRectCompanynoYN
				Case "Y"		sqlStr = sqlStr & " and replace(pp.company_no,'-','') = '2118700620' "
				Case "N"		sqlStr = sqlStr & " and replace(pp.company_no,'-','') <> '2118700620' "
			End Select
		End If

		if (FRectSearchType <> "") and (FRectSearchText <> "") then
			Select Case FRectSearchType
				Case "socname"
					sqlStr = sqlStr + " and m.designerid in ( "
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
					sqlStr = sqlStr + " and m.designerid in ( "
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

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.id, m.designerid, m.groupid, m.yyyymm, m.title, m.ub_cnt,"
		sqlStr = sqlStr + " m.ub_totalsellcash,"
		sqlStr = sqlStr + " m.ub_totalsuplycash,"
		sqlStr = sqlStr + " m.ub_comment,"
		sqlStr = sqlStr + " m.me_cnt, m.me_totalsellcash,"
		sqlStr = sqlStr + " m.me_totalsuplycash, m.me_comment,"
		sqlStr = sqlStr + " m.wi_cnt, m.wi_totalsellcash,"
		sqlStr = sqlStr + " m.wi_totalsuplycash, m.wi_comment,"
		sqlStr = sqlStr + " m.et_cnt, m.et_totalsellcash,"
		sqlStr = sqlStr + " m.et_totalsuplycash, m.et_comment,"
		sqlStr = sqlStr + " m.sh_cnt, m.sh_totalsellcash,"
		sqlStr = sqlStr + " m.sh_totalsuplycash, m.sh_comment,"

		sqlStr = sqlStr + " m.regdate,m.cancelyn,m.finishflag,convert(varchar(10),m.ipkumdate,20) as ipkumdate,"
		sqlStr = sqlStr + " convert(varchar(10),m.taxregdate,20) as taxregdate, m.bigo, "
		sqlStr = sqlStr + " p.jungsan_email,p.jungsan_bank,p.jungsan_date,p.jungsan_acctno,p.jungsan_hp,"
		sqlStr = sqlStr + " p.jungsan_acctname,p.company_name, p.jungsan_gubun,p.company_no, b.billSiteName,"
		sqlStr = sqlStr + " m.taxinputdate, m.differencekey, m.taxtype, m.taxlinkidx, m.neotaxno, m.bankingupflag"
		''sqlStr = sqlStr + " ,IsNULL(m.availneo,0) as availneo,m.ISSU_SEQNO"
		sqlStr = sqlStr + " ,m.BillsiteCode,m.eseroEvalSeq,m.preFixedTaxDate"

		sqlStr = sqlStr + " , m.jgubun, m.itemvatYn, isNULL(m.targetGbn,'ON') as targetGbn"
		sqlStr = sqlStr + " , isNULL(m.wi_totalreducedprice,0) as wi_totalreducedprice"
		sqlStr = sqlStr + " , isNULL(m.ub_totalreducedprice,0) as ub_totalreducedprice"
		sqlStr = sqlStr + " , isNULL(m.et_totalreducedprice,0) as et_totalreducedprice"
		sqlStr = sqlStr + " , isNULL(m.dlv_totalreducedprice,0) as dlv_totalreducedprice"
		sqlStr = sqlStr + " , isNULL(m.dlv_totalsuplycash,0) as dlv_totalsuplycash"
		sqlStr = sqlStr + " , isNULL(m.totalcommission,0) as totalcommission"
        sqlStr = sqlStr + " , m.jacctcd, c.acc_nm"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + " 	inner join [db_partner].[dbo].tbl_partner pp on m.designerid = pp.id"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlStr = sqlStr + "     left join db_jungsan.dbo.tbl_tax_asp_Info b on m.billsiteCode=b.BillSiteCode"
		sqlStr = sqlStr + "     left join db_partner.dbo.tbl_TMS_SL_ACC_CD c on m.jacctcd=c.acc_Use_cd"
		sqlStr = sqlStr + " where m.id<>0"

        if (FRectJGubun<>"") then
            sqlStr = sqlStr + " and m.jgubun='" + FRectJGubun + "'"
        end if

        if (FRecttargetGbn<>"") then
            sqlStr = sqlStr + " and m.targetGbn='" + FRecttargetGbn + "'"
        end if

        if (FRectTaxType<>"") then
		    sqlStr = sqlStr + " and m.taxtype='" + FRectTaxType + "'"
		end if

        if (FRectOnlyYYYYMM<>"") then
            sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
        end if

        if (FRectOnlyCommissionType<>"") then
            sqlStr = sqlStr + " and m.jgubun='CC'"
        end if

        if (Frectfinishflag<>"") then
            sqlStr = sqlStr + " and m.finishflag='" + Frectfinishflag + "'"
        end if

		if (FRectDesigner="") and (FRectGroupID="") and (FRectYYYYMM<>"") then
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectDesignerViewOnly=true then
			sqlStr = sqlStr + " and m.finishflag>0"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and m.designerid='" + FRectDesigner + "'"
		end if

        if (FRectGroupID<>"") then
            sqlStr = sqlStr + " and m.groupid='" + FRectGroupID + "'"
        end if

		if FRectID<>"" then
			sqlStr = sqlStr + " and m.id=" + CStr(FRectID)
		end if

		if FRectState<>"" then
			sqlStr = sqlStr + " and m.finishflag='" + FRectState + "'"
		end if

		if FRectIpkumilNot<>"" then
			sqlStr = sqlStr + " and p.jungsan_date<>'" + FRectIpkumilNot + "'"
		end if

		if FRectNotIncludeWonChon<>"" then
			''sqlStr = sqlStr + " and p.jungsan_gubun<>'원천징수'"
			sqlAdd = sqlAdd + " and m.taxtype<>'03'"
			sqlStr = sqlStr + " and p.jungsan_gubun<>'간이과세'"
			''''sqlStr = sqlStr + " and p.jungsan_gubun<>'면세'"
		end if

		if FRectOnlyIncludeWonChon<>"" then
			''sqlStr = sqlStr + " and p.jungsan_gubun='원천징수'"
			sqlStr = sqlStr + " and m.taxtype='03'"
		end if

		if FRectOnlyIncludeNoTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='면세'"
		end if

		if FRectOnlyIncludeSimpleTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='간이과세'"
		end if

		if FRectJungsanGubun<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='" & FRectJungsanGubun & "'"
		end if

		if FRectOnlyElecTax<>"" then
			sqlStr = sqlStr + " and m.neotaxno is Not NULL"
		end if

		if FRectOnlyNotElecTax<>"" then
			sqlStr = sqlStr + " and m.neotaxno is NULL"
		end if

		if FRectBankingupflag<>"" then
			sqlStr = sqlStr + " and m.bankingupflag='" + FRectBankingupflag + "'"
		end if

		if FRectNotYYYYMM<>"" then
			sqlStr = sqlStr + " and m.yyyymm<>'" + FRectNotYYYYMM + "'"
		end if

		IF (FRectUnderMargin<>"") Then
		    sqlStr = sqlStr + " and (m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash +  m.et_totalsellcash + m.sh_totalsellcash)<>0"
		    sqlStr = sqlStr + " and 100-(m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.et_totalsuplycash + m.sh_totalsuplycash+ m.dlv_totalsuplycash)/(m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash +  m.et_totalsellcash + m.sh_totalsellcash+  m.dlv_totalsellcash)*100<"&FRectUnderMargin
		ENd IF

		if FRectPurchaseType<>"" then
			sqlStr = sqlStr + " and pp.PurchaseType = '" + FRectPurchaseType + "'"
		end if

        if FRectjacctcd<>"" then
			sqlStr = sqlStr + " and m.jacctcd = '" + FRectjacctcd + "'"
		end if

        if (FRectdifferencekey<>"") then
		    sqlStr = sqlStr + " and m.differencekey = '" + FRectdifferencekey + "'"
		end if

		If FRectCompanynoYN <> "" Then
			Select Case FRectCompanynoYN
				Case "Y"		sqlStr = sqlStr & " and replace(pp.company_no,'-','') = '2118700620' "
				Case "N"		sqlStr = sqlStr & " and replace(pp.company_no,'-','') <> '2118700620' "
			End Select
		End If

		if (FRectSearchType <> "") and (FRectSearchText <> "") then
			Select Case FRectSearchType
				Case "socname"
					sqlStr = sqlStr + " and m.designerid in ( "
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
					sqlStr = sqlStr + " and m.designerid in ( "
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

		if FrectOrder="state" then
			sqlStr = sqlStr + " order by m.finishflag "
		elseif FrectOrder="segum" then
			sqlStr = sqlStr + " order by m.taxregdate "
		elseif FrectOrder="designer" then
			sqlStr = sqlStr + " order by m.designerid "
		elseif FrectOrder="jungsanfix" then
			sqlStr = sqlStr + " order by m.yyyymm desc, p.jungsan_date desc"
		elseif FrectOrder="tax" then
			sqlStr = sqlStr + " order by p.jungsan_gubun , m.yyyymm desc"
		elseif FrectOrder="taxinputdate" then
			''sqlStr = sqlStr + " order by p.jungsan_date desc, m.taxinputdate"
			sqlStr = sqlStr + " order by m.yyyymm desc, p.jungsan_date desc, m.neotaxno, m.taxinputdate"
		elseif FrectOrder="taxinputdate_last" then
			sqlStr = sqlStr + " order by m.neotaxno, m.taxinputdate"

		else
			sqlStr = sqlStr + " order by m.id desc"
		end if



'rw sqlStr
		rsget.PageSize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    rsget.absolutepage = FCurrPage
			do until rsget.eof

				set FItemList(i) = new CJungsanMasterItem

				FItemList(i).Fid               = rsget("id")
				FItemList(i).Fdesignerid       = rsget("designerid")
				FItemList(i).Fgroupid		   = rsget("groupid")
				FItemList(i).Fyyyymm           = rsget("yyyymm")
				FItemList(i).Ftitle            = rsget("title")
				FItemList(i).Fub_cnt           = rsget("ub_cnt")
				FItemList(i).Fub_totalsellcash = rsget("ub_totalsellcash")
				FItemList(i).Fub_totalsuplycash= rsget("ub_totalsuplycash")
				FItemList(i).Fub_comment       = db2html(rsget("ub_comment"))
				FItemList(i).Fme_cnt           = rsget("me_cnt")
				FItemList(i).Fme_totalsellcash = rsget("me_totalsellcash")
				FItemList(i).Fme_totalsuplycash= rsget("me_totalsuplycash")
				FItemList(i).Fme_comment       = db2html(rsget("me_comment"))
				FItemList(i).Fwi_cnt           = rsget("wi_cnt")
				FItemList(i).Fwi_totalsellcash = rsget("wi_totalsellcash")
				FItemList(i).Fwi_totalsuplycash= rsget("wi_totalsuplycash")
				FItemList(i).Fwi_comment       = db2html(rsget("wi_comment"))

				FItemList(i).Fet_cnt           = rsget("et_cnt")
				FItemList(i).Fet_totalsellcash = rsget("et_totalsellcash")
				FItemList(i).Fet_totalsuplycash= rsget("et_totalsuplycash")
				FItemList(i).Fet_comment       = db2html(rsget("et_comment"))
				FItemList(i).Fsh_cnt           = rsget("sh_cnt")
				FItemList(i).Fsh_totalsellcash = rsget("sh_totalsellcash")
				FItemList(i).Fsh_totalsuplycash= rsget("sh_totalsuplycash")
				FItemList(i).Fsh_comment       = db2html(rsget("sh_comment"))


				FItemList(i).Fregdate          = rsget("regdate")
				FItemList(i).Fcancelyn         = rsget("cancelyn")
				FItemList(i).Ffinishflag       = rsget("finishflag")
				FItemList(i).Fipkumdate        = rsget("ipkumdate")
				FItemList(i).Ftaxregdate       = rsget("taxregdate")
				FItemList(i).Fbigo			   = db2html(rsget("bigo"))
				FItemList(i).FDesignerEmail		= rsget("jungsan_email")

				FItemList(i).Fjungsan_bank		= rsget("jungsan_bank")
				FItemList(i).Fjungsan_date		= rsget("jungsan_date")
				FItemList(i).Fjungsan_acctno	= rsget("jungsan_acctno")
				FItemList(i).Fjungsan_acctname	= rsget("jungsan_acctname")
				FItemList(i).Fcompany_name	= db2html(rsget("company_name"))

				FItemList(i).Fjungsan_gubun	= db2html(rsget("jungsan_gubun"))
				FItemList(i).Ftaxinputdate	= rsget("taxinputdate")
				FItemList(i).Fcompany_no	= db2html(rsget("company_no"))

				FItemList(i).Fdifferencekey = rsget("differencekey")
				FItemList(i).Ftaxtype      = rsget("taxtype")
				FItemList(i).FTaxLinkidx   = rsget("taxlinkidx")
				FItemList(i).Fneotaxno     = rsget("neotaxno")

				FItemList(i).Fbankingupflag = rsget("bankingupflag")
                ''FItemList(i).Favailneo     = rsget("availneo")
                FItemList(i).Fjungsan_hp   = rsget("jungsan_hp")

                FItemList(i).FBillsiteCode = rsget("BillsiteCode")
                ''FItemList(i).FISSU_SEQNO   = rsget("ISSU_SEQNO")
                FItemList(i).FeseroEvalSeq = rsget("eseroEvalSeq")
                FItemList(i).FbillSiteName = rsget("billSiteName")
                FItemList(i).FpreFixedTaxDate   = rsget("preFixedTaxDate")          ''2012-09-03 추가

                ''2014/01/27 추가
                FItemList(i).FtargetGbn= rsget("targetGbn")
                FItemList(i).Fjgubun                = rsget("jgubun")
                FItemList(i).Fwi_totalreducedprice  = rsget("wi_totalreducedprice")
                FItemList(i).Fub_totalreducedprice  = rsget("ub_totalreducedprice")
                FItemList(i).Fet_totalreducedprice  = rsget("et_totalreducedprice")
                FItemList(i).Fdlv_totalreducedprice = rsget("dlv_totalreducedprice")
                FItemList(i).Fdlv_totalsuplycash    = rsget("dlv_totalsuplycash")
                FItemList(i).Ftotalcommission       = rsget("totalcommission")
                FItemList(i).FitemvatYn             = rsget("itemvatYn")
                FItemList(i).Fjacctcd               = rsget("jacctcd")
                FItemList(i).Fjacc_nm               = rsget("acc_nm")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

	public function JungsanMasterListSimple()
		dim sqlStr,i
		sqlStr = "select count(m.id) as cnt from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + " where m.id<>0"
		if (FRectDesigner="") and (FRectYYYYMM<>"") then
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectDesignerViewOnly=true then
			sqlStr = sqlStr + " and m.finishflag>0"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and m.designerid='" + FRectDesigner + "'"
		end if

		if FRectID<>"" then
			sqlStr = sqlStr + " and m.id=" + CStr(FRectID)
		end if

		if FRectState<>"" then
			sqlStr = sqlStr + " and m.finishflag='" + FRectState + "'"
		end if

		if FRectIpkumilNot<>"" then
			sqlStr = sqlStr + " and p.jungsan_date<>'" + FRectIpkumilNot + "'"
		end if

		rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				FTotalCount = rsget("cnt")
			end if
		rsget.Close


		sqlStr = "select top " + CStr(FCurrPage * FPageSize) + " m.id,m.designerid,m.yyyymm,m.title,"
		sqlStr = sqlStr + " m.ub_cnt,"
		sqlStr = sqlStr + " m.ub_totalsellcash,"
		sqlStr = sqlStr + " m.ub_totalsuplycash,"
		sqlStr = sqlStr + " m.ub_comment,"
		sqlStr = sqlStr + " m.me_cnt,m.me_totalsellcash,"
		sqlStr = sqlStr + " m.me_totalsuplycash, m.me_comment,"
		sqlStr = sqlStr + " m.wi_cnt, m.wi_totalsellcash,"
		sqlStr = sqlStr + " m.wi_totalsuplycash, m.wi_comment,"
		sqlStr = sqlStr + " m.et_cnt, m.et_totalsellcash,"
		sqlStr = sqlStr + " m.et_totalsuplycash, m.et_comment,"
		sqlStr = sqlStr + " m.sh_cnt, m.sh_totalsellcash,"
		sqlStr = sqlStr + " m.sh_totalsuplycash, m.sh_comment,"
		sqlStr = sqlStr + " m.regdate,m.cancelyn,m.finishflag"
		sqlStr = sqlStr + " ,m.dlv_totalsuplycash"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + " where m.id<>0"

		if (FRectDesigner="") and (FRectYYYYMM<>"") then
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectDesignerViewOnly=true then
			sqlStr = sqlStr + " and m.finishflag>0"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and m.designerid='" + FRectDesigner + "'"
		end if

		if FRectID<>"" then
			sqlStr = sqlStr + " and m.id=" + CStr(FRectID)
		end if

		if FRectState<>"" then
			sqlStr = sqlStr + " and m.finishflag='" + FRectState + "'"
		end if

		if FRectIpkumilNot<>"" then
			sqlStr = sqlStr + " and p.jungsan_date<>'" + FRectIpkumilNot + "'"
		end if

		if FrectOrder="state" then
			sqlStr = sqlStr + " order by m.finishflag "
		elseif FrectOrder="segum" then
			sqlStr = sqlStr + " order by m.taxregdate "
		elseif FrectOrder="designer" then
			sqlStr = sqlStr + " order by m.designerid "
		elseif FrectOrder="jungsanfix" then
			sqlStr = sqlStr + " order by m.yyyymm desc, p.jungsan_date desc"
		elseif FrectOrder="tax" then
			sqlStr = sqlStr + " order by p.jungsan_gubun , m.yyyymm desc"
		else
			sqlStr = sqlStr + " order by m.id desc"
		end if



		'response.write sqlStr
		rsget.PageSize = FPageSize
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
				set FItemList(i) = new CJungsanMasterItem

				FItemList(i).Fid               = rsget("id")
				FItemList(i).Fdesignerid       = rsget("designerid")
				FItemList(i).Fyyyymm           = rsget("yyyymm")
				FItemList(i).Ftitle            = rsget("title")
				FItemList(i).Fub_cnt           = rsget("ub_cnt")
				FItemList(i).Fub_totalsellcash = rsget("ub_totalsellcash")
				FItemList(i).Fub_totalsuplycash= rsget("ub_totalsuplycash")
				FItemList(i).Fub_comment       = db2html(rsget("ub_comment"))
				FItemList(i).Fme_cnt           = rsget("me_cnt")
				FItemList(i).Fme_totalsellcash = rsget("me_totalsellcash")
				FItemList(i).Fme_totalsuplycash= rsget("me_totalsuplycash")
				FItemList(i).Fme_comment       = db2html(rsget("me_comment"))
				FItemList(i).Fwi_cnt           = rsget("wi_cnt")
				FItemList(i).Fwi_totalsellcash = rsget("wi_totalsellcash")
				FItemList(i).Fwi_totalsuplycash= rsget("wi_totalsuplycash")
				FItemList(i).Fwi_comment       = db2html(rsget("wi_comment"))

				FItemList(i).Fet_cnt           = rsget("et_cnt")
				FItemList(i).Fet_totalsellcash = rsget("et_totalsellcash")
				FItemList(i).Fet_totalsuplycash= rsget("et_totalsuplycash")
				FItemList(i).Fet_comment       = db2html(rsget("et_comment"))
				FItemList(i).Fsh_cnt           = rsget("sh_cnt")
				FItemList(i).Fsh_totalsellcash = rsget("sh_totalsellcash")
				FItemList(i).Fsh_totalsuplycash= rsget("sh_totalsuplycash")
				FItemList(i).Fsh_comment       = db2html(rsget("sh_comment"))


				FItemList(i).Fregdate          = rsget("regdate")
				FItemList(i).Fcancelyn         = rsget("cancelyn")
				FItemList(i).Ffinishflag       = rsget("finishflag")

				FItemList(i).Fdlv_totalsuplycash = rsget("dlv_totalsuplycash")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

	public function GetMasterCodes()
		dim mastercodes, cnt, i
		cnt = UBound(FItemList)
		mastercodes = ""
		for i=0 to cnt-1
			mastercodes = mastercodes + FItemList(i).FCode + ","
		next

		if Right(mastercodes,1)="," then
			mastercodes = left(mastercodes,Len(mastercodes)-1)
		end if

		mastercodes = replace(mastercodes,",","','")
		GetMasterCodes = mastercodes
	end function

	public sub SearchMaeIpJungsanDetailList()
		dim sqlStr,i

		sqlStr = "select d.id, d.mastercode, d.itemid, d.itemoption,"
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.itemno,"
		sqlStr = sqlStr + " i.itemname, IsNull(d.iitemoptionname,'') as itemoptionname,"
		sqlStr = sqlStr + " m.executedt, m.code, m.indt, m.totalsuplycash, i.vatinclude"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m, [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.socid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and m.code=d.mastercode"
		sqlStr = sqlStr + " and m.divcode='001'"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and datediff(month,m.executedt,getdate())<5"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and d.id not in (select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail where gubuncd='maeip')"
		sqlStr = sqlStr + " order by d.id desc"

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
				set FItemList(i) = new CMaeIpJungsanDetailItem

				FItemList(i).FID         = rsget("id")
				FItemList(i).FMasterCode = rsget("mastercode")
				FItemList(i).FDesignerID = rsget("id")

				FItemList(i).FItemName		= db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemId     = rsget("itemid")
				FItemList(i).FItemOption = rsget("itemoption")
				FItemList(i).FItemNo     = rsget("itemno")
				FItemList(i).Fsellcash   = rsget("sellcash")
				FItemList(i).Fsuplycash  = rsget("suplycash")

				FItemList(i).FExecuteDate  = rsget("executedt")
				FItemList(i).FRegDate   = rsget("indt")
				FItemList(i).FCode		= rsget("code")
				FItemList(i).FTotalsuplycash = rsget("totalsuplycash")

				FItemList(i).Fvatinclude = rsget("vatinclude")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

	public sub SearchChulGoDetailList()
		dim sqlStr,i

		sqlStr = "select d.id, d.mastercode, d.itemid, d.itemoption,"
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.itemno, d.buycash,"
		sqlStr = sqlStr + " d.iitemname as itemname, d.iitemoptionname as itemoptionname, IsNULL(d.mwgubun,'') as mwgubun, d.iitemgubun,"
		sqlStr = sqlStr + " m.executedt, m.code, m.divcode, m.indt, m.totalsuplycash, m.socid, m.socname"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m, [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.code=d.mastercode"
		sqlStr = sqlStr + " and d.imakerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and Left(m.code,2)='SO'"
		sqlStr = sqlStr + " and convert(varchar(7),m.executedt,21)='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and d.itemno<>0"
		sqlStr = sqlStr + " order by d.mastercode desc,d.id "

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
				set FItemList(i) = new CMaeIpJungsanDetailItem

				FItemList(i).FID         = rsget("id")
				FItemList(i).FMasterCode = rsget("mastercode")
				FItemList(i).FDesignerID = rsget("id")

				FItemList(i).FItemName		= db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemId     = rsget("itemid")
				FItemList(i).FItemOption = rsget("itemoption")
				FItemList(i).FItemNo     = rsget("itemno")
				FItemList(i).Fsellcash   = rsget("sellcash")
				FItemList(i).Fsuplycash  = rsget("suplycash")
				FItemList(i).Fbuycash  = rsget("buycash")

				FItemList(i).FExecuteDate  = rsget("executedt")
				FItemList(i).FRegDate   = rsget("indt")
				FItemList(i).FCode		= rsget("code")
				FItemList(i).FDivCode		= rsget("divcode")
				FItemList(i).FTotalsuplycash = rsget("totalsuplycash")

				FItemList(i).Fsocid = rsget("socid")
				FItemList(i).FSocName = db2html(rsget("socname"))
				FItemList(i).Fmwgubun = rsget("mwgubun")
				FItemList(i).Fitemgubun = rsget("iitemgubun")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

	public sub SearchIpGoDetailList()
		dim sqlStr,i

		sqlStr = "select d.id, d.mastercode, d.itemid, d.itemoption,"
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.itemno,"
		sqlStr = sqlStr + " d.iitemname, d.iitemoptionname,"
		sqlStr = sqlStr + " m.executedt, m.code, m.divcode, m.indt, m.totalsuplycash,m.comment"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m, [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.socid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and m.code=d.mastercode"
		sqlStr = sqlStr + " and Left(m.code,2)='ST'"
		sqlStr = sqlStr + " and convert(varchar(7),m.executedt,21)='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and d.itemno<>0"
		sqlStr = sqlStr + " order by d.mastercode desc,d.id "

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
				set FItemList(i) = new CMaeIpJungsanDetailItem

				FItemList(i).FID         = rsget("id")
				FItemList(i).FMasterCode = rsget("mastercode")
				FItemList(i).FDesignerID = rsget("id")

				FItemList(i).FItemName		= db2html(rsget("iitemname"))
				FItemList(i).FItemOptionName= db2html(rsget("iitemoptionname"))
				FItemList(i).FItemId     = rsget("itemid")
				FItemList(i).FItemOption = rsget("itemoption")
				FItemList(i).FItemNo     = rsget("itemno")
				FItemList(i).Fsellcash   = rsget("sellcash")
				FItemList(i).Fsuplycash  = rsget("suplycash")

				FItemList(i).FExecuteDate  = rsget("executedt")
				FItemList(i).FRegDate   = rsget("indt")
				FItemList(i).FCode		= rsget("code")
				FItemList(i).FDivCode		= rsget("divcode")
				FItemList(i).FTotalsuplycash = rsget("totalsuplycash")

				FItemList(i).Fcomment = db2html(rsget("comment"))

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

	public sub SearchIpChulGoDetailList()
		dim sqlStr,i

		sqlStr = "select d.id, d.mastercode, d.itemid, d.itemoption,"
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.itemno,"
		sqlStr = sqlStr + " i.itemname, IsNull(d.iitemoptionname,'') as itemoptionname,"
		sqlStr = sqlStr + " m.executedt, m.code, m.divcode, m.indt, m.totalsuplycash"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m, [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.socid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and m.code=d.mastercode"
		sqlStr = sqlStr + " and Left(m.code,2) in ('ST','SO')"
		sqlStr = sqlStr + " and convert(varchar(7),m.executedt,21)='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and d.itemno<>0"
		sqlStr = sqlStr + " order by d.mastercode desc,d.id "

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
				set FItemList(i) = new CMaeIpJungsanDetailItem

				FItemList(i).FID         = rsget("id")
				FItemList(i).FMasterCode = rsget("mastercode")
				FItemList(i).FDesignerID = rsget("id")

				FItemList(i).FItemName		= db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemId     = rsget("itemid")
				FItemList(i).FItemOption = rsget("itemoption")
				FItemList(i).FItemNo     = rsget("itemno")
				FItemList(i).Fsellcash   = rsget("sellcash")
				FItemList(i).Fsuplycash  = rsget("suplycash")

				FItemList(i).FExecuteDate  = rsget("executedt")
				FItemList(i).FRegDate   = rsget("indt")
				FItemList(i).FCode		= rsget("code")
				FItemList(i).FDivCode		= rsget("divcode")
				FItemList(i).FTotalsuplycash = rsget("totalsuplycash")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

	public sub SearchWitakJungsanItemList()
		dim sqlStr,i
		sqlStr = "select d.idx, d.orderserial, d.itemid, d.itemname, d.itemoption,"
		sqlStr = sqlStr + "  d.itemno, d.itemoptionname, d.itemcost, d.buycash,"
		sqlStr = sqlStr + "  d.currstate, convert(varchar(19),d.beasongdate,20) as beasongdate, m.buyname, convert(varchar(19),m.regdate,20) as regdate, convert(varchar(19),m.ipkumdate,20) as ipkumdate,"
		sqlStr = sqlStr + "  m.jumundiv, m.ipkumdiv"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"

		sqlStr = sqlStr + " ,(select distinct d.itemid"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.code=d.mastercode"
		sqlStr = sqlStr + " and m.divcode='002'"
		sqlStr = sqlStr + " and m.socid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and datediff(month,m.executedt,getdate())<6) as T"

		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and datediff(month,m.regdate,getdate())<3"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.itemid=T.itemid"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " order by d.orderserial desc"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CUpcheJungsanItem
				FItemList(i).FIdx           = rsget("idx")
				FItemList(i).FOrderSerial   = rsget("orderserial")
				FItemList(i).FItemId        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo        = rsget("itemno")
				FItemList(i).FBuyCash       = rsget("buycash")
				FItemList(i).FSellCash      = rsget("itemcost")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FBeasongdate      = rsget("beasongdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FBuyName		= rsget("buyname")
				FItemList(i).FJumunDiv		= rsget("jumundiv")
				FItemList(i).FIpkumDiv		= rsget("ipkumdiv")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub SearchWitakChulgoJungsanList()
		dim sqlStr,i
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.id as masterid, m.socid, m.code, m.divcode,"
		sqlStr = sqlStr + " m.executedt as chulgodate, m.scheduledt, m.chargeid, m.totalsellcash, m.totalsuplycash, m.vatcode,"
		sqlStr = sqlStr + " m.indt, d.id, d.itemid, d.itemoption, d.buycash, d.sellcash, d.itemno,"
		sqlStr = sqlStr + " i.itemname, IsNull(d.iitemoptionname,'') as itemoptionname"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.code=d.mastercode"
		'''sqlStr = sqlStr + " and m.divcode='002'"
		sqlStr = sqlStr + " and itemno<>0"
		sqlStr = sqlStr + " and ((Left(m.code,2)='SO') or (Left(m.code,2)='SR'))"
		sqlStr = sqlStr + " and datediff(month,m.indt,getdate())<5"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and m.executedt is Not NULL"
		sqlStr = sqlStr + " and (Left(m.socid,10) <> 'streetshop' or m.code='SO005826')"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and d.id not in (select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail where gubuncd in ('witakchulgo','maeipchulgo'))"
		sqlStr = sqlStr + " order by m.id desc"

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
				set FItemList(i) = new CMaeIpJungsanItem
				FItemList(i).FID   = rsget("id")
				FItemList(i).FMasterID   = rsget("masterid")
				FItemList(i).FDesignerID = rsget("socid")
				FItemList(i).FCode		 = rsget("code")
				FItemList(i).FDivCode    = rsget("divcode")
				FItemList(i).FExecuteDate   = rsget("chulgodate")
				FItemList(i).FScheduleDate   = rsget("scheduledt")
				FItemList(i).FRegDate	 = rsget("indt")

				FItemList(i).FChargeId   = rsget("chargeid")
				FItemList(i).FTotalsellcash = rsget("totalsellcash")
				FItemList(i).FTotalsuplycash= rsget("totalsuplycash")
				FItemList(i).FVatCode       = rsget("vatcode")

				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).Fsuplycash= rsget("buycash")

				FItemList(i).FItemId= rsget("itemid")
				FItemList(i).FItemOption= rsget("itemoption")
				FItemList(i).FItemNo= rsget("itemno")
				FItemList(i).FItemName= db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public function SearchWitakMaeipChulgoJungsanListGrp()
		dim sqlStr
		'mayjacctcd,makerid,vatinclude,mayitemnosum,mayitemcostsum,mayjungsansum,title,finishflag,jgubun,jacctcd,differencekey,et_cnt,et_totalsellcash,et_totalsuplycash,mayDiff
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_JungsanTarget_EtcChulgo] '"&FRectYYYYMM&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			SearchWitakMaeipChulgoJungsanListGrp = rsget.getRows()
		end if
		rsget.close()
	end function

	public sub SearchWitakMaeipChulgoJungsanList()
		dim sqlStr,i
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.id as masterid, m.socid, m.code, m.divcode,"
		sqlStr = sqlStr + " m.executedt as chulgodate, m.scheduledt, m.chargeid, m.chargename, m.totalsellcash, m.totalsuplycash, m.totalbuycash, m.vatcode,"
		sqlStr = sqlStr + " m.indt, m.comment, d.id, d.itemid, d.itemoption, d.buycash, d.sellcash, d.suplycash, d.itemno,"
		sqlStr = sqlStr + " d.iitemname as itemname, d.mwgubun as mwdiv, d.imakerid, d.iitemgubun, d.iitemoptionname as itemoptionname, i.vatinclude"
		sqlStr = sqlStr + " , od.detailidx"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + "     Join [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + "     on m.code=d.mastercode"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + "     on d.iitemgubun='10'"
		sqlStr = sqlStr + "     and d.itemid=i.itemid"
		sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail od "
		sqlStr = sqlStr + "     on d.id=od.detailidx and od.gubuncd in ('witakchulgo','maeipchulgo')"
		sqlStr = sqlStr + "     left join db_partner.dbo.tbl_partner p on m.socid=p.id and p.tplcompanyid='tplithinkso'"

		sqlStr = sqlStr + " where convert(varchar(7),m.executedt,20)='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and d.mwgubun<>'M'"
	''sqlStr = sqlStr + " and m.socid <> 'itemloss'"  ''주석(2014/04/04)

		''sqlStr = sqlStr + " and ((Left(m.code,2)='SO') or (Left(m.code,2)='SR'))"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and m.executedt is Not NULL"
		sqlStr = sqlStr + " and m.ipchulflag not in ('S','I')" ''' maybe E
		'sqlStr = sqlStr + " and Left(m.socid,11) <> 'streetshop0'"
		'sqlStr = sqlStr + " and Left(m.socid,12) <> 'streetshop80'"
		'sqlStr = sqlStr + " and Left(m.socid,10) <> 'streetshop'"
		'sqlStr = sqlStr + " and Left(m.socid,9) <> 'wholesale'"
		sqlStr = sqlStr + " and m.socid <> 'cafe002'"
		sqlStr = sqlStr + " and m.socid <> 'cafe003'"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and d.itemno<>0"

		sqlStr = sqlStr + " and p.id is NULL" ''3pl 제외.  //2016/03/31

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

        if FRectItemVatYn<>"" then
			sqlStr = sqlStr + " and i.vatinclude='" + FRectItemVatYn + "'"
		end if

		if (FRectNotIncDivcode999<>"") then
		    sqlStr = sqlStr + " and m.divcode<>'999'"
		    '' sqlStr = sqlStr + " and m.socid<>'3pl_its_etc'"  3pl 제외함.
		end if

		sqlStr = sqlStr + " and od.detailidx is NULL"
		'sqlStr = sqlStr + " and d.id not in "
		'sqlStr = sqlStr + "  (select d.detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_master m, [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		'sqlStr = sqlStr + "   where m.id=d.masteridx"
		'sqlStr = sqlStr + "   and m.yyyymm='" + FRectYYYYMM + "'"
		'sqlStr = sqlStr + "   and d.gubuncd in ('witakchulgo','maeipchulgo')"
		'sqlStr = sqlStr + "   and m.cancelyn='N'"
		'sqlStr = sqlStr + "   )"
		sqlStr = sqlStr + " order by m.id desc"
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
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMaeIpJungsanItem
				FItemList(i).FID   = rsget("id")
				FItemList(i).FMasterID   = rsget("masterid")
				FItemList(i).FDesignerID = rsget("socid")
				FItemList(i).FCode		 = rsget("code")
				FItemList(i).FDivCode    = rsget("divcode")
				FItemList(i).FExecuteDate   = rsget("chulgodate")
				FItemList(i).FScheduleDate   = rsget("scheduledt")
				FItemList(i).FRegDate	 = rsget("indt")

				FItemList(i).FChargeId   = rsget("chargeid")
				FItemList(i).Fchargename     = db2html(rsget("chargename"))
				FItemList(i).FTotalsellcash = rsget("totalsellcash")
				FItemList(i).FTotalsuplycash= rsget("totalsuplycash")
				FItemList(i).FTotalbuycash= rsget("totalbuycash")

				FItemList(i).FVatCode       = rsget("vatcode")

				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).Fsuplycash= rsget("buycash")
				FItemList(i).Fsuplycash2= rsget("suplycash")

                FItemList(i).FItemGubun= rsget("iitemgubun")
				FItemList(i).FItemId= rsget("itemid")
				FItemList(i).FItemOption= rsget("itemoption")
				FItemList(i).FItemNo= rsget("itemno")
				FItemList(i).FItemName= db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))

				FItemList(i).FComment= db2html(rsget("comment"))
				FItemList(i).FMWDiv= rsget("mwdiv")
				FItemList(i).FMakerid= rsget("imakerid")
				FItemList(i).Fvatinclude = rsget("vatinclude")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub SearchWitakJungsanList()
		dim sqlStr,i
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.id as masterid, m.socid, m.code, m.divcode,"
		sqlStr = sqlStr + " m.executedt as ipgodate, m.chargeid, m.totalsellcash, m.totalsuplycash, m.vatcode,"
		sqlStr = sqlStr + " m.indt, d.id, d.itemid, d.itemoption, d.suplycash, d.sellcash, d.itemno,"
		sqlStr = sqlStr + " i.itemname, IsNull(d.iitemoptionname,'') as itemoptionname"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.socid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and m.code=d.mastercode"
		sqlStr = sqlStr + " and Left(m.code,2)='ST'"
		sqlStr = sqlStr + " and m.divcode='002'"
		sqlStr = sqlStr + " and datediff(month,m.executedt,getdate())<3"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and d.id not in ("
		sqlStr = sqlStr + " 	select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
		sqlStr = sqlStr + "		where gubuncd='witak'"
		sqlStr = sqlStr + ")"

		sqlStr = sqlStr + " order by m.id desc"

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
				set FItemList(i) = new CMaeIpJungsanItem
				FItemList(i).FID   = rsget("id")
				FItemList(i).FMasterID   = rsget("masterid")
				FItemList(i).FDesignerID = rsget("socid")
				FItemList(i).FCode		 = rsget("code")
				FItemList(i).FDivCode    = rsget("divcode")
				FItemList(i).FExecuteDate   = rsget("ipgodate")
				FItemList(i).FRegDate	 = rsget("indt")
				FItemList(i).FChargeId   = rsget("chargeid")
				FItemList(i).FTotalsellcash = rsget("totalsellcash")
				FItemList(i).FTotalsuplycash= rsget("totalsuplycash")
				FItemList(i).FVatCode       = rsget("vatcode")

				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).Fsuplycash= rsget("suplycash")

				FItemList(i).FItemId= rsget("itemid")
				FItemList(i).FItemOption= rsget("itemoption")
				FItemList(i).FItemNo= rsget("itemno")
				FItemList(i).FItemName= db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub SearchMaeipChulGoJungsanList()
		dim sqlStr,i
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.id as masterid, m.socid, m.code, m.divcode,"
		sqlStr = sqlStr + " m.executedt as chulgodate, m.scheduledt, m.chargeid, m.totalsellcash, m.totalsuplycash, m.vatcode,"
		sqlStr = sqlStr + " m.indt, d.id, d.itemid, d.itemoption, d.suplycash, d.sellcash, d.itemno,"
		sqlStr = sqlStr + " i.itemname, IsNull(d.iitemoptionname,'') as itemoptionname"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.code=d.mastercode"
		'''sqlStr = sqlStr + " and m.divcode='002'"
		sqlStr = sqlStr + " and d.itemno<>0"
		sqlStr = sqlStr + " and ((Left(m.code,2)='SO') or (Left(m.code,2)='SR'))"
		sqlStr = sqlStr + " and datediff(month,m.indt,getdate())<3"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and m.executedt is Not NULL"
		sqlStr = sqlStr + " and Left(m.socid,10) <> 'streetshop'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and d.id not in (select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail where gubuncd in ('maeipchulgo','witakchulgo'))"
		sqlStr = sqlStr + " order by m.id desc"

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
				set FItemList(i) = new CMaeIpJungsanItem
				FItemList(i).FID   = rsget("id")
				FItemList(i).FMasterID   = rsget("masterid")
				FItemList(i).FDesignerID = rsget("socid")
				FItemList(i).FCode		 = rsget("code")
				FItemList(i).FDivCode    = rsget("divcode")
				FItemList(i).FExecuteDate   = rsget("chulgodate")
				FItemList(i).FScheduleDate   = rsget("scheduledt")
				FItemList(i).FRegDate	 = rsget("indt")

				FItemList(i).FChargeId   = rsget("chargeid")
				FItemList(i).FTotalsellcash = rsget("totalsellcash")
				FItemList(i).FTotalsuplycash= rsget("totalsuplycash")
				FItemList(i).FVatCode       = rsget("vatcode")

				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).Fsuplycash= rsget("suplycash")

				FItemList(i).FItemId= rsget("itemid")
				FItemList(i).FItemOption= rsget("itemoption")
				FItemList(i).FItemNo= rsget("itemno")
				FItemList(i).FItemName= db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub SearchChulGoJungsanList()
		dim sqlStr,i

		sqlStr = "select d.id, d.mastercode, d.itemid, d.itemoption,"
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.itemno,"
		sqlStr = sqlStr + " i.itemname, IsNull(d.iitemoptionname,'') as itemoptionname, m.socid, m.executedt, m.scheduledt, m.indt"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.code=d.mastercode"
		sqlStr = sqlStr + " and Left(d.mastercode,2)='SO'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and i.makerid='" + CStr(FRectDesigner) + "'"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and m.indt>='" + FRectStartDay + "'"
		end if

		sqlStr = sqlStr + " order by d.id desc"

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
				set FItemList(i) = new CMaeIpJungsanDetailItem

				FItemList(i).FID         = rsget("id")
				FItemList(i).FMasterCode = rsget("mastercode")
				FItemList(i).FDesignerID = rsget("id")

				FItemList(i).FItemName		= db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemId     = rsget("itemid")
				FItemList(i).FItemOption = rsget("itemoption")
				FItemList(i).FItemNo     = rsget("itemno")
				FItemList(i).Fsellcash   = rsget("sellcash")
				FItemList(i).Fsuplycash  = rsget("suplycash")
				FItemList(i).FSocID 	 = rsget("socid")
				FItemList(i).Fexecutedate  = rsget("executedt")
				FItemList(i).Fscheduledt  = rsget("scheduledt")
				FItemList(i).Fregdate  = rsget("indt")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub SearchMaeIpJungsanList()
		dim sqlStr,i
		sqlStr = "select  m.id, m.socid, m.code, m.divcode,"
		sqlStr = sqlStr + " m.executedt as ipgodate, m.chargeid, m.totalsellcash, m.totalsuplycash, m.vatcode,"
		sqlStr = sqlStr + " m.indt"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " where m.socid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and m.divcode='001'"
		sqlStr = sqlStr + " and datediff(month,m.executedt,getdate())<3"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " order by m.id desc"

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
				set FItemList(i) = new CMaeIpJungsanItem
				FItemList(i).FMasterID   = rsget("id")
				FItemList(i).FDesignerID = rsget("socid")
				FItemList(i).FCode		 = rsget("code")
				FItemList(i).FDivCode    = rsget("divcode")
				FItemList(i).FIpGoDate   = rsget("ipgodate")
				FItemList(i).FRegDate	 = rsget("indt")
				FItemList(i).FChargeId   = rsget("chargeid")
				FItemList(i).FTotalsellcash = rsget("totalsellcash")
				FItemList(i).FTotalsuplycash= rsget("totalsuplycash")
				FItemList(i).FVatCode       = rsget("vatcode")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub


	public sub SearchLectureJungsanList()
		dim sqlStr,i

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " d.detailidx, d.orderserial, d.itemid, d.itemname, d.itemoption,"
		sqlStr = sqlStr + "  d.itemno, d.itemoptionname, d.itemcost, d.buycash,"
		sqlStr = sqlStr + "  d.currstate, convert(varchar(19),d.beasongdate,20) as beasongdate,"
		sqlStr = sqlStr + "  d.songjangno, m.buyname, convert(varchar(19),m.regdate,20) as regdate, convert(varchar(19),m.ipkumdate,20) as ipkumdate,"
		sqlStr = sqlStr + "  m.jumundiv, i.lec_date"
		sqlStr = sqlStr + " from [110.93.128.73].[db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [110.93.128.73].[db_academy].[dbo].tbl_academy_order_detail d,"
		sqlStr = sqlStr + " [110.93.128.73].[db_academy].[dbo].tbl_lec_item i"

		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.itemid=i.idx"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and i.lec_date='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and m.orderserial not in ("
		sqlStr = sqlStr + " 	select d.mastercode from "
		sqlStr = sqlStr + " 	[db_jungsan].[dbo].tbl_designer_jungsan_master m,"
		sqlStr = sqlStr + " 	[db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		sqlStr = sqlStr + " 	where m.id=d.masteridx"
		sqlStr = sqlStr + " 	and m.designerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " 	and d.gubuncd='upche'"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by d.orderserial desc"

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
				set FItemList(i) = new CUpcheJungsanItem
				FItemList(i).FIdx           = rsget("detailidx")
				FItemList(i).FOrderSerial   = rsget("orderserial")
				FItemList(i).FItemId        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo        = rsget("itemno")
				FItemList(i).FBuyCash       = rsget("buycash")
				FItemList(i).FSellCash      = rsget("itemcost")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FBeasongdate      = rsget("beasongdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FBuyName		= rsget("buyname")
				FItemList(i).FJumunDiv		= rsget("jumundiv")
				FItemList(i).FUpcheSongjangNo		= rsget("songjangno")

				FItemList(i).Flec_date		= rsget("lec_date")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub


	public sub SearchUpCheBeasongJungsanList()
		dim sqlStr,i


		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " d.idx, d.orderserial, d.itemid, d.itemname, d.itemoption,"
		sqlStr = sqlStr + "  d.itemno, d.itemoptionname, d.itemcost, d.buycash,"
		sqlStr = sqlStr + "  d.currstate, convert(varchar(19),d.beasongdate,20) as beasongdate,"
		sqlStr = sqlStr + "  d.songjangno, m.buyname, convert(varchar(19),m.regdate,20) as regdate, convert(varchar(19),m.ipkumdate,20) as ipkumdate,"
		sqlStr = sqlStr + "  m.jumundiv, i.vatinclude"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i on d.itemid=i.itemid"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.sitename<>'academy'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	    'sqlStr = sqlStr + " and (d.itemid<>0) and (d.isupchebeasong='Y')"
	    ''sqlStr = sqlStr + " and ((d.itemid<>0) or ((d.itemid=0) and (d.makerid='" + FRectDesigner + "')))"
		sqlStr = sqlStr + " and (((d.itemid<>0) and (d.isupchebeasong='Y')) or ((d.itemid=0) and (d.makerid='" + FRectDesigner + "') and (d.buycash<>0)))"  ''2013/07/01
		sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and d.idx not in ("
		sqlStr = sqlStr + " 	select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
		sqlStr = sqlStr + " 	where gubuncd='upche'"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by d.orderserial desc"

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
				set FItemList(i) = new CUpcheJungsanItem
				FItemList(i).FIdx           = rsget("idx")
				FItemList(i).FOrderSerial   = rsget("orderserial")
				FItemList(i).FItemId        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo        = rsget("itemno")
				FItemList(i).FBuyCash       = rsget("buycash")
				FItemList(i).FSellCash      = rsget("itemcost")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FBeasongdate      = rsget("beasongdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FBuyName		= rsget("buyname")
				FItemList(i).FJumunDiv		= rsget("jumundiv")
				FItemList(i).FUpcheSongjangNo		= rsget("songjangno")

				FItemList(i).Fvatinclude = rsget("vatinclude")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub SearchWitakSellungsanList()
		dim sqlStr,i


		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " d.idx, d.orderserial, d.itemid, d.itemname, d.itemoption,"
		sqlStr = sqlStr + "  d.itemno, d.itemoptionname, d.itemcost, d.buycash,"
		sqlStr = sqlStr + "  d.currstate , m.ipkumdiv , convert(varchar(19),d.beasongdate,20) as beasongdate,"
		sqlStr = sqlStr + "  d.songjangno, m.buyname, convert(varchar(19),m.regdate,20) as regdate, convert(varchar(19),m.ipkumdate,20) as ipkumdate,"
		sqlStr = sqlStr + "  m.jumundiv, d.omwdiv as mwdiv, i.vatinclude"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_designer_jungsan_detail jj"
		sqlStr = sqlStr + "     on d.idx=jj.detailidx and jj.gubuncd ='witaksell'"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i on d.itemid=i.itemid"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.sitename<>'academy'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and d.omwdiv='W'" ''--추가
		sqlStr = sqlStr + " and (d.isupchebeasong<>'Y' or d.isupchebeasong is NULL )"
		sqlStr = sqlStr + " and jj.detailidx is NULL"

'		sqlStr = sqlStr + " and d.idx not in ("
'		sqlStr = sqlStr + "     select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
'		sqlStr = sqlStr + "     where gubuncd ='witaksell'"
'		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by d.orderserial desc"

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
				set FItemList(i) = new CUpcheJungsanItem
				FItemList(i).FIdx           = rsget("idx")
				FItemList(i).FOrderSerial   = rsget("orderserial")
				FItemList(i).FItemId        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo        = rsget("itemno")
				FItemList(i).FBuyCash       = rsget("buycash")
				FItemList(i).FSellCash      = rsget("itemcost")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FBeasongdate      = rsget("beasongdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FBuyName		= rsget("buyname")
				FItemList(i).FJumunDiv		= rsget("jumundiv")
				FItemList(i).FUpcheSongjangNo		= rsget("songjangno")
				FItemList(i).FMWDiv		= rsget("mwdiv")
				FItemList(i).FIpkumdiv  = rsget("ipkumdiv")

				FItemList(i).Fvatinclude = rsget("vatinclude")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub GetWitakJungSanSummary()
		dim sqlStr,i

		sqlStr = "select itemid,itemoption,itemname,itemoptionname, sellcash,suplycash,"
		sqlStr = sqlStr + " sum( Case gubuncd"
		sqlStr = sqlStr + " 	when 'witakchulgo' then itemno"
		sqlStr = sqlStr + " 	else 0"
		sqlStr = sqlStr + "          End"
		sqlStr = sqlStr + " ) as witakchulgo,"
		sqlStr = sqlStr + " sum( Case gubuncd"
		sqlStr = sqlStr + " 	when 'witaksell' then itemno"
		sqlStr = sqlStr + " 	else 0"
		sqlStr = sqlStr + "          End"
		sqlStr = sqlStr + " ) as witaksell,"
		sqlStr = sqlStr + " sum( Case gubuncd"
		sqlStr = sqlStr + " 	when 'witakoffshop' then itemno"
		sqlStr = sqlStr + " 	else 0"
		sqlStr = sqlStr + "          End"
		sqlStr = sqlStr + " ) as witakoffshop"

		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
		sqlStr = sqlStr + " where masteridx=" + Cstr(FRectID)
		sqlStr = sqlStr + " and gubuncd in ('witakchulgo','witaksell','witakoffshop')"
		sqlStr = sqlStr + " group by itemid, itemoption,itemname,itemoptionname,sellcash,suplycash,gubuncd"
		sqlStr = sqlStr + " order by itemid,itemoption"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CWitakJungSanItem
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FSellCash      = rsget("sellcash")
				FItemList(i).FSuplycash       = rsget("suplycash")

				FItemList(i).FChulgoNo = Null2Zero(rsget("witakchulgo"))
				FItemList(i).Fsellno	 = Null2Zero(rsget("witaksell"))
				FItemList(i).Foffsellno	 = Null2Zero(rsget("witakoffshop"))

				'FItemList(i).FIsUsing		= rsget("isusing")
				'FItemList(i).FIsDelete	    = rsget("isdelete")

				'FItemList(i).FDetailidx = rsget("id")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub GetWitakJungSanByItemView()
		dim sqlStr,i

		sqlStr = "select top 2000 i.itemid, i.itemname, w.itemoption, v.opt2name as itemoptionname,"
		sqlStr = sqlStr + " i.sellcash, i.buycash, i.isusing,"
		sqlStr = sqlStr + " w.ipgono,"
		sqlStr = sqlStr + " w.chulgono,"
		sqlStr = sqlStr + " w.sellcash as sellsellcash,"
		sqlStr = sqlStr + " w.suplycash as sellsuplycash,"
		sqlStr = sqlStr + " w.sellno,"
		sqlStr = sqlStr + " w.prejaego,"
		sqlStr = sqlStr + " w.realjaego,"
		sqlStr = sqlStr + " w.ocha,"
		sqlStr = sqlStr + " w.jungsanno,"
		sqlStr = sqlStr + " w.deleteyn as isdelete,"
		sqlStr = sqlStr + " w.id as detailidx"

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i, [db_jungsan].[dbo].tbl_designer_jungsan_witak w"
		sqlStr = sqlStr + " left join [db_item].[dbo].vw_itemoption v on w.itemid=v.itemid and w.itemoption=v.itemoption"
		sqlStr = sqlStr + " where w.masterid=" + FRectID
		sqlStr = sqlStr + " and w.deleteyn='N'"
		sqlStr = sqlStr + " and w.itemid=i.itemid"
		sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"

		sqlStr = sqlStr + " order by i.itemid, w.itemoption"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CWitakJungSanItem
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FSellCash      = rsget("sellcash")
				FItemList(i).FSuplycash       = rsget("buycash")
				FItemList(i).FSellcash_sell = Null2Zero(rsget("sellsellcash"))
				FItemList(i).FSuplycash_sell = Null2Zero(rsget("sellsuplycash"))
				FItemList(i).FIpGoNo = Null2Zero(rsget("ipgono"))
				FItemList(i).FChulgoNo = Null2Zero(rsget("chulgono"))
				FItemList(i).Fsellno	 = Null2Zero(rsget("sellno"))
				FItemList(i).Fprejaego   = Null2Zero(rsget("prejaego"))
				FItemList(i).Frealjaego   = Null2Zero(rsget("realjaego"))

				FItemList(i).FIsUsing		= rsget("isusing")
				FItemList(i).FIsDelete	    = rsget("isdelete")

				FItemList(i).FsysJaeGo = FItemList(i).FPrejaego + FItemList(i).FIpGoNo - FItemList(i).FChulGoNo - FItemList(i).FsellNo
				FItemList(i).FOCha = rsget("ocha")
				FItemList(i).FjungsanNo = rsget("jungsanno")
				FItemList(i).FDetailidx = rsget("detailidx")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub GetMonthJaegoBojung()
		dim sqlStr,i
		dim mastercode, mastercode2

		sqlStr = "select top 1 code"
 		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master"
		sqlStr = sqlStr + " where left(code,2)='ME'"
		sqlStr = sqlStr + " and convert(varchar(7),executedt,21)='" + FRectPreYYYYMM + "'"
		sqlStr = sqlStr + " and deldt is NULL"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			mastercode = rsget("code")
			FPremastercode = mastercode
		end if
		rsget.close

		'response.write mastercode

		sqlStr = "select top 1 code"
 		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master"
		sqlStr = sqlStr + " where left(code,2)='ME'"
		sqlStr = sqlStr + " and convert(varchar(7),executedt,21)='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and deldt is NULL"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			mastercode2 = rsget("code")
			FCurrmastercode = mastercode2
		end if
		rsget.close


		sqlStr = "select top 2000 T.itemid, T.itemname, T.itemoption, T.itemoptionname,"
		sqlStr = sqlStr + " T.sellcash, T.buycash,"
		sqlStr = sqlStr + " T.isusing,"
		sqlStr = sqlStr + " p1.preidx,p1.premastercode,"
		sqlStr = sqlStr + " IsNull(p1.itemno,0) as prejaego,"
		sqlStr = sqlStr + " p2.curridx,p2.currmastercode,"
		sqlStr = sqlStr + " IsNull(p2.itemno,0) as realjaego,"
		sqlStr = sqlStr + " IsNull(p3.ipgono,0) as ipgono,"
		sqlStr = sqlStr + " IsNull(p4.chulgono,0) as chulgono,"
		sqlStr = sqlStr + " s1.sellno"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " (select i.itemid, i.itemname, IsNull(v.itemoption,'0000') as itemoption, v.opt2name as itemoptionname,"
		sqlStr = sqlStr + " i.sellcash, i.buycash, i.isusing"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].vw_itemoption v on i.itemid=v.itemid"
		sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			'sqlStr = sqlStr + " select d.itemid, d.itemoption, sum(d.itemno) as sellno"
			'sqlStr = sqlStr + " from [db_storage].[dbo].vw_acount_sell_month d,"
			'sqlStr = sqlStr + " [dbo].[dbo].tbl_item i"
			'sqlStr = sqlStr + " where d.executedt='" + FRectYYYYMM + "'"
			'sqlStr = sqlStr + " and d.itemid=i.itemid"
			'sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
			'sqlStr = sqlStr + " group by d.itemid, d.itemoption"
			sqlStr = sqlStr + " select SUM(CASE  WHEN d .itemcost<0 THEN d.itemno * - 1 ELSE d.itemno END)  as sellno, d.itemid,d.itemoption"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + " and m.beadaldate>='" + FRectStartDay + "'"
			sqlStr = sqlStr + " and m.beadaldate<'" + FRectEndDay + "'"
			sqlStr = sqlStr + " and m.ipkumdiv>5"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid=i.itemid"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " group by d.itemid,d.itemoption"
			sqlStr = sqlStr + " ) as s1 on s1.itemid=T.itemid and s1.itemoption=T.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select s.id as preidx , s.mastercode as premastercode ,s.itemid, s.itemoption, s.itemno"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.mastercode='" + mastercode + "'"
			sqlStr = sqlStr + " and s.deldt is NULL"
			sqlStr = sqlStr + " and i.itemid=s.itemid"
			sqlStr = sqlStr + " ) as p1 on p1.itemid=T.itemid and p1.itemoption=T.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select s.id as curridx , s.mastercode as currmastercode,s.itemid, s.itemoption, s.itemno"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.mastercode='" + mastercode2 + "'"
			sqlStr = sqlStr + " and s.deldt is NULL"
			sqlStr = sqlStr + " and i.itemid=s.itemid"
			sqlStr = sqlStr + " ) as p2 on p2.itemid=T.itemid and p2.itemoption=T.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select s.itemid, s.itemoption, sum(s.itemno) as ipgono"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail s, [db_storage].[dbo].tbl_acount_storage_master m,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where m.deldt is Null"
			sqlStr = sqlStr + " and Left(m.code,2)='ST'"
			sqlStr = sqlStr + " and m.executedt>='" + FRectStartDay + "'"
			sqlStr = sqlStr + " and m.executedt<'" + FRectEndDay + "'"
			sqlStr = sqlStr + " and s.mastercode=m.code"
			sqlStr = sqlStr + " and s.deldt is NULL"
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and i.itemid=s.itemid"
			sqlStr = sqlStr + " group by s.itemid, s.itemoption"
			sqlStr = sqlStr + " ) as p3 on p3.itemid=T.itemid and p3.itemoption=T.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select s.itemid, s.itemoption, sum(s.itemno) as chulgono"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail s, [db_storage].[dbo].tbl_acount_storage_master m,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where m.deldt is Null"
			sqlStr = sqlStr + " and ((Left(m.code,2)='SO') or (Left(m.code,2)='SR'))"
			sqlStr = sqlStr + " and m.executedt>='" + FRectStartDay + "'"
			sqlStr = sqlStr + " and m.executedt<'" + FRectEndDay + "'"
			sqlStr = sqlStr + " and s.mastercode=m.code"
			sqlStr = sqlStr + " and s.deldt is NULL"
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and i.itemid=s.itemid"
			sqlStr = sqlStr + " group by s.itemid, s.itemoption"
			sqlStr = sqlStr + " ) as p4 on p4.itemid=T.itemid and p4.itemoption=T.itemoption"
		sqlStr = sqlStr + " order by T.itemid, T.itemoption"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CWitakJungSanItem
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FSellCash      = rsget("sellcash")
				FItemList(i).FSuplycash       = rsget("buycash")
				FItemList(i).FIpGoNo = Null2Zero(rsget("ipgono"))
				FItemList(i).FChulgoNo = Null2Zero(rsget("chulgono")) * -1
				FItemList(i).Fsellno	 = Null2Zero(rsget("sellno"))
				FItemList(i).Fprejaego   = Null2Zero(rsget("prejaego"))
				FItemList(i).Frealjaego   = Null2Zero(rsget("realjaego"))

				FItemList(i).FIsUsing		= rsget("isusing")
				FItemList(i).FIsDelete	    = "N"

				FItemList(i).FsysJaeGo = FItemList(i).FPrejaego + FItemList(i).FIpGoNo - FItemList(i).FChulGoNo - FItemList(i).FsellNo
				FItemList(i).FOCha = FItemList(i).FsysJaeGo - FItemList(i).Frealjaego
				FItemList(i).FjungsanNo = FItemList(i).FChulgoNo + ojungsan.FItemList(i).FsellNo + FItemList(i).FOCha

				if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSellcash_sell=0) then
					FItemList(i).FSellcash_sell	= FItemList(i).FSellCash
				end if

				if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSuplycash_sell=0) then
					FItemList(i).FSuplycash_sell	= FItemList(i).FSuplycash
				end if

				FItemList(i).FPreIdx  = rsget("preidx")
				FItemList(i).FCurrIDx = rsget("curridx")
				FItemList(i).FPreMastercode = rsget("premastercode")
				FItemList(i).FCurrMasterCode = rsget("currmastercode")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub GetWitakJungSanByItem()
		dim sqlStr,i,mastercode,mastercode2

		mastercode =""

		sqlStr = "select top 1 masterid from [db_jungsan].[dbo].tbl_designer_jungsan_witak"
		sqlStr = sqlStr + " where masterid=" + FRectID
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			mastercode = rsget("masterid")
		end if
		rsget.close

		if mastercode<>"" then
			sqlStr = "select top 2000 T.itemid, T.itemname, T.itemoption, T.itemoptionname,"
			sqlStr = sqlStr + " T.sellcash, T.buycash,"
			sqlStr = sqlStr + " T.isusing,"
			sqlStr = sqlStr + " w.ipgono,"
			sqlStr = sqlStr + " w.chulgono,"
			sqlStr = sqlStr + " w.sellcash as sellsellcash,"
			sqlStr = sqlStr + " w.suplycash as sellsuplycash,"
			sqlStr = sqlStr + " w.sellno,"
			sqlStr = sqlStr + " w.prejaego,"
			sqlStr = sqlStr + " w.realjaego,"
			sqlStr = sqlStr + " w.ocha,"
			sqlStr = sqlStr + " w.jungsanno,"
			sqlStr = sqlStr + " w.deleteyn as isdelete,"
			sqlStr = sqlStr + " w.id as detailidx"
			sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " (select i.itemid, i.itemname, IsNull(v.itemoption,'0000') as itemoption, v.optionname as itemoptionname,"
				sqlStr = sqlStr + " i.sellcash, i.buycash, i.isusing"
				sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
				sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid"
				sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
				sqlStr = sqlStr + " ) as T"
			sqlStr = sqlStr + " left join "
				sqlStr = sqlStr + " ( select id,itemid,itemoption,sellcash,suplycash,"
				sqlStr = sqlStr + " prejaego, ipgono, chulgono,"
				sqlStr = sqlStr + " sellno, ocha, realjaego, jungsanno, deleteyn"
				sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_witak"
				sqlStr = sqlStr + " where masterid=" + CStr(mastercode)
				sqlStr = sqlStr + " ) as w on w.itemid=T.itemid and w.itemoption=T.itemoption"
			sqlStr = sqlStr + " order by T.itemid, T.itemoption"

			rsget.Open sqlStr,dbget,1

			FtotalPage =  CInt(FTotalCount\FPageSize)
			if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
				FtotalPage = FtotalPage +1
			end if
			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new CWitakJungSanItem
					FItemList(i).Fitemid        = rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname      = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
					FItemList(i).FSellCash      = rsget("sellcash")
					FItemList(i).FSuplycash       = rsget("buycash")
					FItemList(i).FSellcash_sell = Null2Zero(rsget("sellsellcash"))
					FItemList(i).FSuplycash_sell = Null2Zero(rsget("sellsuplycash"))
					FItemList(i).FIpGoNo = Null2Zero(rsget("ipgono"))
					FItemList(i).FChulgoNo = Null2Zero(rsget("chulgono"))
					FItemList(i).Fsellno	 = Null2Zero(rsget("sellno"))
					FItemList(i).Fprejaego   = Null2Zero(rsget("prejaego"))
					FItemList(i).Frealjaego   = Null2Zero(rsget("realjaego"))

					FItemList(i).FIsUsing		= rsget("isusing")
					FItemList(i).FIsDelete	    = rsget("isdelete")

					FItemList(i).FsysJaeGo = FItemList(i).FPrejaego + FItemList(i).FIpGoNo - FItemList(i).FChulGoNo - FItemList(i).FsellNo
					FItemList(i).FOCha = rsget("ocha")
					FItemList(i).FjungsanNo = rsget("jungsanno")
					FItemList(i).FDetailidx = rsget("detailidx")
					i=i+1
					rsget.moveNext
				loop
			end if

			rsget.Close
			FWitakInsserted = true
			Exit sub
		end if

		sqlStr = "select top 1 code"
 		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master"
		sqlStr = sqlStr + " where left(code,2)='ME'"
		sqlStr = sqlStr + " and convert(varchar(7),executedt,21)='" + FRectPreYYYYMM + "'"
		sqlStr = sqlStr + " and deldt is NULL"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			mastercode = rsget("code")
		end if
		rsget.close

		'response.write sqlStr

		sqlStr = "select top 1 code"
 		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master"
		sqlStr = sqlStr + " where left(code,2)='ME'"
		sqlStr = sqlStr + " and convert(varchar(7),executedt,21)='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and deldt is NULL"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			mastercode2 = rsget("code")
		end if
		rsget.close

		'response.write sqlStr

		sqlStr = "select top 2000 T.itemid, T.itemname, T.itemoption, T.itemoptionname,"
		sqlStr = sqlStr + " T.sellcash, T.buycash,"
		sqlStr = sqlStr + " T.isusing,"
		'sqlStr = sqlStr + " w1.sellcash as wiipsellcash, w1.suplycash as wiipsuplycash,"
		sqlStr = sqlStr + " w1.itemno as ipgono,"
		sqlStr = sqlStr + " w2.itemno as chulgono,"
		sqlStr = sqlStr + " s1.itemcost as sellsellcash,"
		sqlStr = sqlStr + " s1.buycash as sellsuplycash,"
		sqlStr = sqlStr + " s1.sellno,"
		'sqlStr = sqlStr + " p.storageno as prejaego,"
		sqlStr = sqlStr + " IsNull(p2.itemno,0) as prejaego,"
		'sqlStr = sqlStr + " p3.storageno as realjaego,"
		sqlStr = sqlStr + " IsNull(p4.itemno,0) as realjaego,"
		sqlStr = sqlStr + " 'N' as isdelete"
		sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " (select i.itemid, i.itemname, IsNull(v.itemoption,'0000') as itemoption, v.optionname as itemoptionname,"
			sqlStr = sqlStr + " i.sellcash, i.buycash, i.isusing"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid"
			sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " (select itemid,itemoption,sum(itemno) as itemno"
			'sqlStr = sqlStr + ",sellcash,suplycash"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail "
			sqlStr = sqlStr + " where masteridx=" + FRectId
			sqlStr = sqlStr + " and gubuncd='witak'"
			sqlStr = sqlStr + " group by itemid,itemoption) w1 on T.itemid=w1.itemid and T.itemoption=w1.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " (select itemid,itemoption,sum(itemno) as itemno"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail "
			sqlStr = sqlStr + " where masteridx=" + FRectId
			sqlStr = sqlStr + " and gubuncd='witakchulgo'"
			sqlStr = sqlStr + " group by itemid,itemoption) w2 on T.itemid=w2.itemid and T.itemoption=w2.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select SUM(CASE  WHEN d.itemcost<0 THEN d.itemno * - 1 ELSE d.itemno END) as sellno, d.itemid,d.itemoption,d.buycash,Abs(d.itemcost) as itemcost"
			''',d.itemcost,d.buycash"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + " and m.beadaldate>='" + FRectStartDay + "'"
			sqlStr = sqlStr + " and m.beadaldate<'" + FRectEndDay + "'"
			sqlStr = sqlStr + " and m.ipkumdiv>5"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.itemid=i.itemid"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
			sqlStr = sqlStr + " group by d.itemid,d.itemoption,d.buycash,d.itemcost"
			sqlStr = sqlStr + " ) as s1 on s1.itemid=T.itemid and s1.itemoption=T.itemoption"
		'sqlStr = sqlStr + " left join "
		'	sqlStr = sqlStr + " ("
		'	sqlStr = sqlStr + " select s.itemid, s.itemoption, s.storageno"
		'	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_month_storage s,"
		'	sqlStr = sqlStr + " [dbo].[dbo].tbl_item i"
		'	sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
		'	sqlStr = sqlStr + " and s.yyyymm='" + FRectPreYYYYMM + "'"
		'	sqlStr = sqlStr + " and i.itemid=s.itemid"
		'	sqlStr = sqlStr + " ) as p on p.itemid=T.itemid and p.itemoption=T.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select s.itemid, s.itemoption, s.itemno"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.mastercode='" + mastercode + "'"
			sqlStr = sqlStr + " and s.deldt is NULL"
			sqlStr = sqlStr + " and i.itemid=s.itemid"
			sqlStr = sqlStr + " ) as p2 on p2.itemid=T.itemid and p2.itemoption=T.itemoption"
		'sqlStr = sqlStr + " left join "
		'	sqlStr = sqlStr + " ("
		'	sqlStr = sqlStr + " select s.itemid, s.itemoption, s.storageno"
		'	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_month_storage s,"
		'	sqlStr = sqlStr + " [dbo].[dbo].tbl_item i"
		'	sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
		'	sqlStr = sqlStr + " and s.yyyymm='" + FRectYYYYMM + "'"
		'	sqlStr = sqlStr + " and i.itemid=s.itemid"
		'	sqlStr = sqlStr + " ) as p3 on p3.itemid=T.itemid and p3.itemoption=T.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select s.itemid, s.itemoption, s.itemno"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.mastercode='" + mastercode2 + "'"
			sqlStr = sqlStr + " and s.deldt is NULL"
			sqlStr = sqlStr + " and i.itemid=s.itemid"
			sqlStr = sqlStr + " ) as p4 on p4.itemid=T.itemid and p4.itemoption=T.itemoption"
		''sqlStr = sqlStr + " where (w1.itemno>0 or w2.itemno>0 or T.sellno>0)"
		sqlStr = sqlStr + " order by T.itemid, T.itemoption"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CWitakJungSanItem
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FSellCash      = rsget("sellcash")
				FItemList(i).FSuplycash       = rsget("buycash")
				FItemList(i).FSellcash_sell = Null2Zero(rsget("sellsellcash"))
				FItemList(i).FSuplycash_sell = Null2Zero(rsget("sellsuplycash"))
				FItemList(i).FIpGoNo = Null2Zero(rsget("ipgono"))
				FItemList(i).FChulgoNo = Null2Zero(rsget("chulgono")) * -1
				FItemList(i).Fsellno	 = Null2Zero(rsget("sellno"))
				FItemList(i).Fprejaego   = Null2Zero(rsget("prejaego"))
				FItemList(i).Frealjaego   = Null2Zero(rsget("realjaego"))

				FItemList(i).FIsUsing		= rsget("isusing")
				FItemList(i).FIsDelete	    = rsget("isdelete")

    			'if (FItemList(i).FSuplycash_sell=0) or ((FItemList(i).FSuplycash = FItemList(i).FSuplycash_sell) and (FItemList(i).FSellcash = FItemList(i).Fsellcash_sell))then
    			'	FItemList(i).FPrejaego_tmp = FItemList(i).FPrejaego
	    		'	FItemList(i).FIpGoNo_tmp = FItemList(i).FIpGoNo
	    		'	FItemList(i).FChulGoNo_tmp = FItemList(i).FChulGoNo
	    		'	FItemList(i).Frealjaego_tmp =FItemList(i).Frealjaego
    			'else
	    		'	FItemList(i).FPrejaego_tmp = 0
	    		'	FItemList(i).FIpGoNo_tmp = 0
	    		'	FItemList(i).FChulGoNo_tmp = 0
	    		'	FItemList(i).Frealjaego_tmp =0
    			'end if

				'FItemList(i).FsysJaeGo = FItemList(i).FPrejaego_tmp + FItemList(i).FIpGoNo_tmp - FItemList(i).FChulGoNo_tmp - FItemList(i).FsellNo
				'FItemList(i).FOCha = FItemList(i).FsysJaeGo - FItemList(i).Frealjaego_tmp
				'FItemList(i).FjungsanNo = FItemList(i).FChulgoNo_tmp + ojungsan.FItemList(i).FsellNo + FItemList(i).FOCha

				'if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSellcash_sell=0) then
				'	FItemList(i).FSellcash_sell	= FItemList(i).FSellCash
				'end if

				'if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSuplycash_sell=0) then
				'	FItemList(i).FSuplycash_sell	= FItemList(i).FSuplycash
				'end if

				FItemList(i).FsysJaeGo = FItemList(i).FPrejaego + FItemList(i).FIpGoNo - FItemList(i).FChulGoNo - FItemList(i).FsellNo
				FItemList(i).FOCha = FItemList(i).FsysJaeGo - FItemList(i).Frealjaego
				FItemList(i).FjungsanNo = FItemList(i).FChulgoNo + ojungsan.FItemList(i).FsellNo + FItemList(i).FOCha

				if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSellcash_sell=0) then
					FItemList(i).FSellcash_sell	= FItemList(i).FSellCash
				end if

				if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSuplycash_sell=0) then
					FItemList(i).FSuplycash_sell	= FItemList(i).FSuplycash
				end if

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub GetWitakJungSanBySell()
		dim sqlStr,i,mastercode,mastercode2

		mastercode =""

		sqlStr = "select top 1 masterid from [db_jungsan].[dbo].tbl_designer_jungsan_witak"
		sqlStr = sqlStr + " where masterid=" + FRectID
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			mastercode = rsget("masterid")
		end if
		rsget.close

		if mastercode<>"" then
			sqlStr = "select top 2000 T.itemid, T.itemname, T.itemoption, T.itemoptionname,"
			sqlStr = sqlStr + " T.sellcash, T.buycash,"
			sqlStr = sqlStr + " T.isusing,"
			sqlStr = sqlStr + " w.ipgono,"
			sqlStr = sqlStr + " w.chulgono,"
			sqlStr = sqlStr + " w.sellcash as sellsellcash,"
			sqlStr = sqlStr + " w.suplycash as sellsuplycash,"
			sqlStr = sqlStr + " w.sellno,"
			sqlStr = sqlStr + " w.prejaego,"
			sqlStr = sqlStr + " w.realjaego,"
			sqlStr = sqlStr + " w.ocha,"
			sqlStr = sqlStr + " w.jungsanno,"
			sqlStr = sqlStr + " w.deleteyn as isdelete,"
			sqlStr = sqlStr + " w.id as detailidx"
			sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " (select i.itemid, i.itemname, IsNull(v.itemoption,'0000') as itemoption, v.optionname as itemoptionname,"
				sqlStr = sqlStr + " i.sellcash, i.buycash, i.isusing"
				sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
				sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid"
				sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
				sqlStr = sqlStr + " ) as T"
			sqlStr = sqlStr + " left join "
				sqlStr = sqlStr + " ( select id,itemid,itemoption,sellcash,suplycash,"
				sqlStr = sqlStr + " prejaego, ipgono, chulgono,"
				sqlStr = sqlStr + " sellno, ocha, realjaego, jungsanno, deleteyn"
				sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_witak"
				sqlStr = sqlStr + " where masterid=" + CStr(mastercode)
				sqlStr = sqlStr + " ) as w on w.itemid=T.itemid and w.itemoption=T.itemoption"
			sqlStr = sqlStr + " order by T.itemid, T.itemoption"

			rsget.Open sqlStr,dbget,1

			FtotalPage =  CInt(FTotalCount\FPageSize)
			if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
				FtotalPage = FtotalPage +1
			end if
			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new CWitakJungSanItem
					FItemList(i).Fitemid        = rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname      = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
					FItemList(i).FSellCash      = rsget("sellcash")
					FItemList(i).FSuplycash       = rsget("buycash")
					FItemList(i).FSellcash_sell = Null2Zero(rsget("sellsellcash"))
					FItemList(i).FSuplycash_sell = Null2Zero(rsget("sellsuplycash"))
					FItemList(i).FIpGoNo = Null2Zero(rsget("ipgono"))
					FItemList(i).FChulgoNo = Null2Zero(rsget("chulgono"))
					FItemList(i).Fsellno	 = Null2Zero(rsget("sellno"))
					FItemList(i).Fprejaego   = Null2Zero(rsget("prejaego"))
					FItemList(i).Frealjaego   = Null2Zero(rsget("realjaego"))

					FItemList(i).FIsUsing		= rsget("isusing")
					FItemList(i).FIsDelete	    = rsget("isdelete")

					FItemList(i).FsysJaeGo = FItemList(i).FPrejaego + FItemList(i).FIpGoNo - FItemList(i).FChulGoNo - FItemList(i).FsellNo
					FItemList(i).FOCha = rsget("ocha")
					FItemList(i).FjungsanNo = rsget("jungsanno")
					FItemList(i).FDetailidx = rsget("detailidx")
					i=i+1
					rsget.moveNext
				loop
			end if

			rsget.Close
			FWitakInsserted = true
			Exit sub
		end if

		sqlStr = "select top 1 code"
 		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master"
		sqlStr = sqlStr + " where left(code,2)='ME'"
		sqlStr = sqlStr + " and convert(varchar(7),executedt,21)='" + FRectPreYYYYMM + "'"
		sqlStr = sqlStr + " and deldt is NULL"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			mastercode = rsget("code")
		end if
		rsget.close

		'response.write sqlStr

		sqlStr = "select top 1 code"
 		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master"
		sqlStr = sqlStr + " where left(code,2)='ME'"
		sqlStr = sqlStr + " and convert(varchar(7),executedt,21)='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and deldt is NULL"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			mastercode2 = rsget("code")
		end if
		rsget.close

		'response.write sqlStr

		sqlStr = "select top 2000 T.itemid, T.itemname, T.itemoption, T.itemoptionname,"
		sqlStr = sqlStr + " T.sellcash, T.buycash,"
		sqlStr = sqlStr + " T.isusing,"
		'sqlStr = sqlStr + " w1.sellcash as wiipsellcash, w1.suplycash as wiipsuplycash,"
		sqlStr = sqlStr + " w1.itemno as ipgono,"
		sqlStr = sqlStr + " w2.itemno as chulgono,"
		sqlStr = sqlStr + " s1.sellcash as sellsellcash,"
		sqlStr = sqlStr + " s1.suplycash as sellsuplycash,"
		sqlStr = sqlStr + " s1.sellno,"
		'sqlStr = sqlStr + " p.storageno as prejaego,"
		sqlStr = sqlStr + " IsNull(p2.itemno,0) as prejaego,"
		'sqlStr = sqlStr + " p3.storageno as realjaego,"
		sqlStr = sqlStr + " IsNull(p4.itemno,0) as realjaego,"
		sqlStr = sqlStr + " 'N' as isdelete"
		sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " (select i.itemid, i.itemname, IsNull(v.itemoption,'0000') as itemoption, v.optionname as itemoptionname,"
			sqlStr = sqlStr + " i.sellcash, i.buycash, i.isusing"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid"
			sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " (select itemid,itemoption,sum(itemno) as itemno"
			'sqlStr = sqlStr + ",sellcash,suplycash"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail "
			sqlStr = sqlStr + " where masteridx=" + FRectId
			sqlStr = sqlStr + " and gubuncd='witak'"
			sqlStr = sqlStr + " group by itemid,itemoption) w1 on T.itemid=w1.itemid and T.itemoption=w1.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " (select itemid,itemoption,sum(itemno*-1) as itemno"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail "
			sqlStr = sqlStr + " where masteridx=" + FRectId
			sqlStr = sqlStr + " and gubuncd='witakchulgo'"
			sqlStr = sqlStr + " group by itemid,itemoption) w2 on T.itemid=w2.itemid and T.itemoption=w2.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select SUM(CASE  WHEN sellcash<0 THEN itemno * - 1 ELSE itemno END) as sellno,"
			sqlStr = sqlStr + " itemid,itemoption,Abs(suplycash) as suplycash, Abs(sellcash) as sellcash"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail "
			sqlStr = sqlStr + " where masteridx=" + FRectId
			sqlStr = sqlStr + " and gubuncd='witaksell'"
			sqlStr = sqlStr + " group by itemid,itemoption,suplycash,sellcash) s1 on T.itemid=s1.itemid and T.itemoption=s1.itemoption"

		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select s.itemid, s.itemoption, s.itemno"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.mastercode='" + mastercode + "'"
			sqlStr = sqlStr + " and s.deldt is NULL"
			sqlStr = sqlStr + " and i.itemid=s.itemid"
			sqlStr = sqlStr + " ) as p2 on p2.itemid=T.itemid and p2.itemoption=T.itemoption"

		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ("
			sqlStr = sqlStr + " select s.itemid, s.itemoption, s.itemno"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.mastercode='" + mastercode2 + "'"
			sqlStr = sqlStr + " and s.deldt is NULL"
			sqlStr = sqlStr + " and i.itemid=s.itemid"
			sqlStr = sqlStr + " ) as p4 on p4.itemid=T.itemid and p4.itemoption=T.itemoption"
		''sqlStr = sqlStr + " where (w1.itemno>0 or w2.itemno>0 or T.sellno>0)"
		sqlStr = sqlStr + " order by T.itemid, T.itemoption"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CWitakJungSanItem
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FSellCash      = rsget("sellcash")
				FItemList(i).FSuplycash       = rsget("buycash")
				FItemList(i).FSellcash_sell = Null2Zero(rsget("sellsellcash"))
				FItemList(i).FSuplycash_sell = Null2Zero(rsget("sellsuplycash"))
				FItemList(i).FIpGoNo = Null2Zero(rsget("ipgono"))
				FItemList(i).FChulgoNo = Null2Zero(rsget("chulgono")) * -1
				FItemList(i).Fsellno	 = Null2Zero(rsget("sellno"))
				FItemList(i).Fprejaego   = Null2Zero(rsget("prejaego"))
				FItemList(i).Frealjaego   = Null2Zero(rsget("realjaego"))

				FItemList(i).FIsUsing		= rsget("isusing")
				FItemList(i).FIsDelete	    = rsget("isdelete")

    			'if (FItemList(i).FSuplycash_sell=0) or ((FItemList(i).FSuplycash = FItemList(i).FSuplycash_sell) and (FItemList(i).FSellcash = FItemList(i).Fsellcash_sell))then
    			'	FItemList(i).FPrejaego_tmp = FItemList(i).FPrejaego
	    		'	FItemList(i).FIpGoNo_tmp = FItemList(i).FIpGoNo
	    		'	FItemList(i).FChulGoNo_tmp = FItemList(i).FChulGoNo
	    		'	FItemList(i).Frealjaego_tmp =FItemList(i).Frealjaego
    			'else
	    		'	FItemList(i).FPrejaego_tmp = 0
	    		'	FItemList(i).FIpGoNo_tmp = 0
	    		'	FItemList(i).FChulGoNo_tmp = 0
	    		'	FItemList(i).Frealjaego_tmp =0
    			'end if

				'FItemList(i).FsysJaeGo = FItemList(i).FPrejaego_tmp + FItemList(i).FIpGoNo_tmp - FItemList(i).FChulGoNo_tmp - FItemList(i).FsellNo
				'FItemList(i).FOCha = FItemList(i).FsysJaeGo - FItemList(i).Frealjaego_tmp
				'FItemList(i).FjungsanNo = FItemList(i).FChulgoNo_tmp + ojungsan.FItemList(i).FsellNo + FItemList(i).FOCha

				'if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSellcash_sell=0) then
				'	FItemList(i).FSellcash_sell	= FItemList(i).FSellCash
				'end if

				'if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSuplycash_sell=0) then
				'	FItemList(i).FSuplycash_sell	= FItemList(i).FSuplycash
				'end if

				FItemList(i).FsysJaeGo = FItemList(i).FPrejaego + FItemList(i).FIpGoNo - FItemList(i).FChulGoNo - FItemList(i).FsellNo
				FItemList(i).FOCha = FItemList(i).FsysJaeGo - FItemList(i).Frealjaego
				FItemList(i).FjungsanNo = FItemList(i).FChulgoNo + ojungsan.FItemList(i).FsellNo + FItemList(i).FOCha

				if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSellcash_sell=0) then
					FItemList(i).FSellcash_sell	= FItemList(i).FSellCash
				end if

				if (FItemList(i).FjungsanNo<>0) and (FItemList(i).FSuplycash_sell=0) then
					FItemList(i).FSuplycash_sell	= FItemList(i).FSuplycash
				end if

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public function CheckDuplicated(byval idx)
		dim i,cnt
		dim iitemid, iitemoption

		iitemid = FItemList(idx).Fitemid
		iitemoption = FItemList(idx).Fitemoption

		CheckDuplicated = false
		cnt = UBound(FItemList)
		for i=0 to cnt -1
			if (idx<>i) and (iitemid= FItemList(i).Fitemid) and (iitemoption= FItemList(i).Fitemoption) then
				CheckDuplicated = true
				Exit for
			end if
		next
	end function

	public function GetMijungsanList()
		dim sqlStr,i
		sqlStr = " select c.userid, c.socname from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " where c.isusing='Y'"
		sqlStr = sqlStr + " and c.userdiv in ('02','03','04','05','06','07','08')"
		'sqlStr = sqlStr + " left join ("
		'sqlStr = sqlStr + " select "
		'sqlStr = sqlStr + " ) as u on u.userid=c.userid"
		sqlStr = sqlStr + " and c.userid not in ("
		sqlStr = sqlStr + " select designerid from [db_jungsan].[dbo].tbl_designer_jungsan_master "
		sqlStr = sqlStr + " where yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " )"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanMasterItem
				'FItemList(i).Fid
				FItemList(i).Fdesignerid         = rsget("userid")
				FItemList(i).Fyyyymm             = FRectYYYYMM
				FItemList(i).Ftitle              = db2html(rsget("socname"))
				'FItemList(i).Fub_cnt
				'FItemList(i).Fub_totalsellcash
				'FItemList(i).Fub_totalsuplycash
				'FItemList(i).Fub_comment
				'FItemList(i).Fme_cnt
				'FItemList(i).Fme_totalsellcash
				'FItemList(i).Fme_totalsuplycash
				'FItemList(i).Fme_comment
				'FItemList(i).Fwi_cnt
				'FItemList(i).Fwi_totalsellcash
				'FItemList(i).Fwi_totalsuplycash
				'FItemList(i).Fwi_comment
				'FItemList(i).Fregdate
				'FItemList(i).Fcancelyn
				'FItemList(i).Ffinishflag
				'FItemList(i).Fipkumdate

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
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
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>
