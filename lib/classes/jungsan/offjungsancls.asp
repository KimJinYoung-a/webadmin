<%
function getPartnerId2GroupID(ipartnerid)
    dim sqlStr
	sqlStr = "select groupid from db_partner.dbo.tbl_partner where id='"&ipartnerid&"'"

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    getPartnerId2GroupID = rsget("groupid")
    end if
    rsget.Close
end function


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

    public Fjgubun
    public FtaxType
    public Ftot_commission
    public FitemVatyn

    ''기본 정산조건.
    public Fchargediv
    public Fdefaultmargin
    public Fdefaultsuplymargin
    public Fautojungsan
    public Fautojungsandiv

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
    public Fvatinclude

    public function GetBarCode()
        GetBarCode = Fitemgubun + Format00(6,Fitemid) + Fitemoption
        if (FItemID >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,FItemId)) + CStr(Fitemoption)
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

    public FFixsegumil
    public FGroupid

    ''2014/01/27 추가  수수료 매출 관련 =================================================
    public FJgubun
    public Ftotalcommission

    public function IsCommissionTax()  ''수수료 매출 세금 계산서 인지.
        IsCommissionTax = false
        if isNULL(Fjgubun) then Exit function

        IsCommissionTax = (Fjgubun="CC")
    end function

    public function getJGugunName
        if isNULL(Fjgubun) then
            getJGugunName = "매입정산"
        elseif Fjgubun="CC" then
            getJGugunName = "수수료정산"
        elseif Fjgubun="MM" then
            getJGugunName = "매입정산"
        else
            getJGugunName = Fjgubun
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


    ''전자 세금계산서 관련
    public function GetTotalTaxSuply()
		if Ftaxtype="01" then
			GetTotalTaxSuply = CLng(Ftot_jungsanprice / 1.1)
		else
			GetTotalTaxSuply = Ftot_jungsanprice
		end if
	end function

	public function GetTotalTaxVat()
		GetTotalTaxVat = Ftot_jungsanprice - GetTotalTaxSuply
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
		if Not(IsNULL(FFixsegumil)) and (FFixsegumil<>"") then
			GetNormalTaxDate = FFixsegumil
		else
		    GetNormalTaxDate = dateserial(Left(Fyyyymm,4),Right(Fyyyymm,2)+1,1-1) ''': 정산월 말일
		end if
	end function

	public function GetPreFixSegumil()
		dim thisdate, maytaxdate
		dim ithis1day , ithis21day, premonth1day, premonth21day

		thisdate = getDbDate()
		maytaxdate = GetNormalTaxDate()

        '' 12일까지 마감할 경우 13으로 세팅
		premonth1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"01")
		premonth21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"13")
		ithis1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"01")
		ithis21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"13")

        ''(매달 12일 까지 발행시 : 정산월 말일)<br>
		''(매달 13일 이후 발행 : 발행월 1일)<br>
		''(이월 내역발행시 12일까지 발행: 발행전월 1일)<br>
		''(이월 내역발행시 13일 이후 발행: 발행월 1일)
		''그외 : 발행일=Today


		if (CStr(FYYYYMM) = Left(CStr(premonth1day),7)) then
		''정상 발행의 경우
		    if (thisdate>=ithis21day) then
		    ''13일 이후 발행건은 이월됨 클릭일 1일
		        GetPreFixSegumil = ithis1day
		    else
		        GetPreFixSegumil = maytaxdate
		    end if
		elseif (CStr(FYYYYMM) < Left(CStr(premonth1day),7)) then
		''이월 발행의 경우
		    if (thisdate>=ithis21day) then
		    ''13일 이후 발행건은 이월됨 클릭일 1일
		        GetPreFixSegumil = ithis1day
		    else
		        GetPreFixSegumil = premonth1day
		    end if
		else
		    GetPreFixSegumil = Left(CStr(thisdate),10)
		end if
	end function

	''==========================================================



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

    '' FRectStartYYYYMM<= RECT <=FRectEndYYYYMM
    public FRectStartYYYYMM
    public FRectEndYYYYMM

    '' FRectStartYYYYMMDD<= RECT <FRectEndYYYYMMDD
    public FRectStartYYYYMMDD
    public FRectEndYYYYMMDD

    public FRectFixStateExiste

    public FRectNotIncludeWonChon
    public FRectOnlyIncludeWonChon
    public FRectNotYYYYMM
    public FRectTaxRegDate

    public function JungsanFixedList()
		dim sqlStr,i
		sqlStr = "select m.*, "
		sqlStr = sqlStr + " p.jungsan_email,p.jungsan_bank,p.jungsan_date_off,p.jungsan_acctno,"
		sqlStr = sqlStr + " p.jungsan_acctname,p.company_name, p.jungsan_gubun,p.company_no"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"

		if FRectfinishflag="ALL" then
		    sqlStr = sqlStr + " where m.finishflag>=3"
		elseif FRectfinishflag<>"" then
		    sqlStr = sqlStr + " where m.finishflag='" + FRectfinishflag + "'"
		else
		    sqlStr = sqlStr + " where m.finishflag='3'"
        end if

        if (FRectTaxRegDate<>"") then
            sqlStr = sqlStr + " and m.taxregdate='" + FRectTaxRegDate + "'"
        end if

        '' AA 전월 정산내역 중 발행일이 전월 & 정산일 수시/15일
        '' BB 전월 정산내역 중 발행일이 전월 & 정산일 말일
        '' CC 전전월 이하 정산내역 중 발행일이 전월
        '' DD 발행일이 현재월 이상
        '' EE 정상발행 전체
        '' FF 이월발행 전체 (비정상발행)
        '' ZZ 발행일이 빈값이거나, 그 외 날짜
        if FRectGubunCd="ZZ" then
            sqlStr = sqlStr + " and m.taxregdate is NULL"
        elseif FRectGubunCd="AA" then
            sqlStr = sqlStr + " and (IsNULL(p.jungsan_date_off,'')='' or p.jungsan_date_off<>'말일')"
            sqlStr = sqlStr + " and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubunCd="BB" then
            sqlStr = sqlStr + " and p.jungsan_date_off='말일'"
            sqlStr = sqlStr + " and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubunCd="CC" then
            sqlStr = sqlStr + " and m.yyyymm<convert(varchar(7),m.taxregdate,21)"
            sqlStr = sqlStr + " and convert(varchar(7),getdate(),21)>convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubunCd="DD" then
            sqlStr = sqlStr + " and convert(varchar(7),getdate(),21)<=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubunCd="EE" then
            sqlStr = sqlStr + " and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubunCd="FF" then
            sqlStr = sqlStr + " and m.yyyymm<>convert(varchar(7),m.taxregdate,21)"
        end if

        if FRectJungsanDate="NULL" then
            sqlStr = sqlStr + " and IsNULL(p.jungsan_date_off,'')=''"
        elseif FRectJungsanDate<>"" then
            sqlStr = sqlStr + " and p.jungsan_date_off='" + FRectJungsanDate + "'"
        end if

        if FRectNotIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun<>'원천징수'"
			sqlStr = sqlStr + " and p.jungsan_gubun<>'간이과세'"
		end if

		if FRectOnlyIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='원천징수'"
		end if

		if FRectbankingupflag<>"" then
		    sqlStr = sqlStr + " and m.bankingupflag='" + FRectBankingupflag + "'"
		end if

		if FRectYYYYMM<>"" then
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectNotYYYYMM<>"" then
			sqlStr = sqlStr + " and m.yyyymm<>'" + FRectNotYYYYMM + "'"
		end if

        sqlStr = sqlStr + " order by m.neotaxno, m.taxinputdate"

		rsget.Open sqlStr,dbget,1

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
        'sqlStr = sqlStr + " sum(case when (m.yyyymm>convert(varchar(7),m.taxregdate,21))  then tot_jungsanprice else 0 end) as next_jungsansum," + VbCrlf
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
            'FItemList(i).Fnext_jungsansum   = rsget("next_jungsansum")

            FItemList(i).Ffixedsum          = rsget("fixedsum")
            FItemList(i).Fipkumsum          = rsget("ipkumsum")

            FItemList(i).Ftot_jungsanprice  = rsget("tot_jungsanprice")


			rsget.MoveNext
			i = i + 1
		loop

	    end if

        rsget.Close
    end Sub

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
        ''정산일 기준으로 입금예정금액 산출.
        ''sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (m.yyyymm=convert(varchar(7),taxregdate,21))  then tot_jungsanprice else 0 end) as fixedthissum," + VbCrlf
        ''sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (m.yyyymm<>convert(varchar(7),taxregdate,21))  then tot_jungsanprice else 0 end) as fixednextsum," + VbCrlf
        ''금월 기준으로 입금예정금액 산출.
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (convert(varchar(7),getdate(),21)>convert(varchar(7),taxregdate,21))  then tot_jungsanprice else 0 end) as fixedthissum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (convert(varchar(7),getdate(),21)<=convert(varchar(7),taxregdate,21))  then tot_jungsanprice else 0 end) as fixednextsum," + VbCrlf

        sqlStr = sqlStr + " sum(case when (m.finishflag <'3') then tot_jungsanprice else 0 end) as waitsum," + VbCrlf
        sqlStr = sqlStr + " sum(tot_jungsanprice) as tot_jungsanprice " + VbCrlf
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m" + VbCrlf
        sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner_group g " + VbCrlf
        sqlStr = sqlStr + " 	on m.groupid=g.groupid" + VbCrlf
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
        sqlStr = sqlStr + " p.jungsan_email,p.jungsan_bank,p.jungsan_date_off,p.jungsan_acctno,"
		sqlStr = sqlStr + " p.jungsan_acctname,p.company_name, p.jungsan_gubun,p.company_no "
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
        sqlStr = sqlStr + " where m.idx=" + CStr(FRectIdx)
        if FRectMakerid<>"" then
            sqlStr = sqlStr + " and m.makerid='" + FRectMakerid + "'"
        end if
        if (FRectGroupid<>"") then
            sqlStr = sqlStr + " and m.groupid='" + FRectGroupid + "'"
        end if

        rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        FTotalCount = FResultCount

        if FResultCount<1 then FResultCount=0

		if  not rsget.EOF  then
			set FOneItem = new COffJungsanMasterItem

			FOneItem.Fidx               = rsget("idx")
            FOneItem.Fyyyymm            = rsget("yyyymm")
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
            FOneItem.Fjgubun            = rsget("jgubun")
            FOneItem.Ftotalcommission   = rsget("totalcommission")
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
        sqlStr = sqlStr + "     left join ("
        sqlStr = sqlStr + "         select distinct makerid, autojungsan "
        sqlStr = sqlStr + "         from [db_shop].[dbo].tbl_shop_designer "
        sqlStr = sqlStr + "         where autojungsan='N' "
        sqlStr = sqlStr + "     ) as T "
        sqlStr = sqlStr + "     on m.makerid=T.makerid"

        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p "
        sqlStr = sqlStr + "     on m.groupid=p.groupid"
        sqlStr = sqlStr + " where 1=1"

        if FRectIdx<>"" then
            sqlStr = sqlStr + " and m.idx=" + CStr(FRectIdx)
        end if

        if FRectMakerid<>"" then
            sqlStr = sqlStr + " and m.makerid='" + FRectMakerid + "'"
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

            if FRectAutojungsan<>"" then
                sqlStr = sqlStr + " and IsNULL(T.autojungsan,'Y')='" + FRectAutojungsan + "'"
            end if

            if FRectJungsanDate<>"" then
                if FRectJungsanDate="NULL" then
                    sqlStr = sqlStr + " and p.jungsan_date_off is NULL"
                else
                    sqlStr = sqlStr + " and p.jungsan_date_off='" + FRectJungsanDate + "'"
                end if
            end if
        end if

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalSum   = rsget("totsum")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.* "
        sqlStr = sqlStr + " , IsNULL(T.autojungsan,'Y') as autojungsan,"
        sqlStr = sqlStr + " p.jungsan_email,p.jungsan_bank,p.jungsan_date_off,p.jungsan_acctno,"
		sqlStr = sqlStr + " p.jungsan_acctname,p.company_name, p.jungsan_gubun,p.company_no "
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_master m"
        sqlStr = sqlStr + "     left join ("
        sqlStr = sqlStr + "         select distinct makerid, autojungsan "
        sqlStr = sqlStr + "         from [db_shop].[dbo].tbl_shop_designer "
        sqlStr = sqlStr + "         where autojungsan='N' "
        sqlStr = sqlStr + "     ) as T "
        sqlStr = sqlStr + "     on m.makerid=T.makerid"

        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p "
        sqlStr = sqlStr + "     on m.groupid=p.groupid"
        sqlStr = sqlStr + " where 1=1"

        if FRectIdx<>"" then
            sqlStr = sqlStr + " and m.idx=" + CStr(FRectIdx)
        end if

        if FRectMakerid<>"" then
            sqlStr = sqlStr + " and m.makerid='" + FRectMakerid + "'"
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

            if FRectAutojungsan<>"" then
                sqlStr = sqlStr + " and IsNULL(T.autojungsan,'Y')='" + FRectAutojungsan + "'"
            end if

            if FRectJungsanDate<>"" then
                if FRectJungsanDate="NULL" then
                    sqlStr = sqlStr + " and p.jungsan_date_off is NULL"
                else
                    sqlStr = sqlStr + " and p.jungsan_date_off='" + FRectJungsanDate + "'"
                end if
            end if
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

            '' 기본정산조건.
            FOneItem.Fchargediv         = rsget("chargediv")
            FOneItem.Fdefaultmargin     = rsget("defaultmargin")
            FOneItem.Fdefaultsuplymargin= rsget("defaultsuplymargin")
            FOneItem.Fautojungsan       = rsget("autojungsan")
            FOneItem.Fautojungsandiv    = rsget("autojungsandiv")

		end if
        rsget.Close
    end Sub

    public Sub GetOffJungsanDetailSummaryList()
        dim sqlStr, i
        sqlStr = "select T.*, "
        sqlStr = sqlStr + " c.comm_name, u.shopname, " + VbCrlf
        sqlStr = sqlStr + " s.chargediv, s.defaultmargin, s.defaultsuplymargin, s.autojungsan, s.autojungsandiv" + VbCrlf
        sqlStr = sqlStr + " from (" + VbCrlf
        sqlStr = sqlStr + " select m.jgubun, m.taxtype, d.shopid, d.gubuncd, d.vatyn " + VbCrlf
        sqlStr = sqlStr + " ,sum(d.itemno) as tot_itemno " + VbCrlf
        sqlStr = sqlStr + " ,sum(d.sellprice*d.itemno) as tot_orgsellprice " + VbCrlf
        sqlStr = sqlStr + " ,sum(d.realsellprice*d.itemno) as tot_realsellprice " + VbCrlf
        sqlStr = sqlStr + " ,sum(d.suplyprice*d.itemno) as tot_jungsanprice " + VbCrlf
        sqlStr = sqlStr + " ,sum(isNULL(d.commission,0)*d.itemno) as tot_commission " + VbCrlf
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_detail d" + VbCrlf
        sqlStr = sqlStr + "     Join [db_jungsan].[dbo].tbl_off_jungsan_master m"
        sqlStr = sqlStr + "     on d.masteridx=m.idx"
        sqlStr = sqlStr + " where masteridx=" + CStr(FRectIdx)
        sqlStr = sqlStr + " group by m.jgubun, m.taxtype, d.shopid, d.gubuncd, d.vatyn "
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

                FItemList(i).Fjgubun           = rsget("jgubun")
                FItemList(i).FtaxType           = rsget("taxType")
                FItemList(i).Ftot_commission    = rsget("tot_commission")
                FItemList(i).FitemVatyn         = rsget("vatyn")

                '' 기본정산조건.
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


    public Sub GetOffJungsanDetailSumList()
        dim sqlStr, i

        sqlStr = "select Top " + CStr(FPageSize*FCurrPage) + " d.itemgubun, d.itemid, d.itemoption, itemname, itemoptionname, realsellprice, suplyprice ,sum(itemno)  as itemno"
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_off_jungsan_detail d" + VbCrlf
        sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectIdx)
        if (FRectGubunCd<>"") then
            sqlStr = sqlStr + " and d.gubuncd='" + FRectGubunCd + "'"
        end if
        if (FRectShopid<>"") then
            sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
        end if
        sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption, itemname, itemoptionname, realsellprice, suplyprice"
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
                FItemList(i).Frealsellprice = rsget("realsellprice")
                FItemList(i).Fsuplyprice    = rsget("suplyprice")
                FItemList(i).Fitemno        = rsget("itemno")

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
                FItemList(i).Fvatinclude    = rsget("vatinclude")

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