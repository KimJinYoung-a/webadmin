<%


function drawSelectBoxCSCommCombo(selectBoxName,selectedId,groupCode,onChangefunction)
   dim tmp_str,sqlStr
   %>
     <select class="select" name="<%=selectBoxName%>" <%= onChangefunction %> >
     <option value='' <%if selectedId="" then response.write " selected" %> >선택</option>
   <%
       sqlStr = " select comm_cd,comm_name "
       sqlStr = sqlStr + " from  "
       sqlStr = sqlStr + " " & TABLE_CS_COMMON_CODE  & " "
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

function drawSelectBoxCancelTypeBox(selectBoxName,selectedId,orgPaymethod,divcd,onChangefunction)
    dim BufStr, selectStr
    BufStr = "<select class='select' name='returnmethod' " + onChangefunction + ">"
    BufStr = BufStr + "<option value=''>선택</option>"

    if (selectedId="R000") then selectStr="selected"
        BufStr = BufStr + "<option value='R000' " + selectStr + ">환불 없음</option>"
    selectStr = ""

    if (orgPaymethod="100") or (orgPaymethod="110") then
        if (selectedId="R100") then selectStr="selected"
        BufStr = BufStr + "<option value='R100' " + selectStr + ">신용카드 취소</option>"

		if True or application("Svr_Info") = "Dev" then
			if (orgPaymethod = "100") then
				selectStr = ""
				if (selectedId="R120") then selectStr="selected"
		        BufStr = BufStr + "<option value='R120' " + selectStr + ">신용카드 부분취소</option>"
		    end if
        end if
    elseif (orgPaymethod="20")  then
        if (selectedId="R020") then selectStr="selected"
        BufStr = BufStr + "<option value='R020' " + selectStr + ">실시간이체 취소</option>"

		if True or application("Svr_Info") = "Dev" then
			if (orgPaymethod = "20") then
				selectStr = ""
				if (selectedId="R022") then selectStr="selected"
		        BufStr = BufStr + "<option value='R022' " + selectStr + ">실시간이체 부분취소</option>"
		    end if
        end if
    elseif (orgPaymethod="400")  then
        if (selectedId="R400") then selectStr="selected"
        BufStr = BufStr + "<option value='R400' " + selectStr + ">휴대폰결제 취소</option>"
    elseif (orgPaymethod="80")  then
        if (selectedId="R080") then selectStr="selected"
        BufStr = BufStr + "<option value='R080' " + selectStr + ">All@카드 취소</option>"
    elseif (orgPaymethod="50") then
        if (selectedId="R050") then selectStr="selected"
        BufStr = BufStr + "<option value='R050' " + selectStr + ">입점몰결제 취소</option>"
    end if

    selectStr = ""

    if (selectedId="R007") then selectStr="selected"
    BufStr = BufStr + "<option value='R007' " + selectStr + ">무통장 환불</option>"

    selectStr = ""

    if (selectedId="R900") then selectStr="selected"
    BufStr = BufStr + "<option value='R900' " + selectStr + ">마일리지 환급</option>"
    BufStr = BufStr + "</select>"

    response.write BufStr
end function


''취소 프로세스
public function fnIsCancelProcess(idivcd)
    fnIsCancelProcess = (idivcd = "A008")
end function

''반품 프로세스(회수, 맞교환 회수)
public function fnIsReturnProcess(idivcd)
    fnIsReturnProcess = (idivcd = "A004") or (idivcd = "A010") or (idivcd = "A011")
end function

public function fnIsRefundProcess(idivcd)
    fnIsRefundProcess = (idivcd = "A003") or (idivcd = "A005")
end function

''누락발송, 서비스발송  프로세스
public function fnIsServiceDeliverProcess(idivcd)
    fnIsServiceDeliverProcess = (idivcd = "A000") or (idivcd = "A001") or (idivcd = "A002")
end function

''Cs Detail 관련 정보
Class CCSASDetailItem
    ''tbl_as_detail's
    public Fid
    public Fmasterid
    public Fgubun01
    public Fgubun02
    public Fgubun01name
    public Fgubun02name
    public Fregdetailstate
    public Fregitemno
    public Fconfirmitemno
    public Fcausediv
    public Fcausedetail
    public Fcausecontent

    ''tbl_order_detail's
    public Forderdetailidx
    public Forderserial
    public Fitemid
    public Fitemoption
    public Fmakerid
    public Fitemname
    public Fitemoptionname
    public Fitemcost
    public Fbuycash
    public Fitemno
	public Fprevreturnno
    public Fisupchebeasong
    public Fcancelyn

    public Foitemdiv
    public FodlvType
    public Fissailitem
    public Fitemcouponidx
    public Fbonuscouponidx

    public ForderDetailcurrstate
    public FdiscountAssingedCost    '' 주문시 할인된가격 ( ALL@ / %할인권 반영)

    ''public FAllAtDiscountedPrice

    ''tbl_item's
    public FSmallImage

    ''업체 개별배송 상품 배송비 인지 여부
    public function IsUpcheParticleDeliverPayCodeItem
        IsUpcheParticleDeliverPayCodeItem = (Fitemid=0) and (Left(Fitemoption,2)="90")
    end function

    ''업체 개별배송 상품인지 여부
    public function IsUpcheParticleDeliverItem
        IsUpcheParticleDeliverItem = (FodlvType=9)
    end function

    ''반품시 사용하는 상품가격(All@ 할인값, %쿠폰 할인값 반영)
    public function GetOrgPayedItemPrice()
        GetOrgPayedItemPrice = Fitemcost

        if (FdiscountAssingedCost=0) then
            ''기존방식
            GetOrgPayedItemPrice = Fitemcost-getAllAtDiscountedPrice
        else
            if (FdiscountAssingedCost<>Fitemcost) then
                GetOrgPayedItemPrice = FdiscountAssingedCost
            end if
        end if
    end function

    ''All@ 할인된가격
    public function getAllAtDiscountedPrice()
        getAllAtDiscountedPrice =0
        ''기존 상품쿠폰 할인되는경우 추가할인없음.
        ''마일리지샵 상품 추가 할인 없음.
	    ''세일상품 추가할인 없음
	    '' 20070901추가 : 정율할인 보너스쿠폰사용시 추가할인 없음.

'	    if (FdiscountAssingedCost=0) then
'	        ''기존방식
'            if (Fitemcouponidx<>0) or (IsMileShopSangpum) or (Fissailitem="Y") then
'    			getAllAtDiscountedPrice = 0
'    		else
'    			getAllAtDiscountedPrice = round(((1-0.94) * FItemCost / 100) * 100 ) * FItemNo
'    		end if
'    	else

'''			'일단 뺀다.
'''    	    if (IsNULL(Fbonuscouponidx) or (Fbonuscouponidx=0)) and (Fitemcost>FdiscountAssingedCost) then
'''    	            getAllAtDiscountedPrice = Fitemcost-FdiscountAssingedCost
'''    	    else
'''    	        getAllAtDiscountedPrice = 0
'''    	    end if

'    	end if
    end function

    '' %할인권 할인금액 or 카드 할인금액
    public function getPercentBonusCouponDiscountedPrice()
        getPercentBonusCouponDiscountedPrice = 0
'        if (Fitemcost>FdiscountAssingedCost) then
'                getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
'        end if

        if (FdiscountAssingedCost=0) then
	        ''기존방식
	        ''getPercentBonusCouponDiscountedPrice = Fitemcost*
	    else
            if (Fbonuscouponidx<>0)  and (Fitemcost>FdiscountAssingedCost) then
                getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
            end if
        end if
    end function

    ''마일리지샵 상품
    public function IsMileShopSangpum()
		IsMileShopSangpum = false

		if Foitemdiv="82" then
			IsMileShopSangpum = true
		end if
	end function

    public function GetDefaultRegNo(IsRegState)
        if (IsRegState) then
            GetDefaultRegNo = Fitemno
        else
            GetDefaultRegNo = Fregitemno
        end if
    end function

    ''CsAction 접수시 상품 갯수 수정 가능여부
    public function IsItemNoEditEnabled(byval idivcd)
        IsItemNoEditEnabled = false

        if (Fcancelyn="Y") then Exit function

        if (fnIsCancelProcess(idivcd)) then
            IsItemNoEditEnabled = true

            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=false
        elseif (fnIsReturnProcess(idivcd)) then
            ''반품 접수
            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=true
        elseif (fnIsServiceDeliverProcess(idivcd)) then
            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=true

        else

        end if
    end function


    ''CsAction 접수시 상품별 체크 가능여부
    public function IsCheckAvailItem(byval iIpkumdiv, byval iMasterCancelYn, byval idivcd)
        IsCheckAvailItem = false

        if (Fcancelyn="Y") then Exit function
        if (iMasterCancelYn<>"N") then Exit function

        if (fnIsCancelProcess(idivcd)) then
            IsCheckAvailItem = true
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=false

        elseif (fnIsReturnProcess(idivcd)) then
            ''반품 접수
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=true

            if (FItemId=0) then IsCheckAvailItem=true
        elseif (idivcd="A006") then
            ''출고시 유의사항
            IsCheckAvailItem=true

            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=false
        elseif (idivcd="A009") then
            ''기타사항(메모) - All case Avail
            IsCheckAvailItem=true
        elseif (idivcd="A700") then
            ''기타정산 - All case Avail
            IsCheckAvailItem=true
        elseif (idivcd = "A002") then
            if Fitemid=0 then
                IsCheckAvailItem=false
            else
                IsCheckAvailItem=true
            end if
        elseif (idivcd = "A001") then
            ''누락, 서비스
            if (ForderDetailcurrstate>=7) or ((Fcancelyn="A") and (iIpkumdiv>=7)) then IsCheckAvailItem=true
        elseif (idivcd = "A000") then
            ''맞교환
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=true
        else

        end if
    end function

    ''CsAction 접수시 상품별 디폴트 체크드
    public function IsDefaultCheckedItem(byval iIpkumdiv, byval iMasterCancelYn, byval idivcd, byval ckAll)
        IsDefaultCheckedItem =false

        if (Not IsCheckAvailItem(iIpkumdiv,iMasterCancelYn,idivcd)) then Exit function

        if (fnIsCancelProcess(idivcd)) then
            if (ckAll<>"") then
                IsDefaultCheckedItem = true
            else
                IsDefaultCheckedItem = false
            end if

            if (Fcancelyn="Y") or (iMasterCancelYn<>"N") then IsDefaultCheckedItem=false

            if (ForderDetailcurrstate>=3) then IsDefaultCheckedItem=false
        elseif (fnIsReturnProcess(idivcd)) then
            ''반품접수인경우 - No action
        elseif (idivcd="A006") then
            ''출고시 유의사항 - No action
        elseif (idivcd="A009") then
            ''기타사항(메모) - No action
        else

        end if
    end function


    public function CancelStateStr()
		CancelStateStr = "정상"

		if Fcancelyn="Y" then
			CancelStateStr ="취소"
		elseif Fcancelyn="D" then
			CancelStateStr ="삭제"
		elseif Fcancelyn="A" then
			CancelStateStr ="추가"
		end if
	end function

	public function CancelStateColor()
		CancelStateColor = "#000000"

		if Fcancelyn="Y" then
			CancelStateColor ="#FF0000"
		elseif Fcancelyn="D" then
			CancelStateColor ="#FF0000"
		elseif Fcancelyn="A" then
			CancelStateColor ="#0000FF"
		end if
	end function

	''order Detail's State Name : 현상태
	Public function GetStateName()
        if ForderDetailcurrstate="2" then
            if (Fisupchebeasong="Y") then
		        GetStateName = "업체통보"
		    else
		        GetStateName = "물류통보"
		    end if
	    elseif ForderDetailcurrstate="3" then
		    GetStateName = "상품준비"
	    elseif ForderDetailcurrstate="7" then
		    GetStateName = "출고완료"
	    else
		    GetStateName = ForderDetailcurrstate
	    end if
	end Function

	'' 등록시 상태..
	Public function GetRegDetailStateName()
        if (Fregdetailstate="2") then
            if (Fisupchebeasong="Y") then
		        GetRegDetailStateName = "업체통보"
		    else
		        GetRegDetailStateName = "물류통보"
		    end if
	    elseif Fregdetailstate="3" then
		    GetRegDetailStateName = "상품준비"
	    elseif Fregdetailstate="7" then
		    GetRegDetailStateName = "출고완료"
	    else
		    GetRegDetailStateName = "----"
	    end if
	end Function

	''order Detail's State color
	public function GetStateColor()
	    if ForderDetailcurrstate="2" then
			GetStateColor="#000000"
		elseif ForderDetailcurrstate="3" then
			GetStateColor="#CC9933"
		elseif ForderDetailcurrstate="7" then
			GetStateColor="#FF0000"
		else
			GetStateColor="#000000"
		end if
	end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class


''환불 관련 정보
Class CCSASRefundInfoItem
    public Fasid

    public Forgsubtotalprice    ''원 주문 결제액
    public Forgitemcostsum      ''원 주문 상품합계
    public Forgbeasongpay       ''원 주문 배송료
    public Forgmileagesum       ''원 주문 사용마일리지
    public Forgcouponsum        ''원 주문 사용쿠폰
    public Forgallatdiscountsum ''원 주문 올엣할인

    public Frefundrequire       ''환불요청액
    public Frefundresult        ''환불  금액
    public Freturnmethod        ''환불  방식

    public Frefundmileagesum    ''취소  마일리지 Frefundmileagesum
    public Frefundcouponsum     ''취소  쿠폰     Frefundcouponsum
    public Fallatsubtractsum    ''취소  카드할인 Fallatsubtractsum

    public Frefunditemcostsum   ''취소 상품합계
    public Frefundbeasongpay    ''취소시 배송비 차감액
    public Frefunddeliverypay   ''취소시 회수 배송비? -> Freturndeliverypay
    public Frefundadjustpay     ''취소시 기타 보정액
    public Fcanceltotal         ''총 취소액

    public Frebankname          ''환불 은행
    public Frebankaccount       ''환불 계좌
    public Frebankownername     ''예금 주
    public FpaygateTid          ''Pg사 T id

    public FencMethod           ''암호화방식
    public FencAccount          ''암호화 계좌번호
    public FdecAccount          ''복호화 계좌번호

    public FpaygateresultTid
    public FpaygateresultMsg

    public FreturnmethodName    ''환불방식명

    public rebankCode

    public Fupfiledate          ''환불파일 작성일

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
End Class

''고객 회수, 맞교환.. 주소지정보
Class CCSDeliveryItem
    public Fasid
    public Freqname
    public Freqphone
    public Freqhp
    public Freqzipcode
    public Freqzipaddr
    public Freqetcaddr
    public Freqetcstr
    public Fsongjangdiv
    public Fsongjangno
    public Fregdate
    public Fsenddate


    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CCSReturnAddressItem
    public Fbrandid
    public Fbrandname

    public Fstreetname_kor
    public Fstreetname_eng

    public FreturnName
    public FreturnPhone
    public Freturnhp
    public FreturnEmail

    public FreturnZipcode
    public FreturnZipaddr
    public FreturnEtcaddr

    public Fsongjangdiv	'택배사
    public Fsongjangno

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

''반품 주소지 정보
Class CCSReturnAddress
	public FItemList()

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public Fbrandid
    public Fbrandname

    public Fstreetname_kor
    public Fstreetname_eng

    public FreturnName
    public FreturnPhone
    public Freturnhp
    public FreturnEmail

    public FreturnZipcode
    public FreturnZipaddr
    public FreturnEtcaddr

    public Fsongjangdiv
    public Fsongjangno

    public FRectMakerid
    public FRectGroupCode

    public sub GetReturnAddress()
        dim sqlStr
        sqlStr = " select company_name, deliver_phone, deliver_hp, return_zipcode, return_address, return_address2"
        sqlStr = sqlStr + " from " & TABLE_PARTNER & ""
        sqlStr = sqlStr + " where id='" + FRectMakerid + "'"

        rsget.Open sqlStr, dbget, 1

        if Not rsget.Eof then
            FreturnName      = db2html(rsget("company_name"))
            FreturnPhone     = db2html(rsget("deliver_phone"))
            Freturnhp        = db2html(rsget("deliver_hp"))
            FreturnZipcode   = rsget("return_zipcode")
            FreturnZipaddr   = db2html(rsget("return_address"))
            FreturnEtcaddr   = db2html(rsget("return_address2"))
            Fsongjangdiv     = ""
            Fsongjangno      = ""

        end if
        rsget.Close
    end sub

    public sub GetBrandReturnAddress()
    	'GetReturnAddress() 에서 company_name 를 FreturnName 에 세팅하므로 별도 함수 생성
        dim sqlStr
        sqlStr = " select id as brandid, company_name as brandname, socname_kor as streetname_kor, socname as streetname_eng, return_zipcode, return_address, return_address2, deliver_phone, deliver_hp, deliver_name, deliver_email, defaultsongjangdiv "
        sqlStr = sqlStr + " from " & TABLE_PARTNER & " p, " & TABLE_USER_C & " c "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and p.id = c.userid "
        sqlStr = sqlStr + " and p.id='" + FRectMakerid + "'"

        rsget.Open sqlStr, dbget, 1

        if Not rsget.Eof then

			Fbrandid         = rsget("brandid")
			Fbrandname       = db2html(rsget("brandname"))

			Fstreetname_kor  = db2html(rsget("streetname_kor"))
			Fstreetname_eng  = db2html(rsget("streetname_eng"))

			FreturnName      = rsget("deliver_name")
			FreturnPhone     = rsget("deliver_phone")
			Freturnhp        = rsget("deliver_hp")
			FreturnEmail     = rsget("deliver_email")

            FreturnZipcode   = rsget("return_zipcode")
            FreturnZipaddr   = db2html(rsget("return_address"))
            FreturnEtcaddr   = db2html(rsget("return_address2"))

            Fsongjangdiv     = rsget("defaultsongjangdiv")

        end if
        rsget.Close
    end sub

    public sub GetReturnAddressList()
        dim sqlStr, i

        sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from " & TABLE_PARTNER & " p, " & TABLE_USER_C & " c "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and p.id = c.userid "
        sqlStr = sqlStr + " and p.groupid ='" + FRectGroupCode + "'"

        rsget.Open sqlStr, dbget, 1
            FTotalCount = rsget("cnt")
        rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr + " id as brandid, company_name as brandname, socname_kor as streetname_kor, socname as streetname_eng, return_zipcode, return_address, return_address2, deliver_phone, deliver_hp, deliver_name, deliver_email, defaultsongjangdiv "
        sqlStr = sqlStr + " from " & TABLE_PARTNER & " p, " & TABLE_USER_C & " c "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and p.id = c.userid "
        sqlStr = sqlStr + " and p.groupid ='" + FRectGroupCode + "'"
        sqlStr = sqlStr + " order by id "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSReturnAddressItem

				FItemList(i).Fbrandid         = rsget("brandid")
				FItemList(i).Fbrandname       = db2html(rsget("brandname"))

				FItemList(i).Fstreetname_kor  = db2html(rsget("streetname_kor"))
				FItemList(i).Fstreetname_eng  = db2html(rsget("streetname_eng"))

				FItemList(i).FreturnName      = rsget("deliver_name")
				FItemList(i).FreturnPhone     = rsget("deliver_phone")
				FItemList(i).Freturnhp        = rsget("deliver_hp")
				FItemList(i).FreturnEmail     = rsget("deliver_email")

	            FItemList(i).FreturnZipcode   = rsget("return_zipcode")
	            FItemList(i).FreturnZipaddr   = db2html(rsget("return_address"))
	            FItemList(i).FreturnEtcaddr   = db2html(rsget("return_address2"))

	            FItemList(i).Fsongjangdiv     = rsget("defaultsongjangdiv")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

    end sub

    Private Sub Class_Initialize()
        FreturnName     = "(주)텐바이텐"
        FreturnPhone    = "1644-6030"
        Freturnhp       = ""

        FreturnZipcode  = "11154"
        FreturnZipaddr  = "경기도 포천시 군내면 용정경제로2길 83"
        FreturnEtcaddr  = "텐바이텐 물류센터"

        Fsongjangdiv    = "24"
        Fsongjangno     = ""

		FCurrPage = 1
		FPageSize = 20
		FScrollCount = 10
    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

''브랜드별 CS 메모
Class CCSBrandMemo
    public Fbrandid

	public Fis_return_allow

	public Fvacation_startday
	public Fvacation_endday

	public Ftel_start
	public Ftel_end

	public Fis_saturday_work

	public Fbrand_comment

	public Flast_modifyday

    public FRectMakerid

    public sub GetBrandMemo()
        dim sqlStr

        sqlStr = " select brandid, is_return_allow, vacation_startday, vacation_endday, tel_start, tel_end, is_saturday_work, brand_comment, last_modifyday "
        sqlStr = sqlStr + " from " & TABLE_CS_BRAND_MEMO & " "
        sqlStr = sqlStr + " where brandid='" + FRectMakerid + "'"
        rsget.Open sqlStr, dbget, 1

        if Not rsget.Eof then
            Fbrandid         		= rsget("brandid")
            Fis_return_allow		= rsget("is_return_allow")
            Fvacation_startday  	= rsget("vacation_startday")
            Fvacation_endday     	= rsget("vacation_endday")
            Ftel_start         		= rsget("tel_start")
            Ftel_end         		= rsget("tel_end")
            Fis_saturday_work       = rsget("is_saturday_work")
            Fbrand_comment          = db2html(rsget("brand_comment"))
            Flast_modifyday         = rsget("last_modifyday")

        end if
        rsget.Close
    end sub

    Private Sub Class_Initialize()
        '
    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CCsConfirmItem
    public Fasid
    public Fcha                 '''2009추가
    public Fconfirmregmsg
    public Fconfirmreguserid
    public Fconfirmregdate
    public Fconfirmfinishmsg
    public Fconfirmfinishuserid
    public Fconfirmfinishdate

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CCSASMasterItem

    public Fid
    public Fdivcd
    public Fgubun01
    public Fgubun02

    public FdivcdName
    public Fgubun01Name
    public Fgubun02Name

    public FdivcdColor
    public Fgubun01Color
    public Fgubun02Color

    public Forderserial
    public Fcustomername
    public Fuserid
    public Fwriteuser
    public Ffinishuser
    public Ftitle
    public Fcontents_jupsu
    public Fcontents_finish
    public Fcurrstate
    public FcurrstateName
    public FcurrstateColor
    public Fregdate
    public Ffinishdate

    public Fsongjangdiv
    public Fsongjangno
    public Fbeasongdate

    public Frequireupche
    public Fmakerid
    public Fdeleteyn
    public Fextsitename

    '' tbl_as_refund_info's
    public Frefundrequire
    public Frefundresult

    '' tbl_as_upcheAddjungsan
    public Fadd_upchejungsandeliverypay
    public Fadd_upchejungsancause

    public Frefminusorderserial  ''2017/03/27
    
    public Fopentitle           ''고객 오픈 Title
    public Fopencontents        ''고객 오픈 내용
    public Fsitegubun           '' 10x10 or theFingers

    public FErrMsg
    public FAuthcode


    public function IsAsRegAvail(byval iIpkumdiv, byval iCancelYn, byref descMsg)
        IsAsRegAvail = false
        if (iIpkumdiv<2) then
            IsAsRegAvail = false
            descMsg      = "실패한 주문건 또는 정상 주문건이 아닙니다. "
            exit function
        end if

        if (IsCancelProcess) then
            IsAsRegAvail = false

            if (iCancelYn<>"N") then
                IsAsRegAvail = false
                descMsg      = "이미 취소된 거래입니다. - 취소 불가능 "
                exit function
            end if

            IsAsRegAvail = true

        elseif (IsReturnProcess) then
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail = false
                descMsg      = "출고 완료/ 일부 출고 상태가 아닙니다. - 반품 접수 불가능 "
                exit function
            end if

            if (iCancelYn<>"N") then
                IsAsRegAvail = false
                descMsg      = "취소된 거래입니다. - 반품 접수 불가능 "
                exit function
            end if

            IsAsRegAvail = true
        elseif (Fdivcd = "A006") then
            '' 출고시 유의사항
            IsAsRegAvail = true

            if (iIpkumdiv>=8) then
                IsAsRegAvail = false
                descMsg      = "출고 이전 상태가 아닙니다. - 출고시 유의사항 접수 불가능 "
                exit function
            end if
        elseif (Fdivcd = "A009") then
            '' 기타사항
            IsAsRegAvail = true
        elseif  (Fdivcd = "A002") then
            ''서비스발송 :모두 가능하게 변경..
            IsAsRegAvail = true
        elseif (Fdivcd = "A001") then
            ''누락재발송,
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail = false
                descMsg      = "출고 완료/ 일부 출고 상태가 아닙니다. - 누락/서비스 발송 접수 불가능 "
                exit function
            end if

            IsAsRegAvail = true
        elseif (Fdivcd = "A000") then
            ''맞교환
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail = false
                descMsg      = "출고 완료/ 일부 출고 상태가 아닙니다. - 맞교환 접수 불가능 "
                exit function
            end if

            IsAsRegAvail = true
        elseif (Fdivcd = "A003") then
            ''환불요청
            IsAsRegAvail = true
        elseif (Fdivcd = "A005") then
            ''접수시 사이트 구분 체크
            IsAsRegAvail = true
         elseif (Fdivcd = "A700") then
            ''업체 기타 정산.
            IsAsRegAvail = true
        else
            descMsg = "정의 되지 않았습니다." + Fdivcd
        end if

    end function

    ''취소 프로세스
    public function IsCancelProcess()
        IsCancelProcess = fnIsCancelProcess(Fdivcd)
    end function

    ''반품 프로세스
    public function IsReturnProcess()
        IsReturnProcess = fnIsReturnProcess(Fdivcd)
    end function

    ''환불 프로세스
    public function IsRefundProcess()
        IsRefundProcess = fnIsRefundProcess(Fdivcd)
    end function

    ''서비스 발송 프로세스
    public function IsServiceDeliverProcess()
        IsServiceDeliverProcess = fnIsServiceDeliverProcess(Fdivcd)
    end function

    public function IsRefundProcessRequire(iIpkumdiv, iCancelyn)
        FErrMsg = ""
        IsRefundProcessRequire = False

        if (iCancelyn ="Y") or (iCancelyn ="D") then Exit function

		if (iIpkumdiv<4) then  Exit function

        '' 취소, 반품접수
        IsRefundProcessRequire = (IsCancelProcess) or (IsReturnProcess)
    end function

    public function IsRefundProcessRequireBeforePay(iIpkumdiv, iCancelyn)
        FErrMsg = ""
        IsRefundProcessRequireBeforePay = False

        if (iCancelyn ="Y") or (iCancelyn ="D") then Exit function

		'주문 일부취소이고 사용한 마일리지가 취소상품의 금액보다 큰경우 결재전에도 취소가 필요하다.
		'사용한 마일리지는 일부취소 할 수 없다.
		'if (iIpkumdiv<4) then  Exit function

        '' 취소, 반품접수
        IsRefundProcessRequireBeforePay = (IsCancelProcess) or (IsReturnProcess)
    end function

    ''송장 필드가 필요한 정보
    public function IsRequireSongjangNO()
        IsRequireSongjangNO = false

        IsRequireSongjangNO = (Fdivcd="A000") or (Fdivcd="A001") or (Fdivcd="A002") or (Fdivcd="A004") or (Fdivcd="A010") or (Fdivcd="A011")
    end function

    public function GetAsDivCDName()
        GetAsDivCDName = FdivcdName


    end function

    public function GetAsDivCDColor()
        GetAsDivCDColor = FdivcdName


    end function


    public function GetCurrstateName()
        GetCurrstateName = FcurrstateName
    end function

     public function GetCurrstateColor()
        GetCurrstateColor = FcurrstateColor
    end function

    public function GetCauseString()
        GetCauseString = Fgubun01Name
    end function

    public function GetCauseDetailString()
        GetCauseDetailString = Fgubun02Name
    end function



    Private Sub Class_Initialize()
        Fadd_upchejungsandeliverypay = 0
    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CCSASList
    public FItemList()
    public FOneItem

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectUserID
    public FRectUserName
    public FRectOrderSerial
    public FRectStartDate
    public FRectEndDate
    public FRectSearchType
    public FRectIdx
    public FRectMakerid

    public FRectDivcd
    public FRectCurrstate

    public FRectCsAsID
    public FRectNotCsID
    ''
    public FDeliverPay
    public IsUpchebeasongExists
    public IsTenbeasongExists

    public FRectOldOrder

    ''업체사용
    public FRectOnlyJupsu


	Public FRectDeleteYN	' 삭제제외여부
	Public FRectWriteUser	' 접수자아이디 검색


    public Sub GetHisOldRefundInfo()
        dim i,sqlStr

        sqlStr = " select count(asid) as cnt "
        sqlStr = sqlStr + " from " & TABLE_CS_REFUND & " r, "
        sqlStr = sqlStr + " " & TABLE_CSMASTER & " a"
        sqlStr = sqlStr + " where a.userid='" + FRectUserID + "'"
        sqlStr = sqlStr + " and a.id=r.asid"
        sqlStr = sqlStr + " and a.divcd='A003'"
        sqlStr = sqlStr + " and r.returnmethod='R007'"
        sqlStr = sqlStr + " and a.deleteyn='N'"


        rsget.Open sqlStr, dbget, 1
            FTotalCount = rsget("cnt")
        rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr + " r.refundrequire, r.rebankname, r.rebankaccount, r.rebankownername, r.encmethod, r.encaccount "
		sqlStr = sqlStr + " , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_academy.dbo.uf_DecAcctPH1(r.encaccount), '') ELSE '' END) as decaccount "
        sqlStr = sqlStr + " from " & TABLE_CS_REFUND & " r, "
        sqlStr = sqlStr + " " & TABLE_CSMASTER & " a"
        sqlStr = sqlStr + " where a.userid='" + FRectUserID + "'"
        sqlStr = sqlStr + " and a.id=r.asid"
        sqlStr = sqlStr + " and a.divcd='A003'"
        sqlStr = sqlStr + " and r.returnmethod='R007'"
        sqlStr = sqlStr + " and a.deleteyn='N'"
        sqlStr = sqlStr + " order by r.asid desc"

        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSASRefundInfoItem

                FItemList(i).Frefundrequire         = rsget("refundrequire")
				FItemList(i).Frebankname            = rsget("rebankname")
                FItemList(i).Frebankaccount         = rsget("rebankaccount")
                FItemList(i).Frebankownername       = rsget("rebankownername")

                FItemList(i).FencMethod             = rsget("encmethod")
                FItemList(i).FencAccount            = rsget("encaccount")
                FItemList(i).FdecAccount            = rsget("decAccount")

                ''FItemList(i).FrebankCode            = rsget("rebankCode")
				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close
    end Sub

    public Sub GetOneRefundInfo()
        dim i,sqlStr

        sqlStr = "select r.* "
        sqlStr = sqlStr + " ,C1.comm_name as returnmethodName"
		sqlStr = sqlStr + " , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_academy.dbo.uf_DecAcctPH1(r.encaccount), '') ELSE '' END) as decaccount "
        sqlStr = sqlStr + " from " & TABLE_CS_REFUND & " r"
        sqlStr = sqlStr + "     left join " & TABLE_CS_COMMON_CODE  & " C1"
        sqlStr = sqlStr + "     on C1.comm_group='Z090'"
        sqlStr = sqlStr + "     and r.returnmethod=C1.comm_cd"
        sqlStr = sqlStr + " where asid=" + CStr(FRectCsAsID)

        rsget.Open sqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CCSASRefundInfoItem
        if Not rsget.Eof then

            FOneItem.Fasid                  = rsget("asid")
            FOneItem.Forgsubtotalprice      = rsget("orgsubtotalprice")
            FOneItem.Forgitemcostsum        = rsget("orgitemcostsum")
            FOneItem.Forgbeasongpay         = rsget("orgbeasongpay")
            FOneItem.Forgmileagesum         = rsget("orgmileagesum")
            FOneItem.Forgcouponsum          = rsget("orgcouponsum")
            FOneItem.Forgallatdiscountsum   = rsget("orgallatdiscountsum")

            FOneItem.Frefundrequire         = rsget("refundrequire")
            FOneItem.Frefundresult          = rsget("refundresult")
            FOneItem.Freturnmethod          = rsget("returnmethod")

            FOneItem.Frefundmileagesum      = rsget("refundmileagesum")
            FOneItem.Frefundcouponsum       = rsget("refundcouponsum")
            FOneItem.Fallatsubtractsum      = rsget("allatsubtractsum")

            FOneItem.Frefunditemcostsum     = rsget("refunditemcostsum")
            FOneItem.Frefundbeasongpay      = rsget("refundbeasongpay")
            FOneItem.Frefunddeliverypay     = rsget("refunddeliverypay")
            FOneItem.Frefundadjustpay       = rsget("refundadjustpay")
            FOneItem.Fcanceltotal           = rsget("canceltotal")

            FOneItem.Frebankname            = rsget("rebankname")
            FOneItem.Frebankaccount         = rsget("rebankaccount")
            FOneItem.Frebankownername       = rsget("rebankownername")
            FOneItem.FpaygateTid            = rsget("paygateTid")

			FOneItem.FencMethod             = rsget("encmethod")
			FOneItem.FencAccount            = rsget("encaccount")
			FOneItem.FdecAccount            = rsget("decAccount")

            FOneItem.FpaygateresultTid      = rsget("paygateresultTid")
            FOneItem.FpaygateresultMsg      = rsget("paygateresultMsg")


            FOneItem.FreturnmethodName      = rsget("returnmethodName")

            FOneItem.Fupfiledate      = rsget("upfiledate")
        end if
        rsget.Close
    end Sub

    public Sub GetCSASMasterList()
        dim i,sqlStr, AddSQL
        AddSQL = ""

        sqlStr = " select count(A.id) as cnt "
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " A"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_REFUND & " r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join " & TABLE_ORDERMASTER & " m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.sitename <> '" & EXCLUDE_SITENAME & "' "

		if (FRectSearchType="") then
		    if (FRectOrderSerial<>"") then
		        AddSQL = AddSQL + " and A.orderserial='" + FRectOrderSerial + "'"
		    end if
		elseif (FRectSearchType="upcheview") then
		    ''업체가 쿼리시
            AddSQL = AddSQL + " and divcd not in ('A005','A007')"
            AddSQL = AddSQL + " and deleteyn='N'"
            AddSQL = AddSQL + " and requireupche='Y' "
            AddSQL = AddSQL + " and makerid='" + CStr(FRectMakerid) + "' "

            if (FRectOnlyJupsu="on") then
                AddSQL = AddSQL + " and currstate='B001'"
            end if

            if (FRectCurrstate = "notfinish") then
	                AddSQL = AddSQL + " and A.currstate < 'B007' "
	        elseif (FRectCurrstate <> "") then
	                AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if

            if (FRectUserName <> "") then
	                AddSQL = AddSQL + " and A.customername='" + CStr(FRectUserName) + "' "
	        end if

	        if (FRectOrderSerial <> "") then
	                AddSQL = AddSQL + " and A.orderserial='" + CStr(FRectOrderSerial) + "' "
	        end if

	        if (FRectUserID <> "") then
	                AddSQL = AddSQL + " and A.userid='" + CStr(FRectUserID) + "' "
	        end if
		elseif (FRectSearchType = "searchfield") then

	        if (FRectUserID <> "") then
	                AddSQL = AddSQL + " and A.userid='" + CStr(FRectUserID) + "' "
	        end if

	        if (FRectUserName <> "") then
	                AddSQL = AddSQL + " and A.customername='" + CStr(FRectUserName) + "' "
	        end if

	        if (FRectOrderSerial <> "") then
	                AddSQL = AddSQL + " and A.orderserial='" + CStr(FRectOrderSerial) + "' "
	        end if

	        if (FRectMakerid<>"") then
	                AddSQL = AddSQL + " and A.requireupche='Y' "
	                AddSQL = AddSQL + " and A.makerid='" + CStr(FRectMakerid) + "' "
	        end if

	        if (FRectStartDate <> "") then
	                AddSQL = AddSQL + " and A.regdate>='" + CStr(FRectStartDate) + "' "
	        end if

	        if (FRectEndDate <> "") then
	                AddSQL = AddSQL + " and A.regdate <'" + CStr(FRectEndDate) + "' "
	        end if

	        if (FRectCurrstate = "notfinish") then
	                AddSQL = AddSQL + " and A.currstate < 'B007' "
	        elseif (FRectCurrstate <> "") then
	                AddSQL = AddSQL + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if


	        if (FRectDivcd <> "") then
	                AddSQL = AddSQL + " and A.divcd ='" + CStr(FRectDivcd) + "' "
	        end if

			if (FRectWriteUser <> "") then
					AddSQL = AddSQL + " and A.writeUser = '" + CStr(FRectWriteUser) + "' "
			end if

			if (FRectDeleteYN <> "") then
					AddSQL = AddSQL + " and A.deleteyn = '" + CStr(FRectDeleteYN) + "' "
			end if

		elseif (FRectSearchType = "notfinish") then
                ''미처리전체
                AddSQL = AddSQL + " and A.currstate<'B007' and A.deleteyn='N' "
        elseif (FRectSearchType = "norefund") then
                '환불미처리
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A003' "
                AddSQL = AddSQL + " and A.deleteyn='N'"
        elseif (FRectSearchType = "norefundmile") then
                '마일리지환불미처리
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A003' "
                AddSQL = AddSQL + " and A.deleteyn='N'"
                AddSQL = AddSQL + " and R.returnmethod='R900'"
        elseif (FRectSearchType = "norefundetc") then
                '마일리지환불미처리
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A005' "
                AddSQL = AddSQL + " and A.deleteyn='N'"
                ''AddSQL = AddSQL + " and R.returnmethod='R050'"
        elseif (FRectSearchType = "cardnocheck") then
                '카드취소미처리
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd='A007' and A.deleteyn='N' "
        elseif (FRectSearchType = "beasongnocheck") then
                '배송유의사항/취소
                AddSQL = AddSQL + " and A.currstate<'B007' and A.divcd in ('A008','A006') and ((A.requireupche is Null) or (A.requireupche='N')) and deleteyn='N' "
        elseif (FRectSearchType = "upchemifinish") then
                '업체미처리
                AddSQL = AddSQL + " and A.currstate<'B006' and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "upchefinish") then
                '업체처리완료
                AddSQL = AddSQL + " and A.currstate='B006' and A.requireupche='Y' and A.deleteyn='N' "
        elseif (FRectSearchType = "returnmifinish") then
                '회수요청미처리
                AddSQL = AddSQL + " and A.currstate<'B002' and A.divcd ='A010' and A.deleteyn='N'  "
        elseif (FRectSearchType = "confirm") then
                '확인요청 미처리
                AddSQL = AddSQL + " and A.currstate='B005' and A.deleteyn='N' "
        elseif (FRectSearchType = "cancelnofinish") then
                '주문취소 미처리
                AddSQL = AddSQL + " and A.divcd='A008'"
                AddSQL = AddSQL + " and A.currstate='B001' and A.deleteyn='N' "
                AddSQL = AddSQL + " and A.regdate>'2008-04-23'"
        end If



        sqlStr = sqlStr + AddSQL

		'rw sqlStr
        rsget.Open sqlStr, dbget, 1

        if  not rsget.EOF  then
            FTotalCount = rsget("cnt")
        else
            FTotalCount = 0
        end if
        rsget.close


        sqlStr = " select      Top " + CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr + "     A.id, A.divcd, A.gubun01, A.gubun02, A.orderserial, A.customername, A.userid, A.finishuser, A.writeuser, A.title, A.currstate"
        sqlStr = sqlStr + "     ,A.regdate, A.finishdate,A.deleteyn "
        sqlStr = sqlStr + "     , A.requireupche, A.makerid, A.songjangdiv ,A.songjangno"
        sqlStr = sqlStr + "     ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult"
        sqlStr = sqlStr + "     ,m.sitename, m.authcode"
        sqlStr = sqlStr + "     ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename, C4.comm_color as currstatecolor"
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " A"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_REFUND & " r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join " & TABLE_ORDERMASTER & " m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
        sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and m.sitename <> '" & EXCLUDE_SITENAME & "' "

        sqlStr = sqlStr + AddSQL

        sqlStr = sqlStr + " order by id desc "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

        FtotalPage =  CLng(FTotalCount\FPageSize)
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
                set FItemList(i) = new CCSASMasterItem

                FItemList(i).Fid                = rsget("id")
                FItemList(i).Fdivcd             = rsget("divcd")
                FItemList(i).FdivcdName         = db2html(rsget("divcdname"))

                FItemList(i).Forderserial       = rsget("orderserial")
                FItemList(i).Fcustomername      = db2html(rsget("customername"))
                FItemList(i).Fuserid            = rsget("userid")
                FItemList(i).Fwriteuser         = rsget("writeuser")
                FItemList(i).Ffinishuser        = rsget("finishuser")
                FItemList(i).Ftitle             = db2html(rsget("title"))
                FItemList(i).Fcurrstate         = rsget("currstate")
                FItemList(i).Fcurrstatename     = rsget("currstatename")
                FItemList(i).FcurrstateColor    = rsget("currstatecolor")

                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Ffinishdate        = rsget("finishdate")

                FItemList(i).Fgubun01           = rsget("gubun01")
                FItemList(i).Fgubun02           = rsget("gubun02")

                FItemList(i).Fgubun01Name       = db2html(rsget("gubun01name"))
                FItemList(i).Fgubun02Name       = db2html(rsget("gubun02name"))

                FItemList(i).Fdeleteyn          = rsget("deleteyn")

                FItemList(i).Frefundrequire     = rsget("refundrequire")
                FItemList(i).Frefundresult      = rsget("refundresult")

                FItemList(i).Fsongjangdiv       = rsget("songjangdiv")
                FItemList(i).Fsongjangno        = rsget("songjangno")

                FItemList(i).Frequireupche      = rsget("requireupche")
                FItemList(i).Fmakerid           = rsget("makerid")

                FItemList(i).FExtsitename          = rsget("sitename")
                FItemList(i).Fauthcode          = rsget("authcode")

                rsget.MoveNext
                i = i + 1
            loop
        end if
        rsget.close
    end sub



    public Sub GetCSASTotalPrevCancelCount()
        dim i,sqlStr

        sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " "
        sqlStr = sqlStr + " where 1 = 1 "

        if (FRectOrderSerial <> "") then
                sqlStr = sqlStr + " and orderserial='" + CStr(FRectOrderSerial) + "' "
        end if

        sqlStr = sqlStr + " and deleteyn='N' and divcd in ('A003','A005','A007') "
        rsget.Open sqlStr, dbget, 1

        if  not rsget.EOF  then
                FResultCount = rsget("cnt")
        else
                FResultCount = 0
        end if
        rsget.close
    end sub

    public Sub GetOneCSASMaster()
        dim i,sqlStr

        sqlStr = " select top 1 A.*, IsNULL(J.add_upchejungsandeliverypay,0) as add_upchejungsandeliverypay, J.add_upchejungsancause "
        sqlStr = sqlStr + " ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult, IsNULL(refminusorderserial,'') as refminusorderserial"  ''refminusorderserial 추가 2017/03/27
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename"
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " A "
        sqlStr = sqlStr + " Left join " & TABLE_UPCHE_ADD_JUNGSAN & " J"
        sqlStr = sqlStr + "  on A.id=J.asid"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_REFUND & " r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"

        sqlStr = sqlStr + " where id= " + CStr(FRectCsAsID) + " "

        if (FRectMakerID<>"") then   ''업체 조회용.
            sqlStr = sqlStr + " and A.makerid='"&FRectMakerID&"'"
        end if
        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount

        if  not rsget.EOF  then
            set FOneItem = new CCSASMasterItem

            FOneItem.Fid                  = rsget("id")
            FOneItem.Fdivcd               = rsget("divcd")
            FOneItem.Fgubun01             = rsget("gubun01")
            FOneItem.Fgubun02             = rsget("gubun02")

            FOneItem.FdivcdName           = db2html(rsget("divcdname"))
            FOneItem.Fgubun01Name         = db2html(rsget("gubun01name"))
            FOneItem.Fgubun02Name         = db2html(rsget("gubun02name"))

            FOneItem.Forderserial         = rsget("orderserial")
            FOneItem.Fcustomername        = db2html(rsget("customername"))
            FOneItem.Fuserid              = rsget("userid")
            FOneItem.Fwriteuser           = rsget("writeuser")
            FOneItem.Ffinishuser          = rsget("finishuser")
            FOneItem.Ftitle               = db2html(rsget("title"))
            FOneItem.Fcontents_jupsu      = db2html(rsget("contents_jupsu"))
            FOneItem.Fcontents_finish     = db2html(rsget("contents_finish"))
            FOneItem.Fcurrstate           = rsget("currstate")
            FOneItem.FcurrstateName       = rsget("currstatename")
            FOneItem.Fregdate             = rsget("regdate")
            FOneItem.Ffinishdate          = rsget("finishdate")

            FOneItem.Fdeleteyn            = rsget("deleteyn")
            FOneItem.Fextsitename         = rsget("extsitename")

            FOneItem.Fopentitle           = db2html(rsget("opentitle"))
            FOneItem.Fopencontents        = db2html(rsget("opencontents"))


            FOneItem.Fsitegubun           = rsget("sitegubun")

            FOneItem.Fsongjangdiv         = rsget("songjangdiv")
            FOneItem.Fsongjangno          = rsget("songjangno")

            FOneItem.Frequireupche        = rsget("requireupche")
            FOneItem.Fmakerid             = rsget("makerid")

            FOneItem.Fadd_upchejungsandeliverypay = rsget("add_upchejungsandeliverypay")
            FOneItem.Fadd_upchejungsancause       = rsget("add_upchejungsancause")
            
            FOneItem.Frefminusorderserial 	= rsget("refminusorderserial")
            
'            FOneItem.Fbeasongdate         = rsget("beasongdate")
'            FOneItem.Frefundrequire       = rsget("refundrequire")
'            FOneItem.Frefundresult        = rsget("refundresult")

        end if
        rsget.close
    end sub

    public Sub GetOrderDetailByCsDetail()
        dim SqlStr, i

		sqlStr = "select d." & FIELD_DETAILIDX & " as orderdetailidx, d.orderserial,d.itemid,d.itemoption,d.itemno,d.itemcost, d.buycash, d.reducedprice as discountAssingedCost"
		sqlStr = sqlStr + " ,d.mileage,d.cancelyn "
		sqlStr = sqlStr + " ,d.itemname, d.makerid, d.itemoptionname "
		sqlStr = sqlStr + " ,d.currstate as orderdetailcurrstate, d.upcheconfirmdate, d.songjangdiv, d.songjangno"
		sqlStr = sqlStr + " ,d.beasongdate, d.isupchebeasong, d.issailitem , d.cancelyn "
		sqlStr = sqlStr + " ,d.oitemdiv, d.odlvType, d." & FIELD_ITEMCOUPONIDX & " as itemcouponidx, d.bonuscouponidx"
		sqlStr = sqlStr + " ,c.id, c.masterid, IsNULL(c.regitemno,0) as regitemno, IsNULL(c.confirmitemno,0) as confirmitemno"
		sqlStr = sqlStr + " ,c.gubun01, c.gubun02, c.regdetailstate"
		sqlStr = sqlStr + " ,C2.comm_name as gubun01name, C3.comm_name as gubun02name"
		sqlStr = sqlStr + " ,i.smallimage "

		sqlStr = sqlStr + " ,IsNull((select top 1 ad.confirmitemno "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 		[db_academy].[dbo].[tbl_academy_as_list] a "
		sqlStr = sqlStr + " 		join [db_academy].[dbo].[tbl_academy_as_detail] ad "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			a.id = ad.masterid "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and a.divcd = 'A004' "
		sqlStr = sqlStr + " 		and a.deleteyn = 'N' "
		sqlStr = sqlStr + " 		and a.orderserial = d.orderserial "
		sqlStr = sqlStr + " 		and ad.itemid = d.itemid "
		sqlStr = sqlStr + " 		and ad.itemoption = d.itemoption),0) as prevreturnno "

		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d "
		else
		    sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " d "
		end if
		sqlStr = sqlStr + " left join " & TABLE_ITEM & " i on d.itemid=i.itemid"
		sqlStr = sqlStr + " left join " & TABLE_CSDETAIL & " c "
		sqlStr = sqlStr + " on c.masterid=" + CStr(FRectCsAsID) + ""
		sqlStr = sqlStr + " and c.orderdetailidx=d." & FIELD_DETAILIDX & " "
		sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C2"
        sqlStr = sqlStr + "  on c.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C3"
        sqlStr = sqlStr + "  on c.gubun02=C3.comm_cd"

		sqlStr = sqlStr + " where d.orderserial='" + CStr(FRectOrderSerial) + "'"

        ''sqlStr = sqlStr + " order by d.isupchebeasong, d.makerid, d.itemid, d.itemoption"
		sqlStr = sqlStr + " order by d.itemid, d.itemoption"
		''response.write sqlStr
		''response.end

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CCSASDetailItem

            ''tbl_as_detail's
            FItemList(i).Fid              = rsget("id")
            FItemList(i).Fmasterid        = rsget("masterid")
            FItemList(i).Fgubun01         = rsget("gubun01")
            FItemList(i).Fgubun02         = rsget("gubun02")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget("regdetailstate")

            ''tbl_order_detail's
            FItemList(i).Forderdetailidx  = rsget("orderdetailidx")
            FItemList(i).Forderserial     = rsget("orderserial")
            FItemList(i).Fitemid          = rsget("itemid")
            FItemList(i).Fitemoption      = rsget("itemoption")
            FItemList(i).Fmakerid         = rsget("makerid")
            FItemList(i).Fitemname        = db2html(rsget("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
            FItemList(i).Fitemcost        = rsget("itemcost")
            FItemList(i).Fbuycash         = rsget("buycash")

			FItemList(i).Fitemno          = rsget("itemno")
			FItemList(i).Fprevreturnno    = rsget("prevreturnno")


            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")
            FItemList(i).FCancelyn        = rsget("cancelyn")
            FItemList(i).ForderDetailcurrstate = rsget("orderdetailcurrstate")

            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")

            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")


            ''쿠폰 사용하거나, 마일리지샵 상품은 할인 안되었음.
''            if (rsget("oitemdiv")="82") or (rsget("itemcouponidx")<>0) or (rsget("issailitem")="Y") then
''                FItemList(i).FAllAtDiscountedPrice = 0
''            else
''                FItemList(i).FAllAtDiscountedPrice = round(((1-0.94) * FItemList(i).Fitemcost / 100) * 100 )
''            end if


            ''tbl_item's
            FItemList(i).FSmallImage  	  = webImgUrl + DIRECTORY_IMAGE_SMALL + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget("gubun01name")
            FItemList(i).Fgubun02name   = rsget("gubun02name")
			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub

    public Sub GetCsDetailList()
        dim SqlStr, i

		sqlStr = "select c.*"
		sqlStr = sqlStr + " ,d.currstate as orderdetailcurrstate"
		sqlStr = sqlStr + " ,d.reducedprice as discountAssingedCost, d.oitemdiv, d.odlvType, d.issailitem, d." & FIELD_ITEMCOUPONIDX & " as itemcouponidx, d.bonuscouponidx"
		sqlStr = sqlStr + " ,IsNULL(d.itemcost,0) as OrderItemcost"
		sqlStr = sqlStr + " ,C2.comm_name as gubun01name, C3.comm_name as gubun02name"
		sqlStr = sqlStr + " ,i.smallimage "

		sqlStr = sqlStr + " from " & TABLE_CSDETAIL & " c "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " left join [db_log].[dbo].tbl_old_order_detail_2003 d"
		    sqlStr = sqlStr + "  on c.orderdetailidx=d." & FIELD_DETAILIDX & ""
		else
		    sqlStr = sqlStr + " left join " & TABLE_ORDERDETAIL & " d"
		    sqlStr = sqlStr + "  on c.orderdetailidx=d." & FIELD_DETAILIDX & ""
		end if

		sqlStr = sqlStr + " left join " & TABLE_ITEM & " i "
		sqlStr = sqlStr + "  on c.itemid=i.itemid"
		sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C2"
        sqlStr = sqlStr + "  on c.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join " & TABLE_CS_COMMON_CODE  & " C3"
        sqlStr = sqlStr + "  on c.gubun02=C3.comm_cd"

		sqlStr = sqlStr + " where c.masterid=" + CStr(FRectCsAsID) + ""
        sqlStr = sqlStr + " order by c.isupchebeasong, c.makerid, c.itemid, c.itemoption"

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CCSASDetailItem

            FItemList(i).Fid              = rsget("id")
            FItemList(i).Fmasterid        = rsget("masterid")
            FItemList(i).Fgubun01         = rsget("gubun01")
            FItemList(i).Fgubun02         = rsget("gubun02")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")

            FItemList(i).Fregdetailstate  = rsget("regdetailstate")   ''접수 당시 진행 상태
            FItemList(i).Forderdetailidx  = rsget("orderdetailidx")
            FItemList(i).Forderserial     = rsget("orderserial")
            FItemList(i).Fitemid          = rsget("itemid")
            FItemList(i).Fitemoption      = rsget("itemoption")
            FItemList(i).Fmakerid         = rsget("makerid")
            FItemList(i).Fitemname        = db2html(rsget("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
            FItemList(i).Fitemcost        = rsget("itemcost")
            FItemList(i).Fbuycash         = rsget("buycash")
            FItemList(i).Fitemno          = rsget("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")

            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")

            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")


            FItemList(i).Forderdetailcurrstate  = rsget("orderdetailcurrstate")

            FItemList(i).FSmallImage  	  = webImgUrl + DIRECTORY_IMAGE_SMALL + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget("gubun01name")
            FItemList(i).Fgubun02name   = rsget("gubun02name")

            if (FItemList(i).Fitemcost=0) then
                FItemList(i).Fitemcost = rsget("OrderItemcost")
            end if

			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub

    public Sub GetCSASTotalCount()
        dim i,sqlStr

        sqlStr = " select count(id) as cnt "
        sqlStr = sqlStr + " from " & TABLE_CSMASTER & " "
        sqlStr = sqlStr + " where 1 = 1 "

        if (FRectNotCsID<> "") then
            sqlStr = sqlStr + " and id<>'" + CStr(FRectNotCsID) + "' "
        end if

        if (FRectUserID <> "") then
                sqlStr = sqlStr + " and userid='" + CStr(FRectUserID) + "' "
        end if

        if (FRectUserName <> "") then
                sqlStr = sqlStr + " and customername='" + CStr(FRectUserName) + "' "
        end if

        if (FRectOrderSerial <> "") then
                sqlStr = sqlStr + " and orderserial='" + CStr(FRectOrderSerial) + "' "
        end if

        if (FRectStartDate <> "") then
                sqlStr = sqlStr + " and regdate>='" + CStr(FRectStartDate) + "' "
        end if

        if (FRectEndDate <> "") then
                sqlStr = sqlStr + " and regdate <'" + CStr(FRectEndDate) + "' "
        end if

        if (FRectSearchType = "norefund") then
                '환불미처리
                sqlStr = sqlStr + " and currstate<7 and divcd in ('3','5') "
        elseif (FRectSearchType = "cardnocheck") then
                '카드취소미처리
                sqlStr = sqlStr + " and currstate<7 and divcd='7' "
        elseif (FRectSearchType = "beasongnocheck") then
                '배송유의사항/취소
                sqlStr = sqlStr + " and currstate<7 and divcd in ('8','6') and ((requireupche is Null) or (requireupche='N')) "
        elseif (FRectSearchType = "upchemifinish") then
                '업체미처리
                sqlStr = sqlStr + " and currstate<6 and requireupche='Y' and deleteyn='N' "
        elseif (FRectSearchType = "upchefinish") then
                '업체처리완료
                sqlStr = sqlStr + " and currstate=6 and requireupche='Y' and deleteyn='N' "
        elseif (FRectSearchType = "returnmifinish") then
                '회수요청미처리
                sqlStr = sqlStr + " and currstate<2 and divcd ='10' "
        end if

        rsget.Open sqlStr, dbget, 1

        if  not rsget.EOF  then
            FResultCount = rsget("cnt")
        else
            FResultCount = 0
        end if
        rsget.close
    end sub

    public Sub GetOneCsDeliveryItem()
        dim i,sqlStr

        sqlStr = " select top 1 A.* "
        sqlStr = sqlStr + " from " & TABLE_CS_DELIVERY & " A "
        sqlStr = sqlStr + " where asid= " + CStr(FRectCsAsID) + " "

        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount

        if  not rsget.EOF  then
            set FOneItem = new CCSDeliveryItem
            FOneItem.Fasid              = rsget("asid")
            FOneItem.Freqname           = db2html(rsget("reqname"))
            FOneItem.Freqphone          = rsget("reqphone")
            FOneItem.Freqhp             = rsget("reqhp")
            FOneItem.Freqzipcode        = rsget("reqzipcode")
            FOneItem.Freqzipaddr        = rsget("reqzipaddr")
            FOneItem.Freqetcaddr        = db2html(rsget("reqetcaddr"))
            FOneItem.Freqetcstr          = db2html(rsget("reqetcstr"))
            FOneItem.Fsongjangdiv       = rsget("songjangdiv")
            FOneItem.Fsongjangno        = rsget("songjangno")
            FOneItem.Fregdate           = rsget("regdate")
            FOneItem.Fsenddate          = rsget("senddate")

        end if
        rsget.close

    end Sub

    public Sub GetOneCsDeliveryItemFromDefaultOrder()
        dim i,sqlStr

        sqlStr = " select m.reqname, m.reqphone, m.reqhp, m.reqzipcode, m.reqzipaddr, m.reqaddress"
        sqlStr = sqlStr + " from db_order.dbo.tbl_order_master m"
        sqlStr = sqlStr + "     Join " & TABLE_CSMASTER & " a"
        sqlStr = sqlStr + "     on m.orderserial=a.orderserial"
        sqlStr = sqlStr + "     and a.id=" + CStr(FRectCsAsID) + " "

        rsget.Open sqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        if  not rsget.EOF  then
            set FOneItem = new CCSDeliveryItem
            FOneItem.Fasid              = FRectCsAsID
            FOneItem.Freqname           = db2html(rsget("reqname"))
            FOneItem.Freqphone          = rsget("reqphone")
            FOneItem.Freqhp             = rsget("reqhp")
            FOneItem.Freqzipcode        = rsget("reqzipcode")
            FOneItem.Freqzipaddr        = rsget("reqzipaddr")
            FOneItem.Freqetcaddr        = db2html(rsget("reqaddress"))
            'FOneItem.Freqetcstr          = db2html(rsget("reqetcstr"))
            'FOneItem.Fsongjangdiv       = rsget("songjangdiv")
            'FOneItem.Fsongjangno        = rsget("songjangno")
            'FOneItem.Fregdate           = rsget("regdate")
            'FOneItem.Fsenddate          = rsget("senddate")

        end if
        rsget.close

        if (FResultCount<1) then
            sqlStr = " select m.reqname, m.reqphone, m.reqhp, m.reqzipcode, m.reqzipaddr, m.reqaddress"
            sqlStr = sqlStr + " from db_log.dbo.tbl_old_order_master_2003 m"
            sqlStr = sqlStr + "     Join " & TABLE_CSMASTER & " a"
            sqlStr = sqlStr + "     on m.orderserial=a.orderserial"
            sqlStr = sqlStr + "     and a.id=" + CStr(FRectCsAsID) + " "

            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            if  not rsget.EOF  then
                set FOneItem = new CCSDeliveryItem
                FOneItem.Fasid              = FRectCsAsID
                FOneItem.Freqname           = db2html(rsget("reqname"))
                FOneItem.Freqphone          = rsget("reqphone")
                FOneItem.Freqhp             = rsget("reqhp")
                FOneItem.Freqzipcode        = rsget("reqzipcode")
                FOneItem.Freqzipaddr        = rsget("reqzipaddr")
                FOneItem.Freqetcaddr        = db2html(rsget("reqaddress"))
                'FOneItem.Freqetcstr          = db2html(rsget("reqetcstr"))
                'FOneItem.Fsongjangdiv       = rsget("songjangdiv")
                'FOneItem.Fsongjangno        = rsget("songjangno")
                'FOneItem.Fregdate           = rsget("regdate")
                'FOneItem.Fsenddate          = rsget("senddate")

            end if
            rsget.close
        end if
    end Sub

    public sub GetOneCsConfirmItem()
        dim sqlStr, i
        sqlStr = " select top 1 * from " & TABLE_CS_CONFIRM & ""
        sqlStr = sqlStr + " where asid=" + CStr(FRectCsAsID)



        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount

        if  not rsget.EOF  then
            set FOneItem = new CCsConfirmItem

            FOneItem.Fasid                  = rsget("asid")
            FOneItem.Fconfirmregmsg         = db2html(rsget("confirmregmsg"))
            FOneItem.Fconfirmreguserid      = rsget("confirmreguserid")
            FOneItem.Fconfirmregdate        = rsget("confirmregdate")
            FOneItem.Fconfirmfinishmsg      = db2html(rsget("confirmfinishmsg"))
            FOneItem.Fconfirmfinishuserid   = rsget("confirmfinishuserid")
            FOneItem.Fconfirmfinishdate     = rsget("confirmfinishdate")

        end if
        rsget.close

    end sub

    Private Sub Class_Initialize()
        FCurrPage       = 1
        FPageSize       = 10
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
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
