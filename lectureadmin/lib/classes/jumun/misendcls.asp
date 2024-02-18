<%
Class COrderMasterWithCSItem
	public FOrderSerial
	public FCancelyn
    public Fbuyname
    public Fbuyhp
    public Fbuyemail


	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COldMiSendItem
	public FOrderSerial
	public FMakerId
	public FItemId
	public FItemName
	public FItemOptionName
	public FItemNo

	public FIsUpcheBeasong
	public FCurrState

	public Fitemlackno
	public FCode
	public FState
	public FRegDate
	public FIpgoDate

	public FBuyName
	public FBuyPhone
	public FBuyHP
	public FReqName
	public FIpkumDate

	public FDeliveryNo
	public FSiteName
	public FUserId
	public FSubTotalPrice
	public Fipkumdiv
	public Fbaljudate

	public FrequestString
	public FfinishString

    ''--2009 추가
    public Fbuyemail
    public Fidx
    public FItemcnt
    public FItemoption
    public Fupcheconfirmdate
    public Fbeasongdate
    public FSongjangno
    public FSongjangdiv

    public FMisendReason
    public FMisendState
    public FMisendipgodate

    public FisSendSMS
    public FisSendEmail
    public FisSendCall

    public Fcompany_name
    public Fcompany_tel
    public Fsmallimage
    public FCancelYn
    public FDetailCancelYn

    public function getBeasongDPlusDateStr()
        getBeasongDPlusDateStr = ""

        if IsNULL(Fbaljudate) then
            exit function
        end if

        if IsNULL(Fbeasongdate) then
            getBeasongDPlusDateStr = "D+" & DateDiff("d",Fbaljudate,now())
            exit function
        end if

        if (DateDiff("d",Fbaljudate,Fbeasongdate)<1) then
            getBeasongDPlusDateStr = "D+0"
        else
            getBeasongDPlusDateStr = "D+" & DateDiff("d",Fbaljudate,Fbeasongdate)
        end if
    end function

    public function getBeasongDPlusDate()
        getBeasongDPlusDate = ""

        if IsNULL(Fbaljudate) then
            exit function
        end if

        if IsNULL(Fbeasongdate) then
            getBeasongDPlusDate = DateDiff("d",Fbaljudate,now())
            exit function
        end if

        getBeasongDPlusDate = DateDiff("d",Fbaljudate,Fbeasongdate)
    end function

    public function getMisendDPlusDate
        dim BaseDate , tmp
        if Not IsNULL(Fbaljudate) then
            BaseDate = Left(CStr(Fbaljudate),10)
        elseIF Not IsNULL(Fupcheconfirmdate) then
            BaseDate = Left(CStr(Fupcheconfirmdate),10)
        else
            BaseDate = Left(CStr(now()),10)
        end if

        tmp = DateDiff("d",BaseDate,FMisendipgodate)
        if (tmp>=0) then
            getMisendDPlusDate = tmp
        else
            getMisendDPlusDate = 0
        end if
    end function

    public function getSMSText()
        dim smstext
        smstext = ""

        if (FMisendipgodate<>"") then
            if (FMisendReason="05") then

            elseif (FMisendReason="02") then  ''주문제작
                ''출고 소요일수 D+2이상
                if (getMisendDPlusDate>1) then
                    smstext = "[더핑거스 출고지연안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"주문제작 상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[더핑거스 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품이 "&VbCrlf
                    smstext = smstext&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."
                end if
            elseif (FMisendReason="03") then  ''출고지연
                ''출고 소요일수 D+2이상
                if (getMisendDPlusDate>1) then
                    smstext = "[더핑거스 출고지연안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품이 "&VbCrlf
                    smstext = smstext&FMisendipgodate&"에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[더핑거스 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품이 "&VbCrlf
                    smstext = smstext&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."

                end if
            elseif (FMisendReason="04") then  ''예약상품
                ''출고 소요일수 D+2이상
                if (getMisendDPlusDate>1) then
                    smstext = "[더핑거스 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"예약배송상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[더핑거스 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"예약배송상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."

                end if
            end if
        end if
        getSMSText = smstext
    end function

    public function isMisendAlreadyInputed()
        isMisendAlreadyInputed = Not (IsNULL(FMisendReason) or (FMisendReason="00") or (FMisendReason=""))
    end function

    public function getDlvCompanyName()
        if FIsUpchebeasong="Y" then
            getDlvCompanyName = Fcompany_name
        else
            getDlvCompanyName = "더핑거스"
        end if
    end function

    Public function getUpcheDeliverStateName()
		 if IsNull(FCurrState) then
		    if (Fipkumdiv<4) then
		        getUpcheDeliverStateName = "주문접수"
		    else
			    getUpcheDeliverStateName = "결제완료"
			end if
		 elseif FCurrState="2" then
			 getUpcheDeliverStateName = "주문통보"
		 elseif FCurrState="3" then
			 getUpcheDeliverStateName = "주문확인"
		 elseif FCurrState="7" then
			 getUpcheDeliverStateName = "출고완료"
		 else
			 getUpcheDeliverStateName = ""
		 end if
	 end Function

	public function getUpCheDeliverStateColor()
		if IsNull(FCurrState) then
		    if (Fipkumdiv<4) then
		        getUpCheDeliverStateColor ="#444444"
		    else
			    getUpCheDeliverStateColor ="#3300CC"
			end if

		elseif FCurrState="2" then
			getUpCheDeliverStateColor="#336600"
		elseif FCurrState="3" then
			getUpCheDeliverStateColor="#CC9933"
		elseif FCurrState="7" then
			getUpCheDeliverStateColor="#FF0000"
		else
			getUpCheDeliverStateColor="#000000"
		end if
	end function

	public function IpkumDivColor()
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
	end function

	Public function IpkumDivName()
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
			IpkumDivName="출고완료"
		end if
	end function

	public function getIpgoMayDay()
		if IsNULL(FIpgoDate) then
			getIpgoMayDay = "&nbsp;"
		else
			getIpgoMayDay = CStr(FIpgoDate)
		end if
	end function

    public function getMiSendCodeColor()
		if FMisendReason="05" then
			getMiSendCodeColor = "#FF0000"
		else
			getMiSendCodeColor = "#000000"
		end if
	end function

	public function getMiSendCodeName()
		if FCode="00" then
			getMiSendCodeName = "입력대기"
		elseif FCode="01" then
			getMiSendCodeName = "재고부족" ''사용안함
		elseif FCode="02" then
			getMiSendCodeName = "주문제작"
		elseif FCode="03" then
			getMiSendCodeName = "출고지연"
		elseif FCode="04" then
			getMiSendCodeName = "예약상품" ''"포장대기" ''사용안함
		elseif FCode="05" then
			getMiSendCodeName = "품절출고불가"
		elseif FCode="06" then
			getMiSendCodeName = "신상품입고지연" ''사용안함
		else
			getMiSendCodeName = "&nbsp;"
		end if
	end function

	public Function GetOptionName()
		if IsNULL(FItemOptionName) or (FItemOptionName="") then
			GetOptionName = "&nbsp;"
		else
			GetOptionName = FItemOptionName
		end if
	end Function

	public Function GetBeagonGubunColor()
		if FIsUpcheBeasong="Y" then
			GetBeagonGubunColor = "#000000"
		else
			GetBeagonGubunColor = "#33EE33"
		end if
	end function

	public Function GetBeagonGubunName()
		if FIsUpcheBeasong="Y" then
			GetBeagonGubunName = "업체"
		else
			GetBeagonGubunName = "10x10"
		end if
	end function

	public Function GetBeagonStateColor()
		if (IsNULL(FCurrState) or (FCurrState=0)) and FIsUpcheBeasong="Y" then
			GetBeagonStateColor = "#EE3333"
		elseif FCurrState="3" then
			GetBeagonStateColor = "#3333EE"
		else
			GetBeagonStateColor = "#000000"
		end if
	end function

	public Function GetBeagonStateName()
		if (IsNULL(FCurrState) or (FCurrState=0)) and FIsUpcheBeasong="Y" then
			GetBeagonStateName = "미확인"
		elseif FCurrState="3" then
			GetBeagonStateName = "업체확인"
		else
			GetBeagonStateName = "&nbsp;"
		end if
	end function

    ''2009년 상태 변경 isSendSMS, isSendEmail, isSendCall
	public Function GetStateString()
		if FState = "0" then
			GetStateString = "미처리"
		elseif FState="1" then
			GetStateString = "SMS완료"
		elseif FState="2" then
			GetStateString = "안내Mail완료"
		elseif FState="3" then
			GetStateString = "통화완료"
		''elseif FState="3" then
		''	GetStateString = "배송실처리"
		elseif FState="4" then
			GetStateString = "고객안내"         '' 2009신규
		elseif FState="6" then
			GetStateString = "CS처리완료"
		elseif FState="7" then
			GetStateString = "배송실 처리완료"
		else
			GetStateString = "&nbsp;"
		end if
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COldMiSend
	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FRectStart
	public FRectEnd

	public FRectDelayDate
	public FRectNotInCludeUpcheCheck
	public FRectInCludeAlreadyInputed
	public FRectDeliveryNo
	public FRectOrderingOpt

	public FRectNotIncludeItemList
	public FRectOrderSerial

    public FRectMakerid
	public FRectItemId
	public FRectIsupchebeasong
	public FRectDetailidx

	''주문내역중 미배송리스트 / 미배송 없는내역도 조회.
	public function getMiSendOrderDetailList()
        dim sqlStr, i
        sqlStr = "exec [db_academy].[dbo].sp_Ten_Mibeasong_Item_GetList '" + CStr(FRectOrderSerial) + "'"
        rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		i=0
		redim FItemList(FResultCount)
		if not rsACADEMYget.EOF then
			do until rsACADEMYget.eof
				set FItemList(i) = new COldMiSendItem

    			FItemList(i).Fidx				  = rsACADEMYget("idx")
    			FItemList(i).FOrderserial		  = rsACADEMYget("orderserial")
    			FItemList(i).FItemid 			  = rsACADEMYget("itemid")
    			FItemList(i).FItemoption     	  = rsACADEMYget("itemoption")
    			FItemList(i).FItemname 		      = db2html(rsACADEMYget("itemname"))
    			FItemList(i).FItemoptionName      = db2html(rsACADEMYget("itemoptionname"))
    			FItemList(i).FItemcnt             = rsACADEMYget("itemno")

    			FItemList(i).FMakerid 			  = rsACADEMYget("makerid")
    			FItemList(i).FBuyname             = db2html(rsACADEMYget("buyname"))
    			FItemList(i).FReqname			  = db2html(rsACADEMYget("reqname"))
    			FItemList(i).FCancelYn		      = rsACADEMYget("cancelyn")
    			FItemList(i).FDetailCancelYn	  = rsACADEMYget("detailcancelyn")
    			FItemList(i).FRegdate			  = rsACADEMYget("regdate")
    			FItemList(i).FIpkumdate		      = rsACADEMYget("ipkumdate")
    			FItemList(i).FBaljudate		      = rsACADEMYget("baljudate")
    			FItemList(i).Fupcheconfirmdate    = rsACADEMYget("upcheconfirmdate")
    			FItemList(i).FCurrstate		      = rsACADEMYget("currstate")      '' DetailState

    			FItemList(i).Fbeasongdate         = rsACADEMYget("beasongdate")

    			FItemList(i).FisUpcheBeasong      = rsACADEMYget("isUpcheBeasong")
    			FItemList(i).FSongjangno          = rsACADEMYget("songjangno")
    			FItemList(i).FSongjangdiv         = rsACADEMYget("songjangdiv")

                FItemList(i).FCode                = rsACADEMYget("code")           '' for old version
                FItemList(i).FState               = rsACADEMYget("state")          '' for old version
                FItemList(i).Fipgodate            = rsACADEMYget("ipgodate")       '' for old version

                FItemList(i).FMisendReason        = rsACADEMYget("code")
                FItemList(i).FMisendState         = rsACADEMYget("state")
                FItemList(i).FMisendipgodate      = rsACADEMYget("ipgodate")

                FItemList(i).FisSendSMS           = rsACADEMYget("isSendSMS")
                FItemList(i).FisSendEmail         = rsACADEMYget("isSendEmail")
                FItemList(i).FisSendCall          = rsACADEMYget("isSendCall")
                FItemList(i).Fbuyemail            = rsACADEMYget("buyemail")
                FItemList(i).FbuyHp               = rsACADEMYget("buyHp")

                FItemList(i).FrequestString       = db2Html(rsACADEMYget("reqstr"))
                FItemList(i).FItemNo              = rsACADEMYget("itemno")
                FItemList(i).Fitemlackno          = rsACADEMYget("itemlackno")
                FItemList(i).FfinishString        = db2Html(rsACADEMYget("finishstr"))


                FItemList(i).Fcompany_name        = db2Html(rsACADEMYget("company_name"))
                FItemList(i).Fcompany_tel         = db2Html(rsACADEMYget("company_tel"))

                'FItemList(i).Fsmallimage          = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsACADEMYget("smallimage")
                FItemList(i).Fsmallimage		  = imgFingers & "/diyitem/webimage/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimage")

                FItemList(i).FCancelYn            = rsACADEMYget("detailcancelyn")
                i=i+1
                rsACADEMYget.MoveNext
            loop

        end if
        rsACADEMYget.Close
    end function

    public function getOneOldMisendItem()
        dim sqlStr
        sqlStr = "exec [db_academy].[dbo].usp_Academy_Mibeasong_Item_GetData " + CStr(FRectDetailidx) + ""
        rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		if not rsACADEMYget.EOF then
            set FOneItem = new COldMiSendItem

			FOneItem.Fidx				  = rsACADEMYget("idx")
			FOneItem.FOrderserial		  = rsACADEMYget("orderserial")
			FOneItem.FItemid 			  = rsACADEMYget("itemid")
			FOneItem.FItemoption     	  = rsACADEMYget("itemoption")
			FOneItem.FItemname 		      = db2html(rsACADEMYget("itemname"))
			FOneItem.FItemoptionName      = db2html(rsACADEMYget("itemoptionname"))
			FOneItem.FItemcnt             = rsACADEMYget("itemno")

			FOneItem.FMakerid 			  = rsACADEMYget("makerid")
			FOneItem.FBuyname             = db2html(rsACADEMYget("buyname"))
			FOneItem.FReqname			  = db2html(rsACADEMYget("reqname"))
			FOneItem.FUserID              = rsACADEMYget("userid")

			FOneItem.FCancelYn		      = rsACADEMYget("cancelyn")  ''master cancelyn
			FOneItem.FDetailCancelYn		      = rsACADEMYget("detailcancelyn")  ''detailcancelyn
			FOneItem.FRegdate			  = rsACADEMYget("regdate")
			FOneItem.FIpkumdate		      = rsACADEMYget("ipkumdate")
			FOneItem.FBaljudate		      = rsACADEMYget("baljudate")
			FOneItem.Fupcheconfirmdate    = rsACADEMYget("upcheconfirmdate")
			FOneItem.FCurrstate		      = rsACADEMYget("currstate")
			FOneItem.Fbeasongdate         = rsACADEMYget("beasongdate")

			FOneItem.FisUpcheBeasong      = rsACADEMYget("isUpcheBeasong")
			FOneItem.FSongjangno          = rsACADEMYget("songjangno")
			FOneItem.FSongjangdiv         = rsACADEMYget("songjangdiv")

            FOneItem.FCode                = rsACADEMYget("code")           '' for old version
            FOneItem.FState               = rsACADEMYget("state")          '' for old version
            FOneItem.Fipgodate            = rsACADEMYget("ipgodate")       '' for old version

            FOneItem.FMisendReason        = rsACADEMYget("code")
            FOneItem.FMisendState         = rsACADEMYget("state")
            FOneItem.FMisendipgodate      = rsACADEMYget("ipgodate")

            FOneItem.FisSendSMS           = rsACADEMYget("isSendSMS")
            FOneItem.FisSendEmail         = rsACADEMYget("isSendEmail")
            FOneItem.FisSendCall          = rsACADEMYget("isSendCall")
            FOneItem.Fbuyemail            = rsACADEMYget("buyemail")
            FOneItem.FbuyHp               = rsACADEMYget("buyHp")

            FOneItem.FrequestString       = db2Html(rsACADEMYget("reqstr"))
            FOneItem.Fitemlackno          = rsACADEMYget("itemlackno")
            FOneItem.FfinishString        = db2Html(rsACADEMYget("finishstr"))

            FOneItem.Fcompany_name        = db2Html(rsACADEMYget("company_name"))
            FOneItem.Fcompany_tel         = db2Html(rsACADEMYget("company_tel"))

            'FOneItem.Fsmallimage          = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemid) + "/" + rsACADEMYget("smallimage")
            FOneItem.Fsmallimage		  = imgFingers & "/diyitem/webimage/small/" + GetImageSubFolderByItemid(FOneItem.FItemid) + "/" + rsACADEMYget("smallimage")
        end if
        rsACADEMYget.Close
    end function


	public sub GetOneOrderMasterWithCS
		dim sqlStr,i
		sqlStr = " select top 1 m.orderserial, m.cancelyn, m.buyname, m.buyhp, m.buyemail from [db_academy].[dbo].tbl_academy_order_master m" + VbCrlf
		if FRectOrderSerial<>"" then
			sqlStr = sqlStr + " where m.orderserial='" + FRectOrderSerial + "'"
		else
			sqlStr = sqlStr + " where m.deliverno='" + FRectDeliveryNo + "'"
		end if
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		set FOneItem = new COrderMasterWithCSItem
		if Not rsACADEMYget.Eof then
			FOneItem.FOrderSerial = rsACADEMYget("orderserial")
			FOneItem.FCancelyn    = rsACADEMYget("cancelyn")

			FOneItem.Fbuyname    = db2Html(rsACADEMYget("buyname"))
			FOneItem.Fbuyhp    = rsACADEMYget("buyhp")
			FOneItem.Fbuyemail    = db2Html(rsACADEMYget("buyemail"))
		end if

		rsACADEMYget.Close
	end sub

	public sub GetOldMisendListMaster
		dim sqlStr, sqlStr1, sqlStr2, i

        '미입력(제한사항:31일이상 미처리된 주문은 잘못된 결과를 출력한다. 입금일을 31일 이내로 제한하므로 사실상 의미는 없다.)
        sqlStr1 = " select distinct top " + CStr(FPageSize) + " m.orderserial, m.buyname,m.ipkumdate,m.regdate, m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp, m.baljudate, m.subtotalprice, m.ipkumdiv, null as code, null as state, null as ipgodate, null as itemid, null as reqstr, null as finishstr "
        sqlStr1 = sqlStr1 + " from [db_academy].[dbo].tbl_academy_order_master m, [db_academy].[dbo].tbl_academy_order_detail d "
        sqlStr1 = sqlStr1 + " where 1 = 1 "
        sqlStr1 = sqlStr1 + " and m.orderserial=d.orderserial "
        sqlStr1 = sqlStr1 + " and m.orderserial not in (select orderserial from [db_academy].[dbo].tbl_academy_mibeasong_list where datediff(d,regdate,getdate())<31) "
        sqlStr1 = sqlStr1 + " and datediff(d,m.ipkumdate,getdate())<31 "
        sqlStr1 = sqlStr1 + " and m.cancelyn='N' "
        sqlStr1 = sqlStr1 + " and m.ipkumdiv<8 "
        sqlStr1 = sqlStr1 + " and m.ipkumdiv>4 "
        sqlStr1 = sqlStr1 + " and m.jumundiv<>9 "
        sqlStr1 = sqlStr1 + " and d.itemid<>0 "
        sqlStr1 = sqlStr1 + " and d.isupchebeasong<>'Y' "
        sqlStr1 = sqlStr1 + " and d.currstate<7"

		if FRectDelayDate <> "" and FRectDelayDate <> "0" then
			sqlStr1 = sqlStr1 + " and (datediff(d,m.baljudate,getdate())>=" + CStr(FRectDelayDate) + " ) "
		end if
		if FRectDeliveryNo <> "" then
			sqlStr1 = sqlStr1 + " and (m.deliverno = '" + FRectDeliveryNo + "' ) "
		end if

        ''입력완료
        sqlStr2 = " select distinct top " + CStr(FPageSize) + " m.orderserial, m.buyname,m.ipkumdate,m.regdate, m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp, m.baljudate, m.subtotalprice, m.ipkumdiv, l.code, l.state,l.ipgodate, l.itemid, l.reqstr, l.finishstr "
        sqlStr2 = sqlStr2 + " from [db_academy].[dbo].tbl_academy_order_master m, [db_academy].[dbo].tbl_academy_order_detail d, [db_academy].[dbo].tbl_academy_mibeasong_list l "
        sqlStr2 = sqlStr2 + " where 1 = 1 "
        sqlStr2 = sqlStr2 + " and m.orderserial=d.orderserial "
        sqlStr2 = sqlStr2 + " and d.idx=l.detailidx "
        ''sqlStr2 = sqlStr2 + " and datediff(d,m.ipkumdate,getdate())<31 "
        sqlStr2 = sqlStr2 + " and m.cancelyn='N' "
        sqlStr2 = sqlStr2 + " and m.ipkumdiv<8 "
        sqlStr2 = sqlStr2 + " and m.ipkumdiv>4 "
        sqlStr2 = sqlStr2 + " and m.jumundiv<>9 "
        sqlStr2 = sqlStr2 + " and d.itemid<>0 "
        sqlStr2 = sqlStr2 + " and d.isupchebeasong<>'Y' "
        sqlStr2 = sqlStr2 + " and d.currstate<7"

		if FRectDelayDate <> "" then
			sqlStr2 = sqlStr2 + " and (datediff(d,m.baljudate,getdate())>=" + CStr(FRectDelayDate) + " ) "
		end if
		if FRectDeliveryNo <> "" then
			sqlStr2 = sqlStr2 + " and (m.deliverno = '" + FRectDeliveryNo + "' ) "
		end if

		if FRectInCludeAlreadyInputed = "N" then
			sqlStr = sqlStr1
			sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		elseif FRectInCludeAlreadyInputed = "Y" then
		    sqlStr = sqlStr2
			sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		elseif FRectInCludeAlreadyInputed = "A" then
					'sqlStr2 = sqlStr2 + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
			sqlStr = " ((" + sqlStr1 + ") union (" + sqlStr2 + ")) "
		end if

		if FRectInCludeAlreadyInputed = "1" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='1' "
			sqlStr = sqlStr + " order by m.ipkumdate , m.orderserial  "
		elseif FRectInCludeAlreadyInputed = "2" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='2' "
			sqlStr = sqlStr + " order by m.ipkumdate , m.orderserial  "
		elseif FRectInCludeAlreadyInputed = "3" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='3' "
			sqlStr = sqlStr + " order by m.ipkumdate , m.orderserial  "
		elseif FRectInCludeAlreadyInputed = "6" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='6' "
			sqlStr = sqlStr + " order by m.ipkumdate , m.orderserial  "
		elseif FRectInCludeAlreadyInputed = "7" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='7' "
			sqlStr = sqlStr + " order by m.ipkumdate , m.orderserial  "
		elseif FRectInCludeAlreadyInputed = "36" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='6' "
			sqlStr = sqlStr + " order by m.ipkumdate , m.orderserial  "
		end if

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

'response.write sqlStr

		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsACADEMYget("orderserial")
				'FItemList(i).FMakerId        = rsACADEMYget("makerid")
				FItemList(i).FItemId         = rsACADEMYget("itemid")
				'FItemList(i).FItemName       = db2html(rsACADEMYget("itemname"))
				'FItemList(i).FItemOptionName = db2html(rsACADEMYget("itemoptionname"))
				'FItemList(i).FIsUpcheBeasong = rsACADEMYget("isupchebeasong")
				'FItemList(i).FCurrState      = rsACADEMYget("currstate")
				FItemList(i).FCode           = rsACADEMYget("code")
				FItemList(i).FState          = rsACADEMYget("state")
				FItemList(i).FIpgoDate       = rsACADEMYget("ipgodate")

				FItemList(i).FBuyName		 = rsACADEMYget("buyname")
				FItemList(i).FReqName		 = rsACADEMYget("reqname")
				FItemList(i).FIpkumDate		 = rsACADEMYget("ipkumdate")
				FItemList(i).FRegDate		 = rsACADEMYget("regdate")
				FItemList(i).FDeliveryNo	 = rsACADEMYget("deliverno")
				FItemList(i).FSiteName	     = rsACADEMYget("sitename")
				FItemList(i).FUserId	     = rsACADEMYget("userid")
				FItemList(i).FSubTotalPrice  = rsACADEMYget("subtotalprice")
				FItemList(i).Fipkumdiv       = rsACADEMYget("ipkumdiv")
				FItemList(i).Fbaljudate      = rsACADEMYget("baljudate")

				FItemList(i).FrequestString = rsACADEMYget("reqstr")
				FItemList(i).FfinishString = rsACADEMYget("finishstr")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end Sub

	public sub GetOldMisendListMasterCS
		dim sqlStr,i
		dim Before3month : Before3month = Left(CStr(DateAdd("m",-3,now())),10)

		sqlStr = " select  top " + CStr(FPageSize) + " m.orderserial,"
		sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno, d.isupchebeasong,d.currstate,d.beasongdate, d.cancelyn as DetailCancelYn,"
		sqlStr = sqlStr + " m.buyname,m.ipkumdate,m.regdate, m.baljudate,m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp, "
		sqlStr = sqlStr + " m.subtotalprice, m.ipkumdiv, l.code, l.state,l.ipgodate, l.itemid, l.reqstr, l.finishstr, l.ItemLackNo "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_master m "
		sqlStr = sqlStr + "     Join [db_academy].[dbo].tbl_academy_order_detail d "
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + "     join [db_academy].[dbo].tbl_academy_mibeasong_list l"
		sqlStr = sqlStr + "     on d.idx=l.detailidx"

		sqlStr = sqlStr + " where m.regdate>'"&Before3month&"'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv>'3'"
		sqlStr = sqlStr + " and m.ipkumdiv<'8'"
		sqlStr = sqlStr + " and m.jumundiv<>'9'"
		sqlStr = sqlStr + " and d.itemid<>0"
		if (FRectItemid<>"") then
		    sqlStr = sqlStr + " and d.itemid="&FRectItemid
		end if

		sqlStr = sqlStr + " and d.isupchebeasong='N'"
		sqlStr = sqlStr + " and d.currstate<7"              ''출고 이전
		''sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		'sqlStr = sqlStr + " and l.reqstr is not NULL "

		if FRectInCludeAlreadyInputed = "N" then
			''(l.reqstr <> '') or
			sqlStr = sqlStr + " and l.code<>'00'"
			sqlStr = sqlStr + " and l.state='0'"
		elseif FRectInCludeAlreadyInputed = "Y" then
			sqlStr = sqlStr + " and l.code is not null"
		elseif FRectInCludeAlreadyInputed = "4" then        '2009고객안내
		    sqlStr = sqlStr + " and l.state in ('1','2','3','4')"
		elseif FRectInCludeAlreadyInputed <> "" then
			sqlStr = sqlStr + " and l.state='"&FRectInCludeAlreadyInputed&"'"
		end if

		if FRectDeliveryNo <> "" then
			sqlStr = sqlStr + " and (m.deliverno = '" + FRectDeliveryNo + "' ) "
		end if
		if FRectOrderingOpt="itidasc" then
			sqlStr = sqlStr + " order by l.itemid "
		elseif FRectOrderingOpt ="itiddesc" then
			sqlStr = sqlStr + " order by l.itemid desc"
		elseif FRectOrderingOpt="cdasc" then
			sqlStr = sqlStr + " order by l.code"
		elseif FRectOrderingOpt="cddesc" then
			sqlStr = sqlStr + " order by l.code desc"
		else

		sqlStr = sqlStr + " order by m.ipkumdate, m.orderserial "
		end if


''response.write sqlStr
        rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenForwardOnly
		rsACADEMYget.LockType = adLockReadOnly

		rsACADEMYget.Open sqlStr,dbACADEMYget
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsACADEMYget("orderserial")
				'FItemList(i).FMakerId        = rsACADEMYget("makerid")
				FItemList(i).FItemId         = rsACADEMYget("itemid")
				FItemList(i).FItemName       = db2html(rsACADEMYget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsACADEMYget("itemoptionname"))
				FItemList(i).FItemNo 		 = rsACADEMYget("itemno")
				FItemList(i).FItemLackNo 	 = rsACADEMYget("itemLackNo")


				FItemList(i).FIsUpcheBeasong = rsACADEMYget("isupchebeasong")
				FItemList(i).FCurrState      = rsACADEMYget("currstate")
				FItemList(i).FCode           = rsACADEMYget("code")
				FItemList(i).FState          = rsACADEMYget("state")
				FItemList(i).FIpgoDate       = rsACADEMYget("ipgodate")

				FItemList(i).FBuyName		 = rsACADEMYget("buyname")
				FItemList(i).FBuyPhone		 = rsACADEMYget("buyphone")
				FItemList(i).FBuyHP		 = rsACADEMYget("buyhp")
				FItemList(i).FReqName		 = rsACADEMYget("reqname")
				FItemList(i).FIpkumDate		 = rsACADEMYget("ipkumdate")
				FItemList(i).FRegDate		 = rsACADEMYget("regdate")
				FItemList(i).FDeliveryNo	 = rsACADEMYget("deliverno")
				FItemList(i).FSiteName	 = rsACADEMYget("sitename")
				FItemList(i).FUserId	 = rsACADEMYget("userid")
				FItemList(i).FSubTotalPrice = rsACADEMYget("subtotalprice")
				FItemList(i).Fipkumdiv = rsACADEMYget("ipkumdiv")

				FItemList(i).FrequestString = rsACADEMYget("reqstr")
				FItemList(i).FfinishString = rsACADEMYget("finishstr")
                FItemList(i).FDetailCancelYn = rsACADEMYget("DetailCancelYn")
                FItemList(i).FBaljudate		      = rsACADEMYget("baljudate")
                FItemList(i).Fbeasongdate         = rsACADEMYget("beasongdate")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end Sub

	public sub GetOldMisendListALL
		dim sqlStr,i
		sqlStr = " select top " + CStr(FPageSize) + " m.orderserial,d.makerid,d.itemid,d.itemname,"
		sqlStr = sqlStr + " d.itemoptionname,d.isupchebeasong,d.currstate,"
		sqlStr = sqlStr + " m.buyname,m.ipkumdate,m.regdate, m.reqname, m.deliverno, m.sitename, m.userid, "
		sqlStr = sqlStr + " m.subtotalprice, m.ipkumdiv, l.code, l.state,l.ipgodate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_mibeasong_list l"
		sqlStr = sqlStr + " on d.idx=l.detailidx"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.idx>350000"
		sqlStr = sqlStr + " and datediff(m,m.ipkumdate,getdate())<2"
		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>" + CStr(FRectDelayDate)
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.oitemdiv<>'90'"
		if FRectNotIncludeItemList<>"" then
			sqlStr = sqlStr + " and i.itemid not in (" + FRectNotIncludeItemList + ")"
		end if

		if FRectNotInCludeUpcheCheck="on" then
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.currstate is NULL))"
		else
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.beasongdate is NULL))"
		end if

		sqlStr = sqlStr + "         or (d.isupchebeasong<>'Y' and m.ipkumdiv<6))"
		sqlStr = sqlStr + " order by d.idx "

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsACADEMYget("orderserial")
				FItemList(i).FMakerId        = rsACADEMYget("makerid")
				FItemList(i).FItemId         = rsACADEMYget("itemid")
				FItemList(i).FItemName       = db2html(rsACADEMYget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsACADEMYget("itemoptionname"))
				FItemList(i).FIsUpcheBeasong = rsACADEMYget("isupchebeasong")
				FItemList(i).FCurrState      = rsACADEMYget("currstate")
				FItemList(i).FCode           = rsACADEMYget("code")
				FItemList(i).FState          = rsACADEMYget("state")
				FItemList(i).FIpgoDate       = rsACADEMYget("ipgodate")

				FItemList(i).FBuyName		 = rsACADEMYget("buyname")
				FItemList(i).FReqName		 = rsACADEMYget("reqname")
				FItemList(i).FIpkumDate		 = rsACADEMYget("ipkumdate")
				FItemList(i).FRegDate		 = rsACADEMYget("regdate")
				FItemList(i).FDeliveryNo	 = rsACADEMYget("deliverno")
				FItemList(i).FSiteName	 = rsACADEMYget("sitename")
				FItemList(i).FUserId	 = rsACADEMYget("userid")
				FItemList(i).FSubTotalPrice = rsACADEMYget("subtotalprice")
				FItemList(i).Fipkumdiv = rsACADEMYget("ipkumdiv")



				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end Sub

	public sub GetOldMisendListSearch
		dim sqlStr,i
		sqlStr = " select top " + CStr(FPageSize) + " d.orderserial,d.makerid,d.itemid,d.itemname,"
		sqlStr = sqlStr + " d.itemoptionname,d.isupchebeasong,d.currstate,"
		sqlStr = sqlStr + " m.buyname,m.ipkumdate,"
		sqlStr = sqlStr + " l.code, l.state,l.ipgodate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_mibeasong_list l"
		sqlStr = sqlStr + " on d.idx=l.detailidx"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.idx>350000"
		sqlStr = sqlStr + " and datediff(m,m.ipkumdate,getdate())<2"
		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>" + CStr(FRectDelayDate)
		sqlStr = sqlStr + " and m.sitename<>'tingmart'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and i.itemdiv<50"

		if FRectNotInCludeUpcheCheck="on" then
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.currstate is NULL))"
		else
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.beasongdate is NULL))"
		end if

		sqlStr = sqlStr + "         or (d.isupchebeasong<>'Y' and m.ipkumdiv<6))"
		sqlStr = sqlStr + " order by d.idx "

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsACADEMYget("orderserial")
				FItemList(i).FMakerId        = rsACADEMYget("makerid")
				FItemList(i).FItemId         = rsACADEMYget("itemid")
				FItemList(i).FItemName       = db2html(rsACADEMYget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsACADEMYget("itemoptionname"))
				FItemList(i).FIsUpcheBeasong = rsACADEMYget("isupchebeasong")
				FItemList(i).FCurrState      = rsACADEMYget("currstate")
				FItemList(i).FCode           = rsACADEMYget("code")
				FItemList(i).FState          = rsACADEMYget("state")
				FItemList(i).FIpgoDate       = rsACADEMYget("ipgodate")

				FItemList(i).FBuyName		 = rsACADEMYget("buyname")
				FItemList(i).FIpkumDate		 = rsACADEMYget("ipkumdate")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end Sub

	public sub GetMiSendOrderByitemid()
		dim sqlStr,i
		sqlStr = " select top 500 m.idx, m.orderserial, m.buyname, m.reqname, m.ipkumdate, m.baljudate, d.itemno,"
		sqlStr = sqlStr + " m.regdate, m.buyphone, m.buyhp, m.deliverno, m.sitename, m.userid,"
		sqlStr = sqlStr + " m.subtotalprice, m.ipkumdiv, "
		sqlStr = sqlStr + " d.currstate, d.makerid, d.itemid, d.isupchebeasong, l.itemlackno, l.code, l.state, l.reqstr, l.finishstr"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m, "
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_mibeasong_list l"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv='5'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"

		if FRectIsupchebeasong = "N" then
			sqlStr = sqlStr + " and d.isupchebeasong='N'"
		elseif FRectIsupchebeasong = "Y" then
			sqlStr = sqlStr + " and d.isupchebeasong='Y'"
		end if

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and d.itemid=" + CStr(FRectItemid)
		end if

		sqlStr = sqlStr + " and d.idx=l.detailidx"
		sqlStr = sqlStr + " order by m.ipkumdate"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FItemList(rsACADEMYget.RecordCount)
		i=0
		do until rsACADEMYget.Eof
			set FItemList(i) = new COldMiSendItem
			FItemList(i).FOrderserial = rsACADEMYget("orderserial")
			FItemList(i).FMakerId     = rsACADEMYget("makerid")
			FItemList(i).FItemId         = rsACADEMYget("itemid")
			FItemList(i).FItemNo = rsACADEMYget("itemno")

			FItemList(i).Fbuyname   = db2html(rsACADEMYget("buyname"))
			FItemList(i).Freqname 	= db2html(rsACADEMYget("reqname"))
			FItemList(i).Fipkumdate = rsACADEMYget("ipkumdate")
			FItemList(i).Fbaljudate = rsACADEMYget("baljudate")
			FItemList(i).FRegDate        = rsACADEMYget("regdate")

			FItemList(i).FIsUpcheBeasong = rsACADEMYget("isupchebeasong")
			FItemList(i).FCurrState      = rsACADEMYget("currstate")
			FItemList(i).Fitemlackno	 = rsACADEMYget("itemlackno")

			FItemList(i).FCode           = rsACADEMYget("code")
			FItemList(i).FState          = rsACADEMYget("state")

			FItemList(i).FBuyPhone      = rsACADEMYget("buyphone")
			FItemList(i).FBuyHP         = rsACADEMYget("buyhp")

			FItemList(i).FDeliveryNo    = rsACADEMYget("deliverno")
			FItemList(i).FSiteName      = rsACADEMYget("sitename")
			FItemList(i).FUserId        = rsACADEMYget("userid")
			FItemList(i).FSubTotalPrice = rsACADEMYget("subtotalprice")
			FItemList(i).Fipkumdiv      = rsACADEMYget("ipkumdiv")

			FItemList(i).FrequestString = rsACADEMYget("reqstr")
			FItemList(i).FfinishString  = rsACADEMYget("finishstr")


			i=i+1
			rsACADEMYget.MoveNext
		loop
		rsACADEMYget.close
	end sub

	Private Sub Class_Initialize()
	redim FItemList(0)
		FRectDelayDate = 5
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
%>