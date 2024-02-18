<%

function MiSendCodeToColor(code)
	if code="05" then
		MiSendCodeToColor = "#FF0000"
	else
		MiSendCodeToColor = "#000000"
	end if
end function

function MiSendCodeToName(code)
	if code="00" then
		MiSendCodeToName = "입력대기"
	elseif code="03" then
		MiSendCodeToName = "출고지연"
	elseif code="06" then
		MiSendCodeToName = "택배파업(취소)"
	elseif code="02" then
		MiSendCodeToName = "주문제작"
	elseif code="08" then
		MiSendCodeToName = "수입"
	elseif code="09" then
		MiSendCodeToName = "가구배송"
	elseif code="04" then
		MiSendCodeToName = "예약배송"
	elseif code="10" then
		MiSendCodeToName = "업체휴가"
	elseif code="07" then
		MiSendCodeToName = "고객지정배송" ''2011-05추가
	elseif code="05" then
		MiSendCodeToName = "품절출고불가"
	elseif code="66" then
		MiSendCodeToName = "가격오류"
	elseif code="11" then
		MiSendCodeToName = "업체확인중"
	else
		MiSendCodeToName = code
	end if
end function

Class COrderMasterWithCSItem
	public FOrderSerial
	public FCancelyn
    public Fbuyname
    public Fbuyhp
    public Fbuyemail
	public Fsitename

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
	public FReqPhone
	public FReqHP
	public FReqName
	public FIpkumDate

	public FDeliveryNo
	public FSiteName
	public FUserId
	public FSubTotalPrice
	public Fipkumdiv
	public Fbaljudate

	public FrequestString			'// 물류 요청내용
	public FupcheRequestString		'// 업체 요청내용
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

	public FPrevMisendReason
    public FMisendReason
    public FMisendState
    public FMisendipgodate

    public FisSendSMS
    public FisSendEmail
    public FisSendCall
	public FisSendAPI

    public Fcompany_name
    public Fcompany_tel
    public Fsmallimage
    public FCancelYn
    public FDetailCancelYn
    public Fdetailidx

	public FisMakeOnOrderItem		'// 주문제작상품
	public FisMakeOnOrderOrgItem	'// 원상품(주문제작상품)
	public Frequiredetail

	public FMiRegDate
	public FMiRegUserid

	public Freqreguserid

	public Fbuyaddr1
	public Fbuyaddr2

	public Freqzipaddr
	public Freqaddress

	public FpojangName
	public FboxNo
	public FpickupReqDate
	public Fcomment
	public Fsongjangprintdate

	public FDday

    public function getNewBeasongDPlusDateStr()
        getNewBeasongDPlusDateStr = "D+" & FDday
    end function

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

            elseif (FMisendReason="06") then	' 택배파업

            elseif (FMisendReason="02") then  ''주문제작(수입)
                ''출고 소요일수 D+2이상
                if (getMisendDPlusDate>1) then
                    smstext = "[텐바이텐 출고지연안내]주문하신 상품중 "&DdotFormat(FItemName,32)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"주문제작(수입) 상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,32)&"("&FItemID&")상품이 "&VbCrlf
                    smstext = smstext&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."
                end if
            elseif (FMisendReason="03") then  ''출고지연
                ''출고 소요일수 D+2이상
                if (getMisendDPlusDate>1) then
                    smstext = "[텐바이텐 출고지연안내]주문하신 상품중 "&DdotFormat(FItemName,32)&"("&FItemID&")상품이 "&VbCrlf
                    smstext = smstext&FMisendipgodate&"에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,32)&"("&FItemID&")상품이 "&VbCrlf
                    smstext = smstext&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."

                end if
            elseif (FMisendReason="04") then  ''예약상품
                ''출고 소요일수 D+2이상
                if (getMisendDPlusDate>1) then
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,32)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"예약배송상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,32)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"예약배송상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."

                end if
            elseif (FMisendReason="07") then  ''고객지정배송
                ''출고 소요일수 D+2이상
                if (getMisendDPlusDate>1) then
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,32)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"고객지정배송상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,32)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"고객지정배송상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."

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
            getDlvCompanyName = "텐바이텐"
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
		getMiSendCodeColor = MiSendCodeToColor(FMisendReason)
	end function

	public function getMiSendCodeName()
		getMiSendCodeName = MiSendCodeToName(FCode)
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
	public FRectLackItemOnly

	public FRectNotIncludeItemList
	public FRectOrderSerial

    public FRectMakerid
	public FRectItemId
    public FRectItemOption
	public FRectIsupchebeasong
	public FRectDetailidx
    public FRectSiteName

	public FRectBaljuCode

	public FRectStartDate
	public FRectEndDate

	public FRectForMail
	public frectdetailcancelyn

	''주문내역중 미배송리스트 / 미배송 없는내역도 조회.
	public function getMiSendOrderDetailList()
        dim sqlStr, i

        sqlStr = "exec [db_order].[dbo].sp_Ten_Mibeasong_Item_GetList '" + CStr(FRectOrderSerial) + "', '"& frectdetailcancelyn &"'"

		'response.write sqlStr & "<br>"
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		i=0
		redim FItemList(FResultCount)
		if not rsget.EOF then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

    			FItemList(i).Fidx				  = rsget("idx")
    			FItemList(i).FOrderserial		  = rsget("orderserial")
    			FItemList(i).FItemid 			  = rsget("itemid")
    			FItemList(i).FItemoption     	  = rsget("itemoption")
    			FItemList(i).FItemname 		      = db2html(rsget("itemname"))
    			FItemList(i).FItemoptionName      = db2html(rsget("itemoptionname"))
    			FItemList(i).FItemcnt             = rsget("itemno")

    			FItemList(i).FMakerid 			  = rsget("makerid")
    			FItemList(i).FBuyname             = db2html(rsget("buyname"))
    			FItemList(i).FReqname			  = db2html(rsget("reqname"))
    			FItemList(i).FCancelYn		      = rsget("cancelyn")
    			FItemList(i).FDetailCancelYn	  = rsget("detailcancelyn")
				FItemList(i).FRegdate			  = rsget("regdate")
    			FItemList(i).FIpkumdate		      = rsget("ipkumdate")
    			FItemList(i).FBaljudate		      = rsget("baljudate")
    			FItemList(i).Fupcheconfirmdate    = rsget("upcheconfirmdate")
    			FItemList(i).FCurrstate		      = rsget("currstate")      '' DetailState

    			FItemList(i).Fbeasongdate         = rsget("beasongdate")

    			FItemList(i).FisUpcheBeasong      = rsget("isUpcheBeasong")
    			FItemList(i).FSongjangno          = rsget("songjangno")
    			FItemList(i).FSongjangdiv         = rsget("songjangdiv")

                FItemList(i).FCode                = rsget("code")           '' for old version
                FItemList(i).FState               = rsget("state")          '' for old version
                FItemList(i).Fipgodate            = rsget("ipgodate")       '' for old version

                FItemList(i).FPrevMisendReason    = rsget("prevcode")
				FItemList(i).FMisendReason        = rsget("code")
                FItemList(i).FMisendState         = rsget("state")
                FItemList(i).FMisendipgodate      = rsget("ipgodate")

                FItemList(i).FisSendSMS           = rsget("isSendSMS")
                FItemList(i).FisSendEmail         = rsget("isSendEmail")
                FItemList(i).FisSendCall          = rsget("isSendCall")
				FItemList(i).FisSendAPI			  = rsget("isSendAPI")
                FItemList(i).Fbuyemail            = rsget("buyemail")
                FItemList(i).FbuyHp               = rsget("buyHp")

                FItemList(i).FrequestString       = db2Html(rsget("reqstr"))
				FItemList(i).FupcheRequestString  = db2Html(rsget("reqaddstr"))

                FItemList(i).FItemNo              = rsget("itemno")
                FItemList(i).Fitemlackno          = rsget("itemlackno")
                FItemList(i).FfinishString        = db2Html(rsget("finishstr"))


                FItemList(i).Fcompany_name        = db2Html(rsget("company_name"))
                FItemList(i).Fcompany_tel         = db2Html(rsget("company_tel"))

                FItemList(i).Fsmallimage          = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsget("smallimage")

                FItemList(i).FCancelYn            = rsget("detailcancelyn")

				FItemList(i).FMiRegDate           = rsget("miregdate")
				FItemList(i).FMiRegUserid         = rsget("mireguserid")

				FItemList(i).Freqreguserid         = rsget("reqreguserid")
				FItemList(i).FDday      		   = rsget("dday")

                i=i+1
                rsget.MoveNext
            loop

        end if
        rsget.Close
    end function

    public function getOneOldMisendItem()
        dim sqlStr
        sqlStr = "exec [db_order].[dbo].sp_Ten_Mibeasong_Item_GetData " + CStr(FRectDetailidx) + ""
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		if not rsget.EOF then
            set FOneItem = new COldMiSendItem

			FOneItem.Fidx				  = rsget("idx")
			FOneItem.FOrderserial		  = rsget("orderserial")
			FOneItem.FItemid 			  = rsget("itemid")
			FOneItem.FItemoption     	  = rsget("itemoption")
			FOneItem.FItemname 		      = db2html(rsget("itemname"))
			FOneItem.FItemoptionName      = db2html(rsget("itemoptionname"))
			FOneItem.FItemcnt             = rsget("itemno")

			FOneItem.FMakerid 			  = rsget("makerid")
			FOneItem.FBuyname             = db2html(rsget("buyname"))
			FOneItem.FReqname			  = db2html(rsget("reqname"))
			FOneItem.FUserID              = rsget("userid")

			FOneItem.FCancelYn		      = rsget("cancelyn")  ''master cancelyn
			FOneItem.FDetailCancelYn		      = rsget("detailcancelyn")  ''detailcancelyn
			FOneItem.FRegdate			  = rsget("regdate")
			FOneItem.FIpkumdate		      = rsget("ipkumdate")
			FOneItem.FBaljudate		      = rsget("baljudate")
			FOneItem.Fupcheconfirmdate    = rsget("upcheconfirmdate")
			FOneItem.FCurrstate		      = rsget("currstate")
			FOneItem.Fbeasongdate         = rsget("beasongdate")

			FOneItem.FisUpcheBeasong      = rsget("isUpcheBeasong")
			FOneItem.FSongjangno          = rsget("songjangno")
			FOneItem.FSongjangdiv         = rsget("songjangdiv")

            FOneItem.FCode                = rsget("code")           '' for old version
            FOneItem.FState               = rsget("state")          '' for old version
            FOneItem.Fipgodate            = rsget("ipgodate")       '' for old version

            FOneItem.FMisendReason        = rsget("code")
            FOneItem.FMisendState         = rsget("state")
            FOneItem.FMisendipgodate      = rsget("ipgodate")

            FOneItem.FisSendSMS           = rsget("isSendSMS")
            FOneItem.FisSendEmail         = rsget("isSendEmail")
            FOneItem.FisSendCall          = rsget("isSendCall")
            FOneItem.Fbuyemail            = rsget("buyemail")
            FOneItem.FbuyHp               = rsget("buyHp")

            FOneItem.FrequestString       = db2Html(rsget("reqstr"))
			FOneItem.FupcheRequestString  = db2Html(rsget("reqaddstr"))

            FOneItem.Fitemlackno          = rsget("itemlackno")
            FOneItem.FfinishString        = db2Html(rsget("finishstr"))

            FOneItem.Fcompany_name        = db2Html(rsget("company_name"))
            FOneItem.Fcompany_tel         = db2Html(rsget("company_tel"))

			if (FRectForMail = "Y") then
				FOneItem.Fsmallimage          = webImgUrlForMAIL + "/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemid) + "/" + rsget("smallimage")
			else
				FOneItem.Fsmallimage          = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemid) + "/" + rsget("smallimage")
			end if
        end if
        rsget.Close
    end function


	public sub GetOneOrderMasterWithCS
		dim sqlStr,i
		sqlStr = " select top 1 m.orderserial, m.cancelyn, m.buyname, m.buyhp, m.buyemail, m.sitename from [db_order].[dbo].tbl_order_master m" + VbCrlf
		if FRectOrderSerial<>"" then
			sqlStr = sqlStr + " where m.orderserial='" + FRectOrderSerial + "'"
		else
			sqlStr = sqlStr + " where m.deliverno='" + FRectDeliveryNo + "'"
		end if
		rsget.Open sqlStr,dbget,1

		set FOneItem = new COrderMasterWithCSItem
		if Not rsget.Eof then
			FOneItem.FOrderSerial 	= rsget("orderserial")
			FOneItem.FCancelyn    	= rsget("cancelyn")
			FOneItem.Fbuyname    	= db2Html(rsget("buyname"))
			FOneItem.Fbuyhp    		= rsget("buyhp")
			FOneItem.Fbuyemail    	= db2Html(rsget("buyemail"))
			FOneItem.Fsitename    	= rsget("sitename")
		end if
		rsget.Close
	end sub

	public sub GetOldMisendListMaster
		dim sqlStr, sqlStr1, sqlStr2, i

        '미입력(제한사항:31일이상 미처리된 주문은 잘못된 결과를 출력한다. 입금일을 31일 이내로 제한하므로 사실상 의미는 없다.)
        sqlStr1 = " select distinct top " + CStr(FPageSize) + " m.orderserial, m.buyname,m.ipkumdate,m.regdate, m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp, m.baljudate, m.subtotalprice, m.ipkumdiv, null as code, null as state, null as ipgodate, null as itemid, null as reqstr, null as finishstr "
        sqlStr1 = sqlStr1 + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d "
        sqlStr1 = sqlStr1 + " where 1 = 1 "
        sqlStr1 = sqlStr1 + " and m.orderserial=d.orderserial "
        sqlStr1 = sqlStr1 + " and m.orderserial not in (select orderserial from [db_temp].[dbo].tbl_mibeasong_list where datediff(d,regdate,getdate())<31) "
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
        sqlStr2 = sqlStr2 + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d, [db_temp].[dbo].tbl_mibeasong_list l "
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

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

'response.write sqlStr

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")
				'FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				'FItemList(i).FItemName       = db2html(rsget("itemname"))
				'FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				'FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
				'FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FCode           = rsget("code")
				FItemList(i).FState          = rsget("state")
				FItemList(i).FIpgoDate       = rsget("ipgodate")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FReqName		 = rsget("reqname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FDeliveryNo	 = rsget("deliverno")
				FItemList(i).FSiteName	     = rsget("sitename")
				FItemList(i).FUserId	     = rsget("userid")
				FItemList(i).FSubTotalPrice  = rsget("subtotalprice")
				FItemList(i).Fipkumdiv       = rsget("ipkumdiv")
				FItemList(i).Fbaljudate      = rsget("baljudate")

				FItemList(i).FrequestString = rsget("reqstr")
				FItemList(i).FfinishString = rsget("finishstr")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetOldMisendListMasterCS
		dim sqlStr,i
		dim Before3month, tmp
		IF (application("Svr_Info")	= "Dev") then
		    Before3month = Left(CStr(DateAdd("m",-200,now())),10)
		ELSE
		    Before3month = Left(CStr(DateAdd("m",-3,now())),10)
	    END IF

		sqlStr = " select  top " + CStr(FPageSize) + " m.orderserial"
		sqlStr = sqlStr + " ,d.itemname, d.itemoptionname, d.itemno, d.isupchebeasong,d.currstate,d.beasongdate, d.cancelyn as DetailCancelYn"
		sqlStr = sqlStr + " ,m.buyname,m.ipkumdate,m.regdate, m.baljudate,m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp "
		sqlStr = sqlStr + " ,m.subtotalprice, m.ipkumdiv, l.code, l.state,l.ipgodate, l.itemid, l.reqstr, l.finishstr, l.ItemLackNo "
		sqlStr = sqlStr + " ,m.cancelyn, l.detailidx, d.makerid "
		sqlStr = sqlStr + " , (select db_order.[dbo].[UF_GetDPlusDateStr] ('T', convert(varchar(10), m.baljudate, 23), convert(varchar(10), isnull(d.beasongdate, getdate()), 23))) as dday "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m "
		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + "     join [db_temp].[dbo].tbl_mibeasong_list l"
		sqlStr = sqlStr + "     on d.idx=l.detailidx and d.orderserial=l.orderserial" ''and d.orderserial=l.orderserial 추가

		sqlStr = sqlStr + " where m.regdate>'"&Before3month&"'"
		if (FRectInCludeAlreadyInputed <> "C") then
		    sqlStr = sqlStr + " and m.cancelyn='N'"
	    end if
		sqlStr = sqlStr + " and m.ipkumdiv>'3'"
		sqlStr = sqlStr + " and m.ipkumdiv<'8'"
		sqlStr = sqlStr + " and m.jumundiv<>'9'"
		sqlStr = sqlStr + " and d.itemid<>0"
		if (FRectItemid<>"") then
		    sqlStr = sqlStr + " and d.itemid="&FRectItemid
		end if
        if (FRectItemOption<>"") then
		    sqlStr = sqlStr + " and d.itemoption='" & FRectItemOption & "' "
		end if
        if (FRectSiteName<>"") then
            if (FRectSiteName="NOTTEN") then
                sqlStr = sqlStr + " and m.sitename<>'10x10'"
            else
                sqlStr = sqlStr + " and m.sitename='"&FRectSiteName&"'"
            end if
        end if
		sqlStr = sqlStr + " and d.isupchebeasong='N'"
		sqlStr = sqlStr + " and d.currstate<7"              ''출고 이전
		if (FRectMakerid <> "") then
			sqlStr = sqlStr + " and d.makerid = '" & FRectMakerid & "' "
		end if
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
		elseif FRectInCludeAlreadyInputed = "C" then
		    sqlStr = sqlStr + " and ((d.cancelyn='Y') or (m.cancelyn='Y'))"
		    sqlStr = sqlStr + " and l.state<>9"
		elseif FRectInCludeAlreadyInputed <> "" then
			sqlStr = sqlStr + " and l.state='"&FRectInCludeAlreadyInputed&"'"
		end if

		if FRectDeliveryNo <> "" then
			sqlStr = sqlStr + " and (m.deliverno = '" + FRectDeliveryNo + "' ) "
		end if

		if FRectLackItemOnly = "Y" then
			sqlStr = sqlStr + " and isnull(l.itemlackno,0) > 0 "
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


''rw sqlStr
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.Open sqlStr,dbget
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

        tmp = ""
        FTotalCount = 0
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")

                if (tmp <> FItemList(i).FOrderSerial) then
                    tmp = FItemList(i).FOrderSerial
                    FTotalCount = FTotalCount + 1
                end if

				FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo 		 = rsget("itemno")
				FItemList(i).FItemLackNo 	 = rsget("itemLackNo")


				FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FCode           = rsget("code")
				FItemList(i).FState          = rsget("state")
				FItemList(i).FIpgoDate       = rsget("ipgodate")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FBuyPhone		 = rsget("buyphone")
				FItemList(i).FBuyHP		 = rsget("buyhp")
				FItemList(i).FReqName		 = rsget("reqname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FDeliveryNo	 = rsget("deliverno")
				FItemList(i).FSiteName	 = rsget("sitename")
				FItemList(i).FUserId	 = rsget("userid")
				FItemList(i).FSubTotalPrice = rsget("subtotalprice")
				FItemList(i).Fipkumdiv = rsget("ipkumdiv")

				FItemList(i).FrequestString = rsget("reqstr")
				FItemList(i).FfinishString = rsget("finishstr")
                FItemList(i).FDetailCancelYn = rsget("DetailCancelYn")
                FItemList(i).FBaljudate		      = rsget("baljudate")
                FItemList(i).Fbeasongdate         = rsget("beasongdate")
                FItemList(i).FCancelYn            = rsget("CancelYn")
                FItemList(i).Fdetailidx           = rsget("detailidx")
				FItemList(i).FDday           = rsget("dday")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetOldMisendListMasterCS_ipkumdiv4
		dim sqlStr,i
		dim Before3month, tmp
		IF (application("Svr_Info")	= "Dev") then
		    Before3month = Left(CStr(DateAdd("m",-200,now())),10)
		ELSE
		    Before3month = Left(CStr(DateAdd("m",-3,now())),10)
	    END IF

		sqlStr = " select  top " + CStr(FPageSize) + " m.orderserial"
		sqlStr = sqlStr + " ,d.itemname, d.itemoptionname, d.itemno, d.isupchebeasong,d.currstate,d.beasongdate, d.cancelyn as DetailCancelYn"
		sqlStr = sqlStr + " ,m.buyname,m.ipkumdate,m.regdate, m.baljudate,m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp "
		sqlStr = sqlStr + " ,m.subtotalprice, m.ipkumdiv, l.code, l.state,l.ipgodate, l.itemid, l.reqstr, l.finishstr, l.ItemLackNo "
		sqlStr = sqlStr + " ,m.cancelyn, l.detailidx, d.makerid "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m "
		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
		sqlStr = sqlStr + "     left join [db_temp].[dbo].tbl_mibeasong_list_ipkumdiv4 l"
		sqlStr = sqlStr + "     on d.idx=l.detailidx and d.orderserial=l.orderserial" ''and d.orderserial=l.orderserial 추가

		sqlStr = sqlStr + " where m.regdate>'"&Before3month&"'"
		if (FRectInCludeAlreadyInputed <> "C") then
		    sqlStr = sqlStr + " and m.cancelyn='N'"
	    end if
		sqlStr = sqlStr + " and m.ipkumdiv = '4' "
		sqlStr = sqlStr + " and m.jumundiv<>'9'"
		sqlStr = sqlStr + " and d.itemid<>0"
		if (FRectItemid<>"") then
		    sqlStr = sqlStr + " and d.itemid="&FRectItemid
		end if
        if (FRectItemOption<>"") then
		    sqlStr = sqlStr + " and d.itemoption='" & FRectItemOption & "' "
		end if
        if (FRectSiteName<>"") then
            if (FRectSiteName="NOTTEN") then
                sqlStr = sqlStr + " and m.sitename<>'10x10'"
            else
                sqlStr = sqlStr + " and m.sitename='"&FRectSiteName&"'"
            end if
        end if
		sqlStr = sqlStr + " and d.isupchebeasong='N'"
		sqlStr = sqlStr + " and d.currstate = '0' "              ''결제완료
		if (FRectMakerid <> "") then
			sqlStr = sqlStr + " and d.makerid = '" & FRectMakerid & "' "
		end if
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
		elseif FRectInCludeAlreadyInputed = "C" then
		    sqlStr = sqlStr + " and ((d.cancelyn='Y') or (m.cancelyn='Y'))"
		    sqlStr = sqlStr + " and l.state<>9"
		elseif FRectInCludeAlreadyInputed <> "" then
			sqlStr = sqlStr + " and l.state='"&FRectInCludeAlreadyInputed&"'"
		end if

		if FRectDeliveryNo <> "" then
			sqlStr = sqlStr + " and (m.deliverno = '" + FRectDeliveryNo + "' ) "
		end if

		if FRectLackItemOnly = "Y" then
			sqlStr = sqlStr + " and isnull(l.itemlackno,0) > 0 "
        elseif FRectLackItemOnly = "N" then
            sqlStr = sqlStr + " and l.itemlackno is NULL "
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


''rw sqlStr
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.Open sqlStr,dbget
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

        tmp = ""
        FTotalCount = 0
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")

                if (tmp <> FItemList(i).FOrderSerial) then
                    tmp = FItemList(i).FOrderSerial
                    FTotalCount = FTotalCount + 1
                end if

				FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo 		 = rsget("itemno")
				FItemList(i).FItemLackNo 	 = rsget("itemLackNo")


				FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FCode           = rsget("code")
				FItemList(i).FState          = rsget("state")
				FItemList(i).FIpgoDate       = rsget("ipgodate")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FBuyPhone		 = rsget("buyphone")
				FItemList(i).FBuyHP		 = rsget("buyhp")
				FItemList(i).FReqName		 = rsget("reqname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FDeliveryNo	 = rsget("deliverno")
				FItemList(i).FSiteName	 = rsget("sitename")
				FItemList(i).FUserId	 = rsget("userid")
				FItemList(i).FSubTotalPrice = rsget("subtotalprice")
				FItemList(i).Fipkumdiv = rsget("ipkumdiv")

				FItemList(i).FrequestString = rsget("reqstr")
				FItemList(i).FfinishString = rsget("finishstr")
                FItemList(i).FDetailCancelYn = rsget("DetailCancelYn")
                FItemList(i).FBaljudate		      = rsget("baljudate")
                FItemList(i).Fbeasongdate         = rsget("beasongdate")
                FItemList(i).FCancelYn            = rsget("CancelYn")
                FItemList(i).Fdetailidx           = rsget("detailidx")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	'// 주문제작(텐배) 발주리스트
	public sub GetBaljuListMakeOnOrder
		dim sqlStr,i
		dim Before3month
		if (application("Svr_Info")	= "Dev") then
		    Before3month = Left(CStr(DateAdd("m",-20,now())),10)
		else
		    Before3month = Left(CStr(DateAdd("m",-3,now())),10)
	    end if

		sqlStr = " select top 500 m.orderserial, m.sitename, m.buyname, m.buyphone, m.buyhp, m.buyemail, m.reqname, m.userid"
		sqlStr = sqlStr + " , d.idx as detailidx, d.itemid, d.itemname, d.itemoptionname, d.itemno, d.isupchebeasong,d.currstate"
		sqlStr = sqlStr & " ,d.beasongdate, isnull(dd.requiredetailUTF8,d.requiredetail) as requiredetail, m.ipkumdate, m.regdate "
		sqlStr = sqlStr + " , (case when d.cancelyn<>'Y' and d.oitemdiv = '06' and d.isupchebeasong = 'N' then 1 else 0 end) as ismakeonorderitem "


		sqlStr = sqlStr + " , (select count(*) from "
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail p "
		sqlStr = sqlStr + " 	, [db_order].[dbo].tbl_order_detail o "
		sqlStr = sqlStr + " 	, db_item.dbo.tbl_PlusSaleLinkItemList l "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and p.orderserial = m.orderserial "
		sqlStr = sqlStr + " 	and p.oitemdiv = '06' and p.isupchebeasong = 'N' "
		sqlStr = sqlStr + " 	and o.orderserial = m.orderserial "
		sqlStr = sqlStr + " 	and o.itemid = d.itemid "
		sqlStr = sqlStr + " 	and l.plusSaleItemID = p.itemid "
		sqlStr = sqlStr + " 	and l.plusSaleLinkItemID = o.itemid) as ismakeonorderorgitem "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + " join [db_order].[dbo].tbl_order_master m "
		sqlStr = sqlStr + " on d.orderserial = m.orderserial "
		sqlStr = sqlStr + " join [db_order].[dbo].tbl_baljudetail bd "
		sqlStr = sqlStr + " on m.orderserial = bd.orderserial "
		sqlStr = sqlStr + " join [db_order].[dbo].tbl_baljumaster bm "
		sqlStr = sqlStr + " on bm.id = bd.baljuid "
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_require dd" & vbcrlf
		sqlStr = sqlStr & "     ON d.idx = dd.detailidx" & vbcrlf
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and d.itemid <> 0 "
		sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		if (FRectBaljuCode <> "") then
			sqlStr = sqlStr + " 	and bd.baljuid = " + CStr(FRectBaljuCode) + " "
		else
			sqlStr = sqlStr + " 	and bm.baljudate >= '" + CStr(FRectStartDate) + "' "
			sqlStr = sqlStr + " 	and bm.baljudate < '" + CStr(FRectEndDate) + "' "
			sqlStr = sqlStr + " 	and (select count(*) from [db_order].[dbo].tbl_order_detail dd where m.orderserial = dd.orderserial and dd.cancelyn <> 'Y' and dd.isupchebeasong = 'N' and oitemdiv = '06') > 0 "
		end if
		''sqlStr = sqlStr + " 	and d.currstate < '7' "                             ''서동석 주석처리
		''sqlStr = sqlStr + " 	and d.beasongdate is NULL "
		sqlStr = sqlStr + " 	and d.isupchebeasong = 'N' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	bd.baljuid, m.orderserial "
		sqlStr = sqlStr + " 	, (CASE WHEN d.oitemdiv = '06' then 999 else 0 end)"
		'sqlStr = sqlStr + " 	, (select count(*) "
		'sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_detail d "
		'sqlStr = sqlStr + "		where d.itemid in (select itemid from db_item.dbo.tbl_item where deliverytype not in (2, 6, 7, 9) and itemdiv = '06' and isusing = 'Y' and regdate >= '2013-01-01')) "

''rw sqlStr
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.Open sqlStr,dbget
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    	= rsget("orderserial")
				'FItemList(i).FMakerId        	= rsget("makerid")
				FItemList(i).FItemId         	= rsget("itemid")
				FItemList(i).FItemName       	= db2html(rsget("itemname"))
				FItemList(i).FItemOptionName 	= db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo 		 	= rsget("itemno")

				FItemList(i).FIsUpcheBeasong 	= rsget("isupchebeasong")
				FItemList(i).FCurrState      	= rsget("currstate")

				FItemList(i).FBuyName		 	= rsget("buyname")
				FItemList(i).FBuyPhone		 	= rsget("buyphone")
				FItemList(i).FBuyHP		 		= rsget("buyhp")
				FItemList(i).FReqName		 	= rsget("reqname")
				FItemList(i).FIpkumDate		 	= rsget("ipkumdate")
				FItemList(i).FRegDate		 	= rsget("regdate")
				FItemList(i).FSiteName	 		= rsget("sitename")
				FItemList(i).FUserId	 		= rsget("userid")

                FItemList(i).Fbeasongdate       = rsget("beasongdate")
                FItemList(i).Fdetailidx         = rsget("detailidx")

				FItemList(i).FisMakeOnOrderOrgItem	= rsget("ismakeonorderorgitem") > 0
				FItemList(i).FisMakeOnOrderItem     = rsget("ismakeonorderitem") > 0
				FItemList(i).Frequiredetail         = rsget("requiredetail")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	'// 퀵배송 출고완료리스트
	public sub GetBaljuListQuickOrder
		dim sqlStr,i
		dim tmpStr

		sqlStr = " SELECT TOP 500 "
		sqlStr = sqlStr + " 	m.orderserial "
		sqlStr = sqlStr + " 	,m.buyname "
		sqlStr = sqlStr + " 	,m.buyphone "

		sqlStr = sqlStr + " 	, '경기도 포천시 군내면 용정경제로2길 83' as buyaddr1 "
		sqlStr = sqlStr + " 	, '텐바이텐 물류센터' as buyaddr2 "

		sqlStr = sqlStr + " 	,m.reqname "
		sqlStr = sqlStr + " 	,m.reqhp "
		sqlStr = sqlStr + " 	,m.reqphone "
		sqlStr = sqlStr + " 	,m.reqzipaddr "
		sqlStr = sqlStr + " 	,m.reqaddress "
		sqlStr = sqlStr + " 	, max(IsNull([db_item].[dbo].[getCateCodeFullDepthName](c.catecode), d.itemname)) as itemname "
		sqlStr = sqlStr + " 	, count(distinct d.itemid) as itemCnt "
		sqlStr = sqlStr + " 	, '소박스' as pojangName "
		sqlStr = sqlStr + " 	, 1 as boxNo "
		sqlStr = sqlStr + " 	, bd.songjangprintdate + ' 13:00' as pickupReqDate "
		sqlStr = sqlStr + " 	, m.comment "
		sqlStr = sqlStr + " 	, bd.songjangprintdate "
		sqlStr = sqlStr + " FROM [db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + " INNER JOIN [db_order].[dbo].tbl_order_master m ON d.orderserial = m.orderserial "
		sqlStr = sqlStr + " INNER JOIN [db_order].[dbo].tbl_baljudetail bd ON m.orderserial = bd.orderserial "
		sqlStr = sqlStr + " left join [db_item].[dbo].[tbl_display_cate_item] c on d.itemid = c.itemid and c.isDefault = 'Y' "
		sqlStr = sqlStr + " WHERE 1 = 1 "
		sqlStr = sqlStr + " 	AND d.itemid not in (0, 100) "
		sqlStr = sqlStr + " 	AND m.cancelyn = 'N' "
		sqlStr = sqlStr + " 	AND d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	AND d.currstate = '7' "
		sqlStr = sqlStr + " 	and bd.songjangprintdate >= '" + CStr(FRectStartDate) + "' "
		sqlStr = sqlStr + " 	and bd.songjangprintdate < '" + CStr(FRectEndDate) + "' "
		sqlStr = sqlStr + " 	AND d.isupchebeasong = 'N' "
		sqlStr = sqlStr + " group by "
		sqlStr = sqlStr + " 	m.orderserial "
		sqlStr = sqlStr + " 	,m.buyname "
		sqlStr = sqlStr + " 	,m.buyphone "
		sqlStr = sqlStr + " 	,m.reqname "
		sqlStr = sqlStr + " 	,m.reqhp "
		sqlStr = sqlStr + " 	,m.reqphone "
		sqlStr = sqlStr + " 	,m.reqzipaddr "
		sqlStr = sqlStr + " 	,m.reqaddress "
		sqlStr = sqlStr + " 	, bd.songjangprintdate "
		sqlStr = sqlStr + " 	, m.comment "
		sqlStr = sqlStr + " 	, bd.songjangprintdate "
		sqlStr = sqlStr + " ORDER BY m.orderserial "
		''rw sqlStr

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.Open sqlStr,dbget
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial  	= rsget("orderserial")
				FItemList(i).Fbuyname    	= db2html(rsget("buyname"))
				FItemList(i).Fbuyphone    	= db2html(rsget("buyphone"))
				FItemList(i).Fbuyaddr1    	= db2html(rsget("buyaddr1"))
				FItemList(i).Fbuyaddr2    	= db2html(rsget("buyaddr2"))
				FItemList(i).Freqname    	= db2html(rsget("reqname"))
				FItemList(i).Freqhp    		= db2html(rsget("reqhp"))
				FItemList(i).Freqphone    	= db2html(rsget("reqphone"))
				FItemList(i).Freqzipaddr   	= db2html(rsget("reqzipaddr"))
				FItemList(i).Freqaddress   	= db2html(rsget("reqaddress"))

				FItemList(i).Fitemname    	= db2html(rsget("itemname"))
				tmpStr = Split(FItemList(i).Fitemname, "^^")
				if (UBound(tmpStr) >= 1) then
					FItemList(i).Fitemname = tmpStr(0) & " " & tmpStr(1)
				end if

				FItemList(i).FItemNo    	= rsget("itemCnt")
				if (FItemList(i).FItemNo > 1) then
					FItemList(i).Fitemname = FItemList(i).Fitemname & " 외 " & (FItemList(i).FItemNo-1) & "종"
				end if

				FItemList(i).FpojangName   	= rsget("pojangName")
				FItemList(i).FboxNo    		= rsget("boxNo")
				FItemList(i).FpickupReqDate = rsget("pickupReqDate")
				FItemList(i).Fcomment    	= db2html(rsget("comment"))
				FItemList(i).Fsongjangprintdate = rsget("songjangprintdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetOldMisendListALL
		dim sqlStr,i
		sqlStr = " select top " + CStr(FPageSize) + " m.orderserial,d.makerid,d.itemid,d.itemname,"
		sqlStr = sqlStr + " d.itemoptionname,d.isupchebeasong,d.currstate,"
		sqlStr = sqlStr + " m.buyname,m.ipkumdate,m.regdate, m.reqname, m.deliverno, m.sitename, m.userid, "
		sqlStr = sqlStr + " m.subtotalprice, m.ipkumdiv, l.code, l.state,l.ipgodate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list l"
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

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")
				FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FCode           = rsget("code")
				FItemList(i).FState          = rsget("state")
				FItemList(i).FIpgoDate       = rsget("ipgodate")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FReqName		 = rsget("reqname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FDeliveryNo	 = rsget("deliverno")
				FItemList(i).FSiteName	 = rsget("sitename")
				FItemList(i).FUserId	 = rsget("userid")
				FItemList(i).FSubTotalPrice = rsget("subtotalprice")
				FItemList(i).Fipkumdiv = rsget("ipkumdiv")



				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetOldMisendListSearch
		dim sqlStr,i
		sqlStr = " select top " + CStr(FPageSize) + " d.orderserial,d.makerid,d.itemid,d.itemname,"
		sqlStr = sqlStr + " d.itemoptionname,d.isupchebeasong,d.currstate,"
		sqlStr = sqlStr + " m.buyname,m.ipkumdate,"
		sqlStr = sqlStr + " l.code, l.state,l.ipgodate "
		sqlStr = sqlStr + " from "
		''sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list l"
		sqlStr = sqlStr + " on d.idx=l.detailidx"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.idx>350000"
		sqlStr = sqlStr + " and datediff(m,m.ipkumdate,getdate())<2"
		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>" + CStr(FRectDelayDate)
		''sqlStr = sqlStr + " and m.sitename<>'tingmart'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		''sqlStr = sqlStr + " and d.itemid=i.itemid"
		''sqlStr = sqlStr + " and i.itemdiv<50"

		if FRectNotInCludeUpcheCheck="on" then
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.currstate is NULL))"
		else
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.beasongdate is NULL))"
		end if

		sqlStr = sqlStr + "         or (d.isupchebeasong<>'Y' and m.ipkumdiv<6))"
		sqlStr = sqlStr + " order by d.idx "

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")
				FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FCode           = rsget("code")
				FItemList(i).FState          = rsget("state")
				FItemList(i).FIpgoDate       = rsget("ipgodate")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetMiSendOrderByitemid()
		dim sqlStr,i
		sqlStr = " select top 500 m.idx, m.orderserial, m.buyname, m.reqname, m.ipkumdate, m.baljudate, d.itemno,"
		sqlStr = sqlStr + " m.regdate, m.buyphone, m.buyhp, m.deliverno, m.sitename, m.userid,"
		sqlStr = sqlStr + " m.subtotalprice, m.ipkumdiv, "
		sqlStr = sqlStr + " d.currstate, d.makerid, d.itemid, d.isupchebeasong, l.itemlackno, l.code, l.state, l.reqstr, l.finishstr"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		sqlStr = sqlStr + " [db_temp].[dbo].tbl_mibeasong_list l"
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

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FItemList(i) = new COldMiSendItem
			FItemList(i).FOrderserial = rsget("orderserial")
			FItemList(i).FMakerId     = rsget("makerid")
			FItemList(i).FItemId         = rsget("itemid")
			FItemList(i).FItemNo = rsget("itemno")

			FItemList(i).Fbuyname   = db2html(rsget("buyname"))
			FItemList(i).Freqname 	= db2html(rsget("reqname"))
			FItemList(i).Fipkumdate = rsget("ipkumdate")
			FItemList(i).Fbaljudate = rsget("baljudate")
			FItemList(i).FRegDate        = rsget("regdate")

			FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
			FItemList(i).FCurrState      = rsget("currstate")
			FItemList(i).Fitemlackno	 = rsget("itemlackno")

			FItemList(i).FCode           = rsget("code")
			FItemList(i).FState          = rsget("state")

			FItemList(i).FBuyPhone      = rsget("buyphone")
			FItemList(i).FBuyHP         = rsget("buyhp")

			FItemList(i).FDeliveryNo    = rsget("deliverno")
			FItemList(i).FSiteName      = rsget("sitename")
			FItemList(i).FUserId        = rsget("userid")
			FItemList(i).FSubTotalPrice = rsget("subtotalprice")
			FItemList(i).Fipkumdiv      = rsget("ipkumdiv")

			FItemList(i).FrequestString = rsget("reqstr")
			FItemList(i).FfinishString  = rsget("finishstr")


			i=i+1
			rsget.MoveNext
		loop
		rsget.close
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
