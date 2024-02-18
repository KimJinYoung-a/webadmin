<%
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
	public FSongJangDivName
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

class CJumunMasterItem
	public Fstatediv
	public Fbeasongetc
    public FOrderserial
    public FItemid
    public FItemname
    public FItemoption
    public FItemcnt
    public FBuyname
    public FReqname
    public FCancelYn
    public FRegdate
    public FIpkumdate
    public FBaljudate
    public FCurrstate
    public Fidx
    public FListimage
	public Fipkumdiv

	public FReqZipCode
	public FReqZipAddr
	public FReqAddress

	public Fdetailidx
	public Fitemno
	public Fitemcost
	public Fitemoptionname
	public Fcanceldate
	public Frefundstate
	public Fbeasongdate
	public Frequiredetail
	public Frequiremakeday
	public Fipgodate
	public Fupcheconfirmdate
	public Fcode
	public Fstate
	public FMCancelYn

	public Fid
	public Fdivcd
	public FdivcdName
	public Fcustomername
	public Fuserid
	public Fwriteuser
	public Ffinishuser
	public Ftitle
	public Fcurrstatename
	public FcurrstateColor
	public Ffinishdate
	public Fgubun01
	public Fgubun02
	public Fgubun01Name
	public Fgubun02Name
	public Fdeleteyn
	public Frefundrequire
	public Frefundresult
	public Fsongjangdiv
	public Fsongjangno
	public Frequireupche
	public Fmakerid
	public FExtsitename
	public Fauthcode

	Public function IpkumDivName()
		If FCancelYn="Y" Then
			IpkumDivName="주문취소"
		Else
			if Fipkumdiv="0" then
				IpkumDivName="주문대기"
			elseif Fipkumdiv="1" then
				IpkumDivName="주문실패"
			elseif Fipkumdiv="2" then
				IpkumDivName="주문접수"
			elseif Fipkumdiv="3" then
				IpkumDivName="주문접수"
			elseif Fipkumdiv="4" then
				IpkumDivName="확인대기"
			elseif Fipkumdiv="5" then
				IpkumDivName="배송통보"
			elseif Fipkumdiv="6" then
				IpkumDivName="배송준비"
			elseif Fipkumdiv="7" then
				IpkumDivName="일부출고"
			elseif Fipkumdiv="8" then
				IpkumDivName="출고완료"
			end If
		end if
	end function

	Public Function CsStateName()
		If Fcurrstate>="B006" Then
			CsStateName="처리완료"
		Else
			CsStateName="미처리"
		End If
	End Function

	Public Function StateClassName()
		If Fstatediv="결제완료" Then
			StateClassName="payFin"
		ElseIf Fstatediv="주문접수" Then
			StateClassName="odrReceive"
		ElseIf Fstatediv="주문취소" Then
			StateClassName="odrCancel"
		ElseIf Fstatediv="출고완료" Then
			StateClassName="releaseFin"
		ElseIf Fstatediv="일부출고" Then
			StateClassName="releaseIng"
		ElseIf Fstatediv="미출고 처리" Then
			StateClassName="undeliver"
		ElseIf Fstatediv="문의글 등록" Then
			StateClassName="qnaRegist"
		ElseIf Fstatediv="문의글 답변" Then
			StateClassName="qnaReply"
		ElseIf Fstatediv="판매완료" Then
			StateClassName="saleFin"
		ElseIf Fstatediv="반품/환불" Then
			StateClassName="refund"
		End If
	End Function

	Public Function GetSongJangDivName()
		Dim sqlStr
		sqlStr = " select top 1 divname from [db_order].[dbo].tbl_songjang_div"
		sqlStr = sqlStr + " where divcd='" + FSongjangdiv + "'"
		rsget.Open sqlStr,dbget,1
		If Not rsget.EOF  Then
			GetSongJangDivName = replace(db2html(rsget("divname")),"'","")
		Else
			GetSongJangDivName=""
		End If
		rsget.close
	End Function

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

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CJumunMaster
	public FMasterItemList()
	public FItemList()
	public FOneItem
	public maxt
	public maxc
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FRectRegStart
	public FRectRegEnd
	public Fipgodate
	public FCurrPage
    public FRectDesignerID
    public FItemCount
	public FItemID
	public FItemName
	public FItemimgsmall
	public FTotalFavoriteCount
	public FSubtotal
	public FItemoption
	public FItemcnt
	public FRegdate
	public FIpkumdate
	public FBaljudate
	public Fupcheconfirmdate
	public FCurrstate
	public FOrderserial
	public FCancelyn
    public Fipkumdiv
    public FItemoptioncode
	public FRectorderlistcount
	public FRectOrderSerial
	public FRectItemid
	public FRectItemoptionno
    public FRectIsAll
	public Fuserid
	public Fbeasongmemo
	public Fcardribbon
	public Fmessage
	public FBuyName
	public FBuyPhone
	public FBuyHp
	public FBuyEmail
	public FReqName
	public FReqPhone
	public FReqHp
	public FReqZipCode
	public FReqZipAddr
	public FReqAddress
	public FComment
	public Fmakerid
	public FItemNo
	public FItemoptionName
	public Fitemcost
	public Fidx
	public Fsearchstate
	public Fbeasongdate
	public FSongjangdiv
	public FSongjangno
	public Fsongjangcnt
	public FSongJangDivName
	public FDetailCancelyn
	public FMisendReason
    public FMisendState
    public FMisendipgodate
    public FisSendSMS
    public FisSendEmail
    public FisSendCall
    public FRectMakerid
    public FRectOnlyJupsu
    public FRectCurrstate
    public FRectUserName
    public FRectUserID
    public FRectDivcd
    public FRectWriteUser
    public FRectDeleteYN
    public Fsmallimage
	public FCancelDate
	public FBeadalDate
	public Ffinishdate
    public FRectSearchType
    public FRectSearchValue
    public FRectMisendReason
    public FRectDetailIDx
	public FRectMasteridx
	public FRectStateDIV
	public FRectSortDIV
	public FRectSortUpDown
	public FRectStartDate
	public FRectEndDate
	public FRectSearchTxt
	public FRectOrderDiv
	public FRectSearchDIV
	public Fid
	public Fdivcd
	public FdivcdName
	public Fcustomername
	public Fwriteuser
	public Ffinishuser
	public Ftitle
	public Fcurrstatename
	public FcurrstateColor
	public Fgubun01
	public Fgubun02
	public Fgubun01Name
	public Fgubun02Name
	public Fdeleteyn
	public Frefundrequire
	public Frefundresult
	public Frequireupche
	public FExtsitename
	public Fauthcode
	public Fcontents_jupsu

    public function isMisendAlreadyInputed()
        isMisendAlreadyInputed = Not (IsNULL(FMisendReason) or (FMisendReason="00") or (FMisendReason=""))
    end function

    public function getMisendText()
        select Case FMisendReason
            CASE "00" : getMisendText = "입력대기"
            CASE "01" : getMisendText = "재고부족"
            CASE "04" : getMisendText = "예약상품"

            CASE "02" : getMisendText = "주문제작"
            CASE "52" : getMisendText = "주문제작"
            CASE "03" : getMisendText = "출고지연"
            CASE "53" : getMisendText = "출고지연"
            CASE "05" : getMisendText = "품절출고불가"
            CASE "55" : getMisendText = "품절출고불가"
            CASE ELSE : getMisendText = FMisendReason
        end Select
    end function

	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		redim FMasterItemList(0)
		redim FItemList(0)
		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function IsAvailAndIpkumOK()
		IsAvailAndIpkumOK = (CInt(Fipkumdiv)>3) and IsAvailJumun
	end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

    public function TimeLineOrderMaster()
        dim sqlStr
        sqlStr = "select top 1 * from [db_academy].[dbo].tbl_academy_order_master where idx='" + CStr(FRectMasteridx) + "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		if not rsACADEMYget.EOF then
            set FOneItem = new CJumunMaster
			FOneItem.FOrderserial = rsACADEMYget("orderserial")
			FOneItem.FCancelYn = rsACADEMYget("cancelyn")
			FOneItem.FCancelDate = rsACADEMYget("canceldate")
			FOneItem.FIpkumdate = rsACADEMYget("ipkumdate")
			FOneItem.FBaljudate = rsACADEMYget("baljudate")
			FOneItem.FBeadalDate = rsACADEMYget("beadaldate")
        end if
        rsACADEMYget.Close
    end function

	public Sub TimeLineOrderDetail()
		dim sqlStr
		dim i
		sqlStr = "select top 100 d.currstate, d.songjangno, d.upcheconfirmdate, d.beasongdate"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d"
	    sqlStr = sqlStr + " where d.masteridx='" + FRectMasteridx + "'"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " order by d.detailidx asc"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
        if (FResultCount<1) then FResultCount=0
		redim preserve FMasterItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			Do Until rsACADEMYget.EOF
				Set FMasterItemList(i) = New CJumunMasterItem
				FMasterItemList(i).Fcurrstate = rsACADEMYget("currstate")
				FMasterItemList(i).Fsongjangno = rsACADEMYget("songjangno")
				FMasterItemList(i).Fupcheconfirmdate = rsACADEMYget("upcheconfirmdate")
				FMasterItemList(i).Fbeasongdate = rsACADEMYget("beasongdate")
				rsACADEMYget.movenext
                i=i+1
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	Public Sub TimeLineOrderDetailPartBeaSong()
		dim sqlStr
		dim i
		sqlStr = "select top 1 d.songjangno,d.songjangdiv,d.beasongdate, count(d.songjangno) as cnt"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d"
	    sqlStr = sqlStr + " where d.masteridx='" + FRectMasteridx + "'"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and isnull(d.songjangno,'')<>''"
		sqlStr = sqlStr + " group by d.songjangno, d.songjangdiv,d.beasongdate"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		set FOneItem = new CJumunMaster
		if not rsACADEMYget.EOF Then
			FOneItem.Fsongjangno = rsACADEMYget("songjangno")
			FOneItem.Fsongjangdiv = rsACADEMYget("songjangdiv")
			FOneItem.Fbeasongdate = rsACADEMYget("beasongdate")
			FOneItem.Fsongjangcnt = rsACADEMYget("cnt")
		Else
			FOneItem.Fsongjangno = ""
			FOneItem.Fsongjangdiv = ""
			FOneItem.Fbeasongdate = ""
			FOneItem.Fsongjangcnt = 0
		End If
		rsACADEMYget.Close
	End Sub

	Public Sub TimeLineOrderMiBeaSong()
		dim sqlStr
		dim i
		sqlStr = "select top 1 regdate, ipgodate"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_mibeasong_list"
	    sqlStr = sqlStr + " where orderserial='" + FRectOrderSerial + "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		set FOneItem = new CJumunMaster
		if not rsACADEMYget.EOF Then
			FOneItem.Fregdate = rsACADEMYget("regdate")
			FOneItem.Fipgodate = rsACADEMYget("ipgodate")
		Else
			FOneItem.Fregdate = ""
			FOneItem.Fipgodate = ""
		End If
		rsACADEMYget.Close
	End Sub

    public function TimeLineOrderMasterComplite()
        dim sqlStr
        sqlStr = "select top 1 dateadd(DD,7,beadaldate) as beadaldate from [db_academy].[dbo].tbl_academy_order_master where idx='" + CStr(FRectMasteridx) + "' and ipkumdiv=8 and datediff(DD,beadaldate,getdate())>=7"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		set FOneItem = new CJumunMaster
		if not rsACADEMYget.EOF then
			FOneItem.FBeadalDate = rsACADEMYget("beadaldate")
        end if
        rsACADEMYget.Close
    end function

    public function TimeLineOrderCSRefund()
        dim sqlStr
        sqlStr = "select top 1 finishdate from [db_academy].[dbo].tbl_academy_as_list where orderserial='" + CStr(FRectOrderSerial) + "' and currstate='B007'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		set FOneItem = new CJumunMaster
		if not rsACADEMYget.EOF then
			FOneItem.Ffinishdate = rsACADEMYget("finishdate")
        end if
        rsACADEMYget.Close
    end function

	public Sub DesignerDateBaljuList()
		dim sqlStr, wheredetail
		dim i
        dim IsFlowerUpche

		wheredetail = ""

		If FRectOrderDiv="S" Then
			If FRectStateDIV="1" Then'확인대기
				wheredetail = wheredetail + " and m.cancelyn<>'Y' and m.ipkumdiv='4'" + vbcrlf
			ElseIf FRectStateDIV="2" Then'주문취소
				wheredetail = wheredetail + " and m.cancelyn='Y' and m.ipkumdiv='4'" + vbcrlf
				wheredetail = wheredetail + " and (isnull(m.canceldate,'')='' or dateadd(dd,3,m.canceldate)>=getdate())" + vbcrlf
			Else'전체
				wheredetail = wheredetail + " and (isnull(m.canceldate,'')='' or dateadd(dd,3,m.canceldate)>=getdate()) and m.ipkumdiv='4'" + vbcrlf
			End If
		Else
			If FRectStateDIV="1" Then'배송대기
				wheredetail = wheredetail + " and m.cancelyn<>'Y' and m.ipkumdiv=6 and d.currstate>=3 and DateDiff(d,DateAdd(d,c.requiremakeday+2,d.upcheconfirmdate),getdate())<=0" + vbcrlf
			ElseIf FRectStateDIV="2" Then'미출고
				wheredetail = wheredetail + " and m.cancelyn<>'Y' and m.ipkumdiv>4 and m.ipkumdiv<8 and DateDiff(d,DateAdd(d,c.requiremakeday+2,d.upcheconfirmdate),getdate())>0" + vbcrlf
			ElseIf FRectStateDIV="3" Then'일부출고
				wheredetail = wheredetail + " and m.cancelyn<>'Y' and m.ipkumdiv=7 and m.jumundiv<>9 and d.currstate=7" + vbcrlf
			ElseIf FRectStateDIV="4" Then'주문취소
				wheredetail = wheredetail + " and m.cancelyn='Y' and d.cancelyn='Y' and (isnull(m.canceldate,'')='' or dateadd(dd,3,m.canceldate)<getdate())" + vbcrlf
			ElseIf FRectStateDIV="5" Then'출고완료
				wheredetail = wheredetail + " and m.cancelyn<>'Y' and m.ipkumdiv=8 and m.jumundiv<>9 and d.currstate=7" + vbcrlf
			Else'전체
				wheredetail = wheredetail + " and m.ipkumdiv>4" + vbcrlf
			End If
		End If

		If FRectStartDate ="" Then
			wheredetail = wheredetail + " and m.regdate>=dateadd(d,-90,getdate())" + vbcrlf
		Else
			wheredetail = wheredetail + " and m.regdate>=@StartDate and m.regdate<=@EndDate" + vbcrlf
		End If

		If FRectSearchDIV<>"" Then
			If FRectSearchDIV=1 Then
				wheredetail = wheredetail + " and m.orderserial like '%" + Cstr(FRectSearchTXT) + "%'" + vbcrlf
			ElseIf FRectSearchDIV=2 Then
				wheredetail = wheredetail + " and d.itemid like '%" + CStr(FRectSearchTXT) + "%'" + vbcrlf
			ElseIf FRectSearchDIV=3 Then
				wheredetail = wheredetail + " and m.buyname like '%" + CStr(FRectSearchTXT) + "%'" + vbcrlf
			ElseIf FRectSearchDIV=4 Then
				wheredetail = wheredetail + " and d.itemname like '%" + CStr(FRectSearchTXT) + "%'" + vbcrlf
			End If
		End If

		''###########################################################################
		''출고요청 리스트 / 업체 미확인건 / 플라워 주문 체크(state NULL 도 보여줌)
		''###########################################################################
		sqlStr = "select m.idx from [db_academy].[dbo].tbl_academy_order_master m" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_order_detail d on m.idx=d.masteridx" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_diy_item_Contents c on d.itemid=c.itemid" + vbcrlf
		sqlStr = sqlStr + " where d.itemid<>0 and d.itemid<>100 and d.isupchebeasong='Y' and d.makerid='" + CStr(FRectDesignerID) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.jumundiv <> '9' and m.sitename='diyitem'" + vbcrlf
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by m.idx" + vbcrlf
		'Response.write sqlStr
		'Response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		rsACADEMYget.Close

		sqlStr = "exec  [db_academy].[dbo].[sp_academy_apps_UpcheOrder_List_New] " + CStr(FPageSize) + "," + CStr(FCurrPage) + ",'" + CStr(FRectDesignerID) + "','" + FRectStartDate + "','" + FRectEndDate + "'," + CStr(FRectSearchDIV) + ",'" + FRectSearchTxt + "'," + Cstr(FRectStateDIV) + ",'" + FRectSortUpDown + "','" + FRectOrderDiv + "'"
		'Response.write sqlStr
		'Response.end
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsACADEMYget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FResultCount<1) then FResultCount=0
		redim preserve FMasterItemList(FResultCount)

		if not rsACADEMYget.EOF then
			Do Until (i >= FResultCount)
				Set FMasterItemList(i) = New CJumunMasterItem
    			FMasterItemList(i).FOrderserial = rsACADEMYget("orderserial")
    			FMasterItemList(i).FItemid 	 = rsACADEMYget("itemid")
    			FMasterItemList(i).FItemname    = db2html(rsACADEMYget("itemname"))
				FMasterItemList(i).FMCancelYn	 = rsACADEMYget("mcancelyn")
    			FMasterItemList(i).FCancelYn	 = rsACADEMYget("cancelyn")
    			FMasterItemList(i).FRegdate  = rsACADEMYget("regdate")
    			FMasterItemList(i).FIpkumdate  = rsACADEMYget("ipkumdate")
    			FMasterItemList(i).FBaljudate  = rsACADEMYget("baljudate")
    			FMasterItemList(i).FCurrstate  = rsACADEMYget("currstate")
				FMasterItemList(i).Fsongjangno  = rsACADEMYget("songjangno")
				FMasterItemList(i).Frequiremakeday  = rsACADEMYget("requiremakeday")
				FMasterItemList(i).Fupcheconfirmdate = rsACADEMYget("upcheconfirmdate")
				FMasterItemList(i).Fcode  = rsACADEMYget("code")
    			FMasterItemList(i).Fidx  = rsACADEMYget("idx")
				FMasterItemList(i).Fipkumdiv = rsACADEMYget("ipkumdiv")
				FMasterItemList(i).FListimage      = rsACADEMYget("listimage")
				if ((Not IsNULL(FMasterItemList(i).FListimage)) and (FMasterItemList(i).FListimage<>"")) then FMasterItemList(i).FListimage = imgFingers & "/diyItem/webimage/list/" + GetImageSubFolderByItemid(FMasterItemList(i).FItemID) + "/"  + FMasterItemList(i).FListimage
				rsACADEMYget.movenext
				i=i+1
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	public Sub GetCSASDetailInfo()
		dim sqlStr,i
		''###########################################################################
		''주문 상품 정보
		''###########################################################################
		sqlStr = "select d.detailidx, d.itemid, c.regitemno, d.itemcost, d.itemname, d.itemoptionname, d.cancelyn, d.canceldate, d.upcheconfirmdate," + vbcrlf
		sqlStr = sqlStr + " d.currstate, d.refundstate, d.songjangdiv, d.songjangno, d.beasongdate, d.requiredetail, i.listimage, l.ipgodate" + vbcrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_detail c" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_order_detail d on c.orderdetailidx=d.detailidx" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_diy_item i on d.itemid=i.itemid" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_mibeasong_list l on d.detailidx=l.detailidx" + vbcrlf
		sqlStr = sqlStr + " where d.itemid<>0 and d.itemid<>100 and d.isupchebeasong='Y' and c.makerid='" + CStr(FRectDesignerID) + "'" + vbcrlf
		sqlStr = sqlStr + " and c.masterid='" + CStr(FRectMasterIDX) + "'"
'Response.write sqlStr
'Response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FMasterItemList(FResultCount)
		if not rsACADEMYget.EOF then
			Do Until rsACADEMYget.eof
				Set FMasterItemList(i) = New CJumunMasterItem
    			FMasterItemList(i).Fdetailidx = rsACADEMYget("detailidx")
    			FMasterItemList(i).FItemid = rsACADEMYget("itemid")
				FMasterItemList(i).Fitemno = rsACADEMYget("regitemno")
				FMasterItemList(i).Fitemcost = rsACADEMYget("itemcost")
    			FMasterItemList(i).FItemname = db2html(rsACADEMYget("itemname"))
				FMasterItemList(i).Fitemoptionname = db2html(rsACADEMYget("itemoptionname"))
    			FMasterItemList(i).FCancelYn = rsACADEMYget("cancelyn")
    			FMasterItemList(i).Fcanceldate = rsACADEMYget("canceldate")
    			FMasterItemList(i).Frefundstate = rsACADEMYget("refundstate")
    			FMasterItemList(i).FCurrstate = rsACADEMYget("currstate")
				FMasterItemList(i).Fsongjangdiv = rsACADEMYget("songjangdiv")
				FMasterItemList(i).Fsongjangno = rsACADEMYget("songjangno")
    			FMasterItemList(i).Fbeasongdate = rsACADEMYget("beasongdate")
				FMasterItemList(i).Fupcheconfirmdate = rsACADEMYget("upcheconfirmdate")
				FMasterItemList(i).Frequiredetail = db2html(rsACADEMYget("requiredetail"))
				FMasterItemList(i).Fipgodate = rsACADEMYget("ipgodate")
				FMasterItemList(i).FListimage = rsACADEMYget("listimage")
				if ((Not IsNULL(FMasterItemList(i).FListimage)) and (FMasterItemList(i).FListimage<>"")) then FMasterItemList(i).FListimage = imgFingers & "/diyItem/webimage/list/" + GetImageSubFolderByItemid(FMasterItemList(i).FItemID) + "/"  + FMasterItemList(i).FListimage
				rsACADEMYget.movenext
				i=i+1
			Loop
		End If
		rsACADEMYget.Close
	End Sub

    public Sub GetCSASMasterList()
        dim i,sqlStr, AddSQL
        AddSQL = ""

        sqlStr = " select count(A.id) as cnt "
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_list A"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_as_refund_info r"
        sqlStr = sqlStr + " on A.id=r.asid"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_order_master m"
        sqlStr = sqlStr + " on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " where 1 = 1"
        sqlStr = sqlStr + " and m.sitename <> 'academy' "

		if (FRectSearchType="") then
		    if (FRectOrderSerial<>"") then
		        AddSQL = AddSQL + " and A.orderserial like '%" + FRectOrderSerial + "%'"
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

            if (FRectStateDIV = "1") then
	                AddSQL = AddSQL + " and A.currstate < 'B006' "
	        elseif (FRectStateDIV = "2") then
	                AddSQL = AddSQL + " and A.currstate >='B006'"
	        end if

            if (FRectUserName <> "") then
	                AddSQL = AddSQL + " and A.customername like '%" + CStr(FRectUserName) + "%'"
	        end if

	        if (FRectOrderSerial <> "") then
	                AddSQL = AddSQL + " and A.orderserial like '%" + FRectOrderSerial + "%'"
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
	                AddSQL = AddSQL + " and A.orderserial like '%" + FRectOrderSerial + "%'"
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
		'Response.write sqlStr
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1
        if  not rsACADEMYget.EOF  then
            FTotalCount = rsACADEMYget("cnt")
        else
            FTotalCount = 0
        end if
        rsACADEMYget.close

        sqlStr = " select Top " + CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr + " A.id, A.divcd, A.gubun01, A.gubun02, A.orderserial, A.customername, A.userid, A.finishuser, A.writeuser, A.title, A.currstate"
        sqlStr = sqlStr + " ,A.regdate, A.finishdate,A.deleteyn "
        sqlStr = sqlStr + " , A.requireupche, A.makerid, A.songjangdiv ,A.songjangno"
        sqlStr = sqlStr + " ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult"
        sqlStr = sqlStr + " ,m.sitename, m.authcode"
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename, C4.comm_color as currstatecolor"
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_list A"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_order_master m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
        sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and m.sitename <> 'academy'"

        sqlStr = sqlStr + AddSQL
		If FRectSortUpDown="u" Then
        sqlStr = sqlStr + " order by id desc"
		Else
		sqlStr = sqlStr + " order by id asc"
		End If
        rsACADEMYget.pagesize = FPageSize
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1

        FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

        redim preserve FItemList(FResultCount)
        if  not rsACADEMYget.EOF  then
            i = 0
			rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.eof
                set FItemList(i) = new CJumunMasterItem

                FItemList(i).Fid                = rsACADEMYget("id")
                FItemList(i).Fdivcd             = rsACADEMYget("divcd")
                FItemList(i).FdivcdName         = db2html(rsACADEMYget("divcdname"))

                FItemList(i).Forderserial       = rsACADEMYget("orderserial")
                FItemList(i).Fcustomername      = db2html(rsACADEMYget("customername"))
                FItemList(i).Fuserid            = rsACADEMYget("userid")
                FItemList(i).Fwriteuser         = rsACADEMYget("writeuser")
                FItemList(i).Ffinishuser        = rsACADEMYget("finishuser")
                FItemList(i).Ftitle             = db2html(rsACADEMYget("title"))
                FItemList(i).Fcurrstate         = rsACADEMYget("currstate")
                FItemList(i).Fcurrstatename     = rsACADEMYget("currstatename")
                FItemList(i).FcurrstateColor    = rsACADEMYget("currstatecolor")

                FItemList(i).Fregdate           = rsACADEMYget("regdate")
                FItemList(i).Ffinishdate        = rsACADEMYget("finishdate")

                FItemList(i).Fgubun01           = rsACADEMYget("gubun01")
                FItemList(i).Fgubun02           = rsACADEMYget("gubun02")

                FItemList(i).Fgubun01Name       = db2html(rsACADEMYget("gubun01name"))
                FItemList(i).Fgubun02Name       = db2html(rsACADEMYget("gubun02name"))

                FItemList(i).Fdeleteyn          = rsACADEMYget("deleteyn")

                FItemList(i).Frefundrequire     = rsACADEMYget("refundrequire")
                FItemList(i).Frefundresult      = rsACADEMYget("refundresult")

                FItemList(i).Fsongjangdiv       = rsACADEMYget("songjangdiv")
                FItemList(i).Fsongjangno        = rsACADEMYget("songjangno")

                FItemList(i).Frequireupche      = rsACADEMYget("requireupche")
                FItemList(i).Fmakerid           = rsACADEMYget("makerid")

                FItemList(i).FExtsitename          = rsACADEMYget("sitename")
                FItemList(i).Fauthcode          = rsACADEMYget("authcode")

                rsACADEMYget.MoveNext
                i = i + 1
            loop
        end if
        rsACADEMYget.close
    end sub

	Public Function OneOrderMasterInfo()
		dim sqlStr
		sqlStr = "select top 1 idx,userid,ipkumdiv,ipkumdate,regdate,beadaldate,canceldate,cancelyn,buyname,buyphone," + vbcrlf
		sqlStr = sqlStr + "buyhp,reqname,reqzipcode,reqzipaddr,reqaddress,reqphone,reqhp,beasongmemo,cardribbon,message" + vbcrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master where orderserial='" + CStr(FRectOrderSerial) + "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		if not rsACADEMYget.EOF then
			set FOneItem = new CJumunMaster
			FOneItem.Fidx = rsACADEMYget("idx")
			FOneItem.Fuserid = rsACADEMYget("userid")
			FOneItem.Fipkumdiv = rsACADEMYget("ipkumdiv")
			FOneItem.Fipkumdate = rsACADEMYget("ipkumdate")
			FOneItem.Fregdate = rsACADEMYget("regdate")
			FOneItem.Fbeadaldate = rsACADEMYget("beadaldate")
			FOneItem.Fcanceldate = rsACADEMYget("canceldate")
			FOneItem.Fcancelyn = rsACADEMYget("cancelyn")
			FOneItem.Fbuyname = rsACADEMYget("buyname")
			FOneItem.Fbuyphone = rsACADEMYget("buyphone")
			FOneItem.Fbuyhp = rsACADEMYget("buyhp")
			FOneItem.Freqname = rsACADEMYget("reqname")
			FOneItem.Freqzipcode = rsACADEMYget("reqzipcode")
			FOneItem.Freqzipaddr = rsACADEMYget("reqzipaddr")
			FOneItem.Freqaddress = rsACADEMYget("reqaddress")
			FOneItem.Freqphone = rsACADEMYget("reqphone")
			FOneItem.Freqhp = rsACADEMYget("reqhp")
			FOneItem.Fbeasongmemo = rsACADEMYget("beasongmemo")
			FOneItem.Fcardribbon = rsACADEMYget("cardribbon")
			FOneItem.Fmessage = rsACADEMYget("message")
		end if
		rsACADEMYget.Close
	end Function
	
	public Sub OrderDetailInfo()
		dim sqlStr,i
		''###########################################################################
		''주문 상품 정보
		''###########################################################################
		sqlStr = "select d.detailidx, d.itemid, d.itemno, d.itemcost, d.itemname, d.itemoptionname, d.cancelyn, d.canceldate, d.upcheconfirmdate," + vbcrlf
		sqlStr = sqlStr + " d.currstate, d.refundstate, d.songjangdiv, d.songjangno, d.beasongdate, d.requiredetail, i.listimage, c.requiremakeday, l.code, l.state, l.ipgodate" + vbcrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_diy_item i on d.itemid=i.itemid" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_diy_item_Contents c on d.itemid=c.itemid" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_mibeasong_list l on d.detailidx=l.detailidx" + vbcrlf
		sqlStr = sqlStr + " where d.itemid<>0 and d.itemid<>100 and d.isupchebeasong='Y' and d.makerid='" + CStr(FRectDesignerID) + "'" + vbcrlf
		sqlStr = sqlStr + " and d.orderserial='" + CStr(FRectOrderSerial) + "'"
'Response.write sqlStr
'Response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FMasterItemList(FResultCount)
		if not rsACADEMYget.EOF then
			Do Until rsACADEMYget.eof
				Set FMasterItemList(i) = New CJumunMasterItem
    			FMasterItemList(i).Fdetailidx = rsACADEMYget("detailidx")
    			FMasterItemList(i).FItemid = rsACADEMYget("itemid")
				FMasterItemList(i).Fitemno = rsACADEMYget("itemno")
				FMasterItemList(i).Fitemcost = rsACADEMYget("itemcost")
    			FMasterItemList(i).FItemname = db2html(rsACADEMYget("itemname"))
				FMasterItemList(i).Fitemoptionname = db2html(rsACADEMYget("itemoptionname"))
    			FMasterItemList(i).FCancelYn = rsACADEMYget("cancelyn")
    			FMasterItemList(i).Fcanceldate = rsACADEMYget("canceldate")
    			FMasterItemList(i).Frefundstate = rsACADEMYget("refundstate")
    			FMasterItemList(i).FCurrstate = rsACADEMYget("currstate")
				FMasterItemList(i).Fsongjangdiv = rsACADEMYget("songjangdiv")
				FMasterItemList(i).Fsongjangno = rsACADEMYget("songjangno")
    			FMasterItemList(i).Fbeasongdate = rsACADEMYget("beasongdate")
				FMasterItemList(i).Fupcheconfirmdate = rsACADEMYget("upcheconfirmdate")
				FMasterItemList(i).Frequiremakeday = rsACADEMYget("requiremakeday")
				FMasterItemList(i).Frequiredetail = db2html(rsACADEMYget("requiredetail"))
				FMasterItemList(i).Fcode = rsACADEMYget("code")
				FMasterItemList(i).Fstate = rsACADEMYget("state")
				FMasterItemList(i).Fipgodate = rsACADEMYget("ipgodate")
				FMasterItemList(i).FListimage = rsACADEMYget("listimage")
				if ((Not IsNULL(FMasterItemList(i).FListimage)) and (FMasterItemList(i).FListimage<>"")) then FMasterItemList(i).FListimage = imgFingers & "/diyItem/webimage/list/" + GetImageSubFolderByItemid(FMasterItemList(i).FItemID) + "/"  + FMasterItemList(i).FListimage
				rsACADEMYget.movenext
				i=i+1
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	public Sub OrderDetailInfoInidx()
		dim sqlStr,i
		''###########################################################################
		''주문 상품 정보
		''###########################################################################
		sqlStr = "select d.detailidx, d.itemid, d.itemno, d.itemcost, d.itemname, d.itemoption, d.itemoptionname, d.cancelyn, d.canceldate," + vbcrlf
		sqlStr = sqlStr + " d.currstate, d.refundstate, d.songjangdiv, d.songjangno, d.beasongdate, d.requiredetail, i.listimage, l.code, l.state, l.ipgodate" + vbcrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_diy_item i on d.itemid=i.itemid" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_mibeasong_list l on d.detailidx=l.detailidx" + vbcrlf
		sqlStr = sqlStr + " where d.itemid<>0 and d.itemid<>100 and d.isupchebeasong='Y'" + vbcrlf
		sqlStr = sqlStr + " and d.detailidx in (" + CStr(FRectDetailidx) + ")"
		sqlStr = sqlStr + " order by detailidx asc"
'Response.write sqlStr
'Response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FMasterItemList(FResultCount)
		if not rsACADEMYget.EOF then
			Do Until rsACADEMYget.eof
				Set FMasterItemList(i) = New CJumunMasterItem
    			FMasterItemList(i).Fdetailidx = rsACADEMYget("detailidx")
    			FMasterItemList(i).FItemid = rsACADEMYget("itemid")
				FMasterItemList(i).Fitemno = rsACADEMYget("itemno")
				FMasterItemList(i).Fitemcost = rsACADEMYget("itemcost")
				FMasterItemList(i).Fitemoption = rsACADEMYget("itemoption")
    			FMasterItemList(i).FItemname = db2html(rsACADEMYget("itemname"))
				FMasterItemList(i).Fitemoptionname = db2html(rsACADEMYget("itemoptionname"))
    			FMasterItemList(i).FCancelYn = rsACADEMYget("cancelyn")
    			FMasterItemList(i).Fcanceldate = rsACADEMYget("canceldate")
    			FMasterItemList(i).Frefundstate = rsACADEMYget("refundstate")
    			FMasterItemList(i).FCurrstate = rsACADEMYget("currstate")
				FMasterItemList(i).Fsongjangdiv = rsACADEMYget("songjangdiv")
				FMasterItemList(i).Fsongjangno = rsACADEMYget("songjangno")
    			FMasterItemList(i).Fbeasongdate = rsACADEMYget("beasongdate")
				FMasterItemList(i).Frequiredetail = db2html(rsACADEMYget("requiredetail"))
				FMasterItemList(i).Fcode = rsACADEMYget("code")
				FMasterItemList(i).Fstate = rsACADEMYget("state")
				FMasterItemList(i).Fipgodate = rsACADEMYget("ipgodate")
				FMasterItemList(i).FListimage = rsACADEMYget("listimage")
				if ((Not IsNULL(FMasterItemList(i).FListimage)) and (FMasterItemList(i).FListimage<>"")) then FMasterItemList(i).FListimage = imgFingers & "/diyItem/webimage/list/" + GetImageSubFolderByItemid(FMasterItemList(i).FItemID) + "/"  + FMasterItemList(i).FListimage
				rsACADEMYget.movenext
				i=i+1
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	public Sub InvoiceBatchWriteList()
		dim sqlStr, wheredetail
		dim i
        dim IsFlowerUpche

		wheredetail = ""

		If FRectSearchDIV<>"" Then
			If FRectSearchDIV=1 Then
				wheredetail = wheredetail + " and m.orderserial like '%" + Cstr(FRectSearchTXT) + "%'" + vbcrlf
			ElseIf FRectSearchDIV=2 Then
				wheredetail = wheredetail + " and d.itemid like '%" + CStr(FRectSearchTXT) + "%'" + vbcrlf
			ElseIf FRectSearchDIV=3 Then
				wheredetail = wheredetail + " and m.buyname like '%" + CStr(FRectSearchTXT) + "%'" + vbcrlf
			ElseIf FRectSearchDIV=4 Then
				wheredetail = wheredetail + " and d.itemname like '%" + CStr(FRectSearchTXT) + "%'" + vbcrlf
			End If
		End If

		''###########################################################################
		''출고요청 리스트 / 업체 미확인건 / 플라워 주문 체크(state NULL 도 보여줌)
		''###########################################################################
		sqlStr = "select m.idx from [db_academy].[dbo].tbl_academy_order_master m" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_order_detail d on m.idx=d.masteridx" + vbcrlf
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_diy_item_Contents c on d.itemid=c.itemid" + vbcrlf
		sqlStr = sqlStr + " where d.itemid<>0 and d.itemid<>100 and d.isupchebeasong='Y' and d.makerid='" + CStr(FRectDesignerID) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.jumundiv<>9 and m.ipkumdiv>5 and m.ipkumdiv<8" + vbcrlf
		sqlStr = sqlStr + " and d.currstate>2 and d.currstate<7 and d.cancelyn<>'Y' and m.cancelyn<>'Y'" + vbcrlf
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by m.idx" + vbcrlf
		'Response.write sqlStr
		'Response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		rsACADEMYget.Close

		sqlStr = "exec [db_academy].[dbo].[sp_academy_apps_InvoiceBatch_List] " + CStr(FPageSize) + "," + CStr(FCurrPage) + ",'" + CStr(FRectDesignerID) + "'," + Cstr(FRectSearchDIV) + ",'" + Cstr(FRectSearchTXT) + "'"
		'Response.write sqlStr
		'Response.end
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsACADEMYget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FResultCount<1) then FResultCount=0
		redim preserve FMasterItemList(FResultCount)

		if not rsACADEMYget.EOF then
			Do Until (i >= FResultCount)
				Set FMasterItemList(i) = New CJumunMasterItem
				FMasterItemList(i).Fdetailidx = rsACADEMYget("idx")
    			FMasterItemList(i).FOrderserial = rsACADEMYget("orderserial")
    			FMasterItemList(i).FItemid 	 = rsACADEMYget("itemid")
				FMasterItemList(i).Fbuyname = rsACADEMYget("buyname")
				FMasterItemList(i).Fipkumdate = rsACADEMYget("ipkumdate")
    			FMasterItemList(i).FItemname = db2html(rsACADEMYget("itemname"))
				FMasterItemList(i).Fitemoptionname = db2html(rsACADEMYget("itemoptionname"))
				FMasterItemList(i).Fitemno	 = rsACADEMYget("itemno")
    			FMasterItemList(i).Frequiredetail = db2html(rsACADEMYget("requiredetail"))
				FMasterItemList(i).Freqzipcode = rsACADEMYget("reqzipcode")
				FMasterItemList(i).Freqzipaddr = rsACADEMYget("reqzipaddr")
				FMasterItemList(i).Freqaddress = rsACADEMYget("reqaddress")
				FMasterItemList(i).FListimage      = rsACADEMYget("listimage")
				if ((Not IsNULL(FMasterItemList(i).FListimage)) and (FMasterItemList(i).FListimage<>"")) then FMasterItemList(i).FListimage = imgFingers & "/diyItem/webimage/list/" + GetImageSubFolderByItemid(FMasterItemList(i).FItemID) + "/"  + FMasterItemList(i).FListimage
				rsACADEMYget.movenext
				i=i+1
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	public Sub OrderTimeLineList()
		dim sqlStr
		dim i

		sqlStr = "exec [db_academy].[dbo].[sp_academy_apps_OrderTimeLine] '" + CStr(FRectOrderSerial) + "','" + CStr(FRectDesignerID) + "'"
		'Response.write sqlStr
		'Response.end
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsACADEMYget.RecordCount

		redim preserve FMasterItemList(FResultCount)
		if not rsACADEMYget.EOF then
			Do Until (i >= FResultCount)
				Set FMasterItemList(i) = New CJumunMasterItem
				FMasterItemList(i).Fregdate = rsACADEMYget("regdate")
    			FMasterItemList(i).Fstatediv = rsACADEMYget("statediv")
    			FMasterItemList(i).Fbeasongetc 	 = rsACADEMYget("beasongetc")
				rsACADEMYget.movenext
				i=i+1
			Loop
		End If
		rsACADEMYget.Close
	End Sub

End Class

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
	public FRectCSID


	Public FRectDeleteYN	' 삭제제외여부
	Public FRectWriteUser	' 접수자아이디 검색

    public Sub GetCSASMasterInfo()
        dim i,sqlStr

        sqlStr = " select top 1 A.id, A.divcd, A.gubun01, A.gubun02, A.orderserial, A.customername, A.userid, A.finishuser, A.writeuser, A.title, A.currstate"
        sqlStr = sqlStr + " ,A.regdate, A.finishdate,A.deleteyn, A.contents_jupsu, A.contents_finish"
        sqlStr = sqlStr + " , A.requireupche, A.makerid, A.songjangdiv ,A.songjangno, s.divname"
        sqlStr = sqlStr + " ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult"
        sqlStr = sqlStr + " ,m.sitename, m.authcode"
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename, C4.comm_color as currstatecolor"
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_list A"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_order_master m"
        sqlStr = sqlStr + "  on A.orderserial=m.orderserial"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
		sqlStr = sqlStr + "  left join [db_academy].[dbo].tbl_songjang_div s on A.songjangdiv=s.divcd"
        sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and m.sitename <> 'academy'"
		sqlStr = sqlStr + " and A.orderserial='" + CStr(FRectOrderSerial) + "' "
		sqlStr = sqlStr + " and A.id='" + CStr(FRectCSID) + "' "
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1
        if  not rsACADEMYget.EOF  then
			set FOneItem = new CCSASMasterItem
			FOneItem.Fid                = rsACADEMYget("id")
			FOneItem.Fdivcd             = rsACADEMYget("divcd")
			FOneItem.FdivcdName         = db2html(rsACADEMYget("divcdname"))

			FOneItem.Forderserial       = rsACADEMYget("orderserial")
			FOneItem.Fcustomername      = db2html(rsACADEMYget("customername"))
			FOneItem.Fuserid            = rsACADEMYget("userid")
			FOneItem.Fwriteuser         = rsACADEMYget("writeuser")
			FOneItem.Ffinishuser        = rsACADEMYget("finishuser")
			FOneItem.Ftitle             = db2html(rsACADEMYget("title"))
			FOneItem.Fcurrstate         = rsACADEMYget("currstate")
			FOneItem.Fcurrstatename     = rsACADEMYget("currstatename")
			FOneItem.FcurrstateColor    = rsACADEMYget("currstatecolor")

			FOneItem.Fregdate           = rsACADEMYget("regdate")
			FOneItem.Ffinishdate        = rsACADEMYget("finishdate")

			FOneItem.Fgubun01           = rsACADEMYget("gubun01")
			FOneItem.Fgubun02           = rsACADEMYget("gubun02")
			FOneItem.Fcontents_jupsu       = db2html(rsACADEMYget("contents_jupsu"))
			FOneItem.Fcontents_finish     = db2html(rsACADEMYget("contents_finish"))
			FOneItem.Fgubun01Name       = db2html(rsACADEMYget("gubun01name"))
			FOneItem.Fgubun02Name       = db2html(rsACADEMYget("gubun02name"))

			FOneItem.Fdeleteyn          = rsACADEMYget("deleteyn")

			FOneItem.Frefundrequire     = rsACADEMYget("refundrequire")
			FOneItem.Frefundresult      = rsACADEMYget("refundresult")

			FOneItem.Fsongjangdiv       = rsACADEMYget("songjangdiv")
			FOneItem.Fsongjangno        = rsACADEMYget("songjangno")
			FOneItem.Fsongjangdivname   = db2html(rsACADEMYget("divname"))

			FOneItem.Frequireupche      = rsACADEMYget("requireupche")
			FOneItem.Fmakerid           = rsACADEMYget("makerid")

			FOneItem.FExtsitename       = rsACADEMYget("sitename")
			FOneItem.Fauthcode          = rsACADEMYget("authcode")
        end if
        rsACADEMYget.close
    end sub

    public Sub GetOneCSASMaster()
        dim i,sqlStr

        sqlStr = " select top 1 A.*, IsNULL(J.add_upchejungsandeliverypay,0) as add_upchejungsandeliverypay, J.add_upchejungsancause "
        sqlStr = sqlStr + " ,IsNULL(r.refundrequire,0) as refundrequire, IsNULL(r.refundresult,0) as refundresult"
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C2.comm_name as gubun01name, C3.comm_name as gubun02name, C4.comm_name as currstatename"
        sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_list A "
        sqlStr = sqlStr + " Left join [db_academy].[dbo].tbl_academy_as_upcheAddjungsan J"
        sqlStr = sqlStr + "  on A.id=J.asid"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_as_refund_info r"
        sqlStr = sqlStr + "  on A.id=r.asid"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C2"
        sqlStr = sqlStr + "  on A.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C3"
        sqlStr = sqlStr + "  on A.gubun02=C3.comm_cd"
        sqlStr = sqlStr + " Left Join [db_academy].[dbo].tbl_academy_cs_comm_code C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"

        sqlStr = sqlStr + " where A.id= " + CStr(FRectCsAsID) + " "

        if (FRectMakerID<>"") then   ''업체 조회용.
            sqlStr = sqlStr + " and A.makerid='"&FRectMakerID&"'"
        end if
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1

        FResultCount = rsACADEMYget.RecordCount

        if  not rsACADEMYget.EOF  then
            set FOneItem = new CCSASMasterItem

            FOneItem.Fid                  = rsACADEMYget("id")
            FOneItem.Fdivcd               = rsACADEMYget("divcd")
            FOneItem.Fgubun01             = rsACADEMYget("gubun01")
            FOneItem.Fgubun02             = rsACADEMYget("gubun02")

            FOneItem.FdivcdName           = db2html(rsACADEMYget("divcdname"))
            FOneItem.Fgubun01Name         = db2html(rsACADEMYget("gubun01name"))
            FOneItem.Fgubun02Name         = db2html(rsACADEMYget("gubun02name"))

            FOneItem.Forderserial         = rsACADEMYget("orderserial")
            FOneItem.Fcustomername        = db2html(rsACADEMYget("customername"))
            FOneItem.Fuserid              = rsACADEMYget("userid")
            FOneItem.Fwriteuser           = rsACADEMYget("writeuser")
            FOneItem.Ffinishuser          = rsACADEMYget("finishuser")
            FOneItem.Ftitle               = db2html(rsACADEMYget("title"))
            FOneItem.Fcontents_jupsu      = db2html(rsACADEMYget("contents_jupsu"))
            FOneItem.Fcontents_finish     = db2html(rsACADEMYget("contents_finish"))
            FOneItem.Fcurrstate           = rsACADEMYget("currstate")
            FOneItem.FcurrstateName       = rsACADEMYget("currstatename")
            FOneItem.Fregdate             = rsACADEMYget("regdate")
            FOneItem.Ffinishdate          = rsACADEMYget("finishdate")

            FOneItem.Fdeleteyn            = rsACADEMYget("deleteyn")
            FOneItem.Fextsitename         = rsACADEMYget("extsitename")

            FOneItem.Fopentitle           = db2html(rsACADEMYget("opentitle"))
            FOneItem.Fopencontents        = db2html(rsACADEMYget("opencontents"))


            FOneItem.Fsitegubun           = rsACADEMYget("sitegubun")

            FOneItem.Fsongjangdiv         = rsACADEMYget("songjangdiv")
            FOneItem.Fsongjangno          = rsACADEMYget("songjangno")

            FOneItem.Frequireupche        = rsACADEMYget("requireupche")
            FOneItem.Fmakerid             = rsACADEMYget("makerid")

            FOneItem.Fadd_upchejungsandeliverypay = rsACADEMYget("add_upchejungsandeliverypay")
            FOneItem.Fadd_upchejungsancause       = rsACADEMYget("add_upchejungsancause")

'            FOneItem.Fbeasongdate         = rsget("beasongdate")
'            FOneItem.Frefundrequire       = rsget("refundrequire")
'            FOneItem.Frefundresult        = rsget("refundresult")

        end if
        rsACADEMYget.close
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

Function SetUpCheOrderConfirm(orderserial,makerid)
	''#################################################
	''업체 선택 주문 확인
	''#################################################
	Dim sqlStr, CheckRecord
	sqlStr = "if exists(select idx from [db_academy].[dbo].tbl_academy_order_master where orderserial='" + orderserial + "' and ipkumdiv in (4,5))" + vbCrlf
	sqlStr = sqlStr + "	begin" + vbCrlf
	sqlStr = sqlStr + "		update [db_academy].[dbo].tbl_academy_order_master" + vbCrlf
	sqlStr = sqlStr + "		 set ipkumdiv=6" + vbCrlf
	sqlStr = sqlStr + "		 where orderserial='" + orderserial + "'" + vbCrlf
	sqlStr = sqlStr + "		 and ipkumdiv in (4,5)" + vbCrlf
	sqlStr = sqlStr + "		 and cancelyn='N'" + vbCrlf
	sqlStr = sqlStr + "		update [db_academy].[dbo].tbl_academy_order_detail" + vbCrlf
	sqlStr = sqlStr + "		 set currstate = '3'" + vbCrlf
	sqlStr = sqlStr + "		 ,upcheconfirmdate=getdate()" + vbCrlf
	sqlStr = sqlStr + "		 where makerid='" + CStr(makerid) + "'" + vbCrlf
	sqlStr = sqlStr + "		 and orderserial='" + orderserial + "'" + vbCrlf
	sqlStr = sqlStr + "		 and ((currstate is NULL) or (currstate ='2'))" + vbCrlf
	sqlStr = sqlStr + "		update [db_academy].[dbo].[tbl_academy_app_iconbadge_count]" + vbCrlf
	sqlStr = sqlStr + "		 set mibaljucnt=mibaljucnt-1, ordercnt=ordercnt+1" + vbCrlf
	sqlStr = sqlStr + "		 where makerid='" + CStr(makerid) + "'" + vbCrlf
	sqlStr = sqlStr + "	end" + vbCrlf
	dbACADEMYget.Execute sqlStr, CheckRecord
	SetUpCheOrderConfirm=CheckRecord
End Function

Function GetOrderStateNum(makerid)
	Dim sqlStr, CheckRecord
	sqlStr = "select mibaljucnt, ordercnt from [db_academy].[dbo].[tbl_academy_app_iconbadge_count] where makerid='" + CStr(makerid) + "'" + vbCrlf
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	if not rsACADEMYget.EOF then
		GetOrderStateNum=rsACADEMYget("mibaljucnt")+rsACADEMYget("ordercnt")
	Else
		GetOrderStateNum=0
	End If
	rsACADEMYget.close
End Function

Sub GetCheckIconBadgeCount(ByVal MakerID, ByRef StandByConfirmCnt, ByRef MiBeasongCnt, ByRef OrderCSCnt, ByRef UpdateCheck)
	Dim sqlStr
	sqlStr = "exec [db_academy].[dbo].[sp_Academy_App_IconBadgeCountOrderCheckSet] '" + CStr(MakerID) + "'"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	if not rsACADEMYget.EOF then
		StandByConfirmCnt=rsACADEMYget("StandByConfirmCnt")
		MiBeasongCnt=rsACADEMYget("MiBeasongCnt")
		OrderCSCnt=rsACADEMYget("OrderCSCnt")
		UpdateCheck=rsACADEMYget("UpdateCheck")
	End If
	rsACADEMYget.close
End Sub

Function GetOrderCancelViewCheck(masteridx)
	Dim sqlStr
	sqlStr = "exec [db_academy].[dbo].[sp_academy_apps_OrderCancelView_Check] " + CStr(masteridx)
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	if not rsACADEMYget.EOF then
		GetOrderCancelViewCheck=rsACADEMYget("checkyn")
	End If
	rsACADEMYget.close
End Function

Sub GetBeasongState(ByVal songjangno, ByVal code, ByVal requiremakeday, ByVal upcheconfirmdate, ByVal CancelYn, ByVal vItemCount, ByVal vBeasongCnt, ByVal vMibeasongCnt, ByVal vMiChulGoCheck, ByVal vOrderCanCelCnt, ByRef BeasongCnt, ByRef MibeasongCnt, ByRef MiChulGoCheck, ByRef OrderCanCelCnt, ByRef BeasongState, ByRef BeasongStateName, ByRef BeasongStateClass)
	'배송 상태 확인
	If songjangno <> "" Then
		vBeasongCnt=vBeasongCnt+1
	End If
	If code <> "" Then
		vMibeasongCnt=vMibeasongCnt+1
	End If
	If (DateDiff("d",DateAdd("d",requiremakeday+2,upcheconfirmdate),now())>0) Then
		vMiChulGoCheck=vMiChulGoCheck+1
	End If
	If CancelYn="Y" Then
		vOrderCanCelCnt=vOrderCanCelCnt+1
	End If

	If vItemCount = vBeasongCnt Then
		BeasongState="0"
		BeasongStateName="출고완료"
		BeasongStateClass="releaseFin"
	ElseIf vItemCount > vBeasongCnt And vBeasongCnt>0 Then
		BeasongState="1"
		BeasongStateName="일부출고"
		BeasongStateClass="releaseIng"
	ElseIf vMibeasongCnt>0 Or vMiChulGoCheck>0 Then
		BeasongState="2"
		BeasongStateName="미출고"
		BeasongStateClass="undeliver"
	ElseIf vItemCount=vOrderCanCelCnt Then
		BeasongState="4"
		BeasongStateName="주문취소"
		BeasongStateClass="odrCancel"
	Else
		BeasongState="3"
		BeasongStateName="배송대기"
		BeasongStateClass="standby"
	End If
	BeasongCnt=vBeasongCnt
	MibeasongCnt=vMibeasongCnt
	MiChulGoCheck=vMiChulGoCheck
	OrderCanCelCnt=vOrderCanCelCnt
End Sub
%>