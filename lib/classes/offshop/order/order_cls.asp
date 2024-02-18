<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.22 한용민 생성
'###########################################################

Class COrderItem
	public frealsellprice
	public Fupcheconfirmdate
	public Forderno
	public Fmasteridx
	public Fjumundiv	
	public Faccountname
	public Faccountdiv
	public Faccountno
	public Ftotalvat
	public Ftotalcost
	public Ftotalmileage
	public Ftotalsum
	public Fipkumdiv
	public Fregdate
	public Fbeadaldiv
	public Fbeadaldate
	public Fcancelyn
	public Fbuyname
	public Fbuyphone
	public Fbuyhp
	public Fbuyemail
	public Freqname
	public Freqzipcode
	public Freqaddress
	public Freqphone
	public Freqhp
	public Freqemail
	public Fcomment
	public Fdeliverno		
	public Fresultmsg
	public Frduserid	
	public Fjungsanflag
	public Freqzipaddr	
	public Fsongjangdiv
	public Fbeasongmemo	
	public FInsureCd
	public Fcashreceiptreq
	public FcashreceiptTid
	public FcashreceiptIdx
	public Finireceipttid
	public Freferip	
	public Flinkorderno	
	public Fsentenceidx
	public Fbaljudate	
	public FcountryNameKr
	public FcountryNameEn
	public FemsAreaCode
    public FemsZipCode
    public FitemGubunName
    public FgoodNames
    public FitemWeigth
    public FitemUsDollar
    public FemsInsureYn
    public FemsInsurePrice
    public FemsDlvCost
    public fshopname
    public fdetailidx
    public Fmakerid
    public fitemid
    public fitemoption
    public fitemno
    public fsellprice
    public FItemName
    public Fitemoptionname
    public FCurrState
    public Fsongjangno
    public Fbeasongdate
    public Fisupchebeasong
    public Fsongjangdivname
    public Ffindurl
    public fitemgubun
	public fcurrsellcash
	public FODlvType
	public Fdivcd
	public FdivcdName
	public Fcustomername
	public Fwriteuser
	public Ffinishuser
	public Ftitle
	public Fcurrstatename
	public FcurrstateColor
	public Ffinishdate
	public Fdeleteyn
	public Frequireupche
	public Fcontents_jupsu
	public Fcontents_finish
	public Fopentitle
	public Fopencontents
	public forgmasteridx
	public Fregitemno
	public Fconfirmitemno
	public Fregdetailstate
	public Forderdetailidx
	public ForderDetailcurrstate
	public FDetailCancelYn
	public FCode
	public FState
	public Fipgodate
	public FMisendReason
	public FMisendState
	public FMisendipgodate
	public FisSendSMS
	public FisSendEmail
	public FisSendCall
	public FrequestString
	public Fitemlackno
	public FfinishString
	public Fcompany_name
	public Fcompany_tel
	public fOrdersellprice
	public Fasid
	public Freqetcaddr
	public Freqetcstr
	public Fsenddate
	public FreturnName
	public FreturnPhone
	public Freturnhp
	public FreturnZipcode
	public FreturnZipaddr
	public FreturnEtcaddr
	public fcancelorgorderno
	public fcancelorgdetailidx
	
	'//배송구분(odlvType) 텐바이텐 배송이냐 업체 배송이냐..
	public function getbeasonggubun()
		if fodlvType = "0" then
			getbeasonggubun = "매장배송"
		elseif fodlvType = "1" then
			getbeasonggubun = "물류배송"
		elseif fodlvType = "4" then
			getbeasonggubun = "텐바이텐무료배송"
		elseif fodlvType = "2" then
			getbeasonggubun = "업체배송"
		elseif fodlvType = "7" then
			getbeasonggubun = "업체착불배송"
		else
			getbeasonggubun = "설정안됨"
		end if

	end function
	
	'' 등록시 상태..
	Public function GetRegDetailStateName_off()
        if (Fregdetailstate="2") then
            if (Fisupchebeasong="Y") then
		        GetRegDetailStateName_off = "업체통보"
		    else
		        GetRegDetailStateName_off = "매장통보"
		    end if
	    elseif Fregdetailstate="3" then
		    GetRegDetailStateName_off = "상품준비"
	    elseif Fregdetailstate="7" then
		    GetRegDetailStateName_off = "출고완료"
	    else
		    GetRegDetailStateName_off = "----"
	    end if
	end Function
	
    public function GetDefaultRegNo_off(IsRegState)
        if (IsRegState) then
            GetDefaultRegNo_off = Fitemno
        else
            GetDefaultRegNo_off = Fregitemno
        end if
    end function
	
    public function getMiSendCodeColor_off()
		if FMisendReason="05" then
			getMiSendCodeColor_off = "#FF0000"
		else
			getMiSendCodeColor_off = "#000000"
		end if
	end function
	
	public function getMiSendCodeName_off()
		if FCode="00" then
			getMiSendCodeName_off = "입력대기"
		elseif FCode="01" then
			getMiSendCodeName_off = "재고부족" ''사용안함
		elseif FCode="02" then
			getMiSendCodeName_off = "주문제작"
		elseif FCode="03" then
			getMiSendCodeName_off = "출고지연"
		elseif FCode="04" then
			getMiSendCodeName_off = "예약상품" ''"포장대기" ''사용안함
		elseif FCode="05" then
			getMiSendCodeName_off = "품절출고불가"
		elseif FCode="06" then
			getMiSendCodeName_off = "신상품입고지연" ''사용안함
		else
			getMiSendCodeName_off = "&nbsp;"
		end if
	end function
	
    public function getBeasongDPlusDate_off()
        getBeasongDPlusDate_off = ""
        
        if IsNULL(Fbaljudate) then
            exit function
        end if
        
        if IsNULL(Fbeasongdate) then
            getBeasongDPlusDate_off = DateDiff("d",Fbaljudate,now())
            exit function
        end if
        
        getBeasongDPlusDate_off = DateDiff("d",Fbaljudate,Fbeasongdate)
    end function
    
    public function getBeasongDPlusDateStr_off()
        getBeasongDPlusDateStr_off = ""
        
        if IsNULL(Fbaljudate) then
            exit function
        end if
        
        if IsNULL(Fbeasongdate) then
            getBeasongDPlusDateStr_off = "D+" & DateDiff("d",Fbaljudate,now())
            exit function
        end if
        
        if (DateDiff("d",Fbaljudate,Fbeasongdate)<1) then
            getBeasongDPlusDateStr_off = "D+0"
        else
            getBeasongDPlusDateStr_off = "D+" & DateDiff("d",Fbaljudate,Fbeasongdate)
        end if
    end function 

	''반품 프로세스(회수, 맞교환 회수)
	public function fnIsReturnProcess_off(idivcd)
	    fnIsReturnProcess_off = (idivcd = "A004") or (idivcd = "A010") or (idivcd = "A011")
	end function

    ''반품 프로세스
    public function IsReturnProcess_off()
        IsReturnProcess_off = fnIsReturnProcess_off(Fdivcd)
    end function

	''취소 프로세스
	public function fnIsCancelProcess_off(idivcd)
	    fnIsCancelProcess_off = (idivcd = "A008")
	end function

    ''취소 프로세스
    public function IsCancelProcess_off()
        IsCancelProcess_off = fnIsCancelProcess_off(Fdivcd)
    end function
    
    public function IsAsRegAvail_off(byval iIpkumdiv, byval iCancelYn, byref descMsg)   
    
    'response.write Fdivcd & "<Br>!!!!!"    
        IsAsRegAvail_off = false
        if (iIpkumdiv<1) then
            IsAsRegAvail_off = false
            descMsg      = "실패한 주문건 또는 정상 주문건이 아닙니다. "
            exit function
        end if

        if (IsCancelProcess_off) then
            IsAsRegAvail_off = false

            if (iCancelYn<>"N") then
                IsAsRegAvail_off = false
                descMsg      = "이미 취소된 거래입니다. - 취소 불가능 "
                exit function
            end if

            if (iIpkumdiv=8) then
                IsAsRegAvail_off = false
                descMsg      = "출고완료 이후에는 회수요청/반품접수 만 가능합니다. - 취소 불가능 "
                exit function
            end if

            IsAsRegAvail_off = true
        elseif (IsReturnProcess_off) then
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail_off = false
                descMsg      = "출고 완료/ 일부 출고 상태가 아닙니다. - 반품 접수 불가능 "
                exit function
            end if

            if (iCancelYn<>"N") then
                IsAsRegAvail_off = false
                descMsg      = "취소된 거래입니다. - 반품 접수 불가능 "
                exit function
            end if

            IsAsRegAvail_off = true
        
        '' 출고시 유의사항        
        elseif (Fdivcd = "A006") then            
            IsAsRegAvail_off = true

            if (iIpkumdiv>=8) then
                IsAsRegAvail_off = false
                descMsg      = "출고 이전 상태가 아닙니다. - 출고시 유의사항 접수 불가능 "
                exit function
            end if
        
        '' 기타사항
        elseif (Fdivcd = "A009") then            
            IsAsRegAvail_off = true
        
        ''서비스발송 :모두 가능하게 변경..
        elseif  (Fdivcd = "A002") then            
            IsAsRegAvail_off = true
        
        ''누락재발송,
        elseif (Fdivcd = "A001") then            
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail_off = false
                descMsg      = "출고 완료/ 일부 출고 상태가 아닙니다. - 누락/서비스 발송 접수 불가능 "
                exit function
            end if

            IsAsRegAvail_off = true
        
        ''맞교환
        elseif (Fdivcd = "A000") then            
            if Not ((iIpkumdiv=7) or (iIpkumdiv=8)) then
                IsAsRegAvail_off = false
                descMsg      = "출고 완료/ 일부 출고 상태가 아닙니다. - 맞교환 접수 불가능 "
                exit function
            end if

            IsAsRegAvail_off = true
        
        ''환불요청    
        elseif (Fdivcd = "A003") then            
            IsAsRegAvail_off = true
        
        ''접수시 사이트 구분 체크
        elseif (Fdivcd = "A005") then            
            IsAsRegAvail_off = true
        
        ''업체 기타 정산.
        elseif (Fdivcd = "A700") then            
            IsAsRegAvail_off = true
        else
            descMsg = "정의 되지 않았습니다." + Fdivcd
        end if
    end function
	
	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function
    
    ''결제했는지 여부
    public function IsPayedOrder()
        IsPayedOrder = (FIpkumdiv>3) and (FIpkumdiv<9)
    end function
	
	Public function GetStateName()
        if FCurrState="2" then
            GetStateName = "업체통보"
	    elseif FCurrState="3" then
		    GetStateName = "상품준비"
	    elseif FCurrState="7" then
		    GetStateName = "출고완료"
		elseif FCurrState="0" then
		    GetStateName = ""
	    else
		    GetStateName = FCurrState
	    end if
	 end Function
	 
	public function GetStateColor()
	    if FCurrState="2" then
			GetStateColor="#000000"
		elseif FCurrState="3" then
			GetStateColor="#CC9933"
		elseif FCurrState="7" then
			GetStateColor="#FF0000"
		else
			GetStateColor="#000000"
		end if
	end function
	
	public function CancelStateColor()
		if FCancelYn="D" then
			CancelStateColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelStateColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelStateColor = "#000000"
		elseif UCase(FCancelYn)="A" then
			CancelStateColor = "#0000FF"
		end if
	end function

    public function GetMasterDeliveryName()
        GetMasterDeliveryName = ""
        if IsNULL(Fsongjangdiv) then Exit function
        
        if Fsongjangdiv="24" then
            GetMasterDeliveryName = "사가와"
        elseif Fsongjangdiv="2" then
            GetMasterDeliveryName = "현대"
        else
            GetMasterDeliveryName = Fsongjangdiv
        end if
    end function

	public function CancelYnName()
		CancelYnName = "정상"

		if Fcancelyn="Y" then
			CancelYnName ="취소"
		elseif Fcancelyn="D" then
			CancelYnName ="삭제"
		elseif Fcancelyn="A" then
			CancelYnName ="추가"
		end if
	end function

	public function CancelYnColor()
		CancelYnColor = "#000000"

		if FCancelYn="D" then
			CancelYnColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelYnColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelYnColor = "#000000"
		end if
	end function

	public function shopIpkumDivColor()
		if Fipkumdiv="0" then
			shopIpkumDivColor="#000000"
		elseif Fipkumdiv="1" then
			shopIpkumDivColor="#000000"
		elseif Fipkumdiv="2" then
			shopIpkumDivColor="#000000"
		elseif Fipkumdiv="3" then
			shopIpkumDivColor="#000000"
		elseif Fipkumdiv="4" then
			shopIpkumDivColor="#0000FF"
		elseif Fipkumdiv="5" then
			shopIpkumDivColor="#0000FF"
		elseif Fipkumdiv="6" then
			shopIpkumDivColor="#444400"
		elseif Fipkumdiv="7" then
			shopIpkumDivColor="#EE2222"
		elseif Fipkumdiv="8" then
			shopIpkumDivColor="#FF00FF"
		end if
	end function

	Public function shopIpkumDivName()
		if Fipkumdiv="1" then
			shopIpkumDivName="배송지입력전"
		elseif Fipkumdiv="2" then
			shopIpkumDivName="배송지입력완료"
		elseif Fipkumdiv="5" then
			shopIpkumDivName="업체통보"
		elseif Fipkumdiv="6" then
			shopIpkumDivName="배송준비"
		elseif Fipkumdiv="7" then
			shopIpkumDivName="일부출고"
		elseif Fipkumdiv="8" then
			shopIpkumDivName="출고완료"
		end if
	end Function

	Public function shopNormalUpcheDeliverState()
		if IsNull(FCurrState) then
			shopNormalUpcheDeliverState = ""
		elseif FCurrState="0" then
			shopNormalUpcheDeliverState = "배송대기"
		elseif FCurrState="2" then
			shopNormalUpcheDeliverState = "업체통보"
		elseif FCurrState="3" then
			shopNormalUpcheDeliverState = "업체확인"
		elseif FCurrState="7" then
			shopNormalUpcheDeliverState = "출고완료"
		else
			shopNormalUpcheDeliverState = ""
		end if
	end Function

	public function shopNormalUpcheDeliverStateColor()
		if FCurrState="0" then
			shopNormalUpcheDeliverStateColor="#000000"
		elseif FCurrState="1" then
			shopNormalUpcheDeliverStateColor="#000000"
		elseif FCurrState="2" then
			shopNormalUpcheDeliverStateColor="#000000"
		elseif FCurrState="3" then
			shopNormalUpcheDeliverStateColor="#000000"
		elseif FCurrState="4" then
			shopNormalUpcheDeliverStateColor="#0000FF"
		elseif FCurrState="5" then
			shopNormalUpcheDeliverStateColor="#0000FF"
		elseif FCurrState="6" then
			shopNormalUpcheDeliverStateColor="#444400"
		elseif FCurrState="7" then
			shopNormalUpcheDeliverStateColor="#EE2222"
		elseif FCurrState="8" then
			shopNormalUpcheDeliverStateColor="#FF00FF"
		end if
	end function
  
    public function shopGetCurrstateName()
        shopGetCurrstateName = FcurrstateName
    end function

     public function shopGetCurrstateColor()
        shopGetCurrstateColor = FcurrstateColor
    end function
    
    public function shopGetAsDivCDName()
        shopGetAsDivCDName = FdivcdName
    end function

    public function shopGetAsDivCDColor()
        shopGetAsDivCDColor = FdivcdName
    end function

    ''송장 필드가 필요한 정보
    public function IsRequireSongjangNO()
        IsRequireSongjangNO = false

        IsRequireSongjangNO = (Fdivcd="A000") or (Fdivcd="A001") or (Fdivcd="A002") or (Fdivcd="A004") or (Fdivcd="A010") or (Fdivcd="A011")
    end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COrder
	public FOneItem
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FRectRegStart
	public FRectRegEnd
	public FRectOldOrder
	public FRectorderno
	public FRectBuyname
	public FRectReqName	
	public FRectBuyHp
	public FRectReqHp
	public FRectBuyPhone
	public FRectReqPhone	
	public frectshopid
	public FRectNotCsID
	public FRectmasteridx
	public frectdetailidx
	public FRectSearchType
	public FRectUserName
	public FRectMakerid
	public FRectDivcd
	public FRectCurrstate
	public FRectWriteUser
	public FRectDeleteYN	
	public FRectStartDate
	public FRectEndDate
	public FRectCsAsID
    public IsUpchebeasongExists
    public IsTenbeasongExists
    public FRectDeliveryNo
	public FRectOnlyJupsu
	
	'//admin/offshop/cscenter/order/misendmaster_main.asp
    public Sub fgetMiSendOrderDetailList()
        dim SqlStr, i , sqlsearch
		
		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch + " and m.masteridx='" + CStr(FRectmasteridx) + "'"
		end if

		sqlStr = "select"
		sqlStr = sqlStr + " m.cancelyn, m.regdate, m.baljudate, m.buyname, m.reqname , m.buyemail,m.buyhp" + vbcrlf
		sqlStr = sqlStr + " ,d.masteridx,d.detailidx, d.itemno, d.orderno, d.itemid, d.itemoption,d.itemgubun" + vbcrlf
		sqlStr = sqlStr + " , d.upcheconfirmdate, d.makerid, d.isupchebeasong,isNull(d.currstate,0) as currstate" + vbcrlf
		sqlStr = sqlStr + " , d.beasongdate, d.songjangno, d.songjangdiv,d.cancelyn as detailcancelyn" + vbcrlf
		sqlStr = sqlStr + " ,T.code, T.state, T.ipgodate,IsNULL(T.isSendSMS,'N') as isSendSMS" + vbcrlf
		sqlStr = sqlStr + " ,IsNULL(T.isSendEmail,'N') as isSendEmail,IsNULL(T.isSendCall,'N') as isSendCall" + vbcrlf
		sqlStr = sqlStr + " ,T.reqstr, IsNULL(T.itemlackno,0) as itemlackno, T.finishstr" + vbcrlf
		sqlStr = sqlStr + " , p.company_name, p.tel as company_tel" + vbcrlf
		sqlStr = sqlStr + " , od.itemname,od.itemoptionname" + vbcrlf
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr + " 	on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr + "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr & " left join [db_shop].dbo.tbl_shopbeasong_mibeasong_list T" +vbcrlf
		sqlStr = sqlStr & " 	on d.detailidx=T.detailidx" +vbcrlf
		sqlStr = sqlStr & " Left Join [db_partner].dbo.tbl_partner p" +vbcrlf
		sqlStr = sqlStr & " 	on d.makerid=p.id" +vbcrlf
		sqlStr = sqlStr & " where m.ipkumdiv >= '5' " & sqlsearch
		sqlStr = sqlStr & " order by d.makerid, d.itemid, d.itemoption" +vbcrlf
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new COrderItem

    			FItemList(i).fdetailidx				  = rsget("detailidx")
    			FItemList(i).forderno		  = rsget("orderno")
				FItemList(i).fitemgubun 			  = rsget("itemgubun")
    			FItemList(i).FItemid 			  = rsget("itemid")
    			FItemList(i).FItemoption     	  = rsget("itemoption")
    			FItemList(i).FItemname 		      = db2html(rsget("itemname"))
    			FItemList(i).FItemoptionName      = db2html(rsget("itemoptionname"))
    			FItemList(i).fitemno             = rsget("itemno")    			
    			FItemList(i).FMakerid 			  = rsget("makerid")
    			FItemList(i).FBuyname             = db2html(rsget("buyname"))
    			FItemList(i).FReqname			  = db2html(rsget("reqname"))
    			FItemList(i).FCancelYn		      = rsget("cancelyn")
    			FItemList(i).FDetailCancelYn	  = rsget("detailcancelyn")
    			FItemList(i).FRegdate			  = rsget("regdate")    			
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
                FItemList(i).FMisendReason        = rsget("code")
                FItemList(i).FMisendState         = rsget("state")
                FItemList(i).FMisendipgodate      = rsget("ipgodate")                
                FItemList(i).FisSendSMS           = rsget("isSendSMS")
                FItemList(i).FisSendEmail         = rsget("isSendEmail")
                FItemList(i).FisSendCall          = rsget("isSendCall")
                FItemList(i).Fbuyemail            = rsget("buyemail")
                FItemList(i).FbuyHp               = rsget("buyHp")                
                FItemList(i).FrequestString       = db2Html(rsget("reqstr"))
                FItemList(i).FItemNo              = rsget("itemno")
                FItemList(i).Fitemlackno          = rsget("itemlackno")
                FItemList(i).FfinishString        = db2Html(rsget("finishstr"))                               
                FItemList(i).Fcompany_name        = db2Html(rsget("company_name"))
                FItemList(i).Fcompany_tel         = db2Html(rsget("company_tel"))
                FItemList(i).FCancelYn            = rsget("detailcancelyn")
			
			rsget.movenext
			i=i+1
		loop
		rsget.close
    end Sub
	
	'//admin/offshop/cscenter/order/misendmaster_main.asp
	public sub fGetOneOrderMasterWithCS
		dim sqlStr,i , sqlsearch
		
		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch & " and m.masteridx='" + FRectmasteridx + "'"
		else
			sqlsearch = sqlsearch + " and m.deliverno='" + FRectDeliveryNo + "'"	
		end if
		
		sqlStr = " select top 1"
		sqlStr = sqlStr + " m.orderno, m.cancelyn, m.buyname, m.buyhp, m.buyemail"
		sqlStr = sqlStr + " ,m.masteridx"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + VbCrlf
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		set FOneItem = new COrderItem
		if Not rsget.Eof then
			
			FOneItem.fmasteridx = rsget("masteridx")
			FOneItem.forderno = rsget("orderno")
			FOneItem.FCancelyn    = rsget("cancelyn")			
			FOneItem.Fbuyname    = db2Html(rsget("buyname"))
			FOneItem.Fbuyhp    = rsget("buyhp")
			FOneItem.Fbuyemail    = db2Html(rsget("buyemail"))
			
		end if

		rsget.Close
	end sub
    	
	'//admin/offshop/cscenter/action/inc_cs_action_item_list.asp
    public Sub fGetOrderDetailByCsDetail()
        dim SqlStr, i , sqlsearch
		
		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch + " and d.masteridx='" + CStr(FRectmasteridx) + "'"
		end if
		
		sqlStr = "select"
		sqlStr = sqlStr + " d.detailidx as orderdetailidx, d.orderno,d.itemid,d.itemoption,d.itemgubun"
		sqlStr = sqlStr + " ,d.itemno,d.cancelyn ,d.makerid, d.songjangno,d.beasongdate"		
		sqlStr = sqlStr + " ,d.currstate as orderdetailcurrstate, d.upcheconfirmdate, d.songjangdiv"
		sqlStr = sqlStr + " , d.isupchebeasong, d.cancelyn , d.odlvType,d.cancelorgdetailidx"		
		sqlStr = sqlStr + " ,od.itemname, od.itemoptionname ,od.sellprice"
		sqlStr = sqlStr + " ,IsNULL(c.regitemno,0) as regitemno, IsNULL(c.confirmitemno,0) as confirmitemno"
		sqlStr = sqlStr + " ,c.detailidx, c.masteridx, c.regdetailstate"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_detail d "
		sqlStr = sqlStr + "	join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx" +vbcrlf    
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shopbeasong_cs_detail c "
		sqlStr = sqlStr + " 	on c.masteridx='" + CStr(FRectCsAsID) + "'"
		sqlStr = sqlStr + " 	and d.detailidx = c.orderdetailidx "    
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
        sqlStr = sqlStr + " order by d.isupchebeasong, d.makerid, d.itemid, d.itemoption"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new COrderItem
			
			FItemList(i).fcancelorgdetailidx       = rsget("cancelorgdetailidx")
            FItemList(i).fdetailidx       = rsget("detailidx")
            FItemList(i).fmasteridx       = rsget("masteridx")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget("regdetailstate")
            FItemList(i).Forderdetailidx  = rsget("orderdetailidx")
            FItemList(i).forderno     = rsget("orderno")
            FItemList(i).Fitemid          = rsget("itemid")
            FItemList(i).Fitemoption      = rsget("itemoption")
            FItemList(i).fitemgubun      = rsget("itemgubun")
            FItemList(i).Fmakerid         = rsget("makerid")
            FItemList(i).Fitemname        = db2html(rsget("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
            FItemList(i).fsellprice        = rsget("sellprice")            
            FItemList(i).Fitemno          = rsget("itemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")
            FItemList(i).FCancelyn        = rsget("cancelyn")
            FItemList(i).ForderDetailcurrstate = rsget("orderdetailcurrstate")            
            FItemList(i).FodlvType        = rsget("odlvType")
			FItemList(i).fcurrstate = rsget("orderdetailcurrstate")
            
            if (FItemList(i).Fitemid=0) then                
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if
			
			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub

	'//admin/offshop/cscenter/action/cs_action_detail.asp
    public sub fGetReturnAddress()
        dim i,sqlStr , sqlsearch
		
        if FRectMakerid <> "" then
            sqlsearch = sqlsearch + " and id='"&FRectMakerid&"'"
        end if		         
	
        sqlStr = " select top 1"
        sqlStr = sqlStr + " company_name, deliver_phone, deliver_hp, return_zipcode, return_address, return_address2"
        sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner"
        sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount

        if not rsget.EOF  then
            set FOneItem = new COrderItem

            FOneItem.FreturnName      = db2html(rsget("company_name"))
            FOneItem.FreturnPhone     = db2html(rsget("deliver_phone"))
            FOneItem.Freturnhp        = db2html(rsget("deliver_hp"))
            FOneItem.FreturnZipcode   = rsget("return_zipcode")
            FOneItem.FreturnZipaddr   = db2html(rsget("return_address"))
            FOneItem.FreturnEtcaddr   = db2html(rsget("return_address2"))
            FOneItem.Fsongjangdiv     = ""
            FOneItem.Fsongjangno      = ""
		end if
        rsget.close
    end sub
	   
	'//admin/offshop/cscenter/action/cs_action_detail.asp '//admin/offshop/cscenter/action/pop_cs_action_new.asp
	'//admin/offshop/cscenter/action/popChangeSongjang.asp '//common/offshop/beasong/upche_csdetail.asp
    public Sub fGetOneCSASMaster()
        dim i,sqlStr , sqlsearch
		
        if FRectMakerID <> "" then   ''업체 조회용.
            sqlsearch = sqlsearch + " and A.makerid='"&FRectMakerID&"'"
        end if		
        if FRectCsAsID <> "" then   ''업체 조회용.
            sqlsearch = sqlsearch + " and a.masteridx='"&FRectCsAsID&"'"
        end if	
            
        if FRectmasteridx <> "" then
            sqlsearch = sqlsearch + " and a.orgmasteridx='"&FRectmasteridx&"'"
        end if	
        	
        sqlStr = " select top 1"
		sqlStr = sqlStr + " a.masteridx,a.orgmasteridx,a.divcd,a.orderno,a.customername,a.writeuser"
		sqlStr = sqlStr + " ,a.finishuser,a.title,a.contents_jupsu,a.contents_finish,a.currstate"
		sqlStr = sqlStr + " ,a.regdate,a.finishdate,a.deleteyn,a.opentitle,a.opencontents"
		sqlStr = sqlStr + " ,a.requireupche,a.makerid,a.songjangdiv,a.songjangno"
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C4.comm_name as currstatename"
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_cs_master A "
        sqlStr = sqlStr + " Left Join [db_shop].[dbo].tbl_cs_comm_code_off C1"
        sqlStr = sqlStr + "  on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_shop].[dbo].tbl_cs_comm_code_off C4"
        sqlStr = sqlStr + "  on A.currstate=C4.comm_cd"
        sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount
        ftotalcount = rsget.RecordCount

        if not rsget.EOF  then
            set FOneItem = new COrderItem
			
			FOneItem.forgmasteridx			= rsget("orgmasteridx")
            FOneItem.fmasteridx				= rsget("masteridx")
            FOneItem.Fdivcd               = rsget("divcd")                        
            FOneItem.FdivcdName           = db2html(rsget("divcdname"))
            FOneItem.forderno         = rsget("orderno")
            FOneItem.Fcustomername        = db2html(rsget("customername"))            
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
            FOneItem.Fopentitle           = db2html(rsget("opentitle"))
            FOneItem.Fopencontents        = db2html(rsget("opencontents"))            
            FOneItem.Fsongjangdiv         = rsget("songjangdiv")
            FOneItem.Fsongjangno          = rsget("songjangno")
            FOneItem.Frequireupche        = rsget("requireupche")
            FOneItem.Fmakerid             = rsget("makerid")

        end if
        rsget.close
    end sub
	
	'//admin/offshop/cscenter/action/cs_action_list.asp '//admin/offshop/cscenter/action/inc_cs_action_prev_cslist.asp
	'//common/offshop/beasong/upche_cslist.asp	'//common/offshop/beasong/shop_cslist.asp
    public Sub fGetCSASMasterList()
        dim i,sqlStr, sqlsearch

		if (FRectSearchType="") then
	        if (frectmasteridx <> "") then
				sqlsearch = sqlsearch + " and A.orgmasteridx='" + CStr(frectmasteridx) + "' "
	        end if
		
		''업체가 쿼리시
		elseif (FRectSearchType="upcheview") then		    
            sqlsearch = sqlsearch + " and a.divcd not in ('A005','A007')"		'/외부몰환불요청	'/카드/이체/휴대폰취소요청 제외
            sqlsearch = sqlsearch + " and a.deleteyn='N'"
            sqlsearch = sqlsearch + " and a.requireupche='Y' "
            sqlsearch = sqlsearch + " and a.makerid='" + CStr(FRectMakerid) + "' "

	        if (frectmasteridx <> "") then
				sqlsearch = sqlsearch + " and A.orgmasteridx='" + CStr(frectmasteridx) + "' "
	        end if
	        
            if (FRectOnlyJupsu="on") then
                sqlsearch = sqlsearch + " and currstate='B001'"		'/접수상태
            end if

            if (FRectCurrstate = "notfinish") then
				sqlsearch = sqlsearch + " and A.currstate < 'B007' "		'/완료상태
	        elseif (FRectCurrstate <> "") then
				sqlsearch = sqlsearch + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if

            if (FRectUserName <> "") then
				sqlsearch = sqlsearch + " and A.customername='" + CStr(FRectUserName) + "' "
	        end if

	        if (frectorderno <> "") then
				sqlsearch = sqlsearch + " and A.orderno='" + CStr(frectorderno) + "' "
	        end if

		''매장 쿼리시
		elseif (FRectSearchType="shopview") then		    
            sqlsearch = sqlsearch + " and a.divcd not in ('A005','A007')"		'/외부몰환불요청	'/카드/이체/휴대폰취소요청 제외
            sqlsearch = sqlsearch + " and a.deleteyn='N'"
            sqlsearch = sqlsearch + " and a.requiremaejang='Y'"            

	        if (frectmasteridx <> "") then
				sqlsearch = sqlsearch + " and A.orgmasteridx='" + CStr(frectmasteridx) + "' "
	        end if
	        
            if (FRectOnlyJupsu="on") then
                sqlsearch = sqlsearch + " and currstate='B001'"		'/접수상태
            end if

            if (FRectCurrstate = "notfinish") then
				sqlsearch = sqlsearch + " and A.currstate < 'B007' "		'/완료상태
	        elseif (FRectCurrstate <> "") then
				sqlsearch = sqlsearch + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if

            if (FRectUserName <> "") then
				sqlsearch = sqlsearch + " and A.customername='" + CStr(FRectUserName) + "' "
	        end if

	        if (frectorderno <> "") then
				sqlsearch = sqlsearch + " and A.orderno='" + CStr(frectorderno) + "' "
	        end if

			if frectshopid <> "" then
				sqlsearch = sqlsearch & " and om.shopid ='"&frectshopid&"'" +vbcrlf
			end if
	        
		elseif (FRectSearchType = "searchfield") then
	        if (FRectUserName <> "") then
				sqlsearch = sqlsearch + " and A.customername='" + CStr(FRectUserName) + "' "
	        end if

	        if (frectmasteridx <> "") then
				sqlsearch = sqlsearch + " and A.orgmasteridx=" + CStr(frectmasteridx) + ""
	        end if

	        if (FRectMakerid<>"") then
				sqlsearch = sqlsearch + " and A.requireupche='Y' "
				sqlsearch = sqlsearch + " and A.makerid='" + CStr(FRectMakerid) + "' "
	        end if

	        if (FRectStartDate <> "") then
				sqlsearch = sqlsearch + " and A.regdate>='" + CStr(FRectStartDate) + "' "
	        end if

	        if (FRectEndDate <> "") then
				sqlsearch = sqlsearch + " and A.regdate <'" + CStr(FRectEndDate) + "' "
	        end if

	        if (FRectCurrstate = "notfinish") then
				sqlsearch = sqlsearch + " and A.currstate < 'B007' "
	        elseif (FRectCurrstate <> "") then
				sqlsearch = sqlsearch + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if

	        if (FRectDivcd <> "") then
				sqlsearch = sqlsearch + " and A.divcd ='" + CStr(FRectDivcd) + "' "
	        end if

			if (FRectWriteUser <> "") then
				sqlsearch = sqlsearch + " and A.writeUser = '" + CStr(FRectWriteUser) + "' "
			end if

			if (FRectDeleteYN <> "") then
				sqlsearch = sqlsearch + " and A.deleteyn = '" + CStr(FRectDeleteYN) + "' "
			end if

        end If
        
        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_cs_master A"
        sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_master om"
        sqlStr = sqlStr + " 	on a.orgmasteridx = om.masteridx"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u" + vbcrlf
		sqlStr = sqlStr & " 	on om.shopid = u.userid and u.isusing='Y'" + vbcrlf        
        sqlStr = sqlStr + " Left Join [db_shop].[dbo].tbl_cs_comm_code_off C1"
        sqlStr = sqlStr + " 	on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_shop].[dbo].tbl_cs_comm_code_off C4"
        sqlStr = sqlStr + " 	on A.currstate=C4.comm_cd"        
        sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1

        if  not rsget.EOF  then
            FTotalCount = rsget("cnt")
        else
            FTotalCount = 0
        end if
        rsget.close
        
        sqlStr = " select Top " + CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr + " A.masteridx, A.divcd, A.orderno, A.customername, A.title,a.orgmasteridx"
        sqlStr = sqlStr + " ,A.regdate, A.finishdate,A.deleteyn, A.finishuser, A.writeuser"
        sqlStr = sqlStr + " ,A.requireupche, A.makerid, A.songjangdiv ,A.songjangno, A.currstate"
        sqlStr = sqlStr + " ,a.contents_jupsu ,a.contents_finish"
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C4.comm_name as currstatename"
        sqlStr = sqlStr + " ,C4.comm_color as currstatecolor , u.shopname"
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_cs_master A"
        sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_master om"
        sqlStr = sqlStr + " 	on a.orgmasteridx = om.masteridx"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u" + vbcrlf
		sqlStr = sqlStr & " 	on om.shopid = u.userid and u.isusing='Y'" + vbcrlf        
        sqlStr = sqlStr + " Left Join [db_shop].[dbo].tbl_cs_comm_code_off C1"
        sqlStr = sqlStr + " 	on A.divcd=C1.comm_cd"
        sqlStr = sqlStr + " Left Join [db_shop].[dbo].tbl_cs_comm_code_off C4"
        sqlStr = sqlStr + " 	on A.currstate=C4.comm_cd"
        sqlStr = sqlStr + " where 1=1 " & sqlsearch
        sqlStr = sqlStr + " order by a.masteridx desc "

		'response.write sqlStr &"<br>"
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
                set FItemList(i) = new COrderItem
				
				FItemList(i).fshopname = rsget("shopname")
				FItemList(i).fcontents_jupsu         = db2html(rsget("contents_jupsu"))
				FItemList(i).fcontents_finish         = db2html(rsget("contents_finish"))
                FItemList(i).forgmasteridx = rsget("orgmasteridx")
                FItemList(i).fmasteridx = rsget("masteridx")
                FItemList(i).Fdivcd             = rsget("divcd")
                FItemList(i).FdivcdName         = db2html(rsget("divcdname"))
                FItemList(i).forderno       = rsget("orderno")
                FItemList(i).Fcustomername      = db2html(rsget("customername"))            
                FItemList(i).Fwriteuser         = rsget("writeuser")
                FItemList(i).Ffinishuser        = rsget("finishuser")
                FItemList(i).Ftitle             = db2html(rsget("title"))
                FItemList(i).Fcurrstate         = rsget("currstate")
                FItemList(i).Fcurrstatename     = rsget("currstatename")
                FItemList(i).FcurrstateColor    = rsget("currstatecolor")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Ffinishdate        = rsget("finishdate")
                FItemList(i).Fdeleteyn          = rsget("deleteyn")                                
                FItemList(i).Fsongjangdiv       = rsget("songjangdiv")
                FItemList(i).Fsongjangno        = rsget("songjangno")
                FItemList(i).Frequireupche      = rsget("requireupche")
                FItemList(i).Fmakerid           = rsget("makerid")

                rsget.MoveNext
                i = i + 1
            loop
        end if
        rsget.close
    end sub
	
	'//admin/offshop/cscenter/order/orderdetailedit.asp
	public Sub fSearchOneJumunDetail()
        dim sqlStr ,sqlsearch ,i

		if frectdetailidx <> "" then
			sqlsearch = sqlsearch + " and d.detailidx="&frectdetailidx&""
		end if
	
		sqlStr = sqlStr + " select top 1 " +vbcrlf
		sqlStr = sqlStr + " d.detailidx,d.masteridx,d.orgdetailidx,d.orderno,d.itemgubun,d.itemid" +vbcrlf
		sqlStr = sqlStr + " ,d.itemoption,d.makerid,d.itemno,d.cancelyn,d.currstate,d.songjangno" +vbcrlf
		sqlStr = sqlStr + " ,d.songjangdiv,d.beasongdate,d.isupchebeasong,d.omwdiv,d.odlvType" +vbcrlf
		sqlStr = sqlStr + " ,d.upcheconfirmdate,d.lastupdateadminid,d.passday" +vbcrlf
		sqlStr = sqlStr + " ,convert(varchar(19),d.upcheconfirmdate,21) as cvupcheconfirmdate" +vbcrlf
		sqlStr = sqlStr + " ,convert(varchar(19),d.beasongdate,21) as cvbeasongdate" +vbcrlf
		sqlStr = sqlStr + " ,i.shopitemprice as currsellcash , od.itemname ,od.itemoptionname,od.sellprice"
		sqlStr = sqlStr + " from [db_shop].dbo.tbl_shopbeasong_order_detail d"
		sqlStr = sqlStr + "	join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "	on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr + "	left join db_shop.dbo.tbl_shop_item i" +vbcrlf
		sqlStr = sqlStr + "	on d.itemid = i.shopitemid" +vbcrlf
		sqlStr = sqlStr + "	and d.itemgubun = i.itemgubun" +vbcrlf
		sqlStr = sqlStr + "	and d.itemoption = i.itemoption" +vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch	

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount

        set FOneItem = new COrderItem

        if Not rsget.Eof then
        	
			FOneItem.fmasteridx = rsget("masteridx")
			FOneItem.forderno = rsget("orderno")
			FOneItem.fdetailidx			= rsget("detailidx")
			FOneItem.Fmakerid      = rsget("makerid")
			FOneItem.Fitemid      = rsget("itemid")
			FOneItem.Fitemoption  = rsget("itemoption")
			FOneItem.fitemgubun  = rsget("itemgubun")
			FOneItem.Fitemno      = rsget("itemno")
			FOneItem.fsellprice    = rsget("sellprice")
			FOneItem.Fcancelyn    = rsget("cancelyn")
			FOneItem.FItemName    = db2html(rsget("itemname"))
			FOneItem.FItemoptionName = db2html(rsget("itemoptionname"))
			FOneItem.Fcurrstate     = rsget("currstate")
			FOneItem.Fsongjangdiv   = rsget("songjangdiv")
			FOneItem.Fsongjangno    = rsget("songjangno")
			FOneItem.Fupcheconfirmdate = rsget("cvupcheconfirmdate")
			FOneItem.Fbeasongdate   = rsget("cvbeasongdate")
			FOneItem.Fisupchebeasong= rsget("isupchebeasong")			
			FOneItem.FcurrSellcash	= rsget("currsellcash")
            FOneItem.FODlvType      = rsget("odlvtype")

        end if
        rsget.Close
    end Sub
	
	'//admin/offshop/cscenter/action/pop_CsDeliveryEdit.asp
    public Sub GetOneCsDeliveryItemFromDefaultOrder()
        dim i,sqlStr

        sqlStr = " select m.reqname, m.reqphone, m.reqhp, m.reqzipcode, m.reqzipaddr, m.reqaddress"
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m"
        sqlStr = sqlStr + "	Join db_shop.dbo.tbl_shopbeasong_cs_master a"
        sqlStr = sqlStr + " 	on m.masteridx=a.orgmasteridx"
        sqlStr = sqlStr + "     and a.masteridx=" + CStr(FRectCsAsID) + " "
		
		'response.write sqlStr &"<br>"
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

        end if
        rsget.close
    end Sub
    
	'//admin/offshop/cscenter/action/cs_action_detail.asp '//admin/offshop/cscenter/action/pop_CsDeliveryEdit.asp
    public Sub fGetOneCsDeliveryItem()
        dim i,sqlStr ,sqlsearch

        if FRectCsAsID <> "" then
        	sqlsearch = sqlsearch + " and asid = "&FRectCsAsID&""
        end if
        
        sqlStr = " select top 1"
        sqlStr = sqlStr & " a.asid , a.reqname ,a.reqphone ,a.reqhp ,a.reqzipcode ,a.reqzipaddr"
        sqlStr = sqlStr & " ,a.reqetcaddr ,a.reqetcstr,a.songjangdiv ,a.songjangno ,a.regdate ,a.senddate"
        sqlStr = sqlStr + " from [db_shop].dbo.tbl_shopbeasong_cs_delivery A"
        sqlStr = sqlStr + " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount

        if not rsget.EOF  then
            set FOneItem = new COrderItem
            
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
	
	'//admin/offshop/cscenter/order/order_process.asp '//admin/offshop/cscenter/action/cs_action_detail.asp
	'//admin/offshop/cscenter/action/inc_cs_action_item_list.asp '//admin/offshop/cscenter/cscenter_mail_Function_off.asp
	'//common/offshop/beasong/upche_csdetail.asp '//common/offshop/beasong/shop_csdetail.asp
    public Sub fGetCsDetailList()
        dim SqlStr, i , sqlsearch
        
        if FRectCsAsID <> "" then
        	sqlsearch = sqlsearch + " and c.masteridx = "&FRectCsAsID&""
        end if
		
		sqlStr = "select"
		sqlStr = sqlStr + " c.detailidx,c.masteridx,c.orderdetailidx,c.jumundetailidx,c.orderno"
		sqlStr = sqlStr + " ,c.itemid,c.itemoption,c.itemgubun,c.makerid,c.regitemno,c.confirmitemno"
		sqlStr = sqlStr + " ,c.orderitemno,c.isupchebeasong,c.regdetailstate,c.currstate"
		sqlStr = sqlStr + " ,d.currstate as orderdetailcurrstate ,d.odlvType ,d.cancelorgdetailidx"
		sqlStr = sqlStr + " ,IsNULL(od.sellprice,0) as Ordersellprice"
		sqlStr = sqlStr + " ,od.itemname , od.itemoptionname , od.sellprice"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_cs_detail c"
	    sqlStr = sqlStr + " left join db_shop.dbo.tbl_shopbeasong_order_detail d"
	    sqlStr = sqlStr + " 	on c.orderdetailidx=d.detailidx"
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr & "		on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr & "	where 1=1 " & sqlsearch
        sqlStr = sqlStr + " order by c.isupchebeasong, c.makerid, c.itemid, c.itemoption"
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new COrderItem
			
			FItemList(i).fcancelorgdetailidx              = rsget("cancelorgdetailidx")
            FItemList(i).fdetailidx              = rsget("detailidx")
            FItemList(i).fmasteridx        = rsget("masteridx")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget("regdetailstate")   ''접수 당시 진행 상태
            FItemList(i).Forderdetailidx  = rsget("orderdetailidx")
            FItemList(i).forderno     = rsget("orderno")
            FItemList(i).Fitemid          = rsget("itemid")
            FItemList(i).Fitemoption      = rsget("itemoption")
            FItemList(i).fitemgubun      = rsget("itemgubun")
            FItemList(i).Fmakerid         = rsget("makerid")
            FItemList(i).Fitemname        = db2html(rsget("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))            
            FItemList(i).Fitemno          = rsget("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")
			FItemList(i).fOrdersellprice = rsget("Ordersellprice")                    
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Forderdetailcurrstate  = rsget("orderdetailcurrstate")
			FItemList(i).fsellprice  = rsget("sellprice")
			
            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if
        		
			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub
	
	'//admin/offshop/cscenter/order/ordermaster_detail.asp
    public Sub fGetCSASTotalCount()
        dim i,sqlStr,sqlsearch

		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch + " and orgmasteridx=" + FRectmasteridx + ""
		end if

        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_cs_master"
        sqlStr = sqlStr + " where 1=1 " & sqlsearch
		
		'response.write sqlStr&"<br>"
        rsget.Open sqlStr, dbget, 1

        if not rsget.EOF  then
            FResultCount = rsget("cnt")
        else
            FResultCount = 0
        end if
        rsget.close
    end sub
	
	'//admin/offshop/cscenter/order/orderitemmaster.asp '//common/pop_order_receipt.asp
	public Sub fQuickSearchOrderDetail()
		dim sqlStr, i , sqlsearch
		
		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch & " and d.masteridx='" + CStr(FRectmasteridx) + "'"
		end if
		
		sqlStr = "select"
		sqlStr = sqlStr + " d.masteridx, d.detailidx, d.orderno,d.itemid,d.itemoption,d.itemno"
		sqlStr = sqlStr + " ,d.itemgubun,d.cancelyn , d.makerid, d.beasongdate, d.isupchebeasong"		
		sqlStr = sqlStr + " ,d.currstate, d.upcheconfirmdate, d.songjangdiv, d.songjangno"		
		sqlStr = sqlStr + " ,od.sellprice, od.realsellprice,od.itemname ,od.itemoptionname"
		sqlStr = sqlStr + " ,s.divname as songjangdivname, s.findurl"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_detail d "
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr & "		on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr + " left join [db_order].[dbo].tbl_songjang_div s"
		sqlStr = sqlStr + " 	on d.songjangdiv=s.divcd"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
        sqlStr = sqlStr + " order by d.isupchebeasong, d.makerid, d.itemid, d.itemoption"
        
        'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new COrderItem

			FItemList(i).forderno = rsget("orderno")
			FItemList(i).fdetailidx         = rsget("detailidx")
			FItemList(i).fmasteridx         = rsget("masteridx")
			FItemList(i).Fmakerid     = rsget("makerid")
			FItemList(i).fitemgubun      = rsget("itemgubun")
			FItemList(i).Fitemid      = rsget("itemid")
			FItemList(i).Fitemoption  = rsget("itemoption")
			FItemList(i).Fitemno      = rsget("itemno")
			FItemList(i).frealsellprice    = rsget("realsellprice")
			FItemList(i).fsellprice    = rsget("sellprice")
			FItemList(i).Fcancelyn    = rsget("cancelyn")
			FItemList(i).FItemName    = db2html(rsget("itemname"))

			if IsNull(rsget("itemoptionname")) then
				FItemList(i).FItemoptionName = "-"
			else
				FItemList(i).FItemoptionName = db2html(rsget("itemoptionname"))
			end if

			FItemList(i).Fcurrstate         = rsget("currstate")
			FItemList(i).Fsongjangdiv       = rsget("songjangdiv")
			FItemList(i).Fsongjangno        = rsget("songjangno")
			FItemList(i).Fbeasongdate       = rsget("beasongdate")
			FItemList(i).Fisupchebeasong    = rsget("isupchebeasong")			
			FItemList(i).Fupcheconfirmdate    = rsget("upcheconfirmdate")			
			FItemList(i).Fsongjangdivname  = db2html(rsget("songjangdivname"))
            FItemList(i).Ffindurl          = db2html(rsget("findurl"))                                    
            
            if Not IsNULL(FItemList(i).Fsongjangno) then
               FItemList(i).Fsongjangno = replace(FItemList(i).Fsongjangno,"-","")
            end if
            
			rsget.movenext
			i=i+1
		loop
		rsget.close
	end sub

	'//admin/offshop/cscenter/order/ordermaster_detail.asp '/admin/offshop/cscenter/order/order_receiver_info.asp
	'//admin/offshop/cscenter/order/orderitemmaster.asp '//admin/offshop/cscenter/order/orderdetailedit.asp
	'//admin/offshop/cscenter/action/cs_action_detail.asp '//admin/offshop/cscenter/action/pop_cs_action_new.asp
	'//admin/offshop/cscenter/action/pop_cs_action_new_process.asp '//common/pop_order_receipt.asp
	public Sub fQuickSearchOrderMaster()
        dim sqlStr ,sqlsearch

		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch & " and m.masteridx="& FRectmasteridx &""
		end if
	
		sqlStr = " select top 1 " +vbcrlf
    	sqlStr = sqlStr + " m.masteridx, m.orderno, m.shopid, m.ipkumdiv, m.regdate, m.beadaldiv" +vbcrlf
    	sqlStr = sqlStr + " ,m.beadaldate ,m.cancelyn ,m.buyname ,m.buyphone ,m.buyhp ,m.buyemail" +vbcrlf
    	sqlStr = sqlStr + " ,m.reqname ,m.reqzipcode ,m.reqzipaddr ,m.reqaddress ,m.reqphone" +vbcrlf
    	sqlStr = sqlStr + " ,m.reqhp ,m.comment ,m.lastupdateadminid ,m.baljudate ,m.cancelorgorderno,u.shopname" +vbcrlf 	
    	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user u" +vbcrlf
		sqlStr = sqlStr + " 	on m.shopid = u.userid and u.isusing='Y'" +vbcrlf    
		sqlStr = sqlStr + " where 1=1 " & sqlsearch	
		sqlStr = sqlStr + " order by m.orderno desc"  

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount

        set FOneItem = new COrderItem

        if Not rsget.Eof then
			
			FOneItem.fcancelorgorderno           = rsget("cancelorgorderno")
			FOneItem.fmasteridx           = rsget("masteridx")
			FOneItem.forderno           = rsget("orderno")
			FOneItem.Fipkumdiv	            = rsget("ipkumdiv")
			FOneItem.Fregdate		        = rsget("regdate")
			FOneItem.Fbaljudate		        = rsget("baljudate")
			FOneItem.Fbeadaldate	        = rsget("beadaldate")
			FOneItem.Fcancelyn	            = rsget("cancelyn")
			FOneItem.Fbuyname		        = db2Html(rsget("buyname"))
			FOneItem.Fbuyphone	            = rsget("buyphone")
			FOneItem.Fbuyhp		            = rsget("buyhp")
			FOneItem.Fbuyemail	            = rsget("buyemail")
			FOneItem.Freqname		        = db2Html(rsget("reqname"))
			FOneItem.Freqzipcode	        = rsget("reqzipcode")
			FOneItem.Freqaddress	        = db2Html(rsget("reqaddress"))
			FOneItem.Freqphone	            = rsget("reqphone")
			FOneItem.Freqhp		            = rsget("reqhp")			
			FOneItem.Fcomment		        = db2Html(rsget("comment"))
			FOneItem.Freqzipaddr		    = db2Html(rsget("reqzipaddr"))
			FOneItem.fshopname  	= rsget("shopname")

        end if
        rsget.Close
    end Sub
	
	'//admin/offshop/cscenter/order/ordermaster_list.asp
	public Sub fQuickSearchOrderList()
		dim sqlStr, i , sqlsearch
		
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid='"&frectshopid&"'"
		end if

		if (FRectorderno<>"") then
			sqlsearch = sqlsearch + " and m.orderno='" + FRectorderno + "'"
		end if
        
		if (FRectRegStart<>"") then
			sqlsearch = sqlsearch + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlsearch = sqlsearch + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if
		
		if (FRectBuyname<>"") then
			sqlsearch = sqlsearch + " and m.reqname = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectReqName<>"") then
			sqlsearch = sqlsearch + " and m.reqname = '" + FRectReqName + "'"  ''like
		end if

		if (FRectBuyHp<>"") then
			sqlsearch = sqlsearch + " and m.buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			sqlsearch = sqlsearch + " and m.reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			sqlsearch = sqlsearch + " and m.buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			sqlsearch = sqlsearch + " and m.reqphone='" + FRectReqPhone + "'"
		end if

		''갯수
		sqlStr = "select count(*) as cnt "
    	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.close

		''데이타.
		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
    	sqlStr = sqlStr + " m.masteridx, m.orderno, m.shopid, m.ipkumdiv, m.regdate, m.beadaldiv" +vbcrlf
    	sqlStr = sqlStr + " ,m.beadaldate ,m.cancelyn ,m.buyname ,m.buyphone ,m.buyhp ,m.buyemail" +vbcrlf
    	sqlStr = sqlStr + " ,m.reqname ,m.reqzipcode ,m.reqzipaddr ,m.reqaddress ,m.reqphone" +vbcrlf
    	sqlStr = sqlStr + " ,m.reqhp ,m.comment ,m.lastupdateadminid ,m.baljudate ,u.shopname" +vbcrlf 	
    	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u" +vbcrlf
		sqlStr = sqlStr & " 	on m.shopid = u.userid and u.isusing='Y'" +vbcrlf    	
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by m.masteridx desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if FResultCount<1 then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderItem
				
				FItemList(i).fmasteridx  	= rsget("masteridx")
				FItemList(i).fshopname  	= rsget("shopname")
				FItemList(i).Forderno       = rsget("orderno")
				FItemList(i).Fipkumdiv	        = rsget("ipkumdiv")			
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Fbaljudate		= rsget("baljudate")
				FItemList(i).Fbeadaldate	= rsget("beadaldate")
				FItemList(i).Fcancelyn	        = rsget("cancelyn")				
				FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
				FItemList(i).Fbuyphone	        = rsget("buyphone")
				FItemList(i).Fbuyhp		= rsget("buyhp")
				FItemList(i).Fbuyemail	        = rsget("buyemail")
				FItemList(i).Freqname		= db2Html(rsget("reqname"))				
				FItemList(i).Freqzipcode	= rsget("reqzipcode")
				FItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
				FItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
				FItemList(i).Freqphone	        = rsget("reqphone")
				FItemList(i).Freqhp		= rsget("reqhp")				
				FItemList(i).Fcomment		= db2Html(rsget("comment"))                
                
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	Private Sub Class_Initialize()
		Redim FItemList(0)
		FCurrPage =1
		FPageSize = 20
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