<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.22 한용민 생성
'###########################################################

class cupchebeasong_item
	public fidx
	public fmasteridx
	public forderno
	public fitemgubun
	public fitemid
	public fitemoption
	public fitemname
	public fitemoptionname
	public fsellprice
	public frealsellprice
	public fsuplyprice
	public fitemno
	public fmakerid
	public fjungsanid
	public fcancelyn
	public fshopidx
	public fitempoint
	public fdiscountKind
	public fdiscountprice
	public fshopbuyprice
	public faddtaxcharge
	public fzoneidx
	public fshopid
	public ftotalsum
	public frealsum
	public fjumundiv
	public fjumunmethod
	public fshopregdate
	public fregdate
	public fspendmile
	public fpointuserno
	public fgainmile
	public ftableno
	public fcashsum
	public fcardsum
	public fGiftCardPaySum
	public fcasherid
	public fCashReceiptNo
	public fCardAppNo
	public fCashreceiptGubun
	public fCardInstallment
	public fIXyyyymmdd
	public fcurrstate
	public fshopname
	public fisupchebeasong
	public fomwdiv
	public fodlvType
	public fipkumdiv
	public fbeadaldiv
	public fbeadaldate
	public fbuyname
	public fbuyphone
	public fbuyhp
	public fbuyemail
	public freqname
	public freqzipcode
	public freqzipaddr
	public freqaddress
	public freqphone
	public freqhp
	public fcomment
	public fipgono
	public frealstock
	public fdetailidx
	public FBaljudate
	public Fsongjangno
	public Fsongjangdiv
	public FMisendReason
	public FMisendState
	public FMisendipgodate
	public Fupcheconfirmdate
	public Fbeasongdate
	public fdetailcancelyn
	public FisSendSMS
	public FisSendEmail
	public FisSendCall
	public FrequestString
	public Fitemlackno
	public FfinishString
	public Fcompany_name
	public Fcompany_tel
	public Fsmallimage
	public Fdefaultbeasongdiv		'배송구분 2:업체배송, 0:매장배송, 1:물류배송
	public fupchesendsms
	public fshopphone
	public fuserSeq
	public fUserName
	public fOnLineUSerID
	public fshopbeasongD_cancelyn
	public fshopbeasongM_cancelyn
	public fAuthIdx
	public fBeaSongcnt
	public fUserHp
	public fSmsYN
	public fKakaoTalkYN
	public fIsUsing
	public fLastUpdate
	public fregdate_beasong
	public fitemno_beasong
	public fmasteridx_beasong
	public fdetailidx_beasong
	public fcomm_cd
	public fCertNo

	Public function getDefaultBeasongDivName()
		if Fdefaultbeasongdiv="0" then
			getDefaultBeasongDivName="매장배송"
		elseif Fdefaultbeasongdiv="1" then
			getDefaultBeasongDivName="물류배송"
		elseif Fdefaultbeasongdiv="2" then
			getDefaultBeasongDivName="업체배송"
		else
			getDefaultBeasongDivName = Fdefaultbeasongdiv
		end if
	end Function

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

	'//물류센타 재고 상태
	public function logicsstockyn()
		if fipgono > 0 or frealstock >0 then
			logicsstockyn = "Y"
		else
			logicsstockyn = "<font color='red'>N</font>"
		end if
	end Function

    public function getDlvCompanyName()
        if FIsUpchebeasong="Y" then
            getDlvCompanyName = Fcompany_name
        else
            getDlvCompanyName = "텐바이텐"
        end if
    end function

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

	'//배송구분(odlvType) 텐바이텐 배송이냐 업체 배송이냐..
	public function getbeasonggubun()
		if fodlvType = "0" then
			getbeasonggubun = "매장배송"
		elseif fodlvType = "1" then
			getbeasonggubun = "물류배송"
		'elseif fodlvType = "4" then
		'	getbeasonggubun = "텐바이텐무료배송"
		elseif fodlvType = "2" then
			getbeasonggubun = "업체배송"
		'elseif fodlvType = "7" then
		'	getbeasonggubun = "업체착불배송"
		else
			getbeasonggubun = "설정안됨"
		end if

	end function

    public function isMisendAlreadyInputed()
        isMisendAlreadyInputed = Not (IsNULL(FMisendReason) or (FMisendReason="00") or (FMisendReason=""))
    end function

	'/미출고사유
    public function getMisendText()
        select Case FMisendReason
            CASE "00" : getMisendText = "입력대기"
            CASE "01" : getMisendText = "재고부족"
            CASE "04" : getMisendText = "예약상품"
            'CASE "02" : getMisendText = "주문제작"
            'CASE "52" : getMisendText = "주문제작"
            CASE "03" : getMisendText = "출고지연"
            CASE "53" : getMisendText = "출고지연"
            CASE "05" : getMisendText = "품절출고불가"
            CASE "55" : getMisendText = "품절출고불가"
            CASE ELSE : getMisendText = FMisendReason
        end Select
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
                    smstext = "[텐바이텐 출고지연안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"주문제작 상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품이 "&VbCrlf
                    smstext = smstext&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."
                end if
            elseif (FMisendReason="03") then  ''출고지연
                ''출고 소요일수 D+2이상
                if (getMisendDPlusDate>1) then
                    smstext = "[텐바이텐 출고지연안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품이 "&VbCrlf
                    smstext = smstext&FMisendipgodate&"에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품이 "&VbCrlf
                    smstext = smstext&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."

                end if
            elseif (FMisendReason="04") then  ''예약상품
                ''출고 소요일수 D+2이상
                if (getMisendDPlusDate>1) then
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"예약배송상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."
                else
                ''출고 소요일수 D+0/D+1
                    smstext = "[텐바이텐 출고예정안내]주문하신 상품중 "&DdotFormat(FItemName,16)&"("&FItemID&")상품은 "&VbCrlf
                    smstext = smstext&"예약배송상품으로 "&FMisendipgodate&"에 발송될 예정입니다. 감사합니다."

                end if
            end if
        end if
        getSMSText = smstext
    end function

	public function IsAvailJumun_off()
		IsAvailJumun_off = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class cupchebeasong_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount
	public frectorderno
	public frectmasteridx
	public frectC_ADMIN_USER
	public frectipkumdiv
	public FRectDesignerID
	public frectdetailidxarr
	public FRectIsAll
	public FRectSearchType
	public FRectSearchValue
	public FRectMisendReason
	public FRectRegStart
	public FRectRegEnd
	public FRectDetailIDx
	public frectshopid
	public FRectDateType
	public FRectIsUpcheBeasong
	public FRectBuyname
	public FRectReqName
	public FRectItemID
	public frectmakerid
	public frectmasteridx_beasong
	public FRectStartDay
	public FRectEndDay
	public frectreqhp

	'//common/offshop/beasong/upche_jumunlist.asp
	public Sub fSearchJumunListByDesigner()
		dim sqlStr , sqlsearch ,i

		if (frectorderno<>"") then
			sqlsearch = sqlsearch + " and m.orderno='" + frectorderno + "'"
		end if

		if (FRectBuyname<>"") then
			sqlsearch = sqlsearch + " and m.buyname = '" + FRectBuyname + "'"
		end if

		if (FRectReqName<>"") then
			sqlsearch = sqlsearch + " and m.reqname = '" + FRectReqName + "'"
		end if

		if (FRectRegStart<>"") then
			if FRectDateType="upbeasongdate" then
				sqlsearch = sqlsearch + " and ((d.isupchebeasong='Y') and (d.beasongdate >='" + CStr(FRectRegStart) + "')) "
			else
				sqlsearch = sqlsearch + " and m.regdate >='" + CStr(FRectRegStart) + "'"
			end if
		end if

		if (FRectRegEnd<>"") then
			if FRectDateType="upbeasongdate" then
				sqlsearch = sqlsearch + " and ((d.isupchebeasong='Y') and (d.beasongdate <'" + CStr(FRectRegEnd) + "')) "
			else
				sqlsearch = sqlsearch + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
			end if
		end if

		if (FRectItemID<>"") then
			sqlsearch = sqlsearch + " and d.itemid="&FRectItemID&""
		end if

        if (FRectIsUpcheBeasong<>"") then
            sqlsearch = sqlsearch + " and d.isupchebeasong='"&FRectIsUpcheBeasong&"'"
		end if

        if (FRectDesignerID<>"") then
            sqlsearch = sqlsearch + " and d.makerid='" + FRectDesignerID + "'"
		end if
		
		''총 갯수
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m"
		sqlStr = sqlStr + " Join [db_shop].dbo.tbl_shopbeasong_order_detail d"
		sqlStr = sqlStr + "     on m.masteridx=d.masteridx"
		sqlStr = sqlStr + "	join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx"
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & "		on m.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr & "	left join db_shop.dbo.tbl_shop_designer s"
		sqlStr = sqlStr & "		on m.shopid=s.shopid"
		sqlStr = sqlStr & " 	and d.makerid=s.makerid"	
		sqlStr = sqlStr + " where m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and m.ipkumdiv>'1' " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1	    
		    FTotalCount = rsget("cnt")		  		    	
		rsget.Close

		''데이타.
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " m.orderno,m.masteridx, m.buyname,m.reqname, m.cancelyn"
		sqlStr = sqlStr + " ,m.ipkumdiv, m.regdate, m.reqphone, m.reqhp"
		sqlStr = sqlStr + " ,d.itemid, d.itemoption, d.itemno , d.itemgubun"
		sqlStr = sqlStr + " ,d.beasongdate,d.isupchebeasong, d.currstate,d.detailidx"
		sqlStr = sqlStr + " ,od.itemname, od.itemoptionname,od.sellprice ,od.realsellprice ,od.shopbuyprice" + vbcrlf
		sqlStr = sqlStr & " ,(CASE when od.suplyprice=0 THEN convert(int,od.sellprice*(100-s.defaultmargin)/100)"
		sqlStr = sqlStr & " 	 ELSE od.suplyprice END) as suplyprice"
		sqlStr = sqlStr & " , u.shopname" & vbcrlf
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m"
		sqlStr = sqlStr + " Join [db_shop].dbo.tbl_shopbeasong_order_detail d"
		sqlStr = sqlStr + "     on m.masteridx=d.masteridx"
		sqlStr = sqlStr + "	join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx"
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & "		on m.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr & "	left join db_shop.dbo.tbl_shop_designer s"
		sqlStr = sqlStr & "		on m.shopid=s.shopid"
		sqlStr = sqlStr & " 	and d.makerid=s.makerid"
		sqlStr = sqlStr + " where m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and m.ipkumdiv>'1' " & sqlsearch
		sqlStr = sqlStr + " order by d.detailidx desc"

		'response.write sqlStr &"<br>"		
		rsget.Open sqlStr,dbget,1		
		
		rsget.pagesize = FPageSize
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
    		do until rsget.eof
    			set FItemList(i) = new cupchebeasong_item

				FItemList(i).fsellprice = rsget("sellprice")
				FItemList(i).frealsellprice = rsget("realsellprice")
				FItemList(i).fshopbuyprice = rsget("shopbuyprice")
				FItemList(i).fsuplyprice = rsget("suplyprice")    			
    			FItemList(i).forderno = rsget("orderno")
    			FItemList(i).fmasteridx	= rsget("masteridx")
    			FItemList(i).fdetailidx	= rsget("detailidx") 			
    			FItemList(i).Fipkumdiv	= rsget("ipkumdiv")    			
    			FItemList(i).Fregdate		= rsget("regdate")
    			FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
    			FItemList(i).Freqname		= db2Html(rsget("reqname"))
    			FItemList(i).Freqphone	= rsget("reqphone")
    			FItemList(i).Freqhp		= rsget("reqhp")
    			FItemList(i).FCancelyn	= rsget("cancelyn")
				FItemList(i).fitemgubun       = rsget("itemgubun")
    			FItemList(i).FItemID       = rsget("itemid")
    			FItemList(i).FItemName     = db2Html(rsget("itemname"))
    			FItemList(i).FItemOption   = rsget("itemoption")
    			FItemList(i).fitemoptionname= db2Html(rsget("itemoptionname"))
    			FItemList(i).FItemNo     = rsget("itemno")
    			FItemList(i).fbeasongdate     = rsget("beasongdate")
    			FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")    			
    			FItemList(i).FCurrState		 = rsget("currstate")
    			FItemList(i).fshopname    = db2html(rsget("shopname"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//common/offshop/beasong/upche_popMisendInput.asp
    public function fOneOldMisendItem()
        dim sqlStr , sqlsearch

		if FRectDetailidx<>"" then
			sqlsearch = sqlsearch + " and d.detailidx='" + FRectDetailidx + "'"
		end if

		if FRectDesignerID <> "" then
			sqlsearch = sqlsearch + " and d.makerid='" + FRectDesignerID + "'"
		end if

		sqlStr = "select top 1" + vbcrlf
		sqlStr = sqlStr + " m.cancelyn, m.regdate,m.buyhp, m.baljudate, m.buyname, m.buyemail, m.reqhp" + vbcrlf
		sqlStr = sqlStr + " ,d.detailidx, d.itemno, m.orderno, d.itemid, d.itemoption,d.itemgubun" + vbcrlf
		sqlStr = sqlStr + " , d.upcheconfirmdate, d.makerid, d.isupchebeasong" + vbcrlf
		sqlStr = sqlStr + " ,isNull(d.currstate,0) as currstate, d.beasongdate, d.songjangno" + vbcrlf
		sqlStr = sqlStr + " , d.songjangdiv, d.cancelyn as detailcancelyn" + vbcrlf
		sqlStr = sqlStr + " ,T.code, T.state, T.ipgodate, IsNULL(T.isSendSMS,'N') as isSendSMS" + vbcrlf
		sqlStr = sqlStr + " ,IsNULL(T.isSendEmail,'N') as isSendEmail, IsNULL(T.isSendCall,'N') as isSendCall" + vbcrlf
		sqlStr = sqlStr + " ,T.reqstr, IsNULL(T.itemlackno,0) as itemlackno, T.finishstr, p.company_name" + vbcrlf
		sqlStr = sqlStr + " , p.tel as company_tel" + vbcrlf
		sqlStr = sqlStr + " , od.itemname,od.itemoptionname" + vbcrlf
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr + " on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr + "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "	on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr & " left join [db_shop].dbo.tbl_shopbeasong_mibeasong_list T" +vbcrlf
		sqlStr = sqlStr & " on d.detailidx=T.detailidx" +vbcrlf
		sqlStr = sqlStr & " Left Join [db_partner].dbo.tbl_partner p" +vbcrlf
		sqlStr = sqlStr & " on d.makerid=p.id" +vbcrlf
		sqlStr = sqlStr & " where m.ipkumdiv >= '5'" +vbcrlf
		sqlStr = sqlStr & " and d.currstate<7 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount

		if not rsget.EOF then
            set FOneItem = new cupchebeasong_item

			FOneItem.fdetailidx				  = rsget("detailidx")
			FOneItem.forderno		  = rsget("orderno")
			FOneItem.FItemid 			  = rsget("itemid")
			FOneItem.fitemgubun 			  = rsget("itemgubun")
			FOneItem.FItemoption     	  = rsget("itemoption")
			FOneItem.FItemname 		      = db2html(rsget("itemname"))
			FOneItem.FItemoptionName      = db2html(rsget("itemoptionname"))
			FOneItem.FItemno            = rsget("itemno")
			FOneItem.FMakerid 			  = rsget("makerid")
			FOneItem.FBuyname             = db2html(rsget("buyname"))
			FOneItem.FCancelYn		      = rsget("cancelyn")
			FOneItem.FDetailCancelYn	  = rsget("detailcancelyn")
			FOneItem.FRegdate			  = rsget("regdate")
			FOneItem.FBaljudate		      = rsget("baljudate")
			FOneItem.Fupcheconfirmdate    = rsget("upcheconfirmdate")
			FOneItem.FCurrstate		      = rsget("currstate")
			FOneItem.Fbeasongdate         = rsget("beasongdate")
			FOneItem.FisUpcheBeasong      = rsget("isUpcheBeasong")
			FOneItem.FSongjangno          = rsget("songjangno")
			FOneItem.FSongjangdiv         = rsget("songjangdiv")
            FOneItem.FMisendReason        = rsget("code")
            FOneItem.FMisendState         = rsget("state")
            FOneItem.FMisendipgodate      = rsget("ipgodate")
            FOneItem.FisSendSMS           = rsget("isSendSMS")
            FOneItem.FisSendEmail         = rsget("isSendEmail")
            FOneItem.FisSendCall          = rsget("isSendCall")
            FOneItem.Fbuyemail            = rsget("buyemail")
            FOneItem.FbuyHp               = rsget("buyHp")
            FOneItem.freqhp               = rsget("reqhp")
            FOneItem.FrequestString       = db2Html(rsget("reqstr"))
            FOneItem.Fitemlackno          = rsget("itemlackno")
            FOneItem.FfinishString        = db2Html(rsget("finishstr"))
            FOneItem.Fcompany_name        = db2Html(rsget("company_name"))
            FOneItem.Fcompany_tel         = db2Html(rsget("company_tel"))
        end if
        rsget.Close
    end function

	'//common/offshop/beasong/upche_viewordermaster.asp
	public Sub fSearchJumunList()
		dim sqlStr , sqlsearch ,i

		if FRectmasteridx<>"" then
			sqlsearch = sqlsearch + " and m.masteridx='" + FRectmasteridx + "'"
		end if

		if FRectDesignerID <> "" then
			sqlsearch = sqlsearch + " and d.makerid='" + FRectDesignerID + "'"
		end if

        if (FRectIpkumdiv<>"") then
            sqlsearch = sqlsearch + FRectIpkumdiv
        end if

		''총 갯수
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr + " 	on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & "		on m.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr + "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr & " left join [db_shop].dbo.tbl_shopbeasong_mibeasong_list T" & vbcrlf
		sqlStr = sqlStr & " 	on d.detailidx=T.detailidx" & vbcrlf
		sqlStr = sqlStr + " where m.masteridx<>0 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		''데이타.
		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.masteridx, m.orderno, m.shopid, m.ipkumdiv, m.beadaldiv, m.beadaldate, m.cancelyn"
		sqlStr = sqlStr + " ,m.buyname, m.buyphone, m.buyhp, m.buyemail, m.reqname, m.reqzipcode, m.reqzipaddr"
		sqlStr = sqlStr + " ,m.reqaddress, m.reqphone, m.reqhp, m.comment, m.lastupdateadminid, m.baljudate"
		sqlStr = sqlStr + " ,convert(varchar,m.regdate,20) as regdate"
		sqlStr = sqlStr + " ,d.itemid , d.Itemoption ,d.itemno,d.itemgubun,d.odlvType,(d.cancelyn) as detailcancelyn"
		sqlStr = sqlStr + " ,d.currstate"
		sqlStr = sqlStr + " ,od.sellprice,od.itemname ,od.itemoptionname"
		sqlStr = sqlStr & " , T.code, T.state, T.ipgodate, u.shopname"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr + " 	on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & "		on m.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr + "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr & " left join [db_shop].dbo.tbl_shopbeasong_mibeasong_list T" & vbcrlf
		sqlStr = sqlStr & " 	on d.detailidx=T.detailidx" & vbcrlf
		sqlStr = sqlStr + " where m.masteridx<>0 " & sqlsearch
		sqlStr = sqlStr + " order by m.masteridx desc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new cupchebeasong_item

                FItemList(i).FMisendReason     = rsget("code")
                FItemList(i).FMisendState      = rsget("state")
                FItemList(i).FMisendipgodate   = rsget("ipgodate")
				FItemList(i).fcurrstate	  = rsget("currstate")
				FItemList(i).fdetailcancelyn 			  = rsget("detailcancelyn")
				FItemList(i).fodlvType 			  = rsget("odlvType")
				FItemList(i).fsellprice 			  = rsget("sellprice")
				FItemList(i).fitemgubun 			  = rsget("itemgubun")
    			FItemList(i).FItemid 			  = rsget("itemid")
    			FItemList(i).FItemoption 			  = rsget("Itemoption")
    			FItemList(i).FItemname 		  = db2html(rsget("itemname"))
    			FItemList(i).fitemoptionname     	  = db2html(rsget("itemoptionname"))
    			FItemList(i).fitemno           = rsget("itemno")
				FItemList(i).forderno = rsget("orderno")
				FItemList(i).Fipkumdiv	= rsget("ipkumdiv")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Fbeadaldiv	= rsget("beadaldiv")
				FItemList(i).Fbeadaldate	= rsget("beadaldate")
				FItemList(i).Fcancelyn	= rsget("cancelyn")
				FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
				FItemList(i).Fbuyphone	= rsget("buyphone")
				FItemList(i).Fbuyhp		= rsget("buyhp")
				FItemList(i).Fbuyemail	= rsget("buyemail")
				FItemList(i).Freqname		= db2Html(rsget("reqname"))
				FItemList(i).Freqzipcode	= rsget("reqzipcode")
				FItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
				FItemList(i).Freqphone	= rsget("reqphone")
				FItemList(i).Freqhp		= rsget("reqhp")
				FItemList(i).Fcomment		= db2Html(rsget("comment"))
				FItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
    			FItemList(i).fshopname    = db2html(rsget("shopname"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/common/offshop/beasong/shop_maejangbeasong.asp
	public sub fshop_maejangbaesong()
		dim sqlStr,i ,sqlsearch

		if FRectSearchType = "orderno" then
			sqlsearch = sqlsearch & " and m.orderno='"&FRectSearchValue&"'"
		elseif FRectSearchType = "buyname" then
			sqlsearch = sqlsearch & " and m.buyname='"&FRectSearchValue&"'"
		elseif FRectSearchType = "reqname" then
			sqlsearch = sqlsearch & " and m.reqname='"&FRectSearchValue&"'"
		elseif FRectSearchType = "itemid" then
			sqlsearch = sqlsearch & " and d.itemid='"&FRectSearchValue&"'"
		end if

		if FRectRegStart <> "" and FRectRegEnd <> "" then
			sqlsearch = sqlsearch & " and d.beasongdate>='"&FRectRegStart&"'"
			sqlsearch = sqlsearch & " and d.beasongdate<'"&FRectRegEnd&"'"
			sqlsearch = sqlsearch & " and d.currstate = '7'"
		else
			sqlsearch = sqlsearch & " and d.currstate <> '7'" +vbcrlf
		end if

		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and om.shopid='"&frectshopid&"'"
		end if

		sqlStr = "select"
		sqlStr = sqlStr & " m.regdate, m.baljudate, m.buyname, m.reqname,m.cancelyn,m.masteridx" +vbcrlf
		sqlStr = sqlStr & " ,d.detailidx, d.itemno, m.orderno, d.itemid, d.beasongdate,d.itemgubun" +vbcrlf
		sqlStr = sqlStr & " ,d.upcheconfirmdate,isNull(d.currstate,0) as currstate,d.Itemoption" +vbcrlf
		sqlStr = sqlStr & " ,d.songjangno, d.songjangdiv,d.cancelyn as detailcancelyn" +vbcrlf
		sqlStr = sqlStr & " ,od.itemname,od.itemoptionname, u.shopname" +vbcrlf
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr + " 	on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr + "	join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr + "	join [db_shop].[dbo].tbl_shopjumun_master om" +vbcrlf
		sqlStr = sqlStr + "		on om.orderno = od.orderno" +vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & "		on m.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr & " where m.ipkumdiv >= 5" +vbcrlf
		sqlStr = sqlStr & " and m.cancelyn = 'N'" +vbcrlf
		sqlStr = sqlStr & " and d.currstate = 3"
		sqlStr = sqlStr & " and d.itemid<>0"
		sqlStr = sqlStr & " and d.isupchebeasong='N'"
		sqlStr = sqlStr & " and d.cancelyn <> 'Y' " & sqlsearch
		sqlStr = sqlStr & " order by m.baljudate, d.detailidx"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cupchebeasong_item

				FItemList(i).fmasteridx = rsget("masteridx")
				FItemList(i).fitemgubun = rsget("itemgubun")
    			FItemList(i).fdetailidx				  = rsget("detailidx")
    			FItemList(i).forderno		  = rsget("orderno")
    			FItemList(i).FItemid 			  = rsget("itemid")
    			FItemList(i).FItemoption 			  = rsget("Itemoption")
    			FItemList(i).FItemname 		  = db2html(rsget("itemname"))
    			FItemList(i).fitemoptionname     	  = db2html(rsget("itemoptionname"))
    			FItemList(i).fitemno           = rsget("itemno")
    			FItemList(i).FBuyname           = db2html(rsget("buyname"))
    			FItemList(i).FReqname			  = db2html(rsget("reqname"))
    			FItemList(i).FCancelYn		  = rsget("cancelyn")
    			FItemList(i).FRegdate			  = rsget("regdate")
    			FItemList(i).FBaljudate		  = rsget("baljudate")
    			FItemList(i).Fupcheconfirmdate  = rsget("upcheconfirmdate")
    			FItemList(i).FCurrstate		  = rsget("currstate")
    			FItemList(i).Fbeasongdate       = rsget("beasongdate")
    			FItemList(i).FSongjangno        = rsget("songjangno")
    			FItemList(i).FSongjangdiv       = rsget("songjangdiv")
    			FItemList(i).fshopname    = db2html(rsget("shopname"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/common/offshop/beasong/upche_datebaljulist.asp '//common/offshop/beasong/upche_sendsongjanginput.asp
	'//common/offshop/beasong/upche_mibeasonglist.asp
	public sub fDesignerDateBaljuinputlist()
		dim sqlStr,i ,sqlsearch

		if FRectSearchType = "orderno" then
			sqlsearch = sqlsearch & " and m.orderno='"&FRectSearchValue&"'"
		elseif FRectSearchType = "buyname" then
			sqlsearch = sqlsearch & " and m.buyname='"&FRectSearchValue&"'"
		elseif FRectSearchType = "reqname" then
			sqlsearch = sqlsearch & " and m.reqname='"&FRectSearchValue&"'"
		elseif FRectSearchType = "itemid" then
			sqlsearch = sqlsearch & " and d.itemid='"&FRectSearchValue&"'"
		end if

		if FRectRegStart <> "" and FRectRegEnd <> "" then
			sqlsearch = sqlsearch & " and d.beasongdate>='"&FRectRegStart&"'"
			sqlsearch = sqlsearch & " and d.beasongdate<'"&FRectRegEnd&"'"
			sqlsearch = sqlsearch & " and d.currstate = '7'"
		else
			sqlsearch = sqlsearch & " and d.currstate <> '7'" +vbcrlf
		end if

		if FRectDesignerID <> "" then
			sqlsearch = sqlsearch & " and d.makerid='"&FRectDesignerID&"'"

		end if

		if FRectMisendReason <> "" and FRectMisendReason <> "AA" and FRectMisendReason <> "NN" then
			sqlsearch = sqlsearch & " and T.code='"&FRectMisendReason&"'"
		end if

		if FRectMisendReason = "NN" then
			sqlsearch = sqlsearch & " and ISNULL(T.code,'00')='00'"
		end if

		sqlStr = "select"
		sqlStr = sqlStr & " m.regdate, m.baljudate, m.buyname, m.reqname,m.cancelyn,m.masteridx" +vbcrlf
		sqlStr = sqlStr & " ,d.detailidx, d.itemno, m.orderno, d.itemid, d.beasongdate,d.itemgubun" +vbcrlf
		sqlStr = sqlStr & " ,d.upcheconfirmdate,isNull(d.currstate,0) as currstate,d.Itemoption" +vbcrlf
		sqlStr = sqlStr & " ,d.songjangno, d.songjangdiv,d.cancelyn as detailcancelyn" +vbcrlf
		sqlStr = sqlStr & " ,od.itemname,od.itemoptionname, od.sellprice ,od.realsellprice ,od.shopbuyprice"
		sqlStr = sqlStr & " ,(CASE when od.suplyprice=0 THEN convert(int,od.sellprice*(100-s.defaultmargin)/100)"
		sqlStr = sqlStr & " 	 ELSE od.suplyprice END) as suplyprice"		
		sqlStr = sqlStr & " ,T.code, T.state, T.ipgodate, u.shopname" & vbcrlf
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr + " 	on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & "		on m.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr + "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr & " left join [db_shop].dbo.tbl_shopbeasong_mibeasong_list T" +vbcrlf
		sqlStr = sqlStr & " 	on d.detailidx=T.detailidx" +vbcrlf
		sqlStr = sqlStr & "	left join db_shop.dbo.tbl_shop_designer s"
		sqlStr = sqlStr & "		on m.shopid=s.shopid"
		sqlStr = sqlStr & " 	and d.makerid=s.makerid"		
		sqlStr = sqlStr & " where m.ipkumdiv >= 5" +vbcrlf
		sqlStr = sqlStr & " and m.cancelyn = 'N'" +vbcrlf
		sqlStr = sqlStr & " and d.currstate = 3"
		sqlStr = sqlStr & " and d.itemid<>0"
		sqlStr = sqlStr & " and d.isupchebeasong='Y'"
		sqlStr = sqlStr & " and d.cancelyn <> 'Y' " & sqlsearch
		sqlStr = sqlStr & " order by m.baljudate, d.detailidx"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cupchebeasong_item

				FItemList(i).fsellprice = rsget("sellprice")
				FItemList(i).frealsellprice = rsget("realsellprice")
				FItemList(i).fshopbuyprice = rsget("shopbuyprice")
				FItemList(i).fsuplyprice = rsget("suplyprice")
				FItemList(i).fmasteridx = rsget("masteridx")
				FItemList(i).fitemgubun = rsget("itemgubun")
    			FItemList(i).fdetailidx				  = rsget("detailidx")
    			FItemList(i).forderno		  = rsget("orderno")
    			FItemList(i).FItemid 			  = rsget("itemid")
    			FItemList(i).FItemoption 			  = rsget("Itemoption")
    			FItemList(i).FItemname 		  = db2html(rsget("itemname"))
    			FItemList(i).fitemoptionname     	  = db2html(rsget("itemoptionname"))
    			FItemList(i).fitemno           = rsget("itemno")
    			FItemList(i).FBuyname           = db2html(rsget("buyname"))
    			FItemList(i).FReqname			  = db2html(rsget("reqname"))
    			FItemList(i).FCancelYn		  = rsget("cancelyn")
    			FItemList(i).FRegdate			  = rsget("regdate")
    			FItemList(i).FBaljudate		  = rsget("baljudate")
    			FItemList(i).Fupcheconfirmdate  = rsget("upcheconfirmdate")
    			FItemList(i).FCurrstate		  = rsget("currstate")
    			FItemList(i).Fbeasongdate       = rsget("beasongdate")
    			FItemList(i).FSongjangno        = rsget("songjangno")
    			FItemList(i).FSongjangdiv       = rsget("songjangdiv")
                FItemList(i).FMisendReason     = rsget("code")
                FItemList(i).FMisendState      = rsget("state")
                FItemList(i).FMisendipgodate   = rsget("ipgodate")
    			FItemList(i).fshopname    = db2html(rsget("shopname"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/common/offshop/beasong/upche_dobeasonglistCSV.asp '//common/offshop/beasong/upche_reselectbaljulist.asp
    public Sub fReDesignerSelectBaljuList()
		dim sqlStr,idxArr , i

        if (Right(frectdetailidxarr,1)=",") then frectdetailidxarr = left(frectdetailidxarr,len(frectdetailidxarr) - 1)

		''업체  발주서 재출력
		sqlStr = "select"
		sqlStr = sqlStr + "  m.orderno, m.buyname, m.reqzipcode, m.reqzipaddr, m.reqaddress, m.comment"
		sqlStr = sqlStr + " ,m.buyphone,m.buyhp, m.buyemail, m.reqname, m.reqphone, m.reqhp, m.regdate"
		sqlStr = sqlStr + " ,d.itemid,d.itemgubun, d.itemno, d.itemoption, d.songjangno,d.songjangdiv"
		sqlStr = sqlStr + " ,d.detailidx"
		sqlStr = sqlStr + " ,od.itemname , od.sellprice ,od.itemoptionname, u.shopname"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr + " 	on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & "		on m.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr + "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx" +vbcrlf
	    sqlStr = sqlStr + " where m.ipkumdiv>=5"
	    sqlStr = sqlStr + " and m.cancelyn='N'"
	    sqlStr = sqlStr + " and d.cancelyn<>'Y'"

	    ''전체출력할 경우. ''(idxArr<>"")조건 추가 선택내역이 없을수 있음.
	    if (FRectIsAll<>"on") and (frectdetailidxarr<>"") then
		    sqlStr = sqlStr + " and d.detailidx in (" & frectdetailidxarr & ")"
		end if

		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and d.currstate='3'"
		sqlStr = sqlStr + " order by m.baljudate, d.detailidx"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		    do until rsget.EOF
				set FItemList(i) = new cupchebeasong_item

				FItemList(i).forderno = rsget("orderno")
				FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
				FItemList(i).Freqzipcode	= rsget("reqzipcode")
				FItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
				FItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
				FItemList(i).Fcomment		= db2Html(rsget("comment"))
				FItemList(i).Fbuyphone	= rsget("buyphone")
				FItemList(i).Fbuyhp		= rsget("buyhp")
				FItemList(i).Fbuyemail	= rsget("buyemail")
				FItemList(i).Freqname		= db2Html(rsget("reqname"))
				FItemList(i).Freqphone	= rsget("reqphone")
				FItemList(i).Freqhp		= rsget("reqhp")
				FItemList(i).FRegDate   = rsget("regdate")
				FItemList(i).fitemgubun      = rsget("itemgubun")
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).FItemName    = db2html(rsget("itemname"))
				FItemList(i).Fitemno      = rsget("itemno")
				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).fsellprice  = rsget("sellprice")
				FItemList(i).Fsongjangno		= rsget("songjangno")

				if IsNull(rsget("itemoptionname")) then
				  FItemList(i).FItemoptionName = "-"
				else
				  FItemList(i).FItemoptionName = db2html(rsget("itemoptionname"))
				end if

                FItemList(i).Fdetailidx = rsget("detailidx")
                FItemList(i).Fsongjangdiv = rsget("songjangdiv")
    			FItemList(i).fshopname    = db2html(rsget("shopname"))

				rsget.movenext
				i=i+1

			loop

		rsget.Close
	end sub

	'//common/offshop/beasong/upche_selectbaljulist.asp
	public Sub fDesignerSelectBaljuList()
		dim sqlStr, i

        if (Right(frectdetailidxarr,1)=",") then frectdetailidxarr = left(frectdetailidxarr,len(frectdetailidxarr) - 1)

        if (Len(frectdetailidxarr)<1) then Exit Sub

        dbget.beginTrans

		''주문 통보 상태가 있을경우  주문확인 상태로 detail 상태 변경
		sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_detail set" & vbCrlf
		sqlStr = sqlStr + " currstate = '3'" & vbCrlf
		sqlStr = sqlStr + " ,upcheconfirmdate=getdate()" & vbCrlf
		sqlStr = sqlStr + " where detailidx in (" & frectdetailidxarr & ")" & vbCrlf
		sqlStr = sqlStr + " and makerid='" + FRectDesignerID + "'"  & vbCrlf
        sqlStr = sqlStr + " and currstate ='2'"

		'response.write sqlStr &"<br>"
		dbget.Execute sqlStr

        ''배송 통보 상태가 있을경우 배송준비로 master 상태 변경
        sqlStr = ""
        sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_master"
        sqlStr = sqlStr + " set ipkumdiv=6"
        sqlStr = sqlStr + " where masteridx in ("
        sqlStr = sqlStr + "     select d.masteridx"
        sqlStr = sqlStr + "     from db_shop.dbo.tbl_shopbeasong_order_detail d"
        sqlStr = sqlStr + "     where d.detailidx in (" & frectdetailidxarr & ")" & vbCrlf
        sqlStr = sqlStr + "     and d.makerid='" + FRectDesignerID + "'"  & vbCrlf
        sqlStr = sqlStr + "     )"
        sqlStr = sqlStr + " and db_shop.dbo.tbl_shopbeasong_order_master.ipkumdiv='5'"
        sqlStr = sqlStr + " and db_shop.dbo.tbl_shopbeasong_order_master.cancelyn='N'"

		'response.write sqlStr &"<br>"
        dbget.Execute sqlStr

		If Err.Number = 0 Then
		    dbget.CommitTrans
		else
		    dbget.rollbackTrans

			response.write "<script language='javascript'>"
			response.write "	alert('값이 일치 하지 않습니다. 관리자 문의 하세요');"
			response.write "	history.back();"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

        sqlStr = ""
		sqlStr = "select" + vbcrlf
		sqlstr = sqlstr + " m.orderno, m.buyname, m.reqzipcode, m.reqzipaddr, m.reqaddress, m.comment" + vbcrlf
		sqlStr = sqlStr + " ,m.buyphone, m.buyhp, m.buyemail, m.reqname, m.reqphone, m.reqhp, m.regdate" + vbcrlf
		sqlStr = sqlStr + " ,d.itemid,  d.itemno, d.itemoption, d.itemgubun" + vbcrlf
		sqlStr = sqlStr + " ,od.itemname,od.itemoptionname ,od.sellprice, u.shopname" + vbcrlf
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr + " 	on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & "		on m.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr + "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx" +vbcrlf
	    sqlStr = sqlStr + " where d.currstate='3'" + vbcrlf
		sqlStr = sqlStr + " and d.detailidx in (" & frectdetailidxarr & ")" + vbcrlf
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'" + vbcrlf
		sqlStr = sqlStr + " order by m.baljudate, d.detailidx "

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0

		do until rsget.EOF

				set FItemList(i) = new cupchebeasong_item

				FItemList(i).fsellprice = rsget("sellprice")
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).forderno = rsget("orderno")
				FItemList(i).Fbuyname		= db2Html(rsget("buyname"))
				FItemList(i).Freqzipcode	= rsget("reqzipcode")
				FItemList(i).Freqzipaddr	= db2Html(rsget("reqzipaddr"))
				FItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
				FItemList(i).Fcomment		= db2Html(rsget("comment"))
				FItemList(i).Fbuyphone	= rsget("buyphone")
				FItemList(i).Fbuyhp		= rsget("buyhp")
				FItemList(i).Fbuyemail	= rsget("buyemail")
				FItemList(i).Freqname		= db2Html(rsget("reqname"))
				FItemList(i).Freqphone	= rsget("reqphone")
				FItemList(i).Freqhp		= rsget("reqhp")
				FItemList(i).FRegDate     = rsget("regdate")
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).FItemName    = db2Html(rsget("itemname"))
				FItemList(i).Fitemno      = rsget("itemno")
				FItemList(i).Fitemoption  = rsget("itemoption")

				if IsNull(rsget("itemoptionname")) then
				  FItemList(i).FItemoptionName = "-"
				else
				  FItemList(i).FItemoptionName =  db2Html(rsget("itemoptionname"))
				end if

    			FItemList(i).fshopname    = db2html(rsget("shopname"))

				rsget.movenext
				i=i+1

			loop

		rsget.Close
	end sub

	'/common/offshop/beasong/upche_datebaljulist.asp
	public sub fDesignerDateBaljuList()
		dim sqlStr,i ,sqlsearch

		if FRectDesignerID <> "" then
			sqlsearch = sqlsearch & " and d.makerid ='"&FRectDesignerID&"'" +vbcrlf
		end if

		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " m.cancelyn, m.regdate, m.baljudate, m.buyname, m.reqname" + vbcrlf
		sqlStr = sqlStr & " ,m.masteridx" + vbcrlf
		sqlStr = sqlStr & " ,d.itemno, m.orderno, d.itemid  ,d.detailidx" + vbcrlf
		sqlStr = sqlStr & " ,d.itemoption ,d.itemgubun,isNull(d.currstate,0) as currstate" + vbcrlf
		sqlStr = sqlStr & " ,od.itemname,od.itemoptionname ,od.sellprice ,od.realsellprice ,od.shopbuyprice" + vbcrlf
		sqlStr = sqlStr & " ,(CASE when od.suplyprice=0 THEN convert(int,od.sellprice*(100-s.defaultmargin)/100)"
		sqlStr = sqlStr & " 	 ELSE od.suplyprice END) as suplyprice" + vbcrlf
		sqlStr = sqlStr & " , u.shopname" & vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & "		on m.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr & "		on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr & "	left join db_shop.dbo.tbl_shop_designer s"
		sqlStr = sqlStr & "		on m.shopid=s.shopid"
		sqlStr = sqlStr & " 	and d.makerid=s.makerid"
		sqlStr = sqlStr & " where m.cancelyn = 'N'" + vbcrlf
		sqlStr = sqlStr & " and m.ipkumdiv >= '5'" + vbcrlf
		sqlStr = sqlStr & " and d.itemid<>0" + vbcrlf
		sqlStr = sqlStr & " and d.isupchebeasong='Y'" + vbcrlf
		sqlStr = sqlStr & " and d.cancelyn <> 'Y'" + vbcrlf
		sqlStr = sqlStr & " and d.currstate = 2 " & sqlsearch
		sqlStr = sqlStr & " order by IsNULL(m.baljudate,getdate()), d.masteridx desc , d.detailidx desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.recordcount

		redim preserve FItemList(FTotalCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cupchebeasong_item
				
				FItemList(i).fsellprice = rsget("sellprice")
				FItemList(i).frealsellprice = rsget("realsellprice")
				FItemList(i).fshopbuyprice = rsget("shopbuyprice")
				FItemList(i).fsuplyprice = rsget("suplyprice")
				FItemList(i).fmasteridx = rsget("masteridx")
    			FItemList(i).fdetailidx  = rsget("detailidx")
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fitemoption = rsget("itemoption")
    			FItemList(i).forderno = rsget("orderno")
    			FItemList(i).FItemid 	 = rsget("itemid")
    			FItemList(i).FItemname    = db2html(rsget("itemname"))
    			FItemList(i).fitemoptionname     = db2html(rsget("itemoptionname"))
    			FItemList(i).fitemno     = rsget("itemno")
    			FItemList(i).FBuyname    = db2html(rsget("buyname"))
    			FItemList(i).FReqname    = db2html(rsget("reqname"))
    			FItemList(i).FCancelYn	 = rsget("cancelyn")
    			FItemList(i).FRegdate  = rsget("regdate")
    			FItemList(i).FBaljudate  = rsget("baljudate")
    			FItemList(i).FCurrstate  = rsget("currstate")
    			FItemList(i).fshopname    = db2html(rsget("shopname"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//common/offshop/beasong/shopbeasong_list.asp
	public sub fbeagsong_list()
		dim sqlStr,i ,sqlsearch

		if frectorderno <> "" then
			sqlsearch = sqlsearch & " and m.orderno ='"&frectorderno&"'" +vbcrlf
		end if
		if frectipkumdiv <> "" then
			if (frectipkumdiv <> "99") then
				sqlsearch = sqlsearch & " and m.ipkumdiv ='"&frectipkumdiv&"'" +vbcrlf
			else
				sqlsearch = sqlsearch & " and m.ipkumdiv <> 8 " +vbcrlf
			end if
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid ='"&frectshopid&"'" +vbcrlf
		end if
		if frectreqhp <> "" then
			sqlsearch = sqlsearch & " and replace(m.reqhp,'-','') ='"& replace(frectreqhp,"-","") &"'" +vbcrlf
		end if

		'총 갯수 구하기
		sqlStr = "select count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_master m" +vbcrlf
        sqlStr = sqlStr & " where m.cancelyn='N' " & sqlsearch

        'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = " select top " & Cstr(FPageSize * FCurrPage) & "" + vbcrlf
		sqlStr = sqlStr & " m.masteridx, m.orderno, m.shopid, m.ipkumdiv, m.regdate	, m.beadaldiv" + vbcrlf
		sqlStr = sqlStr & " , m.beadaldate, m.cancelyn, m.buyname, m.buyphone, m.buyhp, m.buyemail" + vbcrlf
		sqlStr = sqlStr & " , m.reqname, m.reqzipcode, m.reqzipaddr, m.reqaddress, m.reqphone" + vbcrlf
		sqlStr = sqlStr & " , m.reqhp, m.comment, u.shopname" + vbcrlf	
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u" + vbcrlf
		sqlStr = sqlStr & " 	on m.shopid = u.userid" + vbcrlf
		sqlStr = sqlStr & " where m.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " order by m.masteridx Desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cupchebeasong_item

				FItemList(i).fmasteridx = rsget("masteridx")
				FItemList(i).forderno = rsget("orderno")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fshopname = rsget("shopname")
				FItemList(i).fipkumdiv = rsget("ipkumdiv")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fbeadaldiv = rsget("beadaldiv")
				FItemList(i).fbeadaldate = rsget("beadaldate")
				FItemList(i).fcancelyn = rsget("cancelyn")
				FItemList(i).fbuyname = db2html(rsget("buyname"))
				FItemList(i).fbuyphone = rsget("buyphone")
				FItemList(i).fbuyhp = rsget("buyhp")
				FItemList(i).fbuyemail = db2html(rsget("buyemail"))
				FItemList(i).freqname = db2html(rsget("reqname"))
				FItemList(i).freqzipcode = rsget("reqzipcode")
				FItemList(i).freqzipaddr = db2html(rsget("reqzipaddr"))
				FItemList(i).freqaddress = db2html(rsget("reqaddress"))
				FItemList(i).freqphone = rsget("reqphone")
				FItemList(i).freqhp = rsget("reqhp")
				FItemList(i).fcomment = db2html(rsget("comment"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//common/offshop/beasong/shopbeasong_input.asp
	public sub fshopbeasong_input()
		dim sqlStr,i ,sqlsearch

		if frectorderno="" then exit sub

		if frectmasteridx_beasong <> "" then
			sqlsearch = sqlsearch & " and bm.masteridx ="&frectmasteridx_beasong&"" +vbcrlf
		end if
		if frectorderno <> "" then
			sqlsearch = sqlsearch & " and m.orderno ='"&frectorderno&"'" +vbcrlf
		end if

		sqlStr = "select top 500" + vbcrlf
		sqlStr = sqlStr & " m.orderno ,m.shopid" +vbcrlf
		sqlStr = sqlStr & " , d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname" +vbcrlf
		sqlStr = sqlStr & " ,d.sellprice, d.realsellprice, d.suplyprice, d.makerid, d.itemno" +vbcrlf
		sqlStr = sqlStr & " , bm.regdate as regdate_beasong, bd.masteridx as masteridx_beasong, bd.itemno as itemno_beasong,bd.currstate, bd.isupchebeasong" +vbcrlf
		sqlStr = sqlStr & " , bd.omwdiv, isnull(bd.odlvType,'') as odlvType, bd.detailidx as detailidx_beasong, bd.beasongdate, bd.songjangdiv, bd.songjangno" +vbcrlf
		sqlStr = sqlStr & " ,isnull(cl.realstock,0) as realstock ,isnull(cl.ipgono,0) as ipgono" +vbcrlf
		sqlStr = sqlStr & " , isnull(s.defaultbeasongdiv,0) as defaultbeasongdiv, isnull(s.comm_cd,'') as comm_cd" + vbCrLf
		sqlStr = sqlStr & " ,u.shopname" + vbCrLf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_master m" +vbcrlf
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail d" +vbcrlf
		sqlStr = sqlStr & " 	on m.idx = d.masteridx" +vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'" +vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shopbeasong_order_detail bd" +vbcrlf
		sqlStr = sqlStr & "		on d.idx = bd.orgdetailidx" +vbcrlf
		sqlStr = sqlStr & "		and bd.cancelyn='N'" +vbcrlf
		sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shopbeasong_order_master bm" +vbcrlf
		sqlStr = sqlStr & "		on bd.masteridx = bm.masteridx" +vbcrlf
		sqlStr = sqlStr & "		and bm.cancelyn='N'" +vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u" +vbcrlf
		sqlStr = sqlStr & " 	on m.shopid = u.userid and u.isusing='Y'" +vbcrlf
		sqlStr = sqlStr & " left join db_summary.dbo.tbl_current_logisstock_summary cl" +vbcrlf
		sqlStr = sqlStr & "		on d.itemgubun = cl.itemgubun" +vbcrlf
		sqlStr = sqlStr & "		and d.itemid = cl.itemid" +vbcrlf
		sqlStr = sqlStr & "		and d.itemoption = cl.itemoption" +vbcrlf
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_designer s " + vbCrLf
		sqlStr = sqlStr & " 	on s.shopid = m.shopid " + vbCrLf
		sqlStr = sqlStr & " 	and s.makerid = d.makerid " + vbCrLf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by d.idx asc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.recordcount

		redim preserve FItemList(FTotalCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cupchebeasong_item

					FItemList(i).fdetailidx_beasong  	= rsget("detailidx_beasong")
					FItemList(i).frealstock  	= rsget("realstock")
					FItemList(i).fipgono  	= rsget("ipgono")
					FItemList(i).fmasteridx_beasong  	= rsget("masteridx_beasong")
					FItemList(i).fshopname  	= rsget("shopname")
					FItemList(i).fregdate_beasong  	= rsget("regdate_beasong")
					FItemList(i).fshopid  	= rsget("shopid")
					FItemList(i).forderno  	= rsget("orderno")
					FItemList(i).fitemgubun  	= rsget("itemgubun")
					FItemList(i).fitemid  	= rsget("itemid")
					FItemList(i).fitemoption  	= rsget("itemoption")
					FItemList(i).fitemname  	= db2html(rsget("itemname"))
					FItemList(i).fitemoptionname  	= db2html(rsget("itemoptionname"))
					FItemList(i).fsellprice  	= rsget("sellprice")
					FItemList(i).frealsellprice  	= rsget("realsellprice")
					FItemList(i).fsuplyprice  	= rsget("suplyprice")
					FItemList(i).fitemno_beasong  	= rsget("itemno_beasong")
					FItemList(i).fitemno  	= rsget("itemno")
					FItemList(i).fmakerid  	= rsget("makerid")
					FItemList(i).fcurrstate  	= rsget("currstate")
					FItemList(i).fisupchebeasong  	= rsget("isupchebeasong")
					FItemList(i).fomwdiv  	= rsget("omwdiv")
					FItemList(i).fodlvType  	= rsget("odlvType")

					FItemList(i).fbeasongdate  	= rsget("beasongdate")

					FItemList(i).fdefaultbeasongdiv  	= rsget("defaultbeasongdiv")

					FItemList(i).Fsongjangno  	= rsget("songjangno")
					FItemList(i).Fsongjangdiv  	= rsget("songjangdiv")
					FItemList(i).fcomm_cd  	= rsget("comm_cd")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//common/offshop/beasong/shopbeasong_cancel.asp
	public sub fshopbeasong_cancel()
		dim sqlStr,i ,sqlsearch

		if frectorderno <> "" then
			sqlsearch = sqlsearch & " and bm.orderno ='"&frectorderno&"'" +vbcrlf
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and bm.shopid ='"&frectshopid&"'" +vbcrlf
		end if
		if FRectStartDay<>"" and FRectEndDay<>"" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " and bm.regdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " and bm.regdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " bm.orderno ,bm.shopid, bm.orderno , bd.itemgubun, bd.itemid, bd.itemoption, ii.shopitemname as itemname" & vbcrlf
		sqlStr = sqlStr & " , ii.shopitemoptionname as itemoptionname, bd.makerid, bd.itemno , bm.regdate as regdate_beasong, bd.masteridx as masteridx_beasong" & vbcrlf
		sqlStr = sqlStr & " , bd.itemno as itemno_beasong,bd.currstate, bd.isupchebeasong , bd.omwdiv, isnull(bd.odlvType,'') as odlvType" & vbcrlf
		sqlStr = sqlStr & " , bd.detailidx as detailidx_beasong, bd.beasongdate, bd.songjangdiv, bd.songjangno ,u.shopname" & vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_master bm" & vbcrlf
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopbeasong_order_detail bd" & vbcrlf
		sqlStr = sqlStr & " 	on bm.masteridx = bd.masteridx" & vbcrlf
		sqlStr = sqlStr & " 	and bm.cancelyn='N' and bd.cancelyn='N'" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_item ii" & vbcrlf
		sqlStr = sqlStr & " 	on bd.itemgubun = ii.itemgubun" & vbcrlf
		sqlStr = sqlStr & " 	and bd.itemid = ii.shopitemid" & vbcrlf
		sqlStr = sqlStr & " 	and bd.itemoption = ii.itemoption" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shopjumun_master m" & vbcrlf
		sqlStr = sqlStr & " 	on bm.orderno = m.reforderno" & vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N'" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shopjumun_detail d" & vbcrlf
		sqlStr = sqlStr & " 	on m.idx = d. masteridx" & vbcrlf
		sqlStr = sqlStr & " 	and d.cancelyn='N'" & vbcrlf
		sqlStr = sqlStr & " 	and bd.itemgubun = d.itemgubun" & vbcrlf
		sqlStr = sqlStr & " 	and bd.itemid = d.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and bd.itemoption = d.itemoption" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & " 	on bm.shopid = u.userid" & vbcrlf
		sqlStr = sqlStr & " where (d.idx is not null or m.idx is not null) " & sqlsearch
		sqlStr = sqlStr & " order by bm.masteridx desc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.recordcount

		redim preserve FItemList(FTotalCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cupchebeasong_item

					FItemList(i).fdetailidx_beasong  	= rsget("detailidx_beasong")
					FItemList(i).fmasteridx_beasong  	= rsget("masteridx_beasong")
					FItemList(i).fshopname  	= rsget("shopname")
					FItemList(i).fregdate_beasong  	= rsget("regdate_beasong")
					FItemList(i).fshopid  	= rsget("shopid")
					FItemList(i).forderno  	= rsget("orderno")
					FItemList(i).fitemgubun  	= rsget("itemgubun")
					FItemList(i).fitemid  	= rsget("itemid")
					FItemList(i).fitemoption  	= rsget("itemoption")
					FItemList(i).fitemname  	= db2html(rsget("itemname"))
					FItemList(i).fitemoptionname  	= db2html(rsget("itemoptionname"))
					FItemList(i).fitemno_beasong  	= rsget("itemno_beasong")
					FItemList(i).fitemno  	= rsget("itemno")
					FItemList(i).fmakerid  	= rsget("makerid")
					FItemList(i).fcurrstate  	= rsget("currstate")
					FItemList(i).fisupchebeasong  	= rsget("isupchebeasong")
					FItemList(i).fomwdiv  	= rsget("omwdiv")
					FItemList(i).fodlvType  	= rsget("odlvType")
					FItemList(i).fbeasongdate  	= rsget("beasongdate")
					FItemList(i).Fsongjangno  	= rsget("songjangno")
					FItemList(i).Fsongjangdiv  	= rsget("songjangdiv")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//common/offshop/beasong/shopjumun_address.asp
    public Sub fshopjumun_edit()
        dim sqlStr ,sqlsearch

		if frectmasteridx <> "" then
			sqlsearch = sqlsearch & " and m.masteridx="&frectmasteridx&"" +vbcrlf
		end if
		if frectorderno <> "" then
			sqlsearch = sqlsearch & " and m.orderno='"&frectorderno&"'" +vbcrlf
		end if

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " m.masteridx,m.orderno,m.shopid,m.ipkumdiv,m.regdate,m.beadaldiv,m.beadaldate" & vbcrlf
		sqlStr = sqlStr & " ,m.cancelyn,m.buyname,m.buyphone,m.buyhp,m.buyemail,m.reqname,m.reqzipcode" & vbcrlf
		sqlStr = sqlStr & " ,m.reqzipaddr,m.reqaddress,m.reqphone,m.reqhp,m.comment" & vbcrlf
		sqlStr = sqlStr & " ,u.shopname" & vbcrlf
		sqlStr = sqlStr & " , sc.AuthIdx, sc.UserHp, sc.SmsYN, sc.KakaoTalkYN" & vbcrlf
		sqlStr = sqlStr & " , sc.IsUsing, sc.Regdate, sc.LastUpdate, sc.certno" & vbcrlf
		sqlStr = sqlStr & " , (select count(*) from db_shop.[dbo].[tbl_shopjumun_sms_cert] where m.orderno=orderno and isusing='Y') as BeaSongcnt" & vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_master m" +vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shopjumun_sms_cert sc" & vbcrlf
	    sqlStr = sqlStr & " 	on m.orderno=sc.orderno" & vbcrlf
	    sqlStr = sqlStr & " 	and sc.isusing='Y'" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u" +vbcrlf
		sqlStr = sqlStr & " 	on m.shopid = u.userid and u.isusing='Y'" +vbcrlf
        sqlStr = sqlStr & " where cancelyn='N' " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount

        set FOneItem = new cupchebeasong_item

        if Not rsget.Eof then

			FOneItem.fshopname = rsget("shopname")
			FOneItem.fmasteridx = rsget("masteridx")
			FOneItem.forderno = rsget("orderno")
			FOneItem.fshopid = rsget("shopid")
			FOneItem.fipkumdiv = rsget("ipkumdiv")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fbeadaldiv = rsget("beadaldiv")
			FOneItem.fbeadaldate = rsget("beadaldate")
			FOneItem.fcancelyn = rsget("cancelyn")
			FOneItem.fbuyname = db2html(rsget("buyname"))
			FOneItem.fbuyphone = rsget("buyphone")
			FOneItem.fbuyhp = rsget("buyhp")
			FOneItem.fbuyemail = db2html(rsget("buyemail"))
			FOneItem.freqname = db2html(rsget("reqname"))
			FOneItem.freqzipcode = rsget("reqzipcode")
			FOneItem.freqzipaddr = db2html(rsget("reqzipaddr"))
			FOneItem.freqaddress = db2html(rsget("reqaddress"))
			FOneItem.freqphone = rsget("reqphone")
			FOneItem.freqhp = rsget("reqhp")
			FOneItem.fcomment = db2html(rsget("comment"))
			FOneItem.fAuthIdx = rsget("AuthIdx")
			FOneItem.fBeaSongcnt = rsget("BeaSongcnt")
			FOneItem.fUserHp = rsget("UserHp")
			FOneItem.fSmsYN = rsget("SmsYN")
			FOneItem.fKakaoTalkYN = rsget("KakaoTalkYN")
			FOneItem.fIsUsing = rsget("IsUsing")
			FOneItem.fRegdate = rsget("Regdate")
			FOneItem.fLastUpdate = rsget("LastUpdate")
			FOneItem.fCertNo = rsget("certno")

        end if
        rsget.Close
    end Sub

	'/common/offshop/upche/shopjumun_list.asp
	public sub fshopjumun_list()
		dim sqlStr,i ,sqlsearch

		if frectorderno <> "" then
			sqlsearch = sqlsearch & " and m.orderno ='" & frectorderno & "'" + vbCrLf
		else
			'최근한달.
			sqlsearch = sqlsearch & " and IXyyyymmdd >= convert(varchar(10), dateadd(m, -1, getdate()), 21) " + vbCrLf
		end if

		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid ='" & frectshopid & "'" + vbCrLf
		end if

		'XXXXXXXX 주문통보 상태 이전이거나 결제완료인 상태만 노출
		'다 가져온다. 재차 배송입력도 해야한다.
		sqlStr = " select top 200 " + vbCrLf
		sqlStr = sqlStr & " m.orderno , m.shopid , m.IXyyyymmdd , d.masteridx " + vbCrLf
		sqlStr = sqlStr & " , d.itemgubun , d.itemid , d.itemoption , d.itemname " + vbCrLf
		sqlStr = sqlStr & " , d.itemoptionname , d.sellprice , d.realsellprice " + vbCrLf
		sqlStr = sqlStr & " , d.suplyprice , d.itemno , d.makerid , td.currstate " + vbCrLf
		sqlStr = sqlStr & " , td.isupchebeasong , td.omwdiv , td.odlvType , td.detailidx " + vbCrLf
		sqlStr = sqlStr & " , u.shopname , td.masteridx , isnull(cl.realstock,0) as realstock " + vbCrLf
		sqlStr = sqlStr & " , isnull(cl.ipgono,0) as ipgono " + vbCrLf
		sqlStr = sqlStr & " , isnull(s.defaultbeasongdiv,0) as defaultbeasongdiv, isnull(s.comm_cd,'') as comm_cd" + vbCrLf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shopjumun_master m" + vbCrLf
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shopjumun_detail d " + vbCrLf
		sqlStr = sqlStr & " 	on m.idx = d.masteridx " + vbCrLf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shopbeasong_order_detail td " + vbCrLf
		sqlStr = sqlStr & " 	on d.idx = td.orgdetailidx " + vbCrLf
		sqlStr = sqlStr & " 	and td.cancelyn='N' " + vbCrLf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u " + vbCrLf
		sqlStr = sqlStr & " 	on m.shopid = u.userid " + vbCrLf
		sqlStr = sqlStr & " 	and u.isusing='Y' " + vbCrLf
		sqlStr = sqlStr & " left join db_summary.dbo.tbl_current_logisstock_summary cl " + vbCrLf
		sqlStr = sqlStr & " 	on d.itemgubun = cl.itemgubun " + vbCrLf
		sqlStr = sqlStr & " 	and d.itemid = cl.itemid " + vbCrLf
		sqlStr = sqlStr & " 	and d.itemoption = cl.itemoption " + vbCrLf
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_designer s " + vbCrLf
		sqlStr = sqlStr & " 	on s.shopid = m.shopid " + vbCrLf
		sqlStr = sqlStr & " 	and s.makerid = d.makerid " + vbCrLf
		sqlStr = sqlStr & " where m.cancelyn='N' " + vbCrLf
		sqlStr = sqlStr & " and d.cancelyn='N' " + vbCrLf

		sqlStr = sqlStr & sqlsearch

		sqlStr = sqlStr & " order by d.idx asc" + vbCrLf

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.recordcount

		redim preserve FItemList(FTotalCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cupchebeasong_item

					FItemList(i).fdetailidx  	= rsget("detailidx")
					FItemList(i).frealstock  	= rsget("realstock")
					FItemList(i).fipgono  	= rsget("ipgono")
					FItemList(i).fmasteridx  	= rsget("masteridx")
					FItemList(i).fshopname  	= rsget("shopname")
					FItemList(i).fIXyyyymmdd  	= rsget("IXyyyymmdd")
					FItemList(i).fshopid  	= rsget("shopid")
					FItemList(i).forderno  	= rsget("orderno")
					FItemList(i).fitemgubun  	= rsget("itemgubun")
					FItemList(i).fitemid  	= rsget("itemid")
					FItemList(i).fitemoption  	= rsget("itemoption")
					FItemList(i).fitemname  	= db2html(rsget("itemname"))
					FItemList(i).fitemoptionname  	= db2html(rsget("itemoptionname"))
					FItemList(i).fsellprice  	= rsget("sellprice")
					FItemList(i).frealsellprice  	= rsget("realsellprice")
					FItemList(i).fsuplyprice  	= rsget("suplyprice")
					FItemList(i).fitemno  	= rsget("itemno")
					FItemList(i).fmakerid  	= rsget("makerid")
					FItemList(i).fcurrstate  	= rsget("currstate")
					FItemList(i).fisupchebeasong  	= rsget("isupchebeasong")
					FItemList(i).fomwdiv  	= rsget("omwdiv")
					FItemList(i).fodlvType  	= rsget("odlvType")
					FItemList(i).fdefaultbeasongdiv  	= rsget("defaultbeasongdiv")
					FItemList(i).fcomm_cd  	= rsget("comm_cd")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/common/offshop/beasong/popupchejumunsms_off.asp
	public sub fbeasongsmslist()
		dim sqlStr,i ,sqlsearch

		if frectmakerid <> "" then
			sqlsearch = sqlsearch & " and d.makerid='"&frectmakerid&"'"
		end if		

		if frectorderno <> "" then
			sqlsearch = sqlsearch & " and d.orderno='"&frectorderno&"'"
		end if
		
		if frectmasteridx <> "" then
			sqlsearch = sqlsearch & " and m.masteridx="&frectmasteridx&""
		end if
		
		if frectdetailidx <> "" then
			sqlsearch = sqlsearch & " and d.detailidx="&frectdetailidx&""
		end if
		
		if FRectIsUpcheBeasong <> "" then
			sqlsearch = sqlsearch & " and d.isupchebeasong='"&FRectIsUpcheBeasong&"'"
		end if		
		
		sqlStr = "select"
		sqlStr = sqlStr & " m.regdate, m.baljudate, m.buyname, m.reqname,m.cancelyn ,m.masteridx"
		sqlStr = sqlStr & " ,m.shopid, m.orderno,d.detailidx, d.itemno, d.itemid, d.beasongdate,d.itemgubun" +vbcrlf
		sqlStr = sqlStr & " ,d.upcheconfirmdate ,isNull(d.currstate,0) as currstate ,d.Itemoption ,d.makerid" +vbcrlf
		sqlStr = sqlStr & " ,d.songjangno, d.songjangdiv,d.cancelyn as detailcancelyn, d.upchesendsms" +vbcrlf
		sqlStr = sqlStr & " ,od.itemname ,od.itemoptionname, od.sellprice ,od.realsellprice ,od.shopbuyprice"
		sqlStr = sqlStr & " ,u.shopphone"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master m" + vbcrlf
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
		sqlStr = sqlStr + " 	on m.masteridx=d.masteridx" + vbcrlf
		sqlStr = sqlStr + "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
		sqlStr = sqlStr + "		on d.orgdetailidx = od.idx" +vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u" + vbcrlf
		sqlStr = sqlStr & " 	on m.shopid = u.userid" + vbcrlf
		sqlStr = sqlStr & " where m.ipkumdiv <> '8'" +vbcrlf
		sqlStr = sqlStr & " and m.cancelyn = 'N'" +vbcrlf
		sqlStr = sqlStr & " and d.currstate < 3"	'/업체확인이전상태
		sqlStr = sqlStr & " and d.itemid<>0"
		sqlStr = sqlStr & " and d.cancelyn <> 'Y' " & sqlsearch
		sqlStr = sqlStr & " order by d.itemid desc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cupchebeasong_item
				
				FItemList(i).fmakerid = rsget("makerid")
				FItemList(i).fupchesendsms = rsget("upchesendsms")
				FItemList(i).fshopphone = rsget("shopphone")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fsellprice = rsget("sellprice")
				FItemList(i).frealsellprice = rsget("realsellprice")
				FItemList(i).fshopbuyprice = rsget("shopbuyprice")
				FItemList(i).fmasteridx = rsget("masteridx")
				FItemList(i).fitemgubun = rsget("itemgubun")
    			FItemList(i).fdetailidx				  = rsget("detailidx")
    			FItemList(i).forderno		  = rsget("orderno")
    			FItemList(i).FItemid 			  = rsget("itemid")
    			FItemList(i).FItemoption 			  = rsget("Itemoption")
    			FItemList(i).FItemname 		  = db2html(rsget("itemname"))
    			FItemList(i).fitemoptionname     	  = db2html(rsget("itemoptionname"))
    			FItemList(i).fitemno           = rsget("itemno")
    			FItemList(i).FBuyname           = db2html(rsget("buyname"))
    			FItemList(i).FReqname			  = db2html(rsget("reqname"))
    			FItemList(i).FCancelYn		  = rsget("cancelyn")
    			FItemList(i).FRegdate			  = rsget("regdate")
    			FItemList(i).FBaljudate		  = rsget("baljudate")
    			FItemList(i).Fupcheconfirmdate  = rsget("upcheconfirmdate")
    			FItemList(i).FCurrstate		  = rsget("currstate")
    			FItemList(i).Fbeasongdate       = rsget("beasongdate")
    			FItemList(i).FSongjangno        = rsget("songjangno")
    			FItemList(i).FSongjangdiv       = rsget("songjangdiv")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 50
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

'//배송구분(odlvType) 텐바이텐 배송이냐 업체 배송이냐..  차후 무료배송 작업 요망
function Drawbeasonggubun(selectBoxName,selectedId,changeFlag)
dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <%if isnull(selectedId) or selectedId="" then response.write " selected"%> >현장수령</option>
		<option value='0' <% if selectedId="0" then response.write " selected"%> >매장배송</option>
		<option value='1' <% if selectedId="1" then response.write " selected"%> >물류배송</option>
		<option value='2' <% if selectedId="2" then response.write " selected"%> >업체배송</option>

		<!--<option value='4' <%' if selectedId="4" then response.write " selected"%> >텐바이텐무료배송</option>-->
		<!--<option value='5' <%' if selectedId="5" then response.write " selected"%> >업체배송</option>-->
		<!--<option value='7' <%' if selectedId="7" then response.write " selected"%> >업체착불배송</option>-->
	</select>
<%
end function

function Drawupchebeasonggubun(selectBoxName,selectedId,changeFlag)
dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <%if selectedId="" then response.write " selected"%> >선택</option>
		<option value='N' <% if selectedId="N" then response.write " selected"%> >매장배송</option>
		<option value='Y' <% if selectedId="Y" then response.write " selected"%> >업체배송</option>
	</select>
<%
end function

'//출고 상태
function drawshopIpkumDivName(selectBoxName,selectedId,changeFlag)
dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <% if selectedId="" then response.write " selected"%> >선택</option>
		<option value='1' <% if selectedId="1" then response.write " selected"%> >배송지입력전</option>
		<option value='2' <% if selectedId="2" then response.write " selected"%> >배송지입력완료</option>
		<option value='5' <% if selectedId="5" then response.write " selected"%> >업체통보</option>
		<option value='6' <% if selectedId="6" then response.write " selected"%> >업체확인</option>
		<option value='7' <% if selectedId="7" then response.write " selected"%> >일부출고</option>
		<option value='8' <% if selectedId="8" then response.write " selected"%> >출고완료</option>
		<option value='99' <% if selectedId="99" then response.write " selected"%> >미배송건만</option>
	</select>
<%
end Function

'// 인증여부 선택
Function drawcertsendgubun(selectBoxName,selectedId,chplg,dispNotValue)
%>
	<select name="<%=selectBoxName%>" <%= chplg %>>
		<% if dispNotValue="Y" then %>
			<option value="" <% if selectedId="" then response.write "selected" %>>SELECT</option>
		<% end if %>
		<option value="KAKAOTALK" <% if selectedId="KAKAOTALK" then response.write "selected" %>>카카오톡 발송</option>
		<option value="SMS" <% if selectedId="SMS" then response.write "selected" %>>SMS 발송</option>
	</select>
<%
end function
%>
