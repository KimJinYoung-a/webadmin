<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
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
	public frealsum
	public fcashsum
	public fcardsum
	public fpointuserno
	public fonlineuserid
	public Forgdetailidx
	public frequiremaejang
	public fshopid
	
    ''송장 필드가 필요한 정보
    public function IsRequireSongjangNO()
        IsRequireSongjangNO = false

        IsRequireSongjangNO = (Fdivcd="A030") or (Fdivcd="A031")
    end function
    
	''취소 프로세스
	public function fnIsCancelProcess_off(idivcd)
	    fnIsCancelProcess_off = (idivcd = "A008")
	end function

    ''취소 프로세스
    public function IsCancelProcess_off()
        IsCancelProcess_off = fnIsCancelProcess_off(Fdivcd)
    end function

	''반품 프로세스(회수, 맞교환 회수)
	public function fnIsReturnProcess_off(idivcd)
	    fnIsReturnProcess_off = (idivcd = "A030") or (idivcd = "A031")
	end function

    ''반품 프로세스
    public function IsReturnProcess_off()
        IsReturnProcess_off = fnIsReturnProcess_off(Fdivcd)
    end function
    
    public function IsAsRegAvail_off(byval iCancelYn, byref descMsg)   
    
    'response.write Fdivcd & "<Br>!!!!!"    
        IsAsRegAvail_off = false

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

            if (iCancelYn<>"N") then
                IsAsRegAvail_off = false
                descMsg      = "취소된 거래입니다. - 반품 접수 불가능 "
                exit function
            end if

            IsAsRegAvail_off = true

        ''a/s
        elseif (Fdivcd = "A030") then 
			IsAsRegAvail_off = true
			
        else
            descMsg = "정의 되지 않았습니다." + Fdivcd
        end if
    end function
    
	'/취소건 상태값
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
	
	'/취소건 색 처리
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

    public function GetDefaultRegNo_off(IsRegState)
        if (IsRegState) then
            GetDefaultRegNo_off = Fitemno
        else
            GetDefaultRegNo_off = Fregitemno
        end if
    end function

	'//정상건이냐 아니냐 체크
	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

    public function shopGetAsDivCDColor()
        shopGetAsDivCDColor = FdivcdName
    end function
    
    public function shopGetAsDivCDName()
        shopGetAsDivCDName = FdivcdName
    end function

    public function shopGetCurrstateName()
        shopGetCurrstateName = FcurrstateName
    end function

     public function shopGetCurrstateColor()
        shopGetCurrstateColor = FcurrstateColor
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
    public FRectDeliveryNo
	public FRectOnlyJupsu
	public FrectCardNo
	public FrectUserID
	public frectdatefg
	public frectmail
	public FRectcancelyn
	
	'/admin/offshop/shopcscenter/action/pop_cs_action_new.asp		'//admin/offshop/shopcscenter/action/cs_action_detail.asp
	'//common/offshop/shopcscenter/shop_csdetail.asp		'//common/offshop/shopcscenter/upche_csdetail.asp
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
		sqlStr = sqlStr + " ,a.requireupche,a.makerid"
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C4.comm_name as currstatename"
        sqlStr = sqlStr & " ,sd.asid , sd.reqname ,sd.reqphone ,sd.reqhp ,sd.reqzipcode ,sd.reqzipaddr ,sd.reqemail"
        sqlStr = sqlStr & " ,sd.reqetcaddr ,sd.reqetcstr ,sd.songjangdiv ,sd.songjangno ,sd.regdate ,sd.senddate"        
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_cs_master A "
        sqlStr = sqlStr + " left join [db_shop].dbo.tbl_shopjumun_cs_delivery sd"
        sqlStr = sqlStr + " 	on a.masteridx = sd.asid"
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
			
			FOneItem.freqemail           = db2html(rsget("reqemail"))
            FOneItem.Fasid              = rsget("asid")
            FOneItem.Freqname           = db2html(rsget("reqname"))
            FOneItem.Freqphone          = rsget("reqphone")
            FOneItem.Freqhp             = rsget("reqhp")
            FOneItem.Freqzipcode        = rsget("reqzipcode")
            FOneItem.Freqzipaddr        = rsget("reqzipaddr")
            FOneItem.FReqAddress        = db2html(rsget("reqetcaddr"))
            FOneItem.FComment          = db2html(rsget("reqetcstr"))
            FOneItem.Fregdate           = rsget("regdate")
            FOneItem.Fsenddate          = rsget("senddate")
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

	'//admin/offshop/shopcscenter/action/cs_action_detail.asp
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
    
	'//admin/offshop/shopcscenter/action/cs_action_detail.asp	'//admin/offshop/shopcscenter/action/pop_CsDeliveryEdit.asp
    public Sub fGetOneCsDeliveryItem()
        dim i,sqlStr ,sqlsearch

        if FRectCsAsID <> "" then
        	sqlsearch = sqlsearch + " and asid = "&FRectCsAsID&""
        end if
        
        sqlStr = " select top 1"
        sqlStr = sqlStr & " a.asid , a.reqname ,a.reqphone ,a.reqhp ,a.reqzipcode ,a.reqzipaddr ,a.reqemail"
        sqlStr = sqlStr & " ,a.reqetcaddr ,a.reqetcstr,a.songjangdiv ,a.songjangno ,a.regdate ,a.senddate"
        sqlStr = sqlStr + " from [db_shop].dbo.tbl_shopjumun_cs_delivery A"
        sqlStr = sqlStr + " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1

        FResultCount = rsget.RecordCount
        ftotalcount = rsget.RecordCount

        if not rsget.EOF  then
            set FOneItem = new COrderItem

			FOneItem.freqemail           = db2html(rsget("reqemail"))            
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
    
	'//admin/offshop/shopcscenter/action/inc_cs_action_item_list.asp		'//admin/offshop/shopcscenter/action/cs_action_detail.asp
	'//common/offshop/shopcscenter/shop_csdetail.asp		'//common/offshop/shopcscenter/upche_csdetail.asp
    public Sub fGetCsDetailList()
        dim SqlStr, i , sqlsearch
        
        if FRectCsAsID <> "" then
        	sqlsearch = sqlsearch + " and c.masteridx = "&FRectCsAsID&""
        end if
		
		sqlStr = "select"
		sqlStr = sqlStr + " c.detailidx ,c.masteridx ,c.orgdetailidx ,c.orderno"
		sqlStr = sqlStr + " ,c.itemid ,c.itemoption ,c.itemgubun ,c.makerid ,c.regitemno ,c.confirmitemno"
		sqlStr = sqlStr + " ,c.orderitemno ,c.isupchebeasong ,c.regdetailstate ,c.currstate"
		sqlStr = sqlStr + " ,d.itemname , d.itemoptionname , d.sellprice"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_cs_detail c"
	    sqlStr = sqlStr + " left join db_shop.dbo.tbl_shopjumun_detail d"
	    sqlStr = sqlStr + " 	on c.orgdetailidx=d.idx"
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

            FItemList(i).fdetailidx              = rsget("detailidx")
            FItemList(i).fmasteridx        = rsget("masteridx")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget("regdetailstate")   ''접수 당시 진행 상태
            FItemList(i).forderno     = rsget("orderno")
            FItemList(i).Fitemid          = rsget("itemid")
            FItemList(i).Fitemoption      = rsget("itemoption")
            FItemList(i).fitemgubun      = rsget("itemgubun")
            FItemList(i).Fmakerid         = rsget("makerid")
            FItemList(i).Fitemname        = db2html(rsget("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))            
            FItemList(i).Fitemno          = rsget("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")
			FItemList(i).fsellprice  = rsget("sellprice")

			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub
    
	'//admin/offshop/shopcscenter/action/inc_cs_action_item_list.asp
    public Sub fGetOrderDetailByCsDetail()
        dim SqlStr, i , sqlsearch
		
		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch + " and d.masteridx='" + CStr(FRectmasteridx) + "'"
		end if
		
		sqlStr = "select"
		sqlStr = sqlStr + " d.itemname, d.itemoptionname ,d.sellprice ,d.itemno,d.cancelyn ,d.itemid ,d.itemoption ,d.itemgubun"		
		sqlStr = sqlStr + " ,d.makerid ,d.idx as orgdetailidx ,c.detailidx ,c.masteridx ,c.orderno"
		sqlStr = sqlStr + " ,IsNULL(c.regitemno,0) as regitemno ,IsNULL(c.confirmitemno,0) as confirmitemno"
		sqlStr = sqlStr + " ,c.orderitemno ,c.isupchebeasong ,c.regdetailstate ,c.currstate"
		sqlStr = sqlStr + " from db_shop.[dbo].tbl_shopjumun_detail d" +vbcrlf
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shopjumun_cs_detail c "
		sqlStr = sqlStr + " 	on c.masteridx='" + CStr(FRectCsAsID) + "'"
		sqlStr = sqlStr + " 	and d.idx = c.orgdetailidx "    
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
        sqlStr = sqlStr + " order by d.makerid, d.itemid, d.itemoption"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new COrderItem

            FItemList(i).fdetailidx       = rsget("detailidx")
            FItemList(i).fmasteridx       = rsget("masteridx")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")
            FItemList(i).Fregdetailstate  = rsget("regdetailstate")
            FItemList(i).Forgdetailidx  = rsget("orgdetailidx")
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
            
			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub
    
	'//admin/offshop/shopcscenter/action/inc_cs_action_prev_cslist.asp		'//admin/offshop/shopcscenter/action/cs_action_list.asp
	'//common/offshop/shopcenter/shop_cslist.asp	'//common/offshop/shopcscenter/upche_cslist.asp
    public Sub fGetCSASMasterList()
        dim i,sqlStr, sqlsearch

		if (FRectSearchType="") then
	        if (frectmasteridx <> "") then
				sqlsearch = sqlsearch + " and A.orgmasteridx='" + CStr(frectmasteridx) + "' "
	        end if
		
		''업체가 쿼리시
		elseif (FRectSearchType="upcheview") then		    
            'sqlsearch = sqlsearch + " and a.divcd not in ('')"		'/제외구분
            sqlsearch = sqlsearch + " and a.deleteyn='N'"
            sqlsearch = sqlsearch + " and (a.requireupche='Y' or A.divcd='A031')"	'//업체일경우 업체a/s(매장회수) 내역도 보여줌
            sqlsearch = sqlsearch + " and a.makerid='" + CStr(FRectMakerid) + "' "

	        if (frectmasteridx <> "") then
				sqlsearch = sqlsearch + " and A.orgmasteridx='" + CStr(frectmasteridx) + "' "
	        end if

	        if (frectorderno <> "") then
				sqlsearch = sqlsearch + " and om.orderno='" + CStr(frectorderno) + "'"
	        end if
	        
            if (FRectOnlyJupsu="on") then
                sqlsearch = sqlsearch + " and currstate='B001'"		'/접수상태
            end if

            if (FRectCurrstate = "notfinish") then
				sqlsearch = sqlsearch + " and A.currstate <> 'B007' "
            elseif (FRectCurrstate = "notfinal") then
				sqlsearch = sqlsearch + " and (A.currstate = 'B006' or A.currstate = 'B008')"				
	        elseif (FRectCurrstate <> "") then
				sqlsearch = sqlsearch + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if

            if (FRectUserName <> "") then
				sqlsearch = sqlsearch + " and A.customername='" + CStr(FRectUserName) + "' "
	        end if

			if frectshopid <> "" then
				sqlsearch = sqlsearch & " and om.shopid ='"&frectshopid&"'" +vbcrlf
			end if
			
		''매장 쿼리시
		elseif (FRectSearchType="shopview") then		    
            'sqlsearch = sqlsearch + " and a.divcd not in ('')"		'/제외구분
            sqlsearch = sqlsearch + " and a.deleteyn='N'"
            sqlsearch = sqlsearch + " and (a.requiremaejang='Y' or A.divcd='A030')"		'//매장일경우 업체a/s 내역도 보여줌

	        if (frectmasteridx <> "") then
				sqlsearch = sqlsearch + " and A.orgmasteridx='" + CStr(frectmasteridx) + "' "
	        end if

	        if (frectorderno <> "") then
				sqlsearch = sqlsearch + " and om.orderno='" + CStr(frectorderno) + "'"
	        end if
	        
            if (FRectOnlyJupsu="on") then
                sqlsearch = sqlsearch + " and currstate='B001'"		'/접수상태
            end if

            if (FRectCurrstate = "notfinish") then
				sqlsearch = sqlsearch + " and A.currstate <> 'B007' "
            elseif (FRectCurrstate = "notfinal") then
				sqlsearch = sqlsearch + " and (A.currstate = 'B006' or A.currstate = 'B008')"					
	        elseif (FRectCurrstate <> "") then
				sqlsearch = sqlsearch + " and A.currstate ='" + CStr(FRectCurrstate) + "' "
	        end if

            if (FRectUserName <> "") then
				sqlsearch = sqlsearch + " and A.customername='" + CStr(FRectUserName) + "' "
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

	        if (frectorderno <> "") then
				sqlsearch = sqlsearch + " and om.orderno='" + CStr(frectorderno) + "'"
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
				sqlsearch = sqlsearch + " and A.currstate <> 'B007' "
            elseif (FRectCurrstate = "notfinal") then
				sqlsearch = sqlsearch + " and (A.currstate = 'B006' or A.currstate = 'B008')"					
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

			if frectshopid <> "" then
				sqlsearch = sqlsearch & " and om.shopid ='"&frectshopid&"'" +vbcrlf
			end if
			
        end If
        
        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_cs_master A"
        sqlStr = sqlStr + " join db_shop.dbo.tbl_shopjumun_master om"
        sqlStr = sqlStr + " 	on a.orgmasteridx = om.idx"
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
        sqlStr = sqlStr + " ,A.requireupche, A.makerid, A.currstate"
        sqlStr = sqlStr + " ,a.contents_jupsu ,a.contents_finish ,a.requiremaejang"
        sqlStr = sqlStr + " ,C1.comm_name as divcdname, C4.comm_name as currstatename"
        sqlStr = sqlStr + " ,C4.comm_color as currstatecolor , u.shopname"
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_cs_master A"
        sqlStr = sqlStr + " join db_shop.dbo.tbl_shopjumun_master om"
        sqlStr = sqlStr + " 	on a.orgmasteridx = om.idx"
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
				
				FItemList(i).frequiremaejang = rsget("requiremaejang")
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
                FItemList(i).Frequireupche      = rsget("requireupche")
                FItemList(i).Fmakerid           = rsget("makerid")

                rsget.MoveNext
                i = i + 1
            loop
        end if
        rsget.close
    end sub
    
	'//admin/offshop/shopcscenter/order/orderitemmaster.asp		'//admin/offshop/shopcscenter/order/pop_order_receipt.asp
	public Sub fQuickSearchOrderDetail()
		dim sqlStr, i , sqlsearch
		
		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch & " and d.masteridx='" + CStr(FRectmasteridx) + "'"
		end if

		if FRectorderno <> "" then
			sqlsearch = sqlsearch & " and d.orderno='" + FRectorderno + "'"
		end if			

		if FRectcancelyn <> "" then
			sqlsearch = sqlsearch & " and d.cancelyn='" + FRectcancelyn + "'"
		end if

		sqlStr = "select"
		sqlStr = sqlStr + " d.idx ,d.orderno , d.masteridx ,d.makerid ,d.itemgubun, d.itemid , d.itemoption ,d.itemno"
		sqlStr = sqlStr + " ,d.cancelyn ,d.sellprice, d.realsellprice,d.itemname ,d.itemoptionname"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
        sqlStr = sqlStr + " order by d.makerid, d.itemid, d.itemoption"
        
        'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new COrderItem

			FItemList(i).forderno = rsget("orderno")
			FItemList(i).fdetailidx         = rsget("idx")
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
            
			rsget.movenext
			i=i+1
		loop
		rsget.close
	end sub
	
	'//admin/offshop/shopcscenter/order/ordermaster_detail.asp
    public Sub fGetCSASTotalCount()
        dim i,sqlStr,sqlsearch

		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch + " and m.orgmasteridx=" + FRectmasteridx + ""
		end if
		
		if frectcurrstate <> "" then
			sqlsearch = sqlsearch + " and m.currstate in (" + frectcurrstate + ")"
		end if
		
		if frectdeleteyn <> "" then
			sqlsearch = sqlsearch + " and m.deleteyn = '" + frectdeleteyn + "'"
		end if

		if frectshopid <> "" then
			sqlsearch = sqlsearch + " and om.shopid = '" + frectshopid + "'"
		end if		
		
        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_cs_master m"
        sqlStr = sqlStr + " join db_shop.dbo.tbl_shopjumun_master om"
        sqlStr = sqlStr + " 	on m.orgmasteridx = om.idx"
        sqlStr = sqlStr + " 	and om.cancelyn = 'N'"
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
    
	'//admin/offshop/shopcscenter/order/ordermaster_detail.asp	'//admin/offshop/shopcscenter/order/orderitemmaster.asp
	'//admin/offshop/shopcscenter/action/pop_cs_action_new.asp	'//admin/offshop/shopcscenter/action/pop_cs_action_new_process.asp
	'//admin/offshop/shopcscenter/action/cs_action_detail.asp	'//admin/offshop/shopcscenter/order/pop_order_receipt.asp
	public Sub fQuickSearchOrderMaster()
        dim sqlStr ,sqlsearch

		if FRectmasteridx <> "" then
			sqlsearch = sqlsearch & " and m.idx="& FRectmasteridx &""
		end if

		if FRectorderno <> "" then
			sqlsearch = sqlsearch & " and m.orderno="& FRectorderno &""
		end if

		if FRectcancelyn <> "" then
			sqlsearch = sqlsearch & " and m.cancelyn='"& FRectcancelyn &"'"
		end if
		
		sqlStr = " select top 1 " +vbcrlf
    	sqlStr = sqlStr + " m.idx,m.orderno ,m.shopid ,m.totalsum ,m.realsum ,m.jumundiv ,m.jumunmethod"
    	sqlStr = sqlStr + " ,m.shopregdate ,m.cancelyn ,m.regdate ,m.shopidx ,m.spendmile ,m.pointuserno"
    	sqlStr = sqlStr + " ,m.gainmile ,m.tableno ,m.cashsum ,m.cardsum ,m.GiftCardPaySum ,m.TenGiftCardPaySum"
    	sqlStr = sqlStr + " ,m.casherid ,m.CashReceiptNo ,m.CardAppNo ,m.CashreceiptGubun ,m.CardInstallment"
    	sqlStr = sqlStr + " ,m.TenGiftCardMatchCode ,m.refOrderNo ,m.IXyyyymmdd"
    	sqlStr = sqlStr + " ,u.shopname"
		sqlStr = sqlStr + " ,cu.username ,cu.hpno ,cu.email ,cu.onlineuserid"    	
    	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on m.shopid = u.userid"
		sqlStr = sqlStr & " 	and u.isusing='Y'"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_total_shop_card c"
		sqlStr = sqlStr & " 	on m.pointuserno = c.cardno"
		sqlStr = sqlStr & " 	and c.useyn = 'Y'"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_total_shop_user cu"
		sqlStr = sqlStr & " 	on c.userseq = cu.userseq"		
		sqlStr = sqlStr + " where 1=1 " & sqlsearch	
		sqlStr = sqlStr + " order by m.orderno desc"  

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount

        set FOneItem = new COrderItem

        if Not rsget.Eof then

			FOneItem.fmasteridx  	= rsget("idx")
			FOneItem.fshopid  	= rsget("shopid")
			FOneItem.fshopname  	= rsget("shopname")
			FOneItem.Forderno       = rsget("orderno")
			FOneItem.ftotalsum	        = rsget("totalsum")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.frealsum		= rsget("realsum")
			FOneItem.Fcancelyn	        = rsget("cancelyn")
			FOneItem.fcashsum	= rsget("cashsum")
			FOneItem.fcardsum		= rsget("cardsum")
			FOneItem.fpointuserno	= rsget("pointuserno")
			FOneItem.fonlineuserid	= rsget("onlineuserid")
			FOneItem.Fbuyname		= db2Html(rsget("username"))
			FOneItem.Fbuyhp		= db2Html(rsget("hpno"))
			FOneItem.Fbuyemail		= db2Html(rsget("email"))
			FOneItem.Freqname		= db2Html(rsget("username"))
			FOneItem.Freqhp		= db2Html(rsget("hpno"))
			FOneItem.Freqemail		= db2Html(rsget("email"))
			
        end if
        rsget.Close
    end Sub
    
	'/admin/offshop/shopcscenter/order/ordermaster_list.asp
	public Sub fQuickSearchOrderList()
		dim sqlStr, i , sqlsearch
		
		If FrectUserID <> "" Then
			sqlsearch = sqlsearch & " AND cu.onlineuserid = '" & FrectUserID & "' "
		End If
		
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
			sqlsearch = sqlsearch + " and cu.username = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectBuyHp<>"") then
			sqlsearch = sqlsearch + " and replace(cu.hpno,'-','')='" + replace(FRectBuyHp,"-","") + "'"
		end if

		if (FRectmail<>"") then
			sqlsearch = sqlsearch + " and cu.email='" + FRectmail + "'"
		end if

		If FrectCardNo <> "" Then
			sqlsearch = sqlsearch & " AND m.pointuserno = '" & FrectCardNo & "' "
		End If

		''갯수
		sqlStr = "select count(*) as cnt"
    	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on m.shopid = u.userid"
		sqlStr = sqlStr & " 	and u.isusing='Y'"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_total_shop_card c"
		sqlStr = sqlStr & " 	on m.pointuserno = c.cardno"
		sqlStr = sqlStr & " 	and c.useyn = 'Y'"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_total_shop_user cu"
		sqlStr = sqlStr & " 	on c.userseq = cu.userseq"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.close

		''데이타.
		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
    	sqlStr = sqlStr + " m.idx,m.orderno ,m.shopid ,m.totalsum ,m.realsum ,m.jumundiv ,m.jumunmethod"
    	sqlStr = sqlStr + " ,m.shopregdate ,m.cancelyn ,m.regdate ,m.shopidx ,m.spendmile ,m.pointuserno"
    	sqlStr = sqlStr + " ,m.gainmile ,m.tableno ,m.cashsum ,m.cardsum ,m.GiftCardPaySum ,m.TenGiftCardPaySum"
    	sqlStr = sqlStr + " ,m.casherid ,m.CashReceiptNo ,m.CardAppNo ,m.CashreceiptGubun ,m.CardInstallment"
    	sqlStr = sqlStr + " ,m.TenGiftCardMatchCode ,m.refOrderNo ,m.IXyyyymmdd"
    	sqlStr = sqlStr + " ,u.shopname"
    	sqlStr = sqlStr + " ,cu.username ,cu.hpno ,cu.email ,cu.onlineuserid"
    	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on m.shopid = u.userid"
		sqlStr = sqlStr & " 	and u.isusing='Y'"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_total_shop_card c"
		sqlStr = sqlStr & " 	on m.pointuserno = c.cardno"
		sqlStr = sqlStr & " 	and c.useyn = 'Y'"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_total_shop_user cu"
		sqlStr = sqlStr & " 	on c.userseq = cu.userseq"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by m.idx desc"

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
				
				FItemList(i).fmasteridx  	= rsget("idx")
				FItemList(i).fshopname  	= rsget("shopname")
				FItemList(i).Forderno       = rsget("orderno")
				FItemList(i).ftotalsum	        = rsget("totalsum")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).frealsum		= rsget("realsum")
				FItemList(i).Fcancelyn	        = rsget("cancelyn")				
				FItemList(i).Fbuyname		= db2Html(rsget("username"))		
				FItemList(i).fcashsum	= rsget("cashsum")
				FItemList(i).fcardsum		= rsget("cardsum")
				FItemList(i).fpointuserno	= rsget("pointuserno")				
				FItemList(i).fonlineuserid	= rsget("onlineuserid")
				'FItemList(i).Fcomment		= db2Html(rsget("comment"))
                
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