<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2012.03.20 한용민 생성
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
	public Fdefaultbeasongdiv		'배송구분 2:업체배송, 0:매장배송
	public Fshopmisend
	public Fupchemisend
	public fdupchestats0
	public fdupchestats2
	public fonlineuserid
	public Freqemail
	
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

	'//common/offshop/shopcscenter/upche_viewordermaster.asp
	public Sub fSearchJumunList()
		dim sqlStr , sqlsearch ,i

		if FRectmasteridx<>"" then
			sqlsearch = sqlsearch + " and m.idx='" + FRectmasteridx + "'"
		end if

		if FRectDesignerID <> "" then
			sqlsearch = sqlsearch + " and d.makerid='" + FRectDesignerID + "'"
		end if

        if (FRectIpkumdiv<>"") then
            sqlsearch = sqlsearch + FRectIpkumdiv
        end if

		''총 갯수
		sqlStr = "select count(*) as cnt"
    	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m"
    	sqlStr = sqlStr + " join db_shop.dbo.tbl_shopjumun_detail d"
    	sqlStr = sqlStr + " 	on m.orderno = d.orderno"
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
		rsget.Close

		''데이타.
		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
    	sqlStr = sqlStr + " m.idx,m.orderno ,m.shopid ,m.totalsum ,m.realsum ,m.jumundiv ,m.jumunmethod"
    	sqlStr = sqlStr + " ,m.shopregdate ,m.cancelyn ,m.regdate ,m.shopidx ,m.spendmile ,m.pointuserno"
    	sqlStr = sqlStr + " ,m.gainmile ,m.tableno ,m.cashsum ,m.cardsum ,m.GiftCardPaySum ,m.TenGiftCardPaySum"
    	sqlStr = sqlStr + " ,m.casherid ,m.CashReceiptNo ,m.CardAppNo ,m.CashreceiptGubun ,m.CardInstallment"
    	sqlStr = sqlStr + " ,m.TenGiftCardMatchCode ,m.refOrderNo ,m.IXyyyymmdd"
		sqlStr = sqlStr + " ,d.itemid , d.Itemoption ,d.itemno ,d.itemgubun ,(d.cancelyn) as detailcancelyn"  	
    	sqlStr = sqlStr + " ,d.sellprice ,d.itemname ,d.itemoptionname ,u.shopname"
		sqlStr = sqlStr + " ,cu.username ,cu.hpno ,cu.email ,cu.onlineuserid"    	
    	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m"
    	sqlStr = sqlStr + " join db_shop.dbo.tbl_shopjumun_detail d"
    	sqlStr = sqlStr + " 	on m.orderno = d.orderno"
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
				
				FItemList(i).fmasteridx  	= rsget("idx")
				FItemList(i).fshopname  	= rsget("shopname")
				FItemList(i).Forderno       = rsget("orderno")
				FItemList(i).ftotalsum	        = rsget("totalsum")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).frealsum		= rsget("realsum")
				FItemList(i).Fcancelyn	        = rsget("cancelyn")
				FItemList(i).fcashsum	= rsget("cashsum")
				FItemList(i).fcardsum		= rsget("cardsum")
				FItemList(i).fpointuserno	= rsget("pointuserno")
				FItemList(i).fonlineuserid	= rsget("onlineuserid")
				FItemList(i).Fbuyname		= db2Html(rsget("username"))
				FItemList(i).Fbuyhp		= db2Html(rsget("hpno"))
				FItemList(i).Fbuyemail		= db2Html(rsget("email"))
				FItemList(i).Freqname		= db2Html(rsget("username"))
				FItemList(i).Freqhp		= db2Html(rsget("hpno"))
				FItemList(i).Freqemail		= db2Html(rsget("email"))
				FItemList(i).fdetailcancelyn 			  = rsget("detailcancelyn")
				FItemList(i).fsellprice 			  = rsget("sellprice")
				FItemList(i).fitemgubun 			  = rsget("itemgubun")
    			FItemList(i).FItemid 			  = rsget("itemid")
    			FItemList(i).FItemoption 			  = rsget("Itemoption")
    			FItemList(i).FItemname 		  = db2html(rsget("itemname"))
    			FItemList(i).fitemoptionname     	  = db2html(rsget("itemoptionname"))
    			FItemList(i).fitemno           = rsget("itemno")

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
%>