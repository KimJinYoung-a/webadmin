<%
'###########################################################
' Description : 매장 적정재고 관리 클래스
' Hieditor : 2011.07.13 한용민 생성
'###########################################################

Class cshortagestock_item
	Private Sub Class_Initialize()
		FOnlineOptaddbuyprice = 0
		FOnlineOptaddprice    = 0
		
		fdefaultmargin = 0	
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fisusing
	public fshopid
	public fitemgubun
	public fitemid
	public fitemoption
	public flogicsipgono
	public flogicsreipgono
	public fbrandipgono
	public fbrandreipgono
	public fsellno
	public fresellno
	public ferrsampleitemno
	public ferrbaditemno
	public ferrrealcheckno
	public fsysstockno
	public frealstockno
	public frequiredStock
	public fsell7days
	public fsell3days
	public fregdate
	public flastupdate
	public fmakerid
	public fshopitemname
	public fshopitemoptionname
	public fshopitemprice
	public fshopsuplycash
	public forgsellprice
	public fdiscountsellprice
	public fextbarcode
	public foffimgmain
	public foffimglist
	public foffimgsmall
	public fvatinclude
	public fshopbuyprice
	public fcentermwdiv
	public fcatecdl
	public fcatecdm
	public fcatecdn
	public fonofflinkyn
	public FimageSmall
	public frequire3daystock
	public frequire7daystock
	public frequire14daystock
	public frequire28daystock			
	public fchargediv
	public fcomm_cd
	public fdefaultmargin
	public fdefaultsuplymargin
	public fpreorderno
	public fpreordernofix
	public FOnlineSellcash
	public FOnlineBuycash
	public FOnlineOrgprice
	public FOnlineOptaddprice
    public FOnlineOptaddbuyprice
    	
	''가맹점 공급가
	public function GetFranchiseSuplycash()
		dim ishopsupycash

		''가맹점공급가가 0 인경우 기본 마진으로 구한다
		if Fshopbuyprice<>0 then
			ishopsupycash = Fshopbuyprice
		else
		    ''마진이 설정 안되있는경우 매입마진-5%
		    if IsNULL(fdefaultsuplymargin) or (fdefaultsuplymargin=0) then
		        IF (fdefaultmargin=0) then fdefaultmargin=35
		        ishopsupycash = CLng(Fshopitemprice * (100-(fdefaultmargin-5))/100)
		    else
			    ishopsupycash = CLng(Fshopitemprice * (100-fdefaultsuplymargin)/100)
			end if
		end if

		''공급가가 매입가보다 작은경우 공급가를 사용
		if (ishopsupycash<GetFranchiseBuycash) then ishopsupycash = GetFranchiseBuycash

		GetFranchiseSuplycash = ishopsupycash
	end function

	''직영점 공급가 : 가맹점과 동일
	public function GetOfflineSuplycash()
		GetOfflineSuplycash = GetFranchiseSuplycash
	end function

	''가맹점 공급시 매입가(업체로부터 매입하는가격)
	public function GetFranchiseBuycash()
		dim ibuycash
		''가맹점 매입가가 0 인경우 기본 마진으로 구한다
		if Fshopsuplycash<>0 then
			ibuycash = Fshopsuplycash
		else
		    IF (fdefaultmargin=0) then fdefaultmargin=35
			ibuycash = CLng(Fshopitemprice * (100-fdefaultmargin)/100)

			''온라인 매입가보다 큰경우 온라인 매입가를 사용(Fshopsuplycash 가 지정된 경우는 제외)
			''200906 FOnlineOptaddbuyprice 추가
			''온라인만 세일 하는 경우등 // 위탁->매입출고 인 경우 // 이 조건 제외 (2012-02-16)
		    ''if (FOnlinebuycash<>0) and (ibuycash>FOnlinebuycash+FOnlineOptaddbuyprice) then ibuycash=FOnlinebuycash+FOnlineOptaddbuyprice
		end if

		GetFranchiseBuycash = ibuycash
	end function

	''직영점 공급시 매입가(업체로부터 매입하는가격) : 가맹점과 동일
	public function GetOfflineBuycash()
		GetOfflineBuycash = GetFranchiseBuycash
	end function
				
	''유효재고
    public function getAvailStock()
        getAvailStock = FrealstockNo + Ferrsampleitemno + Ferrbaditemno
    end function
    	
	public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function
end class

class cshortagestock_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public Frectshopid
	public FRectisUsing
	public Frectmakerid
	public Frectitemid
	public Frectitemname
	public Frectgeneralbarcode
	public frectcdl
	public frectcdm
	public frectcds
	public Frectshortagetype
	public Frectcomm_cd
	public frectincludepreorder
	public frectsell7days
	public frectipgo
	public FRectOrder
	
	'/common/offshop/stock/shortagestock_shop.asp
	public sub fshortagestock_list()
		dim sqlStr,i ,sqlsearch

		if Frectcomm_cd <> "" then
			if Frectcomm_cd = "B099" then
				sqlsearch = sqlsearch & " and d.comm_cd in ('B011','B031')"
			elseif Frectcomm_cd = "B088" then
				sqlsearch = sqlsearch & " and d.comm_cd in ('B012','B022')"
			else
				sqlsearch = sqlsearch & " and d.comm_cd = '"&Frectcomm_cd&"'"
			end if
		end if
		
		if Frectshopid <> "" then
			sqlsearch = sqlsearch & " and s.shopid = '"&Frectshopid&"'"
		end if
		
		if Frectisusing <> "" then
			sqlsearch = sqlsearch & " and i.isusing = '"&Frectisusing&"'"
		end if
		
		if Frectmakerid <> "" then
			sqlsearch = sqlsearch & " and i.makerid = '"&Frectmakerid&"'"
		end if
		
		if Frectitemid <> "" then
			sqlsearch = sqlsearch & " and i.shopitemid = "&Frectitemid&""
		end if
		
		if Frectitemname <> "" then
			sqlsearch = sqlsearch & " and i.shopitemname like '%"&Frectitemname&"%'"
		end if
		
		if Frectgeneralbarcode <> "" then
			sqlsearch = sqlsearch & " and i.extbarcode = '"&Frectgeneralbarcode&"'"
		end if
		
        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "'"
        end if
        
        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and i.catecdm='" + FRectCDM + "'"
        end if
        
        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and i.catecdn='" + FRectCDS + "'"
        end if
        
		'기주문포함부족상품
		if FRectIncludePreOrder = "on" then
	        if Frectshortagetype="3" then
	    		sqlsearch = sqlsearch + " and (s.sell3days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	elseif Frectshortagetype="7" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	elseif Frectshortagetype="14" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*2) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	elseif Frectshortagetype="28" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*4) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	else
	    		sqlsearch = sqlsearch + " and s.preordernofix + db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) < 1"
	    		'sqlsearch = sqlsearch + " and (s.sell7days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	end if
		else
	        if Frectshortagetype="3" then
	    		sqlsearch = sqlsearch + " and (s.sell3days*1) - db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) > 0"
	    	elseif Frectshortagetype="7" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*1) - db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) > 0"
	    	elseif Frectshortagetype="14" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*2) - db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) > 0"
	    	elseif Frectshortagetype="28" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*4) - db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) > 0"
	    	else

	    	end if
		end if
		
		'/최근7일 판매건만
		if frectsell7days = "on" then
			sqlsearch = sqlsearch + " and s.sell7days > 0"
		end if
		if frectipgo = "on" then
			sqlsearch = sqlsearch + " and (s.logicsipgono+s.logicsreipgono > 0 or s.brandipgono+s.brandreipgono > 0)"
		end if

		'총 갯수 구하기
		sqlStr = "select count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_summary.dbo.tbl_current_shopstock_summary s"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr & " 	on s.itemgubun = i.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid = i.shopitemid"
		sqlStr = sqlStr & " 	and s.itemoption = i.itemoption"
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr & " 	on s.shopid = d.shopid"
		sqlStr = sqlStr & "		and i.makerid = d.makerid"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item ii"
		sqlStr = sqlStr & " 	on i.shopitemid = ii.itemid"
		sqlStr = sqlStr & " 	and i.itemgubun = '10'"
		sqlStr = sqlStr & " where i.shopitemid <> 0 and i.itemgubun <>'70'" & sqlsearch
		'sqlStr = sqlStr & " and (IsNULL(ii.sellyn,'')<>'N')"

		'response.write sqlStr &"<br>"						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " i.itemgubun, i.shopitemid, i.itemoption, i.makerid, i.shopitemname, i.shopitemoptionname"
		sqlStr = sqlStr & " ,i.shopitemprice, i.orgsellprice, i.discountsellprice, i.isusing"
		sqlStr = sqlStr & " ,i.extbarcode, i.offimgmain, i.offimglist, i.offimgsmall, i.vatinclude"
		sqlStr = sqlStr & " ,i.centermwdiv, i.catecdl, i.catecdm, i.catecdn,i.onofflinkyn"
		sqlStr = sqlStr & " ,(CASE when i.shopsuplycash=0 THEN convert(int,i.shopitemprice*(100-IsNULL(d.defaultmargin,35))/100) "
		sqlStr = sqlStr & " 	ELSE i.shopsuplycash END) as shopsuplycash"
		sqlStr = sqlStr & " ,(CASE when i.shopbuyprice=0 THEN convert(int,i.shopitemprice*(100-IsNULL(d.defaultsuplymargin,30))/100) "
		sqlStr = sqlStr & " 	ELSE i.shopbuyprice END) as shopbuyprice"
		sqlStr = sqlStr & " ,ii.smallimage"
		sqlStr = sqlStr & " ,s.shopid, isnull(s.logicsipgono,0) as logicsipgono, isnull(s.logicsreipgono,0) as logicsreipgono"
		sqlStr = sqlStr & " , isnull(s.brandipgono,0) as brandipgono, isnull(s.brandreipgono,0) as brandreipgono"
		sqlStr = sqlStr & " , isnull(s.sellno,0) as sellno, isnull(s.resellno,0) as resellno,isnull(s.errsampleitemno,0) as errsampleitemno"
		sqlStr = sqlStr & " , isnull(s.errbaditemno,0) as errbaditemno, isnull(s.errrealcheckno,0) as errrealcheckno"
		sqlStr = sqlStr & " , isnull(s.sysstockno,0) as sysstockno, isnull(s.realstockno,0) as realstockno, isnull(s.requiredStock,0) as requiredStock"
		sqlStr = sqlStr & " , isnull(s.sell7days,0) as sell7days, isnull(s.sell3days,0) as sell3days ,s.lastupdate"
		sqlStr = sqlStr & " , s.preorderno ,s.preordernofix"
		sqlStr = sqlStr & " ,( (s.sell3days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require3daystock"
		sqlStr = sqlStr & " ,( (s.sell7days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require7daystock"
		sqlStr = sqlStr & " ,( (s.sell7days*2) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require14daystock"
		'sqlStr = sqlStr & " ,( (s.sell7days*4) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require28daystock"
		sqlStr = sqlStr & " ,d.chargediv ,d.comm_cd ,d.defaultmargin ,d.defaultsuplymargin"
		sqlStr = sqlStr & " from db_summary.dbo.tbl_current_shopstock_summary s"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr & " 	on s.itemgubun = i.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid = i.shopitemid"
		sqlStr = sqlStr & " 	and s.itemoption = i.itemoption"
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr & " 	on s.shopid = d.shopid"
		sqlStr = sqlStr & "		and i.makerid = d.makerid"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item ii"
		sqlStr = sqlStr & " 	on i.shopitemid = ii.itemid"
		sqlStr = sqlStr & " 	and i.itemgubun = '10'"
		sqlStr = sqlStr & " where i.shopitemid <> 0 and i.itemgubun <>'70'" & sqlsearch
		'sqlStr = sqlStr & " and (IsNULL(ii.sellyn,'')<>'N')"
		'sqlStr = sqlStr & " order by d.comm_cd asc,i.makerid asc ,i.itemgubun asc,i.shopitemid asc ,i.itemoption asc"
		sqlStr = sqlStr + " order by i.itemgubun asc, i.shopitemid desc, i.itemoption"

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
				set FItemList(i) = new cshortagestock_item

				FItemList(i).fpreorderno = rsget("preorderno")
				FItemList(i).fpreordernofix = rsget("preordernofix")
				FItemList(i).fchargediv = rsget("chargediv")
				FItemList(i).fcomm_cd = rsget("comm_cd")
				FItemList(i).fdefaultmargin = rsget("defaultmargin")
				FItemList(i).fdefaultsuplymargin = rsget("defaultsuplymargin")			
				FItemList(i).frequire3daystock = rsget("require3daystock")
				FItemList(i).frequire7daystock = rsget("require7daystock")
				FItemList(i).frequire14daystock = rsget("require14daystock")
				'FItemList(i).frequire28daystock = rsget("require28daystock")
				FItemList(i).FimageSmall = db2html(rsget("smallimage"))
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fitemid = rsget("shopitemid")
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fmakerid = rsget("makerid")
				FItemList(i).fshopitemname = db2html(rsget("shopitemname"))
				FItemList(i).fshopitemoptionname = db2html(rsget("shopitemoptionname"))
				FItemList(i).fshopitemprice = rsget("shopitemprice")
				FItemList(i).fshopsuplycash = rsget("shopsuplycash")
				FItemList(i).forgsellprice = rsget("orgsellprice")
				FItemList(i).fdiscountsellprice = rsget("discountsellprice")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fextbarcode = rsget("extbarcode")
				FItemList(i).foffimgmain = rsget("offimgmain")
				FItemList(i).foffimglist = db2html(rsget("offimglist"))
				FItemList(i).foffimgsmall = db2html(rsget("offimgsmall"))
				FItemList(i).fvatinclude = rsget("vatinclude")
				FItemList(i).fshopbuyprice = rsget("shopbuyprice")
				FItemList(i).fcentermwdiv = rsget("centermwdiv")
				FItemList(i).fcatecdl = rsget("catecdl")
				FItemList(i).fcatecdm = rsget("catecdm")
				FItemList(i).fcatecdn = rsget("catecdn")
				FItemList(i).fonofflinkyn = rsget("onofflinkyn")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).flogicsipgono = rsget("logicsipgono")
				FItemList(i).flogicsreipgono = rsget("logicsreipgono")
				FItemList(i).fbrandipgono = rsget("brandipgono")
				FItemList(i).fbrandreipgono = rsget("brandreipgono")
				FItemList(i).fsellno = rsget("sellno")
				FItemList(i).fresellno = rsget("resellno")
				FItemList(i).ferrsampleitemno = rsget("errsampleitemno")
				FItemList(i).ferrbaditemno = rsget("errbaditemno")
				FItemList(i).ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).fsysstockno = rsget("sysstockno")
				FItemList(i).frealstockno = rsget("realstockno")
				FItemList(i).frequiredStock = rsget("requiredStock")
				FItemList(i).fsell7days = rsget("sell7days")
				FItemList(i).fsell3days = rsget("sell3days")
				FItemList(i).flastupdate = rsget("lastupdate")
				
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/common/offshop/stock/newitemstock_shop.asp
	public sub fnewitemstock_list()
		dim sqlStr,i ,sqlsearch , iStartDate

        if (FRectOrder="byrecent") then
            sqlStr = "select convert(varchar(10),dateadd(d,-14,getdate()),21) as stdt "
            
            'response.write sqlStr &"<Br>"
            rsget.Open sqlStr,dbget,1
    			iStartDate = rsget("stdt")
    		rsget.Close
        
			sqlsearch = sqlsearch + " and i.itemgubun <>'70'"
			sqlsearch = sqlsearch + " and i.regdate>'" & iStartDate & "'"	
		elseif FRectOrder="byonbest" then
		    sqlsearch = sqlsearch + " and i.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and ii.itemscore>0"
		elseif FRectOrder="byoffbest" then
		    sqlsearch = sqlsearch + " and i.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and s.sell7days>0"
        end if
        
		if Frectcomm_cd <> "" then
			if Frectcomm_cd = "B099" then
				sqlsearch = sqlsearch & " and d.comm_cd in ('B011','B031')"
			elseif Frectcomm_cd = "B077" then
				sqlsearch = sqlsearch & " and d.comm_cd in ('B013','B011','B031')"
			elseif Frectcomm_cd = "B088" then
				sqlsearch = sqlsearch & " and d.comm_cd in ('B012','B022')"
			else
				sqlsearch = sqlsearch & " and d.comm_cd = '"&Frectcomm_cd&"'"
			end if
		end if
		
		if Frectshopid <> "" then
			sqlsearch = sqlsearch & " and d.shopid = '"&Frectshopid&"'"
		end if
		
		if Frectisusing <> "" then
			sqlsearch = sqlsearch & " and i.isusing = '"&Frectisusing&"'"
		end if
		
		if Frectmakerid <> "" then
			sqlsearch = sqlsearch & " and i.makerid = '"&Frectmakerid&"'"
		end if
		
		if Frectitemid <> "" then
			sqlsearch = sqlsearch & " and i.shopitemid = "&Frectitemid&""
		end if
		
		if Frectitemname <> "" then
			sqlsearch = sqlsearch & " and i.shopitemname like '%"&Frectitemname&"%'"
		end if
		
		if Frectgeneralbarcode <> "" then
			sqlsearch = sqlsearch & " and i.extbarcode = '"&Frectgeneralbarcode&"'"
		end if
		
        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "'"
        end if
        
        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and i.catecdm='" + FRectCDM + "'"
        end if
        
        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and i.catecdn='" + FRectCDS + "'"
        end if
        
		'기주문포함부족상품
		if FRectIncludePreOrder = "on" then
	        if Frectshortagetype="3" then
	    		sqlsearch = sqlsearch + " and (s.sell3days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	elseif Frectshortagetype="7" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	elseif Frectshortagetype="14" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*2) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	elseif Frectshortagetype="28" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*4) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	else
	    		sqlsearch = sqlsearch + " and s.preordernofix + db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) < 1"
	    		'sqlsearch = sqlsearch + " and (s.sell7days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) > 0"
	    	end if
		else
	        if Frectshortagetype="3" then
	    		sqlsearch = sqlsearch + " and (s.sell3days*1) - db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) > 0"
	    	elseif Frectshortagetype="7" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*1) - db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) > 0"
	    	elseif Frectshortagetype="14" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*2) - db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) > 0"
	    	elseif Frectshortagetype="28" then
	    		sqlsearch = sqlsearch + " and (s.sell7days*4) - db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno) > 0"
	    	else

	    	end if
		end if
		
		'/최근7일 판매건만
		if frectsell7days = "on" then
			sqlsearch = sqlsearch + " and s.sell7days > 0"
		end if
		if frectipgo = "on" then
			sqlsearch = sqlsearch + " and (s.logicsipgono+s.logicsreipgono > 0 or s.brandipgono+s.brandreipgono > 0)"
		end if

		'총 갯수 구하기
		sqlStr = "select count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr & "		on i.makerid = d.makerid"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item ii"
		sqlStr = sqlStr & " 	on i.shopitemid = ii.itemid"
		sqlStr = sqlStr & " 	and i.itemgubun = '10'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o "
		sqlStr = sqlStr + "     on i.itemgubun='10'"
		sqlStr = sqlStr + "     and i.shopitemid=o.itemid"
		sqlStr = sqlStr + "     and i.itemoption=o.itemoption"
		sqlStr = sqlStr & "	left join db_summary.dbo.tbl_current_shopstock_summary s"		
		sqlStr = sqlStr & " 	on i.itemgubun = s.itemgubun"
		sqlStr = sqlStr & " 	and i.shopitemid = s.itemid"
		sqlStr = sqlStr & " 	and i.itemoption = s.itemoption"
		sqlStr = sqlStr & "		and d.shopid = s.shopid"
		sqlStr = sqlStr & " where i.shopitemid <> 0 and i.itemgubun <>'70' " & sqlsearch	'and (IsNULL(ii.sellyn,'')<>'N')

		'response.write sqlStr &"<br>"					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " i.itemgubun, i.shopitemid, i.itemoption, i.makerid, i.shopitemname, i.shopitemoptionname"
		sqlStr = sqlStr & " ,i.shopitemprice, i.orgsellprice, i.discountsellprice, i.isusing"
		sqlStr = sqlStr & " ,i.extbarcode, i.offimgmain, i.offimglist, i.offimgsmall, i.vatinclude"
		sqlStr = sqlStr & " ,i.centermwdiv, i.catecdl, i.catecdm, i.catecdn,i.onofflinkyn"
		sqlStr = sqlStr & " ,(CASE when i.shopsuplycash=0 THEN convert(int,i.shopitemprice*(100-IsNULL(d.defaultmargin,35))/100) "
		sqlStr = sqlStr & " 	ELSE i.shopsuplycash END) as shopsuplycash"
		sqlStr = sqlStr & " ,(CASE when i.shopbuyprice=0 THEN convert(int,i.shopitemprice*(100-IsNULL(d.defaultsuplymargin,30))/100) "
		sqlStr = sqlStr & " 	ELSE i.shopbuyprice END) as shopbuyprice"
		sqlStr = sqlStr & " ,ii.smallimage,IsNULL(ii.sellcash,0) as sellcash, IsNULL(ii.buycash,0) as buycash,IsNULL(ii.orgprice,0) as orgprice"
		sqlStr = sqlStr & " , isnull(s.logicsipgono,0) as logicsipgono, isnull(s.logicsreipgono,0) as logicsreipgono"
		sqlStr = sqlStr & " , isnull(s.brandipgono,0) as brandipgono, isnull(s.brandreipgono,0) as brandreipgono"
		sqlStr = sqlStr & " , isnull(s.sellno,0) as sellno, isnull(s.resellno,0) as resellno,isnull(s.errsampleitemno,0) as errsampleitemno"
		sqlStr = sqlStr & " , isnull(s.errbaditemno,0) as errbaditemno, isnull(s.errrealcheckno,0) as errrealcheckno"
		sqlStr = sqlStr & " , isnull(s.sysstockno,0) as sysstockno, isnull(s.realstockno,0) as realstockno, isnull(s.requiredStock,0) as requiredStock"
		sqlStr = sqlStr & " , isnull(s.sell7days,0) as sell7days, isnull(s.sell3days,0) as sell3days ,s.lastupdate"
		sqlStr = sqlStr & " , s.preorderno ,s.preordernofix"
		sqlStr = sqlStr & " ,( (s.sell3days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require3daystock"
		sqlStr = sqlStr & " ,( (s.sell7days*1) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require7daystock"
		sqlStr = sqlStr & " ,( (s.sell7days*2) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require14daystock"
		'sqlStr = sqlStr & " ,( (s.sell7days*4) - (db_summary.[dbo].[uf_replacezero](s.realstockNo+s.errsampleitemno+s.errbaditemno)+s.preordernofix) ) as require28daystock"
		sqlStr = sqlStr & " ,d.chargediv ,d.comm_cd ,d.defaultmargin ,d.defaultsuplymargin,d.shopid"
		sqlStr = sqlStr & " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr & "		on i.makerid = d.makerid"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item ii"
		sqlStr = sqlStr & " 	on i.shopitemid = ii.itemid"
		sqlStr = sqlStr & " 	and i.itemgubun = '10'"	
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o "
		sqlStr = sqlStr + "     on i.itemgubun='10'"
		sqlStr = sqlStr + "     and i.shopitemid=o.itemid"
		sqlStr = sqlStr + "     and i.itemoption=o.itemoption"
		sqlStr = sqlStr & "	left join db_summary.dbo.tbl_current_shopstock_summary s"		
		sqlStr = sqlStr & " 	on i.itemgubun = s.itemgubun"
		sqlStr = sqlStr & " 	and i.shopitemid = s.itemid"
		sqlStr = sqlStr & " 	and i.itemoption = s.itemoption"
		sqlStr = sqlStr & "		and d.shopid = s.shopid"
		sqlStr = sqlStr & " where i.shopitemid <> 0 and i.itemgubun <>'70'" & sqlsearch	'and (IsNULL(ii.sellyn,'')<>'N')

        if (FRectOrder="byrecent") then
			sqlStr = sqlStr & " order by i.regdate desc"
		elseif FRectOrder="byonbest" then
			sqlStr = sqlStr & " order by ii.itemscore desc,i.itemgubun asc, i.shopitemid desc, i.itemoption"
		elseif FRectOrder="byoffbest" then
			sqlStr = sqlStr + " order by s.sell7days desc,i.itemgubun asc, i.shopitemid desc, i.itemoption"
		else
			'sqlStr = sqlStr & " order by d.comm_cd asc,i.makerid asc ,i.itemgubun asc,i.shopitemid asc ,i.itemoption asc"
			sqlStr = sqlStr & " order by i.itemgubun asc, i.shopitemid desc, i.itemoption"
        end if
		
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
				set FItemList(i) = new cshortagestock_item

                ''옵션 추가금액
    			FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
    			FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")				
				FItemList(i).FOnlineOrgprice	= rsget("orgprice")
				FItemList(i).FOnlineSellcash	= rsget("sellcash")
				FItemList(i).FOnlineBuycash		= rsget("buycash")
				FItemList(i).fpreorderno = rsget("preorderno")
				FItemList(i).fpreordernofix = rsget("preordernofix")
				FItemList(i).fchargediv = rsget("chargediv")
				FItemList(i).fcomm_cd = rsget("comm_cd")
				FItemList(i).fdefaultmargin = rsget("defaultmargin")
				FItemList(i).fdefaultsuplymargin = rsget("defaultsuplymargin")			
				FItemList(i).frequire3daystock = rsget("require3daystock")
				FItemList(i).frequire7daystock = rsget("require7daystock")
				FItemList(i).frequire14daystock = rsget("require14daystock")
				'FItemList(i).frequire28daystock = rsget("require28daystock")
				FItemList(i).FimageSmall = db2html(rsget("smallimage"))
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fitemid = rsget("shopitemid")
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fmakerid = rsget("makerid")
				FItemList(i).fshopitemname = db2html(rsget("shopitemname"))
				FItemList(i).fshopitemoptionname = db2html(rsget("shopitemoptionname"))
				FItemList(i).fshopitemprice = rsget("shopitemprice")
				FItemList(i).fshopsuplycash = rsget("shopsuplycash")
				FItemList(i).forgsellprice = rsget("orgsellprice")
				FItemList(i).fdiscountsellprice = rsget("discountsellprice")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fextbarcode = rsget("extbarcode")
				FItemList(i).foffimgmain = rsget("offimgmain")
				FItemList(i).foffimglist = db2html(rsget("offimglist"))
				FItemList(i).foffimgsmall = db2html(rsget("offimgsmall"))
				FItemList(i).fvatinclude = rsget("vatinclude")
				FItemList(i).fshopbuyprice = rsget("shopbuyprice")
				FItemList(i).fcentermwdiv = rsget("centermwdiv")
				FItemList(i).fcatecdl = rsget("catecdl")
				FItemList(i).fcatecdm = rsget("catecdm")
				FItemList(i).fcatecdn = rsget("catecdn")
				FItemList(i).fonofflinkyn = rsget("onofflinkyn")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).flogicsipgono = rsget("logicsipgono")
				FItemList(i).flogicsreipgono = rsget("logicsreipgono")
				FItemList(i).fbrandipgono = rsget("brandipgono")
				FItemList(i).fbrandreipgono = rsget("brandreipgono")
				FItemList(i).fsellno = rsget("sellno")
				FItemList(i).fresellno = rsget("resellno")
				FItemList(i).ferrsampleitemno = rsget("errsampleitemno")
				FItemList(i).ferrbaditemno = rsget("errbaditemno")
				FItemList(i).ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).fsysstockno = rsget("sysstockno")
				FItemList(i).frealstockno = rsget("realstockno")
				FItemList(i).frequiredStock = rsget("requiredStock")
				FItemList(i).fsell7days = rsget("sell7days")
				FItemList(i).fsell3days = rsget("sell3days")
				FItemList(i).flastupdate = rsget("lastupdate")
				
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	public sub fnewitemstock_list_datamart2()
		dim sqlStr,i ,sqlsearch , iStartDate
		
		sqlStr = "exec [dbo].[sp_Ten_API_get_ShortageStockShop_For_Foreign_CNT] '"&Frectshopid&"'"
		db3_rsget.Open sqlStr,db3_dbget,1
			FTotalCount = db3_rsget("cnt")
		db3_rsget.Close
		
		if FTotalCount < 1 then exit sub
		
		sqlStr = "exec [dbo].[sp_Ten_API_get_ShortageStockShop_For_Foreign_LIST] "&FPageSize&","&FCurrPage&",'"&Frectshopid&"'"
		
		db3_rsget.CursorLocation = adUseClient
    	db3_rsget.CursorType = adOpenStatic
    	db3_rsget.LockType = adLockOptimistic
    	db3_rsget.Open sqlStr, db3_dbget

        FResultCount =db3_rsget.RecordCount
        if (FResultCount<1) then FResultCount=0

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.EOF
				set FItemList(i) = new cshortagestock_item

                ''옵션 추가금액
    			FItemList(i).FOnlineOptaddprice = db3_rsget("optaddprice")
    			FItemList(i).FOnlineOptaddbuyprice = db3_rsget("optaddbuyprice")				
				FItemList(i).FOnlineOrgprice	= db3_rsget("orgprice")
				FItemList(i).FOnlineSellcash	= db3_rsget("sellcash")
				FItemList(i).FOnlineBuycash		= db3_rsget("buycash")
				FItemList(i).fpreorderno = db3_rsget("preorderno")
				FItemList(i).fpreordernofix = db3_rsget("preordernofix")
				FItemList(i).fchargediv = db3_rsget("chargediv")
				FItemList(i).fcomm_cd = db3_rsget("comm_cd")
				FItemList(i).fdefaultmargin = db3_rsget("defaultmargin")
				FItemList(i).fdefaultsuplymargin = db3_rsget("defaultsuplymargin")			
				FItemList(i).frequire3daystock = db3_rsget("require3daystock")
				FItemList(i).frequire7daystock = db3_rsget("require7daystock")
				FItemList(i).frequire14daystock = db3_rsget("require14daystock")
				'FItemList(i).frequire28daystock = db3_rsget("require28daystock")
				FItemList(i).FimageSmall = db2html(db3_rsget("smallimage"))
				FItemList(i).fitemgubun = db3_rsget("itemgubun")
				FItemList(i).fitemid = db3_rsget("shopitemid")
				FItemList(i).fitemoption = db3_rsget("itemoption")
				FItemList(i).fmakerid = db3_rsget("makerid")
				FItemList(i).fshopitemname = db2html(db3_rsget("shopitemname"))
				FItemList(i).fshopitemoptionname = db2html(db3_rsget("shopitemoptionname"))
				FItemList(i).fshopitemprice = db3_rsget("shopitemprice")
				FItemList(i).fshopsuplycash = db3_rsget("shopsuplycash")
				FItemList(i).forgsellprice = db3_rsget("orgsellprice")
				FItemList(i).fdiscountsellprice = db3_rsget("discountsellprice")
				FItemList(i).fisusing = db3_rsget("isusing")
				FItemList(i).fextbarcode = db3_rsget("extbarcode")
				FItemList(i).foffimgmain = db3_rsget("offimgmain")
				FItemList(i).foffimglist = db2html(db3_rsget("offimglist"))
				FItemList(i).foffimgsmall = db2html(db3_rsget("offimgsmall"))
				FItemList(i).fvatinclude = db3_rsget("vatinclude")
				FItemList(i).fshopbuyprice = db3_rsget("shopbuyprice")
				FItemList(i).fcentermwdiv = db3_rsget("centermwdiv")
				FItemList(i).fcatecdl = db3_rsget("catecdl")
				FItemList(i).fcatecdm = db3_rsget("catecdm")
				FItemList(i).fcatecdn = db3_rsget("catecdn")
				FItemList(i).fonofflinkyn = db3_rsget("onofflinkyn")
				FItemList(i).fshopid = db3_rsget("shopid")
				FItemList(i).flogicsipgono = db3_rsget("logicsipgono")
				FItemList(i).flogicsreipgono = db3_rsget("logicsreipgono")
				FItemList(i).fbrandipgono = db3_rsget("brandipgono")
				FItemList(i).fbrandreipgono = db3_rsget("brandreipgono")
				FItemList(i).fsellno = db3_rsget("sellno")
				FItemList(i).fresellno = db3_rsget("resellno")
				FItemList(i).ferrsampleitemno = db3_rsget("errsampleitemno")
				FItemList(i).ferrbaditemno = db3_rsget("errbaditemno")
				FItemList(i).ferrrealcheckno = db3_rsget("errrealcheckno")
				FItemList(i).fsysstockno = db3_rsget("sysstockno")
				FItemList(i).frealstockno = db3_rsget("realstockno")
				FItemList(i).frequiredStock = db3_rsget("requiredStock")
				FItemList(i).fsell7days = db3_rsget("sell7days")
				FItemList(i).fsell3days = db3_rsget("sell3days")
				FItemList(i).flastupdate = db3_rsget("lastupdate")
				
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
								
				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	end sub
	
	Private Sub Class_Initialize()
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
end class
%>