<%
'####################################################
' Description : 할인관리 클래스
' History : 2010.12.01 한용민 생성
'####################################################

Class csale_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fsaleitem_idx
	public fshopname
	public fcomm_name
	public fsale_startdate
	public fsale_enddate
	public fsale_rate
	public fpoint_rate
	public fsale_margin
	public fsale_marginvalue
	public fsale_status
	public fsale_code
	public fitemid
	public fsaleprice
	public fsalesupplycash
	public fsaleItem_status
	public flimitno
	public forglimityn
	public fmakerid
	public fshopitemname
	public fsmallimage
	public fshopitemprice
	public fshopsuplycash
	public forgsellprice
	public fshopbuyprice
	public fcentermwdiv
	public fisusing
	public fitemoption
	public fshopsuplycash_org
	public fshopid
	public fsaleyn
	public fitemgubun
	public fsale_name
	public fevt_code
	public fsale_shopmargin
	public fsale_shopmarginvalue
	public fsaleshopsupplycash
	public fcomm_cd
	public fpossaleprice

	public function getCalcuMargin(imarginType,iMValue)
	    dim halfbuyprc
	    getCalcuMargin = 0

	    if (imarginType=1) then ''동일마진
	    	if (fcomm_cd="B012") or (fcomm_cd="B013") or (fcomm_cd="B011") then
			    if (fsaleprice=0) then
			        getCalcuMargin =0
			    else
    			    getCalcuMargin = 100-fix(fshopsuplycash/forgsellprice*10000)/100
    			end if
	    	else
            	getCalcuMargin = 0
            end if
        'elseif (imarginType=2) then ''업체부담
        elseif (imarginType=3) then ''반반부담 fsale_rate 필수
			if (fcomm_cd="B012") or (fcomm_cd="B013") or (fcomm_cd="B011") then
			    if (fsaleprice=0) then
			        getCalcuMargin =0
			    else
			        halfbuyprc = (fshopitemprice-fsaleprice)/2
    			    getCalcuMargin = 100-fix((fshopsuplycash - halfbuyprc)/fsaleprice*10000)/100
    			end if
			else
                getCalcuMargin = 0
		    end if
        elseif (imarginType=4) then ''10x10부담
            getCalcuMargin = 0
        elseif (imarginType=5) then ''직접설정
	    	if (fcomm_cd="B012") or (fcomm_cd="B013") or (fcomm_cd="B011") then
			    if (fsaleprice=0) then
			        getCalcuMargin =0
			    else
    			    getCalcuMargin = iMValue
    			end if
	    	else
            	getCalcuMargin = 0
            end if
        end if
    end function

	public function getCalcuShopMargin(imarginType,iMValue)
	    dim halfbuyprc
	    getCalcuShopMargin = 0

	    if (imarginType=1) then ''동일마진
	    	if (fcomm_cd="B012") or (fcomm_cd="B013") or (fcomm_cd="B011") then
			    if (fsaleprice=0) then
			        getCalcuShopMargin =0
			    else
    			    getCalcuShopMargin = 100-fix(fshopbuyprice/forgsellprice*10000)/100
    			end if
	    	else
            	getCalcuShopMargin = 0
            end if
        'elseif (imarginType=2) then ''업체부담
        elseif (imarginType=3) then ''반반부담
			if (fcomm_cd="B012") or (fcomm_cd="B013") or (fcomm_cd="B011") then
			    if (fsaleprice=0) then
			        getCalcuShopMargin =0
			    else
			        halfbuyprc = (fshopitemprice-fsaleprice)/2
    			    getCalcuShopMargin = 100-fix((fshopbuyprice - halfbuyprc)/fsaleprice*10000)/100
    			end if
			else
                getCalcuShopMargin = 0
		    end if
        elseif (imarginType=4) then ''10x10부담
            getCalcuShopMargin = 0
        elseif (imarginType=5) then ''직접설정
	    	if (fcomm_cd="B012") or (fcomm_cd="B013") or (fcomm_cd="B011") then
			    if (fsaleprice=0) then
			        getCalcuShopMargin =0
			    else
    			    getCalcuShopMargin = iMValue
    			end if
	    	else
            	getCalcuShopMargin = 0
            end if
        end if
    end function
end Class

Class csale_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public frectsale_code

	'/admin/offshop/sale/saleitemproc.asp 	'/admin/offshop/sale/saleproc.asp
	public sub getsaledetail()
		dim sqlStr,i , strSearch

		if frectsale_code <> "" then
			strSearch = strSearch + " and sale_code = "&frectsale_code&"" + vbcrlf
		end if

		sqlStr = " SELECT top 1"
		sqlStr = sqlStr & " sale_startdate, sale_enddate, sale_rate, point_rate, sale_margin, sale_marginvalue, sale_status"
		sqlStr = sqlStr & " , shopid, sale_shopmargin, sale_shopmarginvalue, sale_code"
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_sale_off"
		sqlStr = sqlStr & " where 1=1 "  & strSearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FtotalCount = rsget.RecordCount

        set FOneItem = new csale_oneitem

        if Not rsget.Eof then
			FOneItem.fpoint_rate = rsget("point_rate")
			FOneItem.fsale_startdate = db2html(rsget("sale_startdate"))
			FOneItem.fsale_enddate = db2html(rsget("sale_enddate"))
			FOneItem.fsale_rate = rsget("sale_rate")
			FOneItem.fsale_margin = rsget("sale_margin")
			FOneItem.fsale_marginvalue = rsget("sale_marginvalue")
			FOneItem.fsale_status	= rsget("sale_status")
			FOneItem.fshopid	= rsget("shopid")
			FOneItem.fsale_shopmargin	= rsget("sale_shopmargin")
			FOneItem.fsale_shopmarginvalue	= rsget("sale_shopmarginvalue")
        end if
        rsget.Close
	end sub

	 '/admin/offshop/sale/saleproc.asp
	public sub getsalenew()
		dim sqlStr,i , strSearch

		sqlStr = " SELECT top 1"
		sqlStr = sqlStr & " sale_startdate, sale_enddate, sale_rate, point_rate, sale_margin, sale_marginvalue, sale_status"
		sqlStr = sqlStr & " , shopid, sale_shopmargin, sale_shopmarginvalue, sale_code"
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_sale_off"
		sqlStr = sqlStr & " order by sale_code desc"

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FtotalCount = rsget.RecordCount

        set FOneItem = new csale_oneitem

        if Not rsget.Eof then
			FOneItem.fpoint_rate = rsget("point_rate")
			FOneItem.fsale_startdate = db2html(rsget("sale_startdate"))
			FOneItem.fsale_enddate = db2html(rsget("sale_enddate"))
			FOneItem.fsale_rate = rsget("sale_rate")
			FOneItem.fsale_margin = rsget("sale_margin")
			FOneItem.fsale_marginvalue = rsget("sale_marginvalue")
			FOneItem.fsale_status	= rsget("sale_status")
			FOneItem.fshopid	= rsget("shopid")
			FOneItem.fsale_shopmargin	= rsget("sale_shopmargin")
			FOneItem.fsale_shopmarginvalue	= rsget("sale_shopmarginvalue")
			FOneItem.fsale_code	= rsget("sale_code")
        end if
        rsget.Close
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
end Class

Class CSale
	public FSCode
	public FECode
	public FTotCnt
	public FCPage
	public FPSize
	public fshopid
	public FSearchTxt
	public FSearchType
	public FBrand
	public FDateType
	public FSDate
	public FEDate
	public FSStatus
	public FSName
	public FSRate
	public fpoint_rate
	public FSMargin
	public FEGroupCode
	public FSRegdate
	public FSUsing
	public FSAdminid
	public FOpenDate
	public FSMarginValue
	public FCloseDate
	public frectshopid
	public fsale_shopmarginvalue
	public fsale_shopmargin

	'== 할인관리 리스트 가져오기 	'//admin/offshop/sale/salelist.asp
	public Function fnGetSaleList
		Dim strSqlCnt, strSql, strSearch,iDelCnt

		strSearch = ""

		IF frectshopid <> "" THEN
			strSearch  = strSearch & " and shopid = '"&frectshopid&"'"
		END IF

		IF FECode <> "" THEN
			strSearch = " and evt_code ="&FECode
		END IF

		IF FSearchTxt <> "" THEN
			IF FSearchType = 1 THEN
				strSearch = strSearch & " and sale_code = "&FSearchTxt
			ELSEIF FSearchType= 2 THEN
				strSearch = strSearch & " and evt_code = "&FSearchTxt
			ELSEIF FSearchType=3 THEN
				strSearch = strSearch & " and sale_name like '%"& FSearchTxt &"%' "
			END IF
		END IF

		IF FSDate <> "" AND FEDate <> "" THEN
			if CStr(FDateType) = "S" THEN
				strSearch  = strSearch & " and  datediff(day, '"&FSDate&"', sale_startdate) >= 0 and  datediff(day,'"&FEDate&"', sale_startdate) <=0  "
			elseif CStr(FDateType) = "E" THEN
				strSearch  = strSearch & " and  datediff(day,'"&FSDate&"',sale_enddate) >= 0 and  datediff(day,'"&FEDate&"',sale_enddate) <=0  "
			end if
		END IF

		IF FSStatus <> "" THEN
			strSearch = strSearch & " and sale_status = "&FSStatus
		END IF

		strSqlCnt = " SELECT COUNT(*)"&_
					" FROM [db_shop].[dbo].[tbl_sale_off] a"&_
					" left join db_shop.dbo.tbl_shop_user u"&_
					" 	on a.shopid=u.userid"&_
					" WHERE sale_using = 1 " & strSearch

		'response.write strSqlCnt & "<Br>"
		rsget.Open strSqlCnt,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = " SELECT TOP "&FPSize&"  [sale_code], [sale_name], [sale_rate], [sale_margin], [evt_code], [evtgroup_code]"&_
					", [sale_startdate], [sale_enddate], [sale_status], [availPayType], [regdate], [sale_using], [adminid] "&_
					", (select count(itemid) from [db_shop].dbo.tbl_saleItem_off  where sale_code = A.sale_code ) as saleitem_cnt "&_
					", sale_marginvalue, opendate, closedate ,shopid ,sale_shopmargin, point_rate, u.shopname"&_
					" FROM [db_shop].[dbo].[tbl_sale_off] as A "&_
					" left join db_shop.dbo.tbl_shop_user u"&_
					" 	on a.shopid=u.userid"&_
					" WHERE sale_using = 1 AND sale_code <= ( " &_
					" 		SELECT Min(sale_code) FROM (" &_
					" 			SELECT TOP "&iDelCnt&" sale_code "&_
					" 			FROM [db_shop].[dbo].[tbl_sale_off] WHERE sale_using =1 "&strSearch&" ORDER BY sale_code DESC" &_
					" 		) as T" &_
					" 										) " & strSearch &_
					" ORDER BY sale_code DESC "

			'response.write strSql &"<Br>"
			rsget.Open strSql,dbget,3
			IF not rsget.EOF THEN
				fnGetSaleList = rsget.getRows()
			End IF
			rsget.Close
		END IF
	End Function

	'== 할인관리 내용가져오기 	'//admin/offshop/sale/saleReg.asp
	public Function fnGetSaleConts
		Dim strSql

		strSql = " SELECT [sale_code], [sale_name], [sale_rate], point_rate, [sale_margin], [evt_code], [evtgroup_code] "&_
				", [sale_startdate], [sale_enddate], [sale_status], [availPayType], [regdate], [sale_using] "&_
				" , [adminid], [opendate], [closedate],sale_marginvalue , shopid ,sale_shopmargin,sale_shopmarginvalue"&_
				" FROM [db_shop].[dbo].[tbl_sale_off] "&_
				" WHERE sale_code = "&FSCode

		'response.write strSql &"<br>"
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fshopid 	= rsget("shopid")
				FSCode 		= rsget("sale_code")
				FSName 		= rsget("sale_name")
				FSRate 		= rsget("sale_rate")
				fpoint_rate = rsget("point_rate")
				FSMargin 	= rsget("sale_margin")
				FECode 		= rsget("evt_code")
				FEGroupCode = rsget("evtgroup_code")
				FSDate 		= rsget("sale_startdate")
				FEDate		= rsget("sale_enddate")
				FSStatus 	= rsget("sale_status")
				FSRegdate 	= rsget("regdate")
				FSUsing 	= rsget("sale_using")
				FSAdminid 	= rsget("adminid")
				FOpenDate	= rsget("opendate")
				FCloseDate	= rsget("closedate")
				FSMarginValue	= rsget("sale_marginvalue")
				fsale_shopmargin	= rsget("sale_shopmargin")
				fsale_shopmarginvalue	= rsget("sale_shopmarginvalue")
			END IF
		rsget.close
	End Function

END Class

'할인상품
Class CSaleItem
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public FSCode
	public FTotCnt
	public FCPage
	public FPSize
	public FItemid
	public FBrand
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FChkPreSale
	public FSearchType
	public FSearchTxt
	public FSStatus
	public frectshopid
	public FRectDesigner
	public frectitemid
	public FRectItemName
	public frectadminvspos

	'//admin/offshop/sale/saleItemReg.asp 	'//admin/offshop/sale/saleitemlist.asp
	public sub fnGetSaleItemList
		dim sqlStr,i , strSearch

		if frectadminvspos <> "" then
			strSearch = strSearch + " and si.saleitem_status = 6"
			strSearch = strSearch + " and s.sale_status = 6"
			strSearch = strSearch + " and si.saleprice <> li.lcprice"
		end if
		if FRectItemId<>"" then
			strSearch = strSearch + " and i.shopitemid=" + CStr(FRectItemId)
		end if

		if FRectItemName<>"" then
			strSearch = strSearch + " and i.shopitemname like '%" + FRectItemName + "%'"
		end if

		if FRectDesigner<>"" then
			strSearch = strSearch + " and i.makerid='" + FRectDesigner + "'"
		end if

		IF frectshopid <> "" THEN
			strSearch  = strSearch & " and s.shopid = '"&frectshopid&"'"
		END IF

		IF FSearchTxt <> "" THEN
			IF FSearchType = 1 THEN
				strSearch = strSearch & " and si.sale_code = "&FSearchTxt
			ELSEIF FSearchType= 2 THEN
				strSearch = strSearch & " and s.evt_code = "&FSearchTxt
			ELSEIF FSearchType=3 THEN
				strSearch = strSearch & " and s.sale_name like '%"& FSearchTxt &"%' "
			END IF
		END IF

		IF FSStatus <> "" THEN
			strSearch = strSearch & " and s.sale_status = "&FSStatus
		END IF

		IF FSCode <> "" THEN
			strSearch  = strSearch & " and si.sale_code = "&FSCode&""
		END IF

		sqlStr = "select "
		sqlStr = sqlStr & " count(*) as cnt"
		sqlStr = sqlStr & " FROM [db_shop].[dbo].[tbl_saleItem_off] si"
		sqlStr = sqlStr & " join [db_shop].[dbo].[tbl_sale_off] s"
		sqlStr = sqlStr & " 	on si.sale_code = s.sale_code"
		sqlStr = sqlStr & " join [db_shop].[dbo].[tbl_shop_item] i"
		sqlStr = sqlStr & " 	on si.itemid = i.shopitemid"
		sqlStr = sqlStr & " 	and si.itemgubun = i.itemgubun"
		sqlStr = sqlStr & " 	and si.itemoption = i.itemoption"
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item ii"
		sqlStr = sqlStr & " 	on si.itemid=ii.itemid"
		sqlStr = sqlStr & " 	and ii.itemgubun='10'"
		sqlStr = sqlStr & " left join [db_shop].dbo.tbl_shop_locale_item li"
		sqlStr = sqlStr & " 	on si.itemid=li.shopitemid"
		sqlStr = sqlStr & " 	and si.itemgubun = li.itemgubun"
		sqlStr = sqlStr & " 	and si.itemoption = li.itemoption"
		sqlStr = sqlStr & " 	and s.shopid=li.shopid"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_designer sd"
		sqlStr = sqlStr & " 	on s.shopid=sd.shopid"
		sqlStr = sqlStr & " 	and i.makerid=sd.makerid"
		sqlStr = sqlStr & " left join db_jungsan.dbo.tbl_jungsan_comm_code c"
		sqlStr = sqlStr & " 	on sd.comm_cd=c.comm_cd"
		sqlStr = sqlStr & " 	and c.comm_group='Z002'"
		sqlStr = sqlStr & " WHERE 1=1 " & strSearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " si.saleitem_idx, si.sale_code, si.itemid, si.saleprice, si.salesupplycash, si.saleItem_status"
		sqlStr = sqlStr & " , si.limitno, si.orglimityn ,si.saleshopsupplycash, si.point_rate, i.makerid, i.shopitemname"
		sqlStr = sqlStr & " , (case when i.itemgubun = '90' then i.offimgsmall else ii.smallimage end) as smallimage"
		sqlStr = sqlStr & " ,(case when li.shopid is not null then 'Y' else 'N' end) as 'sailyn'"
		sqlStr = sqlStr & " ,li.lcprice as possaleprice ,i.shopitemprice"
		sqlStr = sqlStr & " ,(CASE"
		sqlStr = sqlStr & " 	when sd.comm_cd in ('B012','B013','B011') and i.shopsuplycash=0"	'//업체위탁,출고위탁,텐바이텐위탁이고 매입가가0일경우
		sqlStr = sqlStr & " 		THEN convert(int,i.shopitemprice*(100-IsNULL(sd.defaultmargin,100))/100) "
		sqlStr = sqlStr & " 	ELSE i.shopsuplycash END) as 'shopsuplycash'"
		sqlStr = sqlStr & " ,(CASE"
		sqlStr = sqlStr & " 	when sd.comm_cd in ('B012','B013','B011') and i.shopbuyprice=0"		'//업체위탁,출고위탁,텐바이텐위탁이고 매입가가0일경우
		sqlStr = sqlStr & " 		THEN convert(int,i.shopitemprice*(100-IsNULL(sd.defaultsuplymargin,100))/100) "
		sqlStr = sqlStr & " 	ELSE i.shopbuyprice END) as 'shopbuyprice'"
		sqlStr = sqlStr & " ,i.orgsellprice , sd.comm_cd"
		sqlStr = sqlStr & " ,i.centermwdiv,i.isusing ,si.itemgubun,si.itemoption "
		sqlStr = sqlStr & " ,(i.shopsuplycash) as 'shopsuplycash_org' , s.sale_name ,s.evt_code,s.shopid,s.sale_status"
		sqlStr = sqlStr & " ,c.comm_name"
		sqlStr = sqlStr & " FROM [db_shop].[dbo].[tbl_saleItem_off] si"
		sqlStr = sqlStr & " join [db_shop].[dbo].[tbl_sale_off] s"
		sqlStr = sqlStr & " 	on si.sale_code = s.sale_code"
		sqlStr = sqlStr & " join [db_shop].[dbo].[tbl_shop_item] i"
		sqlStr = sqlStr & " 	on si.itemid = i.shopitemid"
		sqlStr = sqlStr & " 	and si.itemgubun = i.itemgubun"
		sqlStr = sqlStr & " 	and si.itemoption = i.itemoption"
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item ii"
		sqlStr = sqlStr & " 	on si.itemid=ii.itemid"
		sqlStr = sqlStr & " 	and ii.itemgubun='10'"
		sqlStr = sqlStr & " left join [db_shop].dbo.tbl_shop_locale_item li"
		sqlStr = sqlStr & " 	on si.itemid=li.shopitemid"
		sqlStr = sqlStr & " 	and si.itemgubun = li.itemgubun"
		sqlStr = sqlStr & " 	and si.itemoption = li.itemoption"
		sqlStr = sqlStr & " 	and s.shopid=li.shopid"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_designer sd"
		sqlStr = sqlStr & " 	on s.shopid=sd.shopid"
		sqlStr = sqlStr & " 	and i.makerid=sd.makerid"
		sqlStr = sqlStr & " left join db_jungsan.dbo.tbl_jungsan_comm_code c"
		sqlStr = sqlStr & " 	on sd.comm_cd=c.comm_cd"
		sqlStr = sqlStr & " 	and c.comm_group='Z002'"
		sqlStr = sqlStr & " WHERE 1=1 " & strSearch
		sqlStr = sqlStr & " order by s.sale_code desc"

		''response.write sqlStr &"<br>"
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
				set FItemList(i) = new csale_oneitem

				FItemList(i).fsaleitem_idx = rsget("saleitem_idx")
				FItemList(i).fcomm_name = rsget("comm_name")
				FItemList(i).fpoint_rate = rsget("point_rate")
				FItemList(i).fpossaleprice = rsget("possaleprice")
				FItemList(i).fsale_status = rsget("sale_status")
				FItemList(i).fcomm_cd = rsget("comm_cd")
				FItemList(i).fsaleshopsupplycash = rsget("saleshopsupplycash")
				FItemList(i).fsale_name 		= db2html(rsget("sale_name"))
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fsale_code = rsget("sale_code")
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fsaleprice = rsget("saleprice")
				FItemList(i).fsalesupplycash = rsget("salesupplycash")
				FItemList(i).fsaleItem_status = rsget("saleItem_status")
				FItemList(i).flimitno = rsget("limitno")
				FItemList(i).forglimityn = rsget("orglimityn")
				FItemList(i).fmakerid = rsget("makerid")
				FItemList(i).fshopitemname = rsget("shopitemname")

				if rsget("smallimage") <> "" then
			    	IF rsget("itemgubun") = "90" THEN
			    		FItemList(i).fsmallimage = webImgUrl &"/offimage/offsmall/i" & rsget("itemgubun") & "/" & GetImageSubFolderByItemid(rsget("itemid")) &"/" &db2html(rsget("smallimage"))
			    	else
			    		FItemList(i).fsmallimage = webImgUrl &"/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) &"/" &db2html(rsget("smallimage"))
			    	END IF
				end if

				FItemList(i).fsaleyn = rsget("sailyn")
				FItemList(i).fshopitemprice = rsget("shopitemprice")
				FItemList(i).fshopsuplycash = rsget("shopsuplycash")
				FItemList(i).forgsellprice = rsget("orgsellprice")
				FItemList(i).fshopbuyprice = rsget("shopbuyprice")
				FItemList(i).fcentermwdiv = rsget("centermwdiv")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fshopsuplycash_org = rsget("shopsuplycash_org")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
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
End Class
%>
