<%
'####################################################
' Description : 할인 통계 클래스
' History : 2012.10.24 한용민 생성
'####################################################

Class csalereport_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fsale_code
	public fsale_name
	public fsale_rate
	public fpoint_rate
	public fsale_margin
	public fsale_marginvalue
	public fsale_shopmargin
	public fsale_shopmarginvalue
	public fsale_startdate
	public fsale_enddate
	public fsale_status
	public fopendate
	public fsale_using
	public fadminid
	public fclosedate
	public fshopid
	public fshopname
	public ftotsellCnt
	public ftotselljumuncnt
	public ftotsellprice
	public ftotsuplyprice
	public ftotbuyprice
	public fsaleitem_cnt
	public ftotrealsellprice
	public fyyyymmdd
	public ftotitemno
	public fsellsum
	public fsum_cnt
	public fitemgubun
	public fitemoption
	public fitemname
	public fitemoptionname
	public FMakerid
	public fitemid
end Class

class Csalereport_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public fshopid
	public FSearchTxt
	public FSearchType
	public FBrand		
	public frectdatefg   
	public frectevt_startdate		
	public frectevt_enddate		
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
	public FRectsale_code
	public FRectInc3pl

	'/할인 통계테이블 에서 가져옴
	'//admin/offshop/sale/sale_report_detail.asp
	Public Sub getsalebrand_sum
		dim sqlStr , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectsale_code <> "" then
			sqlsearch = sqlsearch & " and s.sale_code ='" & frectsale_code & "'"
		end if
		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and s.Shopid ='" & FRectShopID & "'"
		end if

		IF frectevt_startdate <> "" AND frectevt_enddate <> "" THEN
			sqlsearch  = sqlsearch & " and sd.yyyymmdd >= '"&frectevt_startdate&"'"
			sqlsearch  = sqlsearch & " and sd.yyyymmdd < '"&frectevt_enddate&"'"
		END IF
		
		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.Shopid, s.sale_code, s.sale_name"
		sqlStr = sqlStr & " ,sd.makerid"
		sqlStr = sqlStr & " ,u.shopname"
		sqlStr = sqlStr & " ,isnull(sum(sd.totsellprice),0) as totsellprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totrealsellprice),0) as totrealsellprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totsuplyprice),0) as totsuplyprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totbuyprice),0) as totbuyprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totitemno),0) as totitemno"
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_sale_off s"		
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_sale_off_sellitem_summary_daily sd"
		sqlStr = sqlStr & " 	on s.sale_code=sd.sale_code"
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr & " 	on s.Shopid = u.userid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on s.shopid=p.id "
		sqlStr = sqlStr & " where s.sale_using=1"
		sqlStr = sqlStr & " and s.sale_status>=6 " & sqlsearch
		sqlStr = sqlStr & " group by s.Shopid,s.sale_code,s.sale_name, sd.makerid,u.shopname"
		sqlStr = sqlStr & " order by totsellprice desc"
			
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if not rsget.eof then
			do until rsget.eof
				set FItemList(i) = new csalereport_item
				
				FItemList(i).fShopname 	= db2html(rsget("Shopname"))
				FItemList(i).fShopid 	= rsget("Shopid")
				FItemList(i).fsale_code 	= rsget("sale_code")
				FItemList(i).fsale_name 	= db2html(rsget("sale_name"))
				FItemList(i).fmakerid 	= rsget("makerid")
				FItemList(i).ftotsellprice 	= rsget("totsellprice")
				FItemList(i).ftotrealsellprice 	= rsget("totrealsellprice")
				FItemList(i).ftotsuplyprice 	= rsget("totsuplyprice")
				FItemList(i).ftotbuyprice 	= rsget("totbuyprice")
				FItemList(i).ftotitemno 	= rsget("totitemno")
				
			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close
	end Sub

	'/할인 통계테이블 에서 가져옴
	'//admin/offshop/sale/sale_report_detail.asp
	Public Sub getsaleitem_sum
		dim sqlStr , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectsale_code <> "" then
			sqlsearch = sqlsearch & " and s.sale_code ='" & frectsale_code & "'"
		end if
		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and s.Shopid ='" & FRectShopID & "'"
		end if

		IF frectevt_startdate <> "" AND frectevt_enddate <> "" THEN
			sqlsearch  = sqlsearch & " and sd.yyyymmdd >= '"&frectevt_startdate&"'"
			sqlsearch  = sqlsearch & " and sd.yyyymmdd < '"&frectevt_enddate&"'"
		END IF
		
		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.Shopid, s.sale_code, s.sale_name"
		sqlStr = sqlStr & " ,sd.itemgubun, sd.itemid, sd.itemoption, sd.itemname, sd.itemoptionname, sd.makerid"
		sqlStr = sqlStr & " ,u.shopname"
		sqlStr = sqlStr & " ,isnull(sum(sd.totsellprice),0) as totsellprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totrealsellprice),0) as totrealsellprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totsuplyprice),0) as totsuplyprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totbuyprice),0) as totbuyprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totitemno),0) as totitemno"
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_sale_off s"		
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_sale_off_sellitem_summary_daily sd"
		sqlStr = sqlStr & " 	on s.sale_code=sd.sale_code"
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr & " 	on s.Shopid = u.userid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on s.shopid=p.id "
		sqlStr = sqlStr & " where s.sale_using=1"
		sqlStr = sqlStr & " and s.sale_status>=6 " & sqlsearch
		sqlStr = sqlStr & " group by s.Shopid,s.sale_code,s.sale_name,sd.itemgubun ,sd.itemid ,sd.itemoption"
		sqlStr = sqlStr & " 	,sd.itemname,sd.itemoptionname, sd.makerid,u.shopname"
		sqlStr = sqlStr & " order by totsellprice desc"
			
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if not rsget.eof then
			do until rsget.eof
				set FItemList(i) = new csalereport_item
				
				FItemList(i).fShopname 	= db2html(rsget("Shopname"))
				FItemList(i).fShopid 	= rsget("Shopid")
				FItemList(i).fsale_code 	= rsget("sale_code")
				FItemList(i).fsale_name 	= db2html(rsget("sale_name"))
				FItemList(i).fitemgubun 	= rsget("itemgubun")
				FItemList(i).fitemid 	= rsget("itemid")
				FItemList(i).fitemoption 	= rsget("itemoption")
				FItemList(i).fitemname 	= db2html(rsget("itemname"))
				FItemList(i).fitemoptionname 	= db2html(rsget("itemoptionname"))
				FItemList(i).fmakerid 	= rsget("makerid")
				FItemList(i).ftotsellprice 	= rsget("totsellprice")
				FItemList(i).ftotrealsellprice 	= rsget("totrealsellprice")
				FItemList(i).ftotsuplyprice 	= rsget("totsuplyprice")
				FItemList(i).ftotbuyprice 	= rsget("totbuyprice")
				FItemList(i).ftotitemno 	= rsget("totitemno")
				
			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close
	end Sub
	
	'/할인 통계테이블 에서 가져옴
	'//admin/offshop/sale/sale_report_detail.asp
    Public Sub getsaledate_sum()
		dim sqlStr,i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectsale_code <> "" then
			sqlsearch = sqlsearch & " and s.sale_code ='" & frectsale_code & "'"
		end if
		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and s.Shopid ='" & FRectShopID & "'"
		end if

		IF frectevt_startdate <> "" AND frectevt_enddate <> "" THEN
			sqlsearch  = sqlsearch & " and sd.yyyymmdd >= '"&frectevt_startdate&"'"
			sqlsearch  = sqlsearch & " and sd.yyyymmdd < '"&frectevt_enddate&"'"
		END IF

		'데이터 리스트 
		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.sale_name, s.sale_code, s.Shopid"		
		sqlStr = sqlStr & " ,isnull(sd.totsellCnt,0) as totsellCnt"
		sqlStr = sqlStr & " ,isnull(sd.totselljumuncnt,0) as totselljumuncnt"
		sqlStr = sqlStr & " ,isnull(sd.totsellprice,0) as totsellprice"
		sqlStr = sqlStr & " ,isnull(sd.totrealsellprice,0) as totrealsellprice"		
		sqlStr = sqlStr & " ,isnull(sd.totsuplyprice,0) as totsuplyprice"
		sqlStr = sqlStr & " ,isnull(sd.totbuyprice,0) as totbuyprice"
		sqlStr = sqlStr & " ,sd.yyyymmdd, u.shopname"
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_sale_off s"		
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_sale_off_sell_summary_daily sd"
		sqlStr = sqlStr & " 	on s.sale_code=sd.sale_code"
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr & " 	on s.Shopid = u.userid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on s.shopid=p.id "		
		sqlStr = sqlStr & " where s.sale_using=1"
		sqlStr = sqlStr & " and s.sale_status>=6 " & sqlsearch
		sqlStr = sqlStr & " order by s.Shopid asc, sd.yyyymmdd desc"
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			
			do until rsget.EOF
				set FItemList(i) = new csalereport_item

				FItemList(i).fsale_name = db2html(rsget("sale_name"))
				FItemList(i).fsale_code = rsget("sale_code")
				FItemList(i).fShopid = rsget("Shopid")
				FItemList(i).fyyyymmdd = rsget("yyyymmdd")
				FItemList(i).fshopname = db2html(rsget("shopname"))
				FItemList(i).ftotsellCnt = rsget("totsellCnt")
				FItemList(i).ftotselljumuncnt = rsget("totselljumuncnt")
				FItemList(i).ftotsellprice = rsget("totsellprice")
				FItemList(i).ftotrealsellprice = rsget("totrealsellprice")			
				FItemList(i).ftotsuplyprice = rsget("totsuplyprice")
				FItemList(i).ftotbuyprice = rsget("totbuyprice")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'/할인 통계테이블 에서 가져옴
	'//admin/offshop/sale/sale_report.asp
    Public Sub getsale_sum()
        dim sqlStr,i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		IF frectshopid <> "" THEN
			sqlsearch  = sqlsearch & " and s.shopid = '"&frectshopid&"'"
		END IF

		IF FSearchTxt <> "" THEN
			IF FSearchType = 1 THEN 
				sqlsearch = sqlsearch & " and s.sale_code = "&FSearchTxt
			ELSEIF FSearchType= 2 THEN
				sqlsearch = sqlsearch & " and s.evt_code = "&FSearchTxt	
			ELSEIF FSearchType=3 THEN
				sqlsearch = sqlsearch & " and s.sale_name like '%"& FSearchTxt &"%' "
			END IF	
		END IF	

		IF frectevt_startdate <> "" AND frectevt_enddate <> "" THEN
			if CStr(frectdatefg) = "S" THEN

				sqlsearch  = sqlsearch & " and s.sale_startdate >= '"&frectevt_startdate&"'"
				sqlsearch  = sqlsearch & " and s.sale_startdate < '"&frectevt_enddate&"'"

			elseif CStr(frectdatefg) = "E" THEN
				sqlsearch  = sqlsearch & " and s.sale_enddate >= '"&frectevt_startdate&"'"
				sqlsearch  = sqlsearch & " and s.sale_enddate < '"&frectevt_enddate&"'"				

			end if
		END IF

		IF FSStatus <> "" THEN			
			sqlsearch = sqlsearch & " and s.sale_status = "&FSStatus
		END IF			
		
		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " s.sale_code, s.sale_name, s.sale_rate, s.point_rate, s.sale_margin, s.sale_marginvalue"
		sqlStr = sqlStr & " , s.sale_shopmargin, s.sale_shopmarginvalue, s.sale_startdate, s.sale_enddate"
		sqlStr = sqlStr & " , s.sale_status, s.opendate, s.sale_using, s.adminid, s.closedate, s.shopid"
		sqlStr = sqlStr & " , (select count(*) from [db_shop].dbo.tbl_saleItem_off  where sale_code = s.sale_code ) as saleitem_cnt"
		sqlStr = sqlStr & " ,u.shopname"
		sqlStr = sqlStr & " ,isnull(sum(sd.totsellCnt),0) as totsellCnt"
		sqlStr = sqlStr & " ,isnull(sum(sd.totselljumuncnt),0) as totselljumuncnt"
		sqlStr = sqlStr & " ,isnull(sum(sd.totsellprice),0) as totsellprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totrealsellprice),0) as totrealsellprice"		
		sqlStr = sqlStr & " ,isnull(sum(sd.totsuplyprice),0) as totsuplyprice"
		sqlStr = sqlStr & " ,isnull(sum(sd.totbuyprice),0) as totbuyprice"
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_sale_off s"		
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_sale_off_sell_summary_daily sd"
		sqlStr = sqlStr & " 	on s.sale_code=sd.sale_code"
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr & " 	on s.Shopid = u.userid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on s.Shopid=p.id "
		sqlStr = sqlStr & " where s.sale_using=1"
		sqlStr = sqlStr & " and s.sale_status>=6 " & sqlsearch
		sqlStr = sqlStr & " GROUP BY s.sale_code, s.sale_name, s.sale_rate, s.point_rate, s.sale_margin, s.sale_marginvalue"
		sqlStr = sqlStr & " , s.sale_shopmargin, s.sale_shopmarginvalue, s.sale_startdate, s.sale_enddate"
		sqlStr = sqlStr & " , s.sale_status, s.opendate, s.sale_using, s.adminid, s.closedate, s.shopid"
		sqlStr = sqlStr & " ,u.shopname"
		sqlStr = sqlStr & " order by s.sale_code desc" 

        'response.write sqlStr &"<Br>"
        rsget.open sqlStr,dbget

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new csalereport_item
			
			FItemList(i).fsaleitem_cnt = rsget("saleitem_cnt")
			FItemList(i).fsale_code = rsget("sale_code")
			FItemList(i).fsale_name = db2html(rsget("sale_name"))
			FItemList(i).fsale_rate = rsget("sale_rate")
			FItemList(i).fpoint_rate = rsget("point_rate")
			FItemList(i).fsale_margin = rsget("sale_margin")
			FItemList(i).fsale_marginvalue = rsget("sale_marginvalue")
			FItemList(i).fsale_shopmargin = rsget("sale_shopmargin")
			FItemList(i).fsale_shopmarginvalue = rsget("sale_shopmarginvalue")
			FItemList(i).fsale_startdate = left(rsget("sale_startdate"),10)
			FItemList(i).fsale_enddate = left(rsget("sale_enddate"),10)
			FItemList(i).fsale_status = rsget("sale_status")
			FItemList(i).fopendate = rsget("opendate")
			FItemList(i).fsale_using = rsget("sale_using")
			FItemList(i).fadminid = rsget("adminid")
			FItemList(i).fclosedate = rsget("closedate")
			FItemList(i).fshopid = rsget("shopid")
			FItemList(i).fshopname = db2html(rsget("shopname"))
			FItemList(i).ftotsellCnt = rsget("totsellCnt")
			FItemList(i).ftotselljumuncnt = rsget("totselljumuncnt")
			FItemList(i).ftotsellprice = rsget("totsellprice")
			FItemList(i).ftotrealsellprice = rsget("totrealsellprice")			
			FItemList(i).ftotsuplyprice = rsget("totsuplyprice")
			FItemList(i).ftotbuyprice = rsget("totbuyprice")			
			
		rsget.MoveNext
		i = i + 1
		loop

		rsget.close
    End Sub
    
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