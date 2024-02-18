<%
'#############################################################
'	Description : 클리어런스 세일 통계 클래스
'	History		: 2016.04.27 한용민 생성
'#############################################################

class cStaticclearancesale_oneitem
	public fitemid
	public fitemoption
	public fmakerid
	public fitemname
	public foptionname
	public fsellcash
	public fbuycash
	public fregdate
	public fsmallimage
	public fitemcostsum
	public fbuycashsum
	public frealstock_curryyyy
	public favailsysstock_curryyyy
	public frealstock_beforeyyyy
	public favailsysstock_beforeyyyy
	public frealstock
	public favailsysstock
	public FimageSmall
	public fitemno

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class cStaticclearancesale
	public foneitem
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public ftendb
	public FRectSort
	public FRectCateL
	public FRectCateM
	public FRectCateS
	public FRectIsBanPum
	public FRectPurchasetype
	public FRectDateGijun
	public FRectStartdate
	public FRectEndDate
	public FRectSiteName
	public FRectSellChannelDiv
	public FRectMwDiv
	public FRectMakerid
	public FRectInc3pl
	public FRectDispCate
	public FRectItemid
	public FRectChkShowGubun

	'//admin/maechul/statistic/statistic_clearancesale.asp
	public function fclearancesale_Statistic
		dim i, sql, sqlAdd

		'//////// 클리어런스 등록 이전 재고(월단위) ///////////////////////////////////////
		sql = "select"
		sql = sql & " s.itemid, s.itemoption, sum(sellno) as sellno, sum(resellno) as resellno, sum(realstock) as realstock, sum(availsysstock) as availsysstock"
		sql = sql & " into #boforeclearancestocktable_yyyy"
		sql = sql & " from "& ftendb &"[db_summary].[dbo].[tbl_monthly_logisstock_summary] s"
		sql = sql & " join "& ftendb &"db_sitemaster.dbo.tbl_clearance_sale_item c"
		sql = sql & " 	on s.itemid=c.itemid"
		sql = sql & " 	and c.isusing='Y'"
		sql = sql & " where s.itemgubun='10'"
		sql = sql & " and convert(varchar(7),s.yyyymm,121) < convert(varchar(7),c.regdate,121)"
		sql = sql & " group by s.itemid, s.itemoption"
		
		'response.write sql & "<br>"
		db3_dbget.Execute sql
		'//////// 클리어런스 등록 이전 재고(월단위) ///////////////////////////////////////

		'//////// 클리어런스 등록 이전 재고(일단위) ///////////////////////////////////////
		sql = "select"
		sql = sql & " s.itemid, s.itemoption, sum(sellno) as sellno, sum(resellno) as resellno, sum(realstock) as realstock, sum(availsysstock) as availsysstock"
		sql = sql & " into #boforeclearancestocktable"
		sql = sql & " from "& ftendb &"[db_summary].[dbo].[tbl_daily_logisstock_summary] s"
		sql = sql & " join "& ftendb &"db_sitemaster.dbo.tbl_clearance_sale_item c"
		sql = sql & " 	on s.itemid=c.itemid"
		sql = sql & " 	and c.isusing='Y'"
		sql = sql & " where s.itemgubun='10'"
		sql = sql & " and s.yyyymmdd < c.regdate"
		sql = sql & " and convert(varchar(7),dateadd(m,-1,getdate()),121) <= convert(varchar(7),s.yyyymmdd,121)"
		sql = sql & " group by s.itemid, s.itemoption"

		'response.write sql & "<br>"
		db3_dbget.Execute sql
		'//////// 클리어런스 등록 이전 재고(일단위) ///////////////////////////////////////

		'//////// 매출 ///////////////////////////////////////
		sql = "select"
		sql = sql & " d.itemid, d.itemoption"
		sql = sql & " , sum(d.itemno) as itemno, sum(d.itemcost*d.itemno) as itemcostsum, sum(d.buycash*d.itemno) as buycashsum"
		sql = sql & " into #maechultable"
		sql = sql & " from "& ftendb &"db_order.dbo.tbl_order_master m"
		sql = sql & " join "& ftendb &"db_order.dbo.tbl_order_detail d"
		sql = sql & " 	on m.orderserial=d.orderserial"
		sql = sql & " left join "& ftendb &"db_partner.dbo.tbl_partner p2"
		sql = sql & "	on m.sitename=p2.id"
		sql = sql & " where d.itemid not in (0,100)"
		sql = sql & " and m.ipkumdiv>3"
		sql = sql & " and m.cancelyn='N'"
		sql = sql & " and d.cancelyn<>'Y'"
		'sql = sql & " and m.jumundiv not in (6,9)"

		If FRectSiteName <> "" Then
		    if (FRectSiteName="mobileAll") then
		        sql = sql & " AND left(m.rdsite,6)='mobile'"
		    else
			    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
		    end if
		End If
		if (FRectDateGijun="beasongdate") then
		    sql = sql & "	and d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		else
	    	sql = sql & "	and m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	    end if

	    if (FRectInc3pl<>"") then
	        if (FRectInc3pl="A") then
	        else
	            sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
	        end if
	    else
	        sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
	    end if
		if (FRectSellChannelDiv<>"") then
	    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
	    end if
		If FRectIsBanPum <> "all" Then
			sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
		End If
		if (FRectMwDiv<>"") then
	        sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
	    end if

		sql = sql & " group by d.itemid, d.itemoption"

		'response.write sql & "<br>"
		db3_dbget.Execute sql
		'//////// 매출 ///////////////////////////////////////

		If FRectCateL <> "" Then
			sqlAdd = sqlAdd & " AND i.cate_large = '" & FRectCateL & "' "
		End If
		If FRectCateM <> "" Then
			sqlAdd = sqlAdd & " AND i.cate_mid = '" & FRectCateM & "' "
		End If
		If FRectCateS <> "" Then
			sqlAdd = sqlAdd & " AND i.cate_small = '" & FRectCateS & "' "
		End If
		If FRectMakerid <> "" Then
		    sqlAdd = sqlAdd & " and i.makerid = '" & FRectMakerid &"'"
		end if
		IF FRectItemid <> "" Then
			sqlAdd = sqlAdd & " and c.itemid in ("& FRectItemID&")"
		END IF
		If FRectPurchasetype <> "" Then
			sqlAdd = sqlAdd & " and p.purchasetype = '" & FRectPurchasetype &"'"
		End IF

		sql = "select"
		sql = sql & " count(c.itemid) as cnt"
		sql = sql & " from "& ftendb &"db_sitemaster.dbo.tbl_clearance_sale_item c"
		sql = sql & " join "& ftendb &"db_item.dbo.tbl_item i"
		sql = sql & " 	on c.itemid=i.itemid"
		sql = sql & " join "& ftendb &"db_item.dbo.tbl_item_option o"
		sql = sql & " 	on c.itemid=o.itemid"

		IF FRectDispCate<>"" THEN
			sql = sql & " INNER JOIN "& ftendb &"db_item.dbo.tbl_display_cate_item as dc"
			sql = sql & "  on c.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF

		sql = sql & " LEFT JOIN "& ftendb &"db_partner.dbo.tbl_partner as p"
		sql = sql & " 	on i.makerid = p.id"
		sql = sql & " left join #maechultable as t1"
		sql = sql & " 	on o.itemid=t1.itemid"
		sql = sql & " 	and o.itemoption=t1.itemoption"
		sql = sql & " left join #boforeclearancestocktable as t2"
		sql = sql & " 	on o.itemid=t2.itemid"
		sql = sql & " 	and o.itemoption=t2.itemoption"
		sql = sql & " left join #boforeclearancestocktable_yyyy as t3"
		sql = sql & " 	on o.itemid=t3.itemid"
		sql = sql & " 	and o.itemoption=t3.itemoption"
		sql = sql & " where c.isusing='Y' " & sqlAdd

		'response.write sql & "<br>"
		db3_rsget.Open sql,db3_dbget,1
			FTotalCount = db3_rsget("cnt")
		db3_rsget.Close

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " c.itemid, o.itemoption, i.makerid, i.itemname, o.optionname, i.sellcash, i.buycash, convert(varchar(10),c.regdate,121) as regdate"
		sql = sql & " , i.smallimage"
		sql = sql & " , isnull(t1.itemno,0) as itemno, isnull(t1.itemcostsum,0) as itemcostsum, isnull(t1.buycashsum,0) as buycashsum"
		sql = sql & " , isnull(t2.realstock,0) as realstock_curryyyy, isnull(t2.availsysstock,0) as availsysstock_curryyyy"
		sql = sql & " , isnull(t3.realstock,0) as realstock_beforeyyyy, isnull(t3.availsysstock,0) as availsysstock_beforeyyyy"
		sql = sql & " , isnull(t2.realstock,0)+isnull(t3.realstock,0) as realstock, isnull(t2.availsysstock,0)+isnull(t3.availsysstock,0) as availsysstock"
		sql = sql & " from "& ftendb &"db_sitemaster.dbo.tbl_clearance_sale_item c"
		sql = sql & " join "& ftendb &"db_item.dbo.tbl_item i"
		sql = sql & " 	on c.itemid=i.itemid"
		sql = sql & " join "& ftendb &"db_item.dbo.tbl_item_option o"
		sql = sql & " 	on c.itemid=o.itemid"

		IF FRectDispCate<>"" THEN
			sql = sql & " INNER JOIN "& ftendb &"db_item.dbo.tbl_display_cate_item as dc"
			sql = sql & "  on c.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF

		sql = sql & " LEFT JOIN "& ftendb &"db_partner.dbo.tbl_partner as p"
		sql = sql & " 	on i.makerid = p.id"
		sql = sql & " left join #maechultable as t1"
		sql = sql & " 	on o.itemid=t1.itemid"
		sql = sql & " 	and o.itemoption=t1.itemoption"
		sql = sql & " left join #boforeclearancestocktable as t2"
		sql = sql & " 	on o.itemid=t2.itemid"
		sql = sql & " 	and o.itemoption=t2.itemoption"
		sql = sql & " left join #boforeclearancestocktable_yyyy as t3"
		sql = sql & " 	on o.itemid=t3.itemid"
		sql = sql & " 	and o.itemoption=t3.itemoption"
		sql = sql & " where c.isusing='Y' " & sqlAdd
		sql = sql & " order by "

		IF FRectSort = "itemno" Then
			sql = sql & " itemno desc"
		elseIF FRectSort = "itemcost" Then
			sql = sql & " itemcostsum desc"
		elseIF FRectSort = "stock" Then
			sql = sql & " availsysstock desc"
		else
			sql = sql & " availsysstock desc"
		End If
		
		'response.write sql &"<br>"
		db3_rsget.pagesize = FPageSize		
		db3_rsget.Open sql,db3_dbget,1

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
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.EOF
				set FItemList(i) = new cStaticclearancesale_oneitem
					FItemList(i).fitemid = db3_rsget("itemid")
					FItemList(i).fitemoption = db3_rsget("itemoption")
					FItemList(i).fmakerid = db3_rsget("makerid")
					FItemList(i).fitemname = db2html(db3_rsget("itemname"))
					FItemList(i).foptionname = db2html(db3_rsget("optionname"))
					FItemList(i).fsellcash = db3_rsget("sellcash")
					FItemList(i).fbuycash = db3_rsget("buycash")
					FItemList(i).fregdate = db3_rsget("regdate")
					FItemList(i).fsmallimage = db3_rsget("smallimage")
					FItemList(i).fitemno = db3_rsget("itemno")
					FItemList(i).fitemcostsum = db3_rsget("itemcostsum")
					FItemList(i).fbuycashsum = db3_rsget("buycashsum")
					FItemList(i).frealstock_curryyyy = db3_rsget("realstock_curryyyy")
					FItemList(i).favailsysstock_curryyyy = db3_rsget("availsysstock_curryyyy")
					FItemList(i).frealstock_beforeyyyy = db3_rsget("realstock_beforeyyyy")
					FItemList(i).favailsysstock_beforeyyyy = db3_rsget("availsysstock_beforeyyyy")
					FItemList(i).frealstock = db3_rsget("realstock")
					FItemList(i).favailsysstock = db3_rsget("availsysstock")
					FItemList(i).FimageSmall = db3_rsget("smallimage")
																													
					if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	end function

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		IF application("Svr_Info")="Dev" THEN
			ftendb="tendb."
		end if
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