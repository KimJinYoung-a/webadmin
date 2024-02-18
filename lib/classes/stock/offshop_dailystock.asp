<%
class COffShopRealJaegoMaster
	public Fidx
	public Fshopid
	public Fmakerid
	public Fjeagodate
	public Fregdate
	public Fcancelyn

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class COffShopDailyStockItem
	public Fshopid
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fmakerid
	public Fitemname
	public Fitemoptionname
	public Flastrealdate
	public Flastrealno
	public Fipno
	public Freno
	public Fupcheipno
	public Fupchereno
	public Fsellno
	public Fcurrno
	public Fimgsmall
	public Fregdate
	public Fsell7days
	public Fpreorderno
	public Frequireno
	public Fshortageno
	public Fmaxsellday

	public Fonlinesellcash
	public Fshopitemprice

	public FinputedRealStock

	public Fshopname
	public Fchargediv			'// 사용안함
	public Fcomm_cd
	public Fdefaultmargin
	public Fdefaultsuplymargin

	public function getChargeDivColor()
		if FChargeDiv="2" then
			getChargeDivColor = "#FF4444"
		elseif FChargeDiv="4" then
			getChargeDivColor = "#44FF44"
		elseif FChargeDiv="6" then
			getChargeDivColor = "#4444FF"
		elseif FChargeDiv="8" then
			getChargeDivColor = "#FF44FF"
		end if
	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "텐위"
		elseif FChargeDiv="4" then
			getChargeDivName = "텐매"
		elseif FChargeDiv="6" then
			getChargeDivName = "업위"
		elseif FChargeDiv="8" then
			getChargeDivName = "업매"
		else
			getChargeDivName = FChargeDiv
		end if
	end function

	public function getCommCdColor()
		if Fcomm_cd="B011" then
			getCommCdColor = "#FF4444"
		elseif Fcomm_cd="B031" then
			getCommCdColor = "#44FF44"
		elseif Fcomm_cd="B012" then
			getCommCdColor = "#4444FF"
		elseif Fcomm_cd="B013" then
			getCommCdColor = "#FF44FF"
		end if
	end function

	public function getCommCdName()
		if Fcomm_cd="B011" then
			getCommCdName = "텐위"
		elseif Fcomm_cd="B031" then
			getCommCdName = "매입출고"
		elseif Fcomm_cd="B012" then
			getCommCdName = "업위"
		elseif Fcomm_cd="B013" then
			getCommCdName = "출고위탁"
		else
			getCommCdName = Fcomm_cd
		end if
	end function

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (FItemID >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + Format00(8,FItemId) + FItemOption
    	end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class COffShopDailyStock
	public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectIdx
	public FRectMakerid
	public FRectShopID
	public FRectMinusNo
	public FRecAvailStock
	public FRecOnlyusing

	public FRectItemGubun
	public FRectItemid
	public FRectItemoption

	public function GetCurrentAllShopItemStock()
		dim sqlStr, i
		sqlStr = "select top 100 u.userid, u.shopname,"
		sqlStr = sqlStr + " s.lastrealdate, s.lastrealno,s.ipno, s.reno, s.upcheipno, "
		sqlStr = sqlStr + " s.upchereno, s.sellno, s.currno, s.regdate,"
		sqlStr = sqlStr + " d.chargediv, IsNULL(d.defaultmargin,0) as defaultmargin, IsNULL(d.defaultsuplymargin,0) as defaultsuplymargin, d.comm_cd"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " on u.userid=d.shopid"
		sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_day_stock s"
		sqlStr = sqlStr + " on s.shopid=d.shopid"
		sqlStr = sqlStr + " and s.itemgubun='" + FRectItemGubun + "'"
		sqlStr = sqlStr + " and s.itemid=" + FRectItemid + ""
		sqlStr = sqlStr + " and s.itemoption='" + FRectItemoption + "'"
		sqlStr = sqlStr + " order by u.userid"
		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDailyStockItem
				FItemList(i).Fshopid          = rsget("userid")
				FItemList(i).Flastrealdate    = rsget("lastrealdate")
				FItemList(i).Flastrealno      = rsget("lastrealno")
				FItemList(i).Fipno            = rsget("ipno")
				FItemList(i).Freno            = rsget("reno")
				FItemList(i).Fupcheipno       = rsget("upcheipno")
				FItemList(i).Fupchereno       = rsget("upchereno")
				FItemList(i).Fsellno          = rsget("sellno")
				FItemList(i).Fcurrno          = rsget("currno")
				FItemList(i).Fregdate         = rsget("regdate")

				FItemList(i).Fshopname			= db2html(rsget("shopname"))
				FItemList(i).Fchargediv			 = rsget("chargediv")					'// 사용안함
				FItemList(i).Fcomm_cd			 = rsget("comm_cd")
				FItemList(i).Fdefaultmargin		 = rsget("defaultmargin")
				FItemList(i).Fdefaultsuplymargin = rsget("defaultsuplymargin")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public function GetCurrentAllShopItemStockNEW()
		dim sqlStr, i

		sqlStr = " select u.userid, u.shopname, (c.realstockno + c.errsampleitemno) as realstockno "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_summary].[dbo].[tbl_current_shopstock_summary] c "
		sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shop_user u "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.shopid = u.userid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and u.shopdiv in ('1', '2') "
		sqlStr = sqlStr + " 	and u.isusing = 'Y' "
		sqlStr = sqlStr + " 	and c.itemgubun='" + FRectItemGubun + "'"
		sqlStr = sqlStr + " 	and c.itemid=" + FRectItemid + ""
		sqlStr = sqlStr + " 	and c.itemoption='" + FRectItemoption + "'"
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	u.userid "
		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDailyStockItem
				FItemList(i).Fshopid          = rsget("userid")
				FItemList(i).Fshopname		  = db2html(rsget("shopname"))
				FItemList(i).Fcurrno          = rsget("realstockno")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public function GetCurrentSysStock()
		dim sqlStr, i
		'sqlStr = "select top 1000 s.shopid,i.itemgubun,i.shopitemid,i.itemoption, "
		'sqlStr = sqlStr + " i.makerid, i.shopitemname, i.shopitemoptionname,"
		'sqlStr = sqlStr + " s.lastrealdate, s.lastrealno,s.ipno, s.reno, s.upcheipno, s.upchereno, s.sellno, s.currno,"
		'sqlStr = sqlStr + " i.offimgsmall, s.regdate, o.smallimage, o.sellcash as onlinesellcash, i.shopitemprice"


		sqlStr = "select top 1000 s.itemgubun, s.shopitemid, s.itemoption," + VbCrlf
		sqlStr = sqlStr + " s.shopitemname, s.shopitemoptionname, s.shopitemprice," + VbCrlf
		sqlStr = sqlStr + " IsNULL(T1.sellno,0) + IsNULL(T2.sellno,0) as sellno," + VbCrlf
		sqlStr = sqlStr + " IsNULL(T3.ipno,0) as ipno," + VbCrlf
		sqlStr = sqlStr + " IsNULL(T4.reno,0) as reno," + VbCrlf
		sqlStr = sqlStr + " IsNULL(T5.upcheipno,0) as upcheipno," + VbCrlf
		sqlStr = sqlStr + " IsNULL(T6.upchereipno,0) as upchereno," + VbCrlf
		sqlStr = sqlStr + " IsNULL(T7.offjungsanno,0) as offjungsanno" + VbCrlf

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s" + VbCrlf
		sqlStr = sqlStr + " left join (" + VbCrlf
		sqlStr = sqlStr + " 	select d.itemgubun,d.itemid,d.itemoption,sum(d.itemno) as sellno" + VbCrlf
		sqlStr = sqlStr + " 	from " + VbCrlf
		sqlStr = sqlStr + " 	[db_shoplog].[dbo].tbl_old_shopjumun_master m," + VbCrlf
		sqlStr = sqlStr + " 	[db_shoplog].[dbo].tbl_old_shopjumun_detail d" + VbCrlf
		sqlStr = sqlStr + " 	where m.shopid='" + FRectShopid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and d.makerid='" + FRectMakerid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and m.cancelyn='N'" + VbCrlf
		sqlStr = sqlStr + " 	and m.idx=d.masteridx" + VbCrlf
		sqlStr = sqlStr + " 	and d.cancelyn='N'" + VbCrlf
		sqlStr = sqlStr + " 	group  by d.itemgubun,d.itemid,d.itemoption" + VbCrlf
		sqlStr = sqlStr + " ) T1 on s.itemgubun=T1.itemgubun and s.shopitemid=T1.itemid and s.itemoption=T1.itemoption" + VbCrlf
		sqlStr = sqlStr + " left join (" + VbCrlf
		sqlStr = sqlStr + " 	select  d.itemgubun,d.itemid,d.itemoption,sum(d.itemno) as sellno" + VbCrlf
		sqlStr = sqlStr + " 	from " + VbCrlf
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_master m," + VbCrlf
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d" + VbCrlf
		sqlStr = sqlStr + " 	where m.shopid='" + FRectShopid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and d.makerid='" + FRectMakerid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and m.cancelyn='N'" + VbCrlf
		sqlStr = sqlStr + " 	and m.idx=d.masteridx" + VbCrlf
		sqlStr = sqlStr + " 	and d.cancelyn='N'" + VbCrlf
		sqlStr = sqlStr + " 	group  by d.itemgubun,d.itemid,d.itemoption" + VbCrlf
		sqlStr = sqlStr + " ) T2 on s.itemgubun=T2.itemgubun and s.shopitemid=T2.itemid and s.itemoption=T2.itemoption" + VbCrlf
		sqlStr = sqlStr + " left join (" + VbCrlf
		sqlStr = sqlStr + " 	select d.iitemgubun,d.itemid,d.itemoption,sum(d.itemno*-1) as ipno" + VbCrlf
		sqlStr = sqlStr + " 	from " + VbCrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m," + VbCrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d" + VbCrlf
		sqlStr = sqlStr + " 	where m.ipchulflag='S'" + VbCrlf
		sqlStr = sqlStr + " 	and m.socid='" + FRectShopid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and m.deldt is null" + VbCrlf
		sqlStr = sqlStr + " 	and m.executedt is not null" + VbCrlf
		sqlStr = sqlStr + " 	and m.code=d.mastercode" + VbCrlf
		sqlStr = sqlStr + " 	and d.imakerid='" + FRectMakerid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and d.deldt is null" + VbCrlf
		sqlStr = sqlStr + " 	and d.itemno<0" + VbCrlf
		sqlStr = sqlStr + " 	group  by d.iitemgubun,d.itemid,d.itemoption" + VbCrlf
		sqlStr = sqlStr + " ) T3 on s.itemgubun=T3.iitemgubun and s.shopitemid=T3.itemid and s.itemoption=T3.itemoption" + VbCrlf
		sqlStr = sqlStr + " left join (" + VbCrlf
		sqlStr = sqlStr + " 	select d.iitemgubun,d.itemid,d.itemoption,sum(d.itemno) as reno" + VbCrlf
		sqlStr = sqlStr + " 	from " + VbCrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m," + VbCrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d" + VbCrlf
		sqlStr = sqlStr + " 	where m.ipchulflag='S'" + VbCrlf
		sqlStr = sqlStr + " 	and m.socid='" + FRectShopid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and m.deldt is null" + VbCrlf
		sqlStr = sqlStr + " 	and m.executedt is not null" + VbCrlf
		sqlStr = sqlStr + " 	and m.code=d.mastercode" + VbCrlf
		sqlStr = sqlStr + " 	and d.imakerid='" + FRectMakerid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and d.deldt is null" + VbCrlf
		sqlStr = sqlStr + " 	and d.itemno>0" + VbCrlf
		sqlStr = sqlStr + " 	group  by d.iitemgubun,d.itemid,d.itemoption" + VbCrlf
		sqlStr = sqlStr + " ) T4 on s.itemgubun=T4.iitemgubun and s.shopitemid=T4.itemid and s.itemoption=T4.itemoption" + VbCrlf
		sqlStr = sqlStr + " left join (" + VbCrlf
		sqlStr = sqlStr + " 	select d.itemgubun,d.shopitemid,d.itemoption,sum(d.itemno) as upcheipno" + VbCrlf
		sqlStr = sqlStr + " 	from " + VbCrlf
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_master m," + VbCrlf
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail d," + VbCrlf
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_item s" + VbCrlf
		sqlStr = sqlStr + " 	where m.chargeid<>'10x10'" + VbCrlf
		sqlStr = sqlStr + " 	and m.shopid='" + FRectShopid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and m.deleteyn='N'" + VbCrlf
		sqlStr = sqlStr + " 	and m.execdt is not null" + VbCrlf
		sqlStr = sqlStr + " 	and m.idx=d.masteridx" + VbCrlf
		sqlStr = sqlStr + " 	and d.itemgubun=s.itemgubun" + VbCrlf
		sqlStr = sqlStr + " 	and d.shopitemid=s.shopitemid" + VbCrlf
		sqlStr = sqlStr + " 	and d.itemoption=s.itemoption" + VbCrlf
		sqlStr = sqlStr + " 	and s.makerid='" + FRectMakerid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and d.deleteyn='N'" + VbCrlf
		sqlStr = sqlStr + " 	and d.itemno>0" + VbCrlf
		sqlStr = sqlStr + " 	group  by d.itemgubun,d.shopitemid,d.itemoption" + VbCrlf
		sqlStr = sqlStr + " ) T5 on s.itemgubun=T5.itemgubun and s.shopitemid=T5.shopitemid and s.itemoption=T5.itemoption" + VbCrlf
		sqlStr = sqlStr + " left join (" + VbCrlf
		sqlStr = sqlStr + " 	select d.itemgubun,d.shopitemid,d.itemoption,sum(d.itemno) as upchereipno" + VbCrlf
		sqlStr = sqlStr + " 	from " + VbCrlf
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_master m," + VbCrlf
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail d," + VbCrlf
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_item s" + VbCrlf
		sqlStr = sqlStr + " 	where m.chargeid<>'10x10'" + VbCrlf
		sqlStr = sqlStr + " 	and m.shopid='" + FRectShopid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and m.deleteyn='N'" + VbCrlf
		sqlStr = sqlStr + " 	and m.execdt is not null" + VbCrlf
		sqlStr = sqlStr + " 	and m.idx=d.masteridx" + VbCrlf
		sqlStr = sqlStr + " 	and d.itemgubun=s.itemgubun" + VbCrlf
		sqlStr = sqlStr + " 	and d.shopitemid=s.shopitemid" + VbCrlf
		sqlStr = sqlStr + " 	and d.itemoption=s.itemoption" + VbCrlf
		sqlStr = sqlStr + " 	and s.makerid='" + FRectMakerid + "'" + VbCrlf
		sqlStr = sqlStr + " 	and d.deleteyn='N'" + VbCrlf
		sqlStr = sqlStr + " 	and d.itemno<0" + VbCrlf
		sqlStr = sqlStr + " 	group  by d.itemgubun,d.shopitemid,d.itemoption" + VbCrlf
		sqlStr = sqlStr + " ) T6 on s.itemgubun=T6.itemgubun and s.shopitemid=T6.shopitemid and s.itemoption=T6.itemoption" + VbCrlf
		sqlStr = sqlStr + " left join (" + VbCrlf
		sqlStr = sqlStr + " 	 select d.itemgubun, d.itemid , d.itemoption , " + VbCrlf
		sqlStr = sqlStr + " 	 sum(itemno) as offjungsanno " + VbCrlf
		sqlStr = sqlStr + " 	 from [db_shop].[dbo].tbl_shop_jungsanmaster m, [db_shop].[dbo].tbl_shop_jungsandetail d " + VbCrlf
		sqlStr = sqlStr + " 	 where m.idx = d.masteridx " + VbCrlf
		sqlStr = sqlStr + " 	 and m.shopid='" + FRectShopid + "'" + VbCrlf
		sqlStr = sqlStr + " 	 and m.jungsanid='" + FRectMakerid + "'" + VbCrlf
		sqlStr = sqlStr + " 	 group by d.itemgubun, d.itemid, d.itemoption" + VbCrlf
		sqlStr = sqlStr + " ) T7 on  s.itemgubun=T7.itemgubun and s.shopitemid=T7.itemid and s.itemoption=T7.itemoption" + VbCrlf
		sqlStr = sqlStr + " where s.makerid='" + FRectMakerid + "'" + VbCrlf
		sqlStr = sqlStr + " order by s.itemgubun, s.shopitemid, s.itemoption" + VbCrlf


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDailyStockItem
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("shopitemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fitemname        = db2html(rsget("shopitemname"))
				FItemList(i).Fitemoptionname  = db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice   = rsget("shopitemprice")

				FItemList(i).Fipno            = rsget("ipno")
				FItemList(i).Freno           = rsget("reno")
				FItemList(i).Fupcheipno       = rsget("upcheipno")
				FItemList(i).Fupchereno       = rsget("upchereno")
				FItemList(i).Fsellno          = rsget("sellno")
				FItemList(i).Fcurrno          = FItemList(i).Fipno + FItemList(i).Freno + FItemList(i).Fupcheipno + FItemList(i).Fupchereno - FItemList(i).Fsellno

'				if FItemList(i).Fitemgubun="10" then
'					FItemList(i).Fimgsmall        = rsget("smallimage")
'				else
'					FItemList(i).Fimgsmall        = rsget("offimgsmall")
'				end if

'				FItemList(i).Fregdate        = rsget("regdate")

				if IsNULL(FItemList(i).Flastrealdate) then FItemList(i).Flastrealdate="1901-01-01"
				if IsNULL(FItemList(i).Flastrealno) then FItemList(i).Flastrealno=0
				if IsNULL(FItemList(i).Fipno) then FItemList(i).Fipno=0
				if IsNULL(FItemList(i).Freno) then FItemList(i).Freno=0
				if IsNULL(FItemList(i).Fupcheipno) then FItemList(i).Fupcheipno=0
				if IsNULL(FItemList(i).Fupchereno) then FItemList(i).Fupchereno=0
				if IsNULL(FItemList(i).Fsellno) then FItemList(i).Fsellno=0
				if IsNULL(FItemList(i).Fcurrno) then FItemList(i).Fcurrno=0

'				if IsNULL(FItemList(i).Fimgsmall) then

'				elseif	FItemList(i).Fitemgubun="10" then
'					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
'				else
'					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
'				end if

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

	public function GetDailyStock()
		dim sqlStr, i
		sqlStr = "select top 1000 s.shopid,i.itemgubun,i.shopitemid,i.itemoption, "
		sqlStr = sqlStr + " i.makerid, i.shopitemname, i.shopitemoptionname,"
		sqlStr = sqlStr + " s.lastrealdate, s.lastrealno,s.ipno, s.reno, s.upcheipno, s.upchereno, s.sellno, s.currno,"
		sqlStr = sqlStr + " i.offimgsmall, s.regdate, o.smallimage, o.sellcash as onlinesellcash, i.shopitemprice"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_day_stock s"
		sqlStr = sqlStr + " on shopid='" + FRectShopID + "' and i.itemgubun=s.itemgubun and i.shopitemid=s.itemid and i.itemoption=s.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item o"
		sqlStr = sqlStr + " on i.itemgubun='10' and i.shopitemid=o.itemid"
		sqlStr = sqlStr + " where i.makerid='" + FRectMakerid + "'"
		if FRecAvailStock<>"" then
			sqlStr = sqlStr + " and s.shopid is not null"
		end if

		if FRecOnlyusing<>"" then
			sqlStr = sqlStr + " and i.isusing='Y'"
		end if
		sqlStr = sqlStr + " order by i.itemgubun, i.shopitemid desc, i.itemoption"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDailyStockItem
				FItemList(i).Fshopid          = rsget("shopid")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("shopitemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemname        = db2html(rsget("shopitemname"))
				FItemList(i).Fitemoptionname  = db2html(rsget("shopitemoptionname"))
				FItemList(i).Flastrealdate    = rsget("lastrealdate")
				FItemList(i).Flastrealno      = rsget("lastrealno")
				FItemList(i).Fipno            = rsget("ipno")
				FItemList(i).Freno           = rsget("reno")
				FItemList(i).Fupcheipno       = rsget("upcheipno")
				FItemList(i).Fupchereno       = rsget("upchereno")
				FItemList(i).Fsellno          = rsget("sellno")
				FItemList(i).Fcurrno          = rsget("currno")

				FItemList(i).Fonlinesellcash  = rsget("onlinesellcash")
				FItemList(i).Fshopitemprice   = rsget("shopitemprice")

				if IsNULL(FItemList(i).Fonlinesellcash) then FItemList(i).Fonlinesellcash=0
				if IsNULL(FItemList(i).Fshopitemprice) then FItemList(i).Fshopitemprice=0

				if FItemList(i).Fitemgubun="10" then
					FItemList(i).Fimgsmall        = rsget("smallimage")
				else
					FItemList(i).Fimgsmall        = rsget("offimgsmall")
				end if

				FItemList(i).Fregdate        = rsget("regdate")

				if IsNULL(FItemList(i).Flastrealdate) then FItemList(i).Flastrealdate="1901-01-01"
				if IsNULL(FItemList(i).Flastrealno) then FItemList(i).Flastrealno=0
				if IsNULL(FItemList(i).Fipno) then FItemList(i).Fipno=0
				if IsNULL(FItemList(i).Freno) then FItemList(i).Freno=0
				if IsNULL(FItemList(i).Fupcheipno) then FItemList(i).Fupcheipno=0
				if IsNULL(FItemList(i).Fupchereno) then FItemList(i).Fupchereno=0
				if IsNULL(FItemList(i).Fsellno) then FItemList(i).Fsellno=0
				if IsNULL(FItemList(i).Fcurrno) then FItemList(i).Fcurrno=0

				if IsNULL(FItemList(i).Fimgsmall) then

				elseif	FItemList(i).Fitemgubun="10" then
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				else
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				end if

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function


	public function GetCurrentStockMinusList()
		dim sqlStr, i
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_day_stock s"
		sqlStr = sqlStr + " where shopid='" + FRectShopID + "'"
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and makerid='" + FRectMakerid + "'"
		end if
		sqlStr = sqlStr + " and s.currno<" + CStr(FRectMinusNo)

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " s.shopid,s.itemgubun,s.itemid,s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.itemname, s.itemoptionname,"
		sqlStr = sqlStr + " s.lastrealdate, s.lastrealno,s.ipno, s.reno, s.upcheipno, s.upchereno, s.sellno, s.currno,"
		sqlStr = sqlStr + " s.imgsmall, s.regdate, o.smallimage"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_day_stock s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item o"
		sqlStr = sqlStr + " on s.itemgubun='10' and s.itemid=o.itemid"
		sqlStr = sqlStr + " where s.shopid='" + FRectShopID + "'"
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectMakerid + "'"
		end if
		sqlStr = sqlStr + " and s.currno<" + CStr(FRectMinusNo)

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDailyStockItem
				FItemList(i).Fshopid          = rsget("shopid")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemname        = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
				FItemList(i).Flastrealdate    = rsget("lastrealdate")
				FItemList(i).Flastrealno      = rsget("lastrealno")
				FItemList(i).Fipno            = rsget("ipno")
				FItemList(i).Freno           = rsget("reno")
				FItemList(i).Fupcheipno       = rsget("upcheipno")
				FItemList(i).Fupchereno       = rsget("upchereno")
				FItemList(i).Fsellno          = rsget("sellno")
				FItemList(i).Fcurrno          = rsget("currno")

				if FItemList(i).Fitemgubun="10" then
					FItemList(i).Fimgsmall        = rsget("smallimage")
				else
					FItemList(i).Fimgsmall        = rsget("imgsmall")
				end if

				FItemList(i).Fregdate        = rsget("regdate")

				if IsNULL(FItemList(i).Flastrealdate) then FItemList(i).Flastrealdate="1901-01-01"
				if IsNULL(FItemList(i).Flastrealno) then FItemList(i).Flastrealno=0
				if IsNULL(FItemList(i).Fipno) then FItemList(i).Fipno=0
				if IsNULL(FItemList(i).Freno) then FItemList(i).Freno=0
				if IsNULL(FItemList(i).Fupcheipno) then FItemList(i).Fupcheipno=0
				if IsNULL(FItemList(i).Fupchereno) then FItemList(i).Fupchereno=0
				if IsNULL(FItemList(i).Fsellno) then FItemList(i).Fsellno=0
				if IsNULL(FItemList(i).Fcurrno) then FItemList(i).Fcurrno=0

				if IsNULL(FItemList(i).Fimgsmall) then

				elseif	FItemList(i).Fitemgubun="10" then
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				else
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				end if

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public sub GetDailyStockByInputIdx()
		dim sqlStr, i
		sqlStr = "select top 1000 s.shopid,i.itemgubun,i.shopitemid,i.itemoption, "
		sqlStr = sqlStr + " i.makerid, i.shopitemname, i.shopitemoptionname,"
		sqlStr = sqlStr + " s.lastrealdate, s.lastrealno,s.ipno, s.reno, s.upcheipno, s.upchereno, s.sellno, s.currno,"
		sqlStr = sqlStr + " i.offimgsmall, s.regdate, o.smallimage, r.realjeago"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_day_stock s"
		sqlStr = sqlStr + " on shopid='" + FRectShopID + "' and i.itemgubun=s.itemgubun and i.shopitemid=s.itemid and i.itemoption=s.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item o"
		sqlStr = sqlStr + " on i.itemgubun='10' and i.shopitemid=o.itemid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_realjaego_detail r"
		sqlStr = sqlStr + " on r.masteridx=" + CStr(FRectIdx) + " and i.itemgubun=r.itemgubun and i.shopitemid=r.shopitemid and i.itemoption=r.itemoption"

		sqlStr = sqlStr + " where i.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and r.realjeago is not null"

'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDailyStockItem
				FItemList(i).Fshopid          = rsget("shopid")
				FItemList(i).Fitemgubun       = rsget("itemgubun")
				FItemList(i).Fitemid          = rsget("shopitemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fmakerid         = rsget("makerid")
				FItemList(i).Fitemname        = db2html(rsget("shopitemname"))
				FItemList(i).Fitemoptionname  = db2html(rsget("shopitemoptionname"))
				'FItemList(i).Flastrealdate    = rsget("lastrealdate")
				'FItemList(i).Flastrealno      = rsget("lastrealno")
				'FItemList(i).Fipno            = rsget("ipno")
				'FItemList(i).Freno           = rsget("reno")
				'FItemList(i).Fupcheipno       = rsget("upcheipno")
				'FItemList(i).Fupchereno       = rsget("upchereno")
				'FItemList(i).Fsellno          = rsget("sellno")
				'FItemList(i).Fcurrno          = rsget("currno")

				if FItemList(i).Fitemgubun="10" then
					FItemList(i).Fimgsmall        = rsget("smallimage")
				else
					FItemList(i).Fimgsmall        = rsget("offimgsmall")
				end if

				FItemList(i).Fregdate        = rsget("regdate")

				if IsNULL(FItemList(i).Flastrealdate) then FItemList(i).Flastrealdate="1901-01-01"
				if IsNULL(FItemList(i).Flastrealno) then FItemList(i).Flastrealno=0
				if IsNULL(FItemList(i).Fipno) then FItemList(i).Fipno=0
				if IsNULL(FItemList(i).Freno) then FItemList(i).Freno=0
				if IsNULL(FItemList(i).Fupcheipno) then FItemList(i).Fupcheipno=0
				if IsNULL(FItemList(i).Fupchereno) then FItemList(i).Fupchereno=0
				if IsNULL(FItemList(i).Fsellno) then FItemList(i).Fsellno=0
				if IsNULL(FItemList(i).Fcurrno) then FItemList(i).Fcurrno=0

				if IsNULL(FItemList(i).Fimgsmall) then

				elseif	FItemList(i).Fitemgubun="10" then
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				else
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				end if

				FItemList(i).FinputedRealStock        = rsget("realjeago")
				if IsNULL(FItemList(i).FinputedRealStock) then FItemList(i).FinputedRealStock=0

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub


	public Sub GetRealJaegoList()
		dim i,sqlStr
		sqlStr = "select count(idx) as cnt from [db_shop].[dbo].tbl_shop_realjaego_master"
		sqlStr = sqlStr + " where shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and cancelyn='N'"
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and makerid='" + FRectMakerid + "'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_realjaego_master"
		sqlStr = sqlStr + " where shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and cancelyn='N'"
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and makerid='" + FRectMakerid + "'"
		end if
		sqlStr = sqlStr + " order by jeagodate desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopRealJaegoMaster
				FItemList(i).Fidx       = rsget("idx")
				FItemList(i).Fshopid    = rsget("shopid")
				FItemList(i).Fmakerid   = rsget("makerid")
				FItemList(i).Fjeagodate = rsget("jeagodate")
				FItemList(i).Fregdate   = rsget("regdate")
				FItemList(i).Fcancelyn  = rsget("cancelyn")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetOneJeagoMaster()
		dim i,sqlStr
		sqlStr = " select top 1 idx,shopid,makerid,convert(varchar(32),jeagodate,20) as jeagodate,regdate, cancelyn"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_realjaego_master"
		sqlStr = sqlStr + " where idx='" + FRectIDX + "'"
		rsget.Open sqlStr,dbget,1

		if  not rsget.EOF  then
			set FOneItem = new COffShopRealJaegoMaster
			FOneItem.Fidx       = rsget("idx")
			FOneItem.Fshopid    = rsget("shopid")
			FOneItem.Fmakerid   = rsget("makerid")
			FOneItem.Fjeagodate = rsget("jeagodate")
			FOneItem.Fregdate   = rsget("regdate")
			FOneItem.Fcancelyn  = rsget("cancelyn")
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
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
