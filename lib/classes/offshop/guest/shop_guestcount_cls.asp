<%
'###########################################################
' Description :  매장고객방문카운트 클래스
' History : 2012.05.10 한용민 생성
'###########################################################

Class cguestcount_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fshopid
	public fyyyymmdd
	public fyyyymm
	public fz1_in
	public fz1_out
	public fz1_all
	public fz2_in
	public fz2_out
	public fz2_all
	public fregadminuserid
	public fregdate
	public flastadminuserid
	public flastupdate
	public fshopname
	public fsumtotal
	public FSum
	public FCount
	public FWeather
	public FWeatherComm
end class

class cguestcount_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FRectShopID
	public FRectStartDay
	public FRectEndDay
	public FRectInc3pl
	public FRectExcCancel

	'//common/offshop/guest/shop_guestcount_yyyymm.asp
	public sub fshopguestcount_yyyymm()
		dim sqlStr,i ,sqlsearch

		if FRectStartDay <> "" and FRectEndDay <> "" then
			sqlsearch = sqlsearch & " and yyyymmdd >= '"&FRectStartDay&"'"
			sqlsearch = sqlsearch & " and yyyymmdd < '"&FRectEndDay&"'"
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and shopid = '"&frectshopid&"'"
		end if

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " g.shopid ,convert(varchar(7),g.yyyymmdd,121) as yyyymm"
		sqlStr = sqlStr & " ,sum(g.z1_in) as z1_in, sum(g.z1_out) as z1_out, sum(g.z2_in) as z2_in, sum(g.z2_out) as z2_out"
		sqlStr = sqlStr & " ,sum(round(convert(float,isnull(g.z1_in,0)+isnull(g.z1_out,0))/2,0)) as z1_all"
		sqlStr = sqlStr & " ,sum(round(convert(float,isnull(g.z2_in,0)+isnull(g.z2_out,0))/2,0)) as z2_all"
		sqlStr = sqlStr & " ,u.shopname"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_guestcount g"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on g.shopid = u.userid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " group by shopid ,convert(varchar(7),g.yyyymmdd,121),u.shopname"
		sqlStr = sqlStr & " order by shopid asc , yyyymm desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount
		ftotalcount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cguestcount_item

				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fyyyymm = rsget("yyyymm")
				FItemList(i).fz1_in = rsget("z1_in")
				FItemList(i).fz1_out = rsget("z1_out")
				FItemList(i).fz1_all = rsget("z1_all")
				FItemList(i).fz2_in = rsget("z2_in")
				FItemList(i).fz2_out = rsget("z2_out")
				FItemList(i).fz2_all = rsget("z2_all")
				FItemList(i).fshopname = rsget("shopname")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//common/offshop/guest/shop_guestcount_yyyymmdd.asp
	public sub fshopguestcount_yyyymmdd()
		dim sqlStr,i ,sqlsearch ,sqlsearch2

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch2 = sqlsearch2 & " and isNULL(jp.tplcompanyid,'')<>''"
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch2 = sqlsearch2 & " and isNULL(jp.tplcompanyid,'')=''"
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectStartDay <> "" and FRectEndDay <> "" then
			sqlsearch = sqlsearch & " and g.yyyymmdd >= '"&FRectStartDay&"'"
			sqlsearch = sqlsearch & " and g.yyyymmdd < '"&FRectEndDay&"'"
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and g.shopid = '"&frectshopid&"'"
		end if
		if FRectStartDay <> "" and FRectEndDay <> "" then
			sqlsearch2 = sqlsearch2 & " and m.ixyyyymmdd >= '"&FRectStartDay&"'"
			sqlsearch2 = sqlsearch2 & " and m.ixyyyymmdd < '"&FRectEndDay&"'"
		end if
		if frectshopid <> "" then
			sqlsearch2 = sqlsearch2 & " and m.shopid = '"&frectshopid&"'"
		end if

		if FRectExcCancel = "Y" then
			sqlsearch2 = sqlsearch2 & " and m.cancelyn = 'N' "
			sqlsearch2 = sqlsearch2 & " and d.cancelyn = 'N' "
		end if

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " g.shopid ,convert(varchar(10),g.yyyymmdd,121) as yyyymmdd"
		sqlStr = sqlStr & " ,sum(g.z1_in) as z1_in, sum(g.z1_out) as z1_out, sum(g.z2_in) as z2_in, sum(g.z2_out) as z2_out"
		sqlStr = sqlStr & " ,sum(round(convert(float,isnull(g.z1_in,0)+isnull(g.z1_out,0))/2,0)) as z1_all"
		sqlStr = sqlStr & " ,sum(round(convert(float,isnull(g.z2_in,0)+isnull(g.z2_out,0))/2,0)) as z2_all"
		sqlStr = sqlStr & " ,u.shopname"
		sqlStr = sqlStr & " , isNull(jumun.sumtotal,0) as sumtotal, isNull(jumun.cnt,0) as cnt"
		sqlStr = sqlStr & " , isNull(jumun.sellsum,0) as sellsum, isNull(w.weather,0) as weather"
		sqlStr = sqlStr & " , isNull(convert(varchar(1000),w.comment),'') as comment"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_guestcount g"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on g.shopid = u.userid"
		sqlStr = sqlStr & "	left join"
		sqlStr = sqlStr & "	("
		sqlStr = sqlStr & "		select m.shopid, convert(varchar(10),m.ixyyyymmdd,121) as yyyymmdd"
		sqlStr = sqlStr & "		,isnull(sum((d.sellprice+isnull(d.addtaxcharge,0))*d.itemno),0) as sumtotal, count(distinct(m.idx)) as cnt"
		sqlStr = sqlStr & "		,isnull(sum((d.realsellprice+isnull(d.addtaxcharge,0))*d.itemno),0) as sellsum"
		sqlStr = sqlStr & "		from db_shop.dbo.tbl_shopjumun_master as m"
		sqlStr = sqlStr & "		inner join db_shop.dbo.tbl_shopjumun_detail as d on m.idx = d.masteridx"
		sqlStr = sqlStr & " 	left join db_partner.dbo.tbl_partner jp"
	    sqlStr = sqlStr & "       	on m.shopid=jp.id "
		sqlStr = sqlStr & "		where 1=1 " & sqlsearch2
		sqlStr = sqlStr & "		group by m.shopid ,convert(varchar(10),m.ixyyyymmdd,121)"
		sqlStr = sqlStr & "	) as jumun"
		sqlStr = sqlStr & "		on g.shopid = jumun.shopid and convert(varchar(10),g.yyyymmdd,121) = jumun.yyyymmdd"
		sqlStr = sqlStr & "	left join db_shop.dbo.tbl_shop_weather as w"
		sqlStr = sqlStr & " 	on g.shopid = w.shopid and convert(varchar(10),g.yyyymmdd,121) = convert(varchar(10),w.wdate,121)"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & " 	on g.shopid=p.id "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " group by g.shopid, convert(varchar(10),g.yyyymmdd,121), u.shopname, jumun.sumtotal"
		sqlStr = sqlStr & " 	, jumun.cnt, jumun.sellsum, w.weather, convert(varchar(1000),w.comment)"
		sqlStr = sqlStr & " order by g.shopid asc, convert(varchar(10),g.yyyymmdd,121) desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount
		ftotalcount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cguestcount_item

				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fyyyymmdd = rsget("yyyymmdd")
				FItemList(i).fz1_in = rsget("z1_in")
				FItemList(i).fz1_out = rsget("z1_out")
				FItemList(i).fz1_all = rsget("z1_all")
				FItemList(i).fz2_in = rsget("z2_in")
				FItemList(i).fz2_out = rsget("z2_out")
				FItemList(i).fz2_all = rsget("z2_all")
				FItemList(i).fshopname = rsget("shopname")
				FItemList(i).fsumtotal = rsget("sumtotal")
				FItemList(i).FSum   = rsget("sellsum")
				FItemList(i).FCount = rsget("cnt")
				FItemList(i).FWeather = rsget("weather")
				FItemList(i).FWeatherComm = rsget("comment")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//common/offshop/guest/shop_guestcount_yyyymmddhh.asp
	public sub fshopguestcount_yyyymmddhh()
		dim sqlStr,i ,sqlsearch

		if FRectStartDay <> "" and FRectEndDay <> "" then
			sqlsearch = sqlsearch & " and yyyymmdd >= '"&FRectStartDay&"'"
			sqlsearch = sqlsearch & " and yyyymmdd < '"&FRectEndDay&"'"
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and shopid = '"&frectshopid&"'"
		end if

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " g.shopid ,g.yyyymmdd ,g.z1_in ,g.z1_out ,g.z2_in ,g.z2_out ,g.regadminuserid ,g.regdate"
		sqlStr = sqlStr & " ,round(convert(float,isnull(g.z1_in,0)+isnull(g.z1_out,0))/2,0) as z1_all"
		sqlStr = sqlStr & " ,round(convert(float,isnull(g.z2_in,0)+isnull(g.z2_out,0))/2,0) as z2_all"
		sqlStr = sqlStr & " ,g.lastadminuserid ,g.lastupdate"
		sqlStr = sqlStr & " ,u.shopname"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_guestcount g"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on g.shopid = u.userid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by shopid asc , yyyymmdd desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount
		ftotalcount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cguestcount_item

				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fyyyymmdd = rsget("yyyymmdd")
				FItemList(i).fz1_in = rsget("z1_in")
				FItemList(i).fz1_out = rsget("z1_out")
				FItemList(i).fz1_all = rsget("z1_all")
				FItemList(i).fz2_in = rsget("z2_in")
				FItemList(i).fz2_out = rsget("z2_out")
				FItemList(i).fz2_all = rsget("z2_all")
				FItemList(i).fregadminuserid = rsget("regadminuserid")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastadminuserid = rsget("lastadminuserid")
				FItemList(i).flastupdate = rsget("lastupdate")
				FItemList(i).fshopname = rsget("shopname")

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
end Class

function getzonegubun(shopid ,zone)
	if shopid = "streetshop011" then
		if zone = "z1_in" then
			getzonegubun = "MAIN"
		else
			getzonegubun = "SIDE"
		end if
	elseif shopid = "streetshop017" then
		if zone = "z1_in" then
			getzonegubun = "MAIN"
		else
			getzonegubun = "X"
		end if
	elseif shopid = "streetshop018" then
		if zone = "z1_in" then
			getzonegubun = "우측"
		else
			getzonegubun = "좌측"
		end if
	elseif shopid = "streetshop014" then
		if zone = "z1_in" then
			getzonegubun = "MAIN"
		else
			getzonegubun = "X"
		end if
	elseif shopid = "streetshop020" then
		if zone = "z1_in" then
			getzonegubun = "MAIN"
		else
			getzonegubun = "SIDE"
		end if
	else
		getzonegubun = zone
	end if
end function

function existsguestcountshopid(tmp)
	if tmp = "" then exit function

	if tmp = "streetshop011" then
		existsguestcountshopid = true
	elseif tmp = "streetshop017" then
		existsguestcountshopid = true
	elseif tmp = "streetshop018" then
		existsguestcountshopid = true
	elseif tmp = "streetshop014" then
		existsguestcountshopid = true
	elseif tmp = "streetshop104" then
		existsguestcountshopid = true
	elseif tmp = "streetshop020" then
		existsguestcountshopid = true
	elseif tmp = "streetshop025" then
		existsguestcountshopid = true
	elseif tmp = "streetshop026" then
		existsguestcountshopid = true
	elseif tmp = "streetshop024" then
		existsguestcountshopid = true
	else
		existsguestcountshopid = false
	end if
end function
%>
