<%
'####################################################
' Description :  목표매출 클래스
' History : 2013.03.06 한용민 생성
'####################################################

class ctargetmaechul_item
	public fyyyymm
	public fshopid
	public fgubuntype
	public fgubun
	public ftargetmaechul
	public fregdate
	public flastupdate
	public flastadminid
	public fsolar_date
	public fshopname
	public fgubunname
	public frealsellsum
	public fspendmile
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class ctargetmaechul_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectyyyy
	public FRectyyyy1
	public FRectmm1
	public frectshopid
	public frectgubuntype
	public frectgubun

	'//common/offshop/maechul/targetmaechul/targetmaechul.asp
	public Sub gettargetmaechul()
		dim sqlStr,i
		
		if frectgubuntype = "" or FRectyyyy = "" or FRectshopid = "" or frectgubun = "" then exit Sub
			
		sqlStr = "exec db_shop.dbo.sp_ten_targetmaechul_month_off "&frectgubuntype&",'"&FRectyyyy&"','"&FRectshopid&"',"&frectgubun&""
		
		'Response.write sqlStr &"<br>"
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget
		
		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount
		
		redim preserve FItemList(ftotalcount)

		i=0
		if  not rsget.EOF  then
			
			do until rsget.EOF
				set fitemlist(i) = new ctargetmaechul_item
				
				FItemList(i).frealsellsum          = rsget("realsellsum")
				FItemList(i).fsolar_date          = rsget("solar_date")
				FItemList(i).fyyyymm          = rsget("yyyymm")
				FItemList(i).fshopid          = rsget("shopid")
				FItemList(i).fgubuntype          = rsget("gubuntype")
				FItemList(i).fgubun          = rsget("gubun")
				FItemList(i).ftargetmaechul          = rsget("targetmaechul")
				FItemList(i).fregdate          = rsget("regdate")
				FItemList(i).flastupdate          = rsget("lastupdate")
				FItemList(i).flastadminid          = rsget("lastadminid")
		            
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close

	end sub

	'/common/offshop/maechul/targetmaechul/targetmaechul_sub.asp
	public Sub ftarget_zone()
		dim sqlStr,i
		
		if frectgubuntype = "" or FRectyyyy1 = "" or FRectmm1 = "" or FRectshopid = "" then exit Sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " z.shopid ,isnull(z.idx,0) as gubun,(u.shopname+' 조닝('+z.zonename+')') as gubunname"
		sqlStr = sqlStr & " ,t.gubuntype ,isnull(t.targetmaechul,0) as targetmaechul, t.regdate"
		sqlStr = sqlStr & " ,t.lastupdate ,t.lastadminid ,t.yyyymm, u.shopname"
		sqlStr = sqlStr & " ,isnull(o.realsellsum,0) as realsellsum"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_brand_zone z"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on z.shopid=u.userid"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_targetmaechul_month_off t"
		sqlStr = sqlStr & " 	on z.shopid = t.shopid"
		sqlStr = sqlStr & " 	and z.idx = t.gubun"
		sqlStr = sqlStr & " 	and t.gubuntype="&frectgubuntype&""
		sqlStr = sqlStr & " 	and t.yyyymm = '"&FRectyyyy1&"-"&FRectmm1&"'"
		sqlStr = sqlStr & " left join ("
		sqlStr = sqlStr & " 	select"
		sqlStr = sqlStr & " 	m.shopid ,d.zoneidx"
		sqlStr = sqlStr & " 	, sum( (d.realsellprice+isnull(d.addtaxcharge,0)) *d.itemno) as realsellsum"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " 	join db_shop.dbo.tbl_shopjumun_detail d"
		sqlStr = sqlStr & " 		on m.orderno = d.orderno"
		sqlStr = sqlStr & " 		and m.cancelyn='N'"
		sqlStr = sqlStr & " 		and d.cancelyn='N'"
		sqlStr = sqlStr & " 	where convert(varchar(7),m.IXyyyymmdd,121) = '"&FRectyyyy1-1&"-"&FRectmm1&"'"
		sqlStr = sqlStr & " 	and m.shopid ='"&FRectshopid&"'"
		sqlStr = sqlStr & " 	and isnull(d.zoneidx,'') <> ''"
		sqlStr = sqlStr & " 	group by m.shopid ,d.zoneidx"
		sqlStr = sqlStr & " ) o"
		sqlStr = sqlStr & " 	on z.shopid = o.shopid"
		sqlStr = sqlStr & " 	and z.idx = o.zoneidx"
		sqlStr = sqlStr & " where z.isusing = 'Y'"
		sqlStr = sqlStr & " and z.shopid ='"&FRectshopid&"'"
		sqlStr = sqlStr & " order by idx asc"
		
		'Response.write sqlStr &"<br>"
        rsget.Open sqlStr,dbget,1
		
		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount
		
		redim preserve FItemList(ftotalcount)

		i=0
		if  not rsget.EOF  then
			
			do until rsget.EOF
				set fitemlist(i) = new ctargetmaechul_item
				
				FItemList(i).fyyyymm          = rsget("yyyymm")
				FItemList(i).frealsellsum          = rsget("realsellsum")
				FItemList(i).fshopid          = rsget("shopid")
				FItemList(i).fshopname          = rsget("shopname")
				FItemList(i).fgubuntype          = rsget("gubuntype")
				FItemList(i).fgubun          = rsget("gubun")
				FItemList(i).fgubunname          = db2html(rsget("gubunname"))
				FItemList(i).ftargetmaechul          = rsget("targetmaechul")
				FItemList(i).fregdate          = rsget("regdate")
				FItemList(i).flastupdate          = rsget("lastupdate")
				FItemList(i).flastadminid          = rsget("lastadminid")
		            
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close

	end sub

	'/common/offshop/maechul/targetmaechul/targetmaechul_sub.asp
	public Sub ftarget()
		dim sqlStr,i
		
		if frectgubuntype = "" or FRectyyyy1 = "" or FRectmm1 = "" or FRectshopid = "" then exit Sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " u.userid as shopid, u.shopname, 0 as gubun,u.shopname as gubunname"
		sqlStr = sqlStr & " ,t.gubuntype ,isnull(t.targetmaechul,0) as targetmaechul, t.regdate"
		sqlStr = sqlStr & " ,t.lastupdate ,t.lastadminid ,t.yyyymm"
		sqlStr = sqlStr & " ,isnull(o.realsellsum,0) as realsellsum, isnull(o.spendmile,0) as spendmile"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_targetmaechul_month_off t"
		sqlStr = sqlStr & " 	on u.userid = t.shopid"
		sqlStr = sqlStr & " 	and t.gubuntype="&frectgubuntype&""
		sqlStr = sqlStr & " 	and t.yyyymm = '"&FRectyyyy1&"-"&FRectmm1&"'"
		sqlStr = sqlStr & " left join ("
		sqlStr = sqlStr & " 	select top 1"
		sqlStr = sqlStr & " 	m.shopid"
		sqlStr = sqlStr & " 	, sum(realsum) as realsellsum, sum(spendmile) as spendmile"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " 	where m.cancelyn='N'"
		sqlStr = sqlStr & " 	and convert(varchar(7),m.IXyyyymmdd,121) = '"&FRectyyyy1-1&"-"&FRectmm1&"'"
		sqlStr = sqlStr & " 	and m.shopid ='"&FRectshopid&"'"
		sqlStr = sqlStr & " 	group by m.shopid"
		sqlStr = sqlStr & " ) o"
		sqlStr = sqlStr & " 	on u.userid = o.shopid"
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & " and u.userid ='"&FRectshopid&"'"
		
		'Response.write sqlStr &"<br>"
        rsget.Open sqlStr,dbget,1
		
		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount
		
		redim preserve FItemList(ftotalcount)

		i=0
		if  not rsget.EOF  then
			
			do until rsget.EOF
				set fitemlist(i) = new ctargetmaechul_item
				
				FItemList(i).fspendmile          = rsget("spendmile")
				FItemList(i).frealsellsum          = rsget("realsellsum") + FItemList(i).fspendmile
				FItemList(i).fyyyymm          = rsget("yyyymm")				
				FItemList(i).fshopid          = rsget("shopid")
				FItemList(i).fshopname          = rsget("shopname")
				FItemList(i).fgubuntype          = rsget("gubuntype")
				FItemList(i).fgubun          = rsget("gubun")
				FItemList(i).fgubunname          = db2html(rsget("gubunname"))
				FItemList(i).ftargetmaechul          = rsget("targetmaechul")
				FItemList(i).fregdate          = rsget("regdate")
				FItemList(i).flastupdate          = rsget("lastupdate")
				FItemList(i).flastadminid          = rsget("lastadminid")
		            
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close

	end sub
	
	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
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