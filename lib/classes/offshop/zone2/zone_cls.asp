<%
'###########################################################
' Description : 삽별구역설정
' Hieditor : 2010.12.29 한용민 생성
'###########################################################

Class czone_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fmanagershopyn
	public fusername
	public fmanageridx
	public fparttype
	public fempno	
	public fIXyyyymmdd
	public fyyyy
	public fsuplypricesum
	public frealpyeong
	public fpyeong
	public fidx
	public fshopid
	public fzonename
	public fracktype
	public funit
	public fzonegroup
	public fregdate
	public fisusing
	public fitemgubun
	public fshopitemid
	public fitemoption
	public fzoneidx
	public fmakerid
	public fshopitemname
	public fitemname
	public fitemoptionname
	public fshopitemoptionname
	public fshopitemprice
	public frealsellprice
	public fshopsuplycash
	public forgsellprice
	public fshopbuyprice
	public fitemcnt
	public fsellsum
	public funitvalue
	public fsellprice
	public forderno
	public fitemno
	public fitemidx
	public fstartdate
	public fenddate
	public fcdl_nm
	public fcdm_nm
	public fcds_nm
	public fitemnosum
	public frealmaechul
	public fsuplymaechul
	public FCateName
	public fcatecdl
	public fcatecdm
	public fcatecds
	public fmakername
	public fselect_makerid
	public ftargetmaechul
	public fblock	
end class

class czone_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public frectisusing
	public frectshopid
	public frectidx
	public FRectStartDay
	public FRectEndDay
	public FRectmakerid
	public FCountTotal
	public FSumTotal
	public ftmpselldate
	public ftmpSumTotal
	public fblockcnt
	public fprofitTotal
	public frectitemid
	public frectitemname	
	public frectzonegroup
	public frectracktype
	public FRectstatusdiv
	public FRectCDL
	public FRectCDM
	public FRectCDN
	public frectdatefg
	public FRectsearchtype
	public frectsellgubun
	public maxt
	public maxc
	public FRectzoneidx
	public FRectviewzone
	public frectdategubun
	public frectordertype
	public frectzoneisusing
	public frectsearchgubun
	public frectparttype
	public frectpart_sn
	public frectempno
	public FRectInc3pl
	
	function MaxVal(a,b)
		if a = "" then a = 0
		if b = "" then b = 0		
	
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function
		
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	    maxt = -1
	    maxc = -1
	    FCountTotal = 0
	    FSumTotal = 0
	    fprofitTotal = 0
	    ftmpselldate = ""
	    ftmpSumTotal = ""
	    fblockcnt = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'//admin/offshop/zone2/zone_reg.asp
	public sub Getshopzonemanager()
		dim sqlStr,i , sqlsearch
			
		if frectparttype <> "" then
			sqlsearch = sqlsearch & " and zm.parttype = '"&frectparttype&"'"
		end if
		if frectzoneidx <> "" then
			sqlsearch = sqlsearch & " and zm.zoneidx = '"&frectzoneidx&"'"
		end if
		if frectempno <> "" then
			sqlsearch = sqlsearch & " and zm.empno = '"&frectempno&"'"
		end if
		if frectpart_sn <> "" then
			sqlsearch = sqlsearch & " and t.part_sn in ("&frectpart_sn&")"
		end if
						
		'데이터 리스트 
		sqlStr = "SELECT distinct TOP " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & "	t.posit_sn ,t.empno ,t.username , t.part_sn"
		sqlStr = sqlStr & "	from db_partner.dbo.tbl_user_tenbyten t"
		sqlStr = sqlStr & "	left join db_shop.dbo.tbl_shop_brand_zone_manager zm"
		sqlStr = sqlStr & "		on t.empno = zm.empno"
		sqlStr = sqlStr & "	left join db_partner.dbo.tbl_partner_shopuser su"
		sqlStr = sqlStr & "		on t.empno = su.empno"
		sqlStr = sqlStr & "		and su.firstisusing='Y'"
		sqlStr = sqlStr & "	WHERE su.empno is not null " & sqlsearch

		' 퇴사예정자 처리	' 2018.10.16 한용민
		'sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr & "	order by t.part_sn asc, t.posit_sn asc , t.username asc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		FTotalCount = rsget.recordcount
		fresultcount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new czone_oneitem
				
				FItemList(i).fempno = rsget("empno")
				FItemList(i).fusername = db2html(rsget("username"))
						
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
				
	'//admin/offshop/zone2/zone_reg.asp
    public Sub fzone_oneitem()
        dim sqlStr , sqlsearch

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "&frectidx&""
		end if
        
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " idx, shopid, zonename, unit, regdate, isusing, managershopyn" + vbcrlf		
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_brand_zone" + vbcrlf	
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new czone_oneitem
        
        if Not rsget.Eof then
    		
    		FOneItem.fmanagershopyn = rsget("managershopyn")
			FOneItem.fidx = rsget("idx")
			FOneItem.fshopid = rsget("shopid")
			FOneItem.fzonename = db2html(rsget("zonename"))
			FOneItem.funit = rsget("unit")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fisusing = rsget("isusing")
						           
        end if
        rsget.Close
    end Sub
    
	'//admin/offshop/zone/zone.asp
	public sub fzone_list()
		dim sqlStr,i , sqlsearch
		
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and z.shopid = '"&frectshopid&"'"
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and z.isusing = '"&frectisusing&"'"
		end if

		'총 갯수 구하기
		sqlStr = "select count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_brand_zone z" + vbcrlf
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on z.shopid = u.userid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub
					
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " z.idx,z.shopid,z.zonename,z.unit,z.regdate,z.isusing ,z.managershopyn"
		sqlStr = sqlStr & " ,u.pyeong "
		sqlStr = sqlStr & " ,(select sum(unit)"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shop_brand_zone"
		sqlStr = sqlStr & " 	where isusing='Y' and z.shopid = shopid) as realpyeong"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_brand_zone z" + vbcrlf
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr & " 	on z.shopid = u.userid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by z.shopid asc, z.idx asc" + vbcrlf

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
				set FItemList(i) = new czone_oneitem
				
				FItemList(i).fmanagershopyn = rsget("managershopyn")
				FItemList(i).frealpyeong = rsget("realpyeong")
				FItemList(i).fpyeong = rsget("pyeong")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fshopid = rsget("shopid")		
				FItemList(i).fzonename = db2html(rsget("zonename"))
				FItemList(i).funit = rsget("unit")				
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")

				FSumTotal = FSumTotal + FItemList(i).funit
						
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/offshop/zone/zone_item.asp
	public sub GetoffshopzoneitemMatch()
		dim sqlStr,i , sqlsearch ,sqlsearch2

		if frectmakerid <> "" then
			sqlsearch = sqlsearch & " and d.makerid = '"&frectmakerid&"'"
		end if
		if frectzoneisusing = "Y" then
			sqlsearch = sqlsearch & " and zd.zoneidx is not null"
		elseif frectzoneisusing = "N" then
			sqlsearch = sqlsearch & " and zd.zoneidx is null"
		end if
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and zd.zoneidx = "&frectidx&""
		end if
		
		'데이터 리스트 
		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " d.makerid,zd.regdate, zd.zoneidx, z.managershopyn"
		sqlStr = sqlStr & " ,(SELECT TOP 1 IsNull(zonename,'')"
		sqlStr = sqlStr & " 	FROM [db_shop].[dbo].[tbl_shop_brand_zone]"
		sqlStr = sqlStr & "		WHERE idx = zd.zoneidx AND shopid = '" & frectshopid & "') AS zonename"
		
		'/입고된전체브랜드
		if frectsearchgubun = "A" then
			sqlStr = sqlStr & " FROM db_shop.dbo.tbl_shop_designer d"
			sqlStr = sqlStr & " LEFT JOIN [db_shop].[dbo].[tbl_shop_brand_zone_detail] AS zd"
			sqlStr = sqlStr & " 	ON d.makerid = zd.makerid"
			sqlStr = sqlStr & " 	AND zd.shopid = '" & frectshopid & "'"
			sqlStr = sqlStr & " LEFT join [db_shop].[dbo].[tbl_shop_brand_zone] z"
			sqlStr = sqlStr & " 	on zd.zoneidx = z.idx"
			sqlStr = sqlStr & "	WHERE d.firstipgodate is not null"
			sqlStr = sqlStr & " and d.shopid = '" & frectshopid & "' " & sqlsearch
		
		'/최근3개월판매내역
		else
			sqlStr = sqlStr & " FROM [db_shop].[dbo].[tbl_shopjumun_master] AS m"
			sqlStr = sqlStr & " JOIN [db_shop].[dbo].[tbl_shopjumun_detail] AS d"
			sqlStr = sqlStr & "		ON m.orderno = d.orderno"
			sqlStr = sqlStr & " 	AND m.cancelyn = 'N' AND d.cancelyn = 'N'"
			sqlStr = sqlStr & " 	AND m.regdate >= dateadd(m,-3,getdate())"
			sqlStr = sqlStr & " 	AND m.shopid = '" & frectshopid & "'"
			sqlStr = sqlStr & " LEFT JOIN [db_shop].[dbo].[tbl_shop_brand_zone_detail] AS zd"
			sqlStr = sqlStr & " 	ON d.makerid = zd.makerid AND zd.shopid = '" & frectshopid & "'"
			sqlStr = sqlStr & " LEFT join [db_shop].[dbo].[tbl_shop_brand_zone] z"
			sqlStr = sqlStr & " 	on zd.zoneidx = z.idx"
			sqlStr = sqlStr & "	WHERE 1=1 " & sqlsearch
		end if
		
		sqlStr = sqlStr & "	GROUP BY d.makerid,zd.regdate, zd.zoneidx, z.managershopyn"
		sqlStr = sqlStr & " ORDER BY zd.zoneidx asc, d.makerid asc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		FTotalCount = rsget.recordcount
		fresultcount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new czone_oneitem
				
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fmanagershopyn = rsget("managershopyn")
				FItemList(i).fmakerid = db2html(rsget("makerid"))
				FItemList(i).fzoneidx = rsget("zoneidx")
				FItemList(i).fzonename = db2html(rsget("zonename"))
						
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/offshop/zone/zone_sum_category.asp
	public sub Getoffshopzonesum_category()
		dim sqlStr,i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&frectshopid&"'"
		end if
		if frectmakerid <> "" then
			sqlsearch = sqlsearch & " and d.makerid = '"&frectmakerid&"'"
		end if
		if frectzoneidx <> "" then
			sqlsearch = sqlsearch & " and z.idx = "&frectzoneidx&""
		end if	
						
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
			
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if		
		end if

		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " m.shopid , z.zonename ,z.unit ,z.idx ,z.managershopyn"
		
		if frectdategubun = "D" then 
			sqlStr = sqlStr & " ,m.IXyyyymmdd ,isnull(c.targetmaechul,0) as targetmaechul"
		elseif frectdategubun = "M" then 
			sqlStr = sqlStr & " ,convert(varchar(7),m.IXyyyymmdd) as IXyyyymmdd ,isnull(c.targetmaechul,0) as targetmaechul"
		end if
		
		sqlStr = sqlStr & " ,sum(d.itemno) as itemcnt, sum(d.realsellprice*d.itemno) as sellsum"
		sqlStr = sqlStr & " ,sum(d.suplyprice*d.itemno) as suplypricesum"
		sqlStr = sqlStr & " ,(select sum(unit)"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shop_brand_zone"
		sqlStr = sqlStr & " 	where isusing='Y' and m.shopid = shopid) as realpyeong"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_designer s"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail d"
		sqlStr = sqlStr & " 	on s.makerid = d.makerid"
		sqlStr = sqlStr & " 	and s.shopid = '"&frectshopid&"'"
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " 	on m.orderno=d.orderno"
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "

		'/결제내역기준
		if frectsellgubun = "S" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone z"
			sqlStr = sqlStr & " 	on d.zoneidx = z.idx"
		
		'/현재등록내역기준		
		elseif frectsellgubun = "N" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone_detail zi"
			sqlStr = sqlStr & " 	on m.shopid = zi.shopid"
			sqlStr = sqlStr & " 	and d.makerid = zi.makerid"
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone z"
			sqlStr = sqlStr & " 	on zi.zoneidx = z.idx"
		end if

		if frectdategubun = "D" or frectdategubun = "M" then 
	        sqlStr = sqlStr & " left join db_shop.dbo.tbl_targetmaechul_month_off c"

			'//주문일 기준
			if frectdatefg = "jumun" then
				sqlStr = sqlStr & " 	on convert(varchar(7),m.shopregdate,121) = c.yyyymm"
				
			'//매출일 기준
			elseif frectdatefg = "maechul" then
				sqlStr = sqlStr & " 	on convert(varchar(7),m.IXyyyymmdd,121) = c.yyyymm"

			else
				sqlStr = sqlStr & " 	on convert(varchar(7),m.shopregdate,121) = c.yyyymm"		
			end if

	        sqlStr = sqlStr & " 	and m.shopid = c.shopid"
	        sqlStr = sqlStr & " 	and c.gubuntype = 2"
	        sqlStr = sqlStr & " 	and d.zoneidx = c.gubun"
		end if

		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " group by m.shopid , z.zonename, z.unit, z.idx, z.managershopyn"

		if frectdategubun = "D" then 
			sqlStr = sqlStr & " ,m.IXyyyymmdd,c.targetmaechul"
		elseif frectdategubun = "M" then 
			sqlStr = sqlStr & " ,convert(varchar(7),m.IXyyyymmdd),c.targetmaechul"
		end if

		sqlStr = sqlStr & " order by"

		if frectdategubun = "D" then 
			sqlStr = sqlStr & " m.IXyyyymmdd asc,"
		elseif frectdategubun = "M" then
			sqlStr = sqlStr & " IXyyyymmdd asc,"
		end if
		
		sqlStr = sqlStr & " z.idx asc"
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.recordcount

		redim preserve FItemList(FTotalCount)
		
		i=0
		if  not rsget.EOF  then			
			do until rsget.EOF
				set FItemList(i) = new czone_oneitem
					    
				if frectdategubun = "D" then
					FItemList(i).fIXyyyymmdd = rsget("IXyyyymmdd")
					FItemList(i).ftargetmaechul = rsget("targetmaechul")
				elseif frectdategubun = "M" then
					FItemList(i).fIXyyyymmdd = rsget("IXyyyymmdd")
					FItemList(i).ftargetmaechul = rsget("targetmaechul")
				end if
				
				FItemList(i).fmanagershopyn = rsget("managershopyn")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fzonename = db2html(rsget("zonename"))	
				FItemList(i).funit = rsget("unit")
				FItemList(i).fitemcnt = rsget("itemcnt")
				FItemList(i).fsellsum = rsget("sellsum")
				FItemList(i).fsuplypricesum = rsget("suplypricesum")
				FItemList(i).frealpyeong = rsget("realpyeong")								
				
				'//리스트 중간 합계 처리 , 중간합계를 더해서 변수에 넣음
				if ftmpselldate <> FItemList(i).fIXyyyymmdd and i <> 0 then
					ftmpSumTotal = ftmpSumTotal & FSumTotal & ","
					FSumTotal = 0
					fblockcnt = fblockcnt + 1
				end if
				
				FCountTotal = FCountTotal + FItemList(i).fitemcnt
				FSumTotal = FSumTotal + FItemList(i).fsellsum
				ftmpselldate = FItemList(i).fIXyyyymmdd
				
				'//리스트 맨 마지막일경우 중간합계 더함
				if i = (FTotalCount-1) then
					ftmpSumTotal = ftmpSumTotal & FSumTotal & ","
				end if
				
				'/리스트의 해당 중간 합계의 배열위치를 넣음
				FItemList(i).fblock = fblockcnt
				
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/offshop/zone/zone_sum_category_detail.asp
	public sub Getoffshopzone_detailCategory()
		dim sqlStr,i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + CStr(FRectmakerid) + "'"
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&frectshopid&"'"
		end if
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and z.idx = "&frectidx&""
		else
			sqlsearch = sqlsearch & " and z.idx is null"	
		end if	
		
		if FRectCDL<>"" then
			sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "'"
		end if
		if FRectCDL<>"" and FRectCDM<>"" then
			sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "' and i.catecdm='" + FRectCDM + "'"
		end if
		if FRectCDL<>"" and FRectCDM<>"" and FRectCDN<>"" then
			sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "' and i.catecdm='" + FRectCDM + "' and i.catecdn='" + FRectCDN + "'"
		end if
	
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
			
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if		
		end if
		
		'데이터 리스트 
		sqlStr = "select"
		sqlStr = sqlStr & " m.shopid , z.zonename ,z.unit ,z.idx"
		sqlStr = sqlStr & " ,sum(d.itemno) as itemcnt, sum(d.realsellprice*d.itemno) as sellsum"
		sqlStr = sqlStr & " ,sum(d.suplyprice*d.itemno) as suplypricesum"
		sqlStr = sqlStr & " ,(select sum(unit)"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shop_brand_zone"
		sqlStr = sqlStr & " 	where isusing='Y' and m.shopid = shopid) as realpyeong"

		if FRectCDL<>"" and FRectCDM<>"" then
			sqlStr = sqlStr + " ,i.catecdl,i.catecdm,i.catecdn,cs.code_nm"
		elseif FRectCDL<>"" then
			sqlStr = sqlStr + " ,i.catecdl,i.catecdm,cm.code_nm"
		else
			sqlStr = sqlStr + " ,i.catecdl,cl.code_nm"
		end if

		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_designer s"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail d"
		sqlStr = sqlStr & " 	on s.makerid = d.makerid"
		sqlStr = sqlStr & " 	and s.shopid = '"&frectshopid&"'"
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " 	on m.orderno=d.orderno"
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr + " left Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " 	on d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + " 	and d.itemid=i.shopitemid"
		sqlStr = sqlStr + " 	and d.itemoption=i.itemoption"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "
	    
		'/결제내역기준
		if frectsellgubun = "S" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone z"
			sqlStr = sqlStr & " 	on d.zoneidx = z.idx"
		
		'/현재등록내역기준		
		elseif frectsellgubun = "N" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone_detail zi"
			sqlStr = sqlStr & " 	on m.shopid = zi.shopid"
			sqlStr = sqlStr & " 	and d.makerid = zi.makerid"
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone z"
			sqlStr = sqlStr & " 	on zi.zoneidx = z.idx"
		end if

		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		sqlStr = sqlStr & " group by m.shopid , z.zonename, z.unit, z.idx"		
		
		if FRectCDL<>"" and FRectCDM<>"" then
			sqlStr = sqlStr + " ,i.catecdl,i.catecdm,i.catecdn,cs.code_nm"
		elseif FRectCDL<>"" then
			sqlStr = sqlStr + " ,i.catecdl,i.catecdm,cm.code_nm"
		else
			sqlStr = sqlStr + " ,i.catecdl,cl.code_nm"
		end if
		
		sqlStr = sqlStr + " order by z.idx asc"
		
		if FRectCDL<>"" and FRectCDM<>"" then
			sqlStr = sqlStr + " ,i.catecdl asc ,i.catecdm asc ,i.catecdn asc"
		elseif FRectCDL<>"" then
			sqlStr = sqlStr + " , i.catecdl asc ,i.catecdm asc"
		else
			sqlStr = sqlStr + " ,i.catecdl asc"
		end if
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new czone_oneitem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fzonename = db2html(rsget("zonename"))
				FItemList(i).funit = rsget("unit")
				FItemList(i).fitemcnt = rsget("itemcnt")
				FItemList(i).fsellsum = rsget("sellsum")
				FItemList(i).fsuplypricesum = rsget("suplypricesum")
				FItemList(i).frealpyeong = rsget("realpyeong")								
			    FItemList(i).FCateName 	= db2html(rsget("code_nm"))

				if FRectCDL<>"" and FRectCDM<>"" then
					FItemList(i).fcatecdl 	= rsget("catecdl")
					FItemList(i).fcatecdm 	= rsget("catecdm")
					FItemList(i).fcatecds 	= rsget("catecdn")
				elseif FRectCDL<>"" then
					FItemList(i).fcatecdl 	= rsget("catecdl")
					FItemList(i).fcatecdm 	= rsget("catecdm")
				else
					FItemList(i).fcatecdl 	= rsget("catecdl")
				end if

				if Not IsNull(FItemList(i).frealmaechul) then
					maxt = MaxVal(maxt,FItemList(i).fsellsum)
					maxc = MaxVal(maxc,FItemList(i).fitemcnt)
				end if				
				
				FCountTotal = FCountTotal + FItemList(i).fitemcnt
				FSumTotal = FSumTotal + FItemList(i).fsellsum							
				fprofitTotal = fprofitTotal + (FItemList(i).fsellsum - FItemList(i).fsuplypricesum)
																								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/offshop/zone/zone_sum_brand_detail.asp
	public sub Getoffshopzone_detailbrand()
		dim sqlStr,i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + CStr(FRectmakerid) + "'"
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&frectshopid&"'"
		end if
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and z.idx = "&frectidx&""
		else
			sqlsearch = sqlsearch & " and z.idx is null"	
		end if	
		if FRectCDL<>"" then
			sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "'"
		end if
		if FRectCDL<>"" and FRectCDM<>"" then
			sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "' and i.catecdm='" + FRectCDM + "'"
		end if
		if FRectCDL<>"" and FRectCDM<>"" and FRectCDN<>"" then
			sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "' and i.catecdm='" + FRectCDM + "' and i.catecdn='" + FRectCDN + "'"
		end if
		
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
			
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if		
		end if

		'데이터 리스트 
		sqlStr = "select"
		sqlStr = sqlStr & " m.shopid , z.zonename ,z.unit , z.idx, d.makerid"
		sqlStr = sqlStr & " ,sum(d.itemno) as itemcnt, sum(d.realsellprice*d.itemno) as sellsum"
		sqlStr = sqlStr & " ,sum(d.suplyprice*d.itemno) as suplypricesum"
		sqlStr = sqlStr & " ,(select sum(unit)"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shop_brand_zone"
		sqlStr = sqlStr & " 	where isusing='Y' and m.shopid = shopid) as realpyeong"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_designer s"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail d"
		sqlStr = sqlStr & " 	on s.makerid = d.makerid"
		sqlStr = sqlStr & " 	and s.shopid = '"&frectshopid&"'"
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " 	on m.orderno=d.orderno"
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "
	    
		'/결제내역기준
		if frectsellgubun = "S" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone z"
			sqlStr = sqlStr & " 	on d.zoneidx = z.idx"
		
		'/현재등록내역기준		
		elseif frectsellgubun = "N" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone_detail zi"
			sqlStr = sqlStr & " 	on m.shopid = zi.shopid"
			sqlStr = sqlStr & " 	and d.makerid = zi.makerid"
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone z"
			sqlStr = sqlStr & " 	on zi.zoneidx = z.idx"
		end if

		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		sqlStr = sqlStr & " group by m.shopid , z.zonename, z.unit, z.idx ,d.makerid"
		sqlStr = sqlStr + " order by sellsum desc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new czone_oneitem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fzonename = db2html(rsget("zonename"))
				FItemList(i).funit = rsget("unit")
				FItemList(i).fitemcnt = rsget("itemcnt")
				FItemList(i).fsellsum = rsget("sellsum")
				FItemList(i).fsuplypricesum = rsget("suplypricesum")
				FItemList(i).frealpyeong = rsget("realpyeong")								
				FItemList(i).fmakerid 	= rsget("makerid")

				if Not IsNull(FItemList(i).frealmaechul) then
					maxt = MaxVal(maxt,FItemList(i).fsellsum)
					maxc = MaxVal(maxc,FItemList(i).fitemcnt)
				end if				
				
				FCountTotal = FCountTotal + FItemList(i).fitemcnt
				FSumTotal = FSumTotal + FItemList(i).fsellsum							
				fprofitTotal = fprofitTotal + (FItemList(i).fsellsum - FItemList(i).fsuplypricesum)
																								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/offshop/zone2/zone_sum_item_detail.asp
	public sub Getoffshopzone_detail()
		dim sqlStr,i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + CStr(FRectmakerid) + "'"
		end if
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&frectshopid&"'"
		end if
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and z.idx = "&frectidx&""
		else
			sqlsearch = sqlsearch & " and z.idx is null"
		end if	
		if FRectCDL<>"" then
			sqlsearch = sqlsearch + " and i.catecdl='" + FRectCDL + "'"
		end if
		if FRectCDM<>"" then
			sqlsearch = sqlsearch + " and i.catecdm='" + FRectCDM + "'"
		end if
		if FRectCDN<>"" then
			sqlsearch = sqlsearch + " and i.catecdn='" + FRectCDN + "'"
		end if
		
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
			
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if		
		end if

		if frectitemid<>"" then
			sqlsearch = sqlsearch + " and d.itemid=" + frectitemid + ""
		end if
		if frectitemname<>"" then
			sqlsearch = sqlsearch + " and d.itemname like '%" + frectitemname + "%'"
		end if
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " d.itemgubun , d.itemid ,d.itemoption ,d.makerid, d.itemname"
		sqlStr = sqlStr & " ,d.itemoptionname, d.zoneidx"
		sqlStr = sqlStr & " , z.zonename ,z.unit"
		sqlStr = sqlStr & " , cl.code_nm as cdl_nm, cm.code_nm as cdm_nm, cs.code_nm as cds_nm"
		sqlStr = sqlStr & " ,i.catecdl, i.catecdm, i.catecdn"
		sqlStr = sqlStr & " ,sum(d.itemno) as itemcnt, sum(d.realsellprice*d.itemno) as sellsum ,sum(d.suplyprice*d.itemno) as suplypricesum"
		sqlStr = sqlStr & " ,(select sum(unit)"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shop_brand_zone"
		sqlStr = sqlStr & " 	where isusing='Y' and m.shopid = shopid) as realpyeong"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_designer s"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail d"
		sqlStr = sqlStr & " 	on s.makerid = d.makerid"
		sqlStr = sqlStr & " 	and s.shopid = '"&frectshopid&"'"
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " 	on m.orderno=d.orderno"
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr + " left Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " 	on d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + " 	and d.itemid=i.shopitemid"
		sqlStr = sqlStr + " 	and d.itemoption=i.itemoption"		
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "
	    
		'/결제내역기준
		if frectsellgubun = "S" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone z"
			sqlStr = sqlStr & " 	on d.zoneidx = z.idx"
		
		'/현재등록내역기준		
		elseif frectsellgubun = "N" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone_detail zi"
			sqlStr = sqlStr & " 	on m.shopid = zi.shopid"
			sqlStr = sqlStr & " 	and d.makerid = zi.makerid"
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_brand_zone z"
			sqlStr = sqlStr & " 	on zi.zoneidx = z.idx"
		end if

		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " group by d.itemgubun , d.itemid ,d.itemoption ,d.makerid, d.itemname" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemoptionname, d.zoneidx ,m.shopid" + vbcrlf
		sqlStr = sqlStr & " 	, z.zonename ,z.unit"
		sqlStr = sqlStr & " 	, cl.code_nm, cm.code_nm, cs.code_nm" + vbcrlf
		sqlStr = sqlStr & " 	,i.catecdl, i.catecdm, i.catecdn" + vbcrlf
		sqlStr = sqlStr & " order by d.zoneidx asc"
		
		Select Case FRectOrdertype
			Case "totalprice"
				'매출순
				sqlStr = sqlStr & " ,sellsum Desc"
	    	Case "gain"
	    		'수익순
	            sqlStr = sqlStr & " ,sum(d.realsellprice*d.itemno)-sum(d.suplyprice*d.itemno) Desc"
			Case "unitCost"
				'객단가순
				sqlStr = sqlStr & " ,sellsum Desc"

			Case "ea"
				'수량순
				sqlStr = sqlStr & " ,itemcnt Desc, sellsum desc"
				
			case else	
				if FRectCDL<>"" and FRectCDM<>"" then
					sqlStr = sqlStr + " ,i.catecdl asc ,i.catecdm asc ,i.catecdn asc"
				elseif FRectCDL<>"" then
					sqlStr = sqlStr + " , i.catecdl asc ,i.catecdm asc"
				else
					sqlStr = sqlStr + " ,i.catecdl asc"
				end if
		end Select
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new czone_oneitem

				FItemList(i).funit = rsget("unit")
				FItemList(i).frealpyeong = rsget("realpyeong")
				FItemList(i).fitemcnt = rsget("itemcnt")
				FItemList(i).fsellsum = rsget("sellsum")
				FItemList(i).fsuplypricesum = rsget("suplypricesum")
				FItemList(i).fcdl_nm = db2html(rsget("cdl_nm"))
				FItemList(i).fcdm_nm = db2html(rsget("cdm_nm"))
				FItemList(i).fcds_nm = db2html(rsget("cds_nm"))
				FItemList(i).fzoneidx = rsget("zoneidx")
				FItemList(i).fzonename = db2html(rsget("zonename"))
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fshopitemid = rsget("itemid")
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fmakerid = db2html(rsget("makerid"))
				FItemList(i).fitemname = db2html(rsget("itemname"))
				FItemList(i).fitemoptionname = db2html(rsget("itemoptionname"))
																						
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
		
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


Function zoneselectbox(shopid,boxname,selectedId,changeFlag)
	Dim vBody, vQuery
	If shopid <> "" Then
		vBody = vBody & "* 조닝 : <select name='"&boxname&"' " & changeFlag & " ><option value=''>-전체-</option>"
		
		vQuery = "SELECT idx, zonename FROM [db_shop].[dbo].[tbl_shop_brand_zone] WHERE shopid = '" & shopid & "' ORDER BY idx desc"
		rsget.Open vQuery,dbget,1
		If not rsget.EOF Then
			Do until rsget.EOF
				vBody = vBody & "<option value='" & rsget("idx") & "'"
				If CStr(selectedId) = CStr(rsget("idx")) Then
					vBody = vBody & " selected"
				End If
				vBody = vBody & ">" & rsget("zonename") & "</option>"
			rsget.MoveNext
			Loop
		End If
		rsget.close
		vBody = vBody & "</select>&nbsp;"
	End If
	Response.Write vBody
End Function

Sub drawSelectBoxOffShopOnlyZone(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" <%=changeFlag%>>
     <option value='' <%if selectedId="" then response.write " selected"%>>-매장 선택-</option><%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user"
   query1 = query1 & " where isusing='Y' "
   query1 = query1 & " and userid<>'streetshop000'"
   query1 = query1 & " and userid<>'streetshop800'"
   query1 = query1 & " and userid<>'streetshop870'"
   query1 = query1 & " and userid<>'streetshop700'"
   query1 = query1 & " order by isusing desc, convert(int,shopdiv)+10 asc, user asc"
   
   'response.write query1 &"<br>"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end sub

Function TrColor(no)
	Select Case no
		Case "1"
			TrColor = "#E6E6FA"
		Case "2"
			TrColor = "#B0C4DE"
		Case "3"
			TrColor = "#20B2AA"
		Case "4"
			TrColor = "#ADFF2F"
		Case "5"
			TrColor = "#BDB76B"
		Case "6"
			TrColor = "#DEB48C"
		Case "7"
			TrColor = "#F4A460"
		Case "8"
			TrColor = "#A0522D"
		Case "9"
			TrColor = "#BC8F8F"
		Case "10"
			TrColor = "#FF69B4"
		Case "11"
			TrColor = "#D8BFD8"
		Case "12"
			TrColor = "#EE82EE"
		Case "13"
			TrColor = "#9370DB"
		Case "14"
			TrColor = "#C0C0C0"
		Case "15"
			TrColor = "#FAF0E6"
		Case "16"
			TrColor = "#40E0D0"
		Case "17"
			TrColor = "#32CD32"
		Case "18"
			TrColor = "#F5DEB3"
		Case "19"
			TrColor = "#DA70D6"
		Case "20"
			TrColor = "#7B68EE"
		Case Else
			TrColor = "gray"
	End Select
End Function

'//상품 구역변경
function drawzonechange(selectBoxName,selectedId,shopid,changeFlag)
	dim query1 ,tmp_str
	
	if shopid = "" or isnull(shopid) then exit function
	%>
	<select name="<%=selectBoxName%>" <% if selectedId="" then response.write " selected"%> <%= changeFlag %>>
		<option value=''>구역지정</option>
   <%
	query1 = " select z.idx , z.zonename"
	query1 = query1 & " from db_shop.dbo.tbl_shop_brand_zone z"
	query1 = query1 & " where z.isusing = 'Y' and z.shopid = '"&shopid&"'"   
	query1 = query1 & " order by z.idx desc"
	
	rsget.Open query1,dbget,1
	
	if  not rsget.EOF  then
		do until rsget.EOF
			if Lcase(selectedId) = Lcase(rsget("idx")) then
				tmp_str = " selected"
			end if
			response.write "<option value='"&rsget("idx")&"' "&tmp_str&">" & db2html(rsget("zonename")) & "</option>"
			tmp_str = ""
			
			rsget.MoveNext
		loop
	end if
	rsget.close
	response.write "<option value='0'>현재조닝에서제외</option>"
	response.write "</select>"
end function

'//상품 구역변경
function drawnewipgobrand(shopid)
	dim query1 ,tmp_str
	
	if shopid = "" or isnull(shopid) then exit function

	query1 = " select top 500"
	query1 = query1 & " s.shopid, s.makerid, zd.zoneidx"
	query1 = query1 & " from db_shop.dbo.tbl_shop_designer s"
	query1 = query1 & " left join db_shop.dbo.tbl_shop_brand_zone_detail zd"
	query1 = query1 & " 	on s.shopid = zd.shopid"
	query1 = query1 & " 	and s.makerid = zd.makerid"
	query1 = query1 & " where s.shopid = '"&shopid&"'"   
	query1 = query1 & " and zd.zoneidx is null"
	query1 = query1 & " and s.regdate >= dateadd(m,-3 ,getdate())"
	query1 = query1 & " and s.firstipgodate is not null"
	query1 = query1 & " and itemcount <> 0"
	
	'response.write query1 &"<Br>"
	rsget.Open query1,dbget,1
	
	if not rsget.EOF  then
		do until rsget.EOF
			tmp_str = tmp_str & rsget("makerid") &"\n"
		rsget.MoveNext
		loop
	end if
	rsget.close
	
	if tmp_str <> "" then drawnewipgobrand = tmp_str
end function
%>