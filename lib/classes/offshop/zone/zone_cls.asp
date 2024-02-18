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
	public fzonegroup_name
	public fzonegroup_type
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
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'//admin/offshop/zone/zone_reg.asp
    public Sub fzone_oneitem()
        dim sqlStr , sqlsearch

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "&frectidx&""
		end if
        
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " idx,shopid,zonename,racktype,unit,regdate,isusing,zonegroup" + vbcrlf		
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_zone" + vbcrlf	
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new czone_oneitem
        
        if Not rsget.Eof then
    		    		
			FOneItem.fidx = rsget("idx")
			FOneItem.fshopid = rsget("shopid")
			FOneItem.fzonename = db2html(rsget("zonename"))
			FOneItem.fracktype = rsget("racktype")
			FOneItem.funit = rsget("unit")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fzonegroup = rsget("zonegroup")
						           
        end if
        rsget.Close
    end Sub
    
	'//admin/offshop/zone/zone.asp
	public sub fzone_list()
		dim sqlStr,i , sqlsearch
		
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and shopid = '"&frectshopid&"'"
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing = '"&frectisusing&"'"
		end if
			
		if frectzonegroup <> "" then
			sqlsearch = sqlsearch & " and zonegroup = "&frectzonegroup&""
		end if
		if frectracktype <> "" then
			sqlsearch = sqlsearch & " and racktype = "&frectracktype&""
		end if	
		
		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_zone" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub
					
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,shopid,zonegroup,racktype,zonename,unit,regdate,isusing" + vbcrlf		
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_zone" + vbcrlf	
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by shopid asc" + vbcrlf

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
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fzonegroup = rsget("zonegroup")
				FItemList(i).fracktype = rsget("racktype")				
				FItemList(i).fzonename = db2html(rsget("zonename"))
				FItemList(i).funit = rsget("unit")				
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")
		
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/검색조건이 바뀔경우, 이 클래스 뿐만 아니라, 구역 저장하는 부분도 수정해야 합니다. 검색값을 모두 저장하는 버튼이 있습니다
	'//admin/offshop/zone/zone_item.asp
	public sub GetoffshopzoneitemMatch()
		dim sqlStr,i , sqlsearch ,sqlsearch2

		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&frectshopid&"'"
		end if						
		if FRectmakerid<>"" then
			sqlsearch2 = sqlsearch2 + " and i.makerid='" + CStr(FRectmakerid) + "'"
		end if	
		if frectitemid<>"" then
			sqlsearch2 = sqlsearch2 + " and i.shopitemid=" + frectitemid + ""
		end if
		if frectitemname<>"" then
			sqlsearch2 = sqlsearch2 + " and i.shopitemname like '%" + frectitemname + "%'"
		end if
		if FRectCDL<>"" then
			sqlsearch2 = sqlsearch2 + " and i.catecdl='" + FRectCDL + "'"
		end if
		if FRectCDM<>"" then
			sqlsearch2 = sqlsearch2 + " and i.catecdm='" + FRectCDM + "'"
		end if
		if FRectCDN<>"" then
			sqlsearch2 = sqlsearch2 + " and i.catecdn='" + FRectCDN + "'"
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

		if frectisusing = "Y" then
			sqlsearch2 = sqlsearch2 & " and zi.zoneidx is not null"
		elseif frectisusing = "N" then
			sqlsearch2 = sqlsearch2 & " and zi.zoneidx is null"
		end if
		if frectzonegroup <> "" then
			sqlsearch2 = sqlsearch2 & " and z.zonegroup = "&frectzonegroup&""
		end if
		if frectracktype <> "" then
			sqlsearch2 = sqlsearch2 & " and z.racktype = "&frectracktype&""
		end if
		if frectsearchtype = "M" then
			sqlsearch2 = sqlsearch2 & " and t.shopid is not null"
		end if
			
		'총 갯수 구하기
		sqlStr = "select count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " left Join ("
		sqlStr = sqlStr & " 	select" + vbcrlf
		sqlStr = sqlStr & " 	m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " 	Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " 	group by m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " ) t" + vbcrlf		
		sqlStr = sqlStr + " 	on i.itemgubun=t.itemgubun"
		sqlStr = sqlStr + " 	and i.shopitemid=t.itemid"
		sqlStr = sqlStr + " 	and i.itemoption=t.itemoption"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
		sqlStr = sqlStr & " 	on t.shopid = zi.shopid" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemgubun = zi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemid = zi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemoption = zi.itemoption" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
		sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup"
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr & " where i.isusing='Y' " & sqlsearch2
	
		'response.write sqlStr &"<br>"			
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub
					
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " i.itemgubun ,i.shopitemid ,i.itemoption ,i.makerid ,i.shopitemname" + vbcrlf
		sqlStr = sqlStr & " ,i.shopitemoptionname,t.shopid ,zi.zoneidx" + vbcrlf
		sqlStr = sqlStr & " ,z.zonegroup ,z.racktype ,z.zonename,c.zonegroup_name" + vbcrlf
		sqlStr = sqlStr & " , cl.code_nm as cdl_nm, cm.code_nm as cdm_nm, cs.code_nm as cds_nm" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " left Join ("
		sqlStr = sqlStr & " 	select" + vbcrlf
		sqlStr = sqlStr & " 	m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " 	Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " 	group by m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " ) t" + vbcrlf		
		sqlStr = sqlStr + " 	on i.itemgubun=t.itemgubun"
		sqlStr = sqlStr + " 	and i.shopitemid=t.itemid"
		sqlStr = sqlStr + " 	and i.itemoption=t.itemoption"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
		sqlStr = sqlStr & " 	on t.shopid = zi.shopid" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemgubun = zi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemid = zi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemoption = zi.itemoption" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
		sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup"
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr & " where i.isusing='Y' " & sqlsearch2
		
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

				FItemList(i).fcdl_nm = db2html(rsget("cdl_nm"))
				FItemList(i).fcdm_nm = db2html(rsget("cdm_nm"))
				FItemList(i).fcds_nm = db2html(rsget("cds_nm"))								
				FItemList(i).fzonegroup_name = db2html(rsget("zonegroup_name"))				
				FItemList(i).fzonegroup = rsget("zonegroup")
				FItemList(i).fracktype = rsget("racktype")
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fshopitemid = rsget("shopitemid")
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fmakerid = db2html(rsget("makerid"))
				FItemList(i).fshopitemname = db2html(rsget("shopitemname"))
				FItemList(i).fshopitemoptionname = db2html(rsget("shopitemoptionname"))
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fzoneidx = rsget("zoneidx")
				FItemList(i).fzonename = db2html(rsget("zonename"))
						
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/offshop/zone/zone_sum.asp
	public sub Getoffshopzonesum()
		dim sqlStr,i , sqlsearch
		
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&frectshopid&"'"
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + CStr(FRectmakerid) + "'"
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
		
		sqlStr = "select"
		sqlStr = sqlStr & " sum(d.itemno) as itemcnt, sum(d.realsellprice*d.itemno) as sellsum"
		sqlStr = sqlStr & " ,z.idx ,z.zonegroup, z.racktype, z.zonename ,z.unit"
		sqlStr = sqlStr & " ,(sum(d.realsellprice*d.itemno) / z.unit) as unitvalue"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_shopjumun_detail d"
		sqlStr = sqlStr & " 	on m.orderno=d.orderno"
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " 	on d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + " 	and d.itemid=i.shopitemid"
		sqlStr = sqlStr + " 	and d.itemoption=i.itemoption"

		'/결제내역기준
		if frectsellgubun = "S" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z"
			sqlStr = sqlStr & " 	on d.zoneidx = z.idx"
		
		'/현재등록내역기준		
		elseif frectsellgubun = "N" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem zi"
			sqlStr = sqlStr & " 	on m.shopid = zi.shopid"
			sqlStr = sqlStr & " 	and d.itemid = zi.shopitemid"
			sqlStr = sqlStr & " 	and d.itemgubun = zi.itemgubun"
			sqlStr = sqlStr & " 	and d.itemoption = zi.itemoption"
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z"
			sqlStr = sqlStr & " 	on zi.zoneidx = z.idx"
		end if
		
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " group by z.idx ,z.zonename ,z.unit,z.zonegroup, z.racktype"
		sqlStr = sqlStr & " order by z.idx asc"
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.recordcount

		redim preserve FItemList(FTotalCount)
		
		i=0
		if  not rsget.EOF  then			
			do until rsget.EOF
				set FItemList(i) = new czone_oneitem

				FItemList(i).fzonegroup = rsget("zonegroup")
				FItemList(i).fracktype = rsget("racktype")				
				FItemList(i).funit = rsget("unit")
				FItemList(i).fitemcnt = rsget("itemcnt")
				FItemList(i).fsellsum = rsget("sellsum")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fzonename = db2html(rsget("zonename"))				
				FItemList(i).funitvalue = rsget("unitvalue")
				
				FCountTotal = FCountTotal + FItemList(i).fitemcnt
				FSumTotal = FSumTotal + FItemList(i).fsellsum
						
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/offshop/zone/zone_sum_detail.asp
	public sub Getoffshopzone_detail()
		dim sqlStr,i , sqlsearch
						
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
				
		'총 갯수 구하기
		sqlStr = "select count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " 	on d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + " 	and d.itemid=i.shopitemid"
		sqlStr = sqlStr + " 	and d.itemoption=i.itemoption"		

		'/결제내역기준
		if frectsellgubun = "S" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
			sqlStr = sqlStr & " 	on d.zoneidx = z.idx" + vbcrlf	
		
		'/현재등록내역기준
		elseif frectsellgubun = "N" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
			sqlStr = sqlStr & " 	on m.shopid = zi.shopid" + vbcrlf
			sqlStr = sqlStr & " 	and d.itemid = zi.shopitemid" + vbcrlf
			sqlStr = sqlStr & " 	and d.itemgubun = zi.itemgubun" + vbcrlf
			sqlStr = sqlStr & " 	and d.itemoption = zi.itemoption" + vbcrlf	
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
			sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		end if

		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup" 
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
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
		
		'response.write sqlStr &"<br>"			
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " d.itemgubun , d.itemid ,d.itemoption ,d.makerid, d.itemname" + vbcrlf
		sqlStr = sqlStr & " ,d.itemoptionname, d.sellprice, d.realsellprice, d.suplyprice, d.itemno ,d.orderno, d.zoneidx" + vbcrlf
		sqlStr = sqlStr & " ,z.zonegroup, z.racktype, z.zonename,c.zonegroup_name"
		sqlStr = sqlStr & " , cl.code_nm as cdl_nm, cm.code_nm as cdm_nm, cs.code_nm as cds_nm" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " 	on d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + " 	and d.itemid=i.shopitemid"
		sqlStr = sqlStr + " 	and d.itemoption=i.itemoption"		

		'/결제내역기준
		if frectsellgubun = "S" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
			sqlStr = sqlStr & " 	on d.zoneidx = z.idx" + vbcrlf	
		
		'/현재등록내역기준		
		elseif frectsellgubun = "N" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
			sqlStr = sqlStr & " 	on m.shopid = zi.shopid" + vbcrlf
			sqlStr = sqlStr & " 	and d.itemid = zi.shopitemid" + vbcrlf
			sqlStr = sqlStr & " 	and d.itemgubun = zi.itemgubun" + vbcrlf
			sqlStr = sqlStr & " 	and d.itemoption = zi.itemoption" + vbcrlf	
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
			sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		end if

		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup" 
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
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
				
				FItemList(i).fshopsuplycash = rsget("suplyprice")
				FItemList(i).fcdl_nm = db2html(rsget("cdl_nm"))
				FItemList(i).fcdm_nm = db2html(rsget("cdm_nm"))
				FItemList(i).fcds_nm = db2html(rsget("cds_nm"))
				FItemList(i).fzonegroup_name = db2html(rsget("zonegroup_name"))
				FItemList(i).forderno = rsget("orderno")
				FItemList(i).fitemno = rsget("itemno")
				FItemList(i).fzoneidx = rsget("zoneidx")
				FItemList(i).fzonegroup = rsget("zonegroup")
				FItemList(i).fracktype = rsget("racktype")
				FItemList(i).fzonename = db2html(rsget("zonename"))
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fshopitemid = rsget("itemid")
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fmakerid = db2html(rsget("makerid"))
				FItemList(i).fitemname = db2html(rsget("itemname"))
				FItemList(i).fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).fsellprice = rsget("sellprice")
				FItemList(i).frealsellprice = rsget("realsellprice")
																						
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/offshop/zone/zone_sum_detail.asp
	public sub Getoffshopzone_detailCategory()
		dim sqlStr,i , sqlsearch
	    						
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
		if frectitemid<>"" then
			sqlsearch = sqlsearch + " and d.itemid=" + frectitemid + ""
		end if
		if frectitemname<>"" then
			sqlsearch = sqlsearch + " and d.itemname like '%" + frectitemname + "%'"
		end if
		
		'데이터 리스트 
		sqlStr = "select"
		sqlStr = sqlStr & " sum(d.itemno) as itemnosum, sum(d.realsellprice*d.itemno) as realmaechul" + vbcrlf
		sqlStr = sqlStr & " ,sum(d.suplyprice*d.itemno) as suplymaechul" + vbcrlf

		if FRectCDL<>"" and FRectCDM<>"" then
			sqlStr = sqlStr + " ,i.catecdl,i.catecdm,i.catecdn,cs.code_nm"
		elseif FRectCDL<>"" then
			sqlStr = sqlStr + " ,i.catecdl,i.catecdm,cm.code_nm"
		else
			sqlStr = sqlStr + " ,i.catecdl,cl.code_nm"
		end if

		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " 	on d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + " 	and d.itemid=i.shopitemid"
		sqlStr = sqlStr + " 	and d.itemoption=i.itemoption"

		'/결제내역기준
		if frectsellgubun = "S" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
			sqlStr = sqlStr & " 	on d.zoneidx = z.idx" + vbcrlf	
		
		'/현재등록내역기준		
		elseif frectsellgubun = "N" then
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
			sqlStr = sqlStr & " 	on m.shopid = zi.shopid" + vbcrlf
			sqlStr = sqlStr & " 	and d.itemid = zi.shopitemid" + vbcrlf
			sqlStr = sqlStr & " 	and d.itemgubun = zi.itemgubun" + vbcrlf
			sqlStr = sqlStr & " 	and d.itemoption = zi.itemoption" + vbcrlf	
			sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
			sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
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
		
		if FRectCDL<>"" and FRectCDM<>"" then
			sqlStr = sqlStr + " group by i.catecdl,i.catecdm,i.catecdn,cs.code_nm"
		elseif FRectCDL<>"" then
			sqlStr = sqlStr + " group by i.catecdl,i.catecdm,cm.code_nm"
		else
			sqlStr = sqlStr + " group by i.catecdl,cl.code_nm"
		end if
		
		if FRectCDL<>"" and FRectCDM<>"" then
			sqlStr = sqlStr + " order by i.catecdl asc ,i.catecdm asc ,i.catecdn asc"
		elseif FRectCDL<>"" then
			sqlStr = sqlStr + " order by i.catecdl asc ,i.catecdm asc"
		else
			sqlStr = sqlStr + " order by i.catecdl asc"
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
							
				FItemList(i).fitemnosum	= rsget("itemnosum")
				FItemList(i).frealmaechul	= rsget("realmaechul")
				FItemList(i).fsuplymaechul	= rsget("suplymaechul")
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
				
	
				if IsNULL(FItemList(i).FCateName) then
					if IsNULL(FItemList(i).fcatecd) then FItemList(i).fcatecd=""
					FItemList(i).FCateName = "-"
				end if

				if Not IsNull(FItemList(i).frealmaechul) then
					maxt = MaxVal(maxt,FItemList(i).frealmaechul)
					maxc = MaxVal(maxc,FItemList(i).fitemnosum)
				end if

				FCountTotal = FCountTotal + FItemList(i).fitemnosum
				FSumTotal = FSumTotal + FItemList(i).frealmaechul
																								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
		
	'//admin/offshop/zone/zone_common.asp
	public sub Getoffshopzonecommon_detail()
        dim sqlStr , sqlsearch

		if frectzonegroup <> "" then
			sqlsearch = sqlsearch & " and zonegroup = "&frectzonegroup&""
		end if
        
        sqlStr = "select top 1" & vbcrlf			
		sqlStr = sqlStr & " zonegroup,zonegroup_name,zonegroup_type,isusing,regdate" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_zone_common" + vbcrlf
		sqlStr = sqlStr & " where zonegroup_type='GROUP' " & sqlsearch	

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new czone_oneitem
        
        if Not rsget.Eof then

			FOneItem.fzonegroup = rsget("zonegroup")
			FOneItem.fzonegroup_name = db2html(rsget("zonegroup_name"))
			FOneItem.fzonegroup_type = rsget("zonegroup_type")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fregdate = rsget("regdate")

        end if
        rsget.Close
    end Sub
    
	'//admin/offshop/zone/zone_common.asp
	public sub Getoffshopzonecommon_list()
		dim sqlStr,i
		
		'총 갯수 구하기
		sqlStr = "select count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_zone_common"
		sqlStr = sqlStr & " where zonegroup_type='GROUP'"
		
		'response.write sqlStr &"<br>"			
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub
			
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " zonegroup,zonegroup_name,zonegroup_type,isusing,regdate" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_zone_common"
		sqlStr = sqlStr & " where zonegroup_type='GROUP'"
		sqlStr = sqlStr & " order by zonegroup desc"

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
				
				FItemList(i).fzonegroup = rsget("zonegroup")
				FItemList(i).fzonegroup_name = db2html(rsget("zonegroup_name"))
				FItemList(i).fzonegroup_type = rsget("zonegroup_type")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = rsget("regdate")
																						
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

'//매장내 그룹 셀렉트 박스
function drawSelectBoxOffShopzonegroup(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
	query1 = "select"   
	query1 = query1 & " zonegroup,zonegroup_name,zonegroup_type,isusing,regdate"
	query1 = query1 & " from db_shop.dbo.tbl_shop_zone_common"
	query1 = query1 & " where isusing='Y' and zonegroup_type = 'GROUP'"
	query1 = query1 & " order by zonegroup desc"
	
	rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("zonegroup")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("zonegroup")&"' "&tmp_str&">" + db2html(rsget("zonegroup_name")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

'//매장내 그룹 카테고리
function getOffShopzonegroup(stats)
	dim query1 ,tmp_str

	query1 = "select"   
	query1 = query1 & " zonegroup,zonegroup_name,zonegroup_type,isusing,regdate"
	query1 = query1 & " from db_shop.dbo.tbl_shop_zone_common"
	query1 = query1 & " where isusing='Y' and zonegroup_type = 'GROUP'"
	query1 = query1 & " order by zonegroup desc"
	
	rsget.Open query1,dbget,1
	
	if  not rsget.EOF  then
		do until rsget.EOF
			if Lcase(stats) = Lcase(rsget("zonegroup")) then
			   getOffShopzonegroup = db2html(rsget("zonegroup_name"))
			end if			
			
			rsget.MoveNext
		loop
	end if
	rsget.close	
end function

function drawSelectBoxOffShopracktype(boxname,stats,changeflg)
%>
	<select name="<%=boxname%>" <%=changeflg%>>
		<option value="" <% if stats = "" then response.write " selected"%>>선택</option>
		<option value="0" <% if stats = "0" then response.write " selected"%>>벽매대</option>
		<option value="1" <% if stats = "1" then response.write " selected"%>>아일랜드매대</option>
		<option value="2" <% if stats = "2" then response.write " selected"%>>이벤트매대</option>
		<option value="99" <% if stats = "99" then response.write " selected"%>>기타</option>
	</select>
<%
end function

function getOffShopracktype(stats)
	if stats = "0" then
		getOffShopracktype = "벽매대"
	elseif stats = "1" then
		getOffShopracktype = "아일랜드매대"
	elseif stats = "2" then
		getOffShopracktype = "이벤트매대"
	elseif stats = "99" then
		getOffShopracktype = "기타"			
	end if
end function

'//상품 구역변경
function drawzonechange(selectBoxName,shopid,changeFlag)
	dim query1
	
	if shopid = "" or isnull(shopid) then exit function
	%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value=''>구역지정</option>		
   <%
	query1 = " select z.idx ,z.zonegroup,z.racktype,z.zonename,c.zonegroup_name"
	query1 = query1 & " from db_shop.dbo.tbl_shop_zone z"
	query1 = query1 & " join db_shop.dbo.tbl_shop_zone_common c"
	query1 = query1 & " on z.zonegroup = c.zonegroup" 
	query1 = query1 & " and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
	query1 = query1 & " where z.isusing = 'Y' and z.shopid = '"&shopid&"'"   
	query1 = query1 & " order by z.zonegroup , z.racktype , z.zonename asc"
	
	rsget.Open query1,dbget,1
	
	if  not rsget.EOF  then
		do until rsget.EOF
			response.write("<option value='"&rsget("idx")&"'>["& db2html(rsget("zonegroup_name"))&" , "& getOffShopracktype(rsget("racktype"))&"] " & db2html(rsget("zonename")) & "</option>")
			rsget.MoveNext
		loop
	end if
	rsget.close
	response.write "<option value='0'>구역지정안함(삭제)</option>"
	response.write("</select>")
end function
%>