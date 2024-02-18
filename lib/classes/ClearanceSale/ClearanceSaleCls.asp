<%
'#############################################################
'	Description : 클리어런스 세일 어드민 클래스
'	History		: 2016.01.14 유태욱 생성
'#############################################################
%>
<%
class CClearanceItem
	public Fidx
	public FIsusing
	public Fdispcate1
	public FdispCateName
	public FdispCateNameReal
	public FItemid
	public FRegdate
	public Fsellyn
	public Flimityn
	public Flistimage
	public FMakerid
	public Fitemname
	public Fbasicimage
	public Fsaleyn
	public FmwDiv
	public FitemcouponYn
	public FitemcouponType
	public FitemcouponValue
	public FitemCouponBuyPrice
	public FsellCash
	public FbuyCash
	public ForgPrice
	public ForgSuplyCash
	public FsailPrice
	public FsailSuplyCash
end class

class CClaearanceitem
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectIsusing
	public FRectitemid
	public FRectMakerid
	public FRectItemName
	public FRectSellYN
	public FRectLimityn
	public FRectCatecode
	public FRectSaleYN
	public FRectItemcouponYN


	'###### 클리어런스세일 상품코드 리스트 ######
	public sub fnGetclaearanceitemList
		dim sqlStr,i, sqlsearch

		'' 브랜드명 검색
        if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and i.makerid = '"& FRectMakerid &"'"
        end if

		'' 상품코드 검색
        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

		'' 사용여부 검색
		if FRectIsusing <> "" Then
			sqlsearch = sqlsearch & " AND c.isusing ='"& FRectIsusing &"'"
		end if

		'' 카테고리 검색
		if FRectCatecode <> "" Then
			if FRectCatecode = "999" then
				sqlsearch = sqlsearch & " AND c.dispcate1 ='' "
			else
				sqlsearch = sqlsearch & " AND c.dispcate1 ='"& FRectCatecode &"'"
			end if
		end if

		'' 상품명 검색
        if (FRectItemName <> "") then
            sqlsearch = sqlsearch & " and i.itemname like '%"& FRectItemName &"%'"
        end if

		'' 판매여부 검색
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  sqlsearch = sqlsearch & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if

		'' 한정여부 검색
		if FRectLimityn="Y0" then
            sqlsearch = sqlsearch + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlsearch = sqlsearch + " and i.limityn='" + FRectLimityn + "'"
        end if
        
        ''할인여부
        if FRectSaleYN <> "" then
        	sqlsearch = sqlsearch + " and i.sailyn='" + FRectSaleYN + "'"
    	end if
        
        ''쿠폰여부
        if FRectItemcouponYN <> "" then
        	sqlsearch = sqlsearch + " and i.itemcouponyn='" + FRectItemcouponYN + "'"
    	end if
        
		'총 갯수 구하기
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_clearance_sale_item as c"
		sqlStr = sqlStr & " 	join db_item.dbo.tbl_item as i"
		sqlStr = sqlStr & " 	on i.itemid=c.itemid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'DB 데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " c.idx, c.itemid, c.isusing, c.regdate, i.sellyn, i.limityn, i.listimage, i.makerid, i.itemname, i.basicimage, c.dispcate1, db_item.dbo.getDisplayCateName(c.dispcate1) as dispcateNm, db_item.dbo.getDisplayCateName(i.dispcate1) as dispcateNmReal "
		sqlStr = sqlStr & " , i.sailyn, i.itemcouponyn, i.mwdiv, i.sellcash, i.buycash, i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash, i.itemcoupontype, i.itemcouponvalue "
		sqlStr = sqlStr & " ,  Case i.itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_clearance_sale_item as c with (noLock) "
		sqlStr = sqlStr & " 	join db_item.dbo.tbl_item as i with (noLock) "
		sqlStr = sqlStr & " 	on i.itemid=c.itemid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by c.idx Desc"
		
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
				set FItemList(i) = new CClearanceItem
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).Fitemid = rsget("itemid")
					FItemList(i).FIsusing = rsget("isusing")
					FItemList(i).Fdispcate1 = rsget("dispcate1")
					FItemList(i).FdispCateName = rsget("dispcateNm")
					FItemList(i).FdispCateNameReal = rsget("dispcateNmReal")
					FItemList(i).FRegdate = rsget("regdate")
					
					FItemList(i).Fsellyn = rsget("sellyn")
					FItemList(i).Flimityn = rsget("limityn")
					FItemList(i).FMakerid = rsget("makerid")
					FItemList(i).Fitemname = rsget("itemname")
					
					FItemList(i).Fsaleyn = rsget("sailyn")
					FItemList(i).FitemcouponYn = rsget("itemcouponyn")
					FItemList(i).FitemcouponType = rsget("itemcoupontype")
					FItemList(i).FitemcouponValue = rsget("itemcouponvalue")
					FItemList(i).FitemcouponBuyPrice = rsget("couponbuyprice")
					
					FItemList(i).FmwDiv = rsget("mwdiv")
					FItemList(i).FsellCash = rsget("sellcash")
					FItemList(i).FbuyCash = rsget("buycash")
					FItemList(i).ForgPrice = rsget("orgprice")
					FItemList(i).ForgSuplyCash = rsget("orgsuplycash")
					FItemList(i).FsailPrice = rsget("sailprice")
					FItemList(i).FsailSuplyCash = rsget("sailsuplycash")
					
					FItemList(i).Flistimage = "http://webimage.10x10.co.kr/image/list/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" &db2html(rsget("ListImage"))
					FItemList(i).Fbasicimage = "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" &db2html(rsget("basicimage"))
					
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
end class
%>






	

		