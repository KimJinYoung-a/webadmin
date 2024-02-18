<%
'####################################################
' Description :  오프라인 MD별 매출 클래스
' History : 2012.05.10 강준구 생성(기존매뉴 이전생성)
'			2013.01.24 한용민 수정
'####################################################

Class COffShopSellByTermMD
	public FCount
	public FMakerid
	public FSum
	public fsuplyprice
	public fprofit
	public FaddTaxChargeSum
	public fIXyyyymmdd
	public FChargeDiv
	public Fmdname

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopSellReportMD
	public FItemList()
	public FCountList()
	public FPageCount
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectShopID
	public FRectNormalOnly
	public FRectStartDay
	public FRectEndDay
	public FRectOffgubun
	public FRectOldData
	public frectmakerid
	public frectdatefg
	public frectdategubun
	public frectoffcatecode
	public frectoffmduserid
	public FRectOnlyShop
	public FRectInc3pl

	'//admin/offshop/sellreport_md.asp
	public Sub GetMDSellSumList()
		dim i,sqlStr ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		sqlsearch = sqlsearch + " and p.offmduserid is not null and p.offmduserid <> '' "

		if FRectNormalOnly="on" then
			sqlsearch = sqlsearch + " and m.cancelyn='N'"
			sqlsearch = sqlsearch + " and d.cancelyn='N'"
		end if

		if (FRectOffgubun<>"") then
		    if (FRectOffgubun="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1','2')"
		    elseif (FRectOffgubun="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3','4')"
		    elseif (FRectOffgubun="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5','6')"
		    elseif (FRectOffgubun="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7','8')"
		    elseif (FRectOffgubun="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if

		if FRectOnlyShop<>"" then
			sqlsearch = sqlsearch + " and Left(m.shopid,4)<>'cafe'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
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
		
		If frectoffcatecode <> "" Then
			sqlsearch = sqlsearch + " and p.offcatecode = '" + CStr(frectoffcatecode) + "' "
		End IF
		
		If frectoffmduserid <> "" Then
			sqlsearch = sqlsearch + " and p.offmduserid = '" + CStr(frectoffmduserid) + "' "
		End IF
		

		sqlStr = " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " (select top 1 username from [db_partner].[dbo].tbl_user_tenbyten where userid = p.offmduserid) as mdname, "
		sqlStr = sqlStr + " sum(d.itemno * d.realsellprice) as subtotal"
		sqlStr = sqlStr + " , sum(d.itemno * d.addTaxCharge) as addTaxChargeSum, d.makerid"
		'sqlStr = sqlStr + " , sum(d.itemno) as cnt "
		sqlStr = sqlStr + " ,sum(d.itemno * d.suplyprice) as suplyprice"
		'sqlStr = sqlStr + " ,sum(d.realsellprice*d.itemno-d.suplyprice*d.itemno) as profit"

		if FRectShopid<>"" then
			sqlStr = sqlStr + " ,s.chargediv"
		end if

		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m"
			sqlStr = sqlStr + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d"
			sqlStr = sqlStr + " 	on m.orderno = d.orderno"
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " 	on m.orderno = d.orderno"
		end if

		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " 	on m.shopid = u.userid"
		sqlStr = sqlStr + " join [db_partner].[dbo].tbl_partner p on d.makerid = p.id "

		if FRectShopid<>"" then
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s"
			sqlStr = sqlStr + " 	on s.shopid='" + FRectShopid + "'"
			sqlStr = sqlStr + " 	and d.makerid=s.makerid"
		end if

		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner pp"
	    sqlStr = sqlStr + "       on m.shopid=pp.id "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " group by p.offmduserid, d.makerid "

		if FRectShopid<>"" then
			sqlStr = sqlStr + " ,s.chargediv"
		end if

		sqlStr = sqlStr + " order by"

		sqlStr = sqlStr + " p.offmduserid asc, subtotal desc "

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopSellByTermMD

				FItemList(i).Fmdname = CHKIIF(isNull(rsget("mdname")),"&nbsp;",rsget("mdname"))
				FItemList(i).FMakerid  = rsget("makerid")
				'FItemList(i).FCount = rsget("cnt")
				FItemList(i).FSum   = rsget("subtotal")
				FItemList(i).fsuplyprice  = rsget("suplyprice")
				'FItemList(i).fprofit  = rsget("profit")
				FItemList(i).FaddTaxChargeSum  = rsget("addTaxChargeSum")


				if FRectShopid<>"" then
					FItemList(i).FChargeDiv = rsget("chargediv")
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub
	
	
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		redim  FCountList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
end Class
%>