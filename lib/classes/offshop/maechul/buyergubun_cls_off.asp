<%
'####################################################
' Description :  오프라인 외국인구매통계 클래스
' History : 2013.02.20 한용민 생성
'####################################################

class cbuyeroneitem
	public fshopid
	public fshopname
	public fbuyergubun
	public fcodename
	public fcnt
	public ftotalsum
	public frealsum
	public fspendmile
				
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class cbuyerlist
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount
	
	public maxt
	public maxc
	public FRectStartDay
	public FRectEndDay
	public frectoffgubun
	public FRectOldData
	public frectdatefg
	public frectbuyergubun
	public frectshopid
	public FRectInc3pl
	
	'/common/offshop/maechul/buyergubun_maechulsum_off.asp
	public sub getbuyergubun_list
		dim sqlStr, i, sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
	    
		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
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
		
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if
		
		if frectshopid <> "" then
			sqlsearch = sqlsearch + " and m.shopid = '"&frectshopid&"'"
		end if
		
		sqlStr = " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.shopid, u.shopname"
		sqlStr = sqlStr & " ,isnull(m.buyergubun,-1) as buyergubun, c.codename"
		sqlStr = sqlStr & " ,count(m.idx) as cnt"
		sqlStr = sqlStr & " ,sum(m.totalsum) as totalsum"
		sqlStr = sqlStr & " ,sum(m.realsum) as realsum"
		sqlStr = sqlStr & " ,sum(m.spendmile) as spendmile"

		if FRectOldData="on" then
			sqlStr = sqlStr & " from [db_shoplog].[dbo].tbl_old_shopjumun_master m"
		else
			sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shopjumun_master m"
		end if

		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr & " 	on m.shopid=u.userid"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_offshop_commoncode c"
		sqlStr = sqlStr & " 	on isnull(m.buyergubun,-1)=c.codeid"
		sqlStr = sqlStr & " 	and c.codekind='buyergubun'"
		sqlStr = sqlStr & " 	and c.codegroup='MAIN'"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "		
		sqlStr = sqlStr & " where m.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " group by isnull(m.buyergubun,-1), c.codename, m.shopid, u.shopname, u.shopdiv, c.orderno"
		sqlStr = sqlStr & " order by convert(int,u.shopdiv)+10 asc, m.shopid asc, c.orderno asc"
		
		'response.write sqlStr &"<br>"		
		rsget.Open sqlStr,dbget,1
		
		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount
		
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new cbuyeroneitem

				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fshopname = rsget("shopname")
				FItemList(i).fbuyergubun = rsget("buyergubun")
				FItemList(i).fcodename = rsget("codename")
				FItemList(i).fcnt = rsget("cnt")
				FItemList(i).ftotalsum = rsget("totalsum")
				FItemList(i).frealsum = rsget("realsum")
				FItemList(i).fspendmile = rsget("spendmile")
				
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub
	
	Private Sub Class_Initialize()
		redim  FItemList(0)
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
%>
