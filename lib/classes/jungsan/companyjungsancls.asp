<%
class CJungsanItem
	public FSitename
	public FOrderserial
	public FJungsanOk
	public FUserid
	public FBuyName
	public FSubTotalPrice
	public FBeasongPay
	public FDeasangPay
	public Fjungsansum

	public FIpkumDiv
	public FRegdate
	public FBeaDalDate

	public FCancelyn
	public FrectSiteName
	public FRectRegStart
	public FRectRegEnd
	public FTotaldate
	public FTotaldate2
	public FTotalno
	public FTotaldeasang
	public FTotaljungsansum
	public FIpkumDate
	public FTotalJungsan
	public FTotalBaesong
	public Fmasterid

	public Fjungsantitle
	public Fjungsantype
	public Fetcminussum
	public Fetcplussum
	public Frealipkumsum
	public Fsegumil

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CUpcheJungSan
	public FJungSanList()
	public FCommission
	public FEtcStr

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FrectSiteName
	public FRectRdSite

	public FRectRegStart
	public FRectRegEnd
	public FTotaldate
	public FTotaldate2
	public FTotalno
	public FTotaldeasang
	public FTotaljungsansum
	public FIpkumDiv
	public FIpkumDate

	public FRectYYYY
	public FRectMM

	public FTotalJungsan
	public FTotalBaesong
	public Fmasterid

	public sub getDefaultInfo(byval isitename)
		dim sqlStr
		sqlStr = " select top 1 commission from [db_partner].[dbo].tbl_partner"
		sqlStr = sqlStr + " where id='" + isitename + "'"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			FCommission = rsget("commission")
		end if

		rsget.close

	end sub

	public sub getOldDefaultInfo(byval masterid)
		dim sqlStr
		sqlStr = " select top 1 comission, etcstr from [db_jungsan].[dbo].tbl_etcsite_jungsanmaster"
		sqlStr = sqlStr + " where id=" + CStr(masterid) + ""
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			FCommission = rsget("comission")
			FEtcStr = Db2Html(rsget("etcstr"))
		end if

		rsget.close

	end sub

	public sub JungSanDeasangList()
		dim sqlStr, i
		sqlstr = " select m.orderserial, m.userid, m.buyname,"
		sqlstr = sqlstr + " m.subtotalprice, m.ipkumdiv, m.regdate, IsNull(d.itemcost,0) as beasongpay "
		sqlstr = sqlstr + " from [db_order].[dbo].tbl_order_master m "
		sqlstr = sqlstr + " left join [db_jungsan].[dbo].tbl_etcsite_jungsandetail j"
		sqlstr = sqlstr + " on m.orderserial=j.orderserial"
		sqlstr = sqlstr + " left join [db_order].[dbo].tbl_order_detail d "
		sqlstr = sqlstr + " on m.orderserial=d.orderserial and d.itemid=0 and d.cancelyn<>'Y'"
		sqlstr = sqlstr + " where m.regdate>'2006-07-01'"
		sqlstr = sqlstr + " and ((m.sitename='" + FrectSiteName + "') or (m.rdsite='" + FrectSiteName + "'))"
		sqlstr = sqlstr + " and m.ipkumdiv>=7"
		sqlstr = sqlstr + " and m.cancelyn='N'"
		sqlstr = sqlstr + " and m.regdate<'" + FRectRegEnd + "'"
		sqlstr = sqlstr + " and j.orderserial is Null"
		sqlstr = sqlstr + " order by m.idx desc"

		rsget.PageSize = FPageSize
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount


		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FJungSanList(FResultCount)


		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
				set FJungSanList(i) = new CJungsanItem
				FJungSanList(i).FSitename = FrectSiteName
				FJungSanList(i).FOrderserial = rsget("orderserial")
				FJungSanList(i).FJungsanOk   = "N"
				FJungSanList(i).FUserid      = rsget("userid")
				FJungSanList(i).FBuyName     = rsget("buyname")
				FJungSanList(i).FSubTotalPrice= rsget("subtotalprice")
				FJungSanList(i).FBeasongPay   = rsget("beasongpay")
				FJungSanList(i).FDeasangPay   = FJungSanList(i).FSubTotalPrice - FJungSanList(i).FBeasongPay
				FJungSanList(i).FIpkumDiv     = rsget("ipkumdiv")
				FJungSanList(i).FRegdate      = rsget("regdate")
				FJungSanList(i).FCancelyn	  = "N"
				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close
	end sub

    public sub JungSanDeasangList_OLD()
		dim sqlStr, i
		sqlstr = " select m.orderserial, m.userid, m.buyname,"
		sqlstr = sqlstr + " m.subtotalprice, m.ipkumdiv, m.regdate, IsNull(d.itemcost,0) as beasongpay "
		sqlstr = sqlstr + " from [db_log].[dbo].tbl_old_order_master_2003 m "
		sqlstr = sqlstr + " left join [db_jungsan].[dbo].tbl_etcsite_jungsandetail j"
		sqlstr = sqlstr + " on m.orderserial=j.orderserial"
		sqlstr = sqlstr + " left join [db_log].[dbo].tbl_old_order_detail_2003 d "
		sqlstr = sqlstr + " on m.orderserial=d.orderserial and d.itemid=0 and d.cancelyn<>'Y'"
		sqlstr = sqlstr + " where m.regdate>'2006-01-01'"
		sqlstr = sqlstr + " and ((m.sitename='" + FrectSiteName + "') or (m.rdsite='" + FrectSiteName + "'))"
		sqlstr = sqlstr + " and m.ipkumdiv>=7"
		sqlstr = sqlstr + " and m.cancelyn='N'"
		sqlstr = sqlstr + " and m.regdate<'" + FRectRegEnd + "'"
		sqlstr = sqlstr + " and j.orderserial is Null"
		sqlstr = sqlstr + " order by m.idx desc"

		rsget.PageSize = FPageSize
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount


		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FJungSanList(FResultCount)


		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
				set FJungSanList(i) = new CJungsanItem
				FJungSanList(i).FSitename = FrectSiteName
				FJungSanList(i).FOrderserial = rsget("orderserial")
				FJungSanList(i).FJungsanOk   = "N"
				FJungSanList(i).FUserid      = rsget("userid")
				FJungSanList(i).FBuyName     = rsget("buyname")
				FJungSanList(i).FSubTotalPrice= rsget("subtotalprice")
				FJungSanList(i).FBeasongPay   = rsget("beasongpay")
				FJungSanList(i).FDeasangPay   = FJungSanList(i).FSubTotalPrice - FJungSanList(i).FBeasongPay
				FJungSanList(i).FIpkumDiv     = rsget("ipkumdiv")
				FJungSanList(i).FRegdate      = rsget("regdate")
				FJungSanList(i).FCancelyn	  = "N"
				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close
	end sub
	
	public sub PartnerMiJungSanDeasangList()
		dim sqlStr, i

		''#################################################
		''ÃÑ °¹¼ö. ÃÑ±Ý¾×
		''#################################################
		sqlstr = "select count(m.orderserial) as cnt, sum(m.subtotalprice)as totalprice, sum(IsNull(d.itemcost,0)) as beasongtotal"
		sqlstr = sqlstr + " from [db_order].[dbo].tbl_order_master m "
		sqlstr = sqlstr + " left join [db_jungsan].[dbo].tbl_etcsite_jungsandetail j"
		sqlstr = sqlstr + " on m.orderserial=j.orderserial"
		sqlstr = sqlstr + " left join [db_order].[dbo].tbl_order_detail d "
		sqlstr = sqlstr + " on m.orderserial=d.orderserial and d.itemid=0 and d.cancelyn<>'Y'"

		if FRectRdSite<>"" then
			sqlstr = sqlstr + " where m.rdsite='" + FRectRdSite + "'"
		else
			sqlstr = sqlstr + " where m.sitename='" + FrectSiteName + "'"
		end if

		sqlstr = sqlstr + " and m.ipkumdiv>='7'"
		sqlstr = sqlstr + " and m.cancelyn='N'"
		sqlstr = sqlstr + " and j.orderserial is Null"

		if FRectYYYY<>"" then
			sqlstr = sqlstr + " and year(m.regdate)='" + FRectYYYY + "'"
		end if

		if FRectMM<>"" then
			sqlstr = sqlstr + " and month(m.regdate)='" + FRectMM + "'"
		end if

		rsget.Open sqlStr,dbget,1

		FTotalJungsan = rsget("totalprice")
        FTotalBaesong = rsget("beasongtotal")

			if IsNull(FTotalJungsan) then FTotalJungsan=0
			if IsNull(FTotalBaesong) then FTotalBaesong=0

		rsget.Close


		''#################################################
		''µ¥ÀÌÅ¸.
		''#################################################

		sqlstr = " select m.orderserial, m.userid, m.buyname,"
		sqlstr = sqlstr + " m.subtotalprice, m.ipkumdiv, m.regdate, m.beadaldate, IsNull(d.itemcost,0) as beasongpay "
		sqlstr = sqlstr + " from [db_order].[dbo].tbl_order_master m "
		sqlstr = sqlstr + " left join [db_jungsan].[dbo].tbl_etcsite_jungsandetail j"
		sqlstr = sqlstr + " on m.orderserial=j.orderserial"
		sqlstr = sqlstr + " left join [db_order].[dbo].tbl_order_detail d "
		sqlstr = sqlstr + " on m.orderserial=d.orderserial and d.itemid=0 and d.cancelyn<>'Y'"

		if FRectRdSite<>"" then
			sqlstr = sqlstr + " where m.rdsite='" + FRectRdSite + "'"
		else
			sqlstr = sqlstr + " where m.sitename='" + FrectSiteName + "'"
		end if

		sqlstr = sqlstr + " and m.ipkumdiv>='7'"
		sqlstr = sqlstr + " and m.cancelyn='N'"
		sqlstr = sqlstr + " and j.orderserial is Null"
		if FRectYYYY<>"" then
			sqlstr = sqlstr + " and year(m.regdate)='" + FRectYYYY + "'"
		end if

		if FRectMM<>"" then
			sqlstr = sqlstr + " and month(m.regdate)='" + FRectMM + "'"
		end if

		sqlstr = sqlstr + " order by regdate desc"

		rsget.PageSize = FPageSize
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount


		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FJungSanList(FResultCount)


		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
				set FJungSanList(i) = new CJungsanItem
				FJungSanList(i).FOrderserial = rsget("orderserial")
				FJungSanList(i).FUserid      = rsget("userid")
				FJungSanList(i).FBuyName     = rsget("buyname")
				FJungSanList(i).FSubTotalPrice= rsget("subtotalprice")
				FJungSanList(i).FBeasongPay   = rsget("beasongpay")
				FJungSanList(i).FDeasangPay   = FJungSanList(i).FSubTotalPrice - FJungSanList(i).FBeasongPay
				FJungSanList(i).FRegdate      = rsget("regdate")
				FJungSanList(i).FBeaDalDate      = rsget("beadaldate")

				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close
	end sub

	public sub PartnerOldJungSanDeasangList()
		dim sqlStr, i

		''#################################################
		''ÃÑ °¹¼ö. ÃÑ±Ý¾×
		''#################################################


		sqlstr = " select count(id) as cnt, sum(totaldeasang) as totalprice"
		sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_etcsite_jungsanmaster"
		sqlstr = sqlstr + " where sitename='" + FrectSiteName + "'"

		rsget.Open sqlStr,dbget,1

		FTotalJungsan = rsget("totalprice")

			if IsNull(FTotalJungsan) then FTotalJungsan=0

		rsget.Close


		''#################################################
		''µ¥ÀÌÅ¸.
		''#################################################

		sqlstr = " select top 100 * "
		sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_etcsite_jungsanmaster"
		sqlstr = sqlstr + " where sitename='" + FrectSiteName + "'"
		sqlstr = sqlstr + " order by id desc"

		rsget.PageSize = FPageSize
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount


		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FJungSanList(FResultCount)


		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
				set FJungSanList(i) = new CJungsanItem
				FJungSanList(i).Fmasterid = rsget("id")
 				FJungSanList(i).FTotaldate = rsget("startdate")
                FJungSanList(i).FTotaldate2 = rsget("enddate")
				FJungSanList(i).FTotalno = rsget("totalno")
				FJungSanList(i).FSubTotalPrice= rsget("totalsum")
				FJungSanList(i).FBeasongPay   = rsget("totalbeasongpay")
				FJungSanList(i).FTotaldeasang   = rsget("totaldeasang")
				FJungSanList(i).FTotaljungsansum   = rsget("totaljungsansum")
				FJungSanList(i).FIpkumDiv     = rsget("ipkumdiv")
				FJungSanList(i).FIpkumDate     = rsget("ipkumdate")

				FJungSanList(i).Fsegumil     = rsget("segumil")
				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close
	end sub

	public sub PartnerSiteJungSanDeasangList()
		dim sqlStr, i

		''#################################################
		''ÃÑ °¹¼ö. ÃÑ±Ý¾×
		''#################################################


		sqlstr = " select count(id) as cnt, sum(totaldeasang) as totalprice"
		sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_etcsite_jungsanmaster"
		sqlstr = sqlstr + " where sitename='" + FrectSiteName + "'"

		rsget.Open sqlStr,dbget,1

		FTotalJungsan = rsget("totalprice")

			if IsNull(FTotalJungsan) then FTotalJungsan=0

		rsget.Close


		''#################################################
		''µ¥ÀÌÅ¸.
		''#################################################

		sqlstr = " select top " + CStr(FPageSize*FCurrpage) + " * "
		sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_etcsite_jungsanmaster"
		sqlstr = sqlstr + " where regdate>'" + FRectRegStart + "'"
		sqlstr = sqlstr + " and regdate<'" + FRectRegEnd + "'"
		sqlstr = sqlstr + " order by sitename desc, regdate desc"

'response.write sqlstr

		rsget.PageSize = FPageSize
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount


		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FJungSanList(FResultCount)


		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
				set FJungSanList(i) = new CJungsanItem
				FJungSanList(i).Fmasterid = rsget("id")
				FJungSanList(i).Fsitename = rsget("sitename")
 				FJungSanList(i).FTotaldate = rsget("startdate")
                FJungSanList(i).FTotaldate2 = rsget("enddate")
				FJungSanList(i).FTotalno = rsget("totalno")
				FJungSanList(i).FSubTotalPrice= rsget("totalsum")
				FJungSanList(i).FBeasongPay   = rsget("totalbeasongpay")
				FJungSanList(i).FTotaldeasang   = rsget("totaldeasang")
				FJungSanList(i).FTotaljungsansum   = rsget("totaljungsansum")
				FJungSanList(i).FIpkumDiv     = rsget("ipkumdiv")
				FJungSanList(i).FIpkumDate     = rsget("ipkumdate")

				FJungSanList(i).Fjungsantitle   = rsget("jungsantitle")
				FJungSanList(i).Fjungsantype    = rsget("jungsantype")
				FJungSanList(i).Fetcminussum    = rsget("etcminussum")
				FJungSanList(i).Fetcplussum     = rsget("etcplussum")
				FJungSanList(i).Frealipkumsum   = rsget("realipkumsum")
				FJungSanList(i).Fsegumil      = rsget("segumil")

				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close
	end sub


	public sub PartnerOldDetailJungSanDeasangList()
		dim sqlStr, i

		''#################################################
		''ÃÑ °¹¼ö. ÃÑ±Ý¾×
		''#################################################


		sqlstr = " select count(orderserial) as cnt, sum(deasangsum) as totalprice, sum(jungsansum) as totaljungsansum"
		sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_etcsite_jungsandetail "
		sqlstr = sqlstr + " where masterid='" + Cstr(Fmasterid) + "'"
		sqlstr = sqlstr + " and cancelyn='N'"

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
			FTotalJungsan = rsget("totalprice")
			FTotalJungsansum = rsget("totaljungsansum")

			if IsNull(FTotalJungsan) then FTotalJungsan=0
			if IsNull(FTotalJungsansum) then FTotalJungsansum=0
		end if
		rsget.Close


		''#################################################
		''µ¥ÀÌÅ¸.
		''#################################################


		sqlstr = " select orderserial, userid, buyname,"
		sqlstr = sqlstr + " totalsum, deasangsum, beasongpay, jungsansum"
		sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_etcsite_jungsandetail "
		sqlstr = sqlstr + " where masterid='" + Cstr(Fmasterid) + "'"
		sqlstr = sqlstr + " and cancelyn='N'"
		sqlstr = sqlstr + " order by orderserial desc"

		rsget.PageSize = FPageSize
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount


		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FJungSanList(FResultCount)


		i=0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
				set FJungSanList(i) = new CJungsanItem
				FJungSanList(i).FOrderserial = rsget("orderserial")
				FJungSanList(i).FUserid      = rsget("userid")
				FJungSanList(i).FBuyName     = rsget("buyname")
				FJungSanList(i).FSubTotalPrice= rsget("totalsum")
				FJungSanList(i).FBeasongPay   = rsget("beasongpay")
				FJungSanList(i).FDeasangPay   = rsget("deasangsum")
				FJungSanList(i).Fjungsansum   = rsget("jungsansum")

				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close
	end sub


	public sub PartnerOldJungSanDeasangUpdate()
		dim sqlStr

		''#################################################
		''µ¥ÀÌÅ¸.
		''#################################################


		sqlstr = "update [db_jungsan].[dbo].tbl_etcsite_jungsanmaster"
        sqlstr = sqlstr + " set ipkumdiv = ipkumdiv + 1,"
        sqlstr = sqlstr + " ipkumdate = '" + FRectRegStart + "'"
        sqlstr = sqlstr + " where id = '" + CStr(Fmasterid) + "'"

        rsget.Open sqlstr, dbget, 1

	end sub

	public sub PartnerXLOldJungSanDeasangList()
		dim sqlStr, i

		''#################################################
		''µ¥ÀÌÅ¸.
		''#################################################


		sqlstr = " select orderserial, userid, buyname,"
		sqlstr = sqlstr + " totalsum, deasangsum, beasongpay,jungsansum"
		sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_etcsite_jungsandetail "
		sqlstr = sqlstr + " where masterid='" + Cstr(Fmasterid) + "'"
		sqlstr = sqlstr + " and cancelyn='N'"
		sqlstr = sqlstr + " order by orderserial desc"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FJungSanList(FResultCount)

		i=0
		if not rsget.EOF then
			do until (i >= FResultCount)
				set FJungSanList(i) = new CJungsanItem
				FJungSanList(i).FOrderserial = rsget("orderserial")
				FJungSanList(i).FUserid      = rsget("userid")
				FJungSanList(i).FBuyName     = rsget("buyname")
				FJungSanList(i).FSubTotalPrice= rsget("totalsum")
				FJungSanList(i).FBeasongPay   = rsget("beasongpay")
				FJungSanList(i).FDeasangPay   = rsget("deasangsum")
				FJungSanList(i).Fjungsansum   = rsget("jungsansum")
				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
		'redim preserve FJungSanList(0)
		redim  FJungSanList(0)

		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		FCommission =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class

%>