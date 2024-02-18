<%
class AuctionItem
	private FGetCount
	private FAuctionID()
	private FAuctionType()
	private FAuctionName()
	private FItemID()
	private Flimitno()
	private FStartDate()
	private FFinishDate()
	private FRegDate()
	private FSupplyer()
	private FForceFinish()
	private FItemName()
	private FImageList()
	private FImageSmall()
	private FLinkItem()

	private FPriceStart()
	private FPriceEnd()
	private FPriceFix()

	private FTotalCount
	private FCurrPage
	private FTotalPage
	private FPageSize
	private FResultCount
	private FScrollCount

	Property Get TotalCount()
		TotalCount = FTotalCount
	end Property

	Property Get TotalPage()
		TotalPage = FTotalPage
	end Property

	Property Get CurrPage()
		CurrPage = FCurrPage
	end Property

	Property Let CurrPage(byVal v)
		FCurrPage = v
	end Property

	Property Get PageSize()
		PageSize = FPageSize
	end Property

	Property Let PageSize(byVal v)
		FPageSize = v
	end Property

	Property Get ResultCount()
		ResultCount = FResultCount
	end Property

	Property Get ScrollCount()
		ScrollCount = FScrollCount
	end Property

	Property Let ScrollCount(byVal v)
		FScrollCount = v
	end Property

	Private Sub Class_Initialize()
		FGetCount = 3

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = TotalPage > StarScrollPage + ScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((Currpage-1)\ScrollCount)*ScrollCount +1
	end Function



	Public Sub SetAuctionCount(byVal v)
		FGetCount = v
	end Sub

	Public Sub GetAllAuction()
		dim sqlStr
		sqlStr = " select count(id) as cnt from tbl_board_auction"


	end sub

	Public Sub GetOldAuction(byval userid,byval notthisitem)
		dim sqlStr
		sqlStr = " select a.id, a.auctiontype, a.auctionname, a.itemid, a.limitno, i.itemname, i.listimage,i.smallimage,"
		sqlStr = sqlStr + " a.finishdate, a.forcefinish,"
		sqlStr = sqlStr + " convert(varchar(32),a.startdate,20) as startdate, convert(varchar(32),a.finishdate,20) as finishdate, convert(varchar(32),a.regdate,20) as regdate, a.supplyer"
		sqlStr = sqlStr + " from tbl_board_auction a, tbl_item i"
		sqlStr = sqlStr + " where a.itemid=i.itemid"
		sqlStr = sqlStr + " and (a.forcefinish='Y' or"
		sqlStr = sqlStr + "      a.finishdate<getdate())"
		if userid<>"" then
			sqlStr = sqlStr + " and i.makerid='" + userid +"'"
		end if

		if CStr(notthisitem)<>"" then
			sqlStr = sqlStr + " and i.itemid<>" + CStr(notthisitem)
		end if
		sqlStr = sqlStr + " and a.isusing='Y'"
		sqlStr = sqlStr + " order by id desc"
		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FTotalPage = rsget.PageCount
		FResultCount = rsget.RecordCount - (CurrPage-1)*PageSize

		if (FResultCount>PageSize) then
			FResultCount = PageSize
		end if
		if  not rsget.EOF  then
			redim preserve FAuctionID(FResultCount)
			redim preserve FAuctionType(FResultCount)
			redim preserve FAuctionName(FResultCount)
			redim preserve FItemID(FResultCount)
			redim preserve Flimitno(FResultCount)

			redim preserve FItemName(FResultCount)

			redim preserve FStartDate(FResultCount)
			redim preserve FFinishDate(FResultCount)
			redim preserve FRegDate(FResultCount)
			redim preserve FSupplyer(FResultCount)

			redim preserve FImageList(FResultCount)
			redim preserve FImageSmall(FResultCount)
			redim preserve FForceFinish(FResultCount)

			redim preserve FPriceStart(FResultCount)
			redim preserve FPriceEnd(FResultCount)
			redim preserve FPriceFix(FResultCount)

			dim ix
			ix =0
			rsget.absolutepage = FCurrPage
			do until (rsget.eof or ix>PageSize)
				FAuctionID(ix)   =	rsget("id")
				FAuctionType(ix) =	rsget("auctiontype")
				FAuctionName(ix) =	rsget("auctionname")
				FItemID(ix)      =	rsget("itemid")
				FLimitNo(ix)      =	rsget("limitno")

				FItemName(ix)    =	rsget("itemname")

				FStartDate(ix)    =	rsget("startdate")
				FFinishDate(ix)   =	rsget("finishdate")
				FRegDate(ix)   =	rsget("regdate")
				FSupplyer(ix)   =	rsget("supplyer")

				FImageList(ix)   =	rsget("listimage")
				FImageSmall(ix)   =	rsget("smallimage")
				FForceFinish(ix) =	rsget("forcefinish")
				FFinishDate(ix) =	rsget("finishdate")

				rsget.movenext
				ix = ix +1
			loop
		end if
		rsget.close
	end sub


	Public Sub GetCurrentAuction(byval userid,byval notthisitem)
		dim sqlStr
		sqlStr = " select a.id, a.auctiontype, a.auctionname, a.itemid, a.limitno, i.itemname, i.listimage,i.smallimage,"
		sqlStr = sqlStr + " a.finishdate, a.forcefinish, a.linkitem"
		sqlStr = sqlStr + " from [db_contents].[dbo].tbl_board_auction a, [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where a.itemid=i.itemid"
		sqlStr = sqlStr + " and a.forcefinish='N'"
		sqlStr = sqlStr + " and a.finishdate>getdate()"
		if userid<>"" then
			sqlStr = sqlStr + " and i.makerid='" + userid +"'"
		end if

		if CStr(notthisitem)<>"" then
			sqlStr = sqlStr + " and i.itemid<>" + CStr(notthisitem)
		end if
		sqlStr = sqlStr + " and a.isusing='Y'"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		if  not rsget.EOF  then
			redim preserve FAuctionID(FResultCount)
			redim preserve FAuctionType(FResultCount)
			redim preserve FAuctionName(FResultCount)
			redim preserve FItemID(FResultCount)
			redim preserve Flimitno(FResultCount)
			redim preserve FItemName(FResultCount)
			redim preserve FImageList(FResultCount)
			redim preserve FImageSmall(FResultCount)
			redim preserve FForceFinish(FResultCount)
			redim preserve FFinishDate(FResultCount)
			redim preserve FLinkItem(FResultCount)

			dim ix
			ix =0
			do until rsget.eof
				FAuctionID(ix)   =	rsget("id")
				FAuctionType(ix) =	rsget("auctiontype")
				FAuctionName(ix) =	rsget("auctionname")
				FItemID(ix)      =	rsget("itemid")
				Flimitno(ix)	 =	rsget("limitno")

				FItemName(ix)    =	rsget("itemname")
				FImageList(ix)   =	rsget("listimage")
				FImageSmall(ix)   =	rsget("smallimage")
				FForceFinish(ix) =	rsget("forcefinish")
				FFinishDate(ix) =	rsget("finishdate")
				FLinkItem(ix)	=   rsget("linkitem")

				rsget.movenext
				ix = ix +1
			loop
		end if
		rsget.close
	end sub


	Public Sub GetLastAuction()
		'일단 이걸루..
		FGetCount = 1
		GetAuctionDatas("")
	end Sub

	Public Sub GetAuctionDatas(byval auctionid)
		dim sqlStr
		if FGetCount=0 then
			sqlStr = "select "
		else
			sqlStr = "select top " + CStr(FGetCount)
		end if

		sqlStr = sqlStr + " a.id, a.auctiontype, a.auctionname, a.itemid, a.limitno, "
		sqlStr = sqlStr + " convert(varchar(32),a.startdate,20) as startdate, convert(varchar(32),a.finishdate,20) as finishdate, convert(varchar(32),a.regdate,20) as regdate, a.supplyer, a.forcefinish,"
		sqlStr = sqlStr + " b.itemname, b.listimage,"
		sqlStr = sqlStr + " a.pricestart, a.priceend, a.pricefix, a.linkitem"
		sqlStr = sqlStr + " from [db_contents].[dbo].tbl_board_auction a, [db_item].[dbo].tbl_item b"
		sqlStr = sqlStr + " where a.itemid = b.itemid"
		if auctionid<>"" then
			sqlStr = sqlStr + " and a.id=" + CStr(auctionid)
		end if

		sqlStr = sqlStr + " and a.isusing='Y'"
		sqlStr = sqlStr + " order by a.id desc"

		exec sqlStr
	end Sub

	private Sub Exec(byval sqlStr)
		dim ix
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		if  not rsget.EOF  then
			redim preserve FAuctionID(FResultCount)
			redim preserve FAuctionType(FResultCount)
			redim preserve FAuctionName(FResultCount)
			redim preserve FItemID(FResultCount)
			redim preserve Flimitno(FResultCount)
			redim preserve FStartDate(FResultCount)
			redim preserve FFinishDate(FResultCount)
			redim preserve FRegDate(FResultCount)
			redim preserve FSupplyer(FResultCount)
			redim preserve FForceFinish(FResultCount)
			redim preserve FItemName(FResultCount)
			redim preserve FImageList(FResultCount)
			redim preserve FLinkItem(FResultCount)

			redim preserve FPriceStart(FResultCount)
			redim preserve FPriceEnd(FResultCount)
			redim preserve FPriceFix(FResultCount)


			ix =0
			do until rsget.eof
				FAuctionID(ix)   =	rsget("id")
				FAuctionType(ix) =	rsget("auctiontype")
				FAuctionName(ix) =	rsget("auctionname")
				FItemID(ix)      =	rsget("itemid")
				Flimitno(ix)     =	rsget("limitno")
				FStartDate(ix)   =	rsget("startdate")
				FFinishDate(ix)  =	rsget("finishdate")
				FRegDate(ix)     =	rsget("regdate")
				FSupplyer(ix)    =	rsget("supplyer")
				FForceFinish(ix) =	rsget("forcefinish")
				FItemName(ix)    =	rsget("itemname")
				FImageList(ix)   =	rsget("listimage")
				FLinkItem(ix)	 =  rsget("linkitem")

				FPriceStart(ix)	=	rsget("pricestart")
				FPriceEnd(ix)  	=	rsget("priceend")
				FPriceFix(ix)  	=	rsget("pricefix")

				rsget.movenext
				ix = ix +1
			loop
		end if
		rsget.close
	end Sub



	public property Get AuctionID(byval v)
		AuctionID = FAuctionID(v)
	end property

	public property Get AuctionType(byval v)
		AuctionType = FAuctionType(v)
	end property

	public property Get AuctionTypeName(byval v)
		if CInt(FAuctionType(v)) =1 then
			AuctionTypeName = "찍기"
		elseif CInt(FAuctionType(v)) =2 then
			AuctionTypeName = "큰소리치기"
		end if
	end property

	public property Get AuctionName(byval v)
		AuctionName = FAuctionName(v)
	end property

	public property Get ItemID(byval v)
		ItemID = FItemID(v)
	end property

	public property Get limitno(byval v)
		limitno = Flimitno(v)
	end property

	public property Get StartDate(byval v)
		StartDate = FStartDate(v)
	end property

	public property Get FinishDate(byval v)
		FinishDate = FFinishDate(v)
	end property

	public property Get RegDate(byval v)
		RegDate = FRegDate(v)
	end property

	public property Get Supplyer(byval v)
		Supplyer = FSupplyer(v)
	end property

	public property Get ForceFinish(byval v)
		ForceFinish = FForceFinish(v)
	end property

	public property Get ItemName(byval v)
		ItemName = FItemName(v)
	end property

	public property Get LinkItem(byval v)
		if isNull(FLinkItem(v)) then
			LinkItem = ""
		else
			LinkItem = FLinkItem(v)
		end if
	end property

	public property Get ImageList(byval v)
		if FImageList(v)<>"" then
			ImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(ItemId(v)) + "/" + FImageList(v)
		end if
	end property

	public property Get ImageSmall(byval v)
		if FImageSmall(v)<>"" then
			ImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(ItemId(v)) + "/" + FImageSmall(v)
		end if
	end property

	public property Get PriceStart(byval v)
		PriceStart = FPriceStart(v)
	end property

	public property Get PriceEnd(byval v)
		PriceEnd = FPriceEnd(v)
	end property

	public property Get PriceFix(byval v)
		PriceFix = FPriceFix(v)
	end property

	public function GetFinishDateFormat(byval i, byval v)
		if v=1 then
			GetFinishDateFormat = Mid(FinishDate(i),6,2) + "월 " + Mid(FinishDate(i),9,2) + "일 " + Mid(FinishDate(i),12,2) +"시"
		end if
	end function

	public function RemainTimeFormat(byval i, byval v)
		dim findate
		dim diffval
		dim dd, hh, mm, ss
		dd=0
		hh=0
		mm=0
		ss=0
		findate = CDate(FinishDate(i))

		dim nowerdate
		nowerdate = now

		dd = dateDiff("d",dateAdd("d",1,nowerdate),findate)
		diffval = dateDiff("s",dateAdd("d",dd,nowerdate),findate)

		hh= diffval\(60*60)
		if (hh>=24) then
			dd=dd+hh\24
			hh=hh mod 24
		end if
		mm= (diffval mod (60*60)) \60
		ss= diffval mod 60



		RemainTimeFormat=""
		if v=1 then
			if Cint(dd) >0 then RemainTimeFormat = CStr(dd) + "일 "
			RemainTimeFormat = RemainTimeFormat + CStr(hh) + "시간 " +  CStr(mm) + "분 " +  CStr(ss) + "초"
		end if
	end function

	public Function IsFinish(byval v)
		if (ForceFinish(v) ="Y") then
			IsFinish = true
		end if

		if (now >CDate(FinishDate(v))) then
			IsFinish = true
		end if
	end function

	public Function DoingCount()
		dim i,cnt
		cnt=0
		for i=0 to ResultCount-1
			if IsFinish(i) then
				cnt=cnt+1
			end if
		next
		DoingCount = ResultCount-cnt
	end function

	Public Function GetNavigateBar(byval urllink,byval align)
		dim buf,ix
		if TotalCount<1 then
			GetNavigateBar = ""
			Exit Function
		end if

		buf = "<table border='0' cellspacing='0' cellpadding='0' align='" + align + "'>" + vbCrlf
        buf = buf + "<tr>" + vbCrlf
        buf = buf + "<td class='a'>" + vbCrlf

        if HasPreScroll then
        buf = buf + "<a href='" + urllink + CStr(StarScrollPage-1) + "' class='my10x10'>[pre]</a>" + vbCrlf
        end if

        buf = buf + "<img src='/images/blue_arrow.gif' width='8' height='7'>&nbsp;" + vbCrlf

        for ix=StarScrollPage to StarScrollPage + ScrollCount-1
        	if (ix > TotalPage) then Exit For
            if CStr(ix) = CStr(CurrPage) then
              	buf = buf + "[" + CStr(ix) + "]"
            else
              	buf = buf + "<a href='" + urllink + CStr(ix) + "' class='my10x10'>" + VbCrlf
              	buf = buf + " " + CStr(ix)
              	buf = buf + "</a>"
            end if
        next

        buf = buf + "&nbsp;&nbsp;<img src='/images/blue_arrow2.gif' width='8' height='7'>" + VbCrlf

        if HasNextScroll then
        	buf = buf + "<a href='" + urllink + CStr(StarScrollPage+ScrollCount) + "' class='my10x10'>[next]</a>" + VbCrlf
        end if
        buf = buf + "</td>"
        buf = buf + "</tr>"
        buf = buf + "</table>"
        GetNavigateBar = buf
	end Function

	public Function GetWinnerList(byval v)
		''midesign님께서 <br>2002.05.01 [22:14:05] 에 낙찰되었습니다<br>
		dim sqlStr,buf,bufall
		sqlStr = "select userid,itemno,convert(varchar(20),regdate,20) as rdate from tbl_board_auction_winner"
		sqlStr = sqlStr + " where auctionid=" + CStr(v)
		sqlStr = sqlStr + " and isdelete<>'Y'"
		sqlStr = sqlStr + " order by regdate desc"
		rsget.Open sqlStr,dbget,1
		bufall = ""
		do until rsget.Eof

			buf= "낙찰자 : <b>" + rsget("userid") + "</b><br>" + replace(Left(rsget("rdate"),10),"-",".") + " "
			buf= buf + "[" + Mid(rsget("rdate"),12,64) +"]" + "<br>"
			bufall = bufall + buf
			rsget.MoveNext
		loop

		rsget.close

		GetWinnerList = bufall

	end Function
end Class
%>




            