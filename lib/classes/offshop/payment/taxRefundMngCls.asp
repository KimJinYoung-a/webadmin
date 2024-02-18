<%
'####################################################
' Description :  tax Refund 클래스
' History : 2014.01.17 서동석
'####################################################

class CTaxRefundOneitem
	public Fidx
	public Fshopid
    public Forderno
    public Ftotalsum
    public Frealsum
    public Fjumundiv
    public Fjumunmethod
    public Fshopregdate
    public Fcancelyn
    public Fregdate
    public Fshopidx
    public Fspendmile
    public Fpointuserno
    public Fgainmile
    public Ftableno
    public Fcashsum
    public Fcardsum
    public FGiftCardPaySum
    public FTenGiftCardPaySum
    public Fcasherid
    public FCashReceiptNo
    public FCardAppNo
    public FCashreceiptGubun
    public FCardInstallment
    public FTenGiftCardMatchCode
    public FrefOrderNo
    public FIXyyyymmdd
    public Fmaechulgubun
    public Fbuyergubun
    public Ftaxrefundkey
    public FrefundMonth

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CTaxRefund
    public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount

    public FRectOldData
	public FRectStartDay
	public FRectEndDay
	public frectdatefg
	public frectbuyergubun
	public frectshopid
	public FRectInc3pl
	public frecttaxrefundkey
	public frectschType
	public frectscgRealsum
    public FRectRefundMonth

    public sub GetTaxRefundTargetList
		dim sqlStr, i, sqlsearch

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

		If frectschType <> "" Then
			Select Case frectschType
				Case "0"		sqlsearch = sqlsearch + " and isnull(m.taxrefundkey,'') <> '' "
				Case "1"		sqlsearch = sqlsearch + " and isnull(m.taxrefundkey,'') = '' "
				Case "2"		sqlsearch = sqlsearch + " and m.buyergubun = '200'"
			End Select
		End If

		If frectscgRealsum <> "" Then
			sqlsearch = sqlsearch + " and m.cashsum = '"&frectscgRealsum&"'"
		End If

		If frecttaxrefundkey <> "" Then
			sqlsearch = sqlsearch + " and m.taxrefundkey = '"&frecttaxrefundkey&"'"
		End If

        if (FRectRefundMonth<>"") then
            sqlsearch = sqlsearch + " and m.refundMonth = '"&FRectRefundMonth&"'"
        end if

        sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
        if FRectOldData="on" then
			sqlStr = sqlStr & " from [db_shoplog].[dbo].tbl_old_shopjumun_master m"
		else
			sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shopjumun_master m"
		end if
		sqlStr = sqlStr & " where m.cancelyn='N' " & sqlsearch

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if


		sqlStr=" select top " + CStr(FPageSize*FCurrPage) + " m.* "
		if FRectOldData="on" then
			sqlStr = sqlStr & " from [db_shoplog].[dbo].tbl_old_shopjumun_master m"
		else
			sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shopjumun_master m"
		end if
		sqlStr = sqlStr & " where m.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " ORDER BY m.idx  "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTaxRefundOneitem
				FItemList(i).Fidx	            = rsget("idx")
                FItemList(i).Fshopid            = rsget("shopid")
                FItemList(i).Forderno           = rsget("orderno")
                FItemList(i).Ftotalsum          = rsget("totalsum")
                FItemList(i).Frealsum           = rsget("realsum")
                FItemList(i).Fjumundiv          = rsget("jumundiv")
                FItemList(i).Fjumunmethod       = rsget("jumunmethod")
                FItemList(i).Fshopregdate       = rsget("shopregdate")
                FItemList(i).Fcancelyn          = rsget("cancelyn")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Fshopidx           = rsget("shopidx")
                FItemList(i).Fspendmile         = rsget("spendmile")
                FItemList(i).Fpointuserno       = rsget("pointuserno")
                FItemList(i).Fgainmile          = rsget("gainmile")
                FItemList(i).Ftableno           = rsget("tableno")
                FItemList(i).Fcashsum           = rsget("cashsum")
                FItemList(i).Fcardsum           = rsget("cardsum")
                FItemList(i).FGiftCardPaySum    = rsget("GiftCardPaySum")
                FItemList(i).FTenGiftCardPaySum = rsget("TenGiftCardPaySum")
                FItemList(i).Fcasherid          = rsget("casherid")
                FItemList(i).FCashReceiptNo     = rsget("CashReceiptNo")
                FItemList(i).FCardAppNo         = rsget("CardAppNo")
                FItemList(i).FCashreceiptGubun  = rsget("CashreceiptGubun")
                FItemList(i).FCardInstallment   = rsget("CardInstallment")
                FItemList(i).FTenGiftCardMatchCode = rsget("TenGiftCardMatchCode")
                FItemList(i).FrefOrderNo        = rsget("refOrderNo")
                FItemList(i).FIXyyyymmdd        = rsget("IXyyyymmdd")
                FItemList(i).Fmaechulgubun      = rsget("maechulgubun")
                FItemList(i).Fbuyergubun        = rsget("buyergubun")
                FItemList(i).Ftaxrefundkey      = rsget("taxrefundkey")
                FItemList(i).FrefundMonth       = rsget("refundMonth")
                i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

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
End Class
%>