<%

Class CItemMaeipSaleMarginShareMasterGrpItem
	public Fmakerid
	public FmaySum
	public Ftitle
	public Ffinishflag
	public Fjgubun
	public Fjacctcd
	public Fdifferencekey
	public Fet_cnt
	public Fdlv_totalsuplycash
	public Ftotalcommission
	public Fmaydiff
	Private Sub Class_Initialize()
        ''
	End Sub

	Private Sub Class_Terminate()
        '''
	End Sub
end Class

Class CItemMaeipSaleMarginShareMasterItem
    public Fidx
	public Fmakerid
	public FsaleCode
	public FstartDate
	public FendDate
	public FmeachulGubun
	public FdefaultMargin
	public FsaleMargin
	public Freguserid
	public Fuseyn
	public Fregdate
	public Flastupdate

	public function GetMeachulGubun()
		Select Case FmeachulGubun
			Case "1"
				GetMeachulGubun = "결제일 기준"
			Case "2"
				GetMeachulGubun = "출고일 기준"
			Case else
				GetMeachulGubun = FmeachulGubun
		End Select
	end function

    Private Sub Class_Initialize()
        ''
	End Sub

	Private Sub Class_Terminate()
        '''
	End Sub
end Class

Class CItemMaeipSaleMarginShareDetailItem
    public Fidx
	public Fmasteridx
	public Fitemid
	public Fitemname
	public Forgprice
	public Fsaleprice
	public ForgBuyCash
	public FsaleBuyCash
	public Fuseyn
	public Fregdate
	public Flastupdate

	public Fsmallimage
	public Flistimage
	public Fmakerid
	public Fmwdiv
	public Fcurrmwdiv

	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	public function getCurrMwDivName()
		if Fcurrmwdiv="M" then
			getCurrMwDivName = "매입"
		elseif Fcurrmwdiv="W" then
			getCurrMwDivName = "위탁"
		elseif Fcurrmwdiv="U" then
			getCurrMwDivName = "업체"
		end if
	end function

    Private Sub Class_Initialize()
        ''
	End Sub

	Private Sub Class_Terminate()
        '''
	End Sub
end Class

Class CItemMaeipSaleMarginShare
    public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectMakerid
	public FRectIdx
	public FRectYYYYMM

	public function SearchMaeipSaleMarginShareJungsanListGrp
		dim sqlStr

		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_JungsanTarget_meaipSaleMarginShare] '"&FRectYYYYMM&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemMaeipSaleMarginShareMasterGrpItem

				FItemList(i).Fmakerid				= rsget("makerid")
				FItemList(i).FmaySum				= rsget("maySum")

				FItemList(i).Ftitle					= rsget("title")
				FItemList(i).Ffinishflag			= rsget("finishflag")
				FItemList(i).Fjgubun				= rsget("jgubun")
				FItemList(i).Fjacctcd				= rsget("jacctcd")
				FItemList(i).Fdifferencekey			= rsget("differencekey")
				FItemList(i).Fet_cnt				= rsget("et_cnt")
				FItemList(i).Fdlv_totalsuplycash	= rsget("dlv_totalsuplycash")
				FItemList(i).Ftotalcommission		= rsget("totalcommission")
				FItemList(i).Fmaydiff				= rsget("maydiff")


                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close

	end function

	public function GetMasterList()
        dim sqlStr, addSql, i

        if (FRectMakerid <> "") then
            addSql = addSql & " and m.makerid = '" & FRectMakerid &  "' "
        end if

		'// ====================================================================
		sqlStr = "select count(m.idx) as cnt"
        sqlStr = sqlStr & " from [db_order].[dbo].[tbl_meaipSaleMarginShare_master] m "
		sqlStr = sqlStr & " where 1 = 1 " & addSql

		''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
		rsget.Close


		'// ====================================================================
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " m.* "
        sqlStr = sqlStr & " from [db_order].[dbo].[tbl_meaipSaleMarginShare_master] m "
		sqlStr = sqlStr & " where 1 = 1 " & addSql
		sqlStr = sqlStr & " order by m.idx desc "

		' Response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemMaeipSaleMarginShareMasterItem

				FItemList(i).Fidx            	= rsget("idx")
				FItemList(i).Fmakerid       	= rsget("makerid")
				FItemList(i).FsaleCode          = rsget("saleCode")
				FItemList(i).FstartDate         = rsget("startDate")
				FItemList(i).FendDate          	= rsget("endDate")
				FItemList(i).FmeachulGubun      = rsget("meachulGubun")
				FItemList(i).FdefaultMargin     = rsget("defaultMargin")
				FItemList(i).FsaleMargin        = rsget("saleMargin")
				FItemList(i).Freguserid         = rsget("reguserid")
				FItemList(i).Fuseyn          	= rsget("useyn")
				FItemList(i).Fregdate          	= rsget("regdate")
				FItemList(i).Flastupdate        = rsget("lastupdate")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	public function GetMasterOne()
		dim sqlStr, addSql, i

		sqlStr = "select top 1 m.* "
		sqlStr = sqlStr & " from [db_order].[dbo].[tbl_meaipSaleMarginShare_master] m "
		sqlStr = sqlStr & " where idx =  " & FRectIdx

		set FOneItem = new CItemMaeipSaleMarginShareMasterItem

		if (FRectIdx <> "") then
			rsget.Open sqlStr,dbget,1
			FTotalCount = rsget.RecordCount
			FResultCount = FTotalCount

			if Not rsget.Eof then
				FOneItem.Fidx          		= rsget("idx")
				FOneItem.Fmakerid       	= rsget("makerid")
				FOneItem.FsaleCode          = rsget("saleCode")
				FOneItem.FstartDate         = rsget("startDate")
				FOneItem.FendDate          	= rsget("endDate")
				FOneItem.FmeachulGubun      = rsget("meachulGubun")
				FOneItem.FdefaultMargin     = rsget("defaultMargin")
				FOneItem.FsaleMargin        = rsget("saleMargin")
				FOneItem.Freguserid         = rsget("reguserid")
				FOneItem.Fuseyn          	= rsget("useyn")
				FOneItem.Fregdate          	= rsget("regdate")
				FOneItem.Flastupdate        = rsget("lastupdate")
			end if
			rsget.Close
		end if
	end Function

	public function GetDetailList()
        dim sqlStr, addSql, i

		'// ====================================================================
		sqlStr = "select count(d.idx) as cnt"
        sqlStr = sqlStr & " from [db_order].[dbo].[tbl_meaipSaleMarginShare_detail] d "
		sqlStr = sqlStr & " where d.masteridx = " & FRectIdx

		''rsget.Open sqlStr,dbget,1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
		rsget.Close


		'// ====================================================================
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " d.*, i.smallimage, i.listimage, i.makerid, i.mwdiv as currmwdiv "
        sqlStr = sqlStr & " from [db_order].[dbo].[tbl_meaipSaleMarginShare_detail] d "
		sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item] i on d.itemid = i.itemid "
		sqlStr = sqlStr & " where d.masteridx = " & FRectIdx
		sqlStr = sqlStr & "	and d.useyn = 'Y' "
		sqlStr = sqlStr & " order by d.idx "

		' Response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemMaeipSaleMarginShareDetailItem

				FItemList(i).Fidx            	= rsget("idx")
				FItemList(i).Fmasteridx         = rsget("masteridx")
				FItemList(i).Fitemid            = rsget("itemid")
				FItemList(i).Fitemname          = rsget("itemname")
				FItemList(i).Forgprice          = rsget("orgprice")
				FItemList(i).Fsaleprice         = rsget("saleprice")
				FItemList(i).ForgBuyCash        = rsget("orgBuyCash")
				FItemList(i).FsaleBuyCash       = rsget("saleBuyCash")
				FItemList(i).Fuseyn            	= rsget("useyn")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Flastupdate        = rsget("lastupdate")

				FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).Flistimage        	= webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				FItemList(i).Fmakerid        	= rsget("makerid")
				FItemList(i).Fmwdiv        		= rsget("mwdiv")
				FItemList(i).Fcurrmwdiv    		= rsget("currmwdiv")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage 		= 1
		FPageSize 		= 25
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
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
