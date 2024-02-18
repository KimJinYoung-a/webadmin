<%

Class CRackIpgoItem
	public Fexecutedt

	public Fstatecd0
	public Fstatecd1
	public Fstatecd5
	public Fstatecd7
	public Fstatecd8
	public Fstatecd9
	public Ftotalnotfinishcount

	public Ftotalsellcash
	public Ftotalrackipgo
	public Frackipgo_y
        public Frackipgo_n

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CRackBrandItem
	public Fmakerid
	public Fmakername
	public Frackcode
	public Frackboxno

	public Fitemgubun
	public Fitemid
	public Fitemrackcode
	public Fitemname
	public Fdeliverytype
	public Fmwdiv

	public Fsellyn
	public FIsusing
	public Flimityn

	public Fimgsmall

    public FBrandUsing
    public FBrandUsingExt
    public FwarehouseCd

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CRackIpgo
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectExecuteDtStart
	public FRectExecuteDtEnd
	public FRectOrderByMakerid
	public FRectOrderByCode
	public FRectMakerid
	public FRectIsUsingYN
	public FRectSellYN

	public FRectMWDiv
	public FRectOnlyIsUsing
	public FRectdiffrackcode
    public FRectRackCode
	public FRectMaeipDiv

	public FRectSearchType
	public FRectFromRackcode2
	public FRectToRackcode2
	public FRectPurchaseType
    public FRectWarehouseCd

        '최근두달만 체크한다.
	public Sub GetRackIpgoList()
		dim i, sqlStr

                sqlStr = " select convert(varchar(10), m.executedt, 20) as dt, "
                sqlStr = sqlStr + "     sum(m.totalsellcash) as totalsellcash, "
                sqlStr = sqlStr + "     sum(case when m.rackipgoyn = 'Y' then 1 else 0 end) as rackipgo_y, "
                sqlStr = sqlStr + "     sum(case when m.rackipgoyn <> 'Y' then 1 else 0 end) as rackipgo_n "
                sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m "
                sqlStr = sqlStr + " where m.deldt is null "
                sqlStr = sqlStr + " and m.ipchulflag = 'I' "
                sqlStr = sqlStr + " and m.divcode in ('001','002') "
                sqlStr = sqlStr + " and m.totalsellcash > 0 "
                sqlStr = sqlStr + " and m.executedt is not null "
                sqlStr = sqlStr + " and m.executedt >= '" + Left(dateadd("m",-2,now),10) + "' "

		if FRectExecuteDtStart<>"" then
			sqlStr = sqlStr + " and m.executedt>='" + FRectExecuteDtStart + "'"
		end if

		if FRectExecuteDtEnd<>"" then
			sqlStr = sqlStr + " and m.executedt<'" + FRectExecuteDtEnd + "'"
		end if

                sqlStr = sqlStr + " group by convert(varchar(10), m.executedt, 20) "
                sqlStr = sqlStr + " order by convert(varchar(10), m.executedt, 20) desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CRackIpgoItem

				FItemList(i).Fexecutedt      = rsget("dt")
				FItemList(i).Ftotalsellcash  = rsget("totalsellcash")

				FItemList(i).Frackipgo_y     = rsget("rackipgo_y")
				FItemList(i).Frackipgo_n     = rsget("rackipgo_n")
				FItemList(i).Ftotalrackipgo  = FItemList(i).Frackipgo_y + FItemList(i).Frackipgo_n

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

        '최근두달만 체크한다.
	public Sub GetRackJumunList()
		dim i, sqlStr

        sqlStr = " select "
        sqlStr = sqlStr + "     convert(varchar(10), m.regdate, 20) as dt, "
        sqlStr = sqlStr + "     sum(case when m.statecd = '0' then 1 else 0 end) as statecd0, "
        sqlStr = sqlStr + "     sum(case when m.statecd = '1' then 1 else 0 end) as statecd1, "
        sqlStr = sqlStr + "     sum(case when m.statecd = '5' then 1 else 0 end) as statecd5, "
        sqlStr = sqlStr + "     sum(case when m.statecd = '7' then 1 else 0 end) as statecd7, "
        sqlStr = sqlStr + "     sum(case when m.statecd = '8' then 1 else 0 end) as statecd8, "
        sqlStr = sqlStr + "     sum(case when m.statecd = '9' then 1 else 0 end) as statecd9 "
        sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m "
        sqlStr = sqlStr + " where m.deldt is null "
        sqlStr = sqlStr + " and m.divcode in ('301','302') "
        sqlStr = sqlStr + " and m.regdate >= '" + Left(dateadd("m",-2,now),10) + "' "

		if FRectExecuteDtStart<>"" then
			sqlStr = sqlStr + " and m.executedt>='" + FRectExecuteDtStart + "'"
		end if

		if FRectExecuteDtEnd<>"" then
			sqlStr = sqlStr + " and m.executedt<'" + FRectExecuteDtEnd + "'"
		end if

        sqlStr = sqlStr + " group by convert(varchar(10), m.regdate, 20) "
        sqlStr = sqlStr + " order by convert(varchar(10), m.regdate, 20) desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CRackIpgoItem

				FItemList(i).Fexecutedt      = rsget("dt")

				FItemList(i).Fstatecd0       = rsget("statecd0")
				FItemList(i).Fstatecd1       = rsget("statecd1")
				FItemList(i).Fstatecd5       = rsget("statecd5")
				FItemList(i).Fstatecd7       = rsget("statecd7")
				FItemList(i).Fstatecd8       = rsget("statecd8")
				FItemList(i).Fstatecd9       = rsget("statecd9")

				FItemList(i).Ftotalnotfinishcount= rsget("statecd0") + rsget("statecd1") + rsget("statecd5") + rsget("statecd7") + rsget("statecd8")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	public Sub GetRackMakerList()
		dim i, sqlStr

        sqlStr = " select top 1000 c.userid as makerid, c.socname_kor as makername, c.prtidx as rackcode "
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c "
        sqlStr = sqlStr + " where c.prtidx <> '9999' "

        if (FRectIsUsingYN <> "") then
                sqlStr = sqlStr + " and c.isusing = '" + CStr(FRectIsUsingYN) + "' "
        end if

        if (FRectMakerid <> "") then
                sqlStr = sqlStr + " and c.userid like '" + CStr(FRectMakerid) + "%' "
        end if

        if (FRectOrderByMakerid = "Y") then
                sqlStr = sqlStr + " order by c.makerid "
        elseif (FRectOrderByCode = "Y") then
                sqlStr = sqlStr + " order by c.prtidx, c.makerid "
        end if

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CRackBrandItem

				FItemList(i).Fmakerid        = rsget("makerid")
				FItemList(i).Fmakername      = db2html(rsget("makername"))
				FItemList(i).Frackcode       = Format00(4,rsget("rackcode"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	public function GetMakerByCode(rackcode)
		dim i, result

		result = ""
		for i=0 to FResultcount-1
		        if (FItemList(i).Frackcode = rackcode) then
		                if (result = "") then
		                        result = FItemList(i).Fmakerid
		                else
		                        result = result + "<br>" + FItemList(i).Fmakerid
		                end if
		        end if
		next
                GetMakerByCode = result
	end function

    public Sub GetRackBrandList()
        dim i, sqlStr, addSql

		addSql = ""

		if (FRectSearchType = "R") then
			if Len(FRectFromRackcode2)=2 then
				addSql = addSql + " and left(c.prtidx,2) >= '" & FRectFromRackcode2 &"' "
				addSql = addSql + " and left(c.prtidx,2) <= '" & FRectToRackcode2 &"' "
			else
				addSql = addSql + " and c.prtidx >= '" & FRectFromRackcode2 &"' "
				addSql = addSql + " and c.prtidx <= '" & FRectToRackcode2 &"' "
			end if
		elseif (FRectSearchType = "F") then
			if Len(FRectRackCode)=2 then
				addSql = addSql + " and left(c.prtidx,2) = '" & FRectRackCode &"'"
			else
				addSql = addSql + " and c.prtidx = '" & FRectRackCode &"'"
			end if
		end if

	    if (FRectMaeipDiv<>"") then
	        addSql = addSql + " and c.maeipdiv = '" + FRectMaeipDiv + "'"
	    end if

		if (FRectMakerid<>"") then
	        addSql = addSql + " and c.userid = '" + FRectMakerid + "'"
	    end if

	    if (FRectIsUsingYN<>"") then
	        addSql = addSql + " and c.isusing = '" + FRectIsUsingYN + "'"
	    end if

	    if (FRectPurchaseType <> "") then
	        addSql = addSql + " and p.purchasetype = '" + FRectPurchaseType + "'"
	    end if

        if (FRectWarehouseCd <> "") then
	        addSql = addSql + " and IsNull(c.warehouseCd, 'NUL') = '" + FRectWarehouseCd + "'"
	    end if

        sqlstr = " select count(c.userid) as cnt"
        sqlstr = sqlstr + " from [db_user].[dbo].tbl_user_c c"
		sqlstr = sqlstr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
		sqlStr = sqlStr + " where 1=1"

		sqlStr = sqlStr + addSql

		rsget.Open sqlStr, dbget, 1
		    FTotalCount = rsget("cnt")
		rsget.Close

        sqlstr = " select top  " & (FPageSize*FCurrPage)
		sqlstr = sqlstr + " c.userid as makerid, c.socname_kor as makername, IsNull(c.rackcodeByBrand, c.prtidx) as rackcode, c.rackboxno "
		sqlstr = sqlstr + " ,c.isusing, c.isextusing, c.warehouseCd "
		sqlstr = sqlstr + " from [db_user].[dbo].tbl_user_c c"
		sqlstr = sqlstr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
		sqlStr = sqlStr + " where 1=1"

		sqlStr = sqlStr + addSql

		sqlStr = sqlStr + " order by IsNull(c.rackcodeByBrand, c.prtidx), c.userid "
		''response.write sqlStr

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CRackBrandItem

				FItemList(i).Fmakerid        = rsget("makerid")
				FItemList(i).Fmakername      = db2html(rsget("makername"))
				FItemList(i).Frackcode       = rsget("rackcode")
				FItemList(i).Frackboxno      = rsget("rackboxno")

                FItemList(i).FBrandUsing        = rsget("isusing")
                FItemList(i).FBrandUsingExt     = rsget("isextusing")

                FItemList(i).FwarehouseCd     = rsget("warehouseCd")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

    end Sub

	public Sub GetRackBrandItemList()
		dim i, sqlStr

		sqlstr = "select top 500 i.makerid,i.itemrackcode,i.itemgubun,i.itemid,i.itemname,i.deliverytype,i.mwdiv,i.sellyn,i.isusing,i.limityn,i.smallimage, "
		sqlstr = sqlstr + " c.userid as makerid, c.socname_kor as makername, c.prtidx as rackcode "
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i "
		sqlstr = sqlstr + "     left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlstr + "     on i.makerid=c.userid"
		sqlStr = sqlStr + " where ISNULL(c.prtidx,'9999') <> '9999' "

		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and i.makerid='" + CStr(FRectMakerid) + "'"
		end if

		if FRectSellYN<>"" then
			sqlstr = sqlstr + " and i.sellyn='" +CStr(FRectSellYN) +"'"
		end if

		if FRectIsUsingYN<>"" then
			sqlstr = sqlstr + " and i.isusing='" + CStr(FRectIsUsingYN) +"'"
		end if

		if FRectMWDiv="MW" then
			sqlstr = sqlstr + " and i.mwdiv<>'U'"
		else
			sqlstr = sqlstr + " and i.mwdiv='" + CStr(FRectMWDiv) + "'"
		end if

		if FRectdiffrackcode="on" then
			''sqlstr = sqlstr + " and convert(int,i.itemrackcode/100) <> convert(int,c.prtidx/100) "
			sqlstr = sqlstr + " and Left(i.itemrackcode,2) <> Left(c.prtidx,2) "
		end if

        sqlstr = sqlstr + " order by i.makerid, i.itemid desc "
		''response.write sqlstr


		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CRackBrandItem

				FItemList(i).Fmakerid        = rsget("makerid")
				FItemList(i).Fmakername      = db2html(rsget("makername"))
				FItemList(i).Frackcode       = rsget("rackcode")

				FItemList(i).Fitemgubun     = rsget("itemgubun")
    			FItemList(i).Fitemid        = rsget("itemid")
    			FItemList(i).Fitemrackcode  = rsget("itemrackcode")
    			FItemList(i).Fitemname      = db2html(rsget("itemname"))
    			FItemList(i).Fdeliverytype  = rsget("deliverytype")
    			FItemList(i).Fmwdiv			= rsget("mwdiv")


    			FItemList(i).Fsellyn		= rsget("sellyn")
				FItemList(i).FIsusing		= rsget("isusing")
				FItemList(i).Flimityn		= rsget("limityn")

    			FItemList(i).Fimgsmall      = rsget("smallimage")

    			if isnull(FItemList(i).Fimgsmall) then FItemList(i).Fimgsmall=""
                if FItemList(i).Fimgsmall<>"" then FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall


				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub


	Private Sub Class_Initialize()
		FCurrPage = 1
		FPageSize = 200
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
end Class

%>
