<%
'####################################################
' Description :  오프라인 마일리지 클래스
' History : 2009.04.07 서동석 생성
'			2010.03.26 한용민 수정
'####################################################

Class COffShopMileageItem
	public Fidx
	public Fpointuserno
	public Fpoint
	public Fshopid
	public Fjukyo
	public Fregdate
	public Fdeleteyn
	public Fonlineuserid

	public Fpointusername

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COffShopMileage
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectShopID
	public FRectStartDay
	public FRectEndDay
	public FRectInc3pl
	public FRectLogDesc
	public FRectOnlineUserID

	'//admin/offshop/offmileagelist.asp
	public sub COffShopMileageList()
		dim i, sqlStr

		sqlStr = " select count(Log_idx) as cnt"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_total_shop_log g"
		sqlStr = sqlStr + " Join db_shop.dbo.tbl_total_shop_card c"
	    sqlStr = sqlStr + " 	on g.CardNo=c.CardNo"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_total_shop_user u"
		sqlStr = sqlStr + "     on c.Userseq=u.userSeq"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on g.regshopid=p.id "
		sqlStr = sqlStr + " where 1=1"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and g.regshopid='" + FRectShopID + "'"
		end if
        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')=''"
	    end if

		if (FRectOnlineUserID <> "") then
			sqlStr = sqlStr + " and u.onlineuserid = '" + CStr(FRectOnlineUserID) + "'"
		else
			sqlStr = sqlStr + " and g.regdate>='" + CStr(FRectStartDay) + "'"
			sqlStr = sqlStr + " and g.regdate<'" + CStr(FRectEndDay) + "'"
		end if


		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " g.Log_idx as idx, g.CArdNo as pointuserno, g.point, g.regshopid as shopid,"
		sqlStr = sqlStr + " g.LogDesc as jukyo, g.regdate, 'N' as deleteyn, u.username as pointusername, u.onlineuserid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_total_shop_log g"
		sqlStr = sqlStr + " Join db_shop.dbo.tbl_total_shop_card c"
	    sqlStr = sqlStr + " 	on g.CardNo=c.CardNo"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_total_shop_user u"
		sqlStr = sqlStr + "     on c.Userseq=u.userSeq"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on g.regshopid=p.id "
		sqlStr = sqlStr + " where 1=1"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and g.regshopid='" + FRectShopID + "'"
		end if
        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')=''"
	    end if

		if FRectLogDesc <> "" then
			sqlStr = sqlStr + " and g.logDesc like '" + CStr(FRectLogDesc) + "%' "
		end if

		if (FRectOnlineUserID <> "") then
			sqlStr = sqlStr + " and u.onlineuserid = '" + CStr(FRectOnlineUserID) + "'"
		else
			sqlStr = sqlStr + " and g.regdate>='" + CStr(FRectStartDay) + "'"
			sqlStr = sqlStr + " and g.regdate<'" + CStr(FRectEndDay) + "'"
		end if

		sqlStr = sqlStr + " order by g.Log_idx desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopMileageItem
				FItemList(i).Fidx          = rsget("idx")
				FItemList(i).Fpointuserno  = rsget("pointuserno")
				FItemList(i).Fpoint        = rsget("point")
				FItemList(i).Fshopid       = rsget("shopid")
				FItemList(i).Fjukyo        = db2html(rsget("jukyo"))
				FItemList(i).Fregdate      = rsget("regdate")
				FItemList(i).Fdeleteyn     = rsget("deleteyn")
				FItemList(i).Fonlineuserid     = rsget("onlineuserid")

				FItemList(i).Fpointusername= db2html(rsget("pointusername"))
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim  FItemList(0)
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
