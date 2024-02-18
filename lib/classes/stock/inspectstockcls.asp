<%
Class CInspectStockItem
	public FItemGubun
	public FItemID
	public FItemOption
	public FItemName
	public FItemOptionName

	public FMakerid
	public FMwDiv
	public FSellyn
	public FIsUsing

	public Ftotsellno
	public Ftotipgono
	public Ftotchulgono
	public Foffsellno

	public function GetMwDivName()
		if FMwDiv="M" then
			GetMwDivName = "매입"
		elseif FMwDiv="W" then
			GetMwDivName = "위탁"
		elseif FMwDiv="U" then
			GetMwDivName = "업체"
		end if
	end function

	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end Class

Class CIpchulDetailWithOrderSheet
	public Fid
	public Fmasterid
	public Fmastercode
	public Fitemid
	public Fitemoption
	public Fsellcash
	public Fsuplycash
	public Fitemno
	public Findt
	public Fupdt
	public Fdeldt
	public Fbuycash
	public Fmwgubun
	public Fiitemgubun
	public Fiitemname
	public Fiitemoptionname
	public Fimakerid
	public Fipchulflag

	public Fsheetidx
	public Fsheetitemname
	public Fsheetitemoptionname

	public Fbaljuitemno
	public Frealitemno

	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end Class

Class COrderDetailItem
	public Forderserial
	public Fdetailidx
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fitemno
	public FmcancelYn
	public FdcancelYn
	public Fipkumdiv
	public FIsUpchebeasong

	public FMakerid

	public function IsCancelJumun()
		IsCancelJumun = (FmcancelYn="Y") or (FmcancelYn="D") or (FdcancelYn="Y")
	end function

	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end Class


Class CAcountStorageDetailItem
	public Fid
	public Fmasterid
	public Fmastercode
	public Fitemid
	public Fitemoption
	public Fsellcash
	public Fsuplycash
	public Fitemno
	public Findt
	public Fupdt
	public Fdeldt
	public Fbuycash
	public Fmwgubun
	public Fiitemgubun
	public Fiitemname
	public Fiitemoptionname
	public Fimakerid

	public Fipchulflag
	public FCode
	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end Class



Class COrderSheetDetailItem
	public Fidx
	public Fmasteridx
	public Fitemgubun
	public Fmakerid
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fsellcash
	public Fsuplycash
	public Fbuycash
	public Fbaljuitemno
	public Frealitemno
	public Fregdate
	public Fupdt
	public Fdeldt
	public Fbaljudiv
	public Fcomment
	public Fipgoflag
	public Fdefaultmaginflag
	public Fbuymaginflag
	public Fsuplymaginflag

	public Fbaljucode

	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end Class



Class CInspectStock
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectMakerid
	public FRectIpchulCode
	public FRectIsUsing

	public FRectMwDiv

	public FRectOldOrNew
	public FRectItemGubun
	public FRectItemId
	public FRectItemOption

	public sub GetOnlineSellDetail()
		dim sqlstr, i
		sqlstr = "select top 500 m.orderserial, d.idx as detailidx, d.makerid, d.itemid, d.itemoption,d.itemname, d.itemoptionname, d.itemno,"
		sqlstr = sqlstr + "  m.cancelyn as mcancelyn, d.cancelyn as dcancelyn, m.ipkumdiv, d.isupchebeasong "
		sqlstr = sqlstr + "  from "
		if FRectOldOrNew="old" then
			sqlstr = sqlstr + " [db_log].[dbo].tbl_old_order_master_2003 m,"
			sqlstr = sqlstr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
		else
			sqlstr = sqlstr + " [db_order].[dbo].tbl_order_master m,"
			sqlstr = sqlstr + " [db_order].[dbo].tbl_order_detail d"
		end if

		sqlstr = sqlstr + " where m.orderserial=d.orderserial"
		sqlstr = sqlstr + " and d.itemid=" + CStr(FRectItemId)
		sqlstr = sqlstr + " and d.itemoption='" + FRectItemOption + "'"
		sqlstr = sqlstr + " order by m.orderserial"

'response.write sqlstr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderDetailItem

				FItemList(i).Forderserial    = rsget("orderserial")
				FItemList(i).Fdetailidx      = rsget("detailidx")
				FItemList(i).FMakerid		 = rsget("makerid")
				FItemList(i).Fitemgubun      = "10"
				FItemList(i).Fitemid         = rsget("itemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).Fitemname       = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).Fitemno         = rsget("itemno")
				FItemList(i).FmcancelYn      = rsget("mcancelYn")
				FItemList(i).FdcancelYn      = rsget("dcancelYn")
				FItemList(i).Fipkumdiv       = rsget("ipkumdiv")
				FItemList(i).FIsUpchebeasong = rsget("isupchebeasong")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close


	end sub

	public sub GetIpChulDetail()
		dim sqlstr, i
		sqlstr = "select top 500 "
		sqlstr = sqlstr + " m.code, m.ipchulflag, d.* "
		sqlstr = sqlstr + "  from "
		sqlstr = sqlstr + " [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlstr = sqlstr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlstr = sqlstr + " where m.code=d.mastercode"
		sqlstr = sqlstr + " and m.code='" + FRectIpchulCode + "'"
		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and d.imakerid='" + FRectMakerid + "'"
		end if
		sqlstr = sqlstr + " order by d.iitemgubun, d.itemid, d.itemoption"


		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CAcountStorageDetailItem

				FItemList(i).Fid              = rsget("id")
				FItemList(i).Fmasterid        = rsget("masterid")
				FItemList(i).Fmastercode      = rsget("mastercode")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fsellcash        = rsget("sellcash")
				FItemList(i).Fsuplycash       = rsget("suplycash")
				FItemList(i).Fitemno          = rsget("itemno")
				FItemList(i).Findt            = rsget("indt")
				FItemList(i).Fupdt            = rsget("updt")
				FItemList(i).Fdeldt           = rsget("deldt")
				FItemList(i).Fbuycash         = rsget("buycash")
				FItemList(i).Fmwgubun         = rsget("mwgubun")
				FItemList(i).Fiitemgubun      = rsget("iitemgubun")
				FItemList(i).Fiitemname       = db2html(rsget("iitemname"))
				FItemList(i).Fiitemoptionname = db2html(rsget("iitemoptionname"))
				FItemList(i).Fimakerid        = rsget("imakerid")

				FItemList(i).FCode			= rsget("code")
				FItemList(i).Fipchulflag	= rsget("ipchulflag")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close

	end sub


	public sub GetOrderSheetDetail()
		dim sqlstr, i

		sqlstr = "select top 500 "
		sqlstr = sqlstr + " m.baljucode, d.* "
		sqlstr = sqlstr + "  from "
		sqlstr = sqlstr + " [db_storage].[dbo].tbl_ordersheet_master m,"
		sqlstr = sqlstr + " [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlstr = sqlstr + " where m.idx=d.masteridx"
		sqlstr = sqlstr + " and m.alinkcode='" + FRectIpchulCode + "'"
		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and d.makerid='" + FRectMakerid + "'"
		end if
		sqlstr = sqlstr + " order by d.itemgubun, d.itemid, d.itemoption"


		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderSheetDetailItem

				FItemList(i).Fidx                = rsget("idx")
				FItemList(i).Fmasteridx          = rsget("masteridx")
				FItemList(i).Fitemgubun          = rsget("itemgubun")
				FItemList(i).Fmakerid            = rsget("makerid")
				FItemList(i).Fitemid             = rsget("itemid")
				FItemList(i).Fitemoption         = rsget("itemoption")
				FItemList(i).Fitemname           = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname     = db2html(rsget("itemoptionname"))
				FItemList(i).Fsellcash           = rsget("sellcash")
				FItemList(i).Fsuplycash          = rsget("suplycash")
				FItemList(i).Fbuycash            = rsget("buycash")
				FItemList(i).Fbaljuitemno        = rsget("baljuitemno")
				FItemList(i).Frealitemno         = rsget("realitemno")
				FItemList(i).Fregdate            = rsget("regdate")
				FItemList(i).Fupdt               = rsget("updt")
				FItemList(i).Fdeldt              = rsget("deldt")
				FItemList(i).Fbaljudiv           = rsget("baljudiv")
				FItemList(i).Fcomment            = rsget("comment")
				FItemList(i).Fipgoflag           = rsget("ipgoflag")
				FItemList(i).Fdefaultmaginflag   = rsget("defaultmaginflag")
				FItemList(i).Fbuymaginflag       = rsget("buymaginflag")
				FItemList(i).Fsuplymaginflag     = rsget("suplymaginflag")


				FItemList(i).Fbaljucode				= rsget("baljucode")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub

	public sub GetIpChulDetailWithOrderSheet()
		dim sqlstr, i

		sqlstr = "select top 500 "
		sqlstr = sqlstr + " d.*, m.ipchulflag,"
		sqlstr = sqlstr + " T.idx as sheetidx, T.itemname as sheetitemname, T.itemoptionname as sheetitemoptionname, T.realitemno"
		sqlstr = sqlstr + "  from "
		sqlstr = sqlstr + " [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlstr = sqlstr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlstr = sqlstr + " left join ("
		sqlstr = sqlstr + " 	select top 100 d.* from [db_storage].[dbo].tbl_ordersheet_master m,"
		sqlstr = sqlstr + " 	[db_storage].[dbo].tbl_ordersheet_detail d"
		sqlstr = sqlstr + " 	where m.idx=d.masteridx"
		sqlstr = sqlstr + " 	and d.realitemno<>0"
		sqlstr = sqlstr + " 	and m.alinkcode='" + FRectIpchulCode + "'"
		if FRectMakerid<>"" then
			sqlstr = sqlstr + " 	and d.makerid='" + FRectMakerid + "'"
		end if
		sqlstr = sqlstr + " ) T "
		sqlstr = sqlstr + " on d.iitemgubun=T.itemgubun"
		sqlstr = sqlstr + " and d.itemid=T.itemid"
		sqlstr = sqlstr + " and d.itemoption=T.itemoption"

		sqlstr = sqlstr + " where m.code=d.mastercode"
		sqlstr = sqlstr + " and m.code='" + FRectIpchulCode + "'"
		if FRectMakerid<>"" then
			sqlstr = sqlstr + " and d.imakerid='" + FRectMakerid + "'"
		end if

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CIpchulDetailWithOrderSheet

				FItemList(i).Fid              = rsget("id")
				FItemList(i).Fmasterid        = rsget("masterid")
				FItemList(i).Fmastercode      = rsget("mastercode")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fsellcash        = rsget("sellcash")
				FItemList(i).Fsuplycash       = rsget("suplycash")
				FItemList(i).Fitemno          = rsget("itemno")
				FItemList(i).Findt            = rsget("indt")
				FItemList(i).Fupdt            = rsget("updt")
				FItemList(i).Fdeldt           = rsget("deldt")
				FItemList(i).Fbuycash         = rsget("buycash")
				FItemList(i).Fmwgubun         = rsget("mwgubun")
				FItemList(i).Fiitemgubun      = rsget("iitemgubun")
				FItemList(i).Fiitemname       = db2html(rsget("iitemname"))
				FItemList(i).Fiitemoptionname = db2html(rsget("iitemoptionname"))
				FItemList(i).Fimakerid        = rsget("imakerid")

				FItemList(i).Fipchulflag	= rsget("ipchulflag")

				FItemList(i).Fsheetidx         = rsget("sheetidx")
				FItemList(i).Fsheetitemname			= db2html(rsget("sheetitemname"))
				FItemList(i).Fsheetitemoptionname	= db2html(rsget("sheetitemoptionname"))
				FItemList(i).Frealitemno		= rsget("realitemno")*-1
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close

	end sub



	public function GetErrRegItemList()
		dim sqlstr, i


		sqlstr = "select top " + CStr(FPageSize*FCurrPage) + " s.itemgubun, s.itemid, s.itemoption,"
		sqlstr = sqlstr + " s.totsellno, s.totipgono, s.totchulgono, s.offsellno, "
		sqlstr = sqlstr + " i.itemname, T.optionname as itemoptionname, i.makerid, "
		sqlstr = sqlstr + " i.mwdiv, i.isusing, i.sellyn"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item i,"
		sqlstr = sqlstr + " [db_summary].[dbo].tbl_current_logisstock_summary s"
		sqlstr = sqlstr + "     left join ("
		sqlstr = sqlstr + "         select i.itemid, IsNULL(o.itemoption,'0000') as itemoption, IsNULL(o.optionname,'') as optionname "
		sqlstr = sqlstr + "         from [db_item].[dbo].tbl_item i"
		sqlstr = sqlstr + "         left join [db_item].[dbo].tbl_item_option o"
		sqlstr = sqlstr + "         on i.itemid=o.itemid"
		if (FRectMakerid<>"") then
		    sqlstr = sqlstr + "     where i.makerid='" & FRectMakerid & "'"
		end if
		sqlstr = sqlstr + "     ) T"
		sqlstr = sqlstr + "     on s.itemgubun='10'"
		sqlstr = sqlstr + "     and s.itemid=T.itemid"
		sqlstr = sqlstr + "     and s.itemoption=T.itemoption"
		sqlstr = sqlstr + " where s.itemid<>0"
		sqlstr = sqlstr + " and s.itemgubun='10'"
		sqlstr = sqlstr + " and s.itemid=i.itemid"
		sqlstr = sqlstr + " and T.itemid is null"
		if (FRectMakerid<>"") then
		    sqlstr = sqlstr + "     and i.makerid='" & FRectMakerid & "'"
		end if
		
		if FRectIsUsing="Y" then
			sqlstr = sqlstr + " and i.isusing='Y'"
		end if

		if FRectMwDiv="U" then
			sqlstr = sqlstr + " and i.mwdiv='U'"
		elseif FRectMwDiv="T" then
			sqlstr = sqlstr + " and i.mwdiv<>'U'"
		end if
		sqlstr = sqlstr + " order by s.itemid desc, s.itemoption "

'response.write sqlstr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		FTotalCount = FResultCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInspectStockItem
				FItemList(i).FItemGubun       = rsget("itemgubun")
				FItemList(i).FItemID          = rsget("itemid")
				FItemList(i).FItemOption      = rsget("itemoption")
				FItemList(i).FMakerid      		= rsget("makerid")
				FItemList(i).FItemName        = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName  = db2html(rsget("itemoptionname"))

				FItemList(i).FMwDiv      = rsget("mwdiv")
				FItemList(i).FSellyn      = rsget("sellyn")
				FItemList(i).FIsUsing      = rsget("isusing")

				FItemList(i).Ftotsellno	= rsget("totsellno")
				FItemList(i).Ftotipgono = rsget("totipgono")
				FItemList(i).Ftotchulgono	= rsget("totchulgono")
				FItemList(i).Foffsellno	= rsget("offsellno")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
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