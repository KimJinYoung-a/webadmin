<%
'###########################################################
' Description : 옵션관리
' Hieditor : 서동석 생성
'			 2022.07.06 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################

Class COptionManagerItem
	public Foptioncode01
	public Foptioncode02

	public Fcodename
	public Fcodevalue
	public Fcodeview

	public Foptiondispyn01
	public Foptiondispyn02

	public Fdisporder01
	public Fdisporder02
	public Fkeyword
	public FCode01
	public FCode02

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COptionManager
	public FItemList()
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectOnlyUsing
	public FRectOrderType

	public Sub GetOption01(cdl)
		dim sqlStr,i
		sqlStr = "select top 1 * from [db_item].[dbo].tbl_option_div01"
		sqlStr = sqlStr + " where optioncode01='" + cdl + "'"

		rsget.Open sqlStr, dbget, 1

		redim preserve FItemList(0)
		if not rsget.Eof then
			set FItemList(0) = new COptionManagerItem

			FItemList(0).Foptioncode01 = rsget("optioncode01")
			FItemList(0).Fcodename = db2html(rsget("codename"))
			FItemList(0).Foptiondispyn01 = rsget("optiondispyn")
			FItemList(0).Fdisporder01 = rsget("disporder")
		end if

		rsget.close
	end sub

	public Sub GetOption02(cdl,cdm)
		dim sqlStr,i
		sqlStr = "select top 1 * from [db_item].[dbo].tbl_option_div02"
		sqlStr = sqlStr + " where optioncode01='" + cdl + "'"
		sqlStr = sqlStr + " and optioncode02='" + cdm + "'"
		rsget.Open sqlStr, dbget, 1

		redim preserve FItemList(0)
		if not rsget.Eof then
			set FItemList(0) = new COptionManagerItem

			FItemList(i).Foptioncode01 = cdl
			FItemList(i).Foptioncode02 = rsget("optioncode02")
			FItemList(i).Fcodevalue = db2html(rsget("codevalue"))
			FItemList(i).Fcodeview = db2html(rsget("codeview"))

			FItemList(i).Foptiondispyn02 = rsget("optiondispyn")
			FItemList(i).Fdisporder02 = rsget("disporder")
		end if

		rsget.close
	end sub


	public sub GetOption01List()
		dim sqlStr,i

		sqlStr = "select * from [db_item].[dbo].tbl_option_div01 with (nolock)"
		if FRectOnlyUsing<>"" then
			sqlStr = sqlStr + " where optiondispyn='Y'"
		end if

		if FRectOrderType="d" then
			sqlStr = sqlStr + " order by disporder"
		else
			sqlStr = sqlStr + " order by optioncode01"
		end if

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COptionManagerItem

				FItemList(i).Foptioncode01 = rsget("optioncode01")
				FItemList(i).Fcodename = db2html(rsget("codename"))
				FItemList(i).Foptiondispyn01 = rsget("optiondispyn")
				FItemList(i).Fdisporder01 = rsget("disporder")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetOption02List(cdl)
		dim sqlStr,i

		sqlStr = "select * from [db_item].[dbo].tbl_option_div02 with (nolock)"
		sqlStr = sqlStr + " where optioncode01='" + cdl + "'"
		if FRectOnlyUsing<>"" then
			sqlStr = sqlStr + " and optiondispyn='Y'"
		end if

		if FRectOrderType="d" then
			sqlStr = sqlStr + " order by disporder"
		else
			sqlStr = sqlStr + " order by optioncode02"
		end if

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COptionManagerItem
				FItemList(i).Foptioncode01 = cdl
				FItemList(i).Foptioncode02 = rsget("optioncode02")
				FItemList(i).Fcodevalue = db2html(rsget("codevalue"))
				FItemList(i).Fcodeview = db2html(rsget("codeview"))

				FItemList(i).Foptiondispyn02 = rsget("optiondispyn")
				FItemList(i).Fdisporder02 = rsget("disporder")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetOption01Select()
		dim sqlStr,i
		sqlStr = "select optioncode01,codename from [db_item].[dbo].tbl_option_div01 with (nolock)"
		if FRectOnlyUsing<>"" then
			sqlStr = sqlStr + " where optiondispyn='Y'"
		end if

		if FRectOrderType="d" then
			sqlStr = sqlStr + " order by disporder"
		else
			sqlStr = sqlStr + " order by optioncode01"
		end if

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COptionManagerItem

				FItemList(i).FCode01 = rsget("optioncode01")
				FItemList(i).FCodeName = db2html(rsget("codename"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetOption02Select(byval cdl)
		dim sqlStr,i
		sqlStr = "select optioncode01,optioncode02,codeview from [db_item].[dbo].tbl_option_div02"
		sqlStr = sqlStr + " where optioncode01='" + cdl + "'"
		if FRectOnlyUsing<>"" then
			sqlStr = sqlStr + " and optiondispyn='Y'"
		end if

		if FRectOrderType="d" then
			sqlStr = sqlStr + " order by disporder"
		else
			sqlStr = sqlStr + " order by optioncode02"
		end if

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COptionManagerItem

				FItemList(i).FCode01 = rsget("optioncode01")
				FItemList(i).FCode02 = rsget("optioncode02")
				FItemList(i).FCodeName = db2html(rsget("codeview"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CCatemanageItem
	public Fcdlarge
	public Fcdmid
	public Fcdsmall
	public Fchannel
	public Fnmlarge
	public Fnmmid
	public Fnmsmall
	public Fdescription
	public Fimgtitle
	public Fimgmain
	public Fimgon
	public Fimgoff
	public Fcatecnt
	public Fextcatecnt
	public ForderNo

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CSimpleItem
	public FItemId
	public FItemName
	public Fmakerid
	public FImgSmall
	public FSellyn
	public Fdispyn
	public Fisusing

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCatemanager
	public FItemList()
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FtotalPage
	public FRectDispSailYN

	public Sub GetCategoryKeyword(cdl,cdm,cds)
		dim sqlStr,i
		sqlStr = "select top 1 keyword from [db_item].[dbo].tbl_item_small"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " and code_mid='" + cdm + "'"
		sqlStr = sqlStr + " and code_small='" + cds + "'"

		rsget.Open sqlStr, dbget, 1

		redim preserve FItemList(0)
		if not rsget.Eof then
			set FItemList(0) = new COptionManagerItem

			FItemList(0).Fkeyword = rsget("keyword")

		end if

		rsget.close
	end sub

	public sub GetOrgCateMaster()
		dim sqlStr,i
		sqlStr = "select code_large, code_nm from [db_item].[dbo].tbl_item_large"
		sqlStr = sqlStr + " order by code_large"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetOrgCateMasterMid(cdl)
		dim sqlStr,i
		sqlStr = "select code_large, code_mid, code_nm from [db_item].[dbo].tbl_item_mid"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " order by code_mid"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fcdmid          = rsget("code_mid")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetOrgCateMasterSmall(cdl,cdm)
		dim sqlStr,i
		sqlStr = "select s.code_large, s.code_mid, s.code_small, s.code_nm, IsNULL(T.cnt,0) as catecnt"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_small s"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select itemserial_small, count(itemid) as cnt from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	where itemserial_large='" + cdl + "'"
		sqlStr = sqlStr + " 	and itemserial_mid='" + cdm + "'"
		sqlStr = sqlStr + " 	and itemid not in ("
		sqlStr = sqlStr + " 	select itemid from [db_temp].[dbo].tbl_temp_itemcategory"
		sqlStr = sqlStr + " 	)"
		sqlStr = sqlStr + "		group by itemserial_small"
		sqlStr = sqlStr + "	) as T on s.code_small=T.itemserial_small"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " and code_mid='" + cdm + "'"
		sqlStr = sqlStr + " order by code_small"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fcdmid          = rsget("code_mid")
				FItemList(i).Fcdsmall          = rsget("code_small")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))
				FItemList(i).Fcatecnt        = rsget("catecnt")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetOrgCateItemList(cdl,cdm,cds)
		dim sqlStr,i

		sqlStr = "select top " + CStr(FPageSize) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.dispyn, i.isusing, m.imgsmall "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_image m"
		sqlStr = sqlStr + " on i.itemid=m.itemid"
		sqlStr = sqlStr + " where i.itemserial_large='" + cdl + "'"
		sqlStr = sqlStr + " and i.itemserial_mid='" + cdm + "'"
		sqlStr = sqlStr + " and i.itemserial_small='" + cds + "'"
		sqlStr = sqlStr + " and i.itemid not in ("
		sqlStr = sqlStr + " 	select itemid from [db_temp].[dbo].tbl_temp_itemcategory"
		sqlStr = sqlStr + " )"

		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsget("itemid")
				FItemList(i).FItemName  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")

				FItemList(i).FSellyn    = rsget("sellyn")
				FItemList(i).Fdispyn    = rsget("dispyn")
				FItemList(i).Fisusing   = rsget("isusing")

				FItemList(i).FImgSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("imgsmall")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.Close
	end sub

	public sub GetOrgCateNotMachItemList()
		dim sqlStr,i
		sqlStr = "select count(itemid) as cnt from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.itemid<>0"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.itemid not in ("
		sqlStr = sqlStr + " 	select itemid from [db_temp].[dbo].tbl_temp_itemcategory"
		sqlStr = sqlStr + " )"
		rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.dispyn, i.isusing, m.imgsmall "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_image m"
		sqlStr = sqlStr + " on i.itemid=m.itemid"
		sqlStr = sqlStr + " where i.itemid<>0"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.itemid not in ("
		sqlStr = sqlStr + " 	select itemid from [db_temp].[dbo].tbl_temp_itemcategory"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + "  order by i.itemid desc "

		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsget("itemid")
				FItemList(i).FItemName  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")

				FItemList(i).FSellyn    = rsget("sellyn")
				FItemList(i).Fdispyn    = rsget("dispyn")
				FItemList(i).Fisusing   = rsget("isusing")

				FItemList(i).FImgSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("imgsmall")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.Close
	end sub

	public sub GetNewCateMaster()
		dim sqlStr,i
		sqlStr = "select code_large, code_nm from [db_item].[dbo].tbl_item_large"
		sqlStr = sqlStr + " order by code_large"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetNewCateMasterMid(cdl)
		dim sqlStr,i
		sqlStr = "select code_large, code_mid, code_nm,orderNo from [db_item].[dbo].tbl_item_mid"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " order by orderNo ,code_mid"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fcdmid          = rsget("code_mid")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))
				FItemList(i).FOrderNo				=rsget("orderNo")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub


	public sub GetNewCateMasterSmall(cdl,cdm)
		dim sqlStr,i
		sqlStr = "select s.code_large, s.code_mid, s.code_small, s.code_nm, orderNo ,IsNULL(T.cnt,0) as catecnt"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_small s"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select itemserial_small, count(itemid) as cnt from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	where itemserial_large='" + cdl + "'"
		sqlStr = sqlStr + " 	and itemserial_mid='" + cdm + "'"
		sqlStr = sqlStr + "		group by itemserial_small"
		sqlStr = sqlStr + "	) as T on s.code_small=T.itemserial_small"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		sqlStr = sqlStr + " and code_mid='" + cdm + "'"
		sqlStr = sqlStr + " order by orderNo, code_small"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCatemanageItem

				FItemList(i).Fcdlarge          = rsget("code_large")
				FItemList(i).Fcdmid          = rsget("code_mid")
				FItemList(i).Fcdsmall          = rsget("code_small")
				FItemList(i).Fnmlarge        = db2html(rsget("code_nm"))
				FItemList(i).Fcatecnt        = rsget("catecnt")
				FItemList(i).FOrderNo        = rsget("orderNo")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public function GetNewCateCurrentPos(cdl,cdm,cds)
		dim sqlStr
		sqlStr = "select distinct top 1 code_nm "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_large"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		rsget.Open sqlStr, dbget, 1
		if not rsget.Eof then
			GetNewCateCurrentPos = db2html(rsget("code_nm"))
		end if
		rsget.close


		if cdm<>"" then
			sqlStr = "select distinct top 1 code_nm "
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_mid"
			sqlStr = sqlStr + " where code_large='" + cdl + "'"
			sqlStr = sqlStr + " and code_mid='" + cdm + "'"
			rsget.Open sqlStr, dbget, 1
			if not rsget.Eof then
				GetNewCateCurrentPos = GetNewCateCurrentPos + "-" +  db2html(rsget("code_nm"))
			end if
			rsget.close
		end if

		if cds<>"" then
			sqlStr = "select distinct top 1 code_nm "
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_small"
			sqlStr = sqlStr + " where code_large='" + cdl + "'"
			sqlStr = sqlStr + " and code_mid='" + cdm + "'"
			sqlStr = sqlStr + " and code_small='" + cds + "'"
			rsget.Open sqlStr, dbget, 1
			if not rsget.Eof then
				GetNewCateCurrentPos = GetNewCateCurrentPos + "-" + db2html(rsget("code_nm"))
			end if
			rsget.close
		end if

	end function

	public sub GetNewCateItemList(cdl,cdm,cds)
		dim sqlStr,i

		sqlStr = "select count(i.itemid) as cnt "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_image m"
		sqlStr = sqlStr + " on i.itemid=m.itemid"
		sqlStr = sqlStr + " where i.itemserial_large='" + cdl + "'"
		sqlStr = sqlStr + " and i.itemserial_mid='" + cdm + "'"
		sqlStr = sqlStr + " and i.itemserial_small='" + cds + "'"
		if FRectDispSailYN = "on" then
		sqlStr = sqlStr + " and i.sellyn='Y'"
		sqlStr = sqlStr + " and i.dispyn='Y'"
		end if

		rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellyn, i.dispyn, i.isusing, m.imgsmall "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_image m"
		sqlStr = sqlStr + " on i.itemid=m.itemid"
		sqlStr = sqlStr + " where i.itemserial_large='" + cdl + "'"
		sqlStr = sqlStr + " and i.itemserial_mid='" + cdm + "'"
		sqlStr = sqlStr + " and i.itemserial_small='" + cds + "'"
		if FRectDispSailYN = "on" then
		sqlStr = sqlStr + " and i.sellyn='Y'"
		sqlStr = sqlStr + " and i.dispyn='Y'"
		end if


		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSimpleItem

				FItemList(i).FItemId  = rsget("itemid")
				FItemList(i).FItemName  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")

				FItemList(i).FSellyn    = rsget("sellyn")
				FItemList(i).Fdispyn    = rsget("dispyn")
				FItemList(i).Fisusing   = rsget("isusing")

				FItemList(i).FImgSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("imgsmall")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FtotalPage = 1
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