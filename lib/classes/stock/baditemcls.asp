<%

Class CBadItemItem
	public FItemgubun
	public FItemId
	public FItemOption
	public Fitemno
	public FItemOptionName

	public FItemName
	public Fmakerid
	public Fsellcash
	public FBuycash
	public Fmwdiv
	public Fdeliverytype

	public function GetMwDivName()
		if Fmwdiv="M" then
			GetMwDivName = "¸ÅÀÔ"
		elseif Fmwdiv="W" then
			GetMwDivName = "À§Å¹"
		elseif Fmwdiv="U" then
			GetMwDivName = "¾÷Ã¼"
		end if
	end function

	public function GetdeliverytypeName()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			GetdeliverytypeName = "¾÷¹è"
		else
			GetdeliverytypeName = "ÅÙ¹è"
		end if
	end function

	Private Sub Class_Initialize()
		Fitemno = 0
	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CBadItem
	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
        public FPageCount
	public FTotalCount
	public FResultCount
        public FScrollCount
	public FCurrPage
	public FRectItemID

	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage       = 1
		FPageSize       = 50
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     = 0

		FRectItemID     = ""
	end sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount - 1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount + 1
	end Function

	public sub GetTempItemList()
		dim sqlstr, i
		sqlstr = " select top 1000 t.*, i.makerid, i.itemname, i.deliverytype, i.sellcash , i.buycash, i.mwdiv, "
		sqlstr = sqlstr + " IsNULL(v.optionname,'') as codeview "
		sqlstr = sqlstr + " from [db_summary].[dbo].tbl_temp_baditem t "
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item i on t.itemgubun='10' and t.itemid=i.itemid "
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option v on t.itemoption=v.itemoption and t.itemid=v.itemid"
		sqlstr = sqlstr + " where 1 = 1 "

'		if FRectItemID<>"" then
'			sqlstr = sqlstr + " and t.itemid=" + CStr(FRectItemID)
'		end if

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i = 0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CBadItemItem

				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fitemid            = rsget("itemid")
				FItemList(i).Fitemoption        = rsget("itemoption")
				FItemList(i).Fitemno            = rsget("itemno")

				FItemList(i).FItemName	        = db2html(rsget("itemname"))
				FItemList(i).FMakerid		= rsget("makerid")
				FItemList(i).Fsellcash		= rsget("sellcash")
				FItemList(i).Fbuycash		= rsget("buycash")
				FItemList(i).Fmwdiv		= rsget("mwdiv")
				FItemList(i).Fdeliverytype      = rsget("deliverytype")
				FItemList(i).FItemOptionName    = db2html(rsget("codeview"))

                                if isnull(FItemList(i).Fsellcash) then
                                        FItemList(i).Fsellcash = 0
                                end if

				i = i + 1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub
end Class

%>

