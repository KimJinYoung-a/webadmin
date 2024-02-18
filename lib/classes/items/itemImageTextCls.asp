<%

Class CItemImageTextItem
    public FitemId
	public Freq_yyyymmdd
	public Ffin_yyyymmdd
	public Fimagetext
	public Fmodifiedtext
	public Fupdatecnt
	public Freguserid
	public Flastuserid
	public Fregdate
	public Flastupdate

    public FitemName
    public FmakerId
    public FsmallImage
    public FlistImage

    Private Sub Class_Initialize()
		''
	End Sub

	Private Sub Class_Terminate()
		''
	End Sub
end Class

Class CItemImageText
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectItemId
	public FRectMakerId

	public Sub GetItemImageTextOne()
		dim sqlStr, i

		sqlStr = ""
		sqlStr = sqlStr + " select top 1 t.itemid, t.req_yyyymmdd, t.fin_yyyymmdd, t.imagetext, t.updatecnt, t.reguserid, t.lastuserid, t.regdate, t.lastupdate, i.itemName, i.makerId, i.smallImage, i.listImage, t.modifiedtext "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_contents].[dbo].[tbl_itemImageText] t "
		sqlStr = sqlStr + " 	join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		t.itemid = i.itemid "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + " and t.itemid = " & FRectItemId

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		if Not rsget.Eof then
			set FOneItem = new CItemImageTextItem

			FOneItem.Fitemid	= rsget("itemid")
			FOneItem.Freq_yyyymmdd	= rsget("req_yyyymmdd")
			FOneItem.Ffin_yyyymmdd	= rsget("fin_yyyymmdd")
			FOneItem.Fimagetext	= db2html(rsget("imagetext"))
			FOneItem.Fmodifiedtext	= db2html(rsget("modifiedtext"))
			FOneItem.Fupdatecnt	= rsget("updatecnt")
			FOneItem.Freguserid	= rsget("reguserid")
			FOneItem.Flastuserid	= rsget("lastuserid")
			FOneItem.Fregdate	= rsget("regdate")
			FOneItem.Flastupdate	= rsget("lastupdate")
			FOneItem.FitemName	= db2html(rsget("itemName"))
			FOneItem.FmakerId	= rsget("makerId")
			FOneItem.FsmallImage	= rsget("smallImage")
			FOneItem.FlistImage	= rsget("listImage")

			if ((Not IsNULL(FOneItem.Fsmallimage)) and (FOneItem.Fsmallimage<>"")) then FOneItem.Fsmallimage	= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fsmallimage
			if ((Not IsNULL(FOneItem.Flistimage)) and (FOneItem.Flistimage<>"")) then FOneItem.Flistimage	 = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Flistimage
		end if

		rsget.Close
	end sub

	public function GetItemImageTextList()
        dim sqlStr, addSql, i

		addSql = ""
		if (FRectItemId <> "") then
			addSql = addSql + " 	and t.itemid = " & FRectItemId
		end if
		if (FRectMakerId <> "") then
			addSql = addSql + " 	and i.makerId = '" & FRectMakerId & "' "
		end if

		sqlStr = ""
		sqlStr = sqlStr + " select Count(t.itemid), CEILING(CAST(Count(t.itemid) AS FLOAT)/" & FPageSize & ") "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_contents].[dbo].[tbl_itemImageText] t "
		sqlStr = sqlStr + " 	join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		t.itemid = i.itemid "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + addSql
		''response.write sqlStr

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
        rsget.Close


		sqlStr = ""
		sqlStr = sqlStr + " select top " & Cstr(FPageSize * FCurrPage) & " t.itemid, t.req_yyyymmdd, t.fin_yyyymmdd, t.imagetext, t.updatecnt, t.reguserid, t.lastuserid, t.regdate, t.lastupdate, i.itemName, i.makerId, i.smallImage, i.listImage, t.modifiedtext "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_contents].[dbo].[tbl_itemImageText] t "
		sqlStr = sqlStr + " 	join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		t.itemid = i.itemid "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + addSql
		sqlStr = sqlStr + " order by req_yyyymmdd, t.itemid desc "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if Not(rsget.EOF or rsget.BOF) then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemImageTextItem

                FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).Freq_yyyymmdd	= rsget("req_yyyymmdd")
				FItemList(i).Ffin_yyyymmdd	= rsget("fin_yyyymmdd")
				FItemList(i).Fimagetext	= rsget("imagetext")
				FItemList(i).Fmodifiedtext	= rsget("modifiedtext")
				FItemList(i).Fupdatecnt	= rsget("updatecnt")
				FItemList(i).Freguserid	= rsget("reguserid")
				FItemList(i).Flastuserid	= rsget("lastuserid")
				FItemList(i).Fregdate	= rsget("regdate")
				FItemList(i).Flastupdate	= rsget("lastupdate")
				FItemList(i).FitemName	= db2html(rsget("itemName"))
				FItemList(i).FmakerId	= rsget("makerId")
				FItemList(i).FsmallImage	= rsget("smallImage")
				FItemList(i).FlistImage	= rsget("listImage")

				if ((Not IsNULL(FItemList(i).Fsmallimage)) and (FItemList(i).Fsmallimage<>"")) then FItemList(i).Fsmallimage    = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Fsmallimage
				if ((Not IsNULL(FItemList(i).Flistimage)) and (FItemList(i).Flistimage<>"")) then FItemList(i).Flistimage    = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Flistimage

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage 		= 1
		FPageSize 		= 10
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
	End Sub

	Private Sub Class_Terminate()
		''
    End Sub
end Class

%>
