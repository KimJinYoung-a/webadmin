<%
'#######################################################
'	History	:  2015.11.05 한용민 생성
'	Description : 포장 서비스 클래스
'/배송비 상품코드:100 , 옵션코드:1000
'#######################################################

class Cpack_item
	public fmidx
	public Fuserid
	public Ftitle
	public Fmessage
	public Fpackitemcnt
	public Fregdate
	public fcancelyn
	public FItemID
	public FItemOption
	public FItemEa
	public FItemOptionName
	public FItemName
	public FImageSmall
	public FImageList
	public FBrandName
	public FMakerId
	public fdcancelyn

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Cpack
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FShoppingBagItemCount

	public FRectCancelyn
	public FRectSort
	public FRectOrderSerial
	public FRectuserid
	public frectmidx
	public Fpackitemcnt
	public Fpackcnt

	'//cscenter/pojang_view.asp
	public function Getpojang_itemlist()
		dim sqlStr, sqlsearch, tmpmidx

		if FRectOrderSerial="" then exit function
		
		if FRectOrderSerial<>"" then
			sqlsearch = sqlsearch & " and pm.orderserial='"& FRectOrderSerial &"'"
		end if

		sqlStr = "select pd.itemid, pd.itemoption, pd.itemno, pd.cancelyn as dcancelyn, i.itemname"
		sqlStr = sqlStr & " ,(select itemoptionname from [db_order].[dbo].[tbl_order_detail]"
		sqlStr = sqlStr & " 	where orderserial = '" & FRectOrderSerial & "' and itemid = pd.itemid and itemoption = pd.itemoption) as itemoptionname,"
		sqlStr = sqlStr & " i.smallimage, i.listimage120, i.icon1image, i.brandname, i.makerid"
		sqlStr = sqlStr & " ,pm.midx, pm.userid, isnull(pm.title,'''') as title, isnull(pm.message,'''') as message, pm.packitemcnt, pm.regdate, pm.cancelyn"
		sqlStr = sqlStr & " from [db_order].[dbo].[tbl_order_pack_detail] as pd"
		sqlStr = sqlStr & " join db_order.[dbo].[tbl_order_pack_master] pm"
		sqlStr = sqlStr & " 	on pd.midx=pm.midx"
		sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item] as i"
		sqlStr = sqlStr & " 	on pd.itemid = i.itemid"
		sqlStr = sqlStr & " where pm.packitemcnt>0 " & sqlsearch
		sqlStr = sqlStr & " order by pd.midx desc, pd.didx desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

    	FResultCount = rsget.RecordCount
    	if (FResultCount<1) then FResultCount=0

    	redim FItemList(FResultCount)
    	i=0

    	do Until rsget.Eof
			set FItemList(i) = new Cpack_item

			FItemList(i).fmidx	    = rsget("midx")
			FItemList(i).Fuserid	    = rsget("userid")
			FItemList(i).Ftitle    = db2html(rsget("title"))
			FItemList(i).Fmessage = db2html(rsget("message"))
			FItemList(i).Fpackitemcnt      = rsget("packitemcnt")
			FItemList(i).Fregdate      = rsget("regdate")
			FItemList(i).fcancelyn      = rsget("cancelyn")
			FItemList(i).FItemID	    = rsget("itemid")
			FItemList(i).FItemOption    = rsget("itemoption")
			FItemList(i).FItemEa		= rsget("itemno")
				'/총상품수량합계
				Fpackitemcnt = Fpackitemcnt + FItemList(i).FItemEa

			FItemList(i).fdcancelyn	    = rsget("dcancelyn")
			FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
			FItemList(i).FItemName	    = db2html(rsget("itemname"))
			FItemList(i).FImageSmall		= "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallimage")
			'FItemList(i).FImageList			= "http://webimage.10x10.co.kr/image/List/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("listimage")
			'FItemList(i).FImageList		= "http://webimage.10x10.co.kr/image/List120/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("listimage120")
			'FItemList(i).FImageList		= "http://webimage.10x10.co.kr/image/icon1/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("icon1image")
			FItemList(i).FBrandName		= db2html(rsget("brandname"))
			FItemList(i).FMakerId		= rsget("makerid")

			'/총박스수
			if cstr(tmpmidx)<>cstr(FItemList(i).fmidx) then
				Fpackcnt = Fpackcnt + 1
			end if

			tmpmidx = FItemList(i).fmidx
            i=i+1
    		rsget.movenext
    	loop
    	rsget.Close
	end function

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FShoppingBagItemCount = 0
		FScrollCount = 10
		FTotalCount = 0
		Fpackitemcnt=0
		Fpackcnt=0
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