<%
class CInfoImageList
    public Fitemid
    public FItemName
    public FMakerid
    public Fsmall_img
    public FinfoimgCount
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CInfoImageItem
    public FIDX
    public FITEMID
    public FIMGTYPE
    public FGUBUN
    public FADDIMAGE_400
    
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CInfoImage
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FMakerid
	public FItemid
	
    Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage = 1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
    
    public Sub getOneInfoImageList(byval itemid)
		dim sqlStr, i

		''#################################################
		''데이타
		''#################################################

		sqlStr = "select top 10 " + vbcrlf
		sqlStr = sqlStr & " m.*" + vbcrlf
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_addimage m"
		sqlStr = sqlStr & " where m.itemid='" & itemid & "'"
		sqlStr = sqlStr & " and m.imgtype=1"
		sqlStr = sqlStr & " order by m.gubun"
 
		rsget.Open sqlStr,dbget,1
		
        FResultCount = rsget.RecordCount
        FTotalCount  = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
		    do until rsget.EOF

    			set FItemList(i) = new CInfoImageItem
                FItemList(i).FIDX          = rsget("idx")
                FItemList(i).FITEMID       = rsget("ITEMID")
                FItemList(i).FIMGTYPE      = rsget("IMGTYPE")
                FItemList(i).FGUBUN        = rsget("GUBUN")
                FItemList(i).FADDIMAGE_400 = "http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemList(i).FITEMID) + "/" + rsget("ADDIMAGE_400") 
                
    		    rsget.movenext
    		    i=i+1
			loop
		end if
		rsget.Close
		

	end sub
	
	public Sub getInfoImageList()
		dim sqlStr,whereSQL
		dim i

		'###########################################################################
		'상품 총 갯수 구하기
		'###########################################################################


		if FMakerid<>"" then
			whereSQL = whereSQL + " and i.makerid='" & FMakerid & "'"
		end if

		if FItemid<>"" then
			whereSQL = whereSQL + " and i.itemid='" & FItemid & "'"
		end if

		sqlStr = "select count(m.itemid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " ,[db_item].[dbo].tbl_item_addimage m" + vbcrlf
		sqlStr = sqlStr & " where i.itemid=m.itemid" 
		sqlStr = sqlStr & " and m.imgtype=1" + vbcrlf
		sqlStr = sqlStr & " and m.gubun=1" + vbcrlf
		sqlStr = sqlStr & " and i.itemdiv<>21" + vbcrlf
		
		sqlStr = sqlStr & whereSQL

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		''#################################################
		''데이타
		''#################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage) & "" + vbcrlf
		sqlStr = sqlStr & " i.itemid,i.makerid, i.itemname,i.smallimage,m.addimage_400 as infoimage" + vbcrlf
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " ,[db_item].[dbo].tbl_item_addimage m" + vbcrlf
		sqlStr = sqlStr & " where i.itemid=m.itemid" 
		sqlStr = sqlStr & " and m.imgtype=1" + vbcrlf
		sqlStr = sqlStr & " and m.gubun=1" + vbcrlf
		sqlStr = sqlStr & " and i.itemdiv<>21" + vbcrlf
		sqlStr = sqlStr & whereSQL
		sqlStr = sqlStr & " order by i.itemid Desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
		    do until rsget.EOF

    			set FItemList(i) = new CInfoImageList
    
    			FItemList(i).Fitemid    = rsget("itemid")
    			FItemList(i).FItemName  = db2html(rsget("itemname"))
    			FItemList(i).FMakerid   = rsget("makerid")
    			FItemList(i).Fsmall_img = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")
    		    rsget.movenext
    		    i=i+1
			loop
		end if
		rsget.Close
	end sub


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