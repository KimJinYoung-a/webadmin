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
    public FADDIMAGE
    
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
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item_addimage m"
		sqlStr = sqlStr & " where m.itemid='" & itemid & "'"
		sqlStr = sqlStr & " and m.imgtype=1"
		sqlStr = sqlStr & " order by m.gubun"
 
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		
        FResultCount = rsACADEMYget.RecordCount
        FTotalCount  = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
		    do until rsACADEMYget.EOF

    			set FItemList(i) = new CInfoImageItem
                FItemList(i).FIDX		= rsACADEMYget("idx")
                FItemList(i).FITEMID	= rsACADEMYget("ITEMID")
                FItemList(i).FIMGTYPE	= rsACADEMYget("IMGTYPE")
                FItemList(i).FGUBUN		= rsACADEMYget("GUBUN")
                FItemList(i).FADDIMAGE	= imgFingers & "/diyItem/contentsimage/" + GetImageSubFolderByItemid(FItemList(i).FITEMID) + "/" + rsACADEMYget("ADDIMAGE") 
                
    		    rsACADEMYget.movenext
    		    i=i+1
			loop
		end if
		rsACADEMYget.Close
		

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
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
		sqlStr = sqlStr & " ,db_academy.dbo.tbl_diy_item_addimage m" + vbcrlf
		sqlStr = sqlStr & " where i.itemid=m.itemid" 
		sqlStr = sqlStr & " and m.imgtype=1" + vbcrlf
		sqlStr = sqlStr & " and m.gubun=1" + vbcrlf
		
		sqlStr = sqlStr & whereSQL

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		''#################################################
		''데이타
		''#################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage) & "" + vbcrlf
		sqlStr = sqlStr & " i.itemid,i.makerid, i.itemname,i.smallimage,m.addimage as infoimage" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
		sqlStr = sqlStr & " ,db_academy.dbo.tbl_diy_item_addimage m" + vbcrlf
		sqlStr = sqlStr & " where i.itemid=m.itemid" 
		sqlStr = sqlStr & " and m.imgtype=1" + vbcrlf
		sqlStr = sqlStr & " and m.gubun=1" + vbcrlf
		sqlStr = sqlStr & whereSQL
		sqlStr = sqlStr & " order by i.itemid Desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
		    do until rsACADEMYget.EOF

    			set FItemList(i) = new CInfoImageList
    
    			FItemList(i).Fitemid    = rsACADEMYget("itemid")
    			FItemList(i).FItemName  = db2html(rsACADEMYget("itemname"))
    			FItemList(i).FMakerid   = rsACADEMYget("makerid")
    			FItemList(i).Fsmall_img = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimage")
    		    rsACADEMYget.movenext
    		    i=i+1
			loop
		end if
		rsACADEMYget.Close
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