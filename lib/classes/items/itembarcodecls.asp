<%
class CItemBarCodeSubItem
	public FItemGubun
	public FItemID
	public FItemOption

	public Fmakerid
	public FbrandName
	public FItemName
	public FItemOptionName
	public FSellcash
	public Fitemrackcode
	public FImageSmall
	public FImageList

	public FOpt1Name
	public FOpt2Name
	public FOptionusing

	public FPublicBarcode
    public Fprtidx
    
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class


class CItemBarCode
	public FItemList()
	public FOneItem

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectItemGubun
	public FRectItemID
	public FRectItemoption
    
	public Sub getItemBarcodeInfo()
		dim sqlStr,i
		
		IF (FRectItemGubun="10") THEN
    		sqlStr = "select i.itemid, i.itemname, i.makerid, i.sellcash, i.brandname, i.smallimage, i.listimage, i.itemrackcode," + VbCrlf
    		sqlStr = sqlStr + " IsNULL(v.itemoption,'0000') as itemoption, '' as opt1name," + VbCrlf
    		sqlStr = sqlStr + " IsNULL(v.optionname,'') as opt2name, IsNULL(v.isusing,'N') as optusing, " + VbCrlf
    		sqlStr = sqlStr + " s.barcode, c.prtidx "
    		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + VbCrlf
    		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid" 
    		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid"
    		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s on (s.itemgubun='10') and i.itemid=s.itemid  and IsNULL(s.itemoption,'0000')=IsNULL(v.itemoption,'0000')" + VbCrlf
    		sqlStr = sqlStr + " where i.itemid=" + CStr(FRectItemID) + VbCrlf
        ELSE
            sqlStr = "select i.shopitemid as itemid, i.shopitemname as itemname, i.makerid, i.shopitemprice as sellcash, c.socname as brandname, i.offimgsmall as smallimage, i.offimglist as listimage, '9999' as itemrackcode," + VbCrlf
    		sqlStr = sqlStr + " i.itemoption as itemoption, '' as opt1name," + VbCrlf
    		sqlStr = sqlStr + " i.shopitemoptionname opt2name, i.isusing as optusing, " + VbCrlf
    		sqlStr = sqlStr + " s.barcode, c.prtidx "
    		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i" + VbCrlf
    		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid" 
    		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s on (s.itemgubun='" + FRectItemGubun + "') and s.itemid=i.shopitemid and  s.itemoption=i.itemoption" + VbCrlf
    		sqlStr = sqlStr + " where i.shopitemid=" + CStr(FRectItemID) + VbCrlf
    		sqlStr = sqlStr + " and i.itemgubun='" + FRectItemGubun + "'" 
        END IF
''response.write  sqlStr   
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim  FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemBarCodeSubItem
				FItemList(i).FItemGubun  = FRectItemGubun
				FItemList(i).FItemID     = rsget("itemid")
				FItemList(i).FItemOption = rsget("itemoption")

				FItemList(i).Fmakerid        = rsget("makerid")
				FItemList(i).FbrandName      = db2html(rsget("brandname"))
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FSellcash       = rsget("sellcash")
				FItemList(i).Fitemrackcode	 = rsget("itemrackcode")
                
                if (FRectItemGubun="10") then
    				FItemList(i).FImageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FRectItemID) + "/" + rsget("smallimage")
    				FItemList(i).FImageList      = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FRectItemID) + "/" + rsget("listimage")
                else
                    FItemList(i).FImageSmall     = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")
                    FItemList(i).FImageList     = "http://webimage.10x10.co.kr/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("listimage")
                end if
            
				FItemList(i).FOpt1Name = rsget("opt1name")
				FItemList(i).FOpt2Name = rsget("opt2name")
				FItemList(i).FItemOptionName = FItemList(i).FOpt2Name

				FItemList(i).FOptionusing = rsget("optusing")
				FItemList(i).FPublicBarcode = rsget("barcode")
				
				FItemList(i).Fprtidx        = rsget("prtidx")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)

		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

%>