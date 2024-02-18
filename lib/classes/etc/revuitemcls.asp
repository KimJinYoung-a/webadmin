<%
class CRevuOneItem
	public FItemID
	public FItemName
    public FMakerid
    
	public Fitemserial_large
	public Fitemserial_mid
	public Fitemserial_small
    
    public Fitemserial_largeNm
	public Fitemserial_midNm
	public Fitemserial_smallNm
	
	public Fsourcearea
	public FMakerName
    public FBrandName
    
	public FSellCash
	public Forgsellcash
	public FSuplyCash
	public Fkeywords

	public FListImage
	public FSmallImage
	public FBasicImage
	public Fmainimage
	public Ficon1Image
	public Ficon2Image

	public FSellyn

	public FDesigner

	public FRegdate

	public FLinkCode
	public FItemOption
	public FItemOptionName
	public FItemOptionGubunName

	public FItemContent
	public Fordercomment

	public FUpDate

	public Flimityn
	public Flimitno
	public Flimitsold

	public FSailDispNo
    public Fvatinclude
    
    public Fitemsize 
    public Fitemsource 
    
    public function getItemLink()
        getItemLink = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & FItemID & "&rdsite=revu"
    end function

    public function IsSoldOut()
        IsSoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
    end function

	public function getSellStrNo()
		if (FSellyn="N") then
			getSellStrNo = "3"
		elseif (FSellyn="S") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1)) then
			getSellStrNo = "2"
		else
			getSellStrNo = "1"
		end if
	end function

	public function getkeywords()
		getkeywords = Fkeywords
	end function

	public function get400Image()
		get400Image = ""

		if IsNULL(FBasicImage) or (FBasicImage="") then Exit function

		get400Image = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemID) + "/" + FBasicImage
	end function
    
    public function getItemPreInfodataHTML() 
        dim reStr
        
        if Not IsNULL(Fordercomment) then 
            if Fordercomment<>"" then
                reStr = "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
            end if
        end if
        
        if Fitemsize<>"" then
            reStr = reStr & "- 사이즈 : " & Fitemsize & "<br>"
        end if
        
        if Fitemsource<>"" then
            reStr = reStr & "- 재료 : " &  Fitemsource & "<br>"
        end if
        
        getItemPreInfodataHTML = reStr
    end function

    public function getItemInfoImageHTML()
        getItemInfoImageHTML = ""
        
        if (FItemID=121680) then
            getItemInfoImageHTML = "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo1_121680.jpg' width='600'>"
            getItemInfoImageHTML = getItemInfoImageHTML & "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo2_121680.jpg' width='600'>"
            getItemInfoImageHTML = getItemInfoImageHTML & "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo3_121680.jpg' width='600'>"
            getItemInfoImageHTML = getItemInfoImageHTML & "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo4_121680.jpg' width='600'>"
            Exit function
        end if
        
        if IsNULL(Fmainimage) or (Fmainimage="") then Exit function
        ''if (FMakerid<>"hueplane") then Exit function
        
        getItemInfoImageHTML = "<br><img src=http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemID) + "/" + Fmainimage + ">"
        
    end function

	public function get160Image()
		get160Image = ""

		if IsNULL(Ficon1Image) or (Ficon1Image="") then Exit function

		get160Image = "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(FItemID) + "/" + Ficon1Image

	end function

	public function get85Image()
		get85Image = ""

		if IsNULL(FListImage) or (FListImage="") then Exit function

		get85Image = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemID) + "/" + FListImage
	end function

	public function get60Image()
		get60Image = ""

		if IsNULL(FSmallImage) or (FSmallImage="") then Exit function

		get60Image = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemID) + "/" + FSmallImage
	end function

	public function getAsDeliverInfo()
		getAsDeliverInfo = Fordercomment
	end function

	public function getItemContent()
		getItemContent = FItemContent
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class



Class CRevuItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectDesigner



	public FRectNoRegNate
	public FBufArr
	public FBufOptArr
	public FBufSellcashArr

	public FRectStartItemID

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
    
    
	
	public sub GetAllRevuItemTotalPageRecent()
		dim sqlStr,i
		sqlStr = "select count(i.itemid) as cnt from "
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " ,[db_item].[dbo].tbl_item_Contents c" + vbcrlf
		sqlStr = sqlStr + " where i.itemid=c.itemid"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.sellyn='Y'"
		sqlStr = sqlStr + " and datediff(d,i.regdate,getdate())<21"
        sqlStr = sqlStr + " and datediff(d,i.regdate,getdate())>=7"
        sqlStr = sqlStr + " and c.sellcount>20"
        sqlStr = sqlStr + " and i.danjongyn='N'"
        sqlStr = sqlStr + " and i.cate_large in ('010','020','030')"
        ''sqlStr = sqlStr + " and itemserial_large in ('10','40','25')"
		sqlStr = sqlStr + " and i.basicimage is not null"
		sqlStr = sqlStr + " and i.sellcash>0"
		
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		rsget.Close
	end sub


	public sub GetAllRevuItemListRecent()
		dim sqlStr,i

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.sellcash, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.itemcontent, c.ordercomment,"
		sqlStr = sqlStr + " i.smallimage, i.listimage, i.basicimage, i.icon1image, i.icon2image,"
		sqlStr = sqlStr + " i.regdate ," + vbcrlf
		sqlStr = sqlStr + " i.lastupdate as oregdate,  c.usinghtml, "
		sqlStr = sqlStr + " v.nmlarge, v.nmmid, v.nmsmall"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_Contents c" + vbcrlf
		sqlStr = sqlStr + " ,[db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join db_item.dbo.vw_category v"
		sqlStr = sqlStr + "     on i.cate_large=v.cdlarge"
		sqlStr = sqlStr + "     and i.cate_mid=v.cdmid"
		sqlStr = sqlStr + "     and i.cate_small=v.cdsmall"
		sqlStr = sqlStr + " where i.itemid=c.itemid"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.sellyn='Y'"
		sqlStr = sqlStr + " and datediff(d,i.regdate,getdate())<21"
        sqlStr = sqlStr + " and datediff(d,i.regdate,getdate())>=7"
        sqlStr = sqlStr + " and c.sellcount>20"
        sqlStr = sqlStr + " and i.danjongyn='N'"
        sqlStr = sqlStr + " and i.cate_large in ('010','020','030')"
        ''sqlStr = sqlStr + " and itemserial_large in ('10','40','25')"
		sqlStr = sqlStr + " and i.basicimage is not null"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " order by i.itemid "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CRevuOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				FItemList(i).FUpdate  = rsget("oregdate")

				FItemList(i).Fsellyn  = rsget("sellyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fitemserial_large = rsget("cate_large")
				FItemList(i).Fitemserial_mid = rsget("cate_mid")
				FItemList(i).Fitemserial_small = rsget("cate_small")
				
                FItemList(i).Fitemserial_largeNm = db2html(rsget("nmlarge"))
				FItemList(i).Fitemserial_midNm = db2html(rsget("nmmid"))
				FItemList(i).Fitemserial_smallNm = db2html(rsget("nmsmall"))
                
				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).Flistimage  = rsget("listimage")
				FItemList(i).Fsmallimage  = rsget("smallimage")
				FItemList(i).Ficon1image  = rsget("icon1image")
				FItemList(i).Ficon2image  = rsget("icon2image")

				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).Fordercomment = db2html(rsget("ordercomment"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub


	

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