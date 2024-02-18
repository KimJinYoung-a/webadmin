<%
'############################## 어드민 제휴몰관리 메뉴의 Class ##############################
Class CExtSitemPrdItem
    public Fitemid    
    public Fregdate   
    public Freguserid	
    public Fitemname  
    public Fmakerid   
    public Fsmallimage
    public Fsellcash  
    public Fbuycash   
    
    public FSellyn
    public FLimity
    public Limitno
    public Limitsold
    
    public FExtRegDate
    public FExtLastUpDate
    
    public FExtSiteItemno
    
    public FExtStoreSeq       
    public FExtdispcategory   
    public FExtstorecategory  
    
    public Fatseoulcategory


    public FSailYn
    public FOrgPrice
    
    public function GetprdPrefixStr()
        if (FSailYn="Y") and (FOrgPrice>Fsellcash) then
            GetprdPrefixStr = CStr(CLng(FOrgPrice-Fsellcash/FOrgPrice*100)) + "% 할인중" 
        else
            GetprdPrefixStr = " "
        end if
    end function

    function getExtStoreSeqName
        if IsNULL(FExtStoreSeq) then Exit Function
        
        if (FExtStoreSeq=2) then  
            getExtStoreSeqName = "리빙"
        elseif (FExtStoreSeq=3) then  
            getExtStoreSeqName = "잡화"    
        elseif (FExtStoreSeq=4) then  
            getExtStoreSeqName = "의류"  
        end if
    end function
    
    
    public function IsSoldOut()
        IsSoldOut = (FSellyn<>"Y") or ((FLimity="Y") and (Limitno-Limitsold<1))
    end function

    Private Sub Class_Initialize()

	End Sub


	Private Sub Class_Terminate()

	End Sub
end Class


Class CExtSiteItem
    public FOneItem
    public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectItemId
	public FRectMakerId
	public FRectItemName
	public FRectEventid
	public FRectIsSoldOut
	
	public FRectExtNotReg
	public FRectMatchCate
	
	public FRectCate_large
	public FRectCate_mid
	public FRectCate_small
	
	public FRectNotMatchCategory
    public FRectExtItemID
    public FRectMinusMigin
    
    
    
    public Sub GetAtSeoulRegedItemList()
        dim i,sqlStr
                      
        sqlStr = "select count(s.itemid) as cnt " + vbcrlf
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_atseoul_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		
		if FRectEventid<>"" then 
		    sqlStr = sqlStr + " left join [db_event].[dbo].tbl_eventitem e "
		    sqlStr = sqlStr + " on e.evt_code=" + CStr(FRectEventid)
		    sqlStr = sqlStr + " and e.itemid=i.itemid"
		end if
		    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		
		if FRectMakerid<>"" then
		    sqlStr = sqlStr + " and i.makerid='" & FRectMakerid & "'"
		end if
		
		if FRectItemId<>"" then
		    sqlStr = sqlStr + " and s.itemid=" + CStr(FRectItemId) + vbcrlf
		end if
		
		if FRectItemName<>"" then
		    sqlStr = sqlStr + " and i.itemname like '%" + CStr(FRectItemName) + "%'" + vbcrlf
		end if
		
		if FRectEventid<>"" then
		    sqlStr = sqlStr + " and e.evt_code is Not NULL"
		end if
		
	    ''sqlStr = sqlStr + " and i.basicimage is not null"
		''sqlStr = sqlStr + " and i.itemdiv<50"
		''sqlStr = sqlStr + " and i.itemserial_large<90"
		''sqlStr = sqlStr + " and i.sellcash>0"
		
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.Close


        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + vbcrlf
        sqlStr = sqlStr + " s.itemid, s.regdate, s.reguserid" + vbcrlf
        sqlStr = sqlStr + " ,i.itemname, i.smallimage, i.sellcash, i.buycash, i.makerid " + vbcrlf
        sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold " + vbcrlf
        sqlStr = sqlStr + " ,m.category"
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_atseoul_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_atseoul_category_mapping m " + vbcrlf
	    sqlStr = sqlStr + " on i.cate_large=m.tencdl " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_mid=m.tencdm " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_small=m.tencdn " + vbcrlf

	    if FRectEventid<>"" then 
		    sqlStr = sqlStr + " left join [db_event].[dbo].tbl_eventitem e "
		    sqlStr = sqlStr + " on e.evt_code=" + CStr(FRectEventid)
		    sqlStr = sqlStr + " and e.itemid=i.itemid"
		end if
		
	    sqlStr = sqlStr + " where s.itemid=i.itemid"
	    
	    if FRectMakerid<>"" then
		    sqlStr = sqlStr + " and i.makerid='" & FRectMakerid & "'"
		end if
		
	    if FRectItemId<>"" then
		    sqlStr = sqlStr + " and s.itemid=" + CStr(FRectItemId) + vbcrlf
		end if
		
		if FRectItemName<>"" then
		    sqlStr = sqlStr + " and i.itemname like '%" + CStr(FRectItemName) + "%'" + vbcrlf
		end if
		
		if FRectEventid<>"" then
		    sqlStr = sqlStr + " and e.evt_code is Not NULL"
		end if
		
	    ''sqlStr = sqlStr + " and i.basicimage is not null"
		''sqlStr = sqlStr + " and i.itemdiv<50"
		''sqlStr = sqlStr + " and i.itemserial_large<90"
		''sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " order by s.regdate desc"
		
		rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtSitemPrdItem

				FItemList(i).Fitemid    = rsget("itemid")
				FItemList(i).Fregdate   = rsget("regdate")
				FItemList(i).Freguserid	= rsget("reguserid")
				FItemList(i).Fitemname  = db2html(rsget("itemname"))
				FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).Fsmallimage = "http://webimage.10x10.co.kr/image/small/" + getImageSubFolderByItemId(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Fsellcash   = rsget("sellcash")
                FItemList(i).Fbuycash    = rsget("buycash")
                
                FItemList(i).FSellyn    = rsget("sellyn")
                FItemList(i).FLimity    = rsget("limityn")
                FItemList(i).Limitno    = rsget("limitno")
                FItemList(i).Limitsold  = rsget("limitsold")
                
                FItemList(i).Fatseoulcategory     = rsget("category")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
 
    end Sub
    
    Private Sub Class_Initialize()
	    redim FItemList(0)
		FCurrPage =1
		FPageSize = 5
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub


	Private Sub Class_Terminate()

	End Sub
	
    '// 이전 페이지 검사 //
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function


	'// 다음 페이지 검사 //
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function


	'// 초기 페이지 반환 //
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
	
end Class

%>


<%
'############################## XML 생성 파일 관련 Class ##############################
class CTTLOneItem
	public FItemID
	public FItemName
    public FMakerid
    
	public Fcate_large
	public Fcate_mid
	public Fcate_small

	public Fsourcearea
	public Fitemweight
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
    
    public FInfoImage
    
	public FSellyn
	public FDispyn

	public FDesigner

	public FRegdate

	public FLinkCode
	public FOptionTypeName
	public FItemOption
	public FItemOptionName
	public FItemOptionGubunName

	public FItemContent
	public Fordercomment

	public FUpDate

	public Flimityn
	public Flimitno
	public Flimitsold
	public Fstockqty

	public FSailDispNo
    public Fvatinclude
    
	public FTTLCode
    public Fatseoulcategory 

    
    public Fitemsize 
    public Fitemsource 
    
    public Foptsellyn
    public Foptlimityn
    public Foptlimitno
    public Foptlimitsold
    public Foptaddprice
    
    public FLastUpdate
    public FSellEndDate
    
    public FInfoImage1
    public FInfoImage2
    public FInfoImage3
    public FInfoImage4
    
    public FDetailImage
    public FDetailImage1
    public FDetailImage2
    public FDetailImage3
    public FDetailImage4
    
    public Fitemdiv
    
    public function getdeliCode()
        if (Fitemdiv<>"01") then getdeliCode="7"
    end function

    public function GetSellEndDateStr()
        GetSellEndDateStr = "99991231"
        
        if IsNULL(FSellEndDate) then Exit function
        
        FSellEndDate = Replace(Left(CStr(FSellEndDate),10),"-","")
    end function

    public function GetRealSellprice()
        if (Foptaddprice>0) then
            GetRealSellprice = FSellcash + Foptaddprice
        else
            GetRealSellprice = FSellcash
        end if
    end function

    public function IsOptionSoldOut()
        const CLIMIT_SOLDOUT_NO = 5
        
        IsOptionSoldOut = false
        if (FItemOption="0000") then Exit function
        
        ''옵션추가 금액이 있는것은 뺌
        IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno-Foptlimitsold<CLIMIT_SOLDOUT_NO)) or (Foptaddprice>0)
        
        
    end function

    public function IsSoldOut()
        const CLIMIT_SOLDOUT_NO = 5
        
        IsSoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<CLIMIT_SOLDOUT_NO))
    end function

	public function getSellStrNo()
		if (FDispyn="N") or (FSellyn="N") then
			getSellStrNo = "3"
		elseif ((FLimitYn="Y") and (FLimitNo-FLimitSold<1)) then
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
        dim splited, i, cnt, oneimageName
        
        getItemInfoImageHTML = ""

        if FDetailImage <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage + ">"
        end if
        
        if FDetailImage1 <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage1 + ">"
        end if
        
        if FDetailImage2 <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage2 + ">"
        end if
        
        if FDetailImage3 <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage3 + ">"
        end if
        
        if FDetailImage4 <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage4 + ">"
        end if
        
        
        exit function

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

Class CTTLSpreadItem
	public FGOODS_NM
	public FGOODS_STK_NO
	public FORG_TP
	public FGOOD_COST
	public FSALE_PR
	public FADD_TAX_RT
	public FSPE_TAX_RT
	public FAREA_LMT_YN
	public FDELV_AREA_TP
	public FSALE_STR_DM

	public FSALE_END_DM
	public FSALE_STAT_CL
	public FDISP_YN
	public FMK_ENTR_NO

	public FGOODS_OPT_YN
	public FNO_INT_QUOTA_MON
	public FUSE_YN
	public FZ_ADDTAX_YN
	public fZ_DELV_FEE_TP
	public fZ_MK_ENTR
	public FUPD_DM
	public FREG_DM
	public Flnkurl
	public Fimgurl


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CTTLItem
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
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
    
    public sub GetAllAtSeoulItemTotalPage()
		dim sqlStr,i
		sqlStr = "select count(s.itemid) as cnt from "
		sqlStr = sqlStr + " [db_item].[dbo].tbl_atseoul_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_atseoul_category_mapping m " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=m.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=m.tencdm " + vbcrlf
	    sqlStr = sqlStr + " 	and i.cate_small=m.tencdn " + vbcrlf
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid" + vbcrlf 
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<7) or (dateDiff(d,i.lastupdate,getdate())<7) or (dateDiff(d,s.lastupdate,getdate())<6))" + vbcrlf
		sqlStr = sqlStr + " and i.basicimage is not null" + vbcrlf
		sqlStr = sqlStr + " and i.itemdiv<50" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>''" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>'999'" + vbcrlf
		sqlStr = sqlStr + " and i.sellcash>0" + vbcrlf
		sqlStr = sqlStr + " and m.category is Not NULL" + vbcrlf
		sqlStr = sqlStr + " and i.isExtusing = 'Y'" + VbCrlf
		
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		rsget.Close
	end sub
	
	public sub GetAllAtSeoulItemList4()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid,"
		sqlStr = sqlStr + " IsNULL(t.shopitemprice,i.sellcash+IsNULL(o.optaddprice,0)) as sellcash,"
		sqlStr = sqlStr + " IsNULL(t.orgsellprice,i.orgprice+IsNULL(o.optaddprice,0)) as orgprice,"
		sqlStr = sqlStr + " IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.itemdiv, "
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.lastupdate,  c.usinghtml,m.category, i.itemWeight, "
		sqlStr = sqlStr + " o.optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  isNull((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=1),'') as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  isNull((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=2),'') as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  isNull((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=3),'') as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  isNull((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=4),'') as infoimage4" + vbcrlf
		sqlStr = sqlStr + " ,  isNull((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1),'') as detailimage1" + vbcrlf
        sqlStr = sqlStr + " ,  isNull((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2),'') as detailimage2" + vbcrlf
        sqlStr = sqlStr + " ,  isNull((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3),'') as detailimage3" + vbcrlf
        sqlStr = sqlStr + " ,  isNull((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4),'') as detailimage4" + vbcrlf

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_atseoul_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_atseoul_category_mapping m " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=m.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=m.tencdm " + vbcrlf
	    sqlStr = sqlStr + " 	and i.cate_small=m.tencdn " + vbcrlf
	    ''오프라인 테이블 추가
	    sqlStr = sqlStr + " 	left Join db_shop.dbo.tbl_shop_item t on i.itemid=t.shopitemid and t.itemgubun='10' and o.itemoption=t.itemoption"
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<7) or (dateDiff(d,i.lastupdate,getdate())<7) or (dateDiff(d,s.lastupdate,getdate())<6))" + vbcrlf
		sqlStr = sqlStr + " and i.basicimage is not null"
		sqlStr = sqlStr + " and i.itemdiv<50"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<>'999'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and i.isExtusing = 'Y'"
		''sqlStr = sqlStr + " and IsNULL(t.shopitemprice,i.sellcash+IsNULL(o.optaddprice,0))>0"
		sqlStr = sqlStr + " and m.category is Not NULL"
		
		sqlStr = sqlStr + " order by i.itemid desc, o.itemoption"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTTLOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				'FItemList(i).Fsellyn  = rsget("sellyn")
				If rsget("sellyn") <> "Y" Then
					FItemList(i).Fsellyn = "0"
				Else
					FItemList(i).Fsellyn = "1"
				End IF
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")
				
				If rsget("limityn") = "Y" Then
					FItemList(i).Fstockqty = rsget("limitno") - rsget("limitsold")
				Else
					FItemList(i).Fstockqty = "999"
				End If

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				
				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if
				
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).FOptionTypeName = rsget("optiontypename")
				If rsget("optiontypename") = "" Then
					FItemList(i).FOptionTypeName = "option"
				End If
				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("basicimage")
				FItemList(i).Fmainimage   = rsget("mainimage")
				FItemList(i).FDetailImage   = rsget("mainimage")
				
				'if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"
				    
				'FItemList(i).Flistimage  = rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				'FItemList(i).Ficon1image  = rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")
                
                FItemList(i).Fordercomment = db2html(rsget("ordercomment"))
                
				FItemList(i).FItemContent = db2html(rsget("itemcontent"))
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"","")
				
				'if (FItemList(i).Fitemid=112016) then
				'    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,"본 상품은 무료배송 상품입니다.","")
				'end if
				
				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))
				FItemList(i).Fitemweight  = rsget("itemWeight")
				
				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				
				
				FItemList(i).Fatseoulcategory     = rsget("category")
				
				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))
				
				FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Foptaddprice  = rsget("optaddprice")
				
				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).FSellEndDate  = rsget("sellenddate")
				
				FItemList(i).FInfoImage1  = rsget("InfoImage1")
				If rsget("InfoImage1") <> "" Then
					FItemList(i).FInfoImage1  = "http://webimage.10x10.co.kr/image/add1/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("InfoImage1")
				End IF
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				If rsget("InfoImage2") <> "" Then
					FItemList(i).FInfoImage2  = "http://webimage.10x10.co.kr/image/add2/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("InfoImage2")
				End IF
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				If rsget("InfoImage3") <> "" Then
					FItemList(i).FInfoImage3  = "http://webimage.10x10.co.kr/image/add3/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("InfoImage3")
				End IF
				FItemList(i).FInfoImage4  = rsget("InfoImage4")
				If rsget("InfoImage4") <> "" Then
					FItemList(i).FInfoImage4  = "http://webimage.10x10.co.kr/image/add4/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("InfoImage4")
				End IF
				
				FItemList(i).FDetailImage1  = rsget("detailimage1")
				FItemList(i).FDetailImage2  = rsget("detailimage2")
				FItemList(i).FDetailImage3  = rsget("detailimage3")
				FItemList(i).FDetailImage4  = rsget("detailimage4")
				
				FItemList(i).Fitemdiv  = rsget("itemdiv")
				
				
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