<%
CONST CMAXMARGIN = 14 

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
    
    public Fdnshopmngcategory  
    public Fdnshopdispcategory 
    public Fdnshopstorecategory
    
    public FmayiParkPrice
    public FmayiParkSellYn
    public FitemLastupdate

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

Class CInterParkOneCategory
    public FCate_Large
    public FCate_Mid
    public FCate_Small
    public Fnmlarge
    public FnmMid
    public FnmSmall
    public Finterparkdispcategory
    public Finterparkstorecategory
    public Fdnshopdispcategory
    public Fdnshopstorecategory
    public Fdnshopecategory
    public Fdnshopmngcategory
    
    public FdnshopRcategory
    public FdnshopSpkey
    public FdnshopSeCategory
    
    public FItemCnt
    public FinterparkdispcategoryText
    public FinterparkstorecategoryText
    
    public FSupplyCtrtSeq
    public FIparkCateDispyn
    
    function getSupplyCtrtSeqName
        if IsNULL(FSupplyCtrtSeq) then Exit Function
        
        if (FSupplyCtrtSeq=2) then  
            getSupplyCtrtSeqName = "리빙"
        elseif (FSupplyCtrtSeq=3) then  
            getSupplyCtrtSeqName = "잡화"    
        elseif (FSupplyCtrtSeq=4) then  
            getSupplyCtrtSeqName = "의류"  
        end if
    end function

    function IsNotMatchedDispcategory
        IsNotMatchedDispcategory = IsNULL(Finterparkdispcategory) or (Finterparkdispcategory="")
    end function
    
    function IsNotMatchedStorecategory
        IsNotMatchedStorecategory = IsNULL(Finterparkstorecategory) or (Finterparkstorecategory="")
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
    public FRectMinusMigin15
    
    public FRectExpensive10x10
    public FRectInteryes10x10no
    public FRectOnreginotmapping
    
    public FTemp
    
	public FDelJaeHyu
    
    
    public Sub GetOneInterParkCategoryMaching()
        dim sqlStr,i
        
        sqlStr = "select  top 1 "
        sqlStr = sqlStr + " i.cate_large as tencdl, i.cate_mid as tencdm, i.cate_small as tencdn,"
        sqlStr = sqlStr + " c.nmlarge, c.nmmid, c.nmsmall,  p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory,"
        sqlStr = sqlStr + " ts.storecatename, tp.dispcatename"
        sqlStr = sqlStr + "  from "
        sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr + "     left join [db_item].[dbo].vw_category c "
        sqlStr = sqlStr + "     on i.cate_large=c.cdlarge"
        sqlStr = sqlStr + "     and i.cate_mid=c.cdmid"
        sqlStr = sqlStr + "     and i.cate_small=c.cdsmall"
        
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p"
        sqlStr = sqlStr + "     on i.cate_large=p.tencdl"
        sqlStr = sqlStr + "     and i.cate_mid=p.tencdm"
        sqlStr = sqlStr + "     and i.cate_small=p.tencdn"
        
            
        sqlStr = sqlStr + "     left join [db_temp].dbo.tbl_interpark_Tmp_StoreCategory ts"
        sqlStr = sqlStr + "     on p.interparkstorecategory=ts.storecatecode"
        
        sqlStr = sqlStr + "     left join [db_temp].dbo.tbl_interpark_Tmp_DispCategory tp"
        sqlStr = sqlStr + "     on p.interparkdispcategory=tp.dispcatecode"
        
        sqlStr = sqlStr + " where i.cate_large='" + FRectCate_large + "'"
        sqlStr = sqlStr + " and i.cate_mid='" + FRectCate_mid + "'"
        sqlStr = sqlStr + " and i.cate_small='" + FRectCate_small + "'"
        
        
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
        
        FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
        if (FResultCount<1) then FResultCount=0
        
		i=0
		if  not rsget.EOF  then
			set FOneItem = new CInterParkOneCategory
			FOneItem.FCate_Large             = rsget("tencdl") 
            FOneItem.FCate_Mid               = rsget("tencdm") 
            FOneItem.FCate_Small             = rsget("tencdn")
            FOneItem.Fnmlarge                = db2Html(rsget("nmlarge"))
            FOneItem.FnmMid                  = db2Html(rsget("nmMid"))
            FOneItem.FnmSmall                = db2Html(rsget("nmSmall"))
            FOneItem.Finterparkdispcategory  = rsget("interparkdispcategory")
            FOneItem.Finterparkstorecategory = rsget("interparkstorecategory")
            FOneItem.FSupplyCtrtSeq          = rsget("SupplyCtrtSeq")
            
            FOneItem.FinterparkdispcategoryText  = db2Html(rsget("dispcatename"))
            FOneItem.FinterparkstorecategoryText = db2Html(rsget("storecatename"))
		end if
		rsget.Close
    end Sub
    
    public Sub GetInterParkCategoryMachingList()
        dim sqlStr,i
        
        sqlStr = "select  "
        sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, count(i.itemid) as ItemCnt,"
        sqlStr = sqlStr + " c.nmlarge, c.nmmid, c.nmsmall,  p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory, t.dispyn as IparkCateDispyn"
        sqlStr = sqlStr + "  from "
        sqlStr = sqlStr + " [db_item].[dbo].tbl_interpark_reg_item d"
        sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr + "     on d.itemid=i.itemid"
        if (FRectCate_large<>"") then
            sqlStr = sqlStr + "     and i.cate_large='" & FRectCate_large & "'"
        end if
        sqlStr = sqlStr + "     left join [db_item].[dbo].vw_category c "
        sqlStr = sqlStr + "     on i.cate_large=c.cdlarge"
        sqlStr = sqlStr + "     and i.cate_mid=c.cdmid"
        sqlStr = sqlStr + "     and i.cate_small=c.cdsmall"
            
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p"
        sqlStr = sqlStr + "     on i.cate_large=p.tencdl"
        sqlStr = sqlStr + "     and i.cate_mid=p.tencdm"
        sqlStr = sqlStr + "     and i.cate_small=p.tencdn"
        
        sqlStr = sqlStr + "     left join [db_temp].[dbo].tbl_interpark_Tmp_DispCategory t"
        sqlStr = sqlStr + "     on p.interparkdispcategory=t.DispCateCode"
        
        sqlStr = sqlStr + " where 1=1"
        if (FRectNotMatchCategory="on") then
            sqlStr = sqlStr + " and ((p.interparkdispcategory is NULL) or (p.interparkdispcategory='') or (p.interparkstorecategory is NULL) or (p.interparkstorecategory=''))"
        end if
        sqlStr = sqlStr + " group by i.cate_large, i.cate_mid, i.cate_small,c.nmlarge, c.nmmid, c.nmsmall, p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory, t.dispyn"
        sqlStr = sqlStr + " order by  i.cate_large, i.cate_mid, i.cate_small"
        
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
        
        FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneCategory
				FItemList(i).FCate_Large             = rsget("Cate_Large") 
                FItemList(i).FCate_Mid               = rsget("Cate_Mid") 
                FItemList(i).FCate_Small             = rsget("Cate_Small")
                FItemList(i).FItemCnt                = rsget("ItemCnt")
                FItemList(i).Fnmlarge                = db2Html(rsget("nmlarge"))
                FItemList(i).FnmMid                  = db2Html(rsget("nmMid"))
                FItemList(i).FnmSmall                = db2Html(rsget("nmSmall"))
                FItemList(i).Finterparkdispcategory  = rsget("interparkdispcategory")
                FItemList(i).Finterparkstorecategory = rsget("interparkstorecategory")
                FItemList(i).FSupplyCtrtSeq          = rsget("SupplyCtrtSeq")
                
                FItemList(i).FIparkCateDispyn        = rsget("IparkCateDispyn")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub
    
	public Sub GetInterParkRegedItemList()
	    dim i,sqlStr
	    sqlStr = "select count(s.itemid) as cnt " + vbcrlf
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + " on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_small=p.tencdn " + vbcrlf
	    
		if FRectEventid<>"" then 
		    sqlStr = sqlStr + " left join [db_event].[dbo].tbl_eventitem e "
		    sqlStr = sqlStr + " on e.evt_code=" + CStr(FRectEventid)
		    sqlStr = sqlStr + " and e.itemid=i.itemid"
		end if
		    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		
		if (FRectMinusMigin<>"") then
		    sqlStr = sqlStr + " and i.sellcash<>0"
		    sqlStr = sqlStr + " and ((i.sellcash-i.buycash)/i.sellcash)*100<11" + VbCrlf
		end if
		
		if (FRectMinusMigin15<>"") then
		    sqlStr = sqlStr + " and i.sellcash<>0"
		    sqlStr = sqlStr + " and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN & VbCrlf
		end if
		
		if (FRectExtItemID<>"") then
		    sqlStr = sqlStr + " and s.interparkPrdNo='" & FRectExtItemID & "'" 
		end if
		
		if (FRectExtNotReg="M") then
		    sqlStr = sqlStr + " and s.interparkregdate is NULL"
		elseif (FRectExtNotReg="F") then
		    sqlStr = sqlStr + " and s.interparkregdate is Not NULL"
		elseif (FRectExtNotReg="R") then
		    sqlStr = sqlStr + " and s.interparkregdate is Not NULL"
		    sqlStr = sqlStr + " and s.interparklastupdate<i.lastupdate"
		    'sqlStr = sqlStr + " and ((i.lastupdate<'2010-04-12 00:00:00') or (i.lastupdate>'2010-04-12 03:00:00'))"
		    sqlStr = sqlStr + " and ((i.lastupdate<>'2010-04-22 00:09:48.693'))" '' or ((i.lastupdate='2008-10-21 00:06:16.140') and (i.sellyn<>'Y')) )"
		end if
		
		if (FRectMatchCate="Y") then
		    sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL"
		    sqlStr = sqlStr + " and p.interparkstorecategory is Not NULL"
		elseif (FRectMatchCate="N") then
		    sqlStr = sqlStr + " and (p.interparkdispcategory is NULL or p.interparkstorecategory is NULL)"
		end if
		
		if FRectMakerid<>"" then
		    sqlStr = sqlStr + " and i.makerid='" & FRectMakerid & "'"
		end if
		
		if FRectItemId<>"" then
		    sqlStr = sqlStr + " and s.itemid in(" + CStr(FRectItemId) + ")" + vbcrlf
		end if
		
		if FRectItemName<>"" then
		    sqlStr = sqlStr + " and i.itemname like '%" + CStr(FRectItemName) + "%'" + vbcrlf
		end if
		
		if FRectEventid<>"" then
		    sqlStr = sqlStr + " and e.evt_code is Not NULL"
		end if
		
		if FRectIsSoldOut<>"" then
		    sqlStr = sqlStr + " and i.sellyn<>'Y'"
		end if
		
		if FRectExpensive10x10 <> "" then
		    sqlStr = sqlStr + " and s.mayiParkPrice is Not Null and i.sellcash > s.mayiParkPrice "
		end if
		
		if FRectInteryes10x10no <> "" then
		    sqlStr = sqlStr + " and s.mayiParkPrice is Not Null and s.mayiParkSellYn = 'Y' and i.sellyn = 'N' "
		end if
		
		if FRectOnreginotmapping <> "" then
		    sqlStr = sqlStr + " and s.interParkPrdNo is Not Null and (p.interparkdispcategory is NULL or p.interparkstorecategory is NULL) "
		end if
		
	    ''sqlStr = sqlStr + " and i.basicimage is not null"
		''sqlStr = sqlStr + " and i.itemdiv<50"
		''sqlStr = sqlStr + " and i.itemserial_large<90"
		''sqlStr = sqlStr + " and i.sellcash>0"
		
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.Close


        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + vbcrlf
        sqlStr = sqlStr + " s.itemid, s.regdate, s.reguserid, s.interparkregdate, s.interparklastupdate, s.interParkPrdNo" + vbcrlf
        sqlStr = sqlStr + " ,i.itemname, i.smallimage, i.sellcash, i.buycash, i.makerid, i.sailyn, i.orgprice " + vbcrlf
        sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold, i.lastupdate " + vbcrlf
        sqlStr = sqlStr + " ,p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory " + vbcrlf
        sqlStr = sqlStr + " ,s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory "
        sqlStr = sqlStr + " ,s.mayiParkPrice, s.mayiParkSellYn"
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		
        sqlStr = sqlStr + " left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + " on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_small=p.tencdn " + vbcrlf
	    
	    if FRectEventid<>"" then 
		    sqlStr = sqlStr + " left join [db_event].[dbo].tbl_eventitem e "
		    sqlStr = sqlStr + " on e.evt_code=" + CStr(FRectEventid)
		    sqlStr = sqlStr + " and e.itemid=i.itemid"
		end if
		
	    sqlStr = sqlStr + " where s.itemid=i.itemid"
	    
	    if (FRectMinusMigin<>"") then
	        sqlStr = sqlStr + " and i.sellcash<>0"
		    sqlStr = sqlStr + " and ((i.sellcash-i.buycash)/i.sellcash)*100<11" + VbCrlf
		end if
		
		if (FRectMinusMigin15<>"") then
		    sqlStr = sqlStr + " and i.sellcash<>0"
		    sqlStr = sqlStr + " and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN & VbCrlf
		end if
		
	    if (FRectExtItemID<>"") then
		    sqlStr = sqlStr + " and s.interparkPrdNo='" & FRectExtItemID & "'" 
		end if
		
	    if (FRectExtNotReg="M") then
		        sqlStr = sqlStr + " and s.interparkregdate is NULL"
		elseif (FRectExtNotReg="F") then
		    sqlStr = sqlStr + " and s.interparkregdate is Not NULL"
		elseif (FRectExtNotReg="R") then
		        sqlStr = sqlStr + " and s.interparkregdate is Not NULL"
		        sqlStr = sqlStr + " and s.interparklastupdate<i.lastupdate"
		        'sqlStr = sqlStr + " and ((i.lastupdate<'2010-04-12 00:00:00') or (i.lastupdate>'2010-04-12 03:00:00'))"
		        sqlStr = sqlStr + " and ((i.lastupdate<>'2010-04-22 00:09:48.693'))" '' or ((i.lastupdate='2008-10-21 00:06:16.140') and (i.sellyn<>'Y')) )"
		end if
		
		if (FRectMatchCate="Y") then
		    sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL"
		    sqlStr = sqlStr + " and p.interparkstorecategory is Not NULL"
		elseif (FRectMatchCate="N") then
		    sqlStr = sqlStr + " and (p.interparkdispcategory is NULL or p.interparkstorecategory is NULL)"
		end if
		
	    if FRectMakerid<>"" then
		    sqlStr = sqlStr + " and i.makerid='" & FRectMakerid & "'"
		end if
		
		if FRectItemId<>"" then
		    sqlStr = sqlStr + " and s.itemid in(" + CStr(FRectItemId) + ")" + vbcrlf
		end if
		
		if FRectItemName<>"" then
		    sqlStr = sqlStr + " and i.itemname like '%" + CStr(FRectItemName) + "%'" + vbcrlf
		end if
		
		if FRectEventid<>"" then
		    sqlStr = sqlStr + " and e.evt_code is Not NULL"
		end if
		
		if FRectIsSoldOut<>"" then
		    sqlStr = sqlStr + " and i.sellyn<>'Y'"
		end if
		
		if FRectExpensive10x10 <> "" then
		    sqlStr = sqlStr + " and s.mayiParkPrice is Not Null and i.sellcash > s.mayiParkPrice "
		end if
		
		if FRectInteryes10x10no <> "" then
		    sqlStr = sqlStr + " and s.mayiParkPrice is Not Null and s.mayiParkSellYn = 'Y' and i.sellyn = 'N' "
		end if
		
		if FRectOnreginotmapping <> "" then
		    sqlStr = sqlStr + " and s.interParkPrdNo is Not Null and (p.interparkdispcategory is NULL or p.interparkstorecategory is NULL) "
		end if
		
	    ''sqlStr = sqlStr + " and i.basicimage is not null"
		''sqlStr = sqlStr + " and i.itemdiv<50"
		''sqlStr = sqlStr + " and i.itemserial_large<90"
		''sqlStr = sqlStr + " and i.sellcash>0"
		''sqlStr = sqlStr + " order by s.regdate desc"
		
		if FRectIsSoldOut<>"" then
		    'sqlStr = sqlStr + " order by i.itemid "
		    sqlStr = sqlStr + " order by s.interparklastupdate "
	    else
    		if FRectEventid<>"" then 
    		    sqlStr = sqlStr + " order by i.itemid desc"
    		else
    		    sqlStr = sqlStr + " order by s.regdate desc"
    		end if
		end if
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
                
                FItemList(i).FExtRegDate        = rsget("interparkregdate")
                FItemList(i).FExtLastUpdate     = rsget("interparklastupdate")
                
                FItemList(i).FExtSiteItemno    = rsget("interParkPrdNo")
                
                if IsNULL(rsget("interParkSupplyCtrtSeq")) then
                    FItemList(i).FExtStoreSeq       = rsget("SupplyCtrtSeq")
                else
                    FItemList(i).FExtStoreSeq    = rsget("interParkSupplyCtrtSeq")
                end if
                
                if IsNULL(rsget("regedInterparkstorecategory")) then 
				    FItemList(i).FExtstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).FExtstorecategory   = rsget("regedInterparkstorecategory")
			    end if
                
                FItemList(i).FExtdispcategory   = rsget("interparkdispcategory")
                FItemList(i).FExtstorecategory  = rsget("interparkstorecategory")
                
                FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
                
                FItemList(i).FmayiParkPrice = rsget("mayiParkPrice")
                FItemList(i).FmayiParkSellYn = rsget("mayiParkSellYn")
                FItemList(i).FitemLastupdate = rsget("lastupdate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub
    
    public Sub GetDnshopRegedItemList()
        dim i,sqlStr
                      
        sqlStr = "select count(s.itemid) as cnt " + vbcrlf
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_dnshop_reg_item s," + vbcrlf
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
		    sqlStr = sqlStr + " and s.itemid in(" + CStr(FRectItemId) + ")" + vbcrlf
		end if
		
		if FRectItemName<>"" then
		    sqlStr = sqlStr + " and i.itemname like '%" + CStr(FRectItemName) + "%'" + vbcrlf
		end if
		
		if FRectEventid<>"" then
		    sqlStr = sqlStr + " and e.evt_code is Not NULL"
		end if
		
		If FDelJaeHyu = "o" Then
			sqlStr = sqlStr + " and i.isExtusing = 'N' "
		End IF
		
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
        sqlStr = sqlStr + " ,p.dnshopdispcategory, p.dnshopstorecategory, m.dnshopmngcategory"
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_dnshop_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_dnshop_mngcategory_mapping m " + vbcrlf
	    sqlStr = sqlStr + " on i.cate_large=m.tencdl " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_mid=m.tencdm " + vbcrlf
        
        sqlStr = sqlStr + " left join [db_item].[dbo].tbl_dnshop_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + " on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + " and i.cate_small=p.tencdn " + vbcrlf
	    
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
		    sqlStr = sqlStr + " and s.itemid in(" + CStr(FRectItemId) + ")" + vbcrlf
		end if
		
		if FRectItemName<>"" then
		    sqlStr = sqlStr + " and i.itemname like '%" + CStr(FRectItemName) + "%'" + vbcrlf
		end if
		
		if FRectEventid<>"" then
		    sqlStr = sqlStr + " and e.evt_code is Not NULL"
		end if
		
		If FDelJaeHyu = "o" Then
			sqlStr = sqlStr + " and i.isExtusing = 'N' "
		End IF
		
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
                
                FItemList(i).Fdnshopmngcategory     = rsget("dnshopmngcategory")
                FItemList(i).Fdnshopdispcategory    = rsget("dnshopdispcategory")
                FItemList(i).Fdnshopstorecategory   = rsget("dnshopstorecategory")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
 
    end Sub
    
    public Sub GetDnshopCategoryMachingList()
        dim sqlStr,i
        
        sqlStr = "select  "
        sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, count(i.itemid) as ItemCnt,"
        sqlStr = sqlStr + " c.nmlarge, c.nmmid, c.nmsmall,  p.dnshopdispcategory, p.dnshopstorecategory, p.dnshopEcategory, t.dnshopmngcategory, p.dnshopRcategory, p.dnshopSpkey, p.dnshopSeCategory "
        sqlStr = sqlStr + "  from "
        sqlStr = sqlStr + " [db_item].[dbo].tbl_dnshop_reg_item d"
        sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr + "     on d.itemid=i.itemid"
        if (FRectCate_large<>"") then
            sqlStr = sqlStr + "     and i.cate_large='" & FRectCate_large & "'"
        end if
        sqlStr = sqlStr + "     left join [db_item].[dbo].vw_category c "
        sqlStr = sqlStr + "     on i.cate_large=c.cdlarge"
        sqlStr = sqlStr + "     and i.cate_mid=c.cdmid"
        sqlStr = sqlStr + "     and i.cate_small=c.cdsmall"
            
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_dnshop_dspcategory_mapping p"
        sqlStr = sqlStr + "     on i.cate_large=p.tencdl"
        sqlStr = sqlStr + "     and i.cate_mid=p.tencdm"
        sqlStr = sqlStr + "     and i.cate_small=p.tencdn"
        
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_dnshop_mngcategory_mapping t"
        sqlStr = sqlStr + "     on i.cate_large=t.tencdl"
        sqlStr = sqlStr + "     and i.cate_mid=t.tencdm"
        
        sqlStr = sqlStr + " where 1=1"
        if (FRectNotMatchCategory="on") then
            sqlStr = sqlStr + " and ((p.dnshopdispcategory is NULL) or (p.dnshopdispcategory='') or (p.dnshopstorecategory is NULL) or (p.dnshopstorecategory='')"
            sqlStr = sqlStr + " or (p.dnshopEcategory is NULL) or (p.dnshopEcategory='')"
            sqlStr = sqlStr + " or (p.dnshopRcategory is NULL) or (p.dnshopRcategory='')"
            sqlStr = sqlStr + " or (p.dnshopSpkey is NULL) or (p.dnshopSpkey=''))"
        end if
        sqlStr = sqlStr + " group by i.cate_large, i.cate_mid, i.cate_small,c.nmlarge, c.nmmid, c.nmsmall, p.dnshopdispcategory, p.dnshopstorecategory, p.dnshopEcategory, t.dnshopmngcategory, p.dnshopRcategory, p.dnshopSpkey, p.dnshopSeCategory"
        sqlStr = sqlStr + " order by  i.cate_large, i.cate_mid, i.cate_small"
        'response.write sqlStr
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
        
        FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneCategory
				FItemList(i).FCate_Large             = rsget("Cate_Large") 
                FItemList(i).FCate_Mid               = rsget("Cate_Mid") 
                FItemList(i).FCate_Small             = rsget("Cate_Small")
                FItemList(i).FItemCnt                = rsget("ItemCnt")
                FItemList(i).Fnmlarge                = db2Html(rsget("nmlarge"))
                FItemList(i).FnmMid                  = db2Html(rsget("nmMid"))
                FItemList(i).FnmSmall                = db2Html(rsget("nmSmall"))
                FItemList(i).Fdnshopdispcategory	 = rsget("dnshopdispcategory")
                FItemList(i).Fdnshopstorecategory	 = rsget("dnshopstorecategory")
                FItemList(i).Fdnshopecategory		 = rsget("dnshopEcategory")
                FItemList(i).Fdnshopmngcategory		 = rsget("dnshopmngcategory")
				FItemList(i).FdnshopRcategory		 = rsget("dnshopRcategory")
				FItemList(i).FdnshopSpkey			 = rsget("dnshopSpkey")
				FItemList(i).FdnshopSeCategory		 = rsget("dnshopSeCategory")

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


class CInterParkOneItem
	public FItemID
	public FItemName
    public FMakerid
    
	public Fcate_large
	public Fcate_mid
	public Fcate_small

	public Fsourcearea
	public FMakerName
    public FBrandName
    public FBrandNameKor
    
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
    
	public FTTLCode
    public Fdnshopmngcategory 
    public Fdnshopdispcategory
    public Fdnshopstorecategory
    
    public Finterparkdispcategory
    public Finterparkstorecategory
    
    public Fitemsize 
    public Fitemsource 
    
    public FItemOptionTypeName
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
    
    public FSupplyCtrtSeq
    public Fisusing
    
    public FdeliveryType
    public FInterparkPrdNo
    
    public FSailYn
    public FOrgPrice
    
    public function GetprdPrefixStr()
        if (FSailYn="Y") and (FOrgPrice>Fsellcash) then
            GetprdPrefixStr = "[" + CStr(CLng((FOrgPrice-Fsellcash)/FOrgPrice*100)) + "% 할인 중]" 
        else
            GetprdPrefixStr = " "
        end if
    end function
    
   function getSupplyCtrtSeqName
        if IsNULL(FSupplyCtrtSeq) then Exit Function
        
        if (FSupplyCtrtSeq=2) then  
            getSupplyCtrtSeqName = "리빙"
        elseif (FSupplyCtrtSeq=3) then  
            getSupplyCtrtSeqName = "잡화"    
        elseif (FSupplyCtrtSeq=4) then  
            getSupplyCtrtSeqName = "의류"  
        end if
    end function
    
    public function GetSourcearea()
        if IsNULL(Fsourcearea) or (Fsourcearea="") then
           GetSourcearea = "."
        else
           GetSourcearea = Fsourcearea
        end if
        
    end function

    public function IsFreeBeasong()
        IsFreeBeasong = False
        
        if (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then
            IsFreeBeasong = True
        end if
        
        if (FSellcash>=30000) then IsFreeBeasong=True
        
    end function

    public function GetInterParkentrPoint()
        GetInterParkentrPoint = CLng(Fsellcash*0.01)
        
        if (GetInterParkentrPoint<10) then GetInterParkentrPoint=0
        
        if (Fsellcash<500) then GetInterParkentrPoint=0
        
        
    end function
    
    '' 특정브랜드 IpontMall 제외
    public function GetpointmUseYn()
        GetpointmUseYn = "Y"
        if (FMakerid="elecom") then GetpointmUseYn="N"
        
        if (GetInterParkentrPoint<1) then GetpointmUseYn="N"
    end function

    public function GetSupplyCtrtSeq()
        GetSupplyCtrtSeq = FSupplyCtrtSeq
    end function

    public function getOrderCommentStr()
        dim reStr
        reStr = ""
        
        if Not IsNULL(Fordercomment) then 
            if Fordercomment<>"" then
                reStr = "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
            end if
        end if
        
        getOrderCommentStr = reStr
    end function

    public function GetInterParkLmtQty()
        const CLIMIT_SOLDOUT_NO = 5
        
        ''Max 99999 -> 1000
        if (Flimityn="Y") then
            if (Flimitno-Flimitsold)<CLIMIT_SOLDOUT_NO then
                GetInterParkLmtQty = 0
            else
                GetInterParkLmtQty = Flimitno-Flimitsold-CLIMIT_SOLDOUT_NO
            end if
        else
            GetInterParkLmtQty = 999
        end if
    end function

    ''과세 01, 면세02, 영세 03
    public function GetInterParkTaxTp()
        if (Fvatinclude="Y") then
            GetInterParkTaxTp = "01"
        else
            GetInterParkTaxTp = "02"
        end if
    end function
    
    ''판매중01, 품절02, 판매중지03, 일시품절05
    public function GetInterParkSaleStatTp()
        if (IsSoldOut) then
            if (FSellyn="S") then
                GetInterParkSaleStatTp = "02"       ''품절(02)     SellYN-S
            else
                if (Fisusing="N") then
                    GetInterParkSaleStatTp = "05"   ''판매금지(05)
                else
                    GetInterParkSaleStatTp = "03"   ''판매중지(03) SellYN-N
                end if
            end if
        else
            GetInterParkSaleStatTp = "01"
        end if
    end function

    public function GetSellEndDateStr()
        GetSellEndDateStr = "99991231"
        
        if IsNULL(FSellEndDate) then Exit function
        
        FSellEndDate = Replace(Left(CStr(FSellEndDate),10),"-","")
    end function

    public function GetRealSellprice()
        'if (Foptaddprice>0) then
        '    GetRealSellprice = FSellcash + Foptaddprice
        'else
            GetRealSellprice = FSellcash
        'end if
    end function

    public function IsOptionSoldOut()
        const CLIMIT_SOLDOUT_NO = 5
        
        IsOptionSoldOut = false
        if (FItemOption="0000") then Exit function
        
        IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno-Foptlimitsold<CLIMIT_SOLDOUT_NO))
        
    end function
    
    public function getOptionLimitNo()
        const CLIMIT_SOLDOUT_NO = 3
        
        If (IsOptionSoldOut) then
            getOptionLimitNo = 0
        else
            if (Foptlimityn="Y") then
                if (Foptlimitno-Foptlimitsold<CLIMIT_SOLDOUT_NO) then
                    getOptionLimitNo = 0
                else
                    getOptionLimitNo = Foptlimitno-Foptlimitsold-CLIMIT_SOLDOUT_NO
                end if
            else
                getOptionLimitNo = 999
            end if
        end if
    end function

    public function IsSoldOut()
        const CLIMIT_SOLDOUT_NO = 5
        
        IsSoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<CLIMIT_SOLDOUT_NO))
    end function
    
    ''Dnshop
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
		
		'get400Image = "http://owebimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemID) + "/" + FBasicImage
		
		'if FItemid=98190 then
		'    get400Image = "http://owebimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemID) + "/" + FBasicImage
		'end if
	end function
    
    public function getItemPreInfodataHTML() 
        dim reStr
        reStr = ""
        
        reStr = "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style>"
        
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
        
        if Not (IsNULL(FInfoImage1) and (FInfoImage1<>"")) then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FInfoImage1 + ">"
        end if
        
        if Not (IsNULL(FInfoImage2) and (FInfoImage2<>"")) then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FInfoImage2 + ">"
        end if
        
        if Not (IsNULL(FInfoImage3) and (FInfoImage3<>"")) then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FInfoImage3 + ">"
        end if
        
        if Not (IsNULL(FInfoImage4) and (FInfoImage4<>"")) then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FInfoImage4 + ">"
        end if
        
        ''메인 이미지.
        if (getItemInfoImageHTML="") then
            if Not (IsNULL(Fmainimage) or (Fmainimage="")) then
                getItemInfoImageHTML = "<br><img src=http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemID) + "/" + Fmainimage + ">"
            end if
        end if
        
        '' CS 관련
        if (getItemInfoImageHTML<>"") then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src='http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg' width='546'>"
        else
            getItemInfoImageHTML = "<br><img src='http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg' width='546'>"
        end if
    
        exit function
        
        
        ''' old Style-----------------------------------------------------------
        if IsNULL(FInfoImage) or (FInfoImage="") or (FInfoImage=",,,,") then Exit function
        
        splited = split(FInfoImage,",")
        
        if IsArray(splited) then
            cnt = UBound(splited)
            for i=0 to cnt
                oneimageName = trim(splited(i))
                if (oneimageName<>"") then
                    getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + oneimageName + ">"
                end if    
            next    
        end if
            
        
'        if (FItemID=121680) then
'            getItemInfoImageHTML = "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo1_121680.jpg' width='600'>"
'            getItemInfoImageHTML = getItemInfoImageHTML & "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo2_121680.jpg' width='600'>"
'            getItemInfoImageHTML = getItemInfoImageHTML & "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo3_121680.jpg' width='600'>"
'            getItemInfoImageHTML = getItemInfoImageHTML & "<br><img src='http://webimage.10x10.co.kr/item/contentsimage/12/imginfo4_121680.jpg' width='600'>"
'            Exit function
'        end if
'        
'        if IsNULL(Fmainimage) or (Fmainimage="") then Exit function
'        ''if (FMakerid<>"hueplane") then Exit function
'        
'        getItemInfoImageHTML = "<br><img src=http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemID) + "/" + Fmainimage + ">"
'        
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




Class CiParkRegItem
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
	
    public FJaeHyuPageGubun
    public FBrandID
    
    public FTemp
    
    public sub GetIParkEditItemTotalPage()
		dim sqlStr,i
		sqlStr = "select count(s.itemid) as cnt from "
		sqlStr = sqlStr + " [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호
	    
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid" + vbcrlf 
		sqlStr = sqlStr + " and s.interparklastupdate<i.lastupdate"    
		'sqlStr = sqlStr + " and ((i.lastupdate<'2010-04-12 00:00:00') or (i.lastupdate>'2010-04-12 03:00:00'))"
		sqlStr = sqlStr + " and ((i.lastupdate<>'2010-04-22 00:09:48.693'))" '' or ((i.lastupdate='2008-10-21 00:06:16.140') and (i.sellyn<>'Y')) )"    
		sqlStr = sqlStr + " and i.basicimage is not null" + vbcrlf
		sqlStr = sqlStr + " and i.itemdiv<50" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>''" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>'999'" + vbcrlf
		sqlStr = sqlStr + " and i.sellcash>0" + vbcrlf
		sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL" + vbcrlf
		
	    ''제휴 사용안함인거 걸러냄. isExtusing = 'N'
	    ''sqlStr = sqlStr + " and i.isExtusing = 'Y'"
	    ''일단 수정은 되어야 함..
	    '''sqlStr = sqlStr + "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')"
        
        If FBrandID <> "" Then
    		sqlStr = sqlStr + " and i.makerid = '" & FBrandID & "' " + vbcrlf
    	End If
        
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		rsget.Close
	end sub
	
	public sub GetIParkDelSoldOutItemList()
	    dim sqlStr,i
	    
	    sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, " + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage, 
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, "
		sqlStr = sqlStr + " '0000' as itemoption," + vbcrlf
		sqlStr = sqlStr + " '' as optiontypename," + vbcrlf
		sqlStr = sqlStr + " '' as optionname," + vbcrlf
		sqlStr = sqlStr + " '' as optsellyn," + vbcrlf
		sqlStr = sqlStr + " '' as optlimityn," + vbcrlf
		sqlStr = sqlStr + " '' as optlimitno," + vbcrlf
		sqlStr = sqlStr + " '' as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " '' as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    
	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and i.sellyn ='N'"
		''sqlStr = sqlStr + " and i.isusing ='N'"
		''sqlStr = sqlStr + " and s.InterparkPrdNo is Not NULL"
		sqlStr = sqlStr + " and s.interparkregdate is Not NULL "
	    sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL and p.interparkstorecategory is Not NULL "
		sqlStr = sqlStr + " and datediff(m,s.interparkregdate,getdate())>3" ''--등록된지  4개월이상
		sqlStr = sqlStr + " order by s.interparklastupdate" ''i.itemid "
		
		
		If FTemp = "o" Then
			    sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
				sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, " + vbcrlf
				sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
				sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage, 
				sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
				sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
				sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
				sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
				sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, "
				sqlStr = sqlStr + " '0000' as itemoption," + vbcrlf
				sqlStr = sqlStr + " '' as optiontypename," + vbcrlf
				sqlStr = sqlStr + " '' as optionname," + vbcrlf
				sqlStr = sqlStr + " '' as optsellyn," + vbcrlf
				sqlStr = sqlStr + " '' as optlimityn," + vbcrlf
				sqlStr = sqlStr + " '' as optlimitno," + vbcrlf
				sqlStr = sqlStr + " '' as optlimitsold," + vbcrlf
				sqlStr = sqlStr + " '' as optaddprice" + vbcrlf
				sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
		        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
		        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
		        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf
				sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
				sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
				sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
				
		        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
			    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
			    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
			    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
			    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf
			    sqlStr = sqlStr + " where s.itemid=i.itemid and i.sellcash<>0 and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN
			    sqlStr = sqlStr + " and i.sellyn='Y'"   ''일단 판매중인내역만.
			    sqlStr = sqlStr + " and i.itemid not in (320687,266740,135625,250576,207883,178036,173781,170624)"
			    sqlStr = sqlStr + " and s.interparkregdate is Not NULL "
			    sqlStr = sqlStr + " order by s.regdate desc "
		End If
		
		
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				
				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))
				
				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if
				
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")
				
				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"
				    
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
				
				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				
				if IsNULL(rsget("regedinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("regedinterparkdispcategory")
				end if
				
				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if
				
				if IsNULL(rsget("regedInterparkstorecategory")) then 
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if
			
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
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")
				
				FItemList(i).Fisusing       = rsget("isusing")
				
				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")
				
				FItemList(i).FdeliveryType  = rsget("deliveryType")
				
				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub
    
	public sub GetIParkDelSoldOutItemList_PreVer()
	    dim sqlStr,i
	    
	    sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, " + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage, 
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, "
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    
	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and i.sellyn ='N'"
		''sqlStr = sqlStr + " and i.isusing ='N'"
		''sqlStr = sqlStr + " and s.InterparkPrdNo is Not NULL"
		sqlStr = sqlStr + " and s.interparkregdate is Not NULL "
	    sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL and p.interparkstorecategory is Not NULL "
		
		sqlStr = sqlStr + " order by i.itemid , o.itemoption"
		
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				
				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))
				
				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if
				
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")
				
				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"
				    
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
				
				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				
				if IsNULL(rsget("regedinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("regedinterparkdispcategory")
				end if
				
				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if
				
				if IsNULL(rsget("regedInterparkstorecategory")) then 
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if
			
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
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")
				
				FItemList(i).Fisusing       = rsget("isusing")
				
				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")
				
				FItemList(i).FdeliveryType  = rsget("deliveryType")
				
				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub
    
    

	public sub GetIParkDelJaeHyuItemList()
	    dim sqlStr, sqlSub, i
	    
	    sqlStr = "select top 30 i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, " + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage, 
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, "
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " 	and s.InterparkPrdNo is Not NULL"
		sqlStr = sqlStr + " 	and i.sellyn ='Y'"
		sqlStr = sqlStr + " 	and i.isExtusing ='N'"
		sqlStr = sqlStr + " order by i.itemid"
		
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				
				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))
				
				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if
				
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")
				
				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"
				    
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
				
				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				
				if IsNULL(rsget("regedinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("regedinterparkdispcategory")
				end if
				
				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if
				
				if IsNULL(rsget("regedInterparkstorecategory")) then 
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if
			
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
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")
				
				FItemList(i).Fisusing       = rsget("isusing")
				
				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")
				
				FItemList(i).FdeliveryType  = rsget("deliveryType")
				
				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub

	public sub GetIParkOneItemList(byval iitemid)
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
        
		sqlStr = "select  i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, " + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage, 
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory as regedinterparkDispCategory, "
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    
	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and s.itemid =" & iitemid 
		sqlStr = sqlStr + " order by i.itemid , o.itemoption"
		
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				
				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))
				
				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if
				
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")
				
				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"
				    
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
				
				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				
				if IsNULL(rsget("regedinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("regedinterparkdispcategory")
				end if
				
				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if
				
				if IsNULL(rsget("regedInterparkstorecategory")) then 
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if
			
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
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")
				
				FItemList(i).Fisusing       = rsget("isusing")
				
				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")
				
				FItemList(i).FdeliveryType  = rsget("deliveryType")
				
				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub
	
	
	public sub GetIParkEditItemList()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
        
		sqlStr = "select  i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, " + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate,  i.sailyn, i.orgprice," ''i.infoimage, 
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr + " s.PinterparkDispCategory,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    
	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and s.interparkPrdNo is Not NULL"
		sqlStr = sqlStr + " and s.itemid in ("
		sqlStr = sqlStr + "     select top " + CStr(FPageSize*FCurrPage) + " s.itemid from"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_interpark_reg_item s,"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_interpark_dspcategory_mapping p"
		sqlStr = sqlStr + "     where s.itemid=i.itemid"
		sqlStr = sqlStr + "     and s.interparkPrdNo is Not NULL"
		sqlStr = sqlStr + "     and s.interparklastupdate<i.lastupdate"
		'sqlStr = sqlStr + "     and ((i.lastupdate<'2010-04-12 00:00:00') or (i.lastupdate>'2010-04-12 03:00:00'))"
		sqlStr = sqlStr + "     and ((i.lastupdate<>'2010-04-22 00:09:48.693'))" '' or ((i.lastupdate='2008-10-21 00:06:16.140') and (i.sellyn<>'Y')) )"
		sqlStr = sqlStr + "     and i.basicimage is not null"
		sqlStr = sqlStr + "     and i.itemdiv<50"
		sqlStr = sqlStr + "     and i.cate_large<>''"
		sqlStr = sqlStr + "     and i.cate_large<>'999'"
		sqlStr = sqlStr + "     and i.sellcash>0"
		
	    ''제휴 사용안함인거 걸러냄. isExtusing = 'N'
	    ''sqlStr = sqlStr + " 	and i.isExtusing = 'Y'"
	    ''일단수정은되어야함.
	    '''sqlStr = sqlStr + "		and (i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark') or (i.sellyn<>'Y'))"
		
        If FBrandID <> "" Then
    		sqlStr = sqlStr + " and i.makerid = '" & FBrandID & "' " + vbcrlf
    	End If
		
		''역마진상품은 수정 안함 / 판매중인 아닌것 수정.
        sqlStr = sqlStr + "     and (((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN&" or (i.sellyn<>'Y'))" + VbCrlf
        ''특정상품제외;;
        sqlStr = sqlStr + "     and i.itemid<>171124"+ VbCrlf
        sqlStr = sqlStr + "     and i.itemid<>171659"+ VbCrlf
        sqlStr = sqlStr + "     and i.itemid<>171658"+ VbCrlf
        sqlStr = sqlStr + "     and i.itemid<>172515"+ VbCrlf
        sqlStr = sqlStr + "     and i.itemid<>172794"+ VbCrlf
        
		sqlStr = sqlStr + "     and i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호
		sqlStr = sqlStr + "     order by s.interparkregdate desc "
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by i.itemid , o.itemoption"
		
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				
				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))
				
				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if
				
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")
				
				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"
				    
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
				
				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				
				if IsNULL(rsget("PinterparkDispCategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("Pinterparkdispcategory")
				end if
				
				if IsNULL(rsget("interParkSupplyCtrtSeq")) then
				    FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
				else
				    FItemList(i).FSupplyCtrtSeq             = rsget("interParkSupplyCtrtSeq")
				end if
				
				if IsNULL(rsget("regedInterparkstorecategory")) then 
				    FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				else
				    FItemList(i).Finterparkstorecategory   = rsget("regedInterparkstorecategory")
			    end if
			    
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
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")
				
				FItemList(i).Fisusing       = rsget("isusing")
				
				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")
				
				FItemList(i).FdeliveryType  = rsget("deliveryType")
				
				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub
	
	public sub GetIParkEditItemList_OLD()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, " ''i.infoimage, 
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		
		sqlStr = sqlStr + " and s.interparklastupdate<i.lastupdate"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (i.lastupdate>'2008-04-17 12:00:00'))"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		'sqlStr = sqlStr + " and ((i.isusing='Y' and i.sellyn='Y') or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		sqlStr = sqlStr + " and i.basicimage is not null"
		sqlStr = sqlStr + " and i.itemdiv<50"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<>'999'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL"
		
		sqlStr = sqlStr + " order by i.itemid , o.itemoption"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				
				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if
				
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")
				
				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"
				    
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
				
				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				
				FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
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
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")
				
				FItemList(i).Fisusing       = rsget("isusing")
				
				FItemList(i).FInterparkPrdNo       = rsget("InterparkPrdNo")
				
				FItemList(i).FdeliveryType  = rsget("deliveryType")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

    public sub GetIParkRegItemTotalPage()
		dim sqlStr,i
		sqlStr = "select count(s.itemid) as cnt from "
		sqlStr = sqlStr + " [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호
	    
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid" + vbcrlf 
		sqlStr = sqlStr + " and s.interparkregdate is NULL"     
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (i.lastupdate>'2008-04-17 12:00:00'))"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		'sqlStr = sqlStr + " and ((i.isusing='Y' and i.sellyn='Y') or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		sqlStr = sqlStr + " and i.basicimage is not null" + vbcrlf
		sqlStr = sqlStr + " and i.itemdiv<50" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>''" + vbcrlf
		sqlStr = sqlStr + " and i.cate_large<>'999'" + vbcrlf
		sqlStr = sqlStr + " and i.sellcash>0" + vbcrlf
		sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL" + vbcrlf
		
	    ''제휴 사용안함인거 걸러냄. isExtusing = 'N'
	    sqlStr = sqlStr + " and i.isExtusing = 'Y'"
	    
	    sqlStr = sqlStr + "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')"
        
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		rsget.Close
	end sub
	
	public sub GetIParkRegItemList()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
		sqlStr = "select i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, " + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice," ''i.infoimage, 
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, s.Pinterparkdispcategory, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    
	    sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid " + vbcrlf
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and s.itemid in ("
		sqlStr = sqlStr + "     select top " + CStr(FPageSize*FCurrPage) + " s.itemid from"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_interpark_reg_item s,"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + "     [db_item].[dbo].tbl_interpark_dspcategory_mapping p"
		sqlStr = sqlStr + "     where s.itemid=i.itemid"
		sqlStr = sqlStr + "     and s.interparkregdate is NULL"
		sqlStr = sqlStr + "     and i.basicimage is not null"
		sqlStr = sqlStr + "     and i.itemdiv<50"
		sqlStr = sqlStr + "     and i.cate_large<>''"
		sqlStr = sqlStr + "     and i.cate_large<>'999'"
		sqlStr = sqlStr + "     and i.sellcash>0"
		
	    ''제휴 사용안함인거 걸러냄. isExtusing = 'N'
	    sqlStr = sqlStr + " 	and i.isExtusing = 'Y'"
	    
	    sqlStr = sqlStr + "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')"
		
		sqlStr = sqlStr + "     and i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.interparkdispcategory is Not NULL" + vbcrlf  '' 전시코드
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호
		sqlStr = sqlStr + "     order by s.itemid"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by i.itemid , o.itemoption"
		

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				
				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))
				
				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if
				
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")
				
				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"
				    
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
				
				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				
				if IsNULL(rsget("Pinterparkdispcategory")) then
				    FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				else
				    FItemList(i).Finterparkdispcategory    = rsget("Pinterparkdispcategory")
				end if
				FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
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
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")
				
				FItemList(i).Fisusing       = rsget("isusing")
				FItemList(i).FSailYn    = rsget("sailyn")
                FItemList(i).FOrgPrice  = rsget("orgprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub
	
	public sub GetIParkRegItemList_OLD()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr + " i.basicimage, i.mainimage, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, " ''i.infoimage, 
		sqlStr = sqlStr + " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr + " isNull(o.itemoption,'0000') as itemoption," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optiontypename,'') as optiontypename," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optionname,'') as optionname," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optsellyn,'') as optsellyn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimityn,'') as optlimityn," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitno,'') as optlimitno," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optlimitsold,'') as optlimitsold," + vbcrlf
		sqlStr = sqlStr + " isNull(o.optaddprice,0) as optaddprice" + vbcrlf
		sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3" + vbcrlf
        sqlStr = sqlStr + " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4" + vbcrlf

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item s," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'" + vbcrlf
		
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_interpark_dspcategory_mapping p " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=p.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=p.tencdm " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_small=p.tencdn " + vbcrlf
	    sqlStr = sqlStr + "     and p.SupplyCtrtSeq is Not NULL" + vbcrlf  '' 공급계약번호
	    
		sqlStr = sqlStr + " where s.itemid=i.itemid"
		sqlStr = sqlStr + " and s.interparkregdate is NULL"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (i.lastupdate>'2008-04-17 12:00:00'))"
		'sqlStr = sqlStr + " and ((dateDiff(d,s.regdate,getdate())<3) or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		'sqlStr = sqlStr + " and ((i.isusing='Y' and i.sellyn='Y') or (dateDiff(d,i.lastupdate,getdate())<7))" + vbcrlf
		sqlStr = sqlStr + " and i.basicimage is not null"
		sqlStr = sqlStr + " and i.itemdiv<50"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<>'999'"
		sqlStr = sqlStr + " and i.sellcash>0"
		sqlStr = sqlStr + " and p.interparkdispcategory is Not NULL"
		
		sqlStr = sqlStr + " order by i.itemid , o.itemoption"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterParkOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).FRegdate     = rsget("regdate")
				''FItemList(i).FUpdate  = rsget("lastupdate")

				FItemList(i).Fsellyn  = rsget("sellyn")
				''FItemList(i).Fdispyn  = rsget("dispyn")

				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))
				
				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if
				
				FItemList(i).Fkeywords = db2html(rsget("keywords"))

				FItemList(i).Fitemoption  = rsget("itemoption")
				FItemList(i).FItemOptionTypeName = db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName  = rsget("optionname")
				'FItemList(i).FItemOptionName  = replace(FItemList(i).FItemOptionName,"→","-")

				FItemList(i).Fbasicimage  = rsget("basicimage")
				FItemList(i).Fmainimage   = rsget("mainimage")
				'FItemList(i).FInfoImage   = rsget("infoimage")
				
				if IsNULL(FItemList(i).FInfoImage) then FItemList(i).FInfoImage=",,,,"
				    
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
				
				FItemList(i).Fvatinclude  = rsget("vatinclude")
				FItemList(i).Fkeywords  = db2html(rsget("keywords"))
				
				if (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent = replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				end if
				
				FItemList(i).Finterparkdispcategory    = rsget("interparkdispcategory")
				FItemList(i).Finterparkstorecategory   = rsget("interparkstorecategory")
				FItemList(i).FSupplyCtrtSeq             = rsget("SupplyCtrtSeq")
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
				FItemList(i).FInfoImage2  = rsget("InfoImage2")
				FItemList(i).FInfoImage3  = rsget("InfoImage3")
				FItemList(i).FInfoImage4  = rsget("InfoImage4")
				
				FItemList(i).Fisusing       = rsget("isusing")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub
	
	
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

End Class
%>