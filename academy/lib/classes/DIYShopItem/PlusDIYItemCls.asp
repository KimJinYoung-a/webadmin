<%
'####################################################
' Description :  다이샾 플러스 상품 클래스
' History : 2008.10.31 서동석 생성 
'			2010.11.09 한용민 수정
'####################################################

Class CPlusSaleMainItem    
    public FPlusSaleLinkItemID
    public FItemName
    public Fmakerid
    public Fsellcash
    public FBuycash
    public Fmwdiv 
    public FOrgPrice  
    public FOrgSuplycash     
    public FSellyn    
    public FLimitYn   
    public FLimitNo   
    public FLimitSold 
    public FsaleYn	
    public FisUsing    
    public FSmallImage
    public FPlusSaleItemCount   
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
    
    public function IsCurrentSaleItem()
        IsCurrentSaleItem = (FsaleYn="Y") and (FOrgPrice>Fsellcash)
    end function
    
    '// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = Fsellcash - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = Fsellcash
		end if
	end Function
	
	'// 쿠폰 할인가 
	public Function GetCouponDiscountPrice() '?
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*Fsellcash/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function
    
    '// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function
	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class

Class CPlusSaleSubItem
    public FPlusSaleItemID
    public FPlusSalePro              ''플러스세일 할인율    
    public FPlusSaleMargin           ''할인시 마진설정
    public FPlusSaleMaginFlag        ''할인시 마진설정 구분       
    public FPlusSaleStartDate
    public FPlusSaleEndDate    
    public Fregdate    
    public FPlusSaleLinkItemID
    public FLinkedItemCount    
    public FItemName
    public Fmakerid
    public Fsellcash
    public FBuycash
    public Fmwdiv    
    public FOrgPrice  
    public FOrgSuplycash    
    public FSellyn    
    public FLimitYn   
    public FLimitNo   
    public FLimitSold 
    public FsaleYn	    
    public FImageSmall    
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
    public fPlusdiyItemCount
    public Fitemid
    public Fcate_large
    public Fcate_mid
    public Fcate_small 
	public Fitemdiv
	public Fitemgubun
	public Fsailprice
	public fsailsuplycash
	public Fmileage
	public Flastupdate
	public FisUsing
	public Fdeliverytype
	public Fevalcnt
	public Foptioncnt
	public Fupchemanagecode
	public Fbrandname
	public FSmallImage
	public Flistimage
	public Flistimage120
	public Fcouponbuyprice
	public FinfoimageExists
	public FdefaultFreeBeasongLimit
	public FdefaultDeliverPay
	public FdefaultDeliveryType

    public Function IsSoldOut()
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function

    public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function		

    public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" or Fdeliverytype="7" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function	
	
    '// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = Fsellcash - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = Fsellcash
		end if
	end Function
	
	'// 쿠폰 할인가 
	public Function GetCouponDiscountPrice() '?
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*Fsellcash/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function
    
    '// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function
	
    ''플러스세일 판매가 = 현재판매가*할인율
    public function getPlusSalePrice()
        getPlusSalePrice = CLng(Fsellcash-Fsellcash*FPlusSalePro/100)
    end function
    
    ''플러스세일 매입가 = 
    public function getPlusSaleBuycash()
        
        if (FPlusSaleMaginFlag="4") then
            ''텐바이텐 부담
            getPlusSaleBuycash = FBuycash
            Exit function
        elseif (FPlusSaleMaginFlag="2") then
            ''업체 부담.
            getPlusSaleBuycash = CLng(getPlusSalePrice - (FSellCash-FBuyCash))
        else
            if (FPlusSaleMargin>0) and (FPlusSaleMargin<99) then
                getPlusSaleBuycash = CLng(getPlusSalePrice - getPlusSalePrice*FPlusSaleMargin/100)
            else
                getPlusSaleBuycash = FBuycash
            end if
        end if
    end function

    public function IsCurrentSaleItem()
        IsCurrentSaleItem = (FsaleYn="Y") and (FOrgPrice>Fsellcash)
    end function

    public function getMaginFlagName()
        getMaginFlagName = ""
        
        if FPlusSaleMaginFlag="1" then
	        getMaginFlagName = "동일마진"
	    elseif FPlusSaleMaginFlag="2" then
	        getMaginFlagName = "업체부담"
	    elseif FPlusSaleMaginFlag="3" then
	        getMaginFlagName = "반반부담"
	    elseif FPlusSaleMaginFlag="4" then
	        getMaginFlagName = "텐바이텐부담"
	    elseif FPlusSaleMaginFlag="5" then
	        getMaginFlagName = "직접설정"
	    end if
    end function

    public function getCurrstateName()
        if (now()<FPlusSaleStartDate) then
            getCurrstateName = "진행예정"
        elseif (now()<FPlusSaleEndDate) then
            getCurrstateName = "진행중"
        else
            getCurrstateName = "기간종료"
        end if
    end function
    
    public function getCurrstateColor()
        if (now()<FPlusSaleStartDate) then
            getCurrstateColor = "magenta"
        elseif (now()<FPlusSaleEndDate) then
            getCurrstateColor = "blue"
        else
            getCurrstateColor = "red"
        end if
    end function
    
    ''상시진행여부
    public function IsAlwaysTerms()
        IsAlwaysTerms = (FPlusSaleStartDate="1901-01-01") and (Left(FPlusSaleEndDate,10)="9999-12-31")
    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CPlusSaleItem
    public FItemList()
    public FOneItem   
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount	
	public FRectItemID
	public FRectMakerid
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectItemIDArr
	public FRectItemName
	public FRectOpenState	
	public FRectPlusSaleItemID
	public FRectMwDiv
	public FRectSellyn	
	public FRectPlusSaleLinkItemid
	public FRectisUsing
	public FRectLimityn
	public FRectDeliveryType
	public FRectsaleyn
	public FRectCouponYn
	public FRectCate_Large 
	public FRectCate_Mid
	public FRectCate_Small
	public FRectVatYn
	public FRectSortDiv
	
	public function IsPlusSaleLinkItem()
	    dim sqlStr,i
	    
	    sqlStr = "select count(*) as cnt from db_academy.dbo.tbl_diy_PlusSaleLinkItem"
	    sqlStr = sqlStr + " where PlusSaleLinkItemID=" + CStr(FRectItemID)
	    
	    rsACADEMYget.Open sqlStr,dbACADEMYget,1
	        IsPlusSaleLinkItem = rsACADEMYget("cnt")>0
	    rsACADEMYget.Close
	    
    end function
    
    public function GetPlusSaleMainItemList()
        dim sqlStr,i
        
        sqlStr = " select count(k.PlusSaleLinkItemID) as cnt "
        sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_PlusSaleLinkItem k"
        sqlStr = sqlStr + " , db_academy.dbo.tbl_diy_item i"
        sqlStr = sqlStr + " where k.plusSaleLinkItemID=i.itemid"
        if (FRectPlusSaleItemID<>"") then
            sqlStr = sqlStr + " and k.plusSaleItemid=" & FRectPlusSaleItemID
        end if
        
        if (FRectItemIDArr<>"") then
            sqlStr = sqlStr + " and i.itemid in (" + FRectItemIDArr + ")"
        end if 
        
        if (FRectMakerid<>"") then
            sqlStr = sqlStr + " and i.makerid='" & FRectMakerid & "'"
        end if
        
        if (FRectCDL<>"") then
            sqlStr = sqlStr + " and i.cate_large='" + FRectCDL + "'"
        end if
        
        if (FRectCDM<>"") then
            sqlStr = sqlStr + " and i.cate_mid='" + FRectCDM + "'"
        end if
        
        if (FRectCDS<>"") then
            sqlStr = sqlStr + " and i.cate_small='" + FRectCDS + "'"
        end if
        
        if (FRectItemName<>"") then
            sqlStr = sqlStr + " and i.itemname like '%" + FRectItemName + "%'"
        end if
        
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
        sqlStr = "select distinct top " + CStr(FPageSize*FCurrPage) + " k.PlusSaleLinkItemID, "
        sqlStr = sqlStr + "	i.makerid, i.itemname, i.sellcash, i.buycash, i.orgprice,  i.orgsuplycash,"
        sqlStr = sqlStr + "	i.sellyn, i.limityn, i.saleYn, i.isusing, i.mwdiv, i.smallimage,"
        sqlStr = sqlStr + "	i.itemcouponyn, i.curritemcouponidx, i.itemcoupontype, i.itemcouponvalue,"
        sqlStr = sqlStr + "	(select count(T.PlusSaleItemID) from db_academy.dbo.tbl_diy_PlusSaleLinkItem T where T.PlusSaleLinkItemID=k.plusSaleLinkItemID) as PlusSaleItemCount"
        sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_PlusSaleLinkItem k"
        sqlStr = sqlStr + " , db_academy.dbo.tbl_diy_item i"
        sqlStr = sqlStr + " where k.plusSaleLinkItemID=i.itemid"
        
        if (FRectPlusSaleItemID<>"") then
            sqlStr = sqlStr + " and k.plusSaleItemid=" & FRectPlusSaleItemID
        end if
        
        if (FRectItemIDArr<>"") then
            sqlStr = sqlStr + " and i.itemid in (" + FRectItemIDArr + ")"
        end if 
        
        if (FRectMakerid<>"") then
            sqlStr = sqlStr + " and i.makerid='" & FRectMakerid & "'"
        end if
        
        if (FRectCDL<>"") then
            sqlStr = sqlStr + " and i.cate_large='" + FRectCDL + "'"
        end if
        
        if (FRectCDM<>"") then
            sqlStr = sqlStr + " and i.cate_mid='" + FRectCDM + "'"
        end if
        
        if (FRectCDS<>"") then
            sqlStr = sqlStr + " and i.cate_small='" + FRectCDS + "'"
        end if
        
        if (FRectItemName<>"") then
            sqlStr = sqlStr + " and i.itemname like '%" + FRectItemName + "%'"
        end if
        sqlStr = sqlStr + " order by k.PlusSaleLinkItemID desc"
        
        'response.write sqlStr &"<br>"
        rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CPlusSaleMainItem
				
				FItemList(i).FPlusSaleLinkItemID= rsACADEMYget("PlusSaleLinkItemID")
                FItemList(i).Fmakerid           = rsACADEMYget("makerid")
                FItemList(i).Fitemname          = db2html(rsACADEMYget("itemname"))
                FItemList(i).Fsellcash          = rsACADEMYget("sellcash")
                FItemList(i).Fbuycash           = rsACADEMYget("buycash")
                FItemList(i).Forgprice          = rsACADEMYget("orgprice")
                FItemList(i).Forgsuplycash      = rsACADEMYget("orgsuplycash")
                FItemList(i).Fsellyn            = rsACADEMYget("sellyn")
                FItemList(i).Flimityn           = rsACADEMYget("limityn")
                FItemList(i).FsaleYn            = rsACADEMYget("saleYn")
                FItemList(i).Fisusing           = rsACADEMYget("isusing")
                FItemList(i).Fmwdiv             = rsACADEMYget("mwdiv")               
                FItemList(i).FSmallImage        = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(FItemList(i).FPlusSaleLinkItemID) + "/" + rsACADEMYget("smallimage")
                FItemList(i).Fitemcouponyn      = rsACADEMYget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsACADEMYget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsACADEMYget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsACADEMYget("itemcouponvalue")
                FItemList(i).FPlusSaleItemCount   = rsACADEMYget("PlusSaleItemCount")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.Close
    end function
    
    '//academy/itemmaster/PlusDIYItem/PlusDIYItem_list.asp
    public function GetPlusSaleSubItemListByMainLinkItemID()
        dim sqlStr,i
        
        sqlStr = " select count(s.plusSaleItemID) as Cnt"
        sqlStr = sqlStr + "	from db_academy.dbo.tbl_diy_PlusSaleLinkItem l"
        sqlStr = sqlStr + "		,db_academy.dbo.tbl_diy_PlusSaleRegedItem s"
        sqlStr = sqlStr + "		,db_academy.dbo.tbl_diy_item i"
        sqlStr = sqlStr + "	where s.plusSaleItemID=l.plusSaleItemID"
        sqlStr = sqlStr + "	and s.plusSaleItemID=i.itemid"
        sqlStr = sqlStr + "	and l.plusSaleLinkItemID=" + CStr(FRectItemID)
        
        'response.write sqlStr &"<br>"
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
        sqlStr = " select top " + CStr(FPageSize*FCurrPage) 
        sqlStr = sqlStr + " s.*, i.makerid, i.itemName, i.mwdiv, i.sellcash, i.buycash, i.orgprice, i.OrgSuplycash"
        sqlStr = sqlStr + " , i.sellyn, i.limityn, i.limitno, i.limitsold, i.saleYn, i.smallimage "
        sqlStr = sqlStr + "	from db_academy.dbo.tbl_diy_PlusSaleLinkItem l"
        sqlStr = sqlStr + "		,db_academy.dbo.tbl_diy_PlusSaleRegedItem s"
        sqlStr = sqlStr + "		,db_academy.dbo.tbl_diy_item i"
        sqlStr = sqlStr + "	where s.plusSaleItemID=l.plusSaleItemID"
        sqlStr = sqlStr + "	and s.plusSaleItemID=i.itemid"
        sqlStr = sqlStr + "	and l.plusSaleLinkItemID=" + CStr(FRectItemID)    
        sqlStr = sqlStr + " order by s.regdate desc"
		
		'response.write sqlStr &"<br>"
        rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CPlusSaleSubItem
				
				FItemList(i).FPlusSaleItemID= rsACADEMYget("PlusSaleItemID")
				FItemList(i).FPlusSalePro   = rsACADEMYget("PlusSalePro")
    			FItemList(i).FPlusSaleMargin     = rsACADEMYget("PlusSaleMargin")
                FItemList(i).FPlusSaleMaginFlag  = rsACADEMYget("PlusSaleMaginFlag")
                FItemList(i).FPlusSaleStartDate  = rsACADEMYget("PlusSaleStartDate")
                FItemList(i).FPlusSaleEndDate    = rsACADEMYget("PlusSaleEndDate")
                FItemList(i).Fregdate           = rsACADEMYget("regdate")          
				FItemList(i).FItemName     = db2html(rsACADEMYget("itemname"))
				FItemList(i).Fmakerid      = rsACADEMYget("makerid")
				FItemList(i).FSellCash     = rsACADEMYget("sellcash")
				FItemList(i).FBuyCash     = rsACADEMYget("buycash")
				FItemList(i).Fmwdiv        = rsACADEMYget("mwdiv")
				FItemList(i).FOrgPrice     = rsACADEMYget("orgprice")
				FItemList(i).FOrgSuplycash = rsACADEMYget("OrgSuplycash")
				FItemList(i).FSellyn       = rsACADEMYget("sellyn")
				FItemList(i).FLimitYn      = rsACADEMYget("limityn")
				FItemList(i).FLimitNo      = rsACADEMYget("limitno")
				FItemList(i).FLimitSold    = rsACADEMYget("limitsold")
				FItemList(i).FsaleYn		  = rsACADEMYget("saleYn")            
				FItemList(i).FImageSmall   = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(FItemList(i).FPlusSaleItemID) + "/" + rsACADEMYget("smallimage")
				
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
    end function

	public function GetOnePlusSaleSubItem()
	    dim sqlStr,i
	    
	    sqlStr = "select top 1"
	    sqlStr = sqlStr + "	S.*,  i.makerid, i.itemName, i.mwdiv, i.sellcash, i.buycash, i.orgprice, i.OrgSuplycash, i.sellyn, i.limityn, i.limitno, i.limitsold, i.saleYn, i.smallimage, "
	    sqlStr = sqlStr + "	(select count(T.PlusSaleItemID) from db_academy.dbo.tbl_diy_PlusSaleLinkItem T where T.PlusSaleItemID=S.PlusSaleItemID) as LinkedItemCount"
        sqlStr = sqlStr + "	from  db_academy.dbo.tbl_diy_PlusSaleRegedItem S"
	    sqlStr = sqlStr + "	    Join db_academy.dbo.tbl_diy_item i"
	    sqlStr = sqlStr + "	    on S.PlusSaleItemID=i.itemid"
	    sqlStr = sqlStr + " where S.PlusSaleItemID=" + CStr(FRectItemID)
	    
	    rsACADEMYget.Open sqlStr,dbACADEMYget,1
	    FResultCount = rsACADEMYget.RecordCount
	    if Not rsACADEMYget.Eof then
	        set FOneItem = new CPlusSaleSubItem
	        
			FOneItem.FPlusSaleItemID= rsACADEMYget("PlusSaleItemID")
			FOneItem.FPlusSalePro   = rsACADEMYget("PlusSalePro")
			FOneItem.FPlusSaleMargin     = rsACADEMYget("PlusSaleMargin")
            FOneItem.FPlusSaleMaginFlag  = rsACADEMYget("PlusSaleMaginFlag")
            FOneItem.FPlusSaleStartDate  = rsACADEMYget("PlusSaleStartDate")
            FOneItem.FPlusSaleEndDate    = rsACADEMYget("PlusSaleEndDate")
            FOneItem.Fregdate           = rsACADEMYget("regdate")
			FOneItem.FItemName     = db2html(rsACADEMYget("itemname"))
			FOneItem.Fmakerid      = rsACADEMYget("makerid")
			FOneItem.FSellCash     = rsACADEMYget("sellcash")
			FOneItem.FBuycash      = rsACADEMYget("buycash")
			FOneItem.Fmwdiv        = rsACADEMYget("mwdiv")
			FOneItem.FOrgPrice     = rsACADEMYget("orgprice")
			FOneItem.FOrgSuplycash = rsACADEMYget("OrgSuplycash")
			FOneItem.FSellyn       = rsACADEMYget("sellyn")
			FOneItem.FLimitYn      = rsACADEMYget("limityn")
			FOneItem.FLimitNo      = rsACADEMYget("limitno")
			FOneItem.FLimitSold    = rsACADEMYget("limitsold")
			FOneItem.FsaleYn		  = rsACADEMYget("saleYn")
            FOneItem.FLinkedItemCount = rsACADEMYget("LinkedItemCount")            
			FOneItem.FImageSmall   = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(FOneItem.FPlusSaleItemID) + "/" + rsACADEMYget("smallimage")

	    end if
	    rsACADEMYget.Close
    end function
	
	'//academy/itemmaster/plusdiyitem/pop_plusdiyitem_list.asp
	public function getplusdiyitem_list()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectplusSaleLinkItemID <> "") then
            addSql = addSql & " and i.itemid <> '" + FRectplusSaleLinkItemID + "'"
        end if
        
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if
        
        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
		
		if FRectLimityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            addSql = addSql + " and i.limityn='" + FRectLimityn + "'"
        end if        
        
        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if
        
        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if
        
        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if
        
        if FRectsaleyn<>"" then
            addSql = addSql + " and i.saleyn='" + FRectsaleyn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatYn='" + FRectVatYn + "'"
        end if
        
        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
		sqlStr = sqlStr & " left join db_academy.dbo.tbl_diy_PlusSaleRegedItem s on S.PlusSaleItemID=i.itemid"
        sqlStr = sqlStr & " left join ("
        sqlStr = sqlStr & "  	select plusSaleLinkItemID from db_academy.dbo.tbl_diy_PlusSaleLinkItem group by plusSaleLinkItemID"
        sqlStr = sqlStr & " ) as T"
        sqlStr = sqlStr & " on i.itemid = t.plusSaleLinkItemID"        
        sqlStr = sqlStr & " where s.plusSaleItemID is null" '추가 상품은 제낀다
        sqlStr = sqlStr & " and i.saleYn='N'" '할인 상품은 제낀다
        sqlStr = sqlStr & " and t.plusSaleLinkItemID is null" '플러스 등록 상품은 제낀다
		sqlStr = sqlStr & " and i.itemid<>0" & addSql   
		
		'response.write sqlStr &"<br>"
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliveryPay,0) as defaultDeliverPay, IsNULL(diy_dlv_gubun,'') as defaultDeliveryType"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From db_academy.dbo.tbl_diy_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "        
        sqlStr = sqlStr & " ,(select count(T.PlusSaleItemID) from db_academy.dbo.tbl_diy_PlusSaleLinkItem T where T.PlusSaleLinkItemID=i.itemid) as PlusdiyItemCount"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i "
        sqlStr = sqlStr & " left join db_academy.dbo.tbl_diy_item_addimage A on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & " left join db_academy.dbo.tbl_lec_user c on i.makerid=c.lecturer_id"
		sqlStr = sqlStr & " left join db_academy.dbo.tbl_diy_PlusSaleRegedItem s on S.PlusSaleItemID=i.itemid"
        sqlStr = sqlStr & " left join ("
        sqlStr = sqlStr & "  	select plusSaleLinkItemID from db_academy.dbo.tbl_diy_PlusSaleLinkItem group by plusSaleLinkItemID"
        sqlStr = sqlStr & " ) as T"
        sqlStr = sqlStr & " on i.itemid = t.plusSaleLinkItemID"		                
        sqlStr = sqlStr & " where s.plusSaleItemID is null" '추가 상품은 제낀다
        sqlStr = sqlStr & " and i.saleYn='N'" '할인 상품은 제낀다
        sqlStr = sqlStr & " and t.plusSaleLinkItemID is null" '플러스 등록 상품은 제낀다
		sqlStr = sqlStr & " and i.itemid<>0" & addSql        

		IF FRectSortDiv="new" Then
			sqlStr = sqlStr & " Order by i.itemid desc "
		ELSEIF FRectSortDiv="cashH" Then 
			sqlStr = sqlStr & " Order by i.SellCash desc "
		ELSEIF FRectSortDiv="cashL" Then
			sqlStr = sqlStr & " Order by i.SellCash"
		ELSEIF FRectSortDiv="best" Then
			sqlStr = sqlStr & " Order by i.ItemScore desc "
		ELSE
			sqlStr = sqlStr & " Order by i.itemid desc "
		End IF       

		'response.write sqlStr &"<br>"
        rsACADEMYget.pagesize = FPageSize
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
        
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		
        if (FResultCount<1) then FResultCount=0
        
        redim preserve FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
                set FItemList(i) = new CPlusSaleSubItem
                                
                FItemList(i).fPlusdiyItemCount  = rsACADEMYget("PlusdiyItemCount")
                FItemList(i).Fitemid            = rsACADEMYget("itemid")
                FItemList(i).Fmakerid           = rsACADEMYget("makerid")
                FItemList(i).Fcate_large        = rsACADEMYget("cate_large")
                FItemList(i).Fcate_mid          = rsACADEMYget("cate_mid")
                FItemList(i).Fcate_small        = rsACADEMYget("cate_small")
                FItemList(i).Fitemdiv           = rsACADEMYget("itemdiv")
                FItemList(i).Fitemgubun         = rsACADEMYget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsACADEMYget("itemname"))
                FItemList(i).Fsellcash          = rsACADEMYget("sellcash")
                FItemList(i).Fbuycash           = rsACADEMYget("buycash")
                FItemList(i).Forgprice          = rsACADEMYget("orgprice")
                FItemList(i).Forgsuplycash      = rsACADEMYget("orgsuplycash")
                FItemList(i).Fsailprice         = rsACADEMYget("sailprice")
                FItemList(i).Fsailsuplycash     = rsACADEMYget("sailsuplycash")
                FItemList(i).Fmileage           = rsACADEMYget("mileage")
                FItemList(i).Fregdate           = rsACADEMYget("regdate")
                FItemList(i).Flastupdate        = rsACADEMYget("lastupdate")
                FItemList(i).Fsellyn            = rsACADEMYget("sellyn")
                FItemList(i).Flimityn           = rsACADEMYget("limityn")
                FItemList(i).Fsaleyn            = rsACADEMYget("saleyn")
                FItemList(i).Fisusing           = rsACADEMYget("isusing")
                FItemList(i).Fmwdiv             = rsACADEMYget("mwdiv")
                FItemList(i).Fdeliverytype      = rsACADEMYget("deliverytype")
                FItemList(i).Flimitno           = rsACADEMYget("limitno")
                FItemList(i).Flimitsold         = rsACADEMYget("limitsold")
                FItemList(i).Fevalcnt           = rsACADEMYget("evalcnt")
                FItemList(i).Foptioncnt         = rsACADEMYget("optioncnt")
                FItemList(i).Fupchemanagecode   = rsACADEMYget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsACADEMYget("brandname"))
                FItemList(i).Fsmallimage        = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("smallimage")
                FItemList(i).Flistimage         = imgFingers & "/diyItem/webimage/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("listimage")
                FItemList(i).Flistimage120      = imgFingers & "/diyItem/webimage/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("listimage120")
                FItemList(i).Fitemcouponyn      = rsACADEMYget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsACADEMYget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsACADEMYget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsACADEMYget("itemcouponvalue")
                FItemList(i).Fcouponbuyprice    = rsACADEMYget("couponbuyprice")	'쿠폰적용 매입가
                
                if (rsACADEMYget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if
                
                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsACADEMYget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsACADEMYget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsACADEMYget("defaultDeliveryType")
                
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end function

	public function GetPlusSaleSubItemList()
	    dim sqlStr,i
	    
	    sqlStr = " select count(S.PlusSaleItemID) as cnt "
	    sqlStr = sqlStr + "	from db_academy.dbo.tbl_diy_PlusSaleRegedItem S"
	    sqlStr = sqlStr + "	    Join db_academy.dbo.tbl_diy_item i"
	    sqlStr = sqlStr + "	    on S.PlusSaleItemID=i.itemid"
	    sqlStr = sqlStr + "	where 1=1"
	    
	    if (FRectOpenState<>"") then
	        if (FRectOpenState="open") then
	            sqlStr = sqlStr + "	and s.plusSaleStartDate<getdate()"
	            sqlStr = sqlStr + "	and s.plusSaleEndDate>getdate()"
	        elseif (FRectOpenState="scheduled") then
	            sqlStr = sqlStr + "	and s.plusSaleStartDate>getdate()"
	        elseif (FRectOpenState="expired") then
	            sqlStr = sqlStr + "	and s.plusSaleEndDate<getdate()"
	        elseif (FRectOpenState="openscheduled") then
	            sqlStr = sqlStr + "	and s.plusSaleEndDate>getdate()"
	        end if
	    end if
	    
	    if (FRectMakerid<>"") then
	        sqlStr = sqlStr + "	and i.makerid='" + FRectMakerid + "'"
	    end if
	    
	    if (FRectItemIDArr<>"") then
	        sqlStr = sqlStr + "	and S.PlusSaleItemID in (" + FRectItemIDArr + ")"
	    end if
	    
	    if (FRectItemName<>"") then
	        sqlStr = sqlStr + "	and i.itemname like '%" + FRectItemName + "%'"
	    end if
	    
	    if (FRectCDL<>"") then
	        sqlStr = sqlStr + "	and i.cate_large='" + FRectCDL + "'"
	    end if
	    
	    if (FRectCDM<>"") then
	        sqlStr = sqlStr + "	and i.cate_mid='" + FRectCDM + "'"
	    end if
	    
	    if (FRectCDS<>"") then
	        sqlStr = sqlStr + "	and i.cate_small='" + FRectCDS + "'"
	    end if
	    
	    if (FRectSellyn<>"") then
	        sqlStr = sqlStr + "	and i.sellyn='" & FRectSellyn & "'"
	    end if
	    
	    if (FRectMwDiv<>"") then
	        if (FRectMwDiv="MW") then
	            sqlStr = sqlStr + "	and i.mwdiv<>'U'"
	        else
	            sqlStr = sqlStr + "	and i.mwdiv='" & FRectSellyn & "'"
	        end if
	    end if
	    
	    rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
	    sqlStr = "select top " + CStr(FPageSize*FCurrPage)
	    sqlStr = sqlStr + "	S.*,  i.makerid, i.itemName, i.sellcash, i.buycash, i.mwdiv, i.orgprice, i.OrgSuplycash, i.sellyn, i.limityn, i.limitno, i.limitsold, i.saleYn, i.smallimage, "
	    sqlStr = sqlStr + "	i.itemcouponyn, i.curritemcouponidx, i.itemcoupontype, i.itemcouponvalue,"
	    sqlStr = sqlStr + "	(select count(T.PlusSaleItemID) from db_academy.dbo.tbl_diy_PlusSaleLinkItem T where T.PlusSaleItemID=S.PlusSaleItemID) as LinkedItemCount"
        sqlStr = sqlStr + "	from  db_academy.dbo.tbl_diy_PlusSaleRegedItem S"
	    sqlStr = sqlStr + "	    Join db_academy.dbo.tbl_diy_item i"
	    sqlStr = sqlStr + "	    on S.PlusSaleItemID=i.itemid"
	    sqlStr = sqlStr + " where 1=1"
	    
	    if (FRectOpenState<>"") then
	        if (FRectOpenState="open") then
	            sqlStr = sqlStr + "	and s.plusSaleStartDate<getdate()"
	            sqlStr = sqlStr + "	and s.plusSaleEndDate>getdate()"
	        elseif (FRectOpenState="scheduled") then
	            sqlStr = sqlStr + "	and s.plusSaleStartDate>getdate()"
	        elseif (FRectOpenState="expired") then
	            sqlStr = sqlStr + "	and s.plusSaleEndDate<getdate()"
	        elseif (FRectOpenState="openscheduled") then
	            sqlStr = sqlStr + "	and s.plusSaleEndDate>getdate()"
	        end if
	    end if
	    
	    if (FRectMakerid<>"") then
	        sqlStr = sqlStr + "	and i.makerid='" + FRectMakerid + "'"
	    end if
	    
	    if (FRectItemIDArr<>"") then
	        sqlStr = sqlStr + "	and S.PlusSaleItemID in (" + FRectItemIDArr + ")"
	    end if
	    
	    if (FRectItemName<>"") then
	        sqlStr = sqlStr + "	and i.itemname like '%" + FRectItemName + "%'"
	    end if
	    
	    if (FRectCDL<>"") then
	        sqlStr = sqlStr + "	and i.cate_large='" + FRectCDL + "'"
	    end if
	    
	    if (FRectCDM<>"") then
	        sqlStr = sqlStr + "	and i.cate_mid='" + FRectCDM + "'"
	    end if
	    
	    if (FRectCDS<>"") then
	        sqlStr = sqlStr + "	and i.cate_small='" + FRectCDS + "'"
	    end if
	    
	    if (FRectSellyn<>"") then
	        sqlStr = sqlStr + "	and i.sellyn='" & FRectSellyn & "'"
	    end if
	    
	    if (FRectMwDiv<>"") then
	        if (FRectMwDiv="MW") then
	            sqlStr = sqlStr + "	and i.mwdiv<>'U'"
	        else
	            sqlStr = sqlStr + "	and i.mwdiv='" & FRectSellyn & "'"
	        end if
	    end if
	    
	    sqlStr = sqlStr + " order by S.regdate desc"
	    
	    'response.write sqlStr
	    rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CPlusSaleSubItem
				
				FItemList(i).FPlusSaleItemID= rsACADEMYget("PlusSaleItemID")
				FItemList(i).FPlusSalePro   = rsACADEMYget("PlusSalePro")
    			FItemList(i).FPlusSaleMargin     = rsACADEMYget("PlusSaleMargin")
                FItemList(i).FPlusSaleMaginFlag  = rsACADEMYget("PlusSaleMaginFlag")
                FItemList(i).FPlusSaleStartDate  = rsACADEMYget("PlusSaleStartDate")
                FItemList(i).FPlusSaleEndDate    = rsACADEMYget("PlusSaleEndDate")
                FItemList(i).Fregdate           = rsACADEMYget("regdate")            
				FItemList(i).FItemName     = db2html(rsACADEMYget("itemname"))
				FItemList(i).Fmakerid      = rsACADEMYget("makerid")
				FItemList(i).FSellCash     = rsACADEMYget("sellcash")
				FItemList(i).FBuyCash     = rsACADEMYget("buycash")
				FItemList(i).Fmwdiv        = rsACADEMYget("mwdiv")
				FItemList(i).FOrgPrice     = rsACADEMYget("orgprice")
				FItemList(i).FOrgSuplycash = rsACADEMYget("OrgSuplycash")
				FItemList(i).FSellyn       = rsACADEMYget("sellyn")
				FItemList(i).FLimitYn      = rsACADEMYget("limityn")
				FItemList(i).FLimitNo      = rsACADEMYget("limitno")
				FItemList(i).FLimitSold    = rsACADEMYget("limitsold")
				FItemList(i).FsaleYn		  = rsACADEMYget("saleYn")
                FItemList(i).FLinkedItemCount = rsACADEMYget("LinkedItemCount")                
				FItemList(i).FImageSmall   = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(FItemList(i).FPlusSaleItemID) + "/" + rsACADEMYget("smallimage")
                FItemList(i).Fitemcouponyn      = rsACADEMYget("itemcouponyn")
                FItemList(i).Fitemcoupontype    = rsACADEMYget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsACADEMYget("itemcouponvalue")
                FItemList(i).Fcurritemcouponidx = rsACADEMYget("curritemcouponidx")
            
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.Close		
    end function
    
    public function GetPlusSaleMajorItemList()
	    dim sqlStr,i	    
    end function
    
    Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
End Class
%>