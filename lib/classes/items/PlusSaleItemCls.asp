<%
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
    public Fdanjongyn
    public FSailYN	
    public FisUsing
    
    public FSmallImage
    public FPlusSaleItemCount
    
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
    
    public function IsCurrentSaleItem()
        IsCurrentSaleItem = (FSailYN="Y") and (FOrgPrice>Fsellcash)
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
    
    '''public FPlusSaleBuyCash       ''=> 계산값으로 나와야.. (가격 변동에 대처)
    public FPlusSaleStartDate
    public FPlusSaleEndDate
    ''public Fisusing                ''사용안함
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
    public FSailYN	
    
    public FImageSmall
    
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
    
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
        IsCurrentSaleItem = (FSailYN="Y") and (FOrgPrice>Fsellcash)
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
'        if (Fisusing="N") then
'            getCurrstateName = "사용안함"
'        else
            if (now()<FPlusSaleStartDate) then
                getCurrstateName = "진행예정"
            elseif (now()<FPlusSaleEndDate) then
                getCurrstateName = "진행중"
            else
                getCurrstateName = "기간종료"
            end if
'        end if
    end function
    
    public function getCurrstateColor()
'        if (Fisusing="N") then
'            getCurrstateColor = "gray"
'        else
            if (now()<FPlusSaleStartDate) then
                getCurrstateColor = "magenta"
            elseif (now()<FPlusSaleEndDate) then
                getCurrstateColor = "blue"
            else
                getCurrstateColor = "red"
            end if
'        end if
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
	
	public function IsPlusSaleLinkItem()
	    dim sqlStr,i
	    
	    sqlStr = "select count(*) as cnt from db_item.dbo.tbl_PlusSaleLinkItemList"
	    sqlStr = sqlStr + " where PlusSaleLinkItemID=" + CStr(FRectItemID)
	    rsget.Open sqlStr,dbget,1
	        IsPlusSaleLinkItem = rsget("cnt")>0
	    rsget.Close
    end function
    
    public function GetPlusSaleMainItemList()
        dim sqlStr,i
        sqlStr = " select count(k.PlusSaleLinkItemID) as cnt "
        sqlStr = sqlStr + " from db_item.dbo.tbl_PlusSaleLinkItemList k"
        sqlStr = sqlStr + " , db_item.dbo.tbl_item i"
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
        
        rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close
		
        sqlStr = "select distinct top " + CStr(FPageSize*FCurrPage) + " k.PlusSaleLinkItemID, "
        sqlStr = sqlStr + "	i.makerid, i.itemname, i.sellcash, i.buycash, i.orgprice,  i.orgsuplycash,"
        sqlStr = sqlStr + "	i.sellyn, i.limityn, i.danjongyn, i.sailyn, i.isusing, i.mwdiv, i.smallimage,"
        sqlStr = sqlStr + "	i.itemcouponyn, i.curritemcouponidx, i.itemcoupontype, i.itemcouponvalue,"
        sqlStr = sqlStr + "	(select count(T.PlusSaleItemID) from db_item.dbo.tbl_PlusSaleLinkItemList T where T.PlusSaleLinkItemID=k.plusSaleLinkItemID) as PlusSaleItemCount"
        sqlStr = sqlStr + " from db_item.dbo.tbl_PlusSaleLinkItemList k"
        sqlStr = sqlStr + " , db_item.dbo.tbl_item i"
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
        
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new CPlusSaleMainItem
				FItemList(i).FPlusSaleLinkItemID= rsget("PlusSaleLinkItemID")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                
                FItemList(i).FSmallImage        = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FPlusSaleLinkItemID) + "/" + rsget("smallimage")

                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
                
'                FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가       
         
                FItemList(i).FPlusSaleItemCount   = rsget("PlusSaleItemCount")


				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
		
    end function
    
    public function GetPlusSaleSubItemListByMainLinkItemID()
        dim sqlStr,i
        sqlStr = " select count(s.plusSaleItemID) as Cnt"
        sqlStr = sqlStr + "	from db_item.dbo.tbl_PlusSaleLinkItemList l"
        sqlStr = sqlStr + "		,db_item.dbo.tbl_PlusSaleRegedItem s"
        sqlStr = sqlStr + "		,db_item.dbo.tbl_item i"
        sqlStr = sqlStr + "	where s.plusSaleItemID=l.plusSaleItemID"
        sqlStr = sqlStr + "	and s.plusSaleItemID=i.itemid"
        sqlStr = sqlStr + "	and l.plusSaleLinkItemID=" + CStr(FRectItemID)
        
        rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close
		
        sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " s.*, i.makerid, i.itemName, i.mwdiv, i.sellcash, i.buycash, i.orgprice, i.OrgSuplycash, i.sellyn, i.limityn, i.limitno, i.limitsold, i.sailyn, i.smallimage "
        sqlStr = sqlStr + "	from db_item.dbo.tbl_PlusSaleLinkItemList l"
        sqlStr = sqlStr + "		,db_item.dbo.tbl_PlusSaleRegedItem s"
        sqlStr = sqlStr + "		,db_item.dbo.tbl_item i"
        sqlStr = sqlStr + "	where s.plusSaleItemID=l.plusSaleItemID"
        sqlStr = sqlStr + "	and s.plusSaleItemID=i.itemid"
        sqlStr = sqlStr + "	and l.plusSaleLinkItemID=" + CStr(FRectItemID)
        
        sqlStr = sqlStr + " order by s.regdate desc"
		
		'response.write sqlStr
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new CPlusSaleSubItem
				FItemList(i).FPlusSaleItemID= rsget("PlusSaleItemID")
				FItemList(i).FPlusSalePro   = rsget("PlusSalePro")
    			FItemList(i).FPlusSaleMargin     = rsget("PlusSaleMargin")
                FItemList(i).FPlusSaleMaginFlag  = rsget("PlusSaleMaginFlag")
                FItemList(i).FPlusSaleStartDate  = rsget("PlusSaleStartDate")
                FItemList(i).FPlusSaleEndDate    = rsget("PlusSaleEndDate")
                FItemList(i).Fregdate           = rsget("regdate")
            
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).Fmakerid      = rsget("makerid")
				FItemList(i).FSellCash     = rsget("sellcash")
				FItemList(i).FBuyCash     = rsget("buycash")
				FItemList(i).Fmwdiv        = rsget("mwdiv")
				FItemList(i).FOrgPrice     = rsget("orgprice")
				FItemList(i).FOrgSuplycash = rsget("OrgSuplycash")
				FItemList(i).FSellyn       = rsget("sellyn")
				FItemList(i).FLimitYn      = rsget("limityn")
				FItemList(i).FLimitNo      = rsget("limitno")
				FItemList(i).FLimitSold    = rsget("limitsold")

				FItemList(i).FSailYN		  = rsget("sailyn")
                
				FItemList(i).FImageSmall   = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FPlusSaleItemID) + "/" + rsget("smallimage")


				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end function

	public function GetOnePlusSaleSubItem()
	    dim sqlStr,i
	    sqlStr = "select top 1"
	    sqlStr = sqlStr + "	S.*,  i.makerid, i.itemName, i.mwdiv, i.sellcash, i.buycash, i.orgprice, i.OrgSuplycash, i.sellyn, i.limityn, i.limitno, i.limitsold, i.sailyn, i.smallimage, "
	    sqlStr = sqlStr + "	(select count(T.PlusSaleItemID) from db_item.dbo.tbl_PlusSaleLinkItemList T where T.PlusSaleItemID=S.PlusSaleItemID) as LinkedItemCount"
        sqlStr = sqlStr + "	from  db_item.dbo.tbl_PlusSaleRegedItem S"
	    sqlStr = sqlStr + "	    Join db_item.dbo.tbl_item i"
	    sqlStr = sqlStr + "	    on S.PlusSaleItemID=i.itemid"
	    sqlStr = sqlStr + " where S.PlusSaleItemID=" + CStr(FRectItemID)
	    
	    rsget.Open sqlStr,dbget,1
	    FResultCount = rsget.RecordCount
	    if Not rsget.Eof then
	        set FOneItem = new CPlusSaleSubItem
			FOneItem.FPlusSaleItemID= rsget("PlusSaleItemID")
			FOneItem.FPlusSalePro   = rsget("PlusSalePro")
			FOneItem.FPlusSaleMargin     = rsget("PlusSaleMargin")
            FOneItem.FPlusSaleMaginFlag  = rsget("PlusSaleMaginFlag")
            FOneItem.FPlusSaleStartDate  = rsget("PlusSaleStartDate")
            FOneItem.FPlusSaleEndDate    = rsget("PlusSaleEndDate")
            FOneItem.Fregdate           = rsget("regdate")


			FOneItem.FItemName     = db2html(rsget("itemname"))
			FOneItem.Fmakerid      = rsget("makerid")
			FOneItem.FSellCash     = rsget("sellcash")
			FOneItem.FBuycash      = rsget("buycash")
			FOneItem.Fmwdiv        = rsget("mwdiv")
			FOneItem.FOrgPrice     = rsget("orgprice")
			FOneItem.FOrgSuplycash = rsget("OrgSuplycash")
			FOneItem.FSellyn       = rsget("sellyn")
			FOneItem.FLimitYn      = rsget("limityn")
			FOneItem.FLimitNo      = rsget("limitno")
			FOneItem.FLimitSold    = rsget("limitsold")

			FOneItem.FSailYN		  = rsget("sailyn")
            FOneItem.FLinkedItemCount = rsget("LinkedItemCount")
            
			FOneItem.FImageSmall   = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FOneItem.FPlusSaleItemID) + "/" + rsget("smallimage")

	    end if
	    rsget.Close
    end function

	public function GetPlusSaleSubItemList()
	    dim sqlStr,i
	    sqlStr = " select count(S.PlusSaleItemID) as cnt "
	    sqlStr = sqlStr + "	from db_item.dbo.tbl_PlusSaleRegedItem S"
	    sqlStr = sqlStr + "	    Join db_item.dbo.tbl_item i"
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
	    
	    rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close
		
	    sqlStr = "select top " + CStr(FPageSize*FCurrPage)
	    sqlStr = sqlStr + "	S.*,  i.makerid, i.itemName, i.sellcash, i.buycash, i.mwdiv, i.orgprice, i.OrgSuplycash, i.sellyn, i.limityn, i.limitno, i.limitsold, i.sailyn, i.smallimage, "
	    sqlStr = sqlStr + "	i.itemcouponyn, i.curritemcouponidx, i.itemcoupontype, i.itemcouponvalue,"
	    sqlStr = sqlStr + "	(select count(T.PlusSaleItemID) from db_item.dbo.tbl_PlusSaleLinkItemList T where T.PlusSaleItemID=S.PlusSaleItemID) as LinkedItemCount"
        sqlStr = sqlStr + "	from  db_item.dbo.tbl_PlusSaleRegedItem S"
	    sqlStr = sqlStr + "	    Join db_item.dbo.tbl_item i"
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
	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new CPlusSaleSubItem
				FItemList(i).FPlusSaleItemID= rsget("PlusSaleItemID")
				FItemList(i).FPlusSalePro   = rsget("PlusSalePro")
    			FItemList(i).FPlusSaleMargin     = rsget("PlusSaleMargin")
                FItemList(i).FPlusSaleMaginFlag  = rsget("PlusSaleMaginFlag")
                FItemList(i).FPlusSaleStartDate  = rsget("PlusSaleStartDate")
                FItemList(i).FPlusSaleEndDate    = rsget("PlusSaleEndDate")
                FItemList(i).Fregdate           = rsget("regdate")
            
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).Fmakerid      = rsget("makerid")
				FItemList(i).FSellCash     = rsget("sellcash")
				FItemList(i).FBuyCash     = rsget("buycash")
				FItemList(i).Fmwdiv        = rsget("mwdiv")
				FItemList(i).FOrgPrice     = rsget("orgprice")
				FItemList(i).FOrgSuplycash = rsget("OrgSuplycash")
				FItemList(i).FSellyn       = rsget("sellyn")
				FItemList(i).FLimitYn      = rsget("limityn")
				FItemList(i).FLimitNo      = rsget("limitno")
				FItemList(i).FLimitSold    = rsget("limitsold")

				FItemList(i).FSailYN		  = rsget("sailyn")
                FItemList(i).FLinkedItemCount = rsget("LinkedItemCount")
                
				FItemList(i).FImageSmall   = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FPlusSaleItemID) + "/" + rsget("smallimage")

                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
            
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
		
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
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
End Class
%>