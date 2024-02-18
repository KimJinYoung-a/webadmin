<%
Class CRecomKeywordMasterItem
    public Fgroup_no	
    public Fkeyword	
    public Fitemcnt	
    public Fitemid_list	
    public Fitemname_list

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CRecomKeywordDetailItem
	public Fidx
	public Fkeyword
	public Fsmallimage
	public Fitemid
	public Fitemname
	public Fsellyn
	public Fisusing
	public Fmwdiv
	public Fmakerid

	public Fsellcash
	public Flimityn
	public Flimitno
	public Flimitsold

	public Fsailyn
	public Forgprice
	public Fsailprice
	public Fbuycash
	public Forgsuplycash
	public Fsailsuplycash
	public Fitemcouponyn
	public FitemCouponType
	public FitemCouponValue
	public Fcouponbuyprice

	'// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// 세일포함 실제가격
	public Function getRealPrice() '!
		getRealPrice = FSellCash
	end Function

	'// 쿠폰 할인가
	public Function GetCouponDiscountPrice() '?
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function

	public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CRecomKeywordItem
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	
	public FRectSearchKeyword
	public FRectGroup_no
	
	
	public function getRecomKeywordMasterList
	    dim sqlStr , i
	    Dim paramInfo
	    
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@keyword"		, adVarchar	, adParamInput	,   100   , FRectSearchKeyword) _
        	,Array("@pagesize"		, adInteger	, adParamInput	,		, FPageSize)	_
        	,Array("@pageno"		, adInteger	, adParamInput	,		, FCurrPage) _
        )
        
        sqlStr = "db_temp.[dbo].[usp_TEN_ksearch_keyword_recom_master_list]"
        
        Call fnExecSPReturnRSOutput(sqlStr, paramInfo)
        
        FTotalCount = GetValue(paramInfo, "@RETURN_VALUE")
        
	    FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		
		FResultCount = rsget.RecordCount

        if (FResultCount<1) then FResultCount=0
        redim preserve FItemList(FResultCount)
            
        If Not(rsget.EOF) then
            do until rsget.EOF
                set FItemList(i) = new CRecomKeywordMasterItem

                FItemList(i).Fgroup_no   = rsget("group_no")  
                FItemList(i).Fkeyword   = rsget("keyword")
                FItemList(i).Fitemcnt   = rsget("itemcnt")
                FItemList(i).Fitemid_list   = rsget("itemid_list")
                FItemList(i).Fitemname_list = rsget("itemname_list")
    
                rsget.movenext
                i=i+1
            loop
        end if
    end function

	public function getRecomKeywordItemList
	    dim sqlStr , i
	    Dim paramInfo
	    
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@group_no"		, adInteger	, adParamInput	,		, FRectGroup_no)	_
        	,Array("@pagesize"		, adInteger	, adParamInput	,		, FPageSize)	_
        	,Array("@pageno"		, adInteger	, adParamInput	,		, FCurrPage) _
        )
        
        sqlStr = "db_temp.[dbo].[usp_TEN_ksearch_keyword_recom_detail_list]"
        
        Call fnExecSPReturnRSOutput(sqlStr, paramInfo)
        
        FTotalCount = GetValue(paramInfo, "@RETURN_VALUE")
        
	    FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		
		FResultCount = rsget.RecordCount

        if (FResultCount<1) then FResultCount=0
        redim preserve FItemList(FResultCount)
            
        If Not(rsget.EOF) then
            do until rsget.EOF
                set FItemList(i) = new CRecomKeywordDetailItem
                    FItemList(i).Fidx	= rsget("idx")  
					FItemList(i).Fkeyword		= rsget("keyword")
					FItemList(i).Fitemid		= rsget("itemid")
					FItemList(i).Fsmallimage	= rsget("smallimage")
					
					FItemList(i).Fitemname		= rsget("itemname")
					FItemList(i).Fsellyn		= rsget("sellyn")
					FItemList(i).Fisusing		= rsget("isusing")
					FItemList(i).Fmwdiv			= rsget("mwdiv")
					FItemList(i).Fmakerid		= rsget("makerid")

					FItemList(i).Fsellcash		= rsget("sellcash")
					FItemList(i).Flimityn		= rsget("limityn")
					FItemList(i).Flimitno		= rsget("limitno")
					FItemList(i).Flimitsold		= rsget("limitsold")

					FItemList(i).Fsailyn		= rsget("sailyn")
					FItemList(i).Forgprice		= rsget("orgprice")
					FItemList(i).Fsailprice		= rsget("sailprice")
					FItemList(i).Fbuycash		= rsget("buycash")
					FItemList(i).Forgsuplycash	= rsget("orgsuplycash")
					FItemList(i).Fsailsuplycash = rsget("sailsuplycash")
					
					FItemList(i).Fitemcouponyn	= rsget("itemcouponyn")
					FItemList(i).FitemCouponType	= rsget("itemCouponType")
					FItemList(i).FitemCouponValue	= rsget("itemCouponValue")
					FItemList(i).Fcouponbuyprice	= rsget("couponbuyprice")

					FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                rsget.movenext
                i=i+1
            loop
        end if
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 15
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

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