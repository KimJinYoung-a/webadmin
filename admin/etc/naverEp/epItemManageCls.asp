<%

Class epShopFixedPriceItem
    public Fitemid
    public Fitemname
    public Fmakerid
    public Fsellcash
    public Fbuycash
    public Forgprice
    public Forgsuplycash
    public Fsellyn
    public Fmwdiv
    public Fsailyn
    public FImageSmall
    public Flimityn
    public Flimitno
    public Flimitsold
    public Fregsellcash
    public Ffixedcash
    public Fregdt
    public Fupddt
    public Fuseyn
    public FregUserID

    public FmatchNVMid
    public Fnvminprice
    public Fnvtensellcash
    public FNvMaplastupDt

    public Fitemcouponidx  
    public FcouponGubun    
    public Fitemcoupontype 
    public Fitemcouponvalue
    public Fcouponbuyprice 

    public FNotNvMakerisusing '' N 이면 EP적용안함
    public FNotNvItemisusing  '' Y 이면 EP적용안함

    public function getEpExceptStr()
        Dim ret : ret=""
        if (FNotNvMakerisusing="N") then
            ret ="제외브랜드"
        end if

        if (FNotNvItemisusing="Y") Then
            if (ret<>"") then ret=ret&"<br>"
            ret =ret&"제외상품"
        end if

        getEpExceptStr = ret
    end function

    public function getSellcashHtml()
        Dim ret , maycpnPrice, itemcouponTypeStr
        ret = ""
        maycpnPrice = 0

        if (Forgprice>Fsellcash) then
            ret = "<strike>"&FormatNumber(Forgprice,0)&"</strike><br>"
        end if

        if (NOT isNULL(Fitemcouponidx)) then
            if (Fitemcoupontype=1) then
                maycpnPrice = Fsellcash-CLNG(Fsellcash*Fitemcouponvalue/100)
                itemcouponTypeStr = Fitemcouponvalue&"%"
            elseif (Fitemcoupontype=2) then
                maycpnPrice = Fsellcash-Fitemcouponvalue
                itemcouponTypeStr = FormatNumber(Fitemcouponvalue,0)&"원"
            else
                ret = ret & FormatNumber(Fsellcash,0)
            end if

            
            if (maycpnPrice<>0) then
                ret = ret & "<strike>"&FormatNumber(Fsellcash,0)&"</strike>"
                ret = ret & "<br>"
                ret = ret & "<font color='green'>"&FormatNumber(maycpnPrice,0)&"</font>"
                ''ret = ret & " ("&itemcouponTypeStr&")"
            end if
        else
            ret = ret & FormatNumber(Fsellcash,0)
        end if
        
        
        

        getSellcashHtml = ret
    end function

    public function getDiscountTypeHtml()
        Dim ret , itemcouponTypeStr
        if (Forgprice>Fsellcash) then
            if (Forgprice<>0) then
                ret = "<font color='red'>"&CLNG(100-Fsellcash/Forgprice*100)&"%</font>" & " 할인"
            end if
        end if

        if (NOT isNULL(Fitemcouponidx)) then
            if (Fitemcoupontype=1) then
                itemcouponTypeStr = "<font color='green'>"&Fitemcouponvalue&"%</font>"&" 쿠폰"
            elseif (Fitemcoupontype=2) then
                itemcouponTypeStr = "<font color='green'>"&FormatNumber(Fitemcouponvalue,0)&"원</font>"&" 쿠폰"
            end if

            if (FcouponGubun="V") then itemcouponTypeStr = "<strong>"&itemcouponTypeStr&"</strong>"
        end if

        if (ret<>"") and (itemcouponTypeStr<>"") then 
            ret = ret &"<br>"
        end if
        ret = ret &itemcouponTypeStr

        getDiscountTypeHtml = ret
    end function

    public function getItemLimitStatHtml()
		getItemLimitStatHtml = ""
		if isNULL(Flimityn) then Exit function
		if Flimityn<>"Y" then Exit function

		dim limitea : limitea = Flimitno-Flimitsold
		if (limitea<1) then limitea=0

		getItemLimitStatHtml = "한정"&" <font color='Blue'>"&FormatNumber(limitea,0)&"</font>"

		if (limitea<1) then
			getItemLimitStatHtml = "<strong>"&getItemLimitStatHtml&"</strong>"
		end if
	end function

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

End Class

Class epShopFixedPrice
    public FOneItem
    public FItemList()

    public FResultCount
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
    public FScrollCount

    public FRectSellyn
    public FRectMakerid
    public FRectItemIdArr
    public FRectUseYn
    public FRectMwDiv
    public FRectItemCouponYN
    public FRectPriceCheckType
    public FRectEpExceptBrandItem

    public Sub getNVFixedPriceByNvMapXLLIST()
        Dim sqlStr, i
        ''@itemidArr,@makerid,@itemsellyn,@useyn,@mwdiv,@itemcouponyn,@pageSize int = 30
        sqlStr = "exec [db_temp].[dbo].[usp_TEN_NV_FixedPriceByNVMapItem_CNT] '"&FRectItemIdArr&"','"&FRectMakerid&"','"&FRectSellyn&"','"&FRectUseYn&"','"&FRectMwDiv&"','"&FRectItemCouponYN&"','"&FRectPriceCheckType&"','"&FRectEpExceptBrandItem&"','"&FPageSize&"'"
   
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "exec [db_temp].[dbo].[usp_TEN_NV_FixedPriceByNVMapItem_LIST] '"&FRectItemIdArr&"','"&FRectMakerid&"','"&FRectSellyn&"','"&FRectUseYn&"','"&FRectMwDiv&"','"&FRectItemCouponYN&"','"&FRectPriceCheckType&"','"&FRectEpExceptBrandItem&"','"&FPageSize&"','"&FCurrPage&"'"

		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new epShopFixedPriceItem
					FItemList(i).FItemid			= rsget("itemid")
                    FItemList(i).Fitemname          = rsget("itemname")
                    FItemList(i).Fmakerid           = rsget("makerid")
                    FItemList(i).Fsellcash          = rsget("sellcash")
                    FItemList(i).Fbuycash           = rsget("buycash")
                    FItemList(i).Forgprice          = rsget("orgprice")
                    FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                    FItemList(i).Fsellyn            = rsget("sellyn")
                    FItemList(i).Fmwdiv             = rsget("mwdiv")
                    FItemList(i).Fsailyn            = rsget("sailyn")
                    FItemList(i).Flimityn           = rsget("limityn")
                    FItemList(i).Flimitno           = rsget("limitno")
                    FItemList(i).Flimitsold         = rsget("limitsold")
                    FItemList(i).Fregsellcash       = rsget("regsellcash")  ''의미없음..
                    FItemList(i).Ffixedcash         = rsget("fixedcash")
                    FItemList(i).Fregdt             = rsget("regdt")
                    FItemList(i).Fupddt             = rsget("upddt")
                    FItemList(i).Fuseyn             = rsget("useyn")
                    FItemList(i).FregUserID         = rsget("reguserid")
                    FItemList(i).FmatchNVMid        = rsget("matchNVMid")
                    FItemList(i).Fnvminprice        = rsget("nvminprice")
                    FItemList(i).Fnvtensellcash     = rsget("nvtensellcash")
                    FItemList(i).FNvMaplastupDt     = rsget("NvMaplastupDt")

                    FItemList(i).Fitemcouponidx     = rsget("itemcouponidx")
                    FItemList(i).FcouponGubun       = rsget("couponGubun")
                    FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                    FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
                    FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")

					FItemList(i).FImageSmall		= rsget("smallimage")
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall
                    
                    FItemList(i).FNotNvMakerisusing = rsget("NotNvMakerisusing")
                    FItemList(i).FNotNvItemisusing  = rsget("NotNvItemisusing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
    end Sub

    public Sub getNVFixedPriceLIST()
        Dim sqlStr, i
        ''@itemidArr,@makerid,@itemsellyn,@useyn,@mwdiv,@itemcouponyn,@pageSize int = 30
        sqlStr = "exec [db_temp].[dbo].[usp_TEN_NV_FixedPrice_CNT] '"&FRectItemIdArr&"','"&FRectMakerid&"','"&FRectSellyn&"','"&FRectUseYn&"','"&FRectMwDiv&"','"&FRectItemCouponYN&"','"&FRectPriceCheckType&"','"&FRectEpExceptBrandItem&"','"&FPageSize&"'"
   
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "exec [db_temp].[dbo].[usp_TEN_NV_FixedPrice_LIST] '"&FRectItemIdArr&"','"&FRectMakerid&"','"&FRectSellyn&"','"&FRectUseYn&"','"&FRectMwDiv&"','"&FRectItemCouponYN&"','"&FRectPriceCheckType&"','"&FRectEpExceptBrandItem&"','"&FPageSize&"','"&FCurrPage&"'"

		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new epShopFixedPriceItem
					FItemList(i).FItemid			= rsget("itemid")
                    FItemList(i).Fitemname          = rsget("itemname")
                    FItemList(i).Fmakerid           = rsget("makerid")
                    FItemList(i).Fsellcash          = rsget("sellcash")
                    FItemList(i).Fbuycash           = rsget("buycash")
                    FItemList(i).Forgprice          = rsget("orgprice")
                    FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                    FItemList(i).Fsellyn            = rsget("sellyn")
                    FItemList(i).Fmwdiv             = rsget("mwdiv")
                    FItemList(i).Fsailyn            = rsget("sailyn")
                    FItemList(i).Flimityn           = rsget("limityn")
                    FItemList(i).Flimitno           = rsget("limitno")
                    FItemList(i).Flimitsold         = rsget("limitsold")
                    FItemList(i).Fregsellcash       = rsget("regsellcash")  ''의미없음..
                    FItemList(i).Ffixedcash         = rsget("fixedcash")
                    FItemList(i).Fregdt             = rsget("regdt")
                    FItemList(i).Fupddt             = rsget("upddt")
                    FItemList(i).Fuseyn             = rsget("useyn")
                    FItemList(i).FregUserID         = rsget("reguserid")
                    FItemList(i).FmatchNVMid        = rsget("matchNVMid")
                    FItemList(i).Fnvminprice        = rsget("nvminprice")
                    FItemList(i).Fnvtensellcash     = rsget("nvtensellcash")
                    FItemList(i).FNvMaplastupDt     = rsget("NvMaplastupDt")

                    FItemList(i).Fitemcouponidx     = rsget("itemcouponidx")
                    FItemList(i).FcouponGubun       = rsget("couponGubun")
                    FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                    FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
                    FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")

					FItemList(i).FImageSmall		= rsget("smallimage")
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall
                    
                    FItemList(i).FNotNvMakerisusing = rsget("NotNvMakerisusing")
                    FItemList(i).FNotNvItemisusing  = rsget("NotNvItemisusing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
    end Sub

    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

End Class

Class CNvEpAdItem
    public FAccountId
    public FCampaignId
    public FCampaignNm
    public FAdGroupId
    public FAdGroupNm
    public FAdId
    public FOnOff
    public FProductNo
    public FProductNoMall
    public FProductNm
    public FAdProductNm

    public FImageSmall
    
    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

End Class

Class CNvEpAdList
    public FOneItem
    public FItemList()

    public FResultCount
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
    public FScrollCount

    public FRectSellyn
    public FRectMakerid
    public FRectItemID
    public FRectUseYn
    public FRectMwDiv
    'public FRectItemCouponYN

    public function getEpAdGetOneItem
        Dim sqlStr, i
        sqlStr = "exec [db_naver].[dbo].[usp_SCM_Nvad_Item_Get] "&FRectItemID&""

		db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = db3_rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not db3_rsget.EOF Then
			Do until db3_rsget.EOF
				Set FItemList(i) = new CNvEpAdItem
                    FItemList(i).FAccountId     = db3_rsget("AccountId")
                    FItemList(i).FCampaignId    = db3_rsget("CampaignId")
                    FItemList(i).FCampaignNm    = db3_rsget("CampaignNm")
                    FItemList(i).FAdGroupId     = db3_rsget("AdGroupId")
                    FItemList(i).FAdGroupNm     = db3_rsget("AdGroupNm")
                    FItemList(i).FAdId          = db3_rsget("AdId")
                    FItemList(i).FOnOff         = db3_rsget("OnOff")
                    FItemList(i).FProductNo     = db3_rsget("ProductNo")
                    FItemList(i).FProductNoMall = db3_rsget("ProductNoMall")
                    FItemList(i).FProductNm     = db3_rsget("ProductNm")
                    FItemList(i).FAdProductNm   = db3_rsget("AdProductNm")

					'FItemList(i).FImageSmall		= rsget("smallimage")
					'FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall
                    
				i = i + 1
				db3_rsget.moveNext
			Loop
		End If
		db3_rsget.Close
    end function

    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

End Class
%>