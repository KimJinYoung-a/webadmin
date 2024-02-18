<%
Class CBrandBoostKeyMngItem
    public FBrandBoostKeyword
    public Fmakerid
    public FsellitemCNT
    public Fsocname
    public Fsocname_kor
    public Fbrandboostkeyusing
    public FbrandboostkeyRegdate
    public Freguserid

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CDispCateMngItem
    public FCateCode
	public FDepth
	public FCateName
	public FCateName_E
	public FUseYN
	public FSortNo
	public FIsNew
	public FJaehuname
	public Fsitegubun  ''?
	
	public FMetaKeywords
	public FSafetyInfoType
	public FsearchKeywords
	
	public FCateFullName
	public FSellItemCnt
	
	public FCateBoostKeyword
	public Fcateboostkeyusing
	public FcateboostkeyRegdate
	public Freguserid
	
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class


Class CDispCateKeywordsMng
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	public FRectDispCate
	public FRectCateUsing
	public FRectMi_metakey
	public FRectMi_searchkey
	public FRectSearchKeyword
	public FRectMetaKeyword
	public FRectBoostKateUsing
	public FRectMakerid
	public FRectBoostBrandUsing
	
	public function getBrandBoostKeywordsList
	    dim sqlStr 
	    Dim paramInfo
	    
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
        	,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize)	_
        	,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage) _
        	,Array("@catecode"	, adBigInt	, adParamInput	,      , FRectDispCate) _
        	,Array("@cateusing"	, adVarchar	, adParamInput	,   1   , FRectCateUsing) _
        	,Array("@keycateusing"	, adVarchar	, adParamInput	,   1   , FRectBoostKateUsing) _
        	,Array("@keyword"	, adVarchar	, adParamInput	,   32   , FRectSearchKeyword) _
        	
        )
        
        sqlStr = "db_const.[dbo].[usp_Ten_Const_Brand_BoostKeyword_LIST]"
        
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
                set FItemList(i) = new CBrandBoostKeyMngItem
                FItemList(i).FBrandBoostKeyword   = rsget("keyword")  
                FItemList(i).Fmakerid   = rsget("makerid")
                FItemList(i).FsellitemCNT   = rsget("sellitemCNT")
                FItemList(i).Fsocname   = db2Html(rsget("socname"))
                FItemList(i).Fsocname_kor = db2Html(rsget("socname_kor"))
                FItemList(i).Fbrandboostkeyusing    = rsget("brandboostkeyusing")
                FItemList(i).FbrandboostkeyRegdate   = rsget("regdate")
                FItemList(i).Freguserid = rsget("reguserid")
    
                rsget.movenext
                i=i+1
            loop
        end if
    end function

	public function getDispCateBoostKeywordsList
	    dim sqlStr 
	    Dim paramInfo
	    
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
        	,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize)	_
        	,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage) _
        	,Array("@catecode"	, adBigInt	, adParamInput	,      , FRectDispCate) _
        	,Array("@cateusing"	, adVarchar	, adParamInput	,   1   , FRectCateUsing) _
        	,Array("@keycateusing"	, adVarchar	, adParamInput	,   1   , FRectBoostKateUsing) _
        	,Array("@keyword"	, adVarchar	, adParamInput	,   32   , FRectSearchKeyword) _
        	
        )
        
        sqlStr = "db_const.[dbo].[usp_Ten_Const_display_cate_BoostKeyword_LIST]"
        
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
                set FItemList(i) = new CDispCateMngItem
                    FItemList(i).FCateBoostKeyword = rsget("keyword")
                    FItemList(i).FCateCode      = rsget("catecode")
                	FItemList(i).FDepth         = rsget("depth")
                	FItemList(i).FCateName      = rsget("catename")
                	FItemList(i).FCateName_E    = rsget("catename_e")
                	FItemList(i).FUseYN         = rsget("useyn")
                	FItemList(i).FSortNo        = rsget("sortno")
                	'FItemList(i).FIsNew         = rsget("isnew")
                	FItemList(i).FJaehuname     = rsget("jaehuname")
                	''FItemList(i).Fsitegubun 
                	
                	FItemList(i).FMetaKeywords  = rsget("metakeywords")
                	''FItemList(i).FSafetyInfoType = rsget("metakeywords")
                	FItemList(i).FsearchKeywords = rsget("searchKeywords")
                	
                	FItemList(i).FCateFullName = rsget("cateCodeFullDepthName")
                	FItemList(i).FSellItemCnt = rsget("sellitemcnt")
                	
                	FItemList(i).Fcateboostkeyusing = rsget("cateboostkeyusing")
                	FItemList(i).FcateboostkeyRegdate = rsget("regdate")
                	FItemList(i).Freguserid = rsget("reguserid")
                	if isNULL(FItemList(i).FCateFullName) then FItemList(i).FCateFullName=""
                rsget.movenext
                i=i+1
            loop
        end if
    end function

	public function getDispCateKeywords_CurrentSellitem
	    dim sqlStr 
	    Dim paramInfo
	    
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
        	,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize)	_
        	,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage) _
        	,Array("@catecode"	, adBigInt	, adParamInput	,      , FRectDispCate) _
        	,Array("@cateusing"	, adVarchar	, adParamInput	,   1   , FRectCateUsing) _
        	,Array("@mimetakey"	, adVarchar	, adParamInput	,   10   , FRectMi_metakey) _
        	,Array("@misearchkey"	, adVarchar	, adParamInput	,   10   , FRectMi_searchkey) _
        	,Array("@searchkeyword"	, adVarchar	, adParamInput	,   32   , FRectSearchKeyword) _
        	,Array("@metakeyword"	, adVarchar	, adParamInput	,   32   , FRectmetaKeyword) _
        	
        )
        
        sqlStr = "db_const.[dbo].[usp_Ten_Const_display_cate_ExistsSellitem_LIST]"
        
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
                set FItemList(i) = new CDispCateMngItem
                    FItemList(i).FCateCode      = rsget("catecode")
                	FItemList(i).FDepth         = rsget("depth")
                	FItemList(i).FCateName      = rsget("catename")
                	FItemList(i).FCateName_E    = rsget("catename_e")
                	FItemList(i).FUseYN         = rsget("useyn")
                	FItemList(i).FSortNo        = rsget("sortno")
                	'FItemList(i).FIsNew         = rsget("isnew")
                	FItemList(i).FJaehuname     = rsget("jaehuname")
                	''FItemList(i).Fsitegubun 
                	
                	FItemList(i).FMetaKeywords  = rsget("metakeywords")
                	''FItemList(i).FSafetyInfoType = rsget("metakeywords")
                	FItemList(i).FsearchKeywords = rsget("searchKeywords")
                	
                	FItemList(i).FCateFullName = rsget("cateCodeFullDepthName")
                	FItemList(i).FSellItemCnt = rsget("sellitemcnt")
                	FItemList(i).FCateBoostKeyword = rsget("CateBoostKeyword")
                	
                	if isNULL(FItemList(i).FCateFullName) then FItemList(i).FCateFullName=""
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