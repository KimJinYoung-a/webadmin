<%
Class GiftCardImageObj
'기프트 카드
	public Fidx					'시퀀스값
	public FdesignId
	Public FGiftCardImage		'기프트카드 이미지
	Public FGiftCardAlt		    'alt값
	Public FSortNumber		    '정렬번호
	public FAdminRegister		'등록한 스태프 아이디
	public FAdminName			'등록한 스태프 이름
	public FAdminModifyer		'수정한 스태프 아이디
	public FAdminModifyerName	'수정한 스태프 이름
	public FRegistDate			'등록일
    public FLastUpDate			'수정일
	public FIsusing			

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class GiftCardImageCls
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectIdx
	public FRectIsusing
	
    public Sub GetContentsList()
        dim sqlStr, i, sqlWhere

		sqlwhere = ""		

		if FRectIsusing <> "" then
			sqlWhere = sqlWhere +  " and isusing = '" & FRectIsusing & "'"
			sqlWhere = sqlWhere +  " and designid <> 900 "
		end if 				

		sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].[tbl_giftcard_image] "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere
        
		'response.write sqlStr &"<br>"
		'response.end 

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub        	
			
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		sqlStr = sqlStr + "  idx "
		sqlStr = sqlStr + " , designId "
		sqlStr = sqlStr + " , giftcardImage "
		sqlStr = sqlStr + " , giftcardAlt "
		sqlStr = sqlStr + " , sortNumber "
		sqlStr = sqlStr + " , adminRegister "
		sqlStr = sqlStr + " , adminName "
		sqlStr = sqlStr + " , adminModifyer "
		sqlStr = sqlStr + " , adminModifyerName "
		sqlStr = sqlStr + " , registDate "
		sqlStr = sqlStr + " , lastUpDate "
		sqlStr = sqlStr + " , isusing "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_giftcard_image "
        sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere
        
		sqlStr = sqlStr + " order by sortNumber desc" 

'		response.write sqlStr &"<br>"
		
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new GiftCardImageObj
				
				FItemList(i).Fidx			 	 = rsget("idx")
				FItemList(i).FdesignId		 	 = rsget("designId")
				FItemList(i).FGiftCardImage	 	 = rsget("giftcardImage")
				FItemList(i).FGiftCardAlt	 	 = rsget("giftcardAlt")
				FItemList(i).FSortNumber	 	 = rsget("sortNumber")
				FItemList(i).FAdminRegister	 	 = rsget("adminRegister")
				FItemList(i).FAdminName		 	 = rsget("adminName")
				FItemList(i).FAdminModifyer	 	 = rsget("adminModifyer")
				FItemList(i).FAdminModifyerName	 = rsget("adminModifyerName")
				FItemList(i).FRegistDate		 = rsget("registDate")
				FItemList(i).FLastUpDate		 = rsget("lastUpDate")									
				FItemList(i).FIsusing			 = rsget("isusing")									

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub    
    
    public Sub GetOneContent()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.[tbl_giftcard_image] "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new GiftCardImageObj
        
        if Not rsget.Eof Then	
			FOneItem.Fidx			   = rsget("idx")
			FOneItem.FdesignId		   = rsget("designId")
			FOneItem.FGiftCardImage	   = rsget("giftcardImage")
			FOneItem.FGiftCardAlt	   = rsget("giftcardAlt")
			FOneItem.FSortNumber	   = rsget("sortNumber")
			FOneItem.FAdminRegister	   = rsget("adminRegister")
			FOneItem.FAdminName		   = rsget("adminName")
			FOneItem.FAdminModifyer	   = rsget("adminModifyer")
			FOneItem.FAdminModifyerName= rsget("adminModifyerName")
			FOneItem.FRegistDate	   = rsget("registDate")
			FOneItem.FLastUpDate	   = rsget("lastUpDate")
			FOneItem.FIsusing		   = rsget("isusing")			
        end If
        
        rsget.Close
    end Sub
    
    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
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
end Class
%>