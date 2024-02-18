<%
Class CimageLinkItem
	public fidx
	Public fgubun
	Public Fmenuidx 
	Public Fadminid
	Public Flastadminid
	public Fisusing 
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fusername2
	Public Fmaincopy1
	Public Fmaincopy2

	Public FsubIdx
	Public Flistidx
	Public Fsortnum
	Public FitemName
	Public FsmallImage
	Public FItemid

	Public Fxmlregdate
	Public Fmourl	'모바일 URL
	Public Fappurl	'앱 URL
	Public Fpcurl	'pc URL
	Public Flabel	'라벨딱지
	Public Fldv		'할인 쿠폰
	Public Fis1day

	Public FsubImage1
	Public Fextraurl

	Public Fsubtitle '// 주말특가용
	Public Fsaleper
	Public FbannerImg
	Public FlinkUrl
	Public FaltName
	Public FbannerNameEng
	Public FbannerNameKor
	Public FsubCopy
    Public FWidth
    Public FHeight
	Public FImage
	Public FMasterIdx
	Public FXValue
	Public FYValue
	Public FWValue
	Public FHValue
	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CimageLink
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
       
    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	Public FRectSubIdx
	Public FRectlistidx
	Public FRectMasterIdx
	Public FRectDevice
	
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_event.dbo.tbl_ImageLink_Master "
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr + " a.*, u.username, u2.username as username2 "
        sqlStr = sqlStr + " from db_event.dbo.tbl_ImageLink_Master as a "
        sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten as u on a.reguser = u.userid "
        sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten as u2 on a.modifyuser = u2.userid "
        sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and a.isusing='" + CStr(Fisusing) + "'"
        end If
        
		sqlStr = sqlStr + " order by  a.idx desc, a.LastUpDate DESC" 

		'response.write sqlStr &"<br>"
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
				set FItemList(i) = new CimageLinkItem
				
				FItemList(i).fidx				= rsget("idx")
				FItemList(i).Fadminid			= rsget("reguser")
				FItemList(i).Flastadminid		= rsget("modifyuser")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fusername			= rsget("username")
				FItemList(i).Fusername2			= rsget("username2")
				FItemList(i).FImage				= rsget("image")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub
    
	'//subitem
    public Sub GetOneDetailContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_event.dbo.tbl_ImageLink_Detail "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CimageLinkItem
        
        if Not rsget.Eof Then
			FOneItem.fidx			= rsget("idx")
			FOneItem.FMasterIdx		= rsget("masteridx")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Flastupdate	= rsget("lastupdate")
			FOneItem.Fadminid		= rsget("reguser")
			FOneItem.Flastadminid	= rsget("modifyuser")
			FOneItem.FXValue		= rsget("XValue")
			FOneItem.FYValue		= rsget("YValue")
			FOneItem.FWValue		= rsget("WValue")
			FOneItem.FHValue		= rsget("HValue")
			FOneItem.FLinkURL		= rsget("LinkURL")

        end If
        
        rsget.Close
    end Sub
    
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_event.dbo.tbl_ImageLink_Master "
        sqlStr = sqlStr + " where menuidx=" & CStr(FRectIdx)
		sqlStr = sqlStr + " and device='" & CStr(FRectDevice) & "'"

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CimageLinkItem
        
        if Not rsget.Eof Then
			FOneItem.fidx			= rsget("idx")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Flastupdate	= rsget("lastupdate")
			FOneItem.FImage			= rsget("image")
			FOneItem.Fadminid		= rsget("reguser")
			FOneItem.Flastadminid	= rsget("modifyuser")

        end If
        
        rsget.Close
    end Sub

    public Sub GetMasterContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_event.dbo.tbl_ImageLink_Master "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CimageLinkItem
        
        if Not rsget.Eof Then
			FOneItem.fidx			= rsget("idx")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Flastupdate	= rsget("lastupdate")
			FOneItem.FImage			= rsget("image")
			FOneItem.Fadminid		= rsget("reguser")
			FOneItem.Flastadminid	= rsget("modifyuser")

        end If
        
        rsget.Close
    end Sub
    
    public Sub GetLinkListContents()
       dim sqlStr, addSql, i

		sqlStr = " select count(*) as cnt from db_event.dbo.tbl_ImageLink_Detail "
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr & " and  masterIdx='" & FRectMasterIdx & "'"
        sqlStr = sqlStr + " and isusing='Y'"

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "Select top 1000 s.idx , s.MasterIdx , s.XValue , s.YValue, s.WValue , s.HValue"
		sqlStr = sqlStr & ", s.LinkURL, s.IsUsing, s.RegUser, s.ModifyUser, s.RegDate, s.LastUpDate"
		sqlStr = sqlStr & " ,u.username, u2.username as username2 "
        sqlStr = sqlStr & "From db_event.dbo.tbl_ImageLink_Detail as s "
        sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten as u on s.reguser = u.userid "
        sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten as u2 on s.modifyuser = u2.userid "
        sqlStr = sqlStr & "Where masterIdx='" & FRectMasterIdx & "'"
        sqlStr = sqlStr & " and s.isusing='Y'"
		sqlStr = sqlStr & " order by idx DESC" 

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new CimageLinkItem
				
				FItemList(i).Fidx					= rsget("idx")
	            FItemList(i).FMasterIdx				= rsget("masterIdx")
	            FItemList(i).FLinkURL				= rsget("LinkURL")
	            FItemList(i).FXValue				= rsget("XValue")
				FItemList(i).FYValue				= rsget("YValue")
				FItemList(i).FWValue				= rsget("WValue")
				FItemList(i).FHValue				= rsget("HValue")
	            FItemList(i).FIsUsing				= rsget("IsUsing")
				FItemList(i).Fadminid				= rsget("reguser")
				FItemList(i).Flastadminid			= rsget("modifyuser")
				FItemList(i).Fregdate				= rsget("regdate")
				FItemList(i).Flastupdate			= rsget("lastupdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
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