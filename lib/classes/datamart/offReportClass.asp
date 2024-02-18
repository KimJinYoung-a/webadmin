<%
Class COffMeachulItem
    public FShopid
    public FMakerid
    public FJungsanGubun
    public FJungsanGubunName
    public Ftotalitemcount
    public FtotalSellSum
    public FtotalRealSellSum
    
    public FtotalJungsanitemcount
    public FtotalJungsanSum

	public Fcomm_name
	public Ftotalcnt
	public FallianceYN
    
    Private Sub Class_Initialize()
        Ftotalitemcount     = 0
        FtotalSellSum       = 0
        FtotalRealSellSum   = 0
	End Sub

	Private Sub Class_Terminate()

    End Sub
End Class


Class COffReport
    public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	public FRectYYYYMM
	public FRectJungsanGubun
	public FRectShopid
	public FRectMakerid
	
	public Sub GetShopMeachulByBrandJungsanGubun2()
	    dim sqlStr, i
	    sqlStr = " select IsNULL(S.shopid,IsNULL(J.shopid,'')) as shopid,"
	    sqlStr = sqlStr & " IsNULL(S.makerid,IsNULL(J.makerid,'')) as makerid,"
	    sqlStr = sqlStr & " IsNULL(S.jungsangubun,IsNULL(J.jungsangubun,'')) as jungsangubun,"
	    sqlStr = sqlStr & " IsNULL(S.jungsangubunname,IsNULL(J.jungsangubunname,'')) as jungsangubunname,"
	    sqlStr = sqlStr & " IsNULL(S.totalitemcount,0) as totalSellitemcount,"
        sqlStr = sqlStr & " IsNULL(S.totalSellSum,0) as totalSellSum,"
        sqlStr = sqlStr & " IsNULL(S.totalRealSellSum,0) as totalRealSellSum,"
        
        sqlStr = sqlStr & " IsNULL(J.totalitemcount,0) as totalJungsanitemcount,"
        sqlStr = sqlStr & " IsNULL(J.totalJungsanSum,0) as totalJungsanSum"
        sqlStr = sqlStr & " from ("
	    sqlStr = sqlStr & "     select shopid, jungsangubun, makerid, jungsangubunname,"
	    sqlStr = sqlStr & "     sum(totalitemcount) as totalitemcount,"
	    sqlStr = sqlStr & "     sum(totalSellSum) as totalSellSum,"
	    sqlStr = sqlStr & "     sum(totalRealSellSum) as totalRealSellSum"
	    sqlStr = sqlStr & "     from db_datamart.dbo.tbl_off_monthly_sell_summary"
	    sqlStr = sqlStr & "     where yyyymm='" & FRectYYYYMM & "'"
	    
	    if (FRectShopid<>"") then
	        sqlStr = sqlStr & " and shopid='" & FRectShopid & "'"
	    end if
	    
	    if (FRectMakerid<>"") then
	        sqlStr = sqlStr & " and makerid='" & FRectMakerid & "'"
	    end if
	    
	    if (FRectJungsanGubun="0000") then
            sqlStr = sqlStr & " and jungsangubun=''"
	    elseif (FRectJungsanGubun<>"") then
	        sqlStr = sqlStr & " and jungsangubun='" & FRectJungsanGubun & "'"
	    end if
	    
	    sqlStr = sqlStr & "     group by shopid, makerid, jungsangubun, jungsangubunname"
	    sqlStr = sqlStr & " ) S"
	    sqlStr = sqlStr & " 	full join ("
        sqlStr = sqlStr & " 		select j.shopid, j.makerid, j.jungsangubun, j.jungsangubunname,"
        sqlStr = sqlStr & " 		sum(j.totalitemcount) as totalitemcount,"
        sqlStr = sqlStr & " 		sum(j.totalJungsanSum) as totalJungsanSum"
        sqlStr = sqlStr & "  		from db_datamart.dbo.tbl_off_monthly_jungsan_summary j"
        sqlStr = sqlStr & " 		where j.yyyymm='" & FRectYYYYMM & "'"
        
        if (FRectShopid<>"") then
            sqlStr = sqlStr & " 	and (j.shopid='" & FRectShopid & "' or  j.shopid='')"   ''오프매입건은 Shopid가 없음...
        end if
        
        if (FRectMakerid<>"") then
	        sqlStr = sqlStr & " and makerid='" & FRectMakerid & "'"
	    end if
	    
        if (FRectJungsanGubun<>"") then
	        sqlStr = sqlStr & "     and j.jungsangubun='" & FRectJungsanGubun & "'"
	    end if
        sqlStr = sqlStr & " 		group by j.shopid, j.makerid, j.jungsangubun, j.jungsangubunname"
        sqlStr = sqlStr & " 	) J on J.shopid=s.shopid and J.jungsangubun=S.jungsangubun and J.makerid=S.makerid "
	    sqlStr = sqlStr & " order by shopid, makerid, jungsangubun"

	    db3_rsget.open sqlStr,db3_dbget,1	
	        FTotalCount = db3_rsget.RecordCount
	        FResultCount  = FTotalCount
	        
	        redim preserve FItemList(FResultCount)
    		i=0
    		if  not db3_rsget.EOF  then
    			do until db3_rsget.eof
    				set FItemList(i) = new COffMeachulItem
    				FItemList(i).FShopid            = db3_rsget("Shopid")
                    FItemList(i).FJungsanGubun      = db3_rsget("JungsanGubun")
                    FItemList(i).FJungsanGubunName  = db3_rsget("JungsanGubunName")
                    FItemList(i).FMakerid           = db3_rsget("makerid")
                    FItemList(i).Ftotalitemcount    = db3_rsget("totalSellitemcount")
                    FItemList(i).FtotalSellSum      = db3_rsget("totalSellSum")
                    FItemList(i).FtotalRealSellSum  = db3_rsget("totalRealSellSum")
                    
                    FItemList(i).FtotalJungsanitemcount = db3_rsget("totalJungsanitemcount")
                    FItemList(i).FtotalJungsanSum       = db3_rsget("totalJungsanSum")
                    
    				i=i+1
    				db3_rsget.moveNext
    			loop
    		end if
	    db3_rsget.Close
	    
    end sub
    
	public Sub GetShopMeachulByBrandJungsanGubun()
	    dim sqlStr, i
	    
	    sqlStr = "select shopid, jungsangubun, makerid, jungsangubunname,"
	    sqlStr = sqlStr & " sum(totalitemcount) as totalitemcount,"
	    sqlStr = sqlStr & " sum(totalSellSum) as totalSellSum,"
	    sqlStr = sqlStr & " sum(totalRealSellSum) as totalRealSellSum"
	    sqlStr = sqlStr & " from db_datamart.dbo.tbl_off_monthly_sell_summary"
	    sqlStr = sqlStr & " where yyyymm='" & FRectYYYYMM & "'"
	    
	    if (FRectShopid<>"") then
	        sqlStr = sqlStr & " and shopid='" & FRectShopid & "'"
	    end if
	    
	    if (FRectJungsanGubun="0000") then
            sqlStr = sqlStr & " and jungsangubun=''"
	    elseif (FRectJungsanGubun<>"") then
	        sqlStr = sqlStr & " and jungsangubun='" & FRectJungsanGubun & "'"
	    end if
	    
	    sqlStr = sqlStr & " group by shopid, makerid, jungsangubun, jungsangubunname"
	    sqlStr = sqlStr & " order by shopid, makerid, jungsangubun"

	    db3_rsget.open sqlStr,db3_dbget,1	
	        FTotalCount = db3_rsget.RecordCount
	        FResultCount  = FTotalCount
	        
	        redim preserve FItemList(FResultCount)
    		i=0
    		if  not db3_rsget.EOF  then
    			do until db3_rsget.eof
    				set FItemList(i) = new COffMeachulItem
    				FItemList(i).FShopid            = db3_rsget("Shopid")
                    FItemList(i).FJungsanGubun      = db3_rsget("JungsanGubun")
                    FItemList(i).FJungsanGubunName  = db3_rsget("JungsanGubunName")
                    FItemList(i).FMakerid           = db3_rsget("makerid")
                    FItemList(i).Ftotalitemcount    = db3_rsget("totalitemcount")
                    FItemList(i).FtotalSellSum      = db3_rsget("totalSellSum")
                    FItemList(i).FtotalRealSellSum  = db3_rsget("totalRealSellSum")

    				i=i+1
    				db3_rsget.moveNext
    			loop
    		end if
	    db3_rsget.Close
	    
    end sub
    
    public Sub GetShopMeachulByJungsanGubun2()
        dim sqlStr, i
        sqlStr = " select IsNULL(S.shopid,IsNULL(J.shopid,'')) as shopid, IsNULL(S.jungsangubun,IsNULL(J.jungsangubun,'')) as jungsangubun, IsNULL(S.jungsangubunname,IsNULL(J.jungsangubunname,'')) as jungsangubunname,"
        sqlStr = sqlStr & " IsNULL(S.totalitemcount,0) as totalSellitemcount,"
        sqlStr = sqlStr & " IsNULL(S.totalSellSum,0) as totalSellSum,"
        sqlStr = sqlStr & " IsNULL(S.totalRealSellSum,0) as totalRealSellSum,"
        
        sqlStr = sqlStr & " IsNULL(J.totalitemcount,0) as totalJungsanitemcount,"
        sqlStr = sqlStr & " IsNULL(J.totalJungsanSum,0) as totalJungsanSum"
        sqlStr = sqlStr & " from  ("
        sqlStr = sqlStr & " 	select s.shopid, s.jungsangubun, s.jungsangubunname,"
        sqlStr = sqlStr & " 	sum(s.totalitemcount) as totalitemcount,"
        sqlStr = sqlStr & " 	sum(s.totalSellSum) as totalSellSum,"
        sqlStr = sqlStr & " 	sum(s.totalRealSellSum) as totalRealSellSum"
        sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_off_monthly_sell_summary s"
        sqlStr = sqlStr & " 	where s.yyyymm='" & FRectYYYYMM & "'"
        if (FRectShopid<>"") then
            sqlStr = sqlStr & " 	and s.shopid='" & FRectShopid & "'"
        end if
        
        if (FRectJungsanGubun="0000") then
            sqlStr = sqlStr & " and jungsangubun=''"
        elseif (FRectJungsanGubun<>"") then
	        sqlStr = sqlStr & " and jungsangubun='" & FRectJungsanGubun & "'"
	    end if
        sqlStr = sqlStr & " 	group by s.shopid, s.jungsangubun, s.jungsangubunname"
        	
        sqlStr = sqlStr & " ) S"
        sqlStr = sqlStr & " 	full join ("
        sqlStr = sqlStr & " 		select j.shopid, j.jungsangubun, j.jungsangubunname,"
        sqlStr = sqlStr & " 		sum(j.totalitemcount) as totalitemcount,"
        sqlStr = sqlStr & " 		sum(j.totalJungsanSum) as totalJungsanSum"
        sqlStr = sqlStr & "  		from db_datamart.dbo.tbl_off_monthly_jungsan_summary j"
        sqlStr = sqlStr & " 		where j.yyyymm='" & FRectYYYYMM & "'"
        if (FRectShopid<>"") then
            sqlStr = sqlStr & " 	and (j.shopid='" & FRectShopid & "' or  j.shopid='')"   ''오프매입건은 Shopid가 없음...
        end if
    
        if (FRectJungsanGubun<>"") then
	        sqlStr = sqlStr & "     and j.jungsangubun='" & FRectJungsanGubun & "'"
	    end if
        sqlStr = sqlStr & " 		group by j.shopid, j.jungsangubun, j.jungsangubunname"
        sqlStr = sqlStr & " 	) J on J.shopid=s.shopid and J.jungsangubun=S.jungsangubun  "
        sqlStr = sqlStr & " order by shopid, jungsangubun"
        
        db3_rsget.open sqlStr,db3_dbget,1	
	        FTotalCount = db3_rsget.RecordCount
	        FResultCount  = FTotalCount
	        
	        redim preserve FItemList(FResultCount)
    		i=0
    		if  not db3_rsget.EOF  then
    			do until db3_rsget.eof
    				set FItemList(i) = new COffMeachulItem
    				FItemList(i).FShopid            = db3_rsget("Shopid")
                    FItemList(i).FJungsanGubun      = db3_rsget("JungsanGubun")
                    FItemList(i).FJungsanGubunName  = db3_rsget("JungsanGubunName")
                    FItemList(i).Ftotalitemcount    = db3_rsget("totalSellitemcount")
                    FItemList(i).FtotalSellSum      = db3_rsget("totalSellSum")
                    FItemList(i).FtotalRealSellSum  = db3_rsget("totalRealSellSum")
                    
                    FItemList(i).FtotalJungsanitemcount = db3_rsget("totalJungsanitemcount")
                    FItemList(i).FtotalJungsanSum       = db3_rsget("totalJungsanSum")
    				i=i+1
    				db3_rsget.moveNext
    			loop
    		end if
	    db3_rsget.Close
    end Sub
	
	public Sub GetShopMeachulByJungsanGubun()
	    dim sqlStr, i
	    
	    sqlStr = "select shopid, jungsangubun, jungsangubunname,"
	    sqlStr = sqlStr & " sum(totalitemcount) as totalitemcount, sum(totalSellSum) as totalSellSum, sum(totalRealSellSum) as totalRealSellSum"
	    sqlStr = sqlStr & " from db_datamart.dbo.tbl_off_monthly_sell_summary"
	    sqlStr = sqlStr & " where yyyymm='" & FRectYYYYMM & "'"
	    
	    if (FRectShopid<>"") then
	        sqlStr = sqlStr & " and shopid='" & FRectShopid & "'"
	    end if
	    
	    if (FRectJungsanGubun<>"") then
	        sqlStr = sqlStr & " and jungsangubun='" & FRectJungsanGubun & "'"
	    end if
	    
	    sqlStr = sqlStr & " group by shopid, jungsangubun, jungsangubunname"
	    sqlStr = sqlStr & " order by shopid, jungsangubun"
	    
	    db3_rsget.open sqlStr,db3_dbget,1	
	        FTotalCount = db3_rsget.RecordCount
	        FResultCount  = FTotalCount
	        
	        redim preserve FItemList(FResultCount)
    		i=0
    		if  not db3_rsget.EOF  then
    			do until db3_rsget.eof
    				set FItemList(i) = new COffMeachulItem
    				FItemList(i).FShopid            = db3_rsget("Shopid")
                    FItemList(i).FJungsanGubun      = db3_rsget("JungsanGubun")
                    FItemList(i).FJungsanGubunName  = db3_rsget("JungsanGubunName")
                    FItemList(i).Ftotalitemcount    = db3_rsget("totalitemcount")
                    FItemList(i).FtotalSellSum      = db3_rsget("totalSellSum")
                    FItemList(i).FtotalRealSellSum  = db3_rsget("totalRealSellSum")

    				i=i+1
    				db3_rsget.moveNext
    			loop
    		end if
	    db3_rsget.Close
    end Sub

	public Sub GetSoldoutCancelOrderSet()
	    dim sqlStr, i
	    
		sqlStr = "exec db_datamart.dbo.usp_SCM_Statistic_MonthlySoldoutCancelOrder_Set '" & FRectYYYYMM & "'" & vbcrlf
		db3_dbget.Execute sqlStr
    end Sub

	public Sub GetSoldoutCancelOrderInfo1()
	    dim sqlStr, i
	    sqlStr = "select cancelno" & vbcrlf
	    sqlStr = sqlStr & " from db_datamart.dbo.tbl_monthly_cancelno" & vbcrlf
	    sqlStr = sqlStr & " where yyyymm='" & FRectYYYYMM & "'"
	    db3_rsget.open sqlStr,db3_dbget,1
		if  not db3_rsget.EOF  then
			FTotalCount	= db3_rsget("cancelno")
		end if
	    db3_rsget.Close
    end Sub

	public Sub GetSoldoutCancelOrderInfo2()
	    dim sqlStr, i
		sqlStr = "select comm_name, totalcnt" & vbcrlf
	    sqlStr = sqlStr & " from db_datamart.dbo.tbl_monthly_cancelgubun" & vbcrlf
	    sqlStr = sqlStr & " where yyyymm='" & FRectYYYYMM & "'" & vbcrlf
		sqlStr = sqlStr & " order by gubun desc"
	    db3_rsget.open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount
        redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new COffMeachulItem
				FItemList(i).Fcomm_name = db3_rsget("comm_name")
				FItemList(i).Ftotalcnt = db3_rsget("totalcnt")
				i=i+1
				db3_rsget.moveNext
			loop
		end if
	    db3_rsget.Close
    end Sub

	public Sub GetSoldoutCancelOrderInfo3()
	    dim sqlStr, i
		sqlStr = "select allianceYN, totalcnt" & vbcrlf
	    sqlStr = sqlStr & " from db_datamart.dbo.tbl_monthly_soldout_cancel" & vbcrlf
	    sqlStr = sqlStr & " where yyyymm='" & FRectYYYYMM & "'" & vbcrlf
		sqlStr = sqlStr & " order by gubun desc"
	    db3_rsget.open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount
        redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new COffMeachulItem
				FItemList(i).FallianceYN = db3_rsget("allianceYN")
				FItemList(i).Ftotalcnt = db3_rsget("totalcnt")
				i=i+1
				db3_rsget.moveNext
			loop
		end if
	    db3_rsget.Close
    end Sub

    Private Sub Class_Initialize()
        redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
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