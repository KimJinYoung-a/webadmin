<%
Class CBonusCouponSummaryItem
    public Fbonuscouponidx
    public Fuserlevel
    public FregYYYYMM
    public Fbonuscouponname
    public Fissuedcount
    public Fusingcount
    public Fspendcoupon
    public Fspendmileage
    public Fspendetc
    public Fsubtotalprice
    public Ftotalsum
    public FNotExpiredCount
    public FbaseDate

    public function getUsingPro()
        if Fissuedcount<>0 then
            getUsingPro = CLng(Fusingcount/Fissuedcount*100*100)/100
        end if
    end function
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CBonusCouponSummary

    public FItemList()
	public FOneItem
	
    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
    
    public FRectYYYYMM
    public FRectStartDate
    public FRectEndDate
    public FRectCouponidx
    public FRectUserLevel
    public FRectViewType
    public FRectIncULv
    
    '// 쿠폰 사용 통계(월별_Old)
    public Sub getCouponResultSummary()
        dim sqlStr ,i 
        
        sqlStr = "select count(bonuscouponidx) as cnt "
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_bonuscoupon_using_result "
        sqlStr = sqlStr & " where 1=1"
        if (FRectYYYYMM<>"") then
            sqlStr = sqlStr & " and regyyyymm='" & FRectYYYYMM & "'"
        end if
        
        if (FRectCouponidx<>"") then
            sqlStr = sqlStr & " and bonuscouponidx=" & FRectCouponidx
        end if
        
        if (FRectUserLevel<>"") then
            sqlStr = sqlStr & " and userlevel=" & FRectUserLevel
        end if
        
        db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = db3_rsget("cnt")
        db3_rsget.Close
        
        
        sqlStr = "select top " & (FPagesize*FCurrPage) & " * "
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_bonuscoupon_using_result "
        sqlStr = sqlStr & " where 1=1"
        if (FRectYYYYMM<>"") then
            sqlStr = sqlStr & " and regyyyymm='" & FRectYYYYMM & "'"
        end if
        
        if (FRectCouponidx<>"") then
            sqlStr = sqlStr & " and bonuscouponidx=" & FRectCouponidx
        end if
        
        if (FRectUserLevel<>"") then
            sqlStr = sqlStr & " and userlevel=" & FRectUserLevel
        end if
        sqlStr = sqlStr & " order by regyyyymm desc, bonuscouponidx desc, userlevel"
        
        db3_rsget.pagesize = FPageSize
        db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly
        
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
   'response.write sqlStr     
		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CBonusCouponSummaryItem
				FItemList(i).Fbonuscouponidx      = db3_rsget("bonuscouponidx")
                FItemList(i).Fuserlevel           = db3_rsget("userlevel")
                FItemList(i).FregYYYYMM           = db3_rsget("regYYYYMM")
                FItemList(i).Fbonuscouponname     = db2Html(db3_rsget("bonuscouponname"))
                FItemList(i).Fissuedcount         = db3_rsget("issuedcount")
                FItemList(i).Fusingcount          = db3_rsget("usingcount")
                FItemList(i).Fspendcoupon         = db3_rsget("spendcoupon")
                FItemList(i).Fspendmileage        = db3_rsget("spendmileage")
                FItemList(i).Fspendetc            = db3_rsget("spendetc")
                FItemList(i).Fsubtotalprice       = db3_rsget("subtotalprice")
                FItemList(i).Ftotalsum            = db3_rsget("totalsum")
                FItemList(i).FNotExpiredCount     = db3_rsget("NotExpiredCount")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

    end Sub

    '// 쿠폰 사용 통계(상세_New)
    public Sub getCouponResultSummaryHour()
        dim sqlStr, addSql, grpSql, sortSql, i 

        '검색 조건
        if (FRectStartDate<>"") then
            addSql = addSql & " and baseDate>='" & FRectStartDate & "'"
        end if
        if (FRectEndDate<>"") then
            addSql = addSql & " and baseDate<dateadd(day,1,'" & FRectEndDate & "')"
        end if

        if (FRectCouponidx<>"") then
            addSql = addSql & " and MasterIdx=" & FRectCouponidx
        end if
        
        if (FRectUserLevel<>"") then
            addSql = addSql & " and UserLevel=" & FRectUserLevel
        end if

        grpSql = "YYYYMMDD"
        sortSql = "YYYYMMDD Desc"
        Select Case FRectViewType
            Case "D"    '일자별
                grpSql = "YYYYMMDD"
                sortSql = "YYYYMMDD Desc"
            Case "H"     '시간별
                grpSql = "YYYYMMDD, HH"
                sortSql = "YYYYMMDD Desc, HH Desc"
        End Select

        if FRectIncULv="Y" then
                grpSql = grpSql & ", userlevel"
                sortSql = sortSql & ", userlevel asc"
        end if

        sqlStr = "select COUNT(*) as totCount, CEILING(CAST(COUNT(*) AS FLOAT)/" & FPagesize & ") as totPage "
        sqlStr = sqlStr & " from ("
        sqlStr = sqlStr & "     Select " & grpSql
        sqlStr = sqlStr & "     FROM db_datamart.dbo.tbl_mkt_bonuscoupon_using_result_byHour "
        sqlStr = sqlStr & "     where 1=1" & addSql
        sqlStr = sqlStr & "     group by " & grpSql & " ) as T "
        db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = db3_rsget("totCount")
            FtotalPage  = db3_rsget("totPage")
        db3_rsget.Close
        
        
        sqlStr = "select " & grpSql & ", sum(issuedcount) as issuedcount, sum(UsedCount) as UsedCount, sum(UsedCpnPrice) as UsedCpnPrice "
        sqlStr = sqlStr & " ,sum(UsedTotalMileage) as UsedTotalMileage, sum(UsedTotalPrice) as UsedTotalPrice "
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_mkt_bonuscoupon_using_result_byHour "
        sqlStr = sqlStr & " where 1=1" & addSql
        sqlStr = sqlStr & " group by " & grpSql
        sqlStr = sqlStr & " order by " & sortSql
        sqlStr = sqlStr & " OFFSET " & (FPageSize*(FCurrPage-1)) & " ROWS FETCH NEXT " & FPagesize & " ROWS ONLY"
        db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CBonusCouponSummaryItem
				'FItemList(i).Fbonuscouponidx     = db3_rsget("MasterIdx")
                'FItemList(i).Fuserlevel          = db3_rsget("userlevel")
                'FItemList(i).FbaseDate           = db3_rsget("baseDate")
                FItemList(i).FbaseDate           = db3_rsget("YYYYMMDD")
                if FRectViewType="H" then
                    FItemList(i).FbaseDate = FItemList(i).FbaseDate & " " & db3_rsget("HH") & ":00"
                end if
                if FRectIncULv="Y" then
                    FItemList(i).Fuserlevel = db3_rsget("userlevel")
                end if
                FItemList(i).Fissuedcount        = db3_rsget("issuedcount")
                FItemList(i).Fusingcount         = db3_rsget("UsedCount")
                FItemList(i).Fspendcoupon        = db3_rsget("UsedCpnPrice")
                FItemList(i).Fspendmileage       = db3_rsget("UsedTotalMileage")
                FItemList(i).Fsubtotalprice      = db3_rsget("UsedTotalPrice")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 100
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
end Class

%>