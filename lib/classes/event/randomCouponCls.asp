<%
'###########################################################
' Description : ·£´ý ÄíÆù Å¬·¡½º
' Hieditor : 2023.07.26 Á¤ÅÂÈÆ »ý¼º
'###########################################################

Class RandomCouponContentsCls
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public frate
	public fcoupon
	public fdeleteYN
	public fregdate
    public fcouponname
    public fcouponvalue
    public fminbuyprice
    public ftCount
end class

class RandomCouponCls
	public FItemList()
	public FResultCount
	
	public FRectEvtCode
	
	public sub getRandomCouponList()
		dim sqlStr, sqlsearch, i

		if FRectEvtCode <> "" then
			sqlsearch = sqlsearch & " AND RC.evt_code = "& FRectEvtCode &""
		end if
		
		sqlStr = "SELECT"
		sqlStr = sqlStr & " RC.idx, RC.rate, RC.coupon, RC.deleteYN, RC.regdate, CM.couponname, CM.couponvalue, CM.minbuyprice"
		sqlStr = sqlStr & " FROM [db_event].[dbo].[tbl_event_random_coupon] AS RC"
        sqlStr = sqlStr & " JOIN [db_user].[dbo].[tbl_user_coupon_master] AS CM"
        sqlStr = sqlStr & " ON RC.coupon = CM.idx"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		sqlStr = sqlStr & " ORDER BY RC.idx ASC"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordcount
        redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new RandomCouponContentsCls
				FItemList(i).fidx = rsget("idx")
				FItemList(i).frate = rsget("rate")
				FItemList(i).fcoupon = rsget("coupon")
				FItemList(i).fdeleteYN = rsget("deleteYN")
				FItemList(i).fregdate = rsget("regdate")
                FItemList(i).fcouponname = rsget("couponname")
                FItemList(i).fcouponvalue = rsget("couponvalue")
                FItemList(i).fminbuyprice = rsget("minbuyprice")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub getRandomCouponIssueList()
		dim sqlStr, sqlsearch, i

		if FRectEvtCode <> "" then
			sqlsearch = sqlsearch & " AND RC.evt_code = "& FRectEvtCode &""
		end if
		
		sqlStr = "SELECT"
		sqlStr = sqlStr & " RC.coupon, COUNT(RC.coupon) AS tCount, CONVERT(VARCHAR(10),UC.regdate,21) AS GiveDate"
		sqlStr = sqlStr & " FROM [db_event].[dbo].[tbl_event_random_coupon] AS RC"
        sqlStr = sqlStr & " JOIN [db_user].[dbo].[tbl_user_coupon] AS UC"
        sqlStr = sqlStr & " ON RC.coupon = UC.masteridx"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
        sqlStr = sqlStr & " GROUP BY RC.coupon, CONVERT(VARCHAR(10),UC.regdate,21)"
		sqlStr = sqlStr & " ORDER BY CONVERT(VARCHAR(10),UC.regdate,21) DESC, RC.coupon ASC"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordcount
        redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new RandomCouponContentsCls
				FItemList(i).fcoupon = rsget("coupon")
                FItemList(i).ftCount = rsget("tCount")
                FItemList(i).fregdate = rsget("GiveDate")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		FResultCount = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class
%>