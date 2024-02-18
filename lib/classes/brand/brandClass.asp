<%
'###########################################################
' Description : 월간브랜드서비스지수
' History : 서동석 생성
'			2023.11.16 한용민 수정(페이지에 있던 쿼리 클래스로 변경)
'###########################################################

class CBrandServiceList
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

    public fArrList

    public frectyyyy
    public frectmm
    public frectmakerID
    public frectdispCate

    ' /admin/brandStatic/brandServicePoint.asp
	public sub fBrandServiceList()
		dim sqlStr,i

		sqlStr = "exec db_datamart.dbo.sp_Ten_BrandService_Report '" & frectyyyy & "-" & frectmm & "', '" & frectmakerID & "','"& frectdispCate &"'"

		'response.write sqlStr &"<br>"
		db3_rsget.pagesize = FPageSize
        db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.recordcount
        FtotalCount = db3_rsget.recordcount

		if not db3_rsget.EOF  then
		    fArrList = db3_rsget.getRows()
		end if
		db3_rsget.Close
	end sub

end class
%>