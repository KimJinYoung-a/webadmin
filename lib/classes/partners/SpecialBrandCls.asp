<%
Class SpecialBrandObj
	public Fidx
	public FBrandid
	public FIsexposure
	public FFrequency
	public FExposure_seq
	public FAlways_exposure
	public FStartdate
	public FEnddate
	public FRegdate
	public FBrand_icon

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class SpecialBrandCls
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	PUBLIC FRectBrandId

    public Sub getSpecialBrandInfo()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_brand.dbo.[tbl_special_brand] "
        sqlStr = sqlStr + " where brandid='"& FRectBrandId &"'" 

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new SpecialBrandObj
        
        if Not rsget.Eof Then	
			FOneItem.Fidx				= rsget("idx")
			FOneItem.FBrandid			= rsget("brandid")
			FOneItem.FIsexposure		= rsget("isexposure")
			FOneItem.FFrequency			= rsget("frequency")
			FOneItem.FExposure_seq		= rsget("exposure_seq")
			FOneItem.FAlways_exposure	= rsget("always_exposure")
			FOneItem.FStartdate			= rsget("startdate")
			FOneItem.FEnddate			= rsget("enddate")
			FOneItem.FRegdate			= rsget("regdate")
			FOneItem.FBrand_icon		= rsget("brand_icon")
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
end Class
%>
