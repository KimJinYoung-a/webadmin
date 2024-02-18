<%
Class COSItemItem
    public Fidx
    public Fsitecode
    public Fsiteitemcode
    public Fsiteitemname
    public FrealsellCost
    public ForgsellCost
    public FregTime
    public Fitemid
    public Fitemname
    public Fsellcash
    public Fobrandname
    public Fbrandname

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COSItem
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	public FRectIdx
	public FRectCompareKey
	public FRectRegDate
	public FRectSDate
	public FRectEDate
	public FRectIsMatch
	public FRectSiteCode
	public FRectSiteItemID
	public FItemTotalCount
	
	
    public function fnOtherSiteItemlist()
    	dim sqlStr, sqlsearch
    	
		If FRectRegDate <> "" Then
			sqlsearch = sqlsearch & " and Convert(varchar(10),l.regTime,120) = '" & FRectRegDate & "' "
		End If
		
		If FRectIsMatch <> "" Then
			sqlsearch = sqlsearch & " and m.itemid is " & CHKIIF(FRectIsMatch="o","not","") & " null "
		End If
		
		'// 결과수 카운트
		sqlStr = "select count(l.siteitemcode) as cnt, CEILING(CAST(Count(l.siteitemcode) AS FLOAT)/" & FPageSize & ") AS totPg "
        sqlStr = sqlStr & "FROM [db_analyze_etc].[dbo].[tbl_remote_site_price_log] as l "
		sqlStr = sqlStr & "		LEFT JOIN [db_analyze_etc].[dbo].[tbl_remote_site_price_Item_Match] as m ON l.sitecode = m.sitecode and l.siteitemcode = m.siteitemcode "
		sqlStr = sqlStr & "		LEFT JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i on m.itemid = i.itemid "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
		'response.write sqlStr &"<Br>"
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
            FTotalCount = rsAnalget("cnt")
            FTotalPage	= rsAnalget("totPg")
        rsAnalget.Close

		if FTotalCount < 1 then exit function
			
    	sqlStr = ""
		sqlStr = sqlStr & "SELECT Top " & Cstr(FPageSize * FCurrPage) & " "
		sqlStr = sqlStr & "		l.sitecode, l.siteitemcode, l.siteitemname, l.realsellCost, l.orgsellCost, l.regTime, isNull(m.itemid,0) as itemid, i.itemname "
		sqlStr = sqlStr & "		, isNull(i.sellcash,0) as sellcash, l.brandname as obrandname, isNull(i.brandname,'') as brandname "
		sqlStr = sqlStr & "FROM [db_analyze_etc].[dbo].[tbl_remote_site_price_log] as l "
		sqlStr = sqlStr & "		LEFT JOIN [db_analyze_etc].[dbo].[tbl_remote_site_price_Item_Match] as m ON l.sitecode = m.sitecode and l.siteitemcode = m.siteitemcode "
		sqlStr = sqlStr & "		LEFT JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i on m.itemid = i.itemid "
		sqlStr = sqlStr & "WHERE 1=1 " & sqlsearch & " "
		sqlStr = sqlStr & "ORDER BY l.regTime DESC, l.siteRank ASC"
		'response.write sqlStr
		rsAnalget.pagesize = FPageSize
		rsAnalget.Open sqlStr,dbAnalget,1
		
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsAnalget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        IF not rsAnalget.EOF THEN
            rsAnalget.absolutepage = FCurrPage
            do until rsAnalget.EOF
                set FItemList(i) = new COSItemItem

					FItemList(i).Fsitecode		= rsAnalget("sitecode")
					FItemList(i).Fsiteitemcode	= rsAnalget("siteitemcode")
					FItemList(i).Fsiteitemname	= rsAnalget("siteitemname")
					FItemList(i).FrealsellCost	= rsAnalget("realsellCost")
					FItemList(i).ForgsellCost	= rsAnalget("orgsellCost")
					FItemList(i).FregTime			= rsAnalget("regTime")
					FItemList(i).Fitemid			= rsAnalget("itemid")
					FItemList(i).Fitemname		= rsAnalget("itemname")
					FItemList(i).Fsellcash		= rsAnalget("sellcash")
					FItemList(i).Fobrandname		= rsAnalget("obrandname")
					FItemList(i).Fbrandname		= rsAnalget("brandname")

                rsAnalget.movenext
                i=i+1
            loop
		End IF
    	rsAnalget.Close

    end Function
    
    
    public Sub sbOtherSiteItem()
    	dim sqlStr, addsql
		sqlStr = "select m.itemid, i.itemname from [db_analyze_etc].[dbo].[tbl_remote_site_price_Item_Match] as m "
		sqlStr = sqlStr & " inner join [db_analyze_data_raw].[dbo].[tbl_item] as i on m.itemid = i.itemid "
		sqlStr = sqlStr & " where m.sitecode = '" & FRectSiteCode & "' and m.siteitemcode = '" & FRectSiteItemID & "' "
		'response.write sqlStr
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
	
		set FOneItem = new COSItemItem
		if Not rsAnalget.Eof then
			FOneItem.Fitemid = rsAnalget("itemid")
			FOneItem.Fitemname = db2html(rsAnalget("itemname"))
		end if
		rsAnalget.Close
		
		sqlStr = "select siteitemname from [db_analyze_etc].[dbo].[tbl_remote_site_price_log] "
		sqlStr = sqlStr & " where sitecode = '" & FRectSiteCode & "' and siteitemcode = '" & FRectSiteItemID & "' and regTime = '" & FRectRegDate & "'"
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
		if Not rsAnalget.Eof then
			FOneItem.Fsiteitemname = rsAnalget("siteitemname")
		end if
		rsAnalget.Close
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

Function fnSiteURL(sc,ic)
	Dim vLink
	SELECT CASE sc
		Case "ohou" : vLink = "http://ohou.se/productions/" & ic & "/selling"
	END SELECT
	fnSiteURL = vLink
End Function

Function fnPercentView(o,s)
	Dim vPercent
	vPercent = CInt(((o-s)/o)*100)
	fnPercentView = vPercent
End Function
%>