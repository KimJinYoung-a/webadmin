<%
'###############################################
' PageName :catetag
' Discription : 카테고리 메인 인기 태그
' History : 2014-09-21 이종화 생성
'###############################################

Class CMaincatetagItem
	public Fcatecode 
	public Fisusing 
	Public Fcatename
	public fidx
	Public Fkword1
	Public Fkword2
	Public Fkword3
	Public Fkwordurl1
	Public Fkwordurl2
	Public Fkwordurl3
	Public Fappdiv
	Public Fappcate

	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CMaincatetag
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
	Public Fcatecode
	
	'//admin/mobile/catetag/mc_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_catetag "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMaincatetagItem
        
        if Not rsget.Eof then
    		FOneItem.fidx		= rsget("idx")
			FOneItem.Fcatecode		= rsget("catecode")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fkword1		= rsget("kword1")
			FOneItem.Fkwordurl1		= rsget("kwordurl1")
			FOneItem.Fkwordurl2		= rsget("kwordurl2")
			FOneItem.Fappdiv		= rsget("appdiv")
			FOneItem.Fappcate		= rsget("appcate")

        end If
        
        rsget.Close
    end Sub
	
	'//admin/mobile/catetag/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(c.catecode) as cnt from db_sitemaster.dbo.tbl_mobile_catetag as c inner join db_item.dbo.tbl_display_cate as d on c.catecode = d.catecode"
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and c.isusing='" + CStr(Fisusing) + "'"
        end If

	    if Fcatecode <>"" then
            sqlStr = sqlStr + " and c.catecode='" + CStr(Fcatecode) + "'"
        end If
		
		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.* , d.catename  "
        sqlStr = sqlStr + " , c.kword1 , c.kwordurl1 , c.kwordurl2 "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_catetag as c "
        sqlStr = sqlStr + " inner join db_item.dbo.tbl_display_cate as d on c.catecode = d.catecode "
        sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and c.isusing='" + CStr(Fisusing) + "'"
        end If
        
	    if Fcatecode <>"" then
            sqlStr = sqlStr + " and c.catecode='" + CStr(Fcatecode) + "'"
        end If
        
		sqlStr = sqlStr + " order by c.idx desc"

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
				set FItemList(i) = new CMaincatetagItem
				
				FItemList(i).fidx	= rsget("idx")
				FItemList(i).Fcatecode	= rsget("catecode")
				FItemList(i).Fcatename	= rsget("catename")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fkword1		= rsget("kword1")
				FItemList(i).Fkwordurl1		= rsget("kwordurl1")
				FItemList(i).Fkwordurl2		= rsget("kwordurl2")
				FItemList(i).Fappdiv		= rsget("appdiv")
				FItemList(i).Fappcate		= rsget("appcate")

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