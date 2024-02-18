<%
'###############################################
' PageName :catebanner
' Discription : 사이트 메인 공지 배너 관리
' History : 2013.12.12 이종화 생성
'			2013.12.15 한용민 수정
'			2014.02.04 이종화 추가
'###############################################

Class CMainbannerItem
	public Fcatecode 
	public Fisusing 
	Public Fcateimg
	Public Fcatename
	public fidx
	Public Fkword1
	Public Fkword2
	Public Fkword3
	Public Fkwordurl1
	Public Fkwordurl2
	Public Fkwordurl3

	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CMainbanner
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
	
	'//admin/mobile/cateimg/ci_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_cateimg "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainbannerItem
        
        if Not rsget.Eof then
    		FOneItem.fidx		= rsget("idx")
			FOneItem.Fcatecode		= rsget("catecode")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fcateimg		= staticImgUrl & "/mobile/catecode" & rsget("cateimage")
			FOneItem.Fkword1		= rsget("kword1")
			FOneItem.Fkword2		= rsget("kword2")
			FOneItem.Fkword3		= rsget("kword3")
			FOneItem.Fkwordurl1		= rsget("kwordurl1")
			FOneItem.Fkwordurl2		= rsget("kwordurl2")
			FOneItem.Fkwordurl3		= rsget("kwordurl3")
        end If
        
        rsget.Close
    end Sub
	
	'//admin/mobile/cateimg/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(c.catecode) as cnt from db_sitemaster.dbo.tbl_mobile_cateimg as c inner join db_item.dbo.tbl_display_cate as d on c.catecode = d.catecode"
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
        sqlStr = sqlStr + " , c.kword1 , c.kword2 , c.kword3 , c.kwordurl1 , c.kwordurl2 , c.kwordurl3  "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_cateimg as c "
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
				set FItemList(i) = new CMainbannerItem
				
				FItemList(i).fidx	= rsget("idx")
				FItemList(i).Fcatecode	= rsget("catecode")
				FItemList(i).Fcatename	= rsget("catename")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fcateimg		= staticImgUrl & "/mobile/catecode" & rsget("cateimage")
				FItemList(i).Fkword1		= rsget("kword1")
				FItemList(i).Fkword2		= rsget("kword2")
				FItemList(i).Fkword3		= rsget("kword3")
				FItemList(i).Fkwordurl1		= rsget("kwordurl1")
				FItemList(i).Fkwordurl2		= rsget("kwordurl2")
				FItemList(i).Fkwordurl3		= rsget("kwordurl3")

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