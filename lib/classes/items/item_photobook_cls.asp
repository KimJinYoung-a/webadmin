<%
Class CItemPhotoDetail
    public Fitemid
    public Fitemname
    public Fitemoption
    public Ftplcode
    public Fpcode
    public Ftplname

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
	
end Class


Class CItemPhoto
	public FItemList()
    
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectColorCD
	public FRectItemId

	public function GetPhotoItemList()
        dim sqlStr, addSql, i

		'// 결과수 카운트
		sqlStr = "select Count(A.itemid), CEILING(CAST(Count(A.itemid) AS FLOAT)/" & FPageSize & ") "
        sqlStr = sqlStr & " from (select itemid from [db_item].[dbo].[tbl_fuji_templete_code] group by itemid) AS A "
        sqlStr = sqlStr & " where 1=1 " & addSql

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FTotalPage = rsget(1)
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " p.itemid, i.itemname "
        sqlStr = sqlStr & " from [db_item].[dbo].[tbl_fuji_templete_code] as p "
        sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item] as i on p.itemid = i.itemid "
        sqlStr = sqlStr & " where 1 = 1 " & addSql
        sqlStr = sqlStr & " group by p.itemid, i.itemname "
		sqlStr = sqlStr & " Order by p.itemid desc "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
        redim preserve FItemList(FResultCount)

        i=0
        if Not(rsget.EOF or rsget.BOF) then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemPhotoDetail

                FItemList(i).Fitemid	= rsget("itemid")
                FItemList(i).Fitemname	= rsget("itemname")
                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
    
    
	public function GetPhotoTempleteList()
        dim sqlStr, addSql, i

		addSql = " and p.itemid = '" & FRectItemId & "' "

        '// 본문 내용 접수
        sqlStr = "select "
        sqlStr = sqlStr & " p.itemid, p.itemoption, p.tplcode, p.pcode, p.tplname "
        sqlStr = sqlStr & " from [db_item].[dbo].[tbl_fuji_templete_code] as p "
        sqlStr = sqlStr & " where 1 = 1 " & addSql
		sqlStr = sqlStr & " Order by itemoption asc "

        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-1
        if (FResultCount<1) then FResultCount=0
        
        redim preserve FItemList(FResultCount)

        i=0
        if Not(rsget.EOF or rsget.BOF) then
            do until rsget.EOF
                set FItemList(i) = new CItemPhotoDetail

                FItemList(i).Fitemid		= rsget("itemid")
                FItemList(i).Fitemoption	= rsget("itemoption")
                FItemList(i).Ftplcode		= rsget("tplcode")
                FItemList(i).Fpcode			= rsget("pcode")
                FItemList(i).Ftplname		= rsget("tplname")
                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
    
    
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
    End Sub
    
end Class
%>