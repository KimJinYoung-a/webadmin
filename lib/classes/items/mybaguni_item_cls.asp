<%
Class CMyBaguniItem
	public Fitemid
	public Fitemname
	public Fsellcash
	public Flistimage
	public Fitemcount

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMyBaguni
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectSDate
	public FRectEDate
	

	public function fnGetMyBaguniItemList()
    	dim sqlStr
		sqlStr = "select count(a.itemid) as cnt, CEILING(CAST(Count(a.itemid) AS FLOAT)/20) AS totPg from "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "	select  b.itemid, i.itemname, i.sellcash, i.listimage, count(b.itemid) as cnt  "
		sqlStr = sqlStr & "	from [db_my10x10].[dbo].[tbl_my_baguni] as b  "
		sqlStr = sqlStr & "	inner join [db_item].[dbo].[tbl_item] as i on b.itemid = i.itemid  "
		sqlStr = sqlStr & "	where b.userkey <> '' and b.regdate between '" & FRectSDate & " 00:00:00' and '" & FRectEDate & " 23:59:59' "
		sqlStr = sqlStr & "	group by b.itemid, i.itemname, i.sellcash, i.listimage "
		sqlStr = sqlStr & ") as a "

		'response.write sqlStr &"<Br>"
		'response.end
        db3_rsget.Open sqlStr,db3_dbget,1
            FTotalCount = db3_rsget("cnt")
            FTotalPage	= db3_rsget("totPg")
        db3_rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " b.itemid, i.itemname, i.sellcash, i.listimage, count(b.itemid) as cnt "
        sqlStr = sqlStr & "from [db_my10x10].[dbo].[tbl_my_baguni] as b "
        sqlStr = sqlStr & "inner join [db_item].[dbo].[tbl_item] as i on b.itemid = i.itemid "
		sqlStr = sqlStr & "where b.userkey <> '' and b.regdate between '" & FRectSDate & " 00:00:00' and '" & FRectEDate & " 23:59:59' "
		sqlStr = sqlStr & "group by b.itemid, i.itemname, i.sellcash, i.listimage "
		sqlStr = sqlStr & "order by cnt desc"

		'response.write sqlStr &"<Br>"
		'db3_dbget.close
		'response.end
        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr,db3_dbget,1

        FtotalPage =  CDbl(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not db3_rsget.EOF  then
            db3_rsget.absolutepage = FCurrPage
            do until db3_rsget.EOF
                set FItemList(i) = new CMyBaguniItem

					FItemList(i).Fitemid			= db3_rsget("itemid")
					FItemList(i).Fitemname 		= db3_rsget("itemname")
					FItemList(i).Fsellcash 		= db3_rsget("sellcash")
					FItemList(i).Flistimage 		= "http://webimage.10x10.co.kr/image/List/" & GetImageSubFolderByItemid(db3_rsget("itemid")) & "/" & db3_rsget("listimage")
					FItemList(i).Fitemcount		= db3_rsget("cnt")

                db3_rsget.movenext
                i=i+1
            loop
        end if
        db3_rsget.Close
	end Function
	

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