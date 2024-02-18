<%
'###############################################
' PageName : main_noticebanner
' Discription : 사이트 메인 공지 배너 관리
' History : 2013.04.02
'###############################################
Class CMainbannerItem
	public Fiidx 
	public FutArr 
	public FutnArr
	public Fstartday 
	public Fendday 
	public Ftitle
	public Fsorting 
	public Ftext 
	public Ftexturl 
	public Fisusing 
	public Fwriter 
	public Flastwriter
	Public Fwritedate
	Public Flastupdate
	Public Ftextcopy

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
    public FSearchSdate
    public FSearchEdate
    public Fisusing
    public Fuserlevel
	    
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_noticebanner "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
        
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainbannerItem
        
        if Not rsget.Eof then
    
			FOneItem.Fiidx				= rsget("idx")
			FOneItem.FutArr			= rsget("usertype")
			FOneItem.FutnArr			= rsget("usertypename")
			FOneItem.Fstartday		= rsget("startdate")
			FOneItem.Fendday		= rsget("enddate")
			FOneItem.Ftitle				= rsget("title")
			FOneItem.Fsorting		= rsget("sorting")
			FOneItem.Ftext				= rsget("textcontents")
			FOneItem.Ftexturl			= rsget("texturl")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fwriter			= rsget("writer")
			FOneItem.Fwritedate	= rsget("writedate")
			FOneItem.Flastwriter		= rsget("lastwriter")
			FOneItem.Ftextcopy		= rsget("textcopy")
        end If
        
        rsget.Close
    end Sub

    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_mobile_noticebanner "
		sqlStr = sqlStr + " where 1=1"
        
        if FSearchSdate <> "" then
            sqlStr = sqlStr + " and convert(varchar(10),enddate,120) >='" + CStr(FSearchSdate) + "'"
        end If
        
		if FSearchSdate <> "" then
            sqlStr = sqlStr + " and convert(varchar(10),enddate,120) <= '" + CStr(FSearchEdate) + "'"
        end if

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end if
        
        if Fuserlevel <>"" then
            sqlStr = sqlStr + " and usertype like '%" + CStr(Fuserlevel) + "%'"
        end If

        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_noticebanner "
        sqlStr = sqlStr + " where 1=1"

        if FSearchSdate <> "" then
            sqlStr = sqlStr + " and datediff(d,enddate,'" + CStr(FSearchSdate) + "')<=0"
        end If
        
		if FSearchEdate <> "" then
            sqlStr = sqlStr + " and datediff(d,startdate,'" + CStr(FSearchEdate) + "')>=0"
        end if
        
        if Fisusing<>"" then
        	if Fisusing="Y" then Fisusing="1"
        	if Fisusing="N" then Fisusing="0"
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end if
        
        if Fuserlevel <>"" then
            sqlStr = sqlStr + " and usertype like '%" + CStr(Fuserlevel) + "%'"
        end If

		sqlStr = sqlStr + " order by sorting desc "
        
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

				FItemList(i).Fiidx			= rsget("idx")
				FItemList(i).FutArr			= rsget("usertype")
				FItemList(i).FutnArr		= rsget("usertypename")
				FItemList(i).Fstartday		= rsget("startdate")
				FItemList(i).Fendday		= rsget("enddate")
				FItemList(i).Ftitle			= rsget("title")
				FItemList(i).Fsorting		= rsget("sorting")
				FItemList(i).Ftext			= rsget("textcontents")
				FItemList(i).Ftexturl		= rsget("texturl")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fwriter		= rsget("writer")
				FItemList(i).Fwritedate	= rsget("writedate")
				FItemList(i).Flastwriter	= rsget("lastwriter")

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