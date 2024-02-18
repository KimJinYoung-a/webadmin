<%
Class GoodNoteDiaryContentsCls
	public Fidx
	public Ftitle
	public Fstart_date
	public Fend_date
	public Fisusing
	public Fregdate
End Class

Class GoodNoteDiaryCls

	Public FItemList()
	Public FItem
	public FResultCount
	public FPageSize
	public FCurrPage
	public Ftotalcount
	public FScrollCount
	public FTotalpage
	public FPageCount
	public FRectIsusing
    public FRectSelDate
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	'// brand list
	public Function getStickerList()
        dim sqlStr, addSql, i

        if FRectIsusing<>"" then
            addSql = addSql & " and isusing=" & CStr(FRectIsusing) & ""
        end if

        if FRectSelDate<>"" then
            addSql = addSql & " and '" & FRectSelDate & "' between convert(varchar(10),a.start_date,120) and convert(varchar(10),a.end_date,120) "
        end if

        sqlStr = " select count(idx) as cnt from db_event.dbo.tbl_goodnote WITH(NOLOCK)"
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

       	sqlStr = "select top " & FPageSize * FCurrPage & " idx , title , start_date, end_date, isusing, regdate"
        sqlStr = sqlStr & " from db_event.dbo.tbl_goodnote WITH(NOLOCK)"
		sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & addSql
   		sqlStr = sqlStr & " order by idx desc"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
        rsget.absolutepage = FCurrPage
		if not rsget.EOF  then
		    getStickerList = rsget.getRows()
		end if
		rsget.close
    end Function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class

Class GoodNoteDiaryStickerCls
	public FRectIDX
	public Ftitle
    public Ffont_color
    public Ftitle_image
    public Fbg_color
    public Fcontents_image
    public Fcontents_title
    public Fcontents
    public Fbrand_button_color
    public Fbrand_url
    public Ffile_url
    public Fstart_date
    public Fend_date
    public Fisusing

	public Function fnGetStickerContents
		Dim strSql
		strSql = "SELECT top 1 title,font_color,title_image,bg_color,contents_image,contents_title,contents" & vbcrlf
        strSql = strSql & ",brand_button_color,brand_url,file_url,start_date,end_date,isusing" & vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_goodnote WITH(NOLOCK)" & vbcrlf
		strSql = strSql & " WHERE idx = " & FRectIDX
		rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				Ftitle			    = rsget("title")
				Ffont_color	    	= rsget("font_color")
				Ftitle_image 		= rsget("title_image")
				Fbg_color	    	= rsget("bg_color")
			 	Fcontents_image 	= rsget("contents_image")
			 	Fcontents_title 	= rsget("contents_title")
				Fcontents 		    = rsget("contents")
			  	Fbrand_button_color = rsget("brand_button_color")
			  	Fbrand_url 	        = rsget("brand_url")
			  	Ffile_url		    = rsget("file_url")
                Fstart_date		    = rsget("start_date")
                Fend_date		    = rsget("end_date")
                Fisusing		    = rsget("isusing")
			END IF
		rsget.Close
	End Function
End Class
%>