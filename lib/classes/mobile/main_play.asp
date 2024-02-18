<%
'###############################################
' PageName :play
' Discription : today play
' History : 2015-05-13 이종화 생성
'###############################################

Class CMainbannerItem
	public fidx
	Public Fpimg
	Public Fgubun
	Public Ftitle
	Public Fsubtitle
	Public Furl_mo
	Public Furl_app
	Public Fappdiv
	Public fappcate
	Public Fsortnum
	Public Fstartdate
	Public Fenddate 
	Public Fadminid
	Public Flastadminid
	public Fisusing 
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fordertext
	Public Fxmlregdate

	
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
	Public Fsdt
	Public Fedt
	
	'//admin/mobile/play/play_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_today_play "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainbannerItem
        
        if Not rsget.Eof then
    		FOneItem.fidx			= rsget("idx")
			FOneItem.Fpimg			= staticImgUrl & "/mobile/play" & rsget("pimg")
			FOneItem.Fgubun			= rsget("gubun")
			FOneItem.Ftitle			= rsget("title")
			FOneItem.Fsubtitle		= rsget("subtitle")
			FOneItem.Furl_mo		= rsget("url_mo")
			FOneItem.Furl_app		= rsget("url_app")
			FOneItem.Fappdiv		= rsget("appdiv")
			FOneItem.Fappcate		= rsget("appcate")

			FOneItem.Fsortnum		= rsget("sortnum")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fadminid		= rsget("adminid")
			FOneItem.Flastadminid	= rsget("lastadminid")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fordertext		= rsget("ordertext")
        end If
        
        rsget.Close
    end Sub
	
	'//admin/mobile/cateimg/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_mobile_main_today_play "
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then 
			sqlStr = sqlStr & " and StartDate>='" & Fsdt & " 00:00:00' and StartDate<='" & Fsdt & " 23:59:59' "
		Else 
			sqlStr = sqlStr & " and enddate>getdate() "
		End If 

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr + " t.idx , t.pimg , t.gubun , t.title , t.subtitle , t.url_mo , t.url_app , t.appdiv , t.appcate  , t.sortnum , t.startdate , t.enddate , t.adminid , t.lastadminid , t.isusing ,t.regdate , t.lastupdate , t.xmlregdate "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_today_play as t "
        sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then 
			sqlStr = sqlStr & " and StartDate>='" & Fsdt & " 00:00:00' and StartDate<='" & Fsdt & " 23:59:59' "
		Else
			sqlStr = sqlStr & " and enddate>getdate() "
		End If 
        
		sqlStr = sqlStr + " order by  t.sortnum asc ,  t.idx desc" 

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
				
				FItemList(i).fidx				= rsget("idx")
				FItemList(i).Fpimg				= staticImgUrl & "/mobile/play" & rsget("pimg")
				FItemList(i).Fgubun				= rsget("gubun")
				FItemList(i).Ftitle				= rsget("title")
				FItemList(i).Fsubtitle			= rsget("subtitle")
				FItemList(i).Furl_mo			= rsget("url_mo")
				FItemList(i).Furl_app			= rsget("url_app")
				FItemList(i).Fappdiv			= rsget("appdiv")
				FItemList(i).Fappcate			= rsget("appcate")
				FItemList(i).Fsortnum			= rsget("sortnum")
				FItemList(i).Fstartdate			= rsget("startdate")
				FItemList(i).Fenddate			= rsget("enddate")
				FItemList(i).Fadminid			= rsget("adminid")
				FItemList(i).Flastadminid		= rsget("lastadminid")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fxmlregdate		= rsget("xmlregdate")

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

'// STAFF 이름 접수
public Function getStaffUserName(uid)
	if uid="" or isNull(uid) then
		exit Function
	end if

	Dim strSql
	strSql = "Select top 1 username From db_partner.dbo.tbl_user_tenbyten Where userid='" & uid & "'"
	rsget.Open strSql, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		getStaffUserName = rsget("username")
	end if
	rsget.Close
End Function

'Public Function getGubun(v)
'	
'
'End function
%>