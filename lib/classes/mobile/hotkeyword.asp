<%
'###############################################
' PageName :hotkeyword
' Discription : today hot-keyword
' History : 2014-09-15 이종화 생성
'###############################################

Class CMainbannerItem
	public fidx
	Public Fkwimg
	Public Fkword
	Public Fktitle
	Public Fkcontents
	Public Fkurl_mo
	Public Fkurl_app
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

	Public Fitemid
	Public FitemName
	Public FsmallImage
	Public Fbasicimage
	
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
	Public FRectsedatechk
	
	'//admin/mobile/hotkeyword/hkw_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 hk.* , i.itemname ,  i.smallimage "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_today_hotkeyword as hk "
        sqlStr = sqlStr + " left outer join db_item.dbo.tbl_item as i on hk.itemid = i.itemid and i.itemid<>0 "
        sqlStr = sqlStr + " where hk.idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainbannerItem
        
        if Not rsget.Eof then
    		FOneItem.fidx			= rsget("idx")
			FOneItem.Fkwimg			= staticImgUrl & "/mobile/hotkeyword" & rsget("kwimg")
			FOneItem.Fkword			= rsget("kword")
			FOneItem.Fktitle		= rsget("ktitle")
			FOneItem.Fkcontents		= rsget("kcontents")
			FOneItem.Fkurl_mo		= rsget("kurl_mo")
			FOneItem.Fkurl_app		= rsget("kurl_app")
			FOneItem.Fappdiv		= rsget("appdiv")
			FOneItem.Fappcate		= rsget("appcate")

			FOneItem.Fsortnum		= rsget("sortnum")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fadminid		= rsget("adminid")
			FOneItem.Flastadminid	= rsget("lastadminid")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fordertext		= rsget("ordertext")

			FOneItem.Fitemid		= rsget("itemid")
			FOneItem.FitemName		= rsget("itemname")
            FOneItem.FsmallImage	= chkIIF(Not(rsget("smallimage")="" or isNull(rsget("smallimage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FOneItem.FItemid) & "/" & rsget("smallImage"),"")
        end If
        
        rsget.Close
    end Sub
	
	'//admin/mobile/cateimg/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_mobile_main_today_hotkeyword "
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr + " and startdate = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then 
			sqlStr = sqlStr + " and '" & Fsdt & "' between startdate and enddate "
		End If 

		'response.write sqlStr &"<br>"
		'response.End
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr + " t.idx , t.kwimg , t.kword , t.ktitle , t.kcontents , t.kurl_mo , t.kurl_app , t.appdiv , t.appcate  , t.sortnum , t.startdate , t.enddate , t.adminid , t.lastadminid , t.isusing ,t.regdate , t.lastupdate , t.xmlregdate  , t.itemid , i.basicimage "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_today_hotkeyword as t "
        sqlStr = sqlStr + " left outer join db_item.dbo.tbl_item as i on t.itemid = i.itemid and i.itemid<>0 "
        sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and t.isusing='" + CStr(Fisusing) + "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr + " and t.startdate = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then 
			sqlStr = sqlStr + " and '" & Fsdt & "' between t.startdate and t.enddate "
		End If 
        
		sqlStr = sqlStr + " order by  t.sortnum asc ,  t.idx desc" 

		'response.write sqlStr &"<br>"
		'response.End
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
				FItemList(i).Fkwimg				= staticImgUrl & "/mobile/hotkeyword" & rsget("kwimg")
				FItemList(i).Fkword				= rsget("kword")
				FItemList(i).Fktitle			= rsget("ktitle")
				FItemList(i).Fkcontents			= rsget("kcontents")
				FItemList(i).Fkurl_mo			= rsget("kurl_mo")
				FItemList(i).Fkurl_app			= rsget("kurl_app")
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
				FItemList(i).Fbasicimage		= webImgUrl & "/image/basic/" + GetImageSubFolderByItemid(db2Html(rsget("itemid"))) + "/" + db2Html(rsget("basicimage"))

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
%>