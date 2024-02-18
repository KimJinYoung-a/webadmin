<%
'###############################################
' PageName :onAirCls
' Discription : 사이트 메인 on Air
' History : 2013.12.16 이종화 생성
'###############################################

Class COnAirItem
	public fidx
	Public fgubun
	Public fonairtitle
	Public Fstartdate
	Public Fenddate 
	Public Fadminid
	Public Flastadminid
	public Fisusing 
	Public Fregdate
	Public Fusername
	Public Flastupdate

	Public FsubIdx
	Public Flistidx
	Public Fsortnum
	Public FitemName
	Public FsmallImage
	Public FItemid

	Public Fctitle
	Public Fcper
	Public Fcnum
	Public Fcgubun

	Public Fxmlregdate
	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COnAir
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
	Public FRectSubIdx
	Public FRectlistidx
	
	'//admin/mobile/tpobanner/tpo_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_onair_list "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new COnAirItem
        
        if Not rsget.Eof then
    		FOneItem.fidx				= rsget("idx")
    		FOneItem.fgubun			= rsget("gubun")
    		FOneItem.fonairtitle		= rsget("onairtitle")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fadminid		= rsget("adminid")
			FOneItem.Flastadminid	= rsget("lastadminid")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fctitle			= rsget("ctitle")
			FOneItem.Fcper			= rsget("cper")
			FOneItem.Fcnum			= rsget("cnum")
			FOneItem.Fcgubun		= rsget("cgubun")
        end If
        
        rsget.Close
    end Sub

	'//subitem
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = "Select top 1 s.*, i.itemname, i.smallImage "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_mobile_main_onair_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.Itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        SqlStr = SqlStr + " where subIdx=" + CStr(FRectSubIdx)

		'rw SqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new COnAirItem
        if Not rsget.Eof then
            FOneItem.FsubIdx			= rsget("subIdx")
            FOneItem.Flistidx			= rsget("listIdx")
            FOneItem.FItemid			= rsget("Itemid")
            FOneItem.Fsortnum		= rsget("sortnum")
            FOneItem.Fisusing		= rsget("isusing")
            FOneItem.FitemName	= rsget("itemname")
            FOneItem.FsmallImage	= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FOneItem.FItemid) & "/" & rsget("smallImage"),"")
        end if
        rsget.close
	End Sub
	
	'//admin/mobile/cateimg/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_mobile_main_onair_list "
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >='" & Fsdt & " 00:00:00' and  StartDate <='" & Fsdt & " 23:59:59' "
		'if Fedt<>"" then sqlStr = sqlStr & " and  EndDate <='" & Fedt & " 07:59:59' "

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr + " * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_onair_list "
        sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >='" & Fsdt & " 00:00:00' and  StartDate <='" & Fsdt & " 23:59:59' "
		'if Fedt<>"" then sqlStr = sqlStr & " and  EndDate <='" & Fedt & " 07:59:59' "
        
		sqlStr = sqlStr + " order by  idx desc" 

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
				set FItemList(i) = new COnAirItem
				
				FItemList(i).fidx					= rsget("idx")
				FItemList(i).fgubun			= rsget("gubun")
				FItemList(i).fonairtitle		= rsget("onairtitle")
				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")
				FItemList(i).Fadminid			= rsget("adminid")
				FItemList(i).Flastadminid	= rsget("lastadminid")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fctitle				= rsget("ctitle")
				FItemList(i).Fcper				= rsget("cper")
				FItemList(i).Fcnum			= rsget("cnum")
				FItemList(i).Fcgubun			= rsget("cgubun")
				FItemList(i).Fxmlregdate			= rsget("xmlregdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	'//admin/mobile/cateimg/index.asp
    public Sub GetContentsItemList()
       dim sqlStr, addSql, i

		sqlStr = " select count(listidx) as cnt from db_sitemaster.dbo.tbl_mobile_main_onair_item "
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr & " and  listidx='" & FRectlistidx & "'"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >'" & Fsdt & " 08:00:00' "
		if Fedt<>"" then sqlStr = sqlStr & " and  EndDate <='" & Fedt & " 07:59:59' "

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " s.subidx , s.listidx , s.itemid , s.isusing as itemusing , s.sortnum , i.itemname, i.smallImage "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_mobile_main_onair_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        sqlStr = sqlStr & "Where listidx='" & FRectlistidx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >'" & Fsdt & " 00:00:00' "
		if Fedt<>"" then sqlStr = sqlStr & " and  EndDate <='" & Fedt & " 07:59:59' "
        
		sqlStr = sqlStr + " order by sortnum asc" 

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
				set FItemList(i) = new COnAirItem
				
				FItemList(i).FsubIdx				= rsget("subidx")
	            FItemList(i).Flistidx				= rsget("listidx")
	            FItemList(i).Fitemid				= rsget("itemid")
	            FItemList(i).Fsortnum			= rsget("sortnum")
	            FItemList(i).FIsUsing				= rsget("itemusing")
	            FItemList(i).FitemName			= rsget("itemname")
	            FItemList(i).FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallImage"),"")

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