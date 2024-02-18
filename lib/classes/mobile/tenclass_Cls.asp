<%
'###############################################
' PageName :tenclass
' Discription : ����Ʈ ���� mdpick
' History : 2014.01.28 ����ȭ ����
'###############################################

Class tenClassItem
	public Fidx
	Public Fmainimage
	Public Fmaincopy
	Public Fsubcopy
	Public Fstartdate
	Public Fenddate
	Public Fadminid
	Public Flastadminid
	public Fisusing
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fadminnotice
	Public Fsortno

	Public Fdidx
	Public Fitemid
	Public Fsortnno
	Public FitemName
	Public FsmallImage
	Public Fxmlregdate

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class tenClass
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
	Public FRectdidx

	'//admin/mobile/tpobanner/tpo_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "SELECT TOP 1 * "
        sqlStr = sqlStr + " FROM db_sitemaster.dbo.tbl_mobile_class "
        sqlStr = sqlStr + " WHERE idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new tenClassItem

        if Not rsget.Eof then
    		FOneItem.Fidx			= rsget("idx")
    		FOneItem.Fmainimage 	= staticImgUrl & "/mobile/tenclass"& rsget("mainimage")
			FOneItem.Fmaincopy		= rsget("maincopy")
			FOneItem.Fsubcopy		= rsget("subcopy")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fadminnotice	= rsget("adminnotice")
        end If

        rsget.Close
    end Sub

	'//subitem
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = "SELECT TOP 1 s.*, i.itemname, i.smallImage "
        sqlStr = sqlStr & "FROM db_sitemaster.dbo.tbl_mobile_class_items as s "
        sqlStr = sqlStr & "	LEFT OUTER JOIN db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		ON s.Itemid=i.itemid "
        sqlStr = sqlStr & "			AND i.itemid<>0 "
        SqlStr = SqlStr + " WHERE didx=" + CStr(FRectdidx)

		'rw SqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new tenClassItem
        if Not rsget.Eof then
            FOneItem.Fdidx				= rsget("didx")
            FOneItem.Fidx				= rsget("idx")
            FOneItem.FItemid			= rsget("Itemid")
            FOneItem.Fsortno			= rsget("sortno")
            FOneItem.Fisusing			= rsget("isusing")
            FOneItem.FitemName			= rsget("itemname")
            FOneItem.FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FOneItem.FItemid) & "/" & rsget("smallImage"),"")
        end if
        rsget.close
	End Sub

	'//admin/mobile/cateimg/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " SELECT count(idx) as cnt FROM db_sitemaster.dbo.tbl_mobile_class "
		sqlStr = sqlStr + " WHERE 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " AND isusing=" & Fisusing
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " AND StartDate >='" & Fsdt & " 00:00:00' AND EndDate <='" & Fsdt & " 23:59:59' "

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

        sqlStr = "SELECT TOP " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr + " * "
        sqlStr = sqlStr + " FROM db_sitemaster.dbo.tbl_mobile_class "
        sqlStr = sqlStr + " WHERE 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " AND isusing=" & Fisusing
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " AND StartDate >='" & Fsdt & " 00:00:00' AND EndDate <='" & Fsdt & " 23:59:59' "

		sqlStr = sqlStr + " ORDER BY idx DESC"

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
				set FItemList(i) = new tenClassItem

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fmaincopy		= rsget("maincopy")
				FItemList(i).Fsubcopy		= rsget("subcopy")
				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")
				FItemList(i).Fadminid		= rsget("adminid")
				FItemList(i).Flastadminid	= rsget("lastadminid")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Flastupdate	= rsget("lastupdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	'//admin/mobile/cateimg/index.asp
    public Sub GetContentsItemList()
       dim sqlStr, addSql, i

		sqlStr = " SELECT count(idx) as cnt FROM db_sitemaster.dbo.tbl_mobile_class_items "
		sqlStr = sqlStr + " WHERE 1=1"
		sqlStr = sqlStr & " AND idx='" & FRectidx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " AND isusing=" & Fisusing
        end If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

        sqlStr = "SELECT TOP " + CStr(FPageSize * FCurrPage) + " s.didx , s.idx , s.itemid , s.isusing as itemusing , s.sortno, isnull(s.itemname,i.itemname) as itemname , i.smallImage "
        sqlStr = sqlStr & "FROM db_sitemaster.dbo.tbl_mobile_class_items as s "
        sqlStr = sqlStr & "	LEFT OUTER JOIN db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		ON s.itemid=i.itemid "
        sqlStr = sqlStr & "			AND i.itemid<>0 "
        sqlStr = sqlStr & " WHERE idx='" & FRectidx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " AND isusing=" & Fisusing
        end If

		sqlStr = sqlStr + " ORDER BY sortno ASC"

'		response.write sqlStr &"<br>"
'		Response.end
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
				set FItemList(i) = new tenClassItem

				FItemList(i).Fdidx				= rsget("didx")
	            FItemList(i).Fidx				= rsget("idx")
	            FItemList(i).Fitemid			= rsget("itemid")
	            FItemList(i).Fsortno			= rsget("sortno")
	            FItemList(i).FIsUsing			= rsget("itemusing")
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

'// STAFF �̸� ����
public Function getStaffUserName(uid)
	if uid="" or isNull(uid) then
		exit Function
	end if

	Dim strSql
	strSql = "SELECT TOP 1 username FROM db_partner.dbo.tbl_user_tenbyten WHERE userid='" & uid & "'"
	rsget.Open strSql, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		getStaffUserName = rsget("username")
	end if
	rsget.Close
End Function
%>