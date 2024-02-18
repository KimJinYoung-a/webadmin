<%
'###############################################
' PageName :twinitems
' Discription : 단품 배너
' History : 2017-08-03 이종화 생성
'###############################################

Class CMainbannerItem
	public fidx
	Public Fevtstdate
	Public Fevteddate
	Public Fstartdate
	Public Fenddate 
	Public Fadminid
	Public Flastadminid
	public Fisusing 
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fordertext
	Public Fsortnum
	Public Ftodaybanner
	Public Fevt_code
	Public FL_img
	public FL_maincopy
	public FL_itemname
	public FL_itemid
	public FL_newbest
	public FR_img
	public FR_maincopy
	public FR_itemname
	public FR_itemid
	public FR_newbest

	Public Fiteminfo
	
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
	Public FRectvaliddate
       
    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	Public FRectsedatechk
	
	'//admin/appmanage/today/enjoyevent/enjoy_insert.asp
    public Sub GetOneContents()
        dim sqlStr

        sqlStr = "select top 1 *"
		sqlStr = sqlStr & " , STUFF((   "
        sqlStr = sqlStr & " SELECT ',' + cast(i.itemid as varchar(120)) +'|'+ cast(i.itemname as varchar(120)) +'|'+ cast(i.icon1image as varchar(50))"
        sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock)"
        sqlStr = sqlStr & " WHERE i.itemid in (L_itemid , R_itemid) and i.itemid<>0"
        sqlStr = sqlStr & " FOR XML PATH('')"
        sqlStr = sqlStr & " ), 1, 1, '') AS iteminfo"
        sqlStr = sqlStr & " from db_sitemaster.[dbo].[tbl_mobile_main_twinitems] with (nolock)"
        sqlStr = sqlStr & " where idx=" & CStr(FRectIdx)

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainbannerItem
        
        if Not rsget.Eof then
    		FOneItem.fidx				= rsget("idx")
			FOneItem.Fstartdate			= rsget("startdate")
			FOneItem.Fenddate			= rsget("enddate")
			FOneItem.Fadminid			= rsget("adminid")
			FOneItem.Flastadminid		= rsget("lastadminid")
			FOneItem.Fisusing			= rsget("isusing")
			FOneItem.Fordertext			= rsget("ordertext")
			FOneItem.FL_img				= staticImgUrl & "/mobile/twinitems" & rsget("L_img")
			FOneItem.FL_maincopy		= rsget("L_maincopy")
			FOneItem.FL_itemname		= rsget("L_itemname")
			FOneItem.FL_itemid			= rsget("L_itemid")
			FOneItem.FL_newbest			= rsget("L_newbest")
			FOneItem.FR_img				= staticImgUrl & "/mobile/twinitems" & rsget("R_img")
			FOneItem.FR_maincopy		= rsget("R_maincopy")
			FOneItem.FR_itemname		= rsget("R_itemname")
			FOneItem.FR_itemid			= rsget("R_itemid")
			FOneItem.FR_newbest			= rsget("R_newbest")
			FOneItem.Fiteminfo			= rsget("iteminfo") '2017-07-27 addtype
        end If
        
        rsget.Close
    end Sub
	
	'//admin/appmanage/today/enjoyevent/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.[dbo].[tbl_mobile_main_twinitems] with (nolock)"
		sqlStr = sqlStr & " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr & " and isusing='" & CStr(Fisusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and startdate = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then 
			sqlStr = sqlStr & " and '" & Fsdt & "' between startdate and enddate "
		End If 

		If FRectvaliddate = "on" Then 
			sqlStr = sqlStr & " and enddate > getdate() "
		End If 

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " & CStr(FPageSize * FCurrPage) & " * "
		sqlStr = sqlStr & " , STUFF((   "
        sqlStr = sqlStr & " SELECT '^^' + cast(i.itemid as varchar(120)) +'|'+ cast(i.itemname as varchar(120)) +'|'+ cast(i.icon1image as varchar(50))"
        sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock)"
        sqlStr = sqlStr & " WHERE i.itemid in (L_itemid , R_itemid) and i.itemid<>0"
        sqlStr = sqlStr & " FOR XML PATH('')"
        sqlStr = sqlStr & " ), 1, 1, '') AS iteminfo"
		sqlStr = sqlStr & " from db_sitemaster.[dbo].[tbl_mobile_main_twinitems] as t with (nolock)"
        sqlStr = sqlStr & " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr & " and isusing='" & CStr(Fisusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and startdate = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then 
			sqlStr = sqlStr & " and '" & Fsdt & "' between startdate and enddate "
		End If 

		If FRectvaliddate = "on" Then 
			sqlStr = sqlStr & " and t.enddate > getdate() "
		End If 

		sqlStr = sqlStr & " order by t.startdate desc" 

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				FItemList(i).Fstartdate			= rsget("startdate")
				FItemList(i).Fenddate			= rsget("enddate")
				FItemList(i).Fadminid			= rsget("adminid")
				FItemList(i).Flastadminid		= rsget("lastadminid")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")

				FItemList(i).FL_itemname		= rsget("L_itemname")
				FItemList(i).FR_itemname		= rsget("R_itemname")

				FItemList(i).FL_img				= staticImgUrl & "/mobile/twinitems" & rsget("L_img")
				FItemList(i).FR_img				= staticImgUrl & "/mobile/twinitems" & rsget("R_img")

				FItemList(i).FL_itemid			= rsget("L_itemid")
				FItemList(i).FR_itemid			= rsget("R_itemid")

				FItemList(i).Fiteminfo			= rsget("iteminfo")

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