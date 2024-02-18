<%
'###############################################
' PageName :브랜드 배너
' Discription : 투데이 브랜드배너 영역
' History : 2017-08-03 이종화 생성
'###############################################

Class CMainbannerItem
	public fidx
	Public Flinkurl
	Public Fstartdate
	Public Fenddate 
	Public Fadminid
	Public Flastadminid
	public Fisusing 
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fordertext

	Public Fitemid1
	Public Fitemid2
	Public Fitemid3
	Public Fitemid4

	Public Fiteminfo

	Public Fitemimg1		
	Public Fitemimg2
	Public Fitemimg3
	Public Fitemimg4

	Public Fver_no
	Public Fkeyword
	Public Fpicknum

	Public Fbgcolor
	
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
	Public FRecttype
	
	'//admin/mobile/todaybrand/brandinfo_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 t.* "
        sqlStr = sqlStr & " , STUFF(( "
        sqlStr = sqlStr & " SELECT '^^' + cast(i.itemid as varchar(120)) +'|'+ cast(i.itemname as varchar(120)) +'|'+ cast(i.smallimage as varchar(50)) "
        sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & " WHERE i.itemid in (t.itemid1 , t.itemid2 , t.itemid3 , t.itemid4) and i.itemid<>0 "
        sqlStr = sqlStr & " FOR XML PATH('') "
        sqlStr = sqlStr & " ), 1, 1, '') AS iteminfo "
        sqlStr = sqlStr & " from db_sitemaster.[dbo].[tbl_mobile_main_keyword] as t "
        sqlStr = sqlStr & " where idx=" & CStr(FRectIdx)

'		rw sqlStr & "<Br>"
'		Response.end
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainbannerItem
        
        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
			FOneItem.Fstartdate			= rsget("startdate")
			FOneItem.Fenddate			= rsget("enddate")
			FOneItem.Fisusing			= rsget("isusing")
			FOneItem.Fordertext			= rsget("ordertext")
			FOneItem.Flinkurl			= rsget("linkurl")

			FOneItem.Fver_no			= rsget("ver_no")
			FOneItem.Fkeyword			= rsget("keyword")
			FOneItem.Fpicknum			= rsget("picknum")

			FOneItem.Fitemid1			= rsget("itemid1") '2017-07-27 itemid1
			FOneItem.Fitemid2			= rsget("itemid2") '2017-07-27 itemid2
			FOneItem.Fitemid3			= rsget("itemid3") '2017-07-27 itemid2
			FOneItem.Fitemid4			= rsget("itemid4") '2017-07-27 itemid2

			If rsget("itemimg1") <> "" then
			FOneItem.Fitemimg1			= staticImgUrl & "/mobile/todaykeyword" & rsget("itemimg1")
			Else
			FOneItem.Fitemimg1			= ""
			End If

			If rsget("itemimg2") <> "" then
			FOneItem.Fitemimg2			= staticImgUrl & "/mobile/todaykeyword" & rsget("itemimg2")
			Else
			FOneItem.Fitemimg2			= ""
			End If 

			If rsget("itemimg3") <> "" then
			FOneItem.Fitemimg3			= staticImgUrl & "/mobile/todaykeyword" & rsget("itemimg3")
			Else
			FOneItem.Fitemimg3			= ""
			End If 

			If rsget("itemimg4") <> "" then
			FOneItem.Fitemimg4			= staticImgUrl & "/mobile/todaykeyword" & rsget("itemimg4")
			Else
			FOneItem.Fitemimg4			= ""
			End If 

			FOneItem.Fiteminfo			= rsget("iteminfo") '2017-07-27 addtype
			FOneItem.Fbgcolor			= rsget("bgcolor") 
        end If
        
        rsget.Close
    end Sub
	
	'//admin/mobile/todaybrand/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.[dbo].[tbl_mobile_main_keyword] "
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
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " & CStr(FPageSize * FCurrPage) & " * "
		sqlStr = sqlStr & " , STUFF(( "
        sqlStr = sqlStr & " SELECT '^^' + cast(i.itemid as varchar(120)) +'|'+ cast(i.itemname as varchar(120)) +'|'+ cast(i.smallimage as varchar(50)) "
        sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & " WHERE i.itemid in (t.itemid1 , t.itemid2 , t.itemid3 , t.itemid4) and i.itemid<>0 "
        sqlStr = sqlStr & " FOR XML PATH('') "
        sqlStr = sqlStr & " ), 1, 1, '') AS iteminfo "
		sqlStr = sqlStr & " from db_sitemaster.[dbo].[tbl_mobile_main_keyword] as t "
        sqlStr = sqlStr & " where 1=1"

		'Response.write sqlStr

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

		sqlStr = sqlStr & " order by startdate asc" 

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
				FItemList(i).Fitemid1			= rsget("itemid1")
				FItemList(i).Fitemid2			= rsget("itemid2")
				FItemList(i).Fitemid3			= rsget("itemid3")
				FItemList(i).Fitemid4			= rsget("itemid4")
				FItemList(i).Fstartdate			= rsget("startdate")
				FItemList(i).Fenddate			= rsget("enddate")
				FItemList(i).Fadminid			= rsget("adminid")
				FItemList(i).Flastadminid		= rsget("lastadminid")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fpicknum			= rsget("picknum")
				FItemList(i).Fkeyword			= rsget("keyword")

				If rsget("itemimg1") <> "" then
				FItemList(i).Fitemimg1			= staticImgUrl & "/mobile/todaykeyword" & rsget("itemimg1")
				Else
				FItemList(i).Fitemimg1			= ""
				End If

				If rsget("itemimg2") <> "" then
				FItemList(i).Fitemimg2			= staticImgUrl & "/mobile/todaykeyword" & rsget("itemimg2")
				Else
				FItemList(i).Fitemimg2			= ""
				End If 

				If rsget("itemimg3") <> "" then
				FItemList(i).Fitemimg3			= staticImgUrl & "/mobile/todaykeyword" & rsget("itemimg3")
				Else
				FItemList(i).Fitemimg3			= ""
				End If 

				If rsget("itemimg4") <> "" then
				FItemList(i).Fitemimg4			= staticImgUrl & "/mobile/todaykeyword" & rsget("itemimg4")
				Else
				FItemList(i).Fitemimg4			= ""
				End If 

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