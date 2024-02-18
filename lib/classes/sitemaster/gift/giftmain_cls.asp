<%
class Cgift_item
	public FthemeIdx
	public Fsubject
	public Fregdate
	public Fstartdate
	public Fenddate
	public Fsortno
	public Fidx
	public Fisusing
	public Freguser


    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Cgift_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	Public FRectIsusing
	Public FRectIsOpen
	Public FRectIdx


	'/admin/sitemaster/gift/day/gift.asp
	Public Sub sbHotIssueList
		Dim sqlStr, i, sqladd

		If FRectIsusing <> "" Then
			sqladd = sqladd & " and h.isUsing = '" & FRectIsusing & "' "
		End If

		sqlStr = "SELECT count(*) as cnt"
		sqlStr = sqlStr & " FROM db_board.dbo.tbl_giftmain_hotissue as h with (nolock)"
		sqlStr = sqlStr & " inner join db_board.dbo.tbl_giftShop_theme as t with (nolock) on h.themeIdx = t.themeIdx"
		sqlStr = sqlStr & " where 1=1 " & sqladd

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		If FTotalCount < 1 Then Exit Sub

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " 	h.idx, h.themeIdx, h.subject, h.regdate, h.startdate, h.enddate, h.sortNo, h.isusing "
		sqlStr = sqlStr & " FROM db_board.dbo.tbl_giftmain_hotissue as h with (nolock)"
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY h.sortNo ASC"
		rsget.pagesize = FPageSize

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		
		
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new Cgift_item

					FItemList(i).Fidx = rsget("idx")
					FItemList(i).FthemeIdx = rsget("themeIdx")
					FItemList(i).Fsubject = db2html(rsget("subject"))
					FItemList(i).Fregdate = rsget("regdate")
					FItemList(i).Fstartdate = rsget("startdate")
					FItemList(i).Fenddate = rsget("enddate")
					FItemList(i).Fsortno = rsget("sortno")
					FItemList(i).Fisusing = rsget("isusing")

				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	

	'//subitem
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = "Select top 1 h.idx, h.themeIdx, h.subject, h.regdate, h.startdate, h.enddate, h.sortNo, h.isusing, h.reguserid, t.username "
        sqlStr = sqlStr & "From db_board.dbo.tbl_giftmain_hotissue as h with (nolock)"
        sqlStr = sqlStr & "inner join [db_partner].dbo.tbl_user_tenbyten as t with (nolock) on h.reguserid = t.userid "
        SqlStr = SqlStr & " where idx = '" & FRectIdx & "'"

		'rw SqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new Cgift_item
        if Not rsget.Eof then
			FOneItem.Fidx = rsget("idx")
			FOneItem.FthemeIdx = rsget("themeIdx")
			FOneItem.Fsubject = db2html(rsget("subject"))
			FOneItem.Fregdate = rsget("regdate")
			FOneItem.Fstartdate = rsget("startdate")
			FOneItem.Fenddate = rsget("enddate")
			FOneItem.Fisusing = rsget("isusing")
			FOneItem.Fsortno = rsget("sortNo")
			FOneItem.Freguser = rsget("username")
		else
			FResultCount = 0
        end if
        rsget.close
	End Sub
    	    
	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
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