<%
'#######################################################
'	History	:  2019.10.16 한용민 생성
'	Description : Link 발송 클래스
'#######################################################

Class CLinkSendItem
    public flinkidx
    public ftitle
    public flinkurl
    public fisusing
    public fviewcount
    public fregdate
    public flastupdate
    public flastadminid

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CLinkSend
	public FItemList()
	public FOneItem

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

    public frectlinkidx
    public frecttitle
    public frectisusing

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Sub GetLinkSend()
		dim sqlStr,addSql, i

		if frectlinkidx<>"" then
			addSql = addSql & " and linkidx = " & frectlinkidx & "" & vbcrlf
		end if
		if frecttitle<>"" then
			addSql = addSql & " and title like '%" & frecttitle & "%'" & vbcrlf
		end if
		if frectisusing<>"" then
			addSql = addSql & " and isusing = '" & frectisusing & "'" & vbcrlf
		end if

		'총수 접수
		sqlStr = "select count(linkidx), CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") " & vbcrlf
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_Link_SendList with (nolock)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " + addSql & vbcrlf

        'response.write sqlStr & "<Br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

		'내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) & "" & vbcrlf
		sqlStr = sqlStr & " linkidx, title, linkurl, isusing, viewcount, regdate, lastupdate, lastadminid" & vbcrlf
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_Link_SendList with (nolock)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & addSql & vbcrlf
		sqlStr = sqlStr & " order by linkidx desc"

        'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLinkSendItem

				FItemList(i).flinkidx = rsget("linkidx")
				FItemList(i).ftitle = db2html(rsget("title"))
				FItemList(i).flinkurl = db2html(rsget("linkurl"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fviewcount = rsget("viewcount")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastupdate = rsget("lastupdate")
				FItemList(i).flastadminid = rsget("lastadminid")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public sub GetLinkSend_one()
		dim sqlStr, addSql

		if frectlinkidx<>"" then
			addSql = addSql & " and linkidx = " & frectlinkidx & "" & vbcrlf
		end if

		sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " linkidx, title, linkurl, isusing, viewcount, regdate, lastupdate, lastadminid" & vbcrlf
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_Link_SendList with (nolock)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & addSql & vbcrlf
		sqlStr = sqlStr & " order by linkidx desc"
		
		'Response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		fresultcount = rsget.recordcount
		
		if not rsget.EOF then
			set FOneItem = new CLinkSendItem	

            FOneItem.flinkidx = rsget("linkidx")
            FOneItem.ftitle = db2html(rsget("title"))
            FOneItem.flinkurl = db2html(rsget("linkurl"))
            FOneItem.fisusing = rsget("isusing")
            FOneItem.fviewcount = rsget("viewcount")
            FOneItem.fregdate = rsget("regdate")
            FOneItem.flastupdate = rsget("lastupdate")
            FOneItem.flastadminid = rsget("lastadminid")

		end if
		rsget.close
	end sub	

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>