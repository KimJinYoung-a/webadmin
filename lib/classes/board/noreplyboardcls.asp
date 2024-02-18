<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Class CNoReplyBoardCom
	public FcomID
	public Fwriter
	public FComment
	public FIsDelete
	public FregDate

	Private Sub Class_Initialize()
		'
	End Sub


	Private Sub Class_Terminate()
        '
	End Sub
end Class

Class CNoReplyBoardItem
	public FID
	public FTitle
	public Fmemo
	public Fregdate
	public Fwriter
	public FIsDelete
	public FBuyName
	public FOrderSerial
	public FCheckFlag
	public FcommentCount
	public FHitCount
	public FSitename
	public FMatchDate
	public FPreID
	public FNextID

	public FComItem()

	public Sub ReadComItems()
		Dim sqlstr,i
		sqlstr = "select id,contents,writer,convert(varchar,regdate,20) as regdate ,isdelete from tbl_board_delivery_com"
		sqlstr = sqlstr + " where masterid=" + CStr(FID)
		sqlstr = sqlstr + " and isdelete='N'"

		rsget.Open sqlstr, dbget, 1
		i=0

		FcommentCount = rsget.RecordCount

		redim preserve FComItem(FcommentCount)
		do until rsget.Eof
			set FComItem(i) = new CNoReplyBoardCom
			FComItem(i).FComID = rsget("id")
			FComItem(i).FComment = rsget("contents")
			FComItem(i).FWriter = rsget("writer")
			FComItem(i).FRegDate = rsget("regdate")
			FComItem(i).FIsDelete = rsget("isdelete")
			rsget.MoveNext
			i=i+1
		loop
		rsget.close
	end Sub

	Private Sub Class_Initialize()
		'
		FPreID =0
		FNextID = 0
		'redim preserve FComItem(0)
		redim FComItem(0)
	End Sub


	Private Sub Class_Terminate()
        '
	End Sub
end Class

Class CNoReplyBoard
	private FSitename
	private FTableName
	public FTotalcount
	private FTotalPage
	private FPageSize
	private FCurrentPage
	private FResultCount
	private FScrollCount

	private FStartDate
	private FEndDate
	private FWriter
	private FBuyer
	private FCheckFlag
	private FOrderSerial
	private FDeleteYn

	public FBoardItem()

	Property Get SiteName()
		SiteName = FSiteName
	end Property

	Property Get TableName()
		TableName = FTableName
	end Property

	Property Get TotalPage()
		TotalPage = FTotalPage
	end Property

	Property Get CurrentPage()
		CurrentPage = FCurrentPage
	end Property

	Property Get PageSize()
		PageSize = FPageSize
	end Property

	Property Get ResultCount()
		ResultCount = FResultCount
	end Property

	Property Get ScrollCount()
		ScrollCount = FScrollCount
	end Property

	Property Get StartDate()
		StartDate = FStartDate
	end Property

	Property Get EndDate()
		EndDate = FEndDate
	end Property

	Property Get Writer()
		Writer = FWriter
	end Property

	Property Get Buyer()
		Buyer = FBuyer
	end Property

	Property Get CheckFlag()
		CheckFlag = FCheckFlag
	end Property

	Property Get OrderSerial()
		OrderSerial = FOrderSerial
	end Property

	Property Get DEleteYn()
		DEleteYn = FDEleteYn
	end Property

	Property Let SiteName(byVal v)
		FSiteName = v
	end Property

	Property Let TableName(byVal v)
		FTableName = v
	end Property

	Property Let CurrentPage(byVal v)
		FCurrentPage = v
	end Property

	Property Let PageSize(byVal v)
		FPageSize = v
	end Property

	Property Let ScrollCount(byVal v)
		FScrollCount = v
	end Property

	Property Let StartDate(byVal v)
		FStartDate = v
	end Property

	Property Let EndDate(byVal v)
		FEndDate = v
	end Property

	Property Let Writer(byVal v)
		FWriter = v
	end Property

	Property Let Buyer(byVal v)
		FBuyer = v
	end Property

	Property Let CheckFlag(byVal v)
		FCheckFlag = v
	end Property

	Property Let OrderSerial(byVal v)
		FOrderSerial = v
	end Property

	Property Let DEleteYn(byVal v)
		FDEleteYn = v
	end Property

	Private Sub Class_Initialize()
		'
		FCurrentPage = 1
		FPageSize = 10
		FScrollCount = 10

		'redim preserve FBoardItem(0)
		redim  FBoardItem(0)
	End Sub


	Private Sub Class_Terminate()
        '
	End Sub

	Public function readBoard(byval iid)
		Dim sqlstr
		sqlstr = "update tbl_board_delivery set hitcount=hitcount+1"
		sqlstr = sqlstr + " where id=" + CStr(iid)
		rsget.Open sqlstr, dbget, 1

		sqlstr = "select * from tbl_board_delivery"
		sqlstr = sqlstr + " where id=" + CStr(iid)

		rsget.Open sqlstr, dbget, 1
		if Not rsget.Eof then
			set FBoardItem(0) = new CNoReplyBoardItem

			FBoardItem(0).FID = rsget("id")
		  	FBoardItem(0).FTitle = Db2Html(rsget("title"))
		  	FBoardItem(0).FSitename = rsget("sitename")
		  	FBoardItem(0).Fmemo = replace(Db2Html(rsget("memo")),vbcrlf,"<br>")
		  	FBoardItem(0).Fregdate = rsget("regdate")
		  	FBoardItem(0).FMatchDate = rsget("matchdate")
		  	FBoardItem(0).Fwriter = rsget("writer")
		  	FBoardItem(0).FIsDelete = rsget("deleteyn")
		  	FBoardItem(0).FBuyName = rsget("buyname")
		  	FBoardItem(0).FOrderSerial = rsget("orderserial")
		  	FBoardItem(0).FCheckFlag = rsget("checkflag")
		  	FBoardItem(0).FHitCount = rsget("hitcount")
			rsget.close

			sqlstr = "select * from tbl_board_delivery"
			sqlstr = sqlstr + " where id=" + CStr(iid-1)
		 	rsget.Open sqlstr, dbget, 1
		 		if Not rsget.Eof then
		 			FBoardItem(0).FNextID = iid-1
		 		end if
		 	rsget.close

		 	sqlstr = "select * from tbl_board_delivery"
			sqlstr = sqlstr + " where id=" + CStr(iid+1)
		 	rsget.Open sqlstr, dbget, 1
		 		if Not rsget.Eof then
		 			FBoardItem(0).FPreID = iid+1
		 		end if
		 	rsget.close

		 	FBoardItem(0).readComItems
		else
			rsget.close
		end if


	end function

	Public function delboard(byval iid)
		dim sql
		delboard = ""

        sql = "update tbl_board_delivery set deleteyn='Y' "
        sql = sql + " where id=" + CStR(iid)

        on error resume next

        rsget.Open sql, dbget, 1
        if err then
            delboard = err.description
            on error goto 0
        else
            delboard = ""
        end if
	end function

	Public function checkboard(byval iid,checkflag)
		dim sql
		checkboard = ""

        sql = "update tbl_board_delivery set checkflag='" + checkflag + "' "
        sql = sql + " where id=" + CStR(iid)

        on error resume next

        rsget.Open sql, dbget, 1
        if err then
            checkboard = err.description
            on error goto 0
        else
            checkboard = ""
        end if
	end function

	Public function writeCom(byval masterid,tx_com,writer)
		dim sql
		writeCom = ""

        sql = "insert into tbl_board_delivery_com (masterid, writer, contents) "
        sql = sql + " values(" + CStr(masterid) + ", '" + writer + "', '" + tx_com + "')"

        on error resume next

        rsget.Open sql, dbget, 1
        if err then
            writeCom = err.description
            on error goto 0
        else
            writeCom = ""
        end if
	end function

	Public function writeBoard (byval yyyymmdd,sitename,writer,buyname,orderserial,title,txmemo)
		dim sql
		writeboard = ""

        sql = "insert into tbl_board_delivery (sitename, matchdate, buyname, orderserial, writer, title, memo) "
        sql = sql + " values('" + sitename + "', '" + yyyymmdd + "', '" + buyname + "', '" + orderserial + "', '" + writer + "', '" + title + "','" + txmemo + "')"

        on error resume next

        rsget.Open sql, dbget, 1
        if err then
            writeboard = err.description
            on error goto 0
        else
            writeboard = ""
        end if

	end function

	Public function listBoard()

		  Dim sqlstr, query1
		  dim i

		  query1 = ""
		  sqlstr = "select count(id) cnt from tbl_board_delivery"
		  sqlstr = sqlstr + " where id <> 0 "

		  if FSiteName <> "" then
		  	query1 = query1 + " and sitename='" + FSiteName + "'"
		  end if

		  if FStartDate <> "" then
		      query1 = query1 + " and matchdate >= '" + FStartDate + "'"
		  end if

		  if FEndDate <> "" then
		      query1 = query1 + " and matchdate < '" + FEndDate + "'"
		  end if

		  if FWriter <> "" then
		      query1 = query1 + " and writer = '" + FWriter + "'"
		  end if

		  if FCheckFlag <> "" then
		      query1 = query1 + " and checkflag = '" + FCheckFlag + "'"
		  end if

		  if FBuyer <> "" then
		      query1 = query1 + " and buyname like '%" + FBuyer + "%'"
		  end if

		  if FOrderSerial <> "" then
		      query1 = query1 + " and orderserial like '%" + FOrderSerial + "%'"
		  end if

		  if FDeleteYn <> "" then
		      query1 = query1 + " and deleteyn like '" + FDeleteYn + "'"
		  end if

		  rsget.Open sqlstr+query1,dbget,1
		  FTotalcount = CInt(rsget("cnt"))
		  rsget.close

		  FTotalPage = int((FTotalcount-1)/FPageSize) +1


		  sqlstr = "SELECT TOP " + CStr(FPageSize) + " *, IsNull(T.commentCount,0) as commentCount "
		  sqlstr = sqlstr + " FROM tbl_board_delivery "
		  sqlstr = sqlstr + " left join (select count(id) as commentCount,masterid from tbl_board_delivery_com group by masterid) T on T.masterid =id"
		  sqlstr = sqlstr + " WHERE id not in "
		  sqlstr = sqlstr + "  (SELECT TOP " + CStr((FCurrentPage - 1) * FPageSize) + " id "
		  sqlstr = sqlstr + "   FROM tbl_board_delivery "
		  sqlstr = sqlstr + "   WHERE  id <> 0  "

		  sqlstr = sqlstr + query1

		  sqlstr = sqlstr + " ORDER BY id desc ) "

		  sqlstr = sqlstr + query1

		  sqlstr = sqlstr & " ORDER BY id desc "

		  ''response.write sqlstr
		  rsget.Open sqlstr,dbget,1

		  FResultCount = rsget.recordCount
		  redim preserve FBoardItem(FResultCount)

		  i=0
		  do until rsget.EOF
		  	set FBoardItem(i) = new CNoReplyBoardItem

		  	FBoardItem(i).FID = rsget("id")
		  	FBoardItem(i).FTitle = Db2Html(rsget("title"))
		  	FBoardItem(i).FSitename = rsget("sitename")
		  	FBoardItem(i).Fmemo = replace(Db2Html(rsget("memo")),vbcrlf,"<br>")
		  	FBoardItem(i).Fregdate = rsget("regdate")
		  	FBoardItem(i).FMatchDate = rsget("matchdate")
		  	FBoardItem(i).Fwriter = rsget("writer")
		  	FBoardItem(i).FIsDelete = rsget("deleteyn")
		  	FBoardItem(i).FBuyName = rsget("buyname")
		  	FBoardItem(i).FOrderSerial = rsget("orderserial")
		  	FBoardItem(i).FCheckFlag = rsget("checkflag")
		  	FBoardItem(i).FHitCount = rsget("hitcount")
		  	FBoardItem(i).FcommentCount = rsget("commentCount")

		  	rsget.MoveNext
		  	i=i+1
		  loop

		  rsget.close
	end Function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrentpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class
%>