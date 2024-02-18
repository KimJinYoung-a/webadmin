<%
Class cDownFileItem
	public FfileSn
	public FfileTitle
	public FfileName
	public FfileDownNm		'전송시 파일명
	public FfileSize
	public Fisusing
	public Fregdate
	public FdownCount
	public FlastDownDate
	public Fevt_code
	
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class cDownFile
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectFSN
	public FRectUsing

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function GetfileList()
		dim sqlStr, addSql, i

		'추가 조건절
		if FRectFSN<>"" then
			addSql = addSql & " and fileSn='" & FRectFSN & "'"
		end if
		if FRectUsing<>"" then
			addSql = addSql & " and isUsing='" & FRectUsing & "'"
		end if

		'카운트
		sqlStr = "select count(fileSn) as cnt"
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_DownloadFile as t1" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " fileSn, fileTitle, fileName, fileDownNm, fileSize, downCount, lastDownDate, isusing, regdate, evt_code " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_DownloadFile " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " order by fileSn desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new cDownFileItem

				FItemList(i).FfileSn		= rsget("fileSn")
				FItemList(i).FfileTitle		= db2html(rsget("fileTitle"))
				FItemList(i).FfileName		= rsget("fileName")
				FItemList(i).FfileDownNm	= rsget("fileDownNm")
				FItemList(i).FfileSize		= rsget("fileSize")
				FItemList(i).FdownCount		= rsget("downCount")
				FItemList(i).FlastDownDate	= rsget("lastDownDate")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Fevt_code	= rsget("evt_code")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

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