<%
	'// PR Board 클래스 아이템 //
	Class CboardItem
		public Fbrd_sn
		public Fuserid
		public Fbrd_div
		public Fbrd_subject
		public Fbrd_content
		public Fbrd_regdate
		public Fbrd_hit

		Private Sub Class_Initialize()
		End Sub

		Private Sub Class_Terminate()
		End Sub
	end Class

	'// 첨부파일 클래스 아이템 //
	Class CFileItem
		public Ffile_sn
		public Ffile_name
		public Ffile_size
		public Ffile_ext

		Private Sub Class_Initialize()
		End Sub

		Private Sub Class_Terminate()
		End Sub
	end Class

	'// PR Board 클래스 //
	Class Cboard
		public FitemList()
		public FFileList()
		public FPreNext()
	
		public FTotalCount
		public FResultCount
	
		public FCurrPage
		public FTotalPage
		public FPageSize
		public FScrollCount
	
		public FRectBrdDiv
		public FRectBrdSn
		public FRectSearchArea
		public FRectsearchKeyword
			
		Private Sub Class_Initialize()
			redim	FitemList(0)
			redim	FfileList(0)
			redim	FPreNext(2)
	
			FCurrPage =1
			FPageSize = 10
			FResultCount = 0
			FScrollCount = 10
			FTotalCount =0
		End Sub

		Private Sub Class_Terminate()
		End Sub

		'## 목록 접수
		public Sub getBoardList()
			dim SQL, Add_SQL
			dim oItem, i
	
			'검색 키워드
			if FRectsearchKeyword<>"" then
				ADD_SQL = " and " & FRectSearchArea & " like '%" & FRectsearchKeyword & "%' "
			end if


			''#################################################
			''총 갯수.
			''#################################################
			SQL =	"Select Count(brd_sn), CEILING(CAST(Count(brd_sn) AS FLOAT)/" & FPageSize & ") " &_
					"From db_company.dbo.tbl_pr_board " &_
					"Where brd_div=" & FRectBrdDiv & ADD_SQL
			
			rsCompanyGet.Open SQL,dbCompanyGet,1
				FTotalCount = rsCompanyGet(0)
				FtotalPage = rsCompanyGet(1)
			rsCompanyGet.Close
	
			if Cint(FtotalPage)>0 and Cint(FtotalPage)<Cint(FCurrpage) then
				FCurrpage = FtotalPage
			end if
	
			''#################################################
			''현재 페이지 리스트.
			''#################################################
			SQL =	"Select top " & CStr(FPageSize*FCurrpage) &_
					"	brd_sn, brd_subject, brd_regdate, id, brd_hit " &_
					"From db_company.dbo.tbl_pr_board " &_
					"Where brd_div=" & FRectBrdDiv & ADD_SQL &_
					" order by brd_sn desc"
					
			rsCompanyGet.pagesize = FPageSize
			rsCompanyGet.Open SQL,dbCompanyGet,1

			FResultCount = rsCompanyGet.RecordCount-(FPageSize*(FCurrPage-1))
	
			redim preserve FitemList(FResultCount)
			i=0
			if Not(rsCompanyGet.EOF or rsCompanyGet.BOF) then
				rsCompanyGet.absolutepage = FCurrPage
				do until rsCompanyGet.eof
					set oItem = new CboardItem
	
					oItem.Fbrd_sn		= rsCompanyGet("brd_sn")
					oItem.Fbrd_subject	= rsCompanyGet("brd_subject")
					oItem.Fbrd_regdate	= rsCompanyGet("brd_regdate")
					oItem.Fbrd_hit		= rsCompanyGet("brd_hit")
					oItem.Fuserid			= rsCompanyGet("id")
		
		   			set FitemList(i) = oItem
		   			set oItem = Nothing
	
					rsCompanyGet.MoveNext
		   			i=i+1
				loop
			end if
			rsCompanyGet.close
		end sub
	
		public Function HasPreScroll()
			HasPreScroll = StartScrollPage > 1
		end Function
	
		public Function HasNextScroll()
			HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
		end Function
	
		public Function StartScrollPage()
			StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
		end Function



		'## 게시물 내용 접수
		public Sub getBoardCont()
			dim SQL, oItem

			SQL =	"Select brd_div, brd_subject, brd_content, brd_regdate, id, brd_hit " &_
					"From db_company.dbo.tbl_pr_board " &_
					"Where brd_sn=" & FRectBrdSn
			rsCompanyGet.Open SQL,dbCompanyGet,1

			redim preserve FitemList(1)
			FResultCount = rsCompanyGet.RecordCount
			if Not(rsCompanyGet.EOF or rsCompanyGet.BOF) then
				set oItem = new CboardItem

				oItem.Fbrd_div		= rsCompanyGet("brd_div")
				oItem.Fbrd_subject	= rsCompanyGet("brd_subject")
				oItem.Fbrd_content	= rsCompanyGet("brd_content")
				oItem.Fbrd_regdate	= rsCompanyGet("brd_regdate")
				oItem.Fuserid		= rsCompanyGet("id")
				oItem.Fbrd_hit		= rsCompanyGet("brd_hit")

	   			set FitemList(1) = oItem
	   			set oItem = Nothing
			end if
			rsCompanyGet.close
		end Sub


		'## 첨부파일 목록 접수
		public Sub getBoardFile()
			dim SQL, oItem, i

			SQL =	"Select file_sn, file_name, file_size, file_ext " &_
					"From db_company.dbo.tbl_pr_board_file " &_
					"Where brd_sn=" & FRectBrdSn
			rsCompanyGet.Open SQL,dbCompanyGet,1

			FResultCount = rsCompanyGet.RecordCount
	
			redim preserve FfileList(FResultCount)
			i=0
			if Not(rsCompanyGet.EOF or rsCompanyGet.BOF) then
				do until rsCompanyGet.eof
					set oItem = new CfileItem
	
					oItem.Ffile_sn		= rsCompanyGet("file_sn")
					oItem.Ffile_name	= rsCompanyGet("file_name")
					oItem.Ffile_size	= rsCompanyGet("file_size")
					oItem.Ffile_ext		= rsCompanyGet("file_ext")
		
		   			set FfileList(i) = oItem
		   			set oItem = Nothing
	
					rsCompanyGet.MoveNext
		   			i=i+1
				loop
			end if
			rsCompanyGet.close
		end Sub

	end class


	'// 파일용량 출력 함수 //
	public Function printFileSize(fs)
		if fs="" or isNull(fs) then fs=0
		fs = Clng(fs)

		if (fs/1024)<1 then
			'byte
			printFileSize = FormatNumber(fs,0) & "bytes"
		elseif (fs/1024/1024)<1 then
			'Kilo
			printFileSize = FormatNumber(fs/1024,0) & "KB"
		elseif (fs/1024/1024/1024)<1 then
			'Mega
			printFileSize = FormatNumber(fs/1024/1024,0) & "MB"
		elseif (fs/1024/1024/1024/1024)<1 then
			'Giga
			printFileSize = FormatNumber(fs/1024/1024/1024,0) & "GB"
		end if
	end Function 
%>