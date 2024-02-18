<%
	'// Recruit 클래스 아이템 //
	Class CRecruitItem
		public Frcb_sn
		public Fuserid
		public Frcb_subject
		public Frcb_content
		public Frcb_regdate
		public Frcb_startdate
		public Frcb_enddate
		public Frcb_state
		public Frcb_hit
		public Frcb_career
		public Frcb_jobtype
		public Frcb_recruit_url
		
		public Frcb_always
		public Frcb_personal
		

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

	'// Recruit 클래스 //
	Class CRecruit
		public FitemList()
		public FFileList()
		public FPreNext()
	
		public FTotalCount
		public FResultCount
	
		public FCurrPage
		public FTotalPage
		public FPageSize
		public FScrollCount
	
		public FRectRcbSn
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
		public Sub getRecruitList()
			dim SQL, Add_SQL
			dim oItem, i
	
			'검색 키워드
			if FRectSearchArea<>"" and FRectsearchKeyword<>"" then
				ADD_SQL = " and " & FRectSearchArea & " like '%" & FRectsearchKeyword & "%' "
			end if


			''#################################################
			''총 갯수.
			''#################################################
			SQL =	"Select Count(rcb_sn), CEILING(CAST(Count(rcb_sn) AS FLOAT)/" & FPageSize & ") " &_
					"From db_company.dbo.tbl_recruit_board " &_
					"Where 1=1 " & ADD_SQL
			
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
					"	rcb_sn, rcb_subject, rcb_regdate, id, rcb_hit " &_
					"	,rcb_startdate, rcb_enddate, rcb_state, rcb_career, rcb_jobtype, rcb_recruit_url, rcb_always, rcb_personal " &_
					"From db_company.dbo.tbl_recruit_board " &_
					"Where 1=1 " & ADD_SQL &_
					" order by rcb_sn desc"
					
			rsCompanyGet.pagesize = FPageSize
			rsCompanyGet.Open SQL,dbCompanyGet,1

			FResultCount = rsCompanyGet.RecordCount-(FPageSize*(FCurrPage-1))
	
			redim preserve FitemList(FResultCount)
			i=0
			if Not(rsCompanyGet.EOF or rsCompanyGet.BOF) then
				rsCompanyGet.absolutepage = FCurrPage
				do until rsCompanyGet.eof
					set oItem = new CRecruitItem
	
					oItem.Frcb_sn		= rsCompanyGet("rcb_sn")
					oItem.Frcb_subject	= rsCompanyGet("rcb_subject")
					oItem.Frcb_regdate	= rsCompanyGet("rcb_regdate")
					oItem.Frcb_startdate	= rsCompanyGet("rcb_startdate")
					oItem.Frcb_enddate	= rsCompanyGet("rcb_enddate")
					oItem.Frcb_state	= rsCompanyGet("rcb_state")
					oItem.Frcb_hit		= rsCompanyGet("rcb_hit")
					oItem.Fuserid			= rsCompanyGet("id")
					oItem.Fuserid			= rsCompanyGet("id")

					''2017-02-16 유태욱 추가(rcb_jobtype-직무종류, rcb_career-미선택0,신입1,경력2, rcb_recruit_url-지원하러가기URL)
					oItem.Frcb_career			= rsCompanyGet("rcb_career")
					oItem.Frcb_jobtype			= rsCompanyGet("rcb_jobtype")
					oItem.Frcb_recruit_url		= rsCompanyGet("rcb_recruit_url")
					oItem.Frcb_always			= rsCompanyGet("rcb_always")
					oItem.Frcb_personal			= rsCompanyGet("rcb_personal")
					
					


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
		public Sub getRecruitCont()
			dim SQL, oItem

			SQL =	"Select rcb_subject, rcb_content, rcb_regdate, id, rcb_hit " &_
					"	,rcb_startdate, rcb_enddate, rcb_state, rcb_career, rcb_jobtype, rcb_recruit_url, rcb_always, rcb_personal  " &_
					"From db_company.dbo.tbl_recruit_board " &_
					"Where rcb_sn=" & FRectRcbSn
			rsCompanyGet.Open SQL,dbCompanyGet,1

			redim preserve FitemList(1)
			if Not(rsCompanyGet.EOF or rsCompanyGet.BOF) then
				set oItem = new CRecruitItem

				oItem.Frcb_subject	= rsCompanyGet("rcb_subject")
				oItem.Frcb_content	= rsCompanyGet("rcb_content")
				oItem.Frcb_regdate	= rsCompanyGet("rcb_regdate")
				oItem.Frcb_startdate	= rsCompanyGet("rcb_startdate")
				oItem.Frcb_enddate	= rsCompanyGet("rcb_enddate")
				oItem.Frcb_state	= rsCompanyGet("rcb_state")
				oItem.Fuserid		= rsCompanyGet("id")
				oItem.Frcb_hit		= rsCompanyGet("rcb_hit")

				''2017-02-16 유태욱 추가(rcb_jobtype-직무종류, rcb_career-미선택0,신입1,경력2, rcb_recruit_url-지원하러가기URL)
				oItem.Frcb_career	= rsCompanyGet("rcb_career")
				oItem.Frcb_jobtype	= rsCompanyGet("rcb_jobtype")
				oItem.Frcb_recruit_url	= rsCompanyGet("rcb_recruit_url")
				oItem.Frcb_always	= rsCompanyGet("rcb_always")
				oItem.Frcb_personal	= rsCompanyGet("rcb_personal")
				
				

	   			set FitemList(1) = oItem
	   			set oItem = Nothing
			end if
			rsCompanyGet.close
		end Sub

		'## 첨부파일 목록 접수
		public Sub getRecruitFile()
			dim SQL, oItem, i

			SQL =	"Select file_sn, file_name, file_size, file_ext " &_
					"From db_company.dbo.tbl_recruit_board_file " &_
					"Where rcb_sn=" & FRectRcbSn
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


	'// 현재상태 지정 함수 //
	public Function getRecruitState(stat, sDt, eDt)
		if stat="0" then
			if DateDiff("d",sDt, date())<0 then
				'예정
				getRecruitState = "<font color=darkgreen>예정</font>"
			elseif DateDiff("d",eDt, date())<=0 then
				getRecruitState = "<font color=red>진행중</font>"
			else
				getRecruitState = "<font color=#898989>마감</font>"
			end if
		else
			'강제 종료
			getRecruitState = "<font color=#898989>마감</font>"
		end if
	end Function


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