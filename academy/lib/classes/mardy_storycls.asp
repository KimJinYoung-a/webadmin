<%
'######## 마디이야기 레코드셋 #######
Class CMardyStoryItem
	'변수 선언
	public FstoryId
	public FtitleLong
	public FtitleShort
	public FimgIcon
	public FimgIcon_full
	public Fcontents
	public Fusername
	public FhitCount
	public Fregdate
	public Fisusing

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'######## 마디이야기 레코드셋 #######
Class CMardyStoryImageItem
	'변수 선언
	public FimgId
	public FimgFile
	public FimgFile_full

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'####### 마디이야기 클래스 #######
Class CMardyStory

	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectstoryId
	public FRectsearchKey, FRectsearchString

	'// 마디이야기 목록 접수
	public sub GetMardyStoryList()
		dim SQL, AddSQL, loopList

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		else
			AddSQL = ""
		end if


		'@ 총데이터수
		SQL =	" Select count(storyId) as cnt " &_
				" From [db_academy].[dbo].tbl_mardyStory " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		t1.storyId, t1.titleLong " &_
				"		,t1.ImgIcon, t1.regdate, '' as username, t1.hitCount, t1.isusing " &_
				" From [db_academy].[dbo].tbl_mardyStory as t1 " &_

				" Where 1=1 " & AddSQL &_
				" Order by t1.storyId desc "
''				"		Join db_user.[10x10].tbl_user_n as t2 on t1.userid=t2.userid " &_

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0
			rsACADEMYget.absolutepage = FCurrPage

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CMardyStoryItem

				FItemList(loopList).FstoryId			= rsACADEMYget("storyId")
				FItemList(loopList).FtitleLong			= db2html(rsACADEMYget("titleLong"))

				FItemList(loopList).Fusername		= db2html(rsACADEMYget("username"))
				FItemList(loopList).FhitCount		= db2html(rsACADEMYget("hitCount"))

				FItemList(loopList).Fregdate		= rsACADEMYget("regdate")
				FItemList(loopList).Fisusing		= rsACADEMYget("isusing")

				FItemList(loopList).FimgIcon		= rsACADEMYget("ImgIcon")
				FItemList(loopList).FimgIcon_full	= "http://image.thefingers.co.kr/contents/mardy_story/icon/" & FItemList(loopList).FimgIcon

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close
	end Sub


	'// 마디이야기 내용 접수
	public sub GetMardyStoryView()
		dim SQL

		SQL =	" Select " &_
				"		storyId, titleShort, titleLong, contents " &_
				"		,ImgIcon, regdate, hitCount, isusing " &_
				" From [db_academy].[dbo].tbl_mardyStory " &_
				" Where storyId=" & FRectstoryId
		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FItemList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FItemList(0) = new CMardyStoryItem

			FItemList(0).FstoryId		= rsACADEMYget("storyId")
			FItemList(0).FtitleShort	= db2html(rsACADEMYget("titleShort"))
			FItemList(0).FtitleLong		= db2html(rsACADEMYget("titleLong"))
			FItemList(0).Fcontents		= db2html(rsACADEMYget("contents"))

			FItemList(0).FhitCount		= db2html(rsACADEMYget("hitCount"))
			FItemList(0).Fregdate		= rsACADEMYget("regdate")
			FItemList(0).Fisusing		= rsACADEMYget("isusing")

			FItemList(0).FimgIcon		= rsACADEMYget("ImgIcon")
			FItemList(0).FimgIcon_full	= "http://image.thefingers.co.kr/contents/mardy_story/icon/" & FItemList(0).FimgIcon

		end if
		rsACADEMYget.close

	End Sub



	'// 마디이야기 서브 이미지 목록 접수
	public sub GetMardyStoryImageList()
		dim SQL, AddSQL, loopList

		'@ 총데이터수
		SQL =	" Select count(imgId) as cnt " &_
				" From [db_academy].[dbo].tbl_mardyStoryImage " &_
				" Where storyId = " & FRectstoryId

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select " &_
				"		imgId, imgFile " &_
				" From [db_academy].[dbo].tbl_mardyStoryImage " &_
				" Where storyId = " & FRectstoryId

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim preserve FItemList(FTotalCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CMardyStoryImageItem

				FItemList(loopList).FimgId			= rsACADEMYget("imgId")

				FItemList(loopList).FimgFile		= rsACADEMYget("imgFile")
				FItemList(loopList).FimgFile_full	= "http://image.thefingers.co.kr/contents/mardy_story/image/" & FItemList(loopList).FimgFile

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close
	end Sub




	'// 클래스 초기화
	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub


	'// 클래스 종료
	Private Sub Class_Terminate()

	End Sub


	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function


	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function


	'// 첫페이지 계산
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class
%>