<%
'######## 마디스크랩 레코드셋 #######
Class CMardyScrapItem
	'변수 선언
	public FScrapId
	public Ftitle
	public FscrName
	public FscrDef
	public FscrTime
	public FscrSource
	public FscrTool
	public FimgTitle
	public FimgTitle_full
	public FimgProd
	public FimgProd_full
	public FimgIcon
	public FimgIcon_full
	public FscrTip
	public Fsummary
	public FprintType
	public Fusername
	public FhitCount
	public Fregdate
	public Fisusing

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'######## 마디스크랩 레코드셋 #######
Class CMardyScrapImageItem
	'변수 선언
	public FsubId
	public FsubName
	public FsubCont
	public FsubSort

	public FimgFile1
	public FimgFile1_full
	public FimgFile2
	public FimgFile2_full
	public FimgFile3
	public FimgFile3_full

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'######## 마디스크랩 레코드셋 #######
Class CMardyScrapSubItem
	'변수 선언
	public FsubId
	public FsubName
	public FsubCont
	public FsubSort

	public FimgFile1
	public FimgFile1_full
	public FimgFile2
	public FimgFile2_full
	public FimgFile3
	public FimgFile3_full

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'####### 마디스크랩 클래스 #######
Class CMardyScrap

	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectScrapId
	public FRectsearchKey, FRectsearchString

	'// 마디스크랩 목록 접수
	public sub GetMardyScrapList()
		dim SQL, AddSQL, loopList

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		else
			AddSQL = ""
		end if


		'@ 총데이터수
		SQL =	" Select count(ScrapId) as cnt " &_
				" From [db_academy].[dbo].tbl_mardyScrap " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		t1.ScrapId, t1.title " &_
				"		,t1.ImgIcon, t1.regdate, '' as username, t1.hitCount, t1.isusing " &_
				" From [db_academy].[dbo].tbl_mardyScrap as t1 " &_

				" Where 1=1 " & AddSQL &_
				" Order by t1.ScrapId desc "
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
				set FItemList(loopList) = new CMardyScrapItem

				FItemList(loopList).FScrapId		= rsACADEMYget("ScrapId")
				FItemList(loopList).Ftitle			= db2html(rsACADEMYget("title"))

				FItemList(loopList).Fusername		= db2html(rsACADEMYget("username"))
				FItemList(loopList).FhitCount		= db2html(rsACADEMYget("hitCount"))

				FItemList(loopList).Fregdate		= rsACADEMYget("regdate")
				FItemList(loopList).Fisusing		= rsACADEMYget("isusing")

				FItemList(loopList).FimgIcon		= rsACADEMYget("ImgIcon")
				FItemList(loopList).FimgIcon_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/icon/" & FItemList(loopList).FimgIcon

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close
	end Sub


	'// 마디스크랩 내용 접수
	public sub GetMardyScrapView()
		dim SQL

		SQL =	" Select " &_
				"		ScrapId, title, scrTip, imgTitle, scrName, scrDef, scrTime, scrSource, scrTool, printType  " &_
				"		, imgProd, ImgIcon, regdate, hitCount, isusing, summary " &_
				" From [db_academy].[dbo].tbl_mardyScrap " &_
				" Where ScrapId=" & FRectScrapId
		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FItemList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FItemList(0) = new CMardyScrapItem

			FItemList(0).FScrapId		= rsACADEMYget("ScrapId")
			FItemList(0).Ftitle			= db2html(rsACADEMYget("title"))

			FItemList(0).FscrName		= db2html(rsACADEMYget("scrName"))
			FItemList(0).FscrDef		= rsACADEMYget("scrDef")
			FItemList(0).FscrTime		= rsACADEMYget("scrTime")
			FItemList(0).FscrSource		= db2html(rsACADEMYget("scrSource"))
			FItemList(0).FscrTool		= db2html(rsACADEMYget("scrTool"))
			FItemList(0).FprintType		= rsACADEMYget("printType")
			
			FItemList(0).FscrTip		= db2html(rsACADEMYget("scrTip"))

			FItemList(0).FhitCount		= db2html(rsACADEMYget("hitCount"))
			FItemList(0).Fregdate		= rsACADEMYget("regdate")
			FItemList(0).Fisusing		= rsACADEMYget("isusing")
			
			FItemList(0).Fsummary		= rsACADEMYget("summary")

			FItemList(0).FimgTitle		= rsACADEMYget("ImgTitle")
			FItemList(0).FimgTitle_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/main/" & FItemList(0).FimgTitle
			FItemList(0).FimgProd		= rsACADEMYget("ImgProd")
			FItemList(0).FimgProd_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/main/" & FItemList(0).FimgProd
			FItemList(0).FimgIcon		= rsACADEMYget("ImgIcon")
			FItemList(0).FimgIcon_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/icon/" & FItemList(0).FimgIcon

		end if
		rsACADEMYget.close

	End Sub



	'// 마디스크랩 서브 이미지 목록 접수
	public sub GetMardyScrapImageList()
		dim SQL, AddSQL, loopList

		'@ 총데이터수
		SQL =	" Select count(subId) as cnt " &_
				" From [db_academy].[dbo].tbl_mardyScrapSub " &_
				" Where isusing='Y' and ScrapId = " & FRectScrapId

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select " &_
				"		subId, subName, subCont, imgFile1, imgFile2, imgFile3, subSort " &_
				" From [db_academy].[dbo].tbl_mardyScrapSub " &_
				" Where isusing='Y' and ScrapId = " & FRectScrapId &_
				" Order by subSort "

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim preserve FItemList(FTotalCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CMardyScrapImageItem

				FItemList(loopList).FsubId			= rsACADEMYget("subId")
				FItemList(loopList).FsubName		= rsACADEMYget("subName")
				FItemList(loopList).FsubCont		= rsACADEMYget("subCont")
				FItemList(loopList).FsubSort		= rsACADEMYget("subSort")

				FItemList(loopList).FimgFile1		= rsACADEMYget("imgFile1")
				FItemList(loopList).FimgFile1_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/sub_thumb/" & FItemList(loopList).FimgFile1
				FItemList(loopList).FimgFile2		= rsACADEMYget("imgFile2")
				FItemList(loopList).FimgFile2_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/sub_thumb/" & FItemList(loopList).FimgFile2
				FItemList(loopList).FimgFile3		= rsACADEMYget("imgFile3")
				FItemList(loopList).FimgFile3_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/sub_thumb/" & FItemList(loopList).FimgFile3

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


'####### 마디스크랩 서브 클래스 #######
Class CMardyScrapSub

	public FItemView()

	public FRectScrapId, FRectSubId

	'// 마디스크랩 서브 이미지 내용 접수
	public sub GetMardyScrapImageView()
		dim SQL, AddSQL

		'@ 데이터
		SQL =	" Select " &_
				"		subId, subName, subCont, imgFile1, imgFile2, imgFile3, subSort " &_
				" From [db_academy].[dbo].tbl_mardyScrapSub " &_
				" Where subId = " & FRectSubId

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FItemView(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FItemView(0) = new CMardyScrapSubItem

			FItemView(0).FsubName		= rsACADEMYget("subName")
			FItemView(0).FsubCont		= rsACADEMYget("subCont")
			FItemView(0).FsubSort		= rsACADEMYget("subSort")

			FItemView(0).FimgFile1		= rsACADEMYget("imgFile1")
			FItemView(0).FimgFile1_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/sub_thumb/" & FItemView(0).FimgFile1
			FItemView(0).FimgFile2		= rsACADEMYget("imgFile2")
			FItemView(0).FimgFile2_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/sub_thumb/" & FItemView(0).FimgFile2
			FItemView(0).FimgFile3		= rsACADEMYget("imgFile3")
			FItemView(0).FimgFile3_full	= "http://image.thefingers.co.kr/contents/mardy_Scrap/sub_thumb/" & FItemView(0).FimgFile3

		end if
		rsACADEMYget.close
	end Sub


	'// 클래스 초기화
	Private Sub Class_Initialize()
		redim  FItemView(0)
	End Sub


	'// 클래스 종료
	Private Sub Class_Terminate()

	End Sub

end Class
%>