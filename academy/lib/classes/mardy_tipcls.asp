<%
'######## 마디수첩 레코드셋 #######
Class CMardyTipItem
	'변수 선언
	public FtipId
	public Ftitle
	public FtipName
	public FtipUsage
	public FtipDef
	public FtipTime
	public FtipPrice
	public FtipAttent
	public FtipCont
	public FimgIcon
	public FimgIcon_full
	public Fuserid
	public Fusername
	public FhitCount
	public Fregdate
	public Fisusing

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

'######## 마디수첩 레코드셋 #######
Class CMardyTipImageItem
	'변수 선언
	public FimgId
	public FimgFile
	public FimgFile_full
	public FimgCont

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

'####### 마디수첩 클래스 #######
Class CMardyTip

	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectTipId
	public FRectsearchKey, FRectsearchString


	'// 마디수첩 목록 접수
	public sub GetMardyTipList()
		dim SQL, AddSQL, loopList

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		else
			AddSQL = ""
		end if

		'@ 총데이터수
		SQL =	" Select count(tipId) as cnt " &_
				" From [db_academy].[dbo].tbl_mardyTip " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		t1.tipId, t1.title, t1.tipName, t1.tipDef " &_
				"		,t1.ImgIcon, t1.regdate, '' as username, t1.hitCount ,t1.isusing " &_
				" From [db_academy].[dbo].tbl_mardyTip as t1 " &_

				" Where 1=1 " & AddSQL &_
				" Order by t1.tipId desc "
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
				set FItemList(loopList) = new CMardyTipItem

				FItemList(loopList).FtipId			= rsACADEMYget("tipId")
				FItemList(loopList).Ftitle			= db2html(rsACADEMYget("title"))

				FItemList(loopList).FtipName		= db2html(rsACADEMYget("tipName"))
				FItemList(loopList).FtipDef			= rsACADEMYget("tipDef")
				FItemList(loopList).Fusername		= db2html(rsACADEMYget("username"))
				FItemList(loopList).FhitCount		= db2html(rsACADEMYget("hitCount"))

				FItemList(loopList).Fregdate		= rsACADEMYget("regdate")
				FItemList(loopList).Fisusing		= rsACADEMYget("isusing")

				FItemList(loopList).FimgIcon		= rsACADEMYget("ImgIcon")
				FItemList(loopList).FimgIcon_full	= "http://image.thefingers.co.kr/contents/mardy_tip/icon/" & FItemList(loopList).FimgIcon

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close
	end Sub


	'// 마디수첩 내용 접수
	public sub GetMardyTipView()
		dim SQL

		SQL =	" Select " &_
				"		tipId, title, tipName, tipUsage, tipDef " &_
				"		,tipTime, tipPrice, tipAttent, tipCont " &_
				"		,ImgIcon, regdate, hitCount, isusing " &_
				" From [db_academy].[dbo].tbl_mardyTip " &_
				" Where tipId=" & FRectTipId
		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FItemList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FItemList(0) = new CMardyTipItem

			FItemList(0).FtipId			= rsACADEMYget("tipId")
			FItemList(0).Ftitle			= db2html(rsACADEMYget("title"))

			FItemList(0).FtipName		= db2html(rsACADEMYget("tipName"))
			FItemList(0).FtipUsage		= db2html(rsACADEMYget("tipUsage"))
			FItemList(0).FtipDef		= rsACADEMYget("tipDef")
			FItemList(0).FtipTime		= db2html(rsACADEMYget("tipTime"))
			FItemList(0).FtipPrice		= db2html(rsACADEMYget("tipPrice"))
			FItemList(0).FtipAttent		= db2html(rsACADEMYget("tipAttent"))
			FItemList(0).FtipCont		= db2html(rsACADEMYget("tipCont"))

			FItemList(0).FhitCount		= db2html(rsACADEMYget("hitCount"))

			FItemList(0).Fregdate		= rsACADEMYget("regdate")
			FItemList(0).Fisusing		= rsACADEMYget("isusing")

			FItemList(0).FimgIcon		= rsACADEMYget("ImgIcon")
			FItemList(0).FimgIcon_full	= "http://image.thefingers.co.kr/contents/mardy_tip/icon/" & FItemList(0).FimgIcon

		end if
		rsACADEMYget.close

	End Sub



	'// 마디수첩 서브 이미지 목록 접수
	public sub GetMardyTipImageList()
		dim SQL, AddSQL, loopList

		'@ 총데이터수
		SQL =	" Select count(imgId) as cnt " &_
				" From [db_academy].[dbo].tbl_mardyTipImage " &_
				" Where tipId = " & FRectTipId

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select " &_
				"		imgId, imgFile, imgCont " &_
				" From [db_academy].[dbo].tbl_mardyTipImage " &_
				" Where tipId = " & FRectTipId

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim preserve FItemList(FTotalCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CMardyTipImageItem

				FItemList(loopList).FimgId			= rsACADEMYget("imgId")
				FItemList(loopList).FimgCont		= db2html(rsACADEMYget("imgCont"))

				FItemList(loopList).FimgFile		= rsACADEMYget("imgFile")
				FItemList(loopList).FimgFile_full	= "http://image.thefingers.co.kr/contents/mardy_tip/image/" & FItemList(loopList).FimgFile

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