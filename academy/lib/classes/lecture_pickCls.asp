<%
'######## 강좌Pick 레코드셋 #######
Class CLecPickItem
	'변수 선언
	public FpickSn
	public Fyyyymm
	public FlecLevel
	public FlecLvName
	public Fcdl
	public FcdlNm
	public FlecIdx
	public ForderNo
	public Fregdate
	public FlecTitle

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class


'####### 강좌Pick 클래스 #######
Class CLecPick

	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectPickSn
	public FRectYYYYMM
	public FRectCDL
	public FRectLecLevel

	'// 강좌Pick 목록 접수
	public sub GetLecPickList()
		dim SQL, AddSQL, loopList

		'검색 추가 쿼리
'		if FRectYYYYMM<>"" then
'			AddSQL = AddSQL & " and P.YYYYMM='" & FRectYYYYMM & "' "
'		end if

		if FRectCDL<>"" then
			AddSQL = AddSQL & " and P.code_large='" & FRectCDL & "' "
		end if

		if FRectLecLevel<>"" then
			AddSQL = AddSQL & " and P.lecLevel='" & FRectLecLevel & "' "
		end if

		'@ 총데이터수
		SQL =	" Select count(P.pickSn) as cnt, CEILING(CAST(Count(P.pickSn) AS FLOAT)/'"&FPageSize&"' ) as totPg " &_
				" From [db_academy].[dbo].tbl_lec_pickInfo as P " &_
				" 	join [db_academy].[dbo].tbl_lec_item as L " &_
				" 		on P.lecIdx=L.idx " &_
				" Where 1=1 " & AddSQL
		'response.Write SQL
		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"	P.*, L.lec_title, C.code_nm " &_
				" From [db_academy].[dbo].tbl_lec_pickInfo as P " &_
				" 	join [db_academy].[dbo].tbl_lec_item as L " &_
				" 		on P.lecIdx=L.idx " &_
				" 	join [db_academy].[dbo].tbl_lec_Cate_large as C " &_
				" 		on P.code_large=C.code_large " &_
				" Where 1=1 " & AddSQL &_
				" Order by P.orderNo, P.pickSn "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			loopList = 0
			rsACADEMYget.absolutepage = FCurrPage

			Do Until rsACADEMYget.eof
				set FItemList(loopList) = new CLecPickItem

				FItemList(loopList).FpickSn		= rsACADEMYget("pickSn")
				FItemList(loopList).Fyyyymm		= rsACADEMYget("YYYYMM")
				FItemList(loopList).FlecLevel	= rsACADEMYget("lecLevel")
				FItemList(loopList).FlecLvName	= getLecLevelNm(rsACADEMYget("lecLevel"))
				FItemList(loopList).Fcdl		= rsACADEMYget("code_large")
				FItemList(loopList).FcdlNm		= rsACADEMYget("code_nm")
				FItemList(loopList).FlecIdx		= rsACADEMYget("lecIdx")
				FItemList(loopList).ForderNo	= rsACADEMYget("orderNo")
				FItemList(loopList).Fregdate	= rsACADEMYget("regdate")
				FItemList(loopList).FlecTitle	= rsACADEMYget("lec_title")

				rsACADEMYget.MoveNext
				loopList = loopList + 1
			Loop

		end if
		rsACADEMYget.close
	end Sub

	public Function getLecLevelNm(lv)
		Select Case lv
			Case "L"
				getLecLevelNm = "초급"
			Case "M"
				getLecLevelNm = "중급"
			Case "H"
				getLecLevelNm = "고급"
		End Select
	end Function

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
		HasPreScroll = StartScrollPage > 1
	end Function


	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function


	'// 첫페이지 계산
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class
%>