<%
'##### 공통코드 레코드셋용 클래스 #####
class CCommItem

	public FcommCd
	public FcommNm
	public FgroupCd
	public FgroupNm
	public FisUsing

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### 공통코드 클래스 #####
Class CComm

	public FCommList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectCommCd
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString
	public FRectisUsing

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FCommList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// Comm 분류별 목록 출력
	public Sub GetCommList()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and t1." & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and t1.groupCd='" & FRectsearchDiv & "' "
		end if

		if FRectisUsing<>"" then
			AddSQL = AddSQL & " and t1.isUsing='" & FRectisUsing & "' "
		end if

		'@ 총데이터수
		SQL =	" Select count(CommCd) as cnt " &_
				" From db_academy.dbo.tbl_CommCd as t1 " &_
				"		Join db_academy.dbo.tbl_groupCd as t2 on t1.groupCd=t2.groupCd " &_
				" Where t2.isUsing='Y' " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		CommCd, CommNm, t1.groupCd, groupNm " &_
				"		,Case t1.isusing When 'Y' Then '<font color=darkblue>사용</font>' When 'N' Then '<font color=darkred>삭제</font>' End isusing " &_
				" From db_academy.dbo.tbl_CommCd as t1 " &_
				"		Join db_academy.dbo.tbl_groupCd as t2 on t1.groupCd=t2.groupCd " &_
				" Where t2.isUsing='Y' " & AddSQL &_
				" Order by t1.GroupCd, t1.CommCd"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FCommList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FCommList(lp) = new CCommItem

				FCommList(lp).FCommCd		= rsACADEMYget("CommCd")
				FCommList(lp).FcommNm		= rsACADEMYget("commNm")
				FCommList(lp).FgroupCd		= rsACADEMYget("groupCd")
				FCommList(lp).FgroupNm		= rsACADEMYget("groupNm")
				FCommList(lp).Fisusing		= rsACADEMYget("isusing")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close

	end Sub


	'// Comm 내용 보기
	public Sub GetCommRead()
		dim SQL

		SQL =	" Select t1.commCd, commNm, t1.groupCd, groupNm " &_
				"		,Case t1.isusing When 'Y' Then '사용' When 'N' Then '삭제' End isusing " &_
				" From db_academy.dbo.tbl_CommCd as t1 " &_
				"		Join db_academy.dbo.tbl_groupCd as t2 on t1.groupCd=t2.groupCd " &_
				" Where CommCd = '" & FRectCommCd & "'"

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim FCommList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FCommList(0) = new CCommItem

			FCommList(0).FCommCd	= rsACADEMYget("CommCd")
			FCommList(0).FcommNm	= rsACADEMYget("commNm")
			FCommList(0).FgroupCd	= rsACADEMYget("groupCd")
			FCommList(0).FgroupNm	= rsACADEMYget("groupNm")
			FCommList(0).Fisusing	= rsACADEMYget("isusing")

		end if
		rsACADEMYget.close

	end sub


	'// 그룹 옵션 생성 //
	function optGroupCd(nowCd)
		dim SQL, strOpt

		SQL =	"Select groupCd, groupNm From db_academy.dbo.tbl_groupCd Where isusing='Y'"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			Do Until rsACADEMYget.EOF
				strOpt = strOpt & "<option value='" & rsACADEMYget("groupCd") & "' "

				if nowCd=rsACADEMYget("groupCd") then strOpt = strOpt & "selected"

				strOpt = strOpt & " >" & rsACADEMYget("groupNm") & "</option>"
				rsACADEMYget.MoveNext
			Loop
		end if

		rsACADEMYget.Close

		optGroupCd = strOpt

	end function


	public FPrevID
	public FNextID

	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// 첫페이지 산출
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

%>