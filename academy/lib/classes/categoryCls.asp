<%
'##### 카테고리 코드 레코드셋용 클래스 #####
class CCateItem

	Public FcateLargeCd
	Public FlargeCate_Name
	public FCateCd
	public FCateCD_Name
	public FCateCD_NameEng
	public FisUsing
	public FsortNo

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### 케테고리 코드 클래스 #####
Class CCate

	public FCateList()
	public FTotalCount
	public FCateDiv
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectCateCd
	public FRectsearchKey
	public FRectsearchString
	public FRectisUsing

	Public FRectLargeCateCd

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FCateList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// 카테고리1 목록 출력
	public Sub GetCateList1()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		'@ 총데이터수
		SQL =	" Select count(CateCD1) as cnt " &_
				" From db_academy.dbo.tbl_lec_Cate1 " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close

		'@데이터 접수
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		CateCD1, CateCD1_Name " &_
				" From db_academy.dbo.tbl_lec_Cate1 " &_
				" Where 1=1 " & AddSQL &_
				" Order by CateCD1"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FCateList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FCateList(lp) = new CCateItem

				FCateList(lp).FCateCd		= rsACADEMYget("CateCD1")
				FCateList(lp).FCateCD_Name	= rsACADEMYget("CateCD1_Name")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end Sub

	'// 카테고리2 목록 출력
	public Sub GetCateList2()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectisUsing<>"" then
			AddSQL = AddSQL & " and isUsing='" & FRectisUsing & "' "
		end if

		'@ 총데이터수
		SQL =	" Select count(CateCD2) as cnt " &_
				" From db_academy.dbo.tbl_lec_Cate2 " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close

		'@데이터 접수
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		CateCD2, CateCD2_Name, CateCD2_Name_Eng " &_
				"		,Case isusing When 'Y' Then '<font color=darkblue>사용</font>' When 'N' Then '<font color=darkred>삭제</font>' End isusing " &_
				"		,SortNo" &_
				" From db_academy.dbo.tbl_lec_Cate2 " &_
				" Where 1=1 " & AddSQL &_
				" Order by SortNo, CateCD2"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FCateList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FCateList(lp) = new CCateItem

				FCateList(lp).FCateCd		= rsACADEMYget("CateCD2")
				FCateList(lp).FCateCD_Name	= rsACADEMYget("CateCD2_Name")
				FCateList(lp).FCateCD_NameEng	= rsACADEMYget("CateCD2_Name_Eng")
				FCateList(lp).Fisusing		= rsACADEMYget("isusing")
				FCateList(lp).FsortNo		= rsACADEMYget("SortNo")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end Sub

	'// 카테고리3 목록 출력
	public Sub GetCateList3()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectisUsing<>"" then
			AddSQL = AddSQL & " and isUsing='" & FRectisUsing & "' "
		end if

		'@ 총데이터수
		SQL =	" Select count(CateCD3) as cnt " &_
				" From db_academy.dbo.tbl_lec_Cate3 " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close

		'@데이터 접수
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		CateCD3, CateCD3_Name " &_
				"		,Case isusing When 'Y' Then '<font color=darkblue>사용</font>' When 'N' Then '<font color=darkred>삭제</font>' End isusing " &_
				"		,SortNo" &_
				" From db_academy.dbo.tbl_lec_Cate3 " &_
				" Where 1=1 " & AddSQL &_
				" Order by SortNo, CateCD3"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FCateList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FCateList(lp) = new CCateItem

				FCateList(lp).FCateCd		= rsACADEMYget("CateCD3")
				FCateList(lp).FCateCD_Name	= rsACADEMYget("CateCD3_Name")
				FCateList(lp).Fisusing		= rsACADEMYget("isusing")
				FCateList(lp).FsortNo		= rsACADEMYget("SortNo")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end Sub


	'// Cate 내용 보기
	public Sub GetCateRead()
		dim SQL
		'FRectLargeCateCd
		'구분에 따른 분기
		select Case FCateDiv

			Case "code_large"
				SQL =	" Select code_large, code_nm , orderNo " &_
						" From db_academy.dbo.tbl_lec_Cate_large " &_
						" Where code_large = '" & FRectCateCd & "'"

			Case "code_mid"
				SQL =	" Select code_large, code_mid , code_nm , code_nm_eng " &_
						"		,Case display_yn When 'Y' Then '사용' When 'N' Then '삭제' End display_yn " &_
						"		,orderNo " &_
						" From db_academy.dbo.tbl_lec_Cate_mid " &_
						" Where code_large = '" & FRectLargeCateCd & "' and code_mid = '"& FRectCateCd &"' "

		end Select

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount

		redim FCateList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FCateList(0) = new CCateItem

			if FCateDiv="code_mid" then
			FCateList(0).FcateLargeCd	= rsACADEMYget("code_large")
			End If 
			FCateList(0).FCateCd	= rsACADEMYget(FCateDiv)
			FCateList(0).FCateCD_Name	= rsACADEMYget("code_nm")
			if FCateDiv="code_mid" then
				FCateList(0).FCateCD_NameEng	= rsACADEMYget("code_nm_eng")
			end if
			if FCateDiv="code_mid" then
				FCateList(0).Fisusing	= rsACADEMYget("display_yn")
				FCateList(0).FsortNo	= rsACADEMYget("orderNo")
			end if

		end if
		rsACADEMYget.close

	end sub

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

	'// 대카테고리 목록 출력
	public Sub GetLargeCateList()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		'@ 총데이터수
		SQL =	" Select count(code_large) as cnt " &_
				" From db_academy.dbo.tbl_lec_Cate_large " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close

		'@데이터 접수
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		code_large, code_nm " &_
				"		,Case display_yn When 'Y' Then '<font color=darkblue>사용</font>' When 'N' Then '<font color=darkred>비사용</font>' End display_yn " &_
				"		,orderNo " &_
				" From db_academy.dbo.tbl_lec_Cate_large " &_
				" Where 1=1 " & AddSQL &_
				" Order by orderNo asc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FCateList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FCateList(lp) = new CCateItem

				FCateList(lp).FCateCd		= rsACADEMYget("code_large")
				FCateList(lp).FCateCD_Name	= rsACADEMYget("code_nm")
				FCateList(lp).Fisusing		= rsACADEMYget("display_yn")
				FCateList(lp).FsortNo		= rsACADEMYget("orderNo")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end Sub

	'// 중카테고리 목록 출력
	public Sub GetMidCateList()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and m." & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		'@ 총데이터수
		SQL =	" Select count(M.code_mid) as cnt " &_
				" From db_academy.dbo.tbl_lec_Cate_mid as M " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close

		'@데이터 접수
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		L.code_large , L.code_nm as large_nm  , M.code_mid, M.code_nm as min_nm , M.code_nm_eng" &_
				"		,Case M.display_yn When 'Y' Then '<font color=darkblue>사용</font>' When 'N' Then '<font color=darkred>삭제</font>' End display_yn " &_
				"		,M.orderNo" &_
				" From db_academy.dbo.tbl_lec_Cate_large as L inner join db_academy.dbo.tbl_lec_Cate_mid as M on L.code_large = M.code_large " &_
				" Where 1=1 " & AddSQL &_
				" Order by L.code_large asc ,M.orderNo asc "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FCateList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FCateList(lp) = new CCateItem

				FCateList(lp).FcateLargeCd			= rsACADEMYget("code_large")
				FCateList(lp).FlargeCate_Name		= rsACADEMYget("large_nm")
				FCateList(lp).FCateCd					= rsACADEMYget("code_mid")
				FCateList(lp).FCateCD_Name			= rsACADEMYget("min_nm")
				FCateList(lp).FCateCD_NameEng	= rsACADEMYget("code_nm_eng")
				FCateList(lp).Fisusing					= rsACADEMYget("display_yn")
				FCateList(lp).FsortNo					= rsACADEMYget("orderNo")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end Sub

	public function GetNewCateCurrentPos(cdl,cdm,cds)
		dim sqlStr
		sqlStr = "select distinct top 1 code_nm "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_Cate_large"
		sqlStr = sqlStr + " where code_large='" + cdl + "'"
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		if not rsACADEMYget.Eof then
			GetNewCateCurrentPos = db2html(rsACADEMYget("code_nm"))
		end if
		rsACADEMYget.close


		if cdm<>"" then
			sqlStr = "select distinct top 1 code_nm "
			sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_Cate_mid"
			sqlStr = sqlStr + " where code_large='" + cdl + "'"
			sqlStr = sqlStr + " and code_mid='" + cdm + "'"
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			if not rsACADEMYget.Eof then
				GetNewCateCurrentPos = GetNewCateCurrentPos + "-" +  db2html(rsACADEMYget("code_nm"))
			end if
			rsACADEMYget.close
		end if

		if cds<>"" then
			sqlStr = "select distinct top 1 code_nm "
			sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_Cate_small"
			sqlStr = sqlStr + " where code_large='" + cdl + "'"
			sqlStr = sqlStr + " and code_mid='" + cdm + "'"
			sqlStr = sqlStr + " and code_small='" + cds + "'"
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			if not rsACADEMYget.Eof then
				GetNewCateCurrentPos = GetNewCateCurrentPos + "-" + db2html(rsACADEMYget("code_nm"))
			end if
			rsACADEMYget.close
		end if

	end function

end Class
%>