<%
Class contest_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fdiv
	public fuserid
	public fimgFile1
	public fimgFile2
	public fimgFile3
	public fimgFile4
	public fimgFile5
	public fopt
	public foptText
	public fimgContent
	public fregdate
	public fusername
	public fusernum
	public fsubject
	public fcontents
	public fcontest
	public fuseyn
	public fentry_sdate
	public fentry_edate
	public fvote_sdate
	public fvote_edate
	public fresult_date
	public fcode
	public fcodename
	public fpoll_idx
	public fimg_code
	public fimg_name
	public fimg_name2
	public fsortno


	public function GetOptTypeName
		Select Case fopt
			Case "1"
				GetOptTypeName = "텐바이텐"
			Case "2"
				GetOptTypeName = "각 대학사이트"
			Case "3"
				GetOptTypeName = "공모전사이트"
			Case "4"
				GetOptTypeName = "공모전포스터"
			Case "5"
				GetOptTypeName = "Me2/Tweet/OpenCa"
			Case "6"
				GetOptTypeName = "기타"
			Case Else
				GetOptTypeName = fopt
		end Select
	end function
	

End Class


Class ClsContest

	public FItemList()
	public FOneItem
	public FGubun
	public FUserID
	public FUserNum
	public FIdx
	public FDiv
	public FSubject
	public FContest
	public FUseYN
	public fentry_sdate
	public fentry_edate
	public fvote_sdate
	public fvote_edate
	public fresult_date
	public fregdate
	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	
	
	
	public sub FContestList
		Dim sqlStr, i, vSubQuery
		
		If FUseYN <> "" Then
			vSubQuery = vSubQuery & " AND m.useyn = '" & FUseYN & "' "
		End IF
		
		If FSubject <> "" Then
			vSubQuery = vSubQuery & " AND m.subject like '%" & FSubject & "%' "
		End IF

		sqlStr = "SELECT COUNT(contest) " & _
				 "		FROM [db_event].[dbo].[tbl_contest_master] AS m " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " " & _
				 "			m.contest, m.subject, m.entry_sdate, m.entry_edate, m.vote_sdate, m.vote_edate, m.result_date, m.useyn, m.regdate " & _
				 "		FROM [db_event].[dbo].[tbl_contest_master] AS m " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " " & _
				 "	ORDER BY Cast(replace(m.contest,'con','') as int) DESC "
		
		rsget.Open sqlStr, dbget ,1
		
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		rsget.PageSize= FPageSize
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do Until rsget.Eof
				set FItemList(i) = new contest_oneitem

					FItemList(i).fcontest		= rsget("contest")
					FItemList(i).fsubject		= db2html(rsget("subject"))
					FItemList(i).fentry_sdate	= rsget("entry_sdate")
					FItemList(i).fentry_edate	= rsget("entry_edate")
					FItemList(i).fvote_sdate	= rsget("vote_sdate")
					FItemList(i).fvote_edate	= rsget("vote_edate")
					FItemList(i).fresult_date	= rsget("result_date")
					FItemList(i).fuseyn			= rsget("useyn")
					FItemList(i).fregdate		= rsget("regdate")

				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end sub
	
	
	public sub FContestDetail()
		dim SqlStr,i
		SqlStr = "SELECT m.contest, m.subject, m.entry_sdate, m.entry_edate, m.vote_sdate, m.vote_edate, m.result_date, m.useyn, m.regdate " & _
				 "	FROM [db_event].[dbo].[tbl_contest_master] AS m WHERE m.contest = '" & FContest & "' "
		rsget.Open sqlStr,dbget,1
		
		set FOneItem = new contest_oneitem
		FOneItem.fcontest		= rsget("contest")
		FOneItem.fsubject		= db2html(rsget("subject"))
		FOneItem.fentry_sdate	= rsget("entry_sdate")
		FOneItem.fentry_edate	= rsget("entry_edate")
		FOneItem.fvote_sdate	= rsget("vote_sdate")
		FOneItem.fvote_edate	= rsget("vote_edate")
		FOneItem.fresult_date	= rsget("result_date")
		FOneItem.fuseyn			= rsget("useyn")
		FOneItem.fregdate		= rsget("regdate")

		rsget.Close
	End Sub
	

	public sub FFinalList()
		dim SqlStr,i
		SqlStr = "SELECT c.idx, c.userid, c.subject, c.contents, u.username FROM [db_event].[dbo].[tbl_contest_poll] AS c " & _
				 "		Left JOIN [db_user].[dbo].[tbl_user_n] AS u ON c.userid = u.userid " & _
				 "	WHERE c.contest = '" & FDiv & "' " & _
				 "		GROUP BY c.idx, c.userid, c.subject, c.contents, u.username " & _
				 "	ORDER BY c.idx "
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new contest_oneitem
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fusernum = rsget("idx")
				FItemList(i).fusername = rsget("username")
				FItemList(i).fsubject = db2html(rsget("subject"))
				FItemList(i).fcontents = db2html(rsget("contents"))
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	End Sub
	
	
	public sub fevt_ContestEdit()
		dim SqlStr,i
		SqlStr = "SELECT idx, userid, subject, contents FROM [db_event].[dbo].[tbl_contest_poll] WHERE contest = '" & FDiv & "' AND idx = '" & FUserNum & "'"
		rsget.Open sqlStr,dbget,1
		
		set FOneItem = new contest_oneitem
		FOneItem.fusernum = rsget("idx")
		FOneItem.fuserid = rsget("userid")
		FOneItem.fsubject = db2html(rsget("subject"))
		FOneItem.fcontents = db2html(rsget("contents"))

		rsget.Close
	End Sub
	
	
	public sub FEntryList
		Dim sqlStr, i, vSubQuery
		
		If FDiv <> "" Then
			vSubQuery = vSubQuery & " AND C.div = '" & FDiv & "' "
		End IF
		
		If FUserID <> "" Then
			vSubQuery = vSubQuery & " AND C.userid = '" & FUserID & "' "
		End IF
		
		sqlStr = "SELECT COUNT(idx) " & _
				 "		FROM [db_event].[dbo].[tbl_contest_entry] AS C " & _
				 "	Inner JOIN [db_event].[dbo].[tbl_contest_master] AS M ON C.div = M.contest " & _
				 "	Left JOIN [db_user].[dbo].[tbl_user_n] AS U ON C.userid = U.userid " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " " & _
				 "			C.idx, C.div, C.userid, C.imgFile1, C.imgFile2, C.imgFile3, C.imgFile4, C.imgFile5, C.opt, C.optText, C.imgContent, C.regdate " & _
				 "			, U.username, M.subject " & _
				 "		FROM [db_event].[dbo].[tbl_contest_entry] AS C " & _
				 "	Inner JOIN [db_event].[dbo].[tbl_contest_master] AS M ON C.div = M.contest " & _
				 "	Left JOIN [db_user].[dbo].[tbl_user_n] AS U ON C.userid = U.userid " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " " & _
				 "	ORDER BY C.idx DESC "
		
		rsget.Open sqlStr, dbget ,1
		
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		rsget.PageSize= FPageSize
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do Until rsget.Eof
				set FItemList(i) = new contest_oneitem

					FItemList(i).fidx			= rsget("idx")
					FItemList(i).fdiv			= rsget("div")
					FItemList(i).fsubject		= db2html(rsget("subject"))
					FItemList(i).fuserid		= rsget("userid")
					FItemList(i).fimgFile1		= rsget("imgFile1")
					FItemList(i).fimgFile2		= rsget("imgFile2")
					FItemList(i).fimgFile3		= rsget("imgFile3")
					FItemList(i).fimgFile4		= rsget("imgFile4")
					FItemList(i).fimgFile5		= rsget("imgFile5")
					FItemList(i).fopt			= rsget("opt")
					FItemList(i).foptText		= db2html(rsget("optText"))
					FItemList(i).fimgContent	= db2html(rsget("imgContent"))
					FItemList(i).fregdate		= rsget("regdate")
					FItemList(i).fusername		= rsget("username")
					If isNull(rsget("username")) Then
						FItemList(i).fusername = "데이터없음"
					End IF

				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end sub
	
	
	public sub FPollImageList()
		dim SqlStr,i
		SqlStr = "SELECT i.idx, i.poll_idx, i.contest, i.img_code, i.img_name, m.code_name FROM [db_event].[dbo].[tbl_contest_poll_image] AS i " & _
				 "		Inner JOIN [db_event].[dbo].[tbl_contest_code_master] AS m ON i.img_code = m.code " & _
				 "	WHERE i.contest = '" & FContest & "' AND i.poll_idx = '" & FUserNum & "' " & _
				 "	ORDER BY i.img_code ASC, i.sortno DESC, i.idx ASC"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new contest_oneitem
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fpoll_idx = rsget("poll_idx")
				FItemList(i).fcontest = rsget("contest")
				FItemList(i).fimg_code = rsget("img_code")
				FItemList(i).fimg_name = rsget("img_name")
				FItemList(i).fcodename = rsget("code_name")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	End Sub
	
	
	public sub FPollImageDetail()
		dim SqlStr,i
		SqlStr = "SELECT i.idx, i.poll_idx, i.contest, i.img_code, i.img_name, i.img_name2, i.sortno FROM [db_event].[dbo].[tbl_contest_poll_image] as i WHERE i.idx = '" & FIdx & "' AND i.contest = '" & FContest & "' AND i.poll_idx = '" & FUserNum & "'"
		rsget.Open sqlStr,dbget,1
		
		If Not rsget.Eof Then
			set FOneItem = new contest_oneitem
				FOneItem.fidx = rsget("idx")
				FOneItem.fpoll_idx = rsget("poll_idx")
				FOneItem.fcontest = rsget("contest")
				FOneItem.fimg_code = rsget("img_code")
				FOneItem.fimg_name = rsget("img_name")
				FOneItem.fimg_name2 = rsget("img_name2")
				FOneItem.fsortno	= rsget("sortno")
		End If

		rsget.Close
	End Sub
	
	
	
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class


Function FImageCodeList(vImgCode, vOnChange)
	dim SqlStr,i, vBody
	SqlStr = "SELECT m.code, m.code_name FROM [db_event].[dbo].[tbl_contest_code_master] AS m " & _
			 "	WHERE m.useyn = 'y' " & _
			 "	ORDER BY m.code "
	rsget.Open sqlStr,dbget,1

	vBody = "<select name='code' " & vOnChange & ">"
	vBody = vBody & "<option value=''>-이미지구분-</option>"
	Do Until rsget.Eof
		vBody = vBody & "<option value='" & rsget("code") & "'"
		If CStr(vImgCode) = CStr(rsget("code")) Then
			vBody = vBody & " selected"
		End IF
		vBody = vBody & ">" & rsget("code_name") & "</option>"
	rsget.moveNext
	loop
	rsget.Close
	vBody = vBody & "</select>"
	
	FImageCodeList = vBody
End Function
%>