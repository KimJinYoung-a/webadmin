<%
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// APIs
'//

' 책 목록 얻기
Function API_GetBookList()
	API_GetBookList = GetBookListJSON(Factory.Create("Args"))
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// JSON Functions
'//

' 책 목록 얻기
Function GetBookListJSON(oArgs)
	Dim oBookListRS : Set oBookListRS = GetBookListRS(Factory.Create("Args").SetArgs(Array( _
		"pageSize", 10, _
		"pageNum", 1 _
	)))

	Dim oBookList : Set oBookList = Server.CreateObject("System.Collections.ArrayList")

	Do Until oBookListRS.EOF
		Dim sBookId : sBookId = oBookListRS("idx")
		Dim sVol : sVol = oBookListRS("vol")
		Dim sRev : sRev = oBookListRS("rev")

		Dim oBook : Set oBook = jsObject()
		oBook("bookid") = sBookId
		oBook("vol") = sVol
		oBook("rev") = sRev
		oBookList.Add oBook

		oBookListRS.MoveNext
	Loop

	GetBookListJSON = jsArray().RenderArray(oBookList.ToArray, 1, "")
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// RS Functions
'//

' 책 개수 얻기
Function GetNumBookList(oArgs)
	Dim oRS : Set oRS = GetRsFromArgs(oArgs)

	GetNumBookList = GetTableRows(oRS, "db_contents.dbo.tbl_hhiker_book")
End Function

' 책 목록 얻기
Function GetBookListRS(oTempArgs)
	Dim oArgs : Set oArgs = Factory.Create("Args").SetArgs(Array( _
		"pageSize", 20, _
		"pageNum", 1 _
	))
	oArgs.SetArgs(oTempArgs)

	Dim oRS : Set oRS = GetRsFromArgs(oArgs)

	Dim nPageSize : nPageSize = oArgs.Item("pageSize")
	Dim nPageNum : nPageNum =  IF_(oArgs.Item("pageNum") > 1, oArgs.Item("pageNum"), 1)
	Dim nFirstIndex : nFirstIndex = (nPageNum - 1) * nPageSize
	Dim nLastIndex : nLastIndex = nPageNum * nPageSize + 1

	Dim sSQL : sSQL = "SELECT * FROM (SELECT ROW_NUMBER() OVER (ORDER BY idx DESC) AS ROWINDEX, * FROM db_contents.dbo.tbl_hhiker_book) AS SUB WHERE SUB.ROWINDEX > " & nFirstIndex & " AND SUB.ROWINDEX < " & nLastIndex

	oRS.open sSQL, dbget, 0

	Set GetBookListRS = oRS
End Function


'이하 김진영 생성 20130218
Class HitchItem
	Public Fidx
	Public Fvol
	Public Frev
	Public Fopendate
	Public FopenState
	Public FmTitleName
	Public FmImgURL
	Public FmImgURL2
	Public FzipUrl
	Public FregUserID
	Public Fregdate

	Public Fmidx
	Public Fctgbnname
	Public FctSeq
	Public ForgfileName
	Public FcontURL
	Public FmusicTitle
	Public Fmusician
	Public FlinkURL
	Public Fisusing
	Public ForderNo
	Public Fdevice
	Public Fbannerimg
	Public Fusetype
	Public FclickURL
	Public Fstartdate
	Public Fenddate

	Public FNidx
	Public FNassigndevice
	Public FNstartdate
	Public FNenddate
	Public FNcontents
	Public FNregdate
	Public FNisusing
End Class

Class Hitchhiker
	Public FhitchList()
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	Public Fsearch1
	Public FsearchVol

	Public SMode
	Public Sidx
	Public Sopendate
	Public Sopenstate
	Public Svol

	Public Midx
	Public Ctseq
	Public Ctgbnname

	Public Dctgbnname
	Public DctSeq
	Public DorgfileName
	Public DcontURL
	Public DmusicTitle
	Public Dmusician
	Public DlinkURL
	Public Disusing
	Public DorderNo

	Public FNidx
	Public FNdevice
	Public FNstartdate
	Public FNenddate
	Public FNcontents
	Public FNregdate
	Public FNisusing

	Public SBidx
	Public SBbannerImg
	Public SBclickURL
	Public SBusetype
	Public SBisusing
	Public SBstartdate
	Public SBenddate

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

	Public Sub HitchList
		Dim strSQL, i, where, where2

		If Fsearch1 = "all" Then
			where = where & "" & VBCRLF
			where2 = where2 & "" & VBCRLF
		ElseIf Fsearch1 = "open" Then
			where = where & " and openState = '7' "  & VBCRLF
			where2 = where2 & " and t1.openState = '7' "  & VBCRLF
		End If

		If FsearchVol <> "" Then
			where = where & " and vol = '"&FsearchVol&"' " & VBCRLF
			where2 = where2 & " and t1.vol = '"&FsearchVol&"' " & VBCRLF
		End If

		If Fsearch1 = "lastRev" Then
			strSQL = ""
			strSQL = strSQL & " select count(*) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/"&FPageSize&") as tp " & VBCRLF
			strSQL = strSQL & " from db_contents.dbo.tbl_hhiker_book as t1, " & VBCRLF
			strSQL = strSQL & " (select vol, max(rev) as maxRev from db_contents.dbo.tbl_hhiker_book group by vol) as t2 " & VBCRLF
			strSQL = strSQL & " where t1.rev = t2.maxRev and t1.vol = t2.vol "&where2&" " & VBCRLF
			rsget.open strSQL, dbget, 1
				FTotalCount = rsget("cnt")
				FTotalpage = rsget("tp")
			rsget.close

			strSQL = ""
			strSQL = strSQL & " select top "& Cstr(FPageSize * FCurrPage) &" t1.idx, t1.vol, t1.rev, t1.opendate, t1.openState, t1.mTitleName, t1.mImgURL, t1.mImgURL2, t1.zipUrl, t1.regUserID, t1.regdate " & VBCRLF
			strSQL = strSQL & " from db_contents.dbo.tbl_hhiker_book as t1, " & VBCRLF
			strSQL = strSQL & " (select vol, max(rev) as maxRev from db_contents.dbo.tbl_hhiker_book group by vol) as t2 " & VBCRLF
			strSQL = strSQL & " where t1.rev = t2.maxRev and t1.vol = t2.vol "&where2&" " & VBCRLF
			strSQL = strSQL & " order by t1.vol desc, t1.rev desc " & VBCRLF
		Else
			strSQL = ""
			strSQL = strSQL & " select count(*) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/"&FPageSize&") as tp " & VBCRLF
			strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book "
			strSQL = strSQL & " WHERE 1 = 1 "& where &" " & VBCRLF
			rsget.open strSQL, dbget, 1
				FTotalCount = rsget("cnt")
				FTotalpage = rsget("tp")
			rsget.close

			strSQL = ""
			strSQL = strSQL & " SELECT top "& Cstr(FPageSize * FCurrPage) &" idx, vol, rev, opendate, openState, mTitleName, mImgURL, mImgURL2, zipUrl, regUserID, regdate " & VBCRLF
			strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book " & VBCRLF
			strSQL = strSQL & " WHERE 1 = 1 "& where &" " & VBCRLF
			strSQL = strSQL & " order by vol desc, rev desc"
		End If
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		Redim preserve FhitchList(FResultCount)
		FPageCount = FCurrPage - 1
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FhitchList(i) = new HitchItem
					FhitchList(i).Fidx 			= rsget("idx")
					FhitchList(i).Fvol 			= rsget("vol")
					FhitchList(i).Frev 			= rsget("rev")
					FhitchList(i).Fopendate 	= rsget("opendate")
					FhitchList(i).FopenState 	= rsget("openState")
					FhitchList(i).FmTitleName 	= rsget("mTitleName")
					FhitchList(i).FmImgURL 		= rsget("mImgURL")
					FhitchList(i).FmImgURL2 	= rsget("mImgURL2")
					FhitchList(i).FzipUrl 		= rsget("zipUrl")
					FhitchList(i).FregUserID 	= rsget("regUserID")
					FhitchList(i).Fregdate 		= rsget("regdate")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Sub

	Public Sub HitchModify
		Dim strSql, i
		strSQL = ""
		strSQL = strSQL & " SELECT opendate, openState " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book " & VBCRLF
		strSQL = strSQL & " Where idx = '"&Sidx&"' "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			Sopendate 			= rsget("opendate")
			SopenState			= rsget("openState")
		End If
		rsget.Close
	End Sub

	Public Sub HitchProcess
		Dim strSql, i
		Dim Subidx
		If sMode = "U" Then
			If Sopenstate = "7" Then
				strSQL = ""
				strSQL = strSQL & " SELECT idx, vol, rev FROM db_contents.dbo.tbl_hhiker_book " & VBCRLF
				strSQL = strSQL & " Where vol = '"&Svol&"' and openstate = '7' "
				rsget.Open strSql,dbget,1
				If not rsget.EOF Then
					Subidx = rsget("idx")
					response.write "<script>if(confirm('같은 Vol의 오픈상태가 종료로 변환됩니다\n계속하시겠습니까?')){}else{window.close();}</script>"
					strSQL = ""
					strSQL = strSQL & " Update db_contents.dbo.tbl_hhiker_book SET " & VBCRLF
					strSQL = strSQL & " openstate = '9' " & VBCRLF
					strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book " & VBCRLF
					strSQL = strSQL & " Where idx = '"&Subidx&"' "
					dbget.execute(strSQL)

					strSQL = ""
					strSQL = strSQL & " Update db_contents.dbo.tbl_hhiker_book SET " & VBCRLF
					strSQL = strSQL & " opendate = '"&Sopendate&"', openstate = '"&SopenState&"' " & VBCRLF
					strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book " & VBCRLF
					strSQL = strSQL & " Where idx = '"&Sidx&"' "
					dbget.execute(strSQL)
				Else
					strSQL = ""
					strSQL = strSQL & " Update db_contents.dbo.tbl_hhiker_book SET " & VBCRLF
					strSQL = strSQL & " opendate = '"&Sopendate&"', openstate = '"&SopenState&"' " & VBCRLF
					strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book " & VBCRLF
					strSQL = strSQL & " Where idx = '"&Sidx&"' "
					dbget.execute(strSQL)
				End If
				rsget.Close
			Else
				strSQL = ""
				strSQL = strSQL & " Update db_contents.dbo.tbl_hhiker_book SET " & VBCRLF
				strSQL = strSQL & " opendate = '"&Sopendate&"', openstate = '"&SopenState&"' " & VBCRLF
				strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book " & VBCRLF
				strSQL = strSQL & " Where idx = '"&Sidx&"' "
				dbget.execute(strSQL)
			End If
		End If
	End Sub

	Public Sub HitchDetailList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " select count(*) as cnt, CEILING(CAST(Count(midx) AS FLOAT)/"&FPageSize&") as tp " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book_detail "
		strSQL = strSQL & " WHERE midx = '"&Midx&"' " & VBCRLF
		rsget.open strSQL, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalpage = rsget("tp")
		rsget.close

		strSQL = ""
		strSQL = strSQL & " SELECT top "& Cstr(FPageSize * FCurrPage) &" midx, ctgbnname, ctSeq, orgfileName, contURL, musicTitle, musician, linkURL, isusing, orderNo, isnull(device, '') as device " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book_detail " & VBCRLF
		strSQL = strSQL & " WHERE midx = '"&Midx&"' " & VBCRLF
		strSQL = strSQL & " order by ctgbnname asc, device desc, orderNo asc"
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		Redim preserve FhitchList(FResultCount)
		FPageCount = FCurrPage - 1
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FhitchList(i) = new HitchItem
					FhitchList(i).Fmidx 		= rsget("midx")
					FhitchList(i).Fctgbnname 	= rsget("ctgbnname")
					FhitchList(i).FctSeq 		= rsget("ctSeq")
					FhitchList(i).ForgfileName 	= rsget("orgfileName")
					FhitchList(i).FcontURL 		= rsget("contURL")
					FhitchList(i).FmusicTitle 	= rsget("musicTitle")
					FhitchList(i).Fmusician 	= rsget("musician")
					FhitchList(i).FlinkURL 		= rsget("linkURL")
					FhitchList(i).Fisusing 		= rsget("isusing")
					FhitchList(i).ForderNo 		= rsget("orderNo")
					FhitchList(i).Fdevice 		= rsget("device")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Sub

	Public Sub HitchDetailView
		Dim strSQL
		strSQL = ""
		strSQL = strSQL & " SELECT ctgbnname, ctSeq, orgfileName, contURL, musicTitle, musician, linkURL, isusing, orderNo " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_book_detail " & VBCRLF
		strSQL = strSQL & " WHERE midx = '"&Midx&"' and ctSeq = '"&Ctseq&"' and ctgbnname = '"&Ctgbnname&"' "
		rsget.open strSQL, dbget, 1
			Dctgbnname 		= rsget("ctgbnname")
			DctSeq 			= rsget("ctSeq")
			DorgfileName 	= rsget("orgfileName")
			DcontURL 		= rsget("contURL")
			DmusicTitle 	= rsget("musicTitle")
			Dmusician 		= rsget("musician")
			DlinkURL 		= rsget("linkURL")
			Disusing 		= rsget("isusing")
			DorderNo 		= rsget("orderNo")
		rsget.close
	End Sub

	Public Sub HitchNoticeList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " select count(*) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/"&FPageSize&") as tp " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_notice "
		strSQL = strSQL & " WHERE 1 = 1  " & VBCRLF
		rsget.open strSQL, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalpage = rsget("tp")
		rsget.close

		strSQL = ""
		strSQL = strSQL & " SELECT top "& Cstr(FPageSize * FCurrPage) &" idx, assigndevice, startdate, enddate, contents, regdate, isusing " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_notice " & VBCRLF
		strSQL = strSQL & " WHERE 1 = 1  " & VBCRLF
		strSQL = strSQL & " order by idx desc"
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		Redim preserve FhitchList(FResultCount)
		FPageCount = FCurrPage - 1
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FhitchList(i) = new HitchItem
					FhitchList(i).FNidx 			= rsget("idx")
					FhitchList(i).FNassigndevice	= rsget("assigndevice")
					FhitchList(i).FNstartdate 		= rsget("startdate")
					FhitchList(i).FNenddate 		= rsget("enddate")
					FhitchList(i).FNcontents 		= rsget("contents")
					FhitchList(i).FNregdate 		= rsget("regdate")
					FhitchList(i).FNisusing 		= rsget("isusing")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Sub

	Public Sub FNoticeView
		Dim strSQL
		strSQL = ""
		strSQL = strSQL & " SELECT idx, assigndevice, startdate, enddate, contents, regdate, isusing " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_notice " & VBCRLF
		strSQL = strSQL & " WHERE idx = '"&Midx&"' "
		rsget.open strSQL, dbget, 1
			FNidx	 		= rsget("idx")
			FNdevice 		= rsget("assigndevice")
			FNstartdate 	= rsget("startdate")
			FNenddate 		= rsget("enddate")
			FNcontents 		= rsget("contents")
			FNregdate 		= rsget("regdate")
			FNisusing		= rsget("isusing")
		rsget.close
	End Sub

	Public Sub HitchBannereList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " select count(*) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/"&FPageSize&") as tp " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_bannerImg "
		strSQL = strSQL & " WHERE 1 = 1 " & VBCRLF
		rsget.open strSQL, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalpage = rsget("tp")
		rsget.close

		strSQL = ""
		strSQL = strSQL & " SELECT top "& Cstr(FPageSize * FCurrPage) &" idx, bannerImg, usetype, clickURL, isusing, startdate, enddate " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_bannerImg " & VBCRLF
		strSQL = strSQL & " WHERE 1 = 1  " & VBCRLF
		strSQL = strSQL & " order by idx desc"
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		Redim preserve FhitchList(FResultCount)
		FPageCount = FCurrPage - 1
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FhitchList(i) = new HitchItem
					FhitchList(i).Fidx 			= rsget("idx")
					FhitchList(i).Fusetype		= rsget("usetype")
					FhitchList(i).Fbannerimg	= rsget("bannerimg")
					FhitchList(i).FclickURL 	= rsget("clickURL")
					FhitchList(i).Fisusing 		= rsget("isusing")
					FhitchList(i).Fstartdate	= rsget("startdate")
					FhitchList(i).Fenddate 		= rsget("enddate")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Sub

	Public Sub HitchBannereView
		Dim strSQL
		strSQL = ""
		strSQL = strSQL & " SELECT idx, bannerImg, clickURL, usetype, isusing, startdate, enddate " & VBCRLF
		strSQL = strSQL & " FROM db_contents.dbo.tbl_hhiker_bannerImg " & VBCRLF
		strSQL = strSQL & " WHERE idx = '"&Midx&"' "
		rsget.open strSQL, dbget, 1
			SBidx 		= rsget("idx")
			SBbannerImg = rsget("bannerImg")
			SBclickURL 	= rsget("clickURL")
			SBusetype 	= rsget("usetype")
			SBisusing 	= rsget("isusing")
			SBstartdate	= rsget("startdate")
			SBenddate 	= rsget("enddate")
		rsget.close
	End Sub

End Class
%>