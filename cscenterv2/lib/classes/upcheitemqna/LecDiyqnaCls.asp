<%
Class CQnaItem
	Public FIdx
	Public FPagegubun
	Public FReply_group_idx
	Public FReply_depth
	Public FReply_num
	Public FUserid
	Public FReplyuserid
	Public FTitle
	Public FComment
	Public FQna
	Public FEmailok
	Public FSmsnum
	Public FSmsok
	Public FAnswerYN
	Public FIsusing
	Public FRegdate
	Public FLastRegdate
	Public FLecture_gubun
	
	public Fmakerid
	public Fitemid
	public Flec_idx
	public Flecturer_name
	public Fbrandname
	
	public Fitemname
	public Flec_title
	public Flistimage
	
	Public Function getQnaGubunName()
		Select Case FLecture_gubun
			Case "1"		getQnaGubunName = "강좌문의"
			Case "2"		getQnaGubunName = "재료문의"
			Case "3"		getQnaGubunName = "신설 강좌요청"
			Case "4"		getQnaGubunName = "수강료,재료비 문의"
			Case "5"		getQnaGubunName = "입금확인"
			Case "6"		getQnaGubunName = "강좌취소"
			Case "7"		getQnaGubunName = "강좌대기문의"
			Case "8"		getQnaGubunName = "DIY 주문문의"
			Case "9"		getQnaGubunName = "DIY 주문취소문의"
			Case "10"		getQnaGubunName = "DIY 상품문의"
			Case "11"		getQnaGubunName = "기타 문의"
		End Select
	End Function

	Public Function getAnswerName()
		Select Case FAnswerYN
			Case "Y"		getAnswerName = "완료"
			Case "N"		getAnswerName = "<font color='GREEN'>대기</font>"
		End Select
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CQna
	Public FItemList()
	Public FOneItem
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectisanswer
	Public FRectsearchDiv	''상세 문의 분야(상품)
	Public FRectsearchgubun	''문의 분야(강좌,상품)
	Public FRectsearchKey
	Public FRectsearchString
	Public FRectIdx
	Public FRectGroupIdx
	Public FRectUserid
	
	public FRectmakerid
	public FRectMakerUserid

	Public Sub getMyqnaList()
		Dim i, sqlStr, addSql

		If FRectMakerUserid <> "" Then
			addSql = addSql & " and q.makerid='" & FRectMakerUserid & "'"
		End If

		If FRectisanswer <> "" Then
			addSql = addSql & " and q.answerYN='" & FRectisanswer & "'"
		End If
		
		If FRectsearchDiv <> "" Then
			addSql = addSql & " and q.lecture_gubun='" & FRectsearchDiv & "'"
		End If

		If FRectsearchgubun <> "" Then
			addSql = addSql & " and q.pagegubun='" & FRectsearchgubun & "'"
		else
			addSql = addSql & " and q.pagegubun in ('D','L') "'D : 상품 , L : 강좌
		End If

		If FRectsearchKey <> "" Then
			If FRectsearchString <> "" Then
				If FRectsearchKey = "idx" Then
					addSql = addSql & " and q.idx = '" & FRectsearchString & "'"
				ElseIf FRectsearchKey = "title" Then
					addSql = addSql & " and q.title like '%" & FRectsearchString & "%'"
				ElseIf FRectsearchKey = "comment" Then
					addSql = addSql & " and q.comment like '%" & FRectsearchString & "%'"
				ElseIf FRectsearchKey = "titlecomment" Then
					addSql = addSql & " and (q.title like '%" & FRectsearchString & "%' OR q.comment like '%" & FRectsearchString & "%') "
				ElseIf FRectsearchKey = "regid" Then
					addSql = addSql & " and q.userid = '" & FRectsearchString & "'"
				ElseIf FRectsearchKey = "regname" Then
					addSql = addSql & " and q.username = '" & FRectsearchString & "'"
				ElseIf FRectsearchKey = "searchmakerid" Then
					addSql = addSql & " and q.makerid = '" & FRectsearchString & "'"
				ElseIf FRectsearchKey = "regmakername" Then
					addSql = addSql & " and (l.lecturer_name like '%" & FRectsearchString & "%' OR d.brandname like '%" & FRectsearchString & "%') "
				End If
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as cnt, CEILING(CAST(COUNT(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_qna_NEW] as q"

''브랜드명 검색
		sqlStr = sqlStr & " left join [db_academy].[dbo].tbl_lec_item as l "
		sqlStr = sqlStr & " on q.lec_idx=l.idx "
		sqlStr = sqlStr & " left join [db_academy].[dbo].tbl_diy_item as d "
		sqlStr = sqlStr & " on q.itemid=d.itemid "
		
		sqlStr = sqlStr & " WHERE 1 = 1 "
'		sqlStr = sqlStr & " and pagegubun in ('D','L') "	'D : 상품 , L : 강좌
		sqlStr = sqlStr & " and q.reply_depth = '0' "
		sqlStr = sqlStr & " and q.isusing = 'Y' "
		sqlStr = sqlStr & addSql
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage 	= rsACADEMYget("totPg")
		rsACADEMYget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & "	q.idx, q.pagegubun, q.reply_group_idx, q.reply_depth, q.reply_num, q.userid, q.replyuserid, q.title  "
		sqlStr = sqlStr & " ,q.comment, q.qna, q.emailok, q.smsnum, q.smsok, q.answerYN, q.isusing, q.regdate, q.lecture_gubun, q.makerid "
		sqlStr = sqlStr & " ,l.lecturer_name, d.brandname "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_qna_NEW] as q"

		sqlStr = sqlStr & " left join [db_academy].[dbo].tbl_lec_item as l "
		sqlStr = sqlStr & " on q.lec_idx=l.idx "
		sqlStr = sqlStr & " left join [db_academy].[dbo].tbl_diy_item as d "
		sqlStr = sqlStr & " on q.itemid=d.itemid "

		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and q.pagegubun in ('D','L') "	'D : 상품 , L : 강좌
		sqlStr = sqlStr & " and q.reply_depth = '0' "
		sqlStr = sqlStr & " and q.isusing = 'Y' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY q.idx DESC "
'response.write sqlStr
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new CQnaItem
					FItemList(i).FIdx				= rsACADEMYget("idx")
					
					FItemList(i).Flecturer_name		= rsACADEMYget("lecturer_name")
					FItemList(i).Fbrandname			= rsACADEMYget("brandname")
					
					FItemList(i).FPagegubun			= rsACADEMYget("pagegubun")
					FItemList(i).FReply_group_idx	= rsACADEMYget("reply_group_idx")
					FItemList(i).FReply_depth		= rsACADEMYget("reply_depth")
					FItemList(i).FReply_num			= rsACADEMYget("reply_num")
					FItemList(i).FUserid			= rsACADEMYget("userid")
					FItemList(i).FReplyuserid		= rsACADEMYget("replyuserid")
					FItemList(i).FTitle				= rsACADEMYget("title")
					FItemList(i).FComment			= rsACADEMYget("comment")
					FItemList(i).FQna				= rsACADEMYget("qna")
					FItemList(i).FEmailok			= rsACADEMYget("emailok")
					FItemList(i).FSmsnum			= rsACADEMYget("smsnum")
					FItemList(i).FSmsok				= rsACADEMYget("smsok")
					FItemList(i).FAnswerYN			= rsACADEMYget("answerYN")
					FItemList(i).FIsusing			= rsACADEMYget("isusing")
					FItemList(i).FRegdate			= rsACADEMYget("regdate")
					FItemList(i).FLecture_gubun		= rsACADEMYget("lecture_gubun")
					
					FItemList(i).Fmakerid			= rsACADEMYget("makerid")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	Public Sub getOnemyqna()
		Dim i, sqlStr, addSql
'		sqlStr = ""
'		sqlStr = sqlStr & " SELECT TOP 1 Q.makerid, Q.itemid, Q.lec_idx, Q.pagegubun, Q.answerYN, Q.lecture_gubun, Q.userid "
'		sqlStr = sqlStr & " ,T.regdate, Q.regdate as lastregdate, T.smsok, T.smsnum, Q.title "
'		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_qna_NEW] as Q "
'		sqlStr = sqlStr & " LEFT JOIN ( "
'		sqlStr = sqlStr & " 	SELECT TOP 1 regdate, idx, reply_group_idx, smsok, smsnum "
'		sqlStr = sqlStr & " 	FROM db_academy.[dbo].[tbl_academy_qna_NEW]  "
'		sqlStr = sqlStr & " 	WHERE isusing = 'Y' and reply_group_idx = '" & FRectGroupIdx & "'  "
'		sqlStr = sqlStr & " 	and qna = 'Q' "
'		sqlStr = sqlStr & " 	ORDER BY reply_num DESC  "
'		sqlStr = sqlStr & " ) as T on T.reply_group_idx = Q.reply_group_idx "
'		sqlStr = sqlStr & " WHERE Q.idx= '"&FRectIdx&"' and Q.isusing = 'Y' "
		
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 Q.makerid, Q.itemid, Q.lec_idx, Q.pagegubun, Q.answerYN, Q.lecture_gubun, Q.userid "
		sqlStr = sqlStr & " ,T.regdate, Q.regdate as lastregdate, T.smsok, T.smsnum, Q.title "
		If FRectsearchDiv = "D" Then
			sqlStr = sqlStr & " ,D.itemname, D.listimage as listimage"
		elseif FRectsearchDiv = "L" Then
			sqlStr = sqlStr & " ,L.lec_title, L.listimg as listimage"
		end if
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_qna_NEW] as Q "

		If FRectsearchDiv = "D" Then
			sqlStr = sqlStr & " join [db_academy].[dbo].tbl_diy_item as D "
			sqlStr = sqlStr & " 	on Q.itemid=D.itemid "
		elseif FRectsearchDiv = "L" Then
			sqlStr = sqlStr & " join db_academy.[dbo].[tbl_lec_item] as L "
			sqlStr = sqlStr & " on Q.lec_idx=L.idx"
		end if

		sqlStr = sqlStr & " LEFT JOIN ( "
		sqlStr = sqlStr & " 	SELECT TOP 1 regdate, idx, reply_group_idx, smsok, smsnum "
		sqlStr = sqlStr & " 	FROM db_academy.[dbo].[tbl_academy_qna_NEW]  "
		sqlStr = sqlStr & " 	WHERE isusing = 'Y' and reply_group_idx = '" & FRectGroupIdx & "'  "
		sqlStr = sqlStr & " 	and qna = 'Q' "
		sqlStr = sqlStr & " 	ORDER BY reply_num DESC  "
		sqlStr = sqlStr & " ) as T on T.reply_group_idx = Q.reply_group_idx "
		sqlStr = sqlStr & " WHERE Q.idx= '"&FRectIdx&"' and Q.isusing = 'Y' "
		
'response.write sqlStr
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		If  not rsACADEMYget.EOF  then
			Set FOneItem = new CQnaItem
				FOneItem.Fitemid		= rsACADEMYget("itemid")
				FOneItem.Flec_idx		= rsACADEMYget("lec_idx")
				FOneItem.Fmakerid		= rsACADEMYget("makerid")

				FOneItem.FAnswerYN		= rsACADEMYget("answerYN")
				FOneItem.Fpagegubun	= rsACADEMYget("pagegubun")
				FOneItem.FLecture_gubun	= rsACADEMYget("lecture_gubun")
				FOneItem.FUserid		= rsACADEMYget("userid")
				FOneItem.FRegdate		= rsACADEMYget("regdate")
				FOneItem.FLastRegdate	= rsACADEMYget("lastregdate")
				FOneItem.FSmsnum		= rsACADEMYget("smsnum")
				FOneItem.FSmsok			= rsACADEMYget("smsok")
				FOneItem.FTitle			= rsACADEMYget("title")
				
				If FRectsearchDiv = "D" Then
					FOneItem.Fitemname		= rsACADEMYget("itemname")
					FOneItem.Flistimage     = imgFingers & "/diyitem/webimage/list/" & GetImageSubFolderByItemid(FOneItem.Fitemid)&"/"&rsACADEMYget("listimage")
				elseif FRectsearchDiv = "L" Then
					FOneItem.Flec_title		= rsACADEMYget("lec_title")
					FOneItem.Flistimage     = imgFingers & "/lectureitem/list/"& GetImageSubFolderByItemid(FOneItem.Fitemid)&"/"&rsACADEMYget("listimage")
				end if
		End if
		rsACADEMYget.Close
	End Sub

	Public Sub getqnaDetailList()
		Dim i, sqlStr, addSql
		
		If FRectGroupIdx <> "" Then
			addSql = addSql & " and reply_group_idx='" & FRectGroupIdx & "'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as cnt, CEILING(CAST(COUNT(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_qna_NEW] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and pagegubun in ('D','L') "	'D : 상품 , L : 강좌
		sqlStr = sqlStr & " and isusing = 'Y' "
		sqlStr = sqlStr & addSql
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage 	= rsACADEMYget("totPg")
		rsACADEMYget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & "	idx, pagegubun, reply_group_idx, reply_depth, reply_num, userid, replyuserid, title  "
		sqlStr = sqlStr & " ,comment, qna, emailok, smsnum, smsok, answerYN, isusing, regdate, lecture_gubun "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_qna_NEW] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and pagegubun in ('D','L') "	'D : 상품 , L : 강좌
		sqlStr = sqlStr & " and isusing = 'Y' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY reply_num ASC "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new CQnaItem
					FItemList(i).FIdx				= rsACADEMYget("idx")
					FItemList(i).FPagegubun			= rsACADEMYget("pagegubun")
					FItemList(i).FReply_group_idx	= rsACADEMYget("reply_group_idx")
					FItemList(i).FReply_depth		= rsACADEMYget("reply_depth")
					FItemList(i).FReply_num			= rsACADEMYget("reply_num")
					FItemList(i).FUserid			= rsACADEMYget("userid")
					FItemList(i).FReplyuserid		= rsACADEMYget("replyuserid")
					FItemList(i).FTitle				= rsACADEMYget("title")
					FItemList(i).FComment			= rsACADEMYget("comment")
					FItemList(i).FQna				= rsACADEMYget("qna")
					FItemList(i).FEmailok			= rsACADEMYget("emailok")
					FItemList(i).FSmsnum			= rsACADEMYget("smsnum")
					FItemList(i).FSmsok				= rsACADEMYget("smsok")
					FItemList(i).FAnswerYN			= rsACADEMYget("answerYN")
					FItemList(i).FIsusing			= rsACADEMYget("isusing")
					FItemList(i).FRegdate			= rsACADEMYget("regdate")
					FItemList(i).FLecture_gubun		= rsACADEMYget("lecture_gubun")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	Public Sub getUserQnAList
		Dim i, sqlStr, addSql
		
		If FRectUserid <> "" Then
			addSql = addSql & " and userid = '"&FRectUserid&"' "
		End If
		
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & "	idx, pagegubun, reply_group_idx, reply_depth, reply_num, userid, replyuserid, title  "
		sqlStr = sqlStr & " ,comment, qna, emailok, smsnum, smsok, answerYN, isusing, regdate, lecture_gubun "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_qna_NEW] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and pagegubun in ('D','L') "	'D : 상품 , L : 강좌
		sqlStr = sqlStr & " and reply_depth = '0' "
		sqlStr = sqlStr & " and isusing = 'Y' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new CQnaItem
					FItemList(i).FIdx				= rsACADEMYget("idx")
					FItemList(i).FPagegubun			= rsACADEMYget("pagegubun")
					FItemList(i).FReply_group_idx	= rsACADEMYget("reply_group_idx")
					FItemList(i).FReply_depth		= rsACADEMYget("reply_depth")
					FItemList(i).FReply_num			= rsACADEMYget("reply_num")
					FItemList(i).FUserid			= rsACADEMYget("userid")
					FItemList(i).FReplyuserid		= rsACADEMYget("replyuserid")
					FItemList(i).FTitle				= rsACADEMYget("title")
					FItemList(i).FComment			= rsACADEMYget("comment")
					FItemList(i).FQna				= rsACADEMYget("qna")
					FItemList(i).FEmailok			= rsACADEMYget("emailok")
					FItemList(i).FSmsnum			= rsACADEMYget("smsnum")
					FItemList(i).FSmsok				= rsACADEMYget("smsok")
					FItemList(i).FAnswerYN			= rsACADEMYget("answerYN")
					FItemList(i).FIsusing			= rsACADEMYget("isusing")
					FItemList(i).FRegdate			= rsACADEMYget("regdate")
					FItemList(i).FLecture_gubun		= rsACADEMYget("lecture_gubun")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	'// 머릿말 옵션 생성 //
	Function optPrfCd(grpCd, nowCd)
		Dim SQL, strOpt
		SQL =	" Select t1.commCd, t2.commNm " &_
				" From db_academy.dbo.tbl_preface as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2  on t1.commCd=t2.commCd " &_
				" Where t1.groupCd in (" & grpCd & ") " &_
				" Group by t1.commCd, t2.commNm "
		rsACADEMYget.Open sql, dbACADEMYget, 1
		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			Do Until rsACADEMYget.EOF
				strOpt = strOpt & "<option value='" & rsACADEMYget("commCd") & "' "
	
				if nowCd=rsACADEMYget("commCd") then strOpt = strOpt & "selected"
	
				strOpt = strOpt & " >" & rsACADEMYget("commNm") & "</option>"
				rsACADEMYget.MoveNext
			Loop
		end if
		rsACADEMYget.Close
		optPrfCd = strOpt
	End Function

	'// 공통코드 옵션 생성 //
	function optCommCd(grpCd, nowCd)
		dim SQL, strOpt

		SQL =	"Select commCd, commNm From db_academy.dbo.tbl_commCd Where groupCd in (" & grpCd & ")"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			Do Until rsACADEMYget.EOF
				strOpt = strOpt & "<option value='" & rsACADEMYget("commCd") & "' "

				if nowCd=rsACADEMYget("commCd") then strOpt = strOpt & "selected"

				strOpt = strOpt & " >" & rsACADEMYget("commNm") & "</option>"
				rsACADEMYget.MoveNext
			Loop
		end if

		rsACADEMYget.Close

		optCommCd = strOpt

	end function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
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

Function getMyinfo(iUserid, byref iregname, byref iemail)
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " SELECT TOP 1 username, usermail "
	sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_n "
	sqlStr = sqlStr & " WHERE userid = '"&iUserid&"' "
	rsget.Open sqlStr, dbget, 1
	If not rsget.EOF Then
		iregname	= rsget("username")
		iemail		= rsget("usermail")
	End If
	rsget.close
End Function

Sub sendmail(mailfrom, mailto, mailtitle, mailcontent)
    If (Not IsValidEmailAddress(mailto)) Then Exit Sub
	Dim cdoConfig, cdoMessage
	Set cdoConfig = CreateObject("CDO.Configuration")
		'-> 서버 접근방법을 설정합니다
		If (application("Svr_Info")	= "Dev") then 
		    cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
		Else
		    cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
		End If
		'-> 서버 주소를 설정합니다
		If (application("Svr_Info")	= "Dev") then 
		    cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "61.252.133.2" ''"127.0.0.1"		   
		Else
			cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "110.93.128.94"
		End If
		'-> 접근할 포트번호를 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		'-> 접속시도할 제한시간을 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		'-> SMTP 접속 인증방법을 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		'-> SMTP 서버에 인증할 ID를 입력합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"
		'-> SMTP 서버에 인증할 암호를 입력합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"
		cdoConfig.Fields.Update
	Set cdoMessage = CreateObject("CDO.Message")
		Set cdoMessage.Configuration = cdoConfig
			cdoMessage.To 				= mailto
			cdoMessage.From 			= mailfrom
			cdoMessage.SubJect 	= mailtitle
			'메일 내용이 텍스트일 경우 cdoMessage.TextBody, html일 경우 cdoMessage.HTMLBody
			cdoMessage.HTMLBody	= mailcontent
			cdoMessage.BodyPart.Charset="ks_c_5601-1987"         '/// 한글을 위해선 꼭 넣어 주어야 합니다.
			cdoMessage.HTMLBodyPart.Charset="ks_c_5601-1987"     '/// 한글을 위해선 꼭 넣어 주어야 합니다.
			cdoMessage.Send
		Set cdoMessage = nothing
	Set cdoConfig = nothing
end sub

Function IsValidEmailAddress(imailAddr)
	IsValidEmailAddress = false

	If IsNULL(imailAddr) Then Exit Function
	If (imailAddr="") Then Exit Function
	
	''점 두개 발송시 오류
	If (InStr(imailAddr,"..")>0) Then Exit Function
	
	''테섭인경우. 테스트 외에 발송 금지..
	If (application("Svr_Info")	= "Dev") Then 
	    ''여기 각자 추가 할것.
	    If (imailAddr="yunirang@naver.com") Then IsValidEmailAddress=true
	   	If (imailAddr="kjy8517@naver.com") Then IsValidEmailAddress=true
		If (imailAddr="kjy8517@hanmail.net") Then IsValidEmailAddress=true
		If (imailAddr="kjy8517@10x10.co.kr") Then IsValidEmailAddress=true
	    If (imailAddr="sokangho@korea.com") Then IsValidEmailAddress=true
	    Exit Function
	End If
	IsValidEmailAddress = true
End Function

'// 로컬 디스크의 파일을 읽어 변수에 저장 //
Function ReadLocalFile(file_name, path_name)
	Dim vPath, Filecont
	Dim fso, file
	vPath = Server.MapPath (path_name) & "\"	'로컬 디렉토리를 얻는다.
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
		Set file = fso.OpenTextFile(vPath & file_name)
			Filecont = file.ReadAll
		file.close
		Set file = Nothing
	Set fso = Nothing
	ReadLocalFile = Filecont
End Function

Function getQnaContents(oGridx)
	Dim sqlStr, arrList, i, buf
	sqlStr = ""
	sqlStr = sqlStr & " SELECT reply_num, qna, comment "
	sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_qna_NEW] "
	sqlStr = sqlStr & " WHERE 1 = 1 "
	sqlStr = sqlStr & " and pagegubun in ('L','D') "			
	sqlStr = sqlStr & " and isusing = 'Y' "
	sqlStr = sqlStr & " and reply_group_idx='" & oGridx & "'"
	sqlStr = sqlStr & " ORDER BY reply_num ASC "
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		arrList = rsACADEMYget.getRows()
	rsACADEMYget.Close

	buf = ""
	buf = buf & "<table align=""center"" width=""590"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:590px; background:#f5f5f5;"">"
	For i = 0 to Ubound(arrList, 2)
		If i = 0 Then
			buf = buf & "	<tr>"
			buf = buf & "		<td width=""80"" style=""width:80px; padding:30px 0 0 12px; margin:0; border-top:1px solid #ddd; vertical-align:top;""><img src=""http://image.thefingers.co.kr/2016/mail/txt_q.png"" alt=""질문"" style=""vertical-align:top;"" /></td>"
			buf = buf & "		<td colspan=""2"" width=""510"" style=""width:510px; padding:30px 0; margin:0; font-size:22px; line-height:26px; color:#333; font-family:'맑은고딕','Malgun Gothic','돋움', dotum, sans-serif; border-top:1px solid #ddd; vertical-align:top;"">"&nl2br(arrList(2, i))&"</td>"
			buf = buf & "	</tr>"
		Else
			buf = buf & "	<tr>"
			If arrList(1, i) = "Q" Then
				buf = buf & "	<td width=""80"" style=""width:80px; padding:30px 12px 0 0; margin:0; text-align:right; vertical-align:top;""><img src=""http://image.thefingers.co.kr/2016/mail/blt_reply2.png"" alt="""" style=""vertical-align:top;"" /></td>"
				buf = buf & "	<td width=""80"" style=""width:80px; padding:30px 0; margin:0; border-top:1px solid #ddd; text-align:left; vertical-align:top;""><img src=""http://image.thefingers.co.kr/2016/mail/txt_q.png"" alt=""답변"" style=""vertical-align:top;"" /></td>"
			Else
				buf = buf & "	<td width=""80"" style=""width:80px; padding:30px 12px 0 0; margin:0; text-align:right; vertical-align:top;""><img src=""http://image.thefingers.co.kr/2016/mail/blt_reply.png"" alt="""" style=""vertical-align:top;"" /></td>"
				buf = buf & "	<td width=""80"" style=""width:80px; padding:30px 0; margin:0; border-top:1px solid #ddd; text-align:left; vertical-align:top;""><img src=""http://image.thefingers.co.kr/2016/mail/txt_a.png"" alt=""답변"" style=""vertical-align:top;"" /></td>"
			End If
			buf = buf & "		<td width=""430"" style=""width:430px; padding:30px 0; margin:0; font-size:22px; line-height:26px; color:#333; font-family:'맑은고딕','Malgun Gothic','돋움', dotum, sans-serif; border-top:1px solid #ddd; vertical-align:top;"">"&nl2br(arrList(2, i))&"</td>"
			buf = buf & "	</tr>"
		End If
	Next
	buf = buf & "</table>"
	getQnaContents = buf
End Function
%>