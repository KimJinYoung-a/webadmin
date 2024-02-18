<%
Class cQnaSummaryItem
	Public FID
	Public FItemID
	Public FUserid
	Public Fmakerid
	Public FCdl
	Public Fusername
	Public FContents
	Public FTitle
	Public FQadiv
	Public FReplyuser
	Public FReplycontents
	Public FReplytitle
	Public Fregdate
	Public FBrandName
	Public FReplydate
	Public Fdeliverytype
	Public FItemDiv
	Public FSecretYN
	Public FCSusername

	Public Function GetDeliveryTypeName()
		If Fdeliverytype="1" or Fdeliverytype="3" or Fdeliverytype="4" Then
			GetDeliveryTypeName = "10x10"
		Else
			GetDeliveryTypeName = "업체"
		End If
	End Function

	Public Function GetDeliveryTypeColor()
		If Fdeliverytype = "1" or Fdeliverytype = "3" or Fdeliverytype = "4" Then
			GetDeliveryTypeColor = "#000000"
		Else
			GetDeliveryTypeColor = "#CC3333"
		End if
	End Function

	Public Function GetItemDivNameName()
		If FItemDiv = "90" Then
			GetItemDivNameName = "강좌"
		Else
			GetItemDivNameName = " "
		End If
	End Function

End Class

Class cQnaSummary
	Public fTENDB
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectSdate
	Public FRectEdate
	Public FRectSType
	Public FRectDayMode
	Public FRectMakerid
	Public FRectTypeVal
	Public FRectTopCnt

	Public FRectItemID
	Public FRectCateCode
	Public FRectUserID
	Public FRectstartdate
	Public FRectenddate
	Public FRectOnlyTenBeasong
	Public FRectDPlusDay
	Public FReckMiFinish
	Public FRectsecretYN
	Public FRectSearchType
	Public FRectSearchString

	Public FTotalQnaAllCnt
	Public FTotalQnasecretYCnt
	Public FTotalQnasecretNCnt
	Public FTotalreplyYCnt
	Public FTotalreplyNCnt
	Public FTotalSumReplyDayCnt
	Public FTotalsnssend1Cnt
	Public FTotalsnssend2Cnt
	Public FTotalsnssend3Cnt
	Public FTotalsnssend4Cnt
	Public FTotalsnssend5Cnt

	Public FTermQnaAllCnt
	Public FTermQnasecretYCnt
	Public FTermQnasecretNCnt
	Public FTermreplyYCnt
	Public FTermreplyNCnt
	Public FTermSumReplyDayCnt
	Public FTermsnssend1Cnt
	Public FTermsnssend2Cnt
	Public FTermsnssend3Cnt
	Public FTermsnssend4Cnt
	Public FTermsnssend5Cnt

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

	Private Sub Class_Initialize()
		IF application("Svr_Info")="Dev" THEN
			fTENDB = "TENDB."
		End if

		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	FUNCTION last_day(year,month)
		Dim temp
		temp = CDate(LEFT(dateadd("m",1,year &"-"& month &"-01"),7) &"-01")-1
		last_day = Split(temp,"-")(2)
	END FUNCTION

	Public Function fnQnaSummayReport
		Dim strSql, addSql, addSql2, i

		If FRectSdate <> "" AND FRectEdate <> "" Then
			addSql = addSql & " and convert(varchar(10), p.regdate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), p.regdate, 120) <= '" & (FRectEdate) & "' "
		End If

		'snsSendTBL에 문의글 기준 알람횟수 등록
		strSql = ""
		strSql = strSql & " SELECT qna_idx, count(*) as snsCnt "
		strSql = strSql & " INTO #snsSendTBL "
		strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].[tbl_company_upcheqna_alarm_log] "
		strSql = strSql & " GROUP BY qna_idx "
		strSql = strSql & " CREATE NONCLUSTERED INDEX [IX_snsSendTBL] ON #snsSendTBL( qna_idx ASC ) "
		db3_dbget.Execute strSql
'rw strSql
		'누적 전체 건 조회
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as QnaAllCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
		strSql = strSql & " ,isnull(Sum(Datediff(d, p.regdate, p.replydate)), 0) as sumReplyDayCnt  "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
		strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p  "
		strSql = strSql & " LEFT JOIN #snsSendTBL as s on p.id = s.qna_idx "
		strSql = strSql & " WHERE p.isusing = 'Y' "
		strSql = strSql & " and isnull(p.secretYN, '') <> '' "
		strSql = strSql & " and p.id >= 400000 "
		db3_rsget.Open strSql, db3_dbget, 1
			FTotalQnaAllCnt			= db3_rsget("QnaAllCnt")
			FTotalQnasecretYCnt		= db3_rsget("QnasecretYCnt")
			FTotalQnasecretNCnt		= db3_rsget("QnasecretNCnt")
			FTotalreplyYCnt			= db3_rsget("replyYCnt")
			FTotalreplyNCnt			= db3_rsget("replyNCnt")
			FTotalSumReplyDayCnt	= db3_rsget("sumReplyDayCnt")
			FTotalsnssend1Cnt		= db3_rsget("snssend1Cnt")
			FTotalsnssend2Cnt		= db3_rsget("snssend2Cnt")
			FTotalsnssend3Cnt		= db3_rsget("snssend3Cnt")
			FTotalsnssend4Cnt		= db3_rsget("snssend4Cnt")
			FTotalsnssend5Cnt		= db3_rsget("snssend5Cnt")
		db3_rsget.Close
'rw strSql
		'기간 전체 건 조회
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as QnaAllCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
		strSql = strSql & " ,isnull(Sum(Datediff(d, p.regdate, p.replydate)), 0) as sumReplyDayCnt  "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
		strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p  "
		strSql = strSql & " LEFT JOIN #snsSendTBL as s on p.id = s.qna_idx "
		strSql = strSql & " WHERE p.isusing = 'Y' "
		strSql = strSql & " and isnull(p.secretYN, '') <> '' "
		strSql = strSql & " and p.id >= 400000 "
		strSql = strSql & addSql
		db3_rsget.Open strSql, db3_dbget, 1
			FTermQnaAllCnt			= db3_rsget("QnaAllCnt")
			FTermQnasecretYCnt		= db3_rsget("QnasecretYCnt")
			FTermQnasecretNCnt		= db3_rsget("QnasecretNCnt")
			FTermreplyYCnt			= db3_rsget("replyYCnt")
			FTermreplyNCnt			= db3_rsget("replyNCnt")
			FTermSumReplyDayCnt		= db3_rsget("sumReplyDayCnt")
			FTermsnssend1Cnt		= db3_rsget("snssend1Cnt")
			FTermsnssend2Cnt		= db3_rsget("snssend2Cnt")
			FTermsnssend3Cnt		= db3_rsget("snssend3Cnt")
			FTermsnssend4Cnt		= db3_rsget("snssend4Cnt")
			FTermsnssend5Cnt		= db3_rsget("snssend5Cnt")
		db3_rsget.Close
'rw strSql
		If FRectSType = "category" Then
			'기간 카테고리 그룹바이
			strSql = ""
			strSql = strSql & " SELECT ct.catename, COUNT(*) as QnaAllCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
			strSql = strSql & " ,isnull(Sum(Datediff(d, p.regdate, p.replydate)), 0) as sumReplyDayCnt  "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
			strSql = strSql & " , LEFT(ct.catecode,3) "
			strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p  "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_display_cate_item as ci on p.itemid = ci.itemid and ci.isDefault = 'Y' "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_display_cate as ct on LEFT(ci.catecode,3) = ct.catecode and ct.useyn = 'Y' "
			strSql = strSql & " LEFT JOIN #snsSendTBL as s on p.id = s.qna_idx "
			strSql = strSql & " WHERE p.isusing = 'Y' "
			strSql = strSql & " and isnull(p.secretYN, '') <> '' "
			strSql = strSql & " and p.id >= 400000 "
			strSql = strSql & addSql
			strSql = strSql & " GROUP BY LEFT(ct.catecode,3), ct.catename "
			strSql = strSql & " ORDER BY LEFT(ct.catecode,3)  "
			db3_rsget.Open strSql, db3_dbget, 1
		    If not db3_rsget.EOF Then
				fnQnaSummayReport = db3_rsget.getRows()
		    End If
			db3_rsget.Close
		Else
			If FRectMakerid <> "" Then
				addSql = addSql & " and i.makerid = '"&FRectMakerid&"' "
			End If

			'기간 브랜드 그룹바이
			strSql = ""
			If FRectTopCnt = "Y" Then
				strSql = strSql & " SELECT TOP 200 "
			Else
				strSql = strSql & " SELECT TOP 20 "
			End If
			strSql = strSql & " i.makerid, COUNT(*) as QnaAllCnt,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
			strSql = strSql & " ,isnull(Sum(Datediff(d, p.regdate, p.replydate)), 0) as sumReplyDayCnt  "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
			strSql = strSql & " ,i.makerid "
			strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_item as i on p.itemid = i.itemid "
			strSql = strSql & " LEFT JOIN #snsSendTBL as s on p.id = s.qna_idx "
			strSql = strSql & " WHERE p.isusing = 'Y' "
			strSql = strSql & " and isnull(p.secretYN, '') <> '' "
			strSql = strSql & " and p.id >= 400000 "
			strSql = strSql & addSql
			strSql = strSql & " GROUP BY i.makerid "
			strSql = strSql & " ORDER BY Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End) DESC  "
			db3_rsget.Open strSql, db3_dbget, 1
		    If not db3_rsget.EOF Then
				fnQnaSummayReport = db3_rsget.getRows()
		    End If
			db3_rsget.Close
		End If
'rw strSql
	End Function

	Public Function fnQnaSummayReportByTerm
		Dim strSql, addSql, addSql2, i
		Dim tmpSYear, tmpSMonth, tmpSDay, tmpSdate, dayModeStr

		If FRectSdate <> "" AND FRectEdate <> "" Then
			If FRectDayMode = "D" Then					'일별
				dayModeStr = "convert(varchar(10), p.regdate,21)"
				addSql = addSql & " and convert(varchar(10), p.regdate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), p.regdate, 120) <= '" & (FRectEdate) & "' "
			ElseIf FRectDayMode = "M" Then
				dayModeStr = "convert(varchar(7), p.regdate,21)"
				tmpSYear = LEFT(FRectEdate, 4)
				tmpSMonth = Split(FRectEdate,"-")(1)	'월별
				tmpSDay = last_day(tmpSYear, tmpSMonth)
				tmpSdate = tmpSYear&"-"&tmpSMonth&"-"&tmpSDay

				addSql = addSql & " and convert(varchar(10), p.regdate, 120) >= '" & LEFT(FRectSdate, 7)&"-01' and convert(varchar(10), p.regdate, 120) <= '" & tmpSdate & "' "
			ElseIf FRectDayMode = "Y" Then				'연별
				dayModeStr = "convert(varchar(4), p.regdate,21)"
				addSql = addSql & " and convert(varchar(10), p.regdate, 120) >= '" & LEFT(FRectSdate, 4)&"-01-01' and convert(varchar(10), p.regdate, 120) <= '" & LEFT(FRectEdate, 4)&"-12-31' "
			End If
		End If

		strSql = ""
		strSql = strSql & " SELECT qna_idx, count(*) as snsCnt "
		strSql = strSql & " INTO #snsSendTermTBL "
		strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].[tbl_company_upcheqna_alarm_log] "
		strSql = strSql & " GROUP BY qna_idx "
		strSql = strSql & " CREATE NONCLUSTERED INDEX [IX_snsSendTermTBL] ON #snsSendTermTBL( qna_idx ASC ) "
		db3_dbget.Execute strSql

		'기간 전체 건 조회
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as QnaAllCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
		strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
		strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p  "
		strSql = strSql & " LEFT JOIN #snsSendTermTBL as s on p.id = s.qna_idx "
		strSql = strSql & " WHERE p.isusing = 'Y' "
		strSql = strSql & " and isnull(p.secretYN, '') <> '' "
		strSql = strSql & " and p.id >= 400000 "
		strSql = strSql & addSql
		db3_rsget.Open strSql, db3_dbget, 1
			FTermQnaAllCnt			= db3_rsget("QnaAllCnt")
			FTermQnasecretYCnt		= db3_rsget("QnasecretYCnt")
			FTermQnasecretNCnt		= db3_rsget("QnasecretNCnt")
			FTermreplyYCnt			= db3_rsget("replyYCnt")
			FTermreplyNCnt			= db3_rsget("replyNCnt")
			FTermsnssend1Cnt		= db3_rsget("snssend1Cnt")
			FTermsnssend2Cnt		= db3_rsget("snssend2Cnt")
			FTermsnssend3Cnt		= db3_rsget("snssend3Cnt")
			FTermsnssend4Cnt		= db3_rsget("snssend4Cnt")
			FTermsnssend5Cnt		= db3_rsget("snssend5Cnt")
		db3_rsget.Close

		If FRectSType = "category" Then
			'기간 그룹바이
			strSql = ""
			strSql = strSql & " SELECT " & dayModeStr
			strSql = strSql & " , COUNT(*) as QnaAllCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
			strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p  "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_display_cate_item as ci on p.itemid = ci.itemid and ci.isDefault = 'Y' "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_display_cate as ct on LEFT(ci.catecode,3) = ct.catecode and ct.useyn = 'Y' "
			strSql = strSql & " LEFT JOIN #snsSendTermTBL as s on p.id = s.qna_idx "
			strSql = strSql & " WHERE p.isusing = 'Y' "
			strSql = strSql & " and isnull(p.secretYN, '') <> '' "
			strSql = strSql & " and p.id >= 400000 "
			strSql = strSql & addSql
			strSql = strSql & " GROUP BY " & dayModeStr
			strSql = strSql & " ORDER BY " & dayModeStr
		Else
			If FRectMakerid <> "" Then
				addSql = addSql & " and i.makerid = '"&FRectMakerid&"' "
			End If
			'기간 브랜드 그룹바이
			strSql = ""
			strSql = strSql & " SELECT "& dayModeStr
			strSql = strSql & " , COUNT(*) as QnaAllCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
			strSql = strSql & " ,isnull(Sum(Datediff(d, p.regdate, p.replydate)), 0) as sumReplyDayCnt  "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
			strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_item as i on p.itemid = i.itemid "
			strSql = strSql & " LEFT JOIN #snsSendTermTBL as s on p.id = s.qna_idx "
			strSql = strSql & " WHERE p.isusing = 'Y' "
			strSql = strSql & " and isnull(p.secretYN, '') <> '' "
			strSql = strSql & " and p.id >= 400000 "
			strSql = strSql & addSql
			strSql = strSql & " GROUP BY " & dayModeStr
			strSql = strSql & " ORDER BY " & dayModeStr
		End If
		db3_rsget.Open strSql, db3_dbget, 1
	    If not db3_rsget.EOF Then
			fnQnaSummayReportByTerm = db3_rsget.getRows()
	    End If
		db3_rsget.Close
	End Function

	Public Function fnQnaSummayReportByType
		Dim strSql, addSql, addSql2, i
		Dim tmpSYear, tmpSMonth, tmpSDay, tmpSdate, dayModeStr

		If FRectSdate <> "" AND FRectEdate <> "" Then
			If FRectDayMode = "D" Then					'일별
				dayModeStr = "convert(varchar(10), p.regdate,21)"
				addSql = addSql & " and convert(varchar(10), p.regdate, 120) >= '" & (FRectSdate) & "' and convert(varchar(10), p.regdate, 120) <= '" & (FRectEdate) & "' "
			ElseIf FRectDayMode = "M" Then
				dayModeStr = "convert(varchar(7), p.regdate,21)"
				tmpSYear = LEFT(FRectEdate, 4)
				tmpSMonth = Split(FRectEdate,"-")(1)	'월별
				tmpSDay = last_day(tmpSYear, tmpSMonth)
				tmpSdate = tmpSYear&"-"&tmpSMonth&"-"&tmpSDay

				addSql = addSql & " and convert(varchar(10), p.regdate, 120) >= '" & LEFT(FRectSdate, 7)&"-01' and convert(varchar(10), p.regdate, 120) <= '" & tmpSdate & "' "
			ElseIf FRectDayMode = "Y" Then				'연별
				dayModeStr = "convert(varchar(4), p.regdate,21)"
				addSql = addSql & " and convert(varchar(10), p.regdate, 120) >= '" & LEFT(FRectSdate, 4)&"-01-01' and convert(varchar(10), p.regdate, 120) <= '" & LEFT(FRectEdate, 4)&"-12-31' "
			End If
		End If

		strSql = ""
		strSql = strSql & " SELECT qna_idx, count(*) as snsCnt "
		strSql = strSql & " INTO #snsSendTermTBL "
		strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].[tbl_company_upcheqna_alarm_log] "
		strSql = strSql & " GROUP BY qna_idx "
		strSql = strSql & " CREATE NONCLUSTERED INDEX [IX_snsSendTermTBL] ON #snsSendTermTBL( qna_idx ASC ) "
		db3_dbget.Execute strSql

		If FRectSType = "category" Then
			'기간 전체 건 조회
			strSql = ""
			strSql = strSql & " SELECT COUNT(*) as QnaAllCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
			strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p  "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_display_cate_item as ci on p.itemid = ci.itemid and ci.isDefault = 'Y' "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_display_cate as ct on LEFT(ci.catecode,3) = ct.catecode and ct.useyn = 'Y' "
			strSql = strSql & " LEFT JOIN #snsSendTermTBL as s on p.id = s.qna_idx "
			strSql = strSql & " WHERE p.isusing = 'Y' "
			strSql = strSql & " and isnull(p.secretYN, '') <> '' "
			strSql = strSql & " and p.id >= 400000 "
			strSql = strSql & addSql
			strSql = strSql & " and LEFT(ci.catecode,3) = '"&FRectTypeVal&"' "
			db3_rsget.Open strSql, db3_dbget, 1
				FTermQnaAllCnt			= db3_rsget("QnaAllCnt")
				FTermQnasecretYCnt		= db3_rsget("QnasecretYCnt")
				FTermQnasecretNCnt		= db3_rsget("QnasecretNCnt")
				FTermreplyYCnt			= db3_rsget("replyYCnt")
				FTermreplyNCnt			= db3_rsget("replyNCnt")
				FTermsnssend1Cnt		= db3_rsget("snssend1Cnt")
				FTermsnssend2Cnt		= db3_rsget("snssend2Cnt")
				FTermsnssend3Cnt		= db3_rsget("snssend3Cnt")
				FTermsnssend4Cnt		= db3_rsget("snssend4Cnt")
				FTermsnssend5Cnt		= db3_rsget("snssend5Cnt")
			db3_rsget.Close
		Else
			strSql = ""
			strSql = strSql & " SELECT COUNT(*) as QnaAllCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
			strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_item as i on p.itemid = i.itemid "
			strSql = strSql & " LEFT JOIN #snsSendTermTBL as s on p.id = s.qna_idx "
			strSql = strSql & " WHERE p.isusing = 'Y' "
			strSql = strSql & " and isnull(p.secretYN, '') <> '' "
			strSql = strSql & " and p.id >= 400000 "
			strSql = strSql & addSql
			strSql = strSql & " and i.makerid = '"&FRectTypeVal&"' "
			db3_rsget.Open strSql, db3_dbget, 1
				FTermQnaAllCnt			= db3_rsget("QnaAllCnt")
				FTermQnasecretYCnt		= db3_rsget("QnasecretYCnt")
				FTermQnasecretNCnt		= db3_rsget("QnasecretNCnt")
				FTermreplyYCnt			= db3_rsget("replyYCnt")
				FTermreplyNCnt			= db3_rsget("replyNCnt")
				FTermsnssend1Cnt		= db3_rsget("snssend1Cnt")
				FTermsnssend2Cnt		= db3_rsget("snssend2Cnt")
				FTermsnssend3Cnt		= db3_rsget("snssend3Cnt")
				FTermsnssend4Cnt		= db3_rsget("snssend4Cnt")
				FTermsnssend5Cnt		= db3_rsget("snssend5Cnt")
			db3_rsget.Close
		End If

		If FRectSType = "category" Then
			'기간 그룹바이
			strSql = ""
			strSql = strSql & " SELECT " & dayModeStr
			strSql = strSql & " , COUNT(*) as QnaAllCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
			strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p  "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_display_cate_item as ci on p.itemid = ci.itemid and ci.isDefault = 'Y' "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_display_cate as ct on LEFT(ci.catecode,3) = ct.catecode and ct.useyn = 'Y' "
			strSql = strSql & " LEFT JOIN #snsSendTermTBL as s on p.id = s.qna_idx "
			strSql = strSql & " WHERE p.isusing = 'Y' "
			strSql = strSql & " and isnull(p.secretYN, '') <> '' "
			strSql = strSql & " and p.id >= 400000 "
			strSql = strSql & addSql
			strSql = strSql & " and LEFT(ci.catecode,3) = '"&FRectTypeVal&"' "
			strSql = strSql & " GROUP BY " & dayModeStr
			strSql = strSql & " ORDER BY " & dayModeStr
		Else
			'기간 브랜드 그룹바이
			strSql = ""
			strSql = strSql & " SELECT "& dayModeStr
			strSql = strSql & " , COUNT(*) as QnaAllCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'Y' Then 1 Else 0 End), 0) as QnasecretYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN p.secretYN = 'N' Then 1 Else 0 End), 0) as QnasecretNCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') <> '' Then 1 Else 0 End), 0) as replyYCnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN isnull(convert(varchar(23), p.replydate), '') = '' Then 1 Else 0 End), 0) as replyNCnt "
			strSql = strSql & " ,isnull(Sum(Datediff(d, p.regdate, p.replydate)), 0) as sumReplyDayCnt  "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 1 Then 1 Else 0 End), 0) as snssend1Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 1 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend2Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 2 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend3Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt > 3 AND s.snsCnt <= 5 Then 1 Else 0 End), 0) as snssend4Cnt "
			strSql = strSql & " ,isnull(Sum(Case WHEN s.snsCnt >= 5 Then 1 Else 0 End), 0) as snssend5Cnt "
			strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p "
			strSql = strSql & " JOIN "&fTENDB&"db_item.dbo.tbl_item as i on p.itemid = i.itemid "
			strSql = strSql & " LEFT JOIN #snsSendTermTBL as s on p.id = s.qna_idx "
			strSql = strSql & " WHERE p.isusing = 'Y' "
			strSql = strSql & " and isnull(p.secretYN, '') <> '' "
			strSql = strSql & " and p.id >= 400000 "
			strSql = strSql & addSql
			strSql = strSql & " and i.makerid = '"&FRectTypeVal&"' "
			strSql = strSql & " GROUP BY " & dayModeStr
			strSql = strSql & " ORDER BY " & dayModeStr
		End If
		db3_rsget.Open strSql, db3_dbget, 1
	    If not db3_rsget.EOF Then
			fnQnaSummayReportByType = db3_rsget.getRows()
	    End If
		db3_rsget.Close
	End Function

	Public Sub ItemQnaList()
		Dim strSql, i, addSql, CateCodeLength

		If application("Svr_Info") <> "Dev" and FRectItemID = "" Then
			addSql = addSql & " and p.id >= 400000 "
		End If

		If FRectCateCode <> "" Then
			CateCodeLength = Len(FRectCateCode)
			addSql = addSql & " and LEFT(ci.catecode,"&CateCodeLength&") = '"&FRectCateCode&"' "
		End If

		if FRectuserid<>"" then
			addSql = addSql & " and p.userid = '" + Cstr(FRectuserid) + "'" + vbcrlf
		end if

		if frectstartdate <> "" and frectenddate <> "" then
			addSql = addSql & " and p.regdate between '"+ Cstr(frectstartdate) +"' and '"+ Cstr(dateadd("d",1,frectenddate)) +"'" + vbcrlf
		end if

		if FRectItemID<>"" then
			addSql = addSql & " and p.itemid = " + Cstr(FRectItemID) + "" + vbcrlf
		end if

		if FRectMakerid<>"" then
			addSql = addSql & " and i.makerid = '" + Cstr(FRectMakerid) + "'" + vbcrlf
		end if

		if FRectOnlyTenBeasong<>"" then
			if (FRectOnlyTenBeasong = "Y") then
				addSql = addSql & " and i.deliverytype in ('1', '4')" + vbcrlf
			elseif (FRectOnlyTenBeasong = "N") then
				addSql = addSql & " and i.deliverytype not in ('1', '4')" + vbcrlf
			end if
		end if

		if (FRectDPlusDay <> "") then
			'// D+3 일 초과만
			addSql = addSql & " and DateDiff(d, p.regdate, getdate()) >= 3 " + vbcrlf
		end if

		if FReckMiFinish<>"" then
			if (FReckMiFinish = "N") then
				addSql = addSql & " and p.replydate is null" + vbcrlf
			elseif (FReckMiFinish = "Y") then
				addSql = addSql & " and p.replydate is not null" + vbcrlf
			end if
		end if

		If FRectsecretYN <> "" Then
			addSql = addSql & " and p.secretYN='" + CStr(FRectsecretYN) + "'" + vbcrlf
		End If 

		If FRectSearchType <> "" AND FRectSearchString <> "" Then
			Select Case FRectSearchType
				Case "itemid"			addSql = addSql & " and i.itemid='" & FRectSearchString & "'" + vbcrlf
				Case "qnaContent"		addSql = addSql & " and p.contents like '%" & FRectSearchString & "%'" + vbcrlf
				Case "replyContent"		addSql = addSql & " and p.replycontents like '%" & FRectSearchString & "%'" + vbcrlf
			End Select
		End If

		strSql = ""
		strSql = strSql & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p " & VBCRLF
		strSql = strSql & " LEFT JOIN "&fTENDB&"[db_item].[dbo].tbl_item i on p.itemid=i.itemid " & VBCRLF
		strSql = strSql & " LEFT JOIN "&fTENDB&"[db_item].[dbo].tbl_display_cate_item ci on i.itemid=ci.itemid and ci.isDefault = 'Y' " & VBCRLF
		strSql = strSql & " WHERE 1 = 1 "
		strSql = strSql & addSql
		strSql = strSql & " and p.itemid <> 0 "
		strSql = strSql & " and p.isusing = 'Y' "
		strSql = strSql & " and isnull(p.secretYN, '') <> '' "
		db3_rsget.Open strSql,db3_dbget, 1
			FTotalCount = db3_rsget("cnt")
			FTotalPage = db3_rsget("totPg")
		db3_rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & CStr(FPageSize*FCurrpage) & " p.id,p.userid,p.itemid,i.makerid,p.username,p.cdl,p.contents,"
		strSql = strSql & " p.qadiv, p.replyuser,p.replycontents,p.regdate, p.brandname, p.replydate, i.deliverytype, i.itemdiv, p.secretYN "
		strSql = strSql & " ,(SELECT TOP 1 ut.username FROM "&fTENDB&"[db_partner].[dbo].tbl_user_tenbyten ut WHERE p.replyuser = ut.userid and ut.part_sn = 10 and isnull(p.replyuser, '') <> '') as username "
		strSql = strSql & " FROM "&fTENDB&"[db_cs].[dbo].tbl_my_item_qna p "
		strSql = strSql & " LEFT JOIN "&fTENDB&"[db_item].[dbo].tbl_item i on p.itemid=i.itemid "
		strSql = strSql & " LEFT JOIN "&fTENDB&"[db_item].[dbo].tbl_display_cate_item ci on i.itemid=ci.itemid and ci.isDefault = 'Y' "
		strSql = strSql & " WHERE 1 = 1 "
		strSql = strSql & addSql
		strSql = strSql & " and p.itemid <> 0 "
		strSql = strSql & " and p.isusing = 'Y' "
		strSql = strSql & " and isnull(p.secretYN, '') <> '' "
		strSql = strSql & " ORDER BY p.regdate DESC"
		db3_rsget.pagesize = FPageSize
		db3_rsget.Open strSql, db3_dbget, 1
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not db3_rsget.EOF Then
			db3_rsget.absolutepage = FCurrPage
			Do until db3_rsget.EOF
				Set FItemList(i) = new cQnaSummaryItem
					FItemList(i).FID			= db3_rsget("id")
					FItemList(i).FItemID		= db3_rsget("itemid")
					FItemList(i).FUserid		= db3_rsget("userid")
					FItemList(i).Fmakerid		= db3_rsget("makerid")
					FItemList(i).FCdl			= db3_rsget("cdl")
					FItemList(i).Fusername		= db2html(db3_rsget("username"))
					FItemList(i).FContents		= db2html(db3_rsget("contents"))
					FItemList(i).FTitle			= DdotFormat(FItemList(i).FContents,35)
					FitemList(i).FQadiv			= db3_rsget("qadiv")
					FItemList(i).FReplyuser 	= db3_rsget("replyuser")
					FItemList(i).FReplycontents = db2html(db3_rsget("replycontents"))
					FItemList(i).FReplytitle	= DdotFormat(FItemList(i).FReplycontents,40)
					FItemList(i).Fregdate		= db3_rsget("regdate")
					FItemList(i).FBrandName		= db2html(db3_rsget("brandname"))
					FItemList(i).FReplydate		= db3_rsget("replydate")
					FItemList(i).Fdeliverytype	= db3_rsget("deliverytype")
					FItemList(i).FItemDiv		= db3_rsget("itemdiv")
					FItemList(i).FSecretYN		= db3_rsget("secretYN")
					FItemList(i).FCSusername	= db3_rsget("username")
				i = i + 1
				db3_rsget.moveNext
			Loop
		End If
	End Sub
End Class
%>