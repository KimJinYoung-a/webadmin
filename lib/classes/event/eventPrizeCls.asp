<%
Class CEventPrizeJoinItem
	public FeventGubun
	public Fevt_code
	public Fevt_name
	public Fevt_startdate
	public Fevt_enddate
	public Fevt_state
	public Fmaster_isusing
	public Fevt_prizedate
	public Fuserid
	public Fcomment
	public Fdetail_isusing
	public Fregdate
	public finvaliduserid

	public function GetEventGubunName()
		Select Case FeventGubun
			Case "designfingers"
				GetEventGubunName = "디자인핑거스"
			Case "culturestation"
				GetEventGubunName = "컬쳐스테이션"
			Case "tbl_event_etc"
				GetEventGubunName = "기타"
			Case Else
				GetEventGubunName = "일반"
		End Select
	end function

	public function GetIsUsingStr()
		if (Fmaster_isusing <> "Y") or (Fdetail_isusing <> "Y") then
			GetIsUsingStr = "<font color='red'>삭제</font>"
		else
			GetIsUsingStr = "정상"
		end if
	end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CEventPrize
	public FSUserid
	public FEPType
	public FEPStatus
	public FEKind
	public FEEventCode
	public FEEventName

	public FTotCnt
	public FCPage
	public FPSize
	public FGubun
	public FWinnerOX
	public FResultCount
	public FRectRegDate1
	public FRectRegDate2

    public FItemList()
    public FOneItem

    public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	''public FResultCount
	public FScrollCount
	public FPageCount

	public FRectEventGubun
	public FRectUserid
	public FRectUserName
	public FRectUserCell
	public FRectEventCode
	public FRectEventName
	public FRectStartDate
	public FRectEndDate
	public frectgubun
	public frectinvaliduseryn

	'//admin/eventmanage/event/eventprize_list.asp
	public Function fnGetPrizeList
		Dim strSql,addSql,addSql1
		addSql = ""

		IF FEKind <> "" THEN
			addSql = addSql & " and C.evt_kind ="&FEKind
		END IF

		IF FSUserid <> "" THEN
			addSql = addSql & " and a.evt_winner ='"&FSUserid&"'"
		END IF

		IF FRectUserName <> "" THEN
			addSql = addSql & " and u.username ='"&FRectUserName&"'"
		END IF

		IF FRectUserCell <> "" THEN
			addSql = addSql & " and u.usercell ='"&FRectUserCell&"'"
		END IF

		IF FEPType <> "" THEN
			addSql = addSql & " and evtprize_type = " &FEPType
		END IF

		IF FEPStatus <> "" THEN
			addSql = addSql & " and evtprize_status = " &FEPStatus
		END IF

		IF FEEventCode <> "" THEN
			addSql = addSql & " and (A.evt_code IN (" & FEEventCode & ") OR A.evtgroup_code IN (" & FEEventCode & ")) "
			addSql1 = addSql & " and (D.evt_code IN (" & FEEventCode & ") OR D.evtgroup_code IN (" & FEEventCode & ")) "
		END IF

		IF FEEventName <> "" THEN
			addSql = addSql & " and C.evt_name Like '%" & FEEventName & "%'"
		END IF
		if frectinvaliduseryn="Y" then
			addSql = addSql & " and iu.idx is not null"
		elseif frectinvaliduseryn="N" then
			addSql = addSql & " and iu.idx is null"
		end if

		strSql = " SELECT COUNT(evtprize_code) FROM [db_event].[dbo].[tbl_event_prize] as A "&_
				"		Inner Join [db_event].[dbo].[tbl_event] as C On A.evt_code = C.evt_code "&_
				" left join db_user.dbo.tbl_invalid_user iu"&_
				" 	on A.evt_winner=iu.invaliduserid"&_
				" 	and iu.isusing='Y'"&_
				" 	and iu.gubun='"&frectgubun&"'"&_
				" left join [db_user].[dbo].[tbl_user_n] u on a.evt_winner = u.userid "&_
				"	WHERE 1=1 " &addSql

		'response.write strsql & "<br>"
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End IF
		rsget.Close

		IF FTotCnt >0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )+1
			strSql = " SELECT Top "&FPSize&" A.evtprize_code, A.evtprize_type, A.evt_code, A.evt_rankname, A.evt_winner, A.evtprize_startdate, A.evtprize_enddate "&_
					" ,A.evtprize_status, A.give_evtprizecode, A.evtprize_name, A.evt_regdate, A.lastupdate, B.songjangno, A.evtgroup_code, C.evt_kind, B.id "&_
					" , iu.invaliduserid"&_
					" , ts.itemuse_sdate, ts.itemuse_edate, ts.usewrite_sdate, ts.usewrite_edate "&_
					" FROM [db_event].[dbo].[tbl_event_prize] as A "&_
					"		Inner Join [db_event].[dbo].[tbl_event] as C On A.evt_code = C.evt_code "&_
					"		Left Outer Join [db_sitemaster].[dbo].[tbl_etc_songjang] as B ON A.evtprize_code = B.evtprize_code and B.deleteyn ='N' "&_
					" left join db_user.dbo.tbl_invalid_user iu"&_
					" 	on A.evt_winner=iu.invaliduserid"&_
					" 	and iu.isusing='Y'"&_
					" 	and iu.gubun='"&frectgubun&"'"&_
					" left join [db_user].[dbo].[tbl_user_n] u on a.evt_winner = u.userid "&_
					" left join [db_event].[dbo].[tbl_tester_event_winner] ts on A.evtprize_code = ts.evtprize_code "&_
					" WHERE A.evtprize_code <= (SELECT MIN(evtprize_code) FROM ( SELECT TOP "&iDelCnt&" evtprize_code FROM [db_event].[dbo].[tbl_event_prize] as D"&_
					"	WHERE 1=1 "&addSql1&" Order By evtprize_code DESC) as T ) "&addSql&" ORDER BY A.evtprize_code DESC "

			''response.write strsql & "<br>"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetPrizeList = rsget.getRows()
			End IF
			rsget.Close
		END IF
	End Function

	public FPrizeType
	public FStatus
	public FSongjangno
	public FStatusDesc
	'-----------------------------------------------------------------------
	' fnSetStatus : 이벤트 공통코드 가져오기
	'-----------------------------------------------------------------------
	public Function fnSetStatus
		FStatusDesc =""
	IF FPrizeType = 2 THEN
        FStatusDesc="쿠폰발급완료"
	ELSEIF FPrizeType = 3 THEN
        IF FStatus = 0 THEN
           FStatusDesc ="배송지입력대기"
        ELSEIF FStatus = 3 THEN
            IF  FSongjangno <> "" THEN
            	FStatusDesc="출고완료"
            ELSE
            	FStatusDesc="상품준비중"
            END IF
        END IF
    ELSEIF FPrizeType = 4 THEN
         IF FStatus = 0 THEN
         	FStatusDesc="티켓승인대기"
         ELSEIF FStatus = 3 THEN
         	FStatusDesc="티켓승인확정"
         END IF
    END IF
	End Function

	'####### 이벤트 참여 리스트
	public Function fnGetEventJoinList
		Dim strSql,iDelCnt, addSql, addSqlIn

		If FGubun = "e" Then	'### 일반이벤트
				If FRectRegDate1 <> "" Then
					addSql = addSql & " and A.evtcom_regdate >= '" & FRectRegDate1 & "' "
					addSqlIn = addSqlIn & " and X.evtcom_regdate >= '" & FRectRegDate1 & "' "
				End If

				If FRectRegDate2 <> "" Then
					addSql = addSql & " and A.evtcom_regdate <= '" & FRectRegDate2 & " 23:59:59' "
					addSqlIn = addSqlIn & " and X.evtcom_regdate <= '" & FRectRegDate2 & " 23:59:59' "
				End If

				If FEEventCode <> "" Then
					addSql = addSql & " AND A.evt_code = '" & FEEventCode & "' "
					addSqlIn = addSqlIn & " AND X.evt_code = '" & FEEventCode & "' "
				End IF

				IF FSUserid <> "" Then
					addSql = addSql & " AND A.userid = '" & FSUserid & "' "
					addSqlIn = addSqlIn & " AND X.userid = '" & FSUserid & "' "
				End If

				IF FWinnerOX <> "" Then
					addSql = addSql & " AND B.prizeyn = '" & FWinnerOX & "' "
					addSqlIn = addSqlIn & " AND Y.prizeyn = '" & FWinnerOX & "' "
				End If

				strSql = "SELECT COUNT(A.evtcom_idx) FROM [db_event].[dbo].[tbl_event_comment] AS A " & _
						 "		INNER JOIN [db_event].[dbo].[tbl_event] AS B ON A.evt_code = B.evt_code " & _
						 "WHERE A.evtcom_using = 'Y' " & addSql & " "

				rsget.Open strSql, dbget, 1
				IF not rsget.EOF THEN
					FTotCnt = rsget(0)
				END IF

				rsget.close
				FResultCount = FTotCnt
				IF FTotCnt > 0 THEN
					iDelCnt =  (FCPage - 1) * FPSize
					strSql = "SELECT TOP " & FPSize & " " & _
							 "		B.evt_code, B.evt_kind, B.evt_name, CASE WHEN getdate() Between evt_startdate And evt_enddate THEN '진행중' ELSE '종료' END As evt_state, " & _
							 "		B.evt_prizedate, B.prizeyn AS WinnerOX, A.evtcom_regdate, A.userid, A.evtcom_idx " & _
							 "	FROM [db_event].[dbo].[tbl_event_comment] AS A " & _
							 "		INNER JOIN [db_event].[dbo].[tbl_event] AS B ON A.evt_code = B.evt_code " & _
							 "	WHERE A.evtcom_using = 'Y' " & addSql & " " & _
							 "	AND A.evtcom_idx NOT IN " & _
							 "	( " & _
							 "		SELECT TOP " & iDelCnt & " X.evtcom_idx FROM [db_event].[dbo].[tbl_event_comment] AS X " & _
							 "		INNER JOIN [db_event].[dbo].[tbl_event] AS Y ON X.evt_code = Y.evt_code " & _
							 "		WHERE X.evtcom_using = 'Y' " & addSqlIn & " " & _
							 "		ORDER BY X.evtcom_idx DESC " & _
							 "	) " & _
							 "	ORDER BY A.evtcom_idx DESC "
					'response.write strSql
					rsget.Open strSql, dbget, 1
				IF not rsget.EOF THEN
					fnGetEventJoinList =rsget.getRows()
				END IF
				rsget.close
				END IF
		ElseIf FGubun = "f" Then
				If FRectRegDate1 <> "" Then
					addSql = addSql & " and A.regdate >= '" & FRectRegDate1 & "' "
					addSqlIn = addSqlIn & " and X.regdate >= '" & FRectRegDate1 & "' "
				End If

				If FRectRegDate2 <> "" Then
					addSql = addSql & " and A.regdate <= '" & FRectRegDate2 & " 23:59:59' "
					addSqlIn = addSqlIn & " and X.regdate <= '" & FRectRegDate2 & " 23:59:59' "
				End If

				If FEEventCode <> "" Then
					addSql = addSql & " AND A.masterid = '" & FEEventCode & "' "
					addSqlIn = addSqlIn & " AND X.masterid = '" & FEEventCode & "' "
				End IF

				IF FSUserid <> "" Then
					addSql = addSql & " AND A.userid = '" & FSUserid & "' "
					addSqlIn = addSqlIn & " AND X.userid = '" & FSUserid & "' "
				End If

				IF FWinnerOX = "Y" Then
					addSql = addSql & " AND DateAdd(dd,-1,getdate()) > B.PrizeDate "
					addSqlIn = addSqlIn & " AND DateAdd(dd,-1,getdate()) > Y.PrizeDate "
				ElseIF FWinnerOX = "N" Then
					addSql = addSql & " AND DateAdd(dd,-1,getdate()) <= B.PrizeDate "
					addSqlIn = addSqlIn & " AND DateAdd(dd,-1,getdate()) <= Y.PrizeDate "
				End If

				strSql = "SELECT COUNT(A.id) FROM [db_sitemaster].[dbo].[tbl_zf_comments] AS A " & _
						 "		INNER JOIN [db_sitemaster].[dbo].[tbl_designfingers] AS B ON A.masterid = B.DFSeq " & _
						 "WHERE A.gubuncd = '7' AND A.isdelete = 'N' AND A.sitename = '10x10' " & addSql & " "

				rsget.Open strSql, dbget, 1
				IF not rsget.EOF THEN
					FTotCnt = rsget(0)
				END IF

				rsget.close
				FResultCount = FTotCnt
				IF FTotCnt > 0 THEN
					iDelCnt =  (FCPage - 1) * FPSize
					strSql = "SELECT TOP " & FPSize & " " & _
							 "		A.masterid AS evt_code, '디자인핑거스' AS evt_kind, B.Title AS evt_name, '-' As evt_state, B.PrizeDate, " & _
							 "		CASE WHEN DateAdd(dd,-1,getdate()) <= B.PrizeDate THEN 'N' ELSE 'Y' END As WinnerOX, A.regdate, A.userid, A.id " & _
							 "	FROM [db_sitemaster].[dbo].[tbl_zf_comments] AS A " & _
							 "		INNER JOIN [db_sitemaster].[dbo].[tbl_designfingers] AS B ON A.masterid = B.DFSeq " & _
							 "	WHERE A.gubuncd = '7' AND A.isdelete = 'N' AND A.sitename = '10x10' " & addSql & " " & _
							 "	AND A.id NOT IN " & _
							 "	( " & _
							 "		SELECT TOP " & iDelCnt & " X.id FROM [db_sitemaster].[dbo].[tbl_zf_comments] AS X " & _
							 "		INNER JOIN [db_sitemaster].[dbo].[tbl_designfingers] AS Y ON X.masterid = Y.DFSeq " & _
							 "		WHERE X.gubuncd = '7' AND X.isdelete = 'N' AND X.sitename = '10x10' " & addSqlIn & " " & _
							 "		ORDER BY X.id DESC " & _
							 "	) " & _
							 "	ORDER BY A.id DESC "
					'response.write strSql
					rsget.Open strSql, dbget, 1
				IF not rsget.EOF THEN
					fnGetEventJoinList =rsget.getRows()
				END IF
				rsget.close
				END IF
		ElseIf FGubun = "c" Then
				If FRectRegDate1 <> "" Then
					addSql = addSql & " and A.regdate >= '" & FRectRegDate1 & "' "
					addSqlIn = addSqlIn & " and X.regdate >= '" & FRectRegDate1 & "' "
				End If

				If FRectRegDate2 <> "" Then
					addSql = addSql & " and A.regdate <= '" & FRectRegDate2 & " 23:59:59' "
					addSqlIn = addSqlIn & " and X.regdate <= '" & FRectRegDate2 & " 23:59:59' "
				End If

				If FEEventCode <> "" Then
					addSql = addSql & " AND A.evt_code = '" & FEEventCode & "' "
					addSqlIn = addSqlIn & " AND X.evt_code = '" & FEEventCode & "' "
				End IF

				IF FSUserid <> "" Then
					addSql = addSql & " AND A.userid = '" & FSUserid & "' "
					addSqlIn = addSqlIn & " AND X.userid = '" & FSUserid & "' "
				End If

				strSql = "SELECT COUNT(A.idx) FROM [db_culture_station].[dbo].[tbl_culturestation_event_comment] AS A " & _
						 "		INNER JOIN [db_culture_station].[dbo].[tbl_culturestation_event] AS B ON A.evt_code = B.evt_code " & _
						 "WHERE B.isusing = 'Y' " & addSql & " "

				rsget.Open strSql, dbget, 1
				IF not rsget.EOF THEN
					FTotCnt = rsget(0)
				END IF

				rsget.close
				FResultCount = FTotCnt
				IF FTotCnt > 0 THEN
					iDelCnt =  (FCPage - 1) * FPSize
					strSql = "SELECT TOP " & FPSize & " " & _
							 "		B.evt_code, B.evt_type AS evt_kind, B.evt_name, CASE WHEN getdate() Between startdate And enddate THEN '진행중' ELSE '종료' END As evt_prizedate, " & _
							 "		B.eventdate AS evt_prizedate, B.prizeyn AS WinnerOX, A.regdate, A.userid, A.idx " & _
							 "	FROM [db_culture_station].[dbo].[tbl_culturestation_event_comment] AS A " & _
							 "		INNER JOIN [db_culture_station].[dbo].[tbl_culturestation_event] AS B ON A.evt_code = B.evt_code " & _
							 "	WHERE B.isusing = 'Y' " & addSql & " " & _
							 "	AND A.idx NOT IN " & _
							 "	( " & _
							 "		SELECT TOP " & iDelCnt & " X.idx FROM [db_culture_station].[dbo].[tbl_culturestation_event_comment] AS X " & _
							 "		INNER JOIN [db_culture_station].[dbo].[tbl_culturestation_event] AS Y ON X.evt_code = Y.evt_code " & _
							 "		WHERE X.isusing = 'Y' " & addSqlIn & " " & _
							 "		ORDER BY X.idx DESC " & _
							 "	) " & _
							 "	ORDER BY A.idx DESC "
					'response.write strSql
					rsget.Open strSql, dbget, 1
				IF not rsget.EOF THEN
					fnGetEventJoinList =rsget.getRows()
				END IF
				rsget.close
				END IF
		End If
	End Function

	'// 아이디 검색인 경우
	'//admin/eventmanage/event/eventjoin_list_new.asp
	public Sub GetUserEventJoinListNew
		dim sqlStr, addSqlStr, i
		dim tmpTable

		tmpTable = " select top 200 'tbl_event' as eventGubun, e.evt_code, e.evt_name, e.evt_startdate, e.evt_enddate, e.evt_state, e.evt_using as master_isusing, e.evt_prizedate, c.userid, c.evtcom_txt as comment, c.evtcom_using as detail_isusing, c.evtcom_regdate as regdate "
		tmpTable = tmpTable + " from "
		tmpTable = tmpTable + " 	db_event.dbo.tbl_event e "
		tmpTable = tmpTable + " 	join [db_event].[dbo].[tbl_event_comment] c "
		tmpTable = tmpTable + " 	on "
		tmpTable = tmpTable + " 		e.evt_code = c.evt_code "
		tmpTable = tmpTable + " where "
		tmpTable = tmpTable + " 	1 = 1 "
		tmpTable = tmpTable + " 	and c.evtcom_regdate >= '" + CStr(FRectStartdate) + "' "
		tmpTable = tmpTable + " 	and c.evtcom_regdate < '" + CStr(FRectEndDate) + "' "
		tmpTable = tmpTable + " 	and c.userid = '" + CStr(FRectUserid) + "' "
		tmpTable = tmpTable + " union all "
		tmpTable = tmpTable + " select top 200 'designfingers' as eventGubun, e.dfseq as evt_code, e.title as evt_name, '2001-10-10' as evt_startdate, '2001-10-10' as evt_enddate, e.isdisplay as evt_state, (case when e.isusing = 1 then 'Y' else 'N' end) as master_isusing, e.prizedate as evt_prizedate, c.userid, c.comment, c.isdelete as detail_isusing, c.regdate "
		tmpTable = tmpTable + " from "
		tmpTable = tmpTable + " 	[db_sitemaster].[dbo].[tbl_designfingers] e "
		tmpTable = tmpTable + " 	join [db_sitemaster].[dbo].[tbl_zf_comments] c "
		tmpTable = tmpTable + " 	on "
		tmpTable = tmpTable + " 		e.DFSeq = c.masterid "
		tmpTable = tmpTable + " where "
		tmpTable = tmpTable + " 	1 = 1 "
		tmpTable = tmpTable + " 	and c.regdate >= '" + CStr(FRectStartdate) + "' "
		tmpTable = tmpTable + " 	and c.regdate < '" + CStr(FRectEndDate) + "' "
		tmpTable = tmpTable + " 	and c.userid = '" + CStr(FRectUserid) + "' "
		tmpTable = tmpTable + " union all "
		tmpTable = tmpTable + " select top 200 'culturestation' as eventGubun, e.evt_code, e.evt_name, e.startdate as evt_startdate, e.enddate as evt_enddate, '' as evt_state, e.isusing as master_isusing, e.eventdate as evt_prizedate, c.userid, c.comment, c.isusing as detail_isusing, c.regdate "
		tmpTable = tmpTable + " from "
		tmpTable = tmpTable + " 	[db_culture_station].[dbo].[tbl_culturestation_event] e "
		tmpTable = tmpTable + " 	join [db_culture_station].[dbo].[tbl_culturestation_event_comment] c "
		tmpTable = tmpTable + " 	on "
		tmpTable = tmpTable + " 		e.evt_code = c.evt_code "
		tmpTable = tmpTable + " where "
		tmpTable = tmpTable + " 	1 = 1 "
		tmpTable = tmpTable + " 	and c.regdate >= '" + CStr(FRectStartdate) + "' "
		tmpTable = tmpTable + " 	and c.regdate < '" + CStr(FRectEndDate) + "' "
		tmpTable = tmpTable + " 	and c.userid = '" + CStr(FRectUserid) + "' "
		tmpTable = tmpTable + " union all "
		tmpTable = tmpTable + " select top 200 'tbl_event_etc' as eventGubun, e.evt_code, e.evt_name, e.evt_startdate, e.evt_enddate, e.evt_state, e.evt_using as master_isusing, e.evt_prizedate, c.userid, '' as comment, 'Y' as detail_isusing, c.regdate "
		tmpTable = tmpTable + " from "
		tmpTable = tmpTable + " 	db_event.dbo.tbl_event e "
		tmpTable = tmpTable + " 	join db_event.dbo.tbl_event_subscript c "
		tmpTable = tmpTable + " 	on "
		tmpTable = tmpTable + " 		e.evt_code = c.evt_code "
		tmpTable = tmpTable + " where "
		tmpTable = tmpTable + " 	1 = 1 "
		tmpTable = tmpTable + " 	and c.regdate >= '" + CStr(FRectStartdate) + "' "
		tmpTable = tmpTable + " 	and c.regdate < '" + CStr(FRectEndDate) + "' "
		tmpTable = tmpTable + " 	and c.userid = '" + CStr(FRectUserid) + "' "

		addSqlStr = ""

		if (FRectEventGubun <> "") then
			addSqlStr = addSqlStr + " 	and T.eventGubun = '" + CStr(FRectEventGubun) + "' "
		end if

		if (FRectEventCode <> "") then
			addSqlStr = addSqlStr + " 	and T.evt_code = " + CStr(FRectEventCode) + " "
		end if

		if (FRectEventName <> "") then
			addSqlStr = addSqlStr + " 	and T.evt_name like '%" + CStr(FRectEventName) + "%' "
		end if

		if frectinvaliduseryn="Y" then
			addSqlStr = addSqlStr & " and iu.idx is not null"
		elseif frectinvaliduseryn="N" then
			addSqlStr = addSqlStr & " and iu.idx is null"
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	(" + tmpTable + ") T "
		sqlStr = sqlStr & " left join db_user.dbo.tbl_invalid_user iu"
		sqlStr = sqlStr & " 	on t.userid=iu.invaliduserid"
		sqlStr = sqlStr & " 	and iu.isusing='Y'"
		sqlStr = sqlStr & " 	and iu.gubun='"&frectgubun&"'"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
		'response.write sqlStr &"<br>"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " T.* , iu.invaliduserid"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	(" + tmpTable + ") T "
		sqlStr = sqlStr & " left join db_user.dbo.tbl_invalid_user iu"
		sqlStr = sqlStr & " 	on t.userid=iu.invaliduserid"
		sqlStr = sqlStr & " 	and iu.isusing='Y'"
		sqlStr = sqlStr & " 	and iu.gubun='"&frectgubun&"'"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
		sqlStr = sqlStr + " order by T.regdate desc "
		'response.write sqlStr &"<br>"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEventPrizeJoinItem

				''eventGubun, evt_code, evt_name, evt_startdate, evt_enddate, evt_state, master_isusing, evt_prizedate, userid, comment, detail_isusing, regdate

				FItemList(i).finvaliduserid = rsget("invaliduserid")
				FItemList(i).FeventGubun 		= rsget("eventGubun")
				FItemList(i).Fevt_code 			= rsget("evt_code")
				FItemList(i).Fevt_name 			= rsget("evt_name")
				FItemList(i).Fevt_startdate 	= rsget("evt_startdate")
				FItemList(i).Fevt_enddate 		= rsget("evt_enddate")
				FItemList(i).Fevt_state 		= rsget("evt_state")
				FItemList(i).Fmaster_isusing 	= rsget("master_isusing")
				FItemList(i).Fevt_prizedate 	= rsget("evt_prizedate")
				FItemList(i).Fuserid 			= rsget("userid")
				FItemList(i).Fcomment 			= rsget("comment")
				FItemList(i).Fdetail_isusing 	= rsget("detail_isusing")
				FItemList(i).Fregdate 			= rsget("regdate")

				if Left(FItemList(i).Fevt_startdate, 10) = "1900-01-01" or Left(FItemList(i).Fevt_startdate, 10) = "2001-10-10" then
					FItemList(i).Fevt_startdate = ""
				end if

				if Left(FItemList(i).Fevt_enddate, 10) = "1900-01-01" or Left(FItemList(i).Fevt_enddate, 10) = "2001-10-10" then
					FItemList(i).Fevt_enddate = ""
				end if

				if Left(FItemList(i).Fevt_prizedate, 10) = "1900-01-01" or Left(FItemList(i).Fevt_prizedate, 10) = "2001-10-10" then
					FItemList(i).Fevt_prizedate = ""
				end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close

	end Sub

	'//admin/eventmanage/event/eventjoin_list_new.asp
	public Sub GetEventJoinListNew
		dim sqlStr, addSqlStr, i
		dim tmpTable

		Select Case FRectEventGubun
			Case "designfingers"
				if (FRectStartdate <> "") then
					addSqlStr = addSqlStr + " 	and c.regdate >= '" + CStr(FRectStartdate) + "' "
				end if

				if (FRectEndDate <> "") then
					addSqlStr = addSqlStr + " 	and c.regdate < '" + CStr(FRectEndDate) + "' "
				end if

				if (FRectEventCode <> "") then
					addSqlStr = addSqlStr + " 	and e.dfseq = " + CStr(FRectEventCode) + " "
				end if

				if (FRectEventName <> "") then
					addSqlStr = addSqlStr + " 	and e.title like '%" + CStr(FRectEventName) + "%' "
				end if

				if (FRectUserid <> "") then
					addSqlStr = addSqlStr + " 	and c.userid = '" + CStr(FRectUserid) + "' "
				end if
			Case "culturestation"
				if (FRectStartdate <> "") then
					addSqlStr = addSqlStr + " 	and c.regdate >= '" + CStr(FRectStartdate) + "' "
				end if

				if (FRectEndDate <> "") then
					addSqlStr = addSqlStr + " 	and c.regdate < '" + CStr(FRectEndDate) + "' "
				end if

				if (FRectEventName <> "") then
					addSqlStr = addSqlStr + " 	and e.evt_name like '%" + CStr(FRectEventName) + "%' "
				end if

				if (FRectEventCode <> "") then
					addSqlStr = addSqlStr + " 	and e.evt_code = " + CStr(FRectEventCode) + " "
				end if

				if (FRectUserid <> "") then
					addSqlStr = addSqlStr + " 	and c.userid = '" + CStr(FRectUserid) + "' "
				end if
			Case "tbl_event_etc"
				if (FRectStartdate <> "") then
					addSqlStr = addSqlStr + " 	and c.regdate >= '" + CStr(FRectStartdate) + "' "
				end if

				if (FRectEndDate <> "") then
					addSqlStr = addSqlStr + " 	and c.regdate < '" + CStr(FRectEndDate) + "' "
				end if

				if (FRectEventName <> "") then
					addSqlStr = addSqlStr + " 	and e.evt_name like '%" + CStr(FRectEventName) + "%' "
				end if

				if (FRectEventCode <> "") then
					addSqlStr = addSqlStr + " 	and e.evt_code = " + CStr(FRectEventCode) + " "
				end if

				if (FRectUserid <> "") then
					addSqlStr = addSqlStr + " 	and c.userid = '" + CStr(FRectUserid) + "' "
				end if
			Case Else	'// tbl_event
				if (FRectStartdate <> "") then
					addSqlStr = addSqlStr + " 	and c.evtcom_regdate >= '" + CStr(FRectStartdate) + "' "
				end if

				if (FRectEndDate <> "") then
					addSqlStr = addSqlStr + " 	and c.evtcom_regdate < '" + CStr(FRectEndDate) + "' "
				end if

				if (FRectEventName <> "") then
					addSqlStr = addSqlStr + " 	and e.evt_name like '%" + CStr(FRectEventName) + "%' "
				end if

				if (FRectEventCode <> "") then
					addSqlStr = addSqlStr + " 	and e.evt_code = " + CStr(FRectEventCode) + " "
				end if

				if (FRectUserid <> "") then
					addSqlStr = addSqlStr + " 	and c.userid = '" + CStr(FRectUserid) + "' "
				end if
		End Select

		if frectinvaliduseryn="Y" then
			addSqlStr = addSqlStr & " and iu.idx is not null"
		elseif frectinvaliduseryn="N" then
			addSqlStr = addSqlStr & " and iu.idx is null"
		end if

		sqlStr = " select count(*) as cnt "

		Select Case FRectEventGubun
			Case "designfingers"
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	[db_sitemaster].[dbo].[tbl_designfingers] e "
				sqlStr = sqlStr + " 	join [db_sitemaster].[dbo].[tbl_zf_comments] c "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		e.DFSeq = c.masterid "
			Case "culturestation"
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	[db_culture_station].[dbo].[tbl_culturestation_event] e "
				sqlStr = sqlStr + " 	join [db_culture_station].[dbo].[tbl_culturestation_event_comment] c "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		e.evt_code = c.evt_code "
			Case "tbl_event_etc"
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_event.dbo.tbl_event e "
				sqlStr = sqlStr + " 	join db_event.dbo.tbl_event_subscript c "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		e.evt_code = c.evt_code "
			Case Else
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_event.dbo.tbl_event e "
				sqlStr = sqlStr + " 	join [db_event].[dbo].[tbl_event_comment] c "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		e.evt_code = c.evt_code "
		End Select

		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		Select Case FRectEventGubun
			Case "designfingers"
				'// ------------------------------------------------------------
				sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " 'designfingers' as eventGubun, e.dfseq as evt_code, e.title as evt_name, '2001-10-10' as evt_startdate, '2001-10-10' as evt_enddate, e.isdisplay as evt_state, (case when e.isusing = 1 then 'Y' else 'N' end) as master_isusing, e.prizedate as evt_prizedate, c.userid, c.comment, (case when c.isdelete = 'N' then 'Y' else 'N' end) as detail_isusing, c.regdate , iu.invaliduserid"
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	[db_sitemaster].[dbo].[tbl_designfingers] e "
				sqlStr = sqlStr + " 	join [db_sitemaster].[dbo].[tbl_zf_comments] c "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		e.DFSeq = c.masterid "
			Case "culturestation"
				'// ------------------------------------------------------------
				sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " 'culturestation' as eventGubun, e.evt_code, e.evt_name, e.startdate as evt_startdate, e.enddate as evt_enddate, '' as evt_state, e.isusing as master_isusing, e.eventdate as evt_prizedate, c.userid, c.comment, c.isusing as detail_isusing, c.regdate , iu.invaliduserid"
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	[db_culture_station].[dbo].[tbl_culturestation_event] e "
				sqlStr = sqlStr + " 	join [db_culture_station].[dbo].[tbl_culturestation_event_comment] c "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		e.evt_code = c.evt_code "
			Case "tbl_event_etc"
				'// ------------------------------------------------------------
				sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " 'tbl_event_etc' as eventGubun, e.evt_code, e.evt_name, e.evt_startdate, e.evt_enddate, e.evt_state, e.evt_using as master_isusing, e.evt_prizedate, c.userid, '' as comment, 'Y' as detail_isusing, c.regdate , iu.invaliduserid"
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_event.dbo.tbl_event e "
				sqlStr = sqlStr + " 	join db_event.dbo.tbl_event_subscript c "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		e.evt_code = c.evt_code "
			Case Else
				'// ------------------------------------------------------------
				sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " 'tbl_event' as eventGubun, e.evt_code, e.evt_name, e.evt_startdate, e.evt_enddate, e.evt_state, e.evt_using as master_isusing, e.evt_prizedate, c.userid, c.evtcom_txt as comment, c.evtcom_using as detail_isusing, c.evtcom_regdate as regdate , iu.invaliduserid"
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_event.dbo.tbl_event e "
				sqlStr = sqlStr + " 	join [db_event].[dbo].[tbl_event_comment] c "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		e.evt_code = c.evt_code "
		End Select

		sqlStr = sqlStr & " left join db_user.dbo.tbl_invalid_user iu"
		sqlStr = sqlStr & " 	on c.userid=iu.invaliduserid"
		sqlStr = sqlStr & " 	and iu.isusing='Y'"
		sqlStr = sqlStr & " 	and iu.gubun='"&frectgubun&"'"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr

		Select Case FRectEventGubun
			Case "designfingers"
				sqlStr = sqlStr + " order by c.regdate desc "
			Case "culturestation"
				sqlStr = sqlStr + " order by c.regdate desc "
			Case "tbl_event_etc"
				sqlStr = sqlStr + " order by c.regdate desc "
			Case Else
				sqlStr = sqlStr + " order by c.evtcom_regdate desc "
		End Select

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEventPrizeJoinItem

				''eventGubun, evt_code, evt_name, evt_startdate, evt_enddate, evt_state, master_isusing, evt_prizedate, userid, comment, detail_isusing, regdate
				FItemList(i).finvaliduserid = rsget("invaliduserid")
				FItemList(i).FeventGubun 		= rsget("eventGubun")
				FItemList(i).Fevt_code 			= rsget("evt_code")
				FItemList(i).Fevt_name 			= rsget("evt_name")
				FItemList(i).Fevt_startdate 	= rsget("evt_startdate")
				FItemList(i).Fevt_enddate 		= rsget("evt_enddate")
				FItemList(i).Fevt_state 		= rsget("evt_state")
				FItemList(i).Fmaster_isusing 	= rsget("master_isusing")
				FItemList(i).Fevt_prizedate 	= rsget("evt_prizedate")
				FItemList(i).Fuserid 			= rsget("userid")
				FItemList(i).Fcomment 			= rsget("comment")
				FItemList(i).Fdetail_isusing 	= rsget("detail_isusing")
				FItemList(i).Fregdate 			= rsget("regdate")

				if Left(FItemList(i).Fevt_startdate, 10) = "1900-01-01" or Left(FItemList(i).Fevt_startdate, 10) = "2001-10-10" then
					FItemList(i).Fevt_startdate = ""
				end if

				if Left(FItemList(i).Fevt_enddate, 10) = "1900-01-01" or Left(FItemList(i).Fevt_enddate, 10) = "2001-10-10" then
					FItemList(i).Fevt_enddate = ""
				end if

				if Left(FItemList(i).Fevt_prizedate, 10) = "1900-01-01" or Left(FItemList(i).Fevt_prizedate, 10) = "2001-10-10" then
					FItemList(i).Fevt_prizedate = ""
				end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
    end Sub

    Private Sub Class_Initialize()
		ReDim FItemList(0)

		FCurrPage		= 1
		FPageSize 		= 20
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
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

Function fnGetCommCodeArrDescCulture(ByVal iCodeValue)
	select Case iCodeValue
		Case 0	'느껴봐
		 	fnGetCommCodeArrDescCulture = "느껴봐"
		Case 1	'읽어봐
			fnGetCommCodeArrDescCulture = "읽어봐"
		Case 2	'들어봐
			fnGetCommCodeArrDescCulture = "들어봐"
		Case Else
			fnGetCommCodeArrDescCulture = ""
	End Select
End Function
%>
