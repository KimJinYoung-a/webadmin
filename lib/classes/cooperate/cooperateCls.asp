<%
'####################################################
' Description :  업무협조
' History : 강준구 생성
'			2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
'####################################################

	Class CCooperate
		public FCPage	'Set 현재 페이지
		public FPSize	'Set 페이지 사이즈
		public FTotCnt
		public FComeCnt
		public FSendCnt
		public FComeNewCnt
		public FSendNewCnt
		public FGubun
		public FUserID

		public FTeam
		public FRectDepartmentID
		public FRectWorker

		public FDoc_Idx
		public FDoc_Subj
		public FDoc_Id
		public FDoc_Regdate
		public FDoc_Start
		public FDoc_Status
		public FDoc_Status1
		public FDoc_End
		public FDoc_Name
		public FDoc_Type
		public FDoc_Import
		public FDoc_Diffi
		public FDoc_Content
		public FDoc_UseYN
		public FDoc_WorkerName
		public FDoc_Worker
		public FDoc_WorkerViewdate
		public FDoc_ReferViewdate
		public FDoc_Refer
		public FDoc_ReferName
		public FDoc_AnsOX
		public FDoc_MineOX
		public FDoc_UserName
		public FDoc_Searching
		public FDoc_IsRefer

		public FAns_Idx
		public FAns_Content

		public FState1Cnt
		public FState2Cnt
		public FState3Cnt
		public FState4Cnt
		public FState5Cnt
		public FReferCnt
		public FOnlyNewList

		public FDoc_reportidx
		public FDoc_reportstate
		public FSys_reportidx
		public FSys_reportstate



		'####### 업무협조리스트 #######
		public Function fnGetCooperateList
			Dim strSql,iDelCnt, strSubSql
			If FDoc_IsRefer = "o" Then		'####### 팝업용. 참조자, 작업자 완전 분리
					If FDoc_Status <> "x" Then
						strSubSql = " AND B.worker_id = '" & session("ssBctId") & "' AND A.doc_useyn = 'Y' "
						If FOnlyNewList = "o" Then
							strSubSql = strSubSql & " AND B.worker_viewdate is Null "
						End IF
					Else
						strSubSql = " AND D.refer_id = '" & session("ssBctId") & "' AND A.doc_useyn = 'Y' "
						If FOnlyNewList = "o" Then
							strSubSql = strSubSql & " AND D.refer_viewdate is Null "
						End IF
					End If
			Else
					If FDoc_MineOX <> "o" Then
						'strSubSql = " AND B.worker_id = '" & session("ssBctId") & "' AND A.doc_useyn = 'Y' "
						strSubSql = " AND (B.worker_id = '" & session("ssBctId") & "' OR D.refer_id = '" & session("ssBctId") & "') AND A.doc_useyn = 'Y' "
						If g_TeamJang = "o" Then
							If FTeam = "" Then
								strSubSql = " AND (B.part_sn IN(" & g_MyTeam & ") OR D.part_sn IN(" & g_MyTeam & ")) AND A.doc_useyn = 'Y' "
							Else
								strSubSql = " AND B.part_sn IN(" & FTeam & ") AND A.doc_useyn = 'Y' "
							End If
						Else
							If g_PartJang = "o" Then
								strSubSql = " AND B.part_sn = '" & g_MyPart & "' AND A.doc_useyn = 'Y' "
							End If
						End If
					Else
						strSubSql = " AND (B.worker_id = '" & session("ssBctId") & "' OR D.refer_id = '" & session("ssBctId") & "') AND A.doc_useyn = 'Y' "
					End If
			End If

			If FDoc_Status <> "x" Then
				If FDoc_Status = "0" Then
					strSubSql = strSubSql & " AND A.doc_status IN (1,2,4) "
				Else
					strSubSql = strSubSql & " AND A.doc_status = '" & FDoc_Status & "' "
				End If
			Else
				If FDoc_Status1 = "0" Then
					strSubSql = strSubSql & " AND A.doc_status IN (1,2,4) "
				Else
					If FDoc_Status = "6" AND FDoc_Status1 <> "x" Then
						strSubSql = strSubSql & " AND A.doc_status = '" & FDoc_Status1 & "' "
					End If
				End If
			End If

			If FDoc_Type <> "" Then
				strSubSql = strSubSql & " AND A.doc_type = '" & FDoc_Type & "' "
			End If

			If FDoc_AnsOX <> "" Then
				strSubSql = strSubSql & " AND A.doc_ans_ox = '" & FDoc_AnsOX & "' "
			End IF

			If FDoc_AnsOX <> "" Then
				strSubSql = strSubSql & " AND A.doc_ans_ox = '" & FDoc_AnsOX & "' "
			End IF

			If FDoc_UserName <> "" Then
				strSubSql = strSubSql & " AND C.username like '%" & FDoc_UserName & "%' "
			End IF

			If FDoc_Searching <> "" AND FDoc_Content <> "" Then
				If FDoc_Searching = "title" Then
					strSubSql = strSubSql & " AND A.doc_subject like '%" & FDoc_Content & "%' "
				ElseIf FDoc_Searching = "content" Then
					strSubSql = strSubSql & " AND A.doc_content like '%" & FDoc_Content & "%' "
				ElseIf FDoc_Searching = "doc_idx" Then
					strSubSql = strSubSql & " AND A.doc_idx = '" & FDoc_Content & "' "
				End IF
			End IF


			strSql = " SELECT COUNT(DISTINCT A.doc_idx) From " & _
					" 		[db_partner].[dbo].tbl_cooperate_document AS A " & _
					"		INNER JOIN [db_partner].[dbo].tbl_cooperate_worker AS B ON A.doc_idx = B.doc_idx " & _
					"		LEFT JOIN [db_partner].[dbo].tbl_cooperate_refer AS D ON A.doc_idx = D.doc_idx " & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS C ON A.id = C.userid "&_
					"	WHERE " & _
					"		1=1 " & strSubSql & " "

			rsget.Open strSql,dbget
			'response.write strSql
			IF not rsget.EOF THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close

			IF FTotCnt > 0 THEN
				iDelCnt =  ((FCPage - 1) * FPSize )
				'2014.03.06 품의서정보 정윤정 추가
				strSql = "	SELECT DISTINCT TOP "&FPSize&" " & _
						"		A.doc_idx, A.doc_subject, A.doc_type, A.doc_important, A.doc_difficult, A.doc_status, A.doc_regdate, C.username, A.doc_ans_ox, A.doc_workername, A.doc_refername "&_
						"		, B.worker_viewdate " & CHKIIF(FDoc_Status = "x" AND FDoc_IsRefer = "o",", D.refer_viewdate","") & " "&_
						"		, E1.reportidx, E1.reportstate, E2.reportidx, E2.reportstate "&_
						"	FROM [db_partner].[dbo].tbl_cooperate_document AS A "&_
						"	INNER JOIN [db_partner].[dbo].tbl_cooperate_worker AS B ON A.doc_idx = B.doc_idx "&_
						"	LEFT JOIN [db_partner].[dbo].tbl_cooperate_refer AS D ON A.doc_idx = D.doc_idx " & _
						"	INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS C ON A.id = C.userid "&_
						"		left outer join db_partner.dbo.tbl_eappreport AS E1 ON A.doc_idx = E1.scmlinkNo and E1.isUsing =1  and E1.edmsidx = 37"&_
						"		left outer join db_partner.dbo.tbl_eappreport as E2 on A.doc_idx = E2.scmlinkNo and E2.isUsing =1   and E2.edmsidx = 38 "&_
						"	WHERE " & _
						"		1=1 " & strSubSql & " AND " & _
						"		A.doc_idx NOT IN "&_
						"	( "&_
						"			SELECT DISTINCT TOP "&iDelCnt&" X.doc_idx FROM [db_partner].[dbo].tbl_cooperate_document AS X "&_
						"			INNER JOIN [db_partner].[dbo].tbl_cooperate_worker AS Y ON X.doc_idx = Y.doc_idx "&_
						"			LEFT JOIN [db_partner].[dbo].tbl_cooperate_refer AS Z ON X.doc_idx = Z.doc_idx " & _
						"			INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS W ON X.id = W.userid "&_
						"			WHERE 1=1 " & Replace(Replace(Replace(Replace(strSubSql,"B.","Y."),"A.","X."),"D.","Z."),"C.","W.") & " " & _
						"			ORDER BY X.doc_idx DESC "&_
						"	) "&_
						"	ORDER BY A.doc_idx DESC "
						' response.write strSql
				 rsget.Open strSql,dbget
				IF not rsget.EOF THEN
					fnGetCooperateList = rsget.getRows()
				END IF
				rsget.close
			END IF

		End Function


		'####### 내가작성한업무협조 #######
		public Function fnGetMyCooperateList
			Dim strSql,iDelCnt, strSubSql

			strSubSql = ""
			If FDoc_MineOX <> "o" Then
				strSubSql = " A.id = '" & session("ssBctId") & "' "
				If g_TeamJang = "o" Then
					If FTeam = "" Then
						strSubSql = " P.part_sn IN(" & g_MyTeam & ") "
					Else
						strSubSql = " P.part_sn IN(" & FTeam & ") "
					End If
				Else
					If g_PartJang = "o" Then
						strSubSql = " P.part_sn = '" & g_MyPart & "' "
					End If
				End If
			Else
				strSubSql = " A.id = '" & session("ssBctId") & "' "
			End If

			If FDoc_Status <> "x" Then
				If FDoc_Status = "0" Then
					strSubSql = strSubSql & " AND A.doc_status IN (1,2,4) "
				Else
					strSubSql = strSubSql & " AND A.doc_status = '" & FDoc_Status & "' "
				End If
			End If

			If FDoc_Type <> "" Then
				strSubSql = strSubSql & " AND A.doc_type = '" & FDoc_Type & "' "
			End If

			If FDoc_AnsOX <> "" Then
				strSubSql = strSubSql & " AND A.doc_ans_ox = '" & FDoc_AnsOX & "' "
			End IF

			strSql = " SELECT COUNT(A.doc_idx) From " & _
					" 		[db_partner].[dbo].tbl_cooperate_document AS A " & _
					"		JOIN [db_partner].[dbo].tbl_user_tenbyten AS C ON A.id = C.userid "&_
					"		JOIN [db_partner].[dbo].tbl_partner AS P ON A.id = P.id "&_
					"	WHERE " & strSubSql & " "
			rsget.Open strSql,dbget
			'response.write strSql
			IF not rsget.EOF THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close

			IF FTotCnt > 0 THEN
				iDelCnt =  ((FCPage - 1) * FPSize )
				'2014.03.06 품의서정보 정윤정 추가
				strSql = "	SELECT TOP "&FPSize&" A.doc_idx, A.doc_subject, A.doc_type, A.doc_important, "&_
						"		A.doc_difficult, A.doc_status, A.doc_regdate, A.doc_ans_ox, C.username "&_
						"		, E1.reportidx, E1.reportstate "&_
						"	FROM [db_partner].[dbo].tbl_cooperate_document AS A "&_
						"		JOIN [db_partner].[dbo].tbl_user_tenbyten AS C ON A.id = C.userid "&_
						"		JOIN [db_partner].[dbo].tbl_partner AS P ON A.id = P.id "&_
						"		left outer join db_partner.dbo.tbl_eappreport as E1 on A.doc_idx = E1.scmlinkNo and E1.isUsing =1   and E1.edmsidx = 37 "&_
						"	WHERE " & strSubSql & " AND " & _
						"		A.doc_idx NOT IN "&_
						"	( "&_
						"			SELECT TOP "&iDelCnt&" X.doc_idx FROM [db_partner].[dbo].tbl_cooperate_document AS X "&_
						"				JOIN [db_partner].[dbo].tbl_user_tenbyten AS Y ON X.id = Y.userid "&_
						"				JOIN [db_partner].[dbo].tbl_partner AS P ON X.id = P.id "&_
						"			WHERE " & Replace(Replace(strSubSql,"A.","X."),"C.","Y.") & " " & _
						"			ORDER BY X.doc_idx DESC "&_
						"	) "&_
						"	ORDER BY A.doc_idx DESC "
				rsget.Open strSql,dbget
				IF not rsget.EOF THEN
					fnGetMyCooperateList = rsget.getRows()
				END IF
				rsget.close
			END IF

		End Function


		'####### 협조문 보기 #######
		public Function fnGetCooperateView
			Dim strSql
			strSql = "	SELECT A.id, B.username, A.doc_regdate, A.doc_status, A.doc_startdate, A.doc_enddate, " & _
					"			A.doc_type, A.doc_important, A.doc_difficult, A.doc_subject, A.doc_content, A.doc_useyn, A.doc_refername " & _
					"				, E1.reportidx as reportidx1, E1.reportstate as reportstate1 , E2.reportidx as reportidx2, E2.reportstate as reportstate2 "&_
					"		FROM [db_partner].[dbo].tbl_cooperate_document AS A " & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B ON A.id = B.userid " & _
					"		left outer join db_partner.dbo.tbl_eappreport AS E1 ON A.doc_idx = E1.scmlinkNo and E1.isUsing =1  and E1.edmsidx = 37"&_
					"		left outer join db_partner.dbo.tbl_eappreport as E2 on A.doc_idx = E2.scmlinkNo and E2.isUsing =1   and E2.edmsidx = 38 "&_
					"	WHERE A.doc_idx = '" & FDoc_Idx & "' "
			rsget.Open strSql,dbget,1
			'response.write strSql
			IF not rsget.EOF THEN
				 FDoc_Id 		= rsget("id")
				 FDoc_Name		= rsget("username")
				 FDoc_Status	= rsget("doc_status")
				 FDoc_Start		= rsget("doc_startdate")
				 FDoc_End		= rsget("doc_enddate")
				 FDoc_Type		= rsget("doc_type")
				 FDoc_Import	= rsget("doc_important")
				 FDoc_Diffi		= rsget("doc_difficult")
				 FDoc_Subj		= db2html(rsget("doc_subject"))
				 FDoc_Content	= db2html(rsget("doc_content"))
				 FDoc_UseYN		= rsget("doc_useyn")
				 FDoc_Regdate	= rsget("doc_regdate")
				 FDoc_ReferName	= rsget("doc_refername")
				 FDoc_reportidx   = rsget("reportidx1")
				 FDoc_reportstate = rsget("reportstate1")
				 FSys_reportidx   = rsget("reportidx2")
				 FSys_reportstate = rsget("reportstate2")
			END IF
			rsget.close

			strSql = "	SELECT A.worker_id, B.username, A.part_sn, Convert(varchar(20),A.worker_viewdate,120) AS worker_viewdate " & _
					"		FROM [db_partner].[dbo].tbl_cooperate_worker AS A " & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B ON A.worker_id = B.userid " & _
					"	WHERE doc_idx = '" & FDoc_Idx & "' " & _
					"	ORDER BY A.idx ASC "
			rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				Do Until rsget.Eof
					FDoc_WorkerName = FDoc_WorkerName & rsget("username") & ","
					FDoc_Worker		= FDoc_Worker & rsget("worker_id") & "|" & rsget("part_sn") & ","
					FDoc_WorkerViewdate = FDoc_WorkerViewdate & rsget("worker_viewdate") & ","
				rsget.MoveNext
				Loop
			END IF
			rsget.close

			FDoc_WorkerName = Left(FDoc_WorkerName,Len(FDoc_WorkerName)-1)
			FDoc_Worker = Left(FDoc_Worker,Len(FDoc_Worker)-1)
			FDoc_WorkerViewdate = Left(FDoc_WorkerViewdate,Len(FDoc_WorkerViewdate)-1)

			If FDoc_ReferName <> "" Then
				strSql = "	SELECT A.refer_id, B.username, A.part_sn, Convert(varchar(20),A.refer_viewdate,120) AS refer_viewdate " & _
						"		FROM [db_partner].[dbo].tbl_cooperate_refer AS A " & _
						"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B ON A.refer_id = B.userid " & _
						"	WHERE doc_idx = '" & FDoc_Idx & "' " & _
						"	ORDER BY A.idx ASC "
				rsget.Open strSql,dbget,1
				IF not rsget.EOF THEN
					FDoc_ReferName = ""
					Do Until rsget.Eof
						FDoc_ReferName 		= FDoc_ReferName & rsget("username") & ","
						FDoc_Refer			= FDoc_Refer & rsget("refer_id") & "|" & rsget("part_sn") & ","
						FDoc_ReferViewdate	= FDoc_ReferViewdate & rsget("refer_viewdate") & ","
					rsget.MoveNext
					Loop

					FDoc_ReferName = Left(FDoc_ReferName,Len(FDoc_ReferName)-1)
					FDoc_Refer = Left(FDoc_Refer,Len(FDoc_Refer)-1)
					FDoc_ReferViewdate = Left(FDoc_ReferViewdate,Len(FDoc_ReferViewdate)-1)
				END IF
				rsget.close
			End If

		End Function



		'####### 협조문답변리스트 #######
		public Function fnGetCooperateAnsList
			Dim strSql,iDelCnt, strSubSql

			strSubSql = " AND A.doc_idx = '" & FDoc_Idx & "' "

			strSql = " SELECT COUNT(A.ans_idx) From " & _
					" 		[db_partner].[dbo].tbl_cooperate_ans AS A with (nolock)" & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B with (nolock) ON A.id = B.userid " & _
					"	WHERE " & _
					"		1=1 AND ans_useyn = 'Y' " & strSubSql & " "

			'response.write strSql & "<Br>"
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

			IF not rsget.EOF THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close

			IF FTotCnt > 0 THEN
				iDelCnt =  ((FCPage - 1) * FPSize )

				strSql = "	SELECT TOP "&FPSize&" " & _
						"			 A.ans_idx, A.ans_type, A.ans_content, A.ans_regdate, B.username, A.id "&_
						"	FROM [db_partner].[dbo].tbl_cooperate_ans AS A with (nolock)"&_
						"	INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B with (nolock) ON A.id = B.userid "&_
						"	WHERE " & _
						"		1=1 AND A.ans_useyn = 'Y' " & strSubSql & " AND " & _
						"		A.ans_idx NOT IN "&_
						"	( "&_
						"			SELECT TOP "&iDelCnt&" X.ans_idx FROM [db_partner].[dbo].tbl_cooperate_ans AS X with (nolock)"&_
						"			INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS Y with (nolock) ON X.id = Y.userid "&_
						"			WHERE 1=1 AND X.ans_useyn = 'Y' " & Replace(strSubSql,"B.","Y.") & " " & _
						"			ORDER BY X.ans_idx DESC "&_
						"	) "&_
						"	ORDER BY A.ans_idx DESC "

				'response.write strSql & "<Br>"
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

				IF not rsget.EOF THEN
					fnGetCooperateAnsList = rsget.getRows()
				END IF
				rsget.close
			END IF

		End Function


		'####### 협조문답변보기 #######
		public Function fnGetCooperateAnsView
			Dim strSql
			strSql = "	SELECT ans_content " & _
					"		FROM [db_partner].[dbo].tbl_cooperate_ans with (nolock)" & _
					"	WHERE  ans_useyn = 'Y' AND ans_idx = '" & FAns_Idx & "' AND id = '" & session("ssBctId") & "' "

			'response.write strSql & "<Br>"
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

			IF not rsget.EOF THEN
				 FAns_Content	= db2html(rsget("ans_content"))
			END IF
			rsget.close

		End Function


		'##### 부서 목록 접수 #####
		'//common/offshop/member/PoptenbytenuserList.asp
		public Function fnPartList
			Dim strSql, i
			strSql = "SELECT part_sn, part_name From [db_partner].dbo.tbl_partinfo " &_
					 " WHERE part_isDel = 'N' " & _
					 "ORDER BY part_sort "
			rsget.Open strSql,dbget,1
			'response.write strSql
			IF not rsget.EOF THEN
				fnPartList = rsget.getRows()
			END IF
			rsget.Close
		End Function


		'####### 직원리스트 #######
		'//common/offshop/member/PoptenbytenuserList.asp
		public Function fnGetMemberList
			Dim strSql, addsql
			'2016-06-29 김진영 수정, AND (D.department_id = '" &FRectDepartmentID & "' or D.username like '%" & FRectWorker & "%' ) 이런 형식으로 했을 때, 팀선택시 like검색으로 인해 전체 인원 출력됨
			If FRectWorker <> "" Then
				addsql = addsql & "  or D.username like '%" & FRectWorker & "%' "
			End If
			If FTeam <> "" Then
				addsql = addsql & "  or D.part_sn='" & FTeam & "' "
			End If

			strSql = "	SELECT A.id, B.departmentname, C.posit_name, D.username AS company_name, D.department_id, " & _
					"		isNull((select Convert(varchar(20),worker_viewdate,120) from [db_partner].[dbo].tbl_cooperate_worker where doc_idx = '" & FDoc_Idx & "' and worker_id = A.id),'x') AS worker_viewdate " & _
					"		, isNull(E.halfgubun,'') AS vacation, D.mywork, D.empno " & _
					"		FROM [db_partner].[dbo].tbl_partner AS A " & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS D" & _
					"			ON A.id = D.userid " & _
					"			and d.isUsing=1 and (d.statediv ='Y' or (d.statediv ='N' and datediff(dd,d.retireday,getdate())<=0))" & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_department AS B ON D.department_id = B.cid " & _
					"		INNER JOIN [db_partner].[dbo].tbl_positInfo AS C ON D.posit_sn = C.posit_sn " & _
					"		LEFT JOIN ( " & _
					"			SELECT mm.userid, dd.halfgubun FROM [db_partner].[dbo].tbl_vacation_master AS mm " & _
					"			INNER JOIN [db_partner].[dbo].tbl_vacation_detail AS dd on mm.idx = dd.masteridx " & _
					"			WHERE mm.deleteyn <> 'Y' AND dd.deleteyn <> 'Y' AND dd.statedivcd IN ('R','A') AND ('" & date() & "' between convert(varchar(10),dd.startday,120) and convert(varchar(10),dd.endday,120)) " & _
					"		) AS E ON A.id = E.userid " & _
					"	WHERE A.isusing = 'Y' AND A.userdiv < 999 AND A.id <> '' AND Left(A.id,10) <> 'streetshop'   " & _
					"			AND (D.department_id = '" &FRectDepartmentID & "' "& addsql &" ) " & _
					"	ORDER BY   D.posit_sn ASC, D.regdate ASC "

			'response.write strSql & "<br>"
			rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				fnGetMemberList = rsget.getRows()
			END IF
			rsget.close
		End Function

		public Function fnGetMemberListNew
			Dim strSql
			dim tmpID

			tmpID = FRectDepartmentID
			if (tmpID = "") then
				tmpID = "-1"
			end if

			strSql = "	SELECT A.id, dv.departmentNameFull as part_name, C.posit_name, D.username AS company_name, A.part_sn, " & _
					"		isNull((select Convert(varchar(20),worker_viewdate,120) from [db_partner].[dbo].tbl_cooperate_worker where doc_idx = '" & FDoc_Idx & "' and worker_id = A.id),'x') AS worker_viewdate " & _
					"		, isNull(E.halfgubun,'') AS vacation, D.mywork " & _
					"		FROM [db_partner].[dbo].tbl_partner AS A " & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS D ON A.id = D.userid " & _
					"		INNER JOIN [db_partner].[dbo].tbl_partInfo AS B ON D.part_sn = B.part_sn " & _
					"		INNER JOIN [db_partner].[dbo].tbl_positInfo AS C ON D.posit_sn = C.posit_sn " & _
					"		LEFT JOIN ( " & _
					"			SELECT mm.userid, dd.halfgubun FROM [db_partner].[dbo].tbl_vacation_master AS mm " & _
					"			INNER JOIN [db_partner].[dbo].tbl_vacation_detail AS dd on mm.idx = dd.masteridx " & _
					"			WHERE mm.deleteyn <> 'Y' AND dd.deleteyn <> 'Y' AND dd.statedivcd IN ('R','A') AND ('" & date() & "' between convert(varchar(10),dd.startday,120) and convert(varchar(10),dd.endday,120)) " & _
					"		) AS E ON A.id = E.userid " & _
					"		JOIN db_partner.dbo.vw_user_department dv on D.department_id = dv.cid " & _
					"	WHERE A.isusing = 'Y' AND A.userdiv < 999 AND A.id <> '' " & _
					"			AND (dv.cid1 = " + CStr(tmpID) + " or dv.cid2 = " + CStr(tmpID) + " or dv.cid3 = " + CStr(tmpID) + " or dv.cid4 = " + CStr(tmpID) + ") " & _
					"			AND d.isUsing=1 and (d.statediv ='Y' or (d.statediv ='N' and datediff(dd,d.retireday,getdate())<=0))" & _
					"	ORDER BY D.part_sn ASC, D.posit_sn ASC, D.regdate ASC "
			''response.write strSql
			rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				fnGetMemberListNew = rsget.getRows()
			END IF
			rsget.close
		End Function

		'####### 협조문첨부파일리스트 #######
		public Function fnGetFileList
			Dim strSql
			strSql = "	SELECT file_idx, file_name, real_name " & _
					"		FROM [db_partner].[dbo].tbl_cooperate_file " & _
					"	WHERE doc_idx = '" & FDoc_Idx & "' " & _
					"	ORDER BY file_idx ASC "
			rsget.Open strSql,dbget,1
			'response.write strSql
			IF not rsget.EOF THEN
				fnGetFileList = rsget.getRows()
			END IF
			rsget.close
		End Function


		'####### 나에게 온 새로운 협조문 갯수 #######
		public Function fnGetCooperateCount
			Dim strSql
			strSql = "	SELECT COUNT(DISTINCT A.doc_idx) " & _
					"		, sum(Case When datediff(d,doc_regdate,getdate())<3 then 1 else 0 end) as newCnt " & _
					"		From [db_partner].[dbo].tbl_cooperate_document AS A with (nolock)" & _
					"		INNER JOIN [db_partner].[dbo].tbl_cooperate_worker AS B with (nolock) ON A.doc_idx = B.doc_idx " & _
					"	WHERE B.worker_id = '" & FDoc_Id & "' AND A.doc_status IN(1,2) AND A.doc_useyn = 'Y' "

			'response.write strSql & "<Br>"
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				FComeCnt = rsget(0)
				FComeNewCnt = rsget(1)
			Else
				FComeCnt = 0
				FComeNewCnt = 0
			END IF
			rsget.close
			strSql = "	SELECT COUNT(DISTINCT A.doc_idx) " & _
					"		, sum(Case When datediff(d,doc_regdate,getdate())<3 then 1 else 0 end) as newCnt " & _
					"		From [db_partner].[dbo].tbl_cooperate_document AS A with (nolock)" & _
					"	WHERE A.id = '" & FDoc_Id & "' AND A.doc_status IN(1,2) AND A.doc_useyn = 'Y' "

			'response.write strSql & "<Br>"
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				FSendCnt = rsget(0)
				FSendNewCnt = rsget(1)
			Else
				FSendCnt = 0
				FSendNewCnt = 0
			END IF
			rsget.close
		End Function


		'####### POPUP 버전 협조문 전체 카운트 #######
		public Function fnGetCooperateCount_PopVer
			Dim strSql
			strSql = "[db_partner].[dbo].[sp_Ten_Cooperate_MainCount]('" & FGubun & "', '" & FUserID & "')"
			'response.write strSql
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF not rsget.EOF THEN
				FState1Cnt		= rsget(0)
				FState2Cnt		= rsget(1)
				FState3Cnt		= rsget(2)
				FState4Cnt		= rsget(3)
				FState5Cnt		= rsget(4)
				FReferCnt		= rsget(5)
			END IF
			rsget.close
		End Function

	End Class


	'####### 코드 매니저 페이지용. 리스트. #######
	Class CCoopCommonCode
	public FCodeType
	public FCodeValue
	public FCodeDesc
	public FCodeUsing
	public FCodeSort

		'####### 공통코드 리스트 #######
		public Function fnGetCoopCodeList
			IF FCodeType = "" THEN Exit Function
			Dim strSql
			strSql = "SELECT code_type, code_value, code_desc, code_useyn, code_sort "&_
					" From [db_partner].[dbo].[tbl_cooperate_comCode] "&_
					" WHERE code_type = '"&FCodeType&"' ORDER BY code_sort ASC "
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetCoopCodeList = rsget.getRows()
			End IF
			rsget.Close
		End Function

		'####### 선택한 코드 내용 가져오기 #######
		public Function fnGetCoopCodeCont
			IF FCodeValue = "" or FCodeType = ""  THEN Exit Function
			Dim strSql
			strSql = " SELECT code_type, code_value, code_desc, code_useyn, code_sort "&_
					" From  [db_partner].[dbo].[tbl_cooperate_comCode] "&_
					" WHERE code_value = "&FCodeValue&" and code_type ='"&FCodeType&"'"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				FCodeType 	= rsget("code_type")
				FCodeValue 	= rsget("code_value")
				FCodeDesc 	= rsget("code_desc")
				FCodeUsing 	= rsget("code_useyn")
				FCodeSort	= rsget("code_sort")
			End IF
			rsget.Close
		End Function
	End Class



	'####### 로그 저장. #######
	Public Function LogInsert(ByVal iDoc_idx, ByVal sLog_Type, ByVal sAction)
		Dim strSql
		SELECT CASE sAction
			CASE "1"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), 협조문 작성"
			CASE "2"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), 협조문 수정"
			CASE "3"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), 협조문 삭제"
			CASE "4"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), 협조문 확인"
			CASE "5"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " 협조문 답변 작성"
			CASE "6"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " 협조문 답변 수정"
			CASE "7"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " 협조문 답변 삭제"
			CASE "8"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " 협조문 작업자에게 SMS 전송"
			CASE "9"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " 협조문 참조자에게 SMS 전송"
			CASE "91"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), 코드 생성"
			CASE "92"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), 코드 수정"
			CASE "93"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), 코드 삭제"
			CASE ELSE
				sAction = "기타 Action"
		END SELECT
		strSql = " INSERT INTO [db_partner].[dbo].tbl_cooperate_log(doc_idx, log_type, log_content, log_ip, log_regdate) " & _
				 "	VALUES('" & iDoc_idx & "','" & sLog_Type & "','" & sAction & "','" & Request.ServerVariables("REMOTE_ADDR") & "', getdate()) "
		dbget.execute strSql
	End Function


	'####### 자신한테 온 협조문 맨 처음 확인시 읽은 시간 저장. 로그도 처음 확인일때 입력. (update:2, 협조문 확인) #######
	Public Function WorkerView(ByVal iDoc_idx)
		Dim strSql
		strSql = " DECLARE @ISNULL datetime " & _
				 "	IF EXISTS(SELECT worker_viewdate FROM [db_partner].[dbo].tbl_cooperate_worker WHERE doc_idx = '" & iDoc_idx & "' AND worker_id = '" & session("ssBctId") & "') " & _
				 "	BEGIN " & _
				 "		SELECT @ISNULL = worker_viewdate FROM [db_partner].[dbo].tbl_cooperate_worker WHERE doc_idx = '" & iDoc_idx & "' AND worker_id = '" & session("ssBctId") & "' " & _
				 "		If @ISNULL is NULL " & _
				 "			BEGIN " & _
				 "			UPDATE [db_partner].[dbo].tbl_cooperate_worker SET worker_viewdate = getdate() WHERE doc_idx = '" & iDoc_idx & "' AND worker_id = '" & session("ssBctId") & "' " & _
				 "			INSERT INTO [db_partner].[dbo].tbl_cooperate_log(doc_idx, log_type, log_content, log_ip, log_regdate) VALUES('" & iDoc_idx & "', '2', '" & session("ssBctCname") & "(" & session("ssBctId") & "), 협조문 확인', '" & Request.ServerVariables("REMOTE_ADDR") & "', getdate()) " & _
				 "			END " & _
				 "	END "
		dbget.execute strSql
	End Function



	'####### 참조 확인일 저장. 상세내용은 위 작업확인일 저장과 같음. #######
	Public Function ReferView(ByVal iDoc_idx)
		Dim strSql
		strSql = " DECLARE @ISNULL datetime " & _
				 "	IF EXISTS(SELECT refer_viewdate FROM [db_partner].[dbo].tbl_cooperate_refer WHERE doc_idx = '" & iDoc_idx & "' AND refer_id = '" & session("ssBctId") & "') " & _
				 "	BEGIN " & _
				 "		SELECT @ISNULL = refer_viewdate FROM [db_partner].[dbo].tbl_cooperate_refer WHERE doc_idx = '" & iDoc_idx & "' AND refer_id = '" & session("ssBctId") & "' " & _
				 "		If @ISNULL is NULL " & _
				 "			BEGIN " & _
				 "			UPDATE [db_partner].[dbo].tbl_cooperate_refer SET refer_viewdate = getdate() WHERE doc_idx = '" & iDoc_idx & "' AND refer_id = '" & session("ssBctId") & "' " & _
				 "			INSERT INTO [db_partner].[dbo].tbl_cooperate_log(doc_idx, log_type, log_content, log_ip, log_regdate) VALUES('" & iDoc_idx & "', '2', '" & session("ssBctCname") & "(" & session("ssBctId") & "), 협조문 확인', '" & Request.ServerVariables("REMOTE_ADDR") & "', getdate()) " & _
				 "			END " & _
				 "	END "
		dbget.execute strSql
	End Function



	'####### 코드 select 박스 #######
	Sub sbOptCodeType(ByVal selCodeType)
%>
		<option value="doc_status" <%IF Cstr(selCodeType)="doc_status" THEN%>selected<%END IF%>>현재상태값</option>
		<option value="doc_type" <%IF Cstr(selCodeType)="doc_type" THEN%>selected<%END IF%>>업무구분</option>
		<option value="doc_important" <%IF Cstr(selCodeType)="doc_important" THEN%>selected<%END IF%>>업무중요도</option>
		<option value="doc_difficult" <%IF Cstr(selCodeType)="doc_difficult" THEN%>selected<%END IF%>>업무난이도</option>
		<option value="ans_type" <%IF Cstr(selCodeType)="ans_type" THEN%>selected<%END IF%>>답변구분</option>
		<option value="log_type" <%IF Cstr(selCodeType)="log_type" THEN%>selected<%END IF%>>로그구분</option>
<%
	End Sub


	'####### 코드매니저 페이지 외 하나씩 불러쓰는 공통 코드 관리. write 용, view 용. #######
	Public Function CommonCode(ByVal sUse, ByVal sType, ByVal sCode)
		Dim strSql, sBody, i, vTemp
		sBody = ""
		i = 0
		vTemp = sUse
		If sUse = "wnew" Then
			sUse = "w"
		End If

		'### sUse = "w" write 용
		If sUse = "w" Then
			If vTemp = "wnew" Then
				strSql = " SELECT code_value, code_desc From [db_partner].[dbo].[tbl_cooperate_comCode] WHERE code_type ='"&sType&"' AND code_value Not IN('4','5') AND code_useyn = 'Y' ORDER BY code_sort ASC"
			Else
				strSql = " SELECT code_value, code_desc From [db_partner].[dbo].[tbl_cooperate_comCode] WHERE code_type ='"&sType&"' AND code_useyn = 'Y' ORDER BY code_sort ASC"
			End IF

			If sType = "doc_status" AND sCode = "" Then
				sBody = "<input type='hidden' name='doc_status' value='1'>협조문 작성"
			Else
				rsget.Open strSql,dbget,1
				Do Until rsget.Eof
					'####### 업무구분은 select박스 형으로. 나머지는 radio 로. #######
					If sType = "doc_type" Then
						If i = 0 Then
							sBody = "<select name='doc_type' class='select' onChange='issystem(this.value)'>"
							If GetFileName() = "index" OR GetFileName() = "my_cooperate" Then
								sBody = sBody & "<option value=''>-요청구분-</option> "
							End IF
						End IF
						sBody = sBody & "<option value='" & rsget("code_value") & "' "
						If CStr(sCode) = CStr(rsget("code_value")) Then
							sBody = sBody & "selected"
						End If
						sBody = sBody & ">" & rsget("code_desc") & "</option>"
						If i = rsget.RecordCount-1 Then
							sBody = sBody & "</select>"
						End IF
					Else
						If Left(sCode,1) = "s" Then
							'####### index 페이지에 현재상태 검색용. select박스. #######
							If i = 0 Then
								sBody = "<select name='doc_status' class='select'><option value='x'>-처리상태-</option><option value='0' "
								If CStr(Replace(sCode,"s","")) = CStr(0) Then
									sBody = sBody & "selected"
								End If
								sBody = sBody & ">미처리 전체</option>"
							End IF
							sBody = sBody & "<option value='" & rsget("code_value") & "' "
							If CStr(Replace(sCode,"s","")) = CStr(rsget("code_value")) Then
								sBody = sBody & "selected"
							End If
							sBody = sBody & ">" & rsget("code_desc") & "</option>"
							If i = rsget.RecordCount-1 Then
								sBody = sBody & "</select>"
							End IF
						Else
							sBody = sBody & "<label id='" & sType & rsget("code_value") & "'>" & _
											"<input type='radio' name='" & sType & "' id='" & sType & rsget("code_value") & "' value='" & rsget("code_value") & "' "
							If CStr(sCode) = CStr(rsget("code_value")) Then
								sBody = sBody & "checked"
							End If
							sBody = sBody & ">" & rsget("code_desc") & "</label>&nbsp;&nbsp;"
						End If
					End If
				rsget.MoveNext
				i = i + 1
				Loop
				rsget.Close
			End If
		Else
		'### sUse = "v" view 용
			strSql = " SELECT code_desc From [db_partner].[dbo].[tbl_cooperate_comCode] WHERE code_type ='"&sType&"' AND code_value = '" & sCode & "' AND code_useyn = 'Y'"
			rsget.Open strSql,dbget
			If Not rsget.Eof Then
				sBody = rsget(0)
			End If
			rsget.Close
		End If
		CommonCode = sBody
	End Function


	'####### 직원 연락처 Get. SMS 발송용. #######
	public Function fnGetMemberHp(id)
		Dim strSql
		strSql = "	SELECT isNull(usercell,'0') AS manager_hp FROM [db_partner].[dbo].tbl_user_tenbyten WHERE userid = '" & id & "' and userid <> '' "
		rsget.Open strSql,dbget,1
		'response.write strSql
		IF not rsget.EOF THEN
			If rsget("manager_hp") = "" Then
				fnGetMemberHp = "0"
			Else
				fnGetMemberHp = rsget("manager_hp")
			End If
		Else
			fnGetMemberHp = "0"
		END IF
		rsget.close
	End Function


	'####### 직원 연락처 Get. 웹훅 발송용. #######
	public Function fnGetMemberEmail(id)
		Dim strSql
		strSql = "	SELECT isNull(usermail,'') AS email FROM [db_partner].[dbo].tbl_user_tenbyten WHERE userid = '" & id & "' and userid <> '' "
		rsget.Open strSql,dbget,1

		IF not rsget.EOF THEN
			If rsget("email") = "" Then
				fnGetMemberEmail = ""
			Else
				fnGetMemberEmail = rsget("email")
			End If
		Else
			fnGetMemberEmail = ""
		END IF
		rsget.close
	End Function


	Public Function MyTeamDocTypeExpl
		Dim vBody, vTemp_MyPart, vArr, i

		vTemp_MyPart = g_MyPart
		If g_TeamJang = "x" Then
			vArr = Split(g_MyTeam,",")
		Else
			vArr = Split(g_MyPart,",")
		End If

		For i = 0 To UBOUND(vArr)
			vTemp_MyPart = vArr(i)
			Select Case vTemp_MyPart
				Case "7"	'시스템
					vBody = vBody & "<br>시스템팀에 관한 업무구분 설명은 없습니다.<br>"
				Case "8"	'경영지원
					vBody = vBody & "<br>경영지원팀에 관한 업무구분 설명은 없습니다.<br>"
				Case "10"	'고객센터
					vBody = vBody & "<br>- 이벤트 문의<br>&nbsp;&nbsp;&nbsp;☞ <b>시스템개발및수정,디자인개발및수정,<br>&nbsp;&nbsp;&nbsp;데이터요청및문의,상품관련요청및문의</b><br>"
					vBody = vBody & "- 이미지&상품설명오류 ☞ <b>디자인개발및수정</b><br>- 품절 및 업체배송관련 ☞ <b>상품관련요청및문의</b><br>- 고객의 소리 ☞ <b>모든 항목 해당</b><br>"
					vBody = vBody & "- 회원관련문의 ☞ <b>모든 항목 해당</b><br>"
				Case "11"	'MD
					vBody = vBody & "<br>- 오류정정<br>&nbsp;&nbsp;&nbsp;☞ <b>시스템개발및수정,디자인개발및수정</b><br>- 일부수정/개선<br>&nbsp;&nbsp;&nbsp;☞ <b>시스템개발및수정,디자인개발및수정</b><br>"
					vBody = vBody & "- 신규개발<br>&nbsp;&nbsp;&nbsp;☞ <b>시스템개발및수정,디자인개발및수정</b><br>- 이벤트 관련 ☞ <b>모든 항목 해당</b><br>"
				Case "12"	'웹디
					vBody = vBody & "<br>웹디자인파트에 관한 업무구분 설명은 없습니다.<br>"
				Case "13"	'오프라인
					vBody = vBody & "<br>- 브랜드 정보수정 ☞ <b>시스템개발및수정</b><br>- 시스템 개발 및 수정 ☞ <b>시스템개발및수정</b><br>- 정보 요청 ☞ <b>데이터요청및문의</b><br>"
					vBody = vBody & "- Shop별 주문관련 ☞ <b>매장관련요청및문의</b><br>- 상품관련 ☞ <b>상품관련요청및문의</b><br>"
					vBody = vBody & "- 사이트 관련 요청<br>&nbsp;&nbsp;&nbsp;☞ <b>시스템개발및수정,디자인개발및수정</b><br>- 세금계산서 발행 요청 ☞ <b>경영지원관련업무</b><br>"
					vBody = vBody & "- 상품 제작관련 요청 ☞ <b>상품관련요청및문의</b><br>- 프로모션관련 협조 요청 ☞ <b>사업및업무제안</b><br>"
				Case "14"	'마케팅
					vBody = vBody & "<br>- [제휴]상품및이벤트 협조 ☞ <b>사업및업무제안</b><br>- 업무공지 ☞ <b>기타공지사항</b><br>- 상품기획 ☞ <b>상품관련요청및문의</b><br>"
				Case "15"	'아이띵소
					vBody = vBody & "<br>- 히치하이커 관련 ☞ <b>상품관련요청및문의</b><br>- 인쇄물 제작의뢰 ☞ <b>상품관련요청및문의</b><br>- 정산내역 결제요청 ☞ <b>경영지원관련업무</b><br>"
					vBody = vBody & "- 세금계산서 발행 요청 ☞ <b>경영지원관련업무</b><br>- 선결제 요청 ☞ <b>경영지원관련업무</b><br>- 기타 결제 요청 ☞ <b>경영지원관련업무</b><br>"
					vBody = vBody & "- 정기 메일링<br>&nbsp;&nbsp;&nbsp;☞ <b>시스템개발및수정,디자인개발및수정</b><br>- 페이지 수정<br>&nbsp;&nbsp;&nbsp;☞ <b>시스템개발및수정,디자인개발및수정</b><br>"
					vBody = vBody & "- 이벤트 개발및요청<br>&nbsp;&nbsp;&nbsp;☞ <b>시스템개발및수정,디자인개발및수정</b><br>- 어드민 개발및수정 ☞ <b>시스템개발및수정</b><br>"
					vBody = vBody & "- 상품 입고 확인 ☞ <b>상품관련요청및문의</b><br>- 상품출고 및 재고이동 ☞ <b>상품관련요청및문의</b><br>"
				Case Else
					vBody = vBody & "<br>"
			End Select
		Next
		MyTeamDocTypeExpl = vBody
	End Function


	Function NaviTitle(status)
		SELECT CASE status
			CASE "0"
				NaviTitle = "전체리스트"
			CASE "1"
				NaviTitle = "기안"
			CASE "2"
				NaviTitle = "작업중"
			CASE "3"
				NaviTitle = "작업완료"
			CASE "4"
				NaviTitle = "반려"
			CASE "5"
				NaviTitle = "반려 후 최종완료"
			CASE "6"
				NaviTitle = "참조"
			CASE ELSE
				NaviTitle = ""
		END SELECT
	End Function


	' 현재 페이지 URL에서 파일명 뽑기
	Function GetFileName()
		On Error Resume Next
		Dim vUrl			'/소스 경로저장 변수
		Dim FullFilename		'파일이름
		Dim strName			'확장자를 제외한 파일이름

		vUrl = Request.ServerVariables("SCRIPT_NAME")
		FullFilename = mid(vUrl,instrrev(vUrl,"/")+1)
		strName = Mid(FullFilename, 1, Instr(FullFilename, ".") - 1)

		GetFileName = strName
	End Function


	Function fnProgramWriteCount(didx)
		Dim strSql, vTemp
		vTemp = 0
		If didx > 0 Then
			strSql = "SELECT count(pidx) FROM [db_board].[dbo].tbl_program_change WHERE doc_idx = '" & didx & "' "
			rsget.Open strSql,dbget,1
			IF rsget(0) > 0 THEN
				vTemp = 1
			END IF
			rsget.close
		END IF
		fnProgramWriteCount = vTemp
	End Function
%>
