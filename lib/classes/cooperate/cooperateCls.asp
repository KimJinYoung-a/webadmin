<%
'####################################################
' Description :  ��������
' History : ���ر� ����
'			2022.07.11 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'####################################################

	Class CCooperate
		public FCPage	'Set ���� ������
		public FPSize	'Set ������ ������
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



		'####### ������������Ʈ #######
		public Function fnGetCooperateList
			Dim strSql,iDelCnt, strSubSql
			If FDoc_IsRefer = "o" Then		'####### �˾���. ������, �۾��� ���� �и�
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
				'2014.03.06 ǰ�Ǽ����� ������ �߰�
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


		'####### �����ۼ��Ѿ������� #######
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
				'2014.03.06 ǰ�Ǽ����� ������ �߰�
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


		'####### ������ ���� #######
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



		'####### �������亯����Ʈ #######
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


		'####### �������亯���� #######
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


		'##### �μ� ��� ���� #####
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


		'####### ��������Ʈ #######
		'//common/offshop/member/PoptenbytenuserList.asp
		public Function fnGetMemberList
			Dim strSql, addsql
			'2016-06-29 ������ ����, AND (D.department_id = '" &FRectDepartmentID & "' or D.username like '%" & FRectWorker & "%' ) �̷� �������� ���� ��, �����ý� like�˻����� ���� ��ü �ο� ��µ�
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

		'####### ������÷�����ϸ���Ʈ #######
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


		'####### ������ �� ���ο� ������ ���� #######
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


		'####### POPUP ���� ������ ��ü ī��Ʈ #######
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


	'####### �ڵ� �Ŵ��� ��������. ����Ʈ. #######
	Class CCoopCommonCode
	public FCodeType
	public FCodeValue
	public FCodeDesc
	public FCodeUsing
	public FCodeSort

		'####### �����ڵ� ����Ʈ #######
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

		'####### ������ �ڵ� ���� �������� #######
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



	'####### �α� ����. #######
	Public Function LogInsert(ByVal iDoc_idx, ByVal sLog_Type, ByVal sAction)
		Dim strSql
		SELECT CASE sAction
			CASE "1"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), ������ �ۼ�"
			CASE "2"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), ������ ����"
			CASE "3"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), ������ ����"
			CASE "4"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), ������ Ȯ��"
			CASE "5"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " ������ �亯 �ۼ�"
			CASE "6"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " ������ �亯 ����"
			CASE "7"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " ������ �亯 ����"
			CASE "8"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " ������ �۾��ڿ��� SMS ����"
			CASE "9"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), No." & iDoc_idx & " ������ �����ڿ��� SMS ����"
			CASE "91"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), �ڵ� ����"
			CASE "92"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), �ڵ� ����"
			CASE "93"
				sAction = "" & session("ssBctCname") & "(" & session("ssBctId") & "), �ڵ� ����"
			CASE ELSE
				sAction = "��Ÿ Action"
		END SELECT
		strSql = " INSERT INTO [db_partner].[dbo].tbl_cooperate_log(doc_idx, log_type, log_content, log_ip, log_regdate) " & _
				 "	VALUES('" & iDoc_idx & "','" & sLog_Type & "','" & sAction & "','" & Request.ServerVariables("REMOTE_ADDR") & "', getdate()) "
		dbget.execute strSql
	End Function


	'####### �ڽ����� �� ������ �� ó�� Ȯ�ν� ���� �ð� ����. �α׵� ó�� Ȯ���϶� �Է�. (update:2, ������ Ȯ��) #######
	Public Function WorkerView(ByVal iDoc_idx)
		Dim strSql
		strSql = " DECLARE @ISNULL datetime " & _
				 "	IF EXISTS(SELECT worker_viewdate FROM [db_partner].[dbo].tbl_cooperate_worker WHERE doc_idx = '" & iDoc_idx & "' AND worker_id = '" & session("ssBctId") & "') " & _
				 "	BEGIN " & _
				 "		SELECT @ISNULL = worker_viewdate FROM [db_partner].[dbo].tbl_cooperate_worker WHERE doc_idx = '" & iDoc_idx & "' AND worker_id = '" & session("ssBctId") & "' " & _
				 "		If @ISNULL is NULL " & _
				 "			BEGIN " & _
				 "			UPDATE [db_partner].[dbo].tbl_cooperate_worker SET worker_viewdate = getdate() WHERE doc_idx = '" & iDoc_idx & "' AND worker_id = '" & session("ssBctId") & "' " & _
				 "			INSERT INTO [db_partner].[dbo].tbl_cooperate_log(doc_idx, log_type, log_content, log_ip, log_regdate) VALUES('" & iDoc_idx & "', '2', '" & session("ssBctCname") & "(" & session("ssBctId") & "), ������ Ȯ��', '" & Request.ServerVariables("REMOTE_ADDR") & "', getdate()) " & _
				 "			END " & _
				 "	END "
		dbget.execute strSql
	End Function



	'####### ���� Ȯ���� ����. �󼼳����� �� �۾�Ȯ���� ����� ����. #######
	Public Function ReferView(ByVal iDoc_idx)
		Dim strSql
		strSql = " DECLARE @ISNULL datetime " & _
				 "	IF EXISTS(SELECT refer_viewdate FROM [db_partner].[dbo].tbl_cooperate_refer WHERE doc_idx = '" & iDoc_idx & "' AND refer_id = '" & session("ssBctId") & "') " & _
				 "	BEGIN " & _
				 "		SELECT @ISNULL = refer_viewdate FROM [db_partner].[dbo].tbl_cooperate_refer WHERE doc_idx = '" & iDoc_idx & "' AND refer_id = '" & session("ssBctId") & "' " & _
				 "		If @ISNULL is NULL " & _
				 "			BEGIN " & _
				 "			UPDATE [db_partner].[dbo].tbl_cooperate_refer SET refer_viewdate = getdate() WHERE doc_idx = '" & iDoc_idx & "' AND refer_id = '" & session("ssBctId") & "' " & _
				 "			INSERT INTO [db_partner].[dbo].tbl_cooperate_log(doc_idx, log_type, log_content, log_ip, log_regdate) VALUES('" & iDoc_idx & "', '2', '" & session("ssBctCname") & "(" & session("ssBctId") & "), ������ Ȯ��', '" & Request.ServerVariables("REMOTE_ADDR") & "', getdate()) " & _
				 "			END " & _
				 "	END "
		dbget.execute strSql
	End Function



	'####### �ڵ� select �ڽ� #######
	Sub sbOptCodeType(ByVal selCodeType)
%>
		<option value="doc_status" <%IF Cstr(selCodeType)="doc_status" THEN%>selected<%END IF%>>������°�</option>
		<option value="doc_type" <%IF Cstr(selCodeType)="doc_type" THEN%>selected<%END IF%>>��������</option>
		<option value="doc_important" <%IF Cstr(selCodeType)="doc_important" THEN%>selected<%END IF%>>�����߿䵵</option>
		<option value="doc_difficult" <%IF Cstr(selCodeType)="doc_difficult" THEN%>selected<%END IF%>>�������̵�</option>
		<option value="ans_type" <%IF Cstr(selCodeType)="ans_type" THEN%>selected<%END IF%>>�亯����</option>
		<option value="log_type" <%IF Cstr(selCodeType)="log_type" THEN%>selected<%END IF%>>�αױ���</option>
<%
	End Sub


	'####### �ڵ�Ŵ��� ������ �� �ϳ��� �ҷ����� ���� �ڵ� ����. write ��, view ��. #######
	Public Function CommonCode(ByVal sUse, ByVal sType, ByVal sCode)
		Dim strSql, sBody, i, vTemp
		sBody = ""
		i = 0
		vTemp = sUse
		If sUse = "wnew" Then
			sUse = "w"
		End If

		'### sUse = "w" write ��
		If sUse = "w" Then
			If vTemp = "wnew" Then
				strSql = " SELECT code_value, code_desc From [db_partner].[dbo].[tbl_cooperate_comCode] WHERE code_type ='"&sType&"' AND code_value Not IN('4','5') AND code_useyn = 'Y' ORDER BY code_sort ASC"
			Else
				strSql = " SELECT code_value, code_desc From [db_partner].[dbo].[tbl_cooperate_comCode] WHERE code_type ='"&sType&"' AND code_useyn = 'Y' ORDER BY code_sort ASC"
			End IF

			If sType = "doc_status" AND sCode = "" Then
				sBody = "<input type='hidden' name='doc_status' value='1'>������ �ۼ�"
			Else
				rsget.Open strSql,dbget,1
				Do Until rsget.Eof
					'####### ���������� select�ڽ� ������. �������� radio ��. #######
					If sType = "doc_type" Then
						If i = 0 Then
							sBody = "<select name='doc_type' class='select' onChange='issystem(this.value)'>"
							If GetFileName() = "index" OR GetFileName() = "my_cooperate" Then
								sBody = sBody & "<option value=''>-��û����-</option> "
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
							'####### index �������� ������� �˻���. select�ڽ�. #######
							If i = 0 Then
								sBody = "<select name='doc_status' class='select'><option value='x'>-ó������-</option><option value='0' "
								If CStr(Replace(sCode,"s","")) = CStr(0) Then
									sBody = sBody & "selected"
								End If
								sBody = sBody & ">��ó�� ��ü</option>"
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
		'### sUse = "v" view ��
			strSql = " SELECT code_desc From [db_partner].[dbo].[tbl_cooperate_comCode] WHERE code_type ='"&sType&"' AND code_value = '" & sCode & "' AND code_useyn = 'Y'"
			rsget.Open strSql,dbget
			If Not rsget.Eof Then
				sBody = rsget(0)
			End If
			rsget.Close
		End If
		CommonCode = sBody
	End Function


	'####### ���� ����ó Get. SMS �߼ۿ�. #######
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


	'####### ���� ����ó Get. ���� �߼ۿ�. #######
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
				Case "7"	'�ý���
					vBody = vBody & "<br>�ý������� ���� �������� ������ �����ϴ�.<br>"
				Case "8"	'�濵����
					vBody = vBody & "<br>�濵�������� ���� �������� ������ �����ϴ�.<br>"
				Case "10"	'������
					vBody = vBody & "<br>- �̺�Ʈ ����<br>&nbsp;&nbsp;&nbsp;�� <b>�ý��۰��߹׼���,�����ΰ��߹׼���,<br>&nbsp;&nbsp;&nbsp;�����Ϳ�û�׹���,��ǰ���ÿ�û�׹���</b><br>"
					vBody = vBody & "- �̹���&��ǰ������� �� <b>�����ΰ��߹׼���</b><br>- ǰ�� �� ��ü��۰��� �� <b>��ǰ���ÿ�û�׹���</b><br>- ���� �Ҹ� �� <b>��� �׸� �ش�</b><br>"
					vBody = vBody & "- ȸ�����ù��� �� <b>��� �׸� �ش�</b><br>"
				Case "11"	'MD
					vBody = vBody & "<br>- ��������<br>&nbsp;&nbsp;&nbsp;�� <b>�ý��۰��߹׼���,�����ΰ��߹׼���</b><br>- �Ϻμ���/����<br>&nbsp;&nbsp;&nbsp;�� <b>�ý��۰��߹׼���,�����ΰ��߹׼���</b><br>"
					vBody = vBody & "- �ű԰���<br>&nbsp;&nbsp;&nbsp;�� <b>�ý��۰��߹׼���,�����ΰ��߹׼���</b><br>- �̺�Ʈ ���� �� <b>��� �׸� �ش�</b><br>"
				Case "12"	'����
					vBody = vBody & "<br>����������Ʈ�� ���� �������� ������ �����ϴ�.<br>"
				Case "13"	'��������
					vBody = vBody & "<br>- �귣�� �������� �� <b>�ý��۰��߹׼���</b><br>- �ý��� ���� �� ���� �� <b>�ý��۰��߹׼���</b><br>- ���� ��û �� <b>�����Ϳ�û�׹���</b><br>"
					vBody = vBody & "- Shop�� �ֹ����� �� <b>������ÿ�û�׹���</b><br>- ��ǰ���� �� <b>��ǰ���ÿ�û�׹���</b><br>"
					vBody = vBody & "- ����Ʈ ���� ��û<br>&nbsp;&nbsp;&nbsp;�� <b>�ý��۰��߹׼���,�����ΰ��߹׼���</b><br>- ���ݰ�꼭 ���� ��û �� <b>�濵�������þ���</b><br>"
					vBody = vBody & "- ��ǰ ���۰��� ��û �� <b>��ǰ���ÿ�û�׹���</b><br>- ���θ�ǰ��� ���� ��û �� <b>����׾�������</b><br>"
				Case "14"	'������
					vBody = vBody & "<br>- [����]��ǰ���̺�Ʈ ���� �� <b>����׾�������</b><br>- �������� �� <b>��Ÿ��������</b><br>- ��ǰ��ȹ �� <b>��ǰ���ÿ�û�׹���</b><br>"
				Case "15"	'���̶��
					vBody = vBody & "<br>- ��ġ����Ŀ ���� �� <b>��ǰ���ÿ�û�׹���</b><br>- �μ⹰ �����Ƿ� �� <b>��ǰ���ÿ�û�׹���</b><br>- ���곻�� ������û �� <b>�濵�������þ���</b><br>"
					vBody = vBody & "- ���ݰ�꼭 ���� ��û �� <b>�濵�������þ���</b><br>- ������ ��û �� <b>�濵�������þ���</b><br>- ��Ÿ ���� ��û �� <b>�濵�������þ���</b><br>"
					vBody = vBody & "- ���� ���ϸ�<br>&nbsp;&nbsp;&nbsp;�� <b>�ý��۰��߹׼���,�����ΰ��߹׼���</b><br>- ������ ����<br>&nbsp;&nbsp;&nbsp;�� <b>�ý��۰��߹׼���,�����ΰ��߹׼���</b><br>"
					vBody = vBody & "- �̺�Ʈ ���߹׿�û<br>&nbsp;&nbsp;&nbsp;�� <b>�ý��۰��߹׼���,�����ΰ��߹׼���</b><br>- ���� ���߹׼��� �� <b>�ý��۰��߹׼���</b><br>"
					vBody = vBody & "- ��ǰ �԰� Ȯ�� �� <b>��ǰ���ÿ�û�׹���</b><br>- ��ǰ��� �� ����̵� �� <b>��ǰ���ÿ�û�׹���</b><br>"
				Case Else
					vBody = vBody & "<br>"
			End Select
		Next
		MyTeamDocTypeExpl = vBody
	End Function


	Function NaviTitle(status)
		SELECT CASE status
			CASE "0"
				NaviTitle = "��ü����Ʈ"
			CASE "1"
				NaviTitle = "���"
			CASE "2"
				NaviTitle = "�۾���"
			CASE "3"
				NaviTitle = "�۾��Ϸ�"
			CASE "4"
				NaviTitle = "�ݷ�"
			CASE "5"
				NaviTitle = "�ݷ� �� �����Ϸ�"
			CASE "6"
				NaviTitle = "����"
			CASE ELSE
				NaviTitle = ""
		END SELECT
	End Function


	' ���� ������ URL���� ���ϸ� �̱�
	Function GetFileName()
		On Error Resume Next
		Dim vUrl			'/�ҽ� ������� ����
		Dim FullFilename		'�����̸�
		Dim strName			'Ȯ���ڸ� ������ �����̸�

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
