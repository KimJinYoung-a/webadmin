<%
'####################################################
' Description :  공지사항
' History : 2011.02.28 김진영 생성
'           2018.07.12 한용민 수정
'####################################################

Class board_item
	Public Fbrd_subject
	Public Fbrd_hit
	Public Fbrd_regdate
	Public Fbrd_fixed
	Public Fbrd_sn
	Public Fbrd_username
	Public Fbrd_team
	Public Fbrd_type
	Public Fcnt

	Public Fevt_code
	Public FpartMDid
	Public Fevt_name
	Public Fevt_winner
	Public Fevt_regdate
	Public Fevt_laterdate
End Class

Class CBoardScmItem
	Public Fidx
	Public FscheduleDate
	Public Ftitle
	Public Fcontents
	Public Freguserid
	Public Fmodiuserid
	Public Fisusing
	Public Fdispno
	Public Fregdate
	Public Flastupdate
End Class

Class board
	Public FbrdList()
	Public FcmtList()
	Public Fbrd_sn
	Public FRegUserID
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount
	Public Fbrd_subject
	Public Fbrd_content
	Public Fbrd_hit
	Public Fbrd_regdate
	Public Fbrd_fixed
	Public Fbrd_team
	Public Fbrd_isusing
	Public Fbrd_type
	Public Fpart_sn
	Public FJob_sn
	Public FJob_name
	Public FPosit_name
	Public FPositsn
	Public Fusername
	Public Fbid
	Public Fbrd_username
	Public Flevel_name
	Public FAdminlsn
	Public FPartpsn
	public Fdepartment_id
	Public Fsearch_team
	Public Fsearch_type
	Public FisMine
	Public Fdetail_search
	Public Fsearchstr
	Public FstartDate
	Public FendDate

	Public Fcmt_sn
	Public FOnecmt
	Public Fid
	Public Fcmt_content
	Public Fcmt_regdate
	Public Fjob

	public FisAuth
	public frectprizeyn
    public fposit_sn
	public FIsNotify

	Public Sub fnBoardcontent
		Dim strSql
		strSql = "select max(brd_sn) as brd_sn from db_partner.dbo.tbl_cooperate_board"
		rsget.open strSql, dbget, 1

		If not rsget.EOF Then
			Fbrd_sn = rsget("brd_sn")
		End If

		If IsNull(rsget("brd_sn")) Then
			Fbrd_sn = "0"
		End If

		rsget.close
	End Sub


	Public Function fnBoardlist
		Dim strSql, i, where

		''우정원 요청 인사팀 lovesay999 이냥반 인경우 계약직이라 정직원 공지를 못본다고 함.. (인사팀에서 올린공지를 볼수 있게.)
		dim inSaUidARR : inSaUidARR = "wahahajw,lovesay999,jhw7980,icommang"
		if (InStr((","&inSaUidARR&","),","&session("ssBctId")&",")>0) then
		    if (FPositsn>8) then
		        FPositsn=8 ''사원 보다 직급이 낮으면 사원으로 ..
		        rw "::"&FPositsn&"::"
		    end if
		end if

		' 열람선택 검색 부 where절'
		If Fsearch_team = "all" Then
			where = " and (p.part_sn = 1 or (p.part_sn = 0 and p.department_id = '0'))"
		ElseIf Fsearch_team = "team" Then
			where = " and (p.part_sn <> 1 or   (p.part_sn = 0 and p.department_id <> '0' ))"
		ElseIf Fsearch_team = "" Then
			where = ""
		End If

		' 상세 검색 부 where절'
		If Fdetail_search = "subject" Then
			where = where & " and B.brd_subject like '%"&Fsearchstr&"%' "
			FAdminlsn = "7777"
		ElseIf Fdetail_search = "content" Then
			where = where & " and B.brd_content like '%"&Fsearchstr&"%' "
			FAdminlsn = "7777"
		ElseIf Fdetail_search = "writer" Then
			where = where & " and T.username like '%"&Fsearchstr&"%' "
			FAdminlsn = "7777"
		ElseIf Fdetail_search = "" Then
			where = where
		End If

		If Fdetail_search = "" Then
			If FJob_sn <> "0" Then
				Fjob = " and (p.job_sn >= '" & FJob_sn & "' or P.job_sn = '0' ) and (('" & FAdminlsn & "' = '1' or '"&FPositsn&"' = '2') or (p.part_sn = '1' or p.part_sn = '" & FPartpsn & "') or (p.part_sn = 0 and (p.department_id = 0 or p.department_id = '" & Fdepartment_id & "'))) "
			Else
				Fjob = " and (p.job_sn <= '" & FJob_sn & "') and (('" & FAdminlsn & "' = '1' or '"&FPositsn&"' = '2') or (p.part_sn = '1' or p.part_sn = '" & FPartpsn & "') or (p.part_sn = 0 and (p.department_id = 0 or p.department_id = '" & Fdepartment_id & "'))) "
			End If
		End If
		If Fsearch_type <> "" Then
			where = where & " and B.brd_type = '" & Fsearch_type & "' "
			If Fjob <> "" Then
				Fjob = ""
			End IF
		End IF
		If FPositsn <> "" Then
			'/ Assistant는 사원과 동급으로 처리한다.
			if FPositsn="11" then
				where = where & " and p.posit_sn >= '8'"
			else
				where = where & " and p.posit_sn >= '"&FPositsn&"'"
			end if
		End IF

'		strSql = "select count(distinct B.brd_sn) as cnt " + vbcrlf
'		strSql = strSql & " from db_partner.dbo.tbl_cooperate_board as B " & vbcrlf
'		strSql = strSql & " Inner Join db_partner.dbo.tbl_cooperate_board_part as P on B.brd_sn = P.brd_sn and B.brd_isusing = 'N'  " & vbcrlf
'		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on B.id = T.userid " & vbcrlf
'		strSql = strSql & " where B.brd_isusing = 'N' "&Fjob&"  " & vbcrlf
'		strSql = strSql & where
		'response.write strSql & "<br>"
'		rsget.Open strSql,dbget,1
		strSql = "EXEC [db_partner].[dbo].[sp_Ten_cooperate_board_list] 'count','" & (FPageSize * FCurrPage) & "','" & Fsearch_team & "','" & Fdetail_search & "',"
		strSql = strSql & "'" & Fsearchstr & "','" & FJob_sn & "','" & FAdminlsn & "','" & FPositsn & "','" & FPartpsn & "','" & Fdepartment_id & "','" & Fsearch_type & "' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly

		If not rsget.EOF Then
			FTotalCount = rsget("cnt")
		Else
			FTotalCount = 0
		End If
		rsget.Close

'		strSql = "select top "& Cstr(FPageSize * FCurrPage) &" B.brd_sn, T.username, B.brd_subject, B.brd_hit, B.brd_regdate, B.brd_fixed, B.brd_team, (select count(*) from db_partner.dbo.tbl_cooperate_board_comment where brd_sn = B.brd_sn) as cnt " & vbcrlf
'		strSql = strSql & " from db_partner.dbo.tbl_cooperate_board as B " & vbcrlf
'		strSql = strSql & " Inner Join db_partner.dbo.tbl_cooperate_board_part as P on B.brd_sn = P.brd_sn and B.brd_isusing = 'N'  " & vbcrlf
'		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on B.id = T.userid " & vbcrlf
'		strSql = strSql & " Left Join db_partner.dbo.tbl_cooperate_board_comment as C on B.brd_sn = C.brd_sn " & vbcrlf
'		strSql = strSql & " where B.brd_isusing = 'N' "&Fjob&"  " & vbcrlf
'		strSql = strSql & where
'		strSql = strSql & " group by T.username, B.brd_sn, B.brd_subject, B.brd_hit, B.brd_regdate, B.brd_fixed, B.brd_team  " & vbcrlf
'		strSql = strSql & " order by B.brd_fixed asc, B.brd_sn desc"
'		response.write strSql & "<br>"
		strSql = "EXEC [db_partner].[dbo].[sp_Ten_cooperate_board_list] 'list','" & (FPageSize * FCurrPage) & "','" & Fsearch_team & "','" & Fdetail_search & "',"
		strSql = strSql & "'" & Fsearchstr & "','" & FJob_sn & "','" & FAdminlsn & "','" & FPositsn & "','" & FPartpsn & "','" & Fdepartment_id & "','" & Fsearch_type & "' "
'		response.write strSql & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1
		Redim preserve FbrdList(FResultCount)
		FPageCount = FCurrPage - 1

		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FbrdList(i) = new board_item
				FbrdList(i).Fbrd_sn 		= rsget("brd_sn")
				FbrdList(i).Fbrd_username	= rsget("username")
				FbrdList(i).Fbrd_subject	= rsget("brd_subject")
				FbrdList(i).Fbrd_hit 		= rsget("brd_hit")
				FbrdList(i).Fbrd_regdate 	= rsget("brd_regdate")
				FbrdList(i).Fbrd_fixed 		= rsget("brd_fixed")
				FbrdList(i).Fbrd_team 		= rsget("brd_team")
				FbrdList(i).Fcnt 			= rsget("cnt")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Function

	Public Function fnBoardview
		Dim strSql, i

		strSql = " select Top 1 B.brd_sn, B.id, B.brd_subject, B.brd_content, B.brd_hit, B.brd_regdate, B.brd_fixed, B.brd_team" & vbcrlf
		strSql = strSql & " , B.brd_type, P.part_sn,p.department_id, T.username, J.job_name, F.posit_name, F.posit_sn, B.isNotify " & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_cooperate_board as B " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_cooperate_board_part as P on B.brd_sn = P.brd_sn  " & vbcrlf
		strSql = strSql & " Left Join db_partner.dbo.tbl_JobInfo as J on P.job_sn = J.job_sn  " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_positinfo as F on P.posit_sn = F.posit_sn " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on T.userid = B.id " & vbcrlf
		strSql = strSql & " where B.brd_sn = '"& Fbrd_sn &"'" & vbcrlf

		'response.write strSql & "<br>"
		rsget.Open strSql,dbget,1

        FResultCount = rsget.recordcount
        FtotalCount = rsget.recordcount

		If not rsget.EOF Then
			Fbrd_sn 		= rsget("brd_sn")
			Fbid 			= rsget("id")
			Fbrd_subject	= rsget("brd_subject")
			Fbrd_content 	= rsget("brd_content")
			Fbrd_hit 		= rsget("brd_hit")
			Fbrd_regdate 	= rsget("brd_regdate")
			Fbrd_fixed 		= rsget("brd_fixed")
			Fbrd_team 		= rsget("brd_team")
			Fpart_sn 		= rsget("part_sn")
			Fdepartment_id  = rsget("department_id")
			Fusername 		= rsget("username")
			FJob_name		= rsget("job_name")
			FPosit_name		= rsget("posit_name")
			Fbrd_type		= rsget("brd_type")
			fposit_sn		= rsget("posit_sn")
			FIsNotify		= rsget("isNotify")
		End If
		rsget.Close
	End Function

	Public Function fnBoardmodify
		Dim strSql, i
		strSql = ""
		strSql = strSql & " select top 1 B.id, T.username, B.brd_subject, B.brd_content, B.brd_hit, B.brd_regdate, B.brd_fixed, B.brd_team, B.brd_type, P.level_sn, P.posit_sn, P.job_sn, B.brd_isusing, P.department_id, B.startDate, B.endDate, B.isNotify " & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_cooperate_board as B " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_cooperate_board_part as P on B.brd_sn = P.brd_sn  " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on B.id = T.userid " & vbcrlf
		strSql = strSql & " where B.brd_sn = '"& Fbrd_sn &"'"& vbcrlf
		if not FisAuth then
		strSql = strSql & "		 AND B.id = '" & FRegUserID & "'" & vbcrlf
		end if
	'	 response.write strSql
		rsget.Open strSql,dbget,1

		If not rsget.EOF Then
			Fbid 			= rsget("id")
			Fbrd_username	= rsget("username")
			Fbrd_subject	= rsget("brd_subject")
			Fbrd_content 	= rsget("brd_content")
			Fbrd_hit 		= rsget("brd_hit")
			Fbrd_regdate 	= rsget("brd_regdate")
			Fbrd_fixed 		= rsget("brd_fixed")
			Fbrd_team 		= rsget("brd_team")
			FPositsn 		= rsget("posit_sn")
			FJob_sn			= rsget("job_sn")
			Fbrd_isusing	= rsget("brd_isusing")
			Fbrd_type		= rsget("brd_type")
            Fdepartment_id  = rsget("department_id")
			FstartDate		= rsget("startDate")
			FendDate		= rsget("endDate")
			FIsNotify		= rsget("isNotify")
			FResultCount = 1
		Else
			FResultCount = 0
		End If
		rsget.Close
	End Function

	public Function fnGetDepartmentid
	    dim strSql
	    strSql = " select department_id from db_partner.dbo.tbl_cooperate_board_part where brd_sn ='"& Fbrd_sn &"'"
	    rsget.Open strSql,dbget,1
	    If not rsget.EOF Then
	        fnGetDepartmentid = rsget.getRows()
	    end if
	    rsget.Close
    end function

	Public Sub fnBoardreplylist
		Dim strSql

		strSql = ""
		strSql = strSql & " select C.cmt_sn, C.id, C.cmt_content, C.cmt_regdate, C.brd_sn, T.username, p.part_sn " & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_cooperate_board_comment as C " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on C.id = T.userid   " & vbcrlf
		strSql = strSql & " left Join db_partner.dbo.tbl_partner as p on C.id = p.id   " & vbcrlf
		strSql = strSql & " where brd_sn = '"& Fbrd_sn &"' " & vbcrlf
		strSql = strSql & " ORDER BY C.cmt_sn ASC" & vbcrlf

		rsget.Open strSql,dbget,1

		FResultCount = rsget.recordcount
		FTotalCount = FResultCount

		Redim preserve FcmtList(FResultCount)
		FPageCount = FCurrPage - 1

		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				set FcmtList(i) = new board

				FcmtList(i).Fcmt_sn 		= rsget("cmt_sn")
				FcmtList(i).Fid 			= rsget("id")
				FcmtList(i).Fcmt_content 	= db2html(rsget("cmt_content"))
				FcmtList(i).Fcmt_regdate 	= rsget("cmt_regdate")
				FcmtList(i).Fbrd_sn 		= db2html(rsget("brd_sn"))
				FcmtList(i).Fusername 		= rsget("username")
				FcmtList(i).fpart_sn 		= rsget("part_sn")

				rsget.Movenext
				i = i + 1
			Loop
		End If
		rsget.Close

	End Sub

    Public Sub fnBoardreplymodify()
		dim strSql
		strSql = ""
		strSql = strSql & " SELECT top 1 cmt_content" & vbcrlf
		strSql = strSql & " FROM db_partner.dbo.tbl_cooperate_board_comment " & vbcrlf
		strSql = strSql & " WHERE cmt_sn = '" & Fcmt_sn & "' AND id = '" & session("ssBctId") & "' "

		rsget.Open strSql, dbget, 1

		Set FOnecmt = new board

		If Not rsget.Eof Then
			FOnecmt.Fcmt_content = rsget("cmt_content")
		End If
		rsget.Close
	End Sub

	Public Function fnGetFileList
		Dim strSql
		strSql = "	SELECT file_idx, file_name, real_name " & _
				"		FROM [db_partner].[dbo].tbl_cooperate_file " & _
				"	WHERE brd_sn = '" & Fbrd_sn & "' " & _
				"	ORDER BY brd_sn ASC "
		rsget.Open strSql,dbget,1
		'response.write strSql
		IF not rsget.EOF THEN
			fnGetFileList = rsget.getRows()
		END IF
		rsget.close
	End Function

	Public Function fnmain_notice_list
		Dim strSql, i

		If FJob_sn <> "0" Then
			Fjob = " (p.job_sn >= '" & FJob_sn & "' or P.job_sn = '0' ) "
		Else
			Fjob = " (p.job_sn <= '" & FJob_sn & "') "
		End If
		'게시글 리스트 구하기'

		strSql = ""
		strSql = strSql & " select top 10 B.brd_sn, T.username, B.brd_subject, B.brd_hit, B.brd_regdate, B.brd_fixed, B.brd_team, (select count(*) from db_partner.dbo.tbl_cooperate_board_comment where brd_sn = B.brd_sn) as cnt  " & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_cooperate_board as B with (nolock)" & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_cooperate_board_part as P with (nolock) on B.brd_sn = P.brd_sn and B.brd_isusing = 'N' " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T with (nolock) on B.id = T.userid " & vbcrlf
		strSql = strSql & " Left Join db_partner.dbo.tbl_cooperate_board_comment as C with (nolock) on B.brd_sn = C.brd_sn " & vbcrlf
		strSql = strSql & " where  ('" & FAdminlsn & "' = '1' or '"&FPositsn&"' = '2') or ('"& C_PSMngPart &"' = 'True') "& vbcrlf	'시스템팀일 때, 는 전부 보임
		strSql = strSql & " or ((p.part_sn = '1' or p.part_sn = '" & FPartpsn & "') or (p.part_sn = 0 and (p.department_id = 0 or p.department_id = '" & Fdepartment_id & "')) )and  "&Fjob&"  " & vbcrlf
		strSql = strSql & " and p.posit_sn >= '"&FPositsn&"' and B.brd_isusing = 'N' " & vbcrlf
		strSql = strSql & " group by T.username, B.brd_sn, B.brd_subject, B.brd_hit, B.brd_regdate, B.brd_fixed, B.brd_team  " & vbcrlf
		strSql = strSql & " order by B.brd_regdate desc"

		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.recordcount
		Redim preserve FbrdList(FResultCount)

		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FbrdList(i) = new board_item
				FbrdList(i).Fbrd_sn 		= rsget("brd_sn")
				FbrdList(i).Fbrd_username	= rsget("username")
				FbrdList(i).Fbrd_subject	= rsget("brd_subject")
				FbrdList(i).Fbrd_hit 		= rsget("brd_hit")
				FbrdList(i).Fbrd_regdate 	= rsget("brd_regdate")
				FbrdList(i).Fbrd_fixed 		= rsget("brd_fixed")
				FbrdList(i).Fbrd_team 		= rsget("brd_team")
				FbrdList(i).Fcnt 			= rsget("cnt")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
		'게시글 리스트 구하기 끝'
	End Function

	Public Function fnmain_event_MDlist
		Dim strSql, i, sqlsearch

        if frectprizeyn<>"" Then
            sqlsearch = sqlsearch & " and prizeyn<>'"& frectprizeyn &"'" & vbcrlf
        end if

		strSql = ""
		strSql = strSql & " Select E.evt_code,D.partMDid, E.evt_name, P.evt_winner, E.evt_regdate, datediff(d,'"& date() &"', E.evt_prizedate) as laterdate " & vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_event as E " & vbcrlf
		strSql = strSql & " Inner Join db_event.dbo.tbl_event_display as D On E.evt_code = D.evt_code " & vbcrlf
		strSql = strSql & " Left Join db_event.dbo.tbl_event_prize as P On E.evt_code = P.evt_code " & vbcrlf
		strSql = strSql & " where ( D.partMDid = '" & session("ssBctId") & "' or '"&session("ssBctId")&"' = 'hrkang97' or '"&session("ssBctId")&"' = 'tozzinet')"  & vbcrlf
		'strSql = strSql & " where D.partMDid = 'monkeytn' "  & vbcrlf
		strSql = strSql & " and E.evt_using = 'Y'" & sqlsearch
		strSql = strSql & " and E.evt_prizedate != '' and P.evt_winner IS NULL "  & vbcrlf
		strSql = strSql & " and E.evt_prizedate <= '"& date()+3 &"' "  & vbcrlf
		strSql = strSql & " and E.evt_prizedate >= '"& date()-60 &"' "  & vbcrlf
		strSql = strSql & " order by E.evt_prizedate ASC"

		'response.write strSql & "<br>"
		rsget.Open strSql,dbget,1
		FResultCount = rsget.recordcount
		Redim preserve FbrdList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FbrdList(i) = new board_item
				FbrdList(i).Fevt_code 		= rsget("evt_code")
				FbrdList(i).FpartMDid		= rsget("partMDid")
				FbrdList(i).Fevt_name		= rsget("evt_name")
				FbrdList(i).Fevt_winner 	= rsget("evt_winner")
				FbrdList(i).Fevt_regdate 	= rsget("evt_regdate")
				FbrdList(i).Fevt_laterdate 	= rsget("laterdate")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Function

	Public Function fnGetScmNoticeList
		Dim strSql, i, sqlsearch

		strSql = ""
		strSql = strSql & " select top 20 * "
		strSql = strSql & " from [db_board].[dbo].[tbl_scm_notice] with (nolock)"
		strSql = strSql & " where 1=1 "
		strSql = strSql & " and isusing = 'Y' "
		strSql = strSql & " order by dispno, idx desc "

		'response.write strSql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordcount
		Redim preserve FbrdList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				'// idx, scheduleDate, title, contents, reguserid, modiuserid, isusing, dispno, regdate, lastupdate
				Set FbrdList(i) = new CBoardScmItem
				FbrdList(i).Fidx 			= rsget("idx")
				FbrdList(i).FscheduleDate 	= db2html(rsget("scheduleDate"))
				FbrdList(i).Ftitle 			= db2html(rsget("title"))
				FbrdList(i).Fcontents 		= db2html(rsget("contents"))
				FbrdList(i).Freguserid 		= rsget("reguserid")
				FbrdList(i).Fmodiuserid 	= rsget("modiuserid")
				FbrdList(i).Fisusing 		= rsget("isusing")
				FbrdList(i).Fdispno 		= rsget("dispno")
				FbrdList(i).Fregdate 		= rsget("regdate")
				FbrdList(i).Flastupdate 	= rsget("lastupdate")

				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Function

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

End Class
%>

<%
	Function DrawPartinfoCombo(checkBoxName,checkboxId, bsn)
	Dim tmp_str, strSql, i, Sn, k
		If checkboxId = "" Then
			strSql = "select part_sn, part_name from db_partner.dbo.tbl_partinfo where part_isDel = 'N' and part_sn not in (1,2,3,6) order by part_sort asc"
			rsget.Open strSql,dbget,1
			i = 0
			If not rsget.EOF Then
				Do Until rsget.EOF
					If Lcase(checkboxId) = Lcase(rsget("part_sn")) Then
						tmp_str = " checked"
					End If
					response.write ("<td><label><input type = 'checkbox' name="& checkBoxName &" value="&rsget("part_sn")&">"&rsget("part_name")&"</label></td>")
					If (i mod 3) = 2  Then response.write "</tr><tr>" End If
					rsget.MoveNext
					i = i + 1
				Loop

			End if
			rsget.close
		Else
		Dim A, B
			strSql = ""
			strSql = strSql & " select O.part_sn as osn , P.part_sn, O.part_name " &vbcrlf
			strSql = strSql & " from db_partner.dbo.tbl_partinfo as O " &vbcrlf
			strSql = strSql & " Left Join db_partner.dbo.tbl_cooperate_board_part as P " &vbcrlf
			strSql = strSql & " on O.part_sn = P.part_sn and P.brd_sn = '"& bsn &"' " &vbcrlf
			strSql = strSql & " where O.part_isDel = 'N' and O.part_sn not in (1,2,3,6) " &vbcrlf

			rsget.Open strSql,dbget,1
			i = 0
			If not rsget.EOF Then
				Do Until rsget.EOF
					A = db2html(rsget("part_sn"))
					B = db2html(rsget("osn"))

					If A = B Then
						tmp_str = " checked "
					Else
						tmp_str = ""
					End If

					response.write ("<td><label><input type = 'checkbox' name='"& checkBoxName &"' value='"&rsget("osn")&"' "& tmp_str &" >"&rsget("part_name")&"</label></td>")
					If (i mod 3) = 2  Then response.write "</tr><tr>" End If
					rsget.MoveNext
					i = i + 1
				Loop

			End if
			rsget.close
		End If
	End Function

	Function DrawJobCombo(selectBoxName,selectedId)
	Dim tmp_str, strSql

%>
	<select name="<%=selectBoxName%>" class="select">
<%
	response.write("<option value='0' selected>일반</option>")
	strSql = "select job_sn, job_name from db_partner.dbo.tbl_JobInfo where job_isDel = 'N'"
	rsget.Open strSql,dbget,1

	If not rsget.EOF Then
		Do Until rsget.EOF
			If rsget("job_sn") = selectedId Then
				tmp_str = " selected"
			End If
			response.write("<option value='"&rsget("job_sn")&"' "&tmp_str&">" + db2html(rsget("job_name")) + "</option>")
			tmp_str = ""
			rsget.MoveNext
		Loop
	End If
	rsget.close

	response.write("</select>")

	End Function

	Function DrawPositCombo(selectBoxName,selectedId)
	Dim tmp_str, strSql
%>
	<select name="<%=selectBoxName%>">
<%
	strSql = "select posit_sn, posit_name from db_partner.dbo.tbl_positinfo where posit_isDel = 'N' and posit_sn not in (14,15)"
	rsget.Open strSql,dbget,1

	If not rsget.EOF Then
		Do Until rsget.EOF
			If selectedId = "" Then
				If rsget("posit_sn") = 8 Then
					tmp_str = " selected"
				End If
			Else
				If rsget("posit_sn") = selectedId Then
					tmp_str = " selected"
				End If
			End If
			response.write("<option value='"&rsget("posit_sn")&"' "&tmp_str&">" + db2html(rsget("posit_name")) + "</option>")
			tmp_str = ""
			rsget.MoveNext
		Loop
	End If
	rsget.close
	response.write("</select>")

	End Function

	Function fnBrdType(page, isAll, value, onchange)
		Dim vBody
		If page = "w" Then	'### write
			vBody = "<select name=""brd_type"" class=""select"" " & onchange & ">"
			If isAll = "Y" Then
				vBody = vBody & "	<option value="""">-선택-</option>"
			End If
			vBody = vBody & "	<option value=""0"" " & CHKIIF(value="0","selected","") & ">일반 안내</option>"
			vBody = vBody & "	<option value=""5"" " & CHKIIF(value="5","selected","") & ">업무 공지</option>"
			vBody = vBody & "	<option value=""1"" " & CHKIIF(value="1","selected","") & ">인사 공지</option>"
			vBody = vBody & "	<option value=""2"" " & CHKIIF(value="2","selected","") & ">경영제도</option>"
			vBody = vBody & "	<option value=""3"" " & CHKIIF(value="3","selected","") & ">회사내규</option>"
			vBody = vBody & "	<option value=""6"" " & CHKIIF(value="6","selected","") & ">보안 공지</option>"
			vBody = vBody & "	<option value=""4"" " & CHKIIF(value="4","selected","") & ">경조사</option>"
			vBody = vBody & "	<option value=""11"" " & CHKIIF(value="11","selected","") & ">인사규정</option>"
			vBody = vBody & "	<option value=""12"" " & CHKIIF(value="12","selected","") & ">근태</option>"
			vBody = vBody & "	<option value=""13"" " & CHKIIF(value="13","selected","") & ">복리후생</option>"
			vBody = vBody & "	<option value=""90"" " & CHKIIF(value="90","selected","") & ">기타</option>"
			vBody = vBody & "</select>"
		ElseIf page = "v" Then	'### view
			Select Case value
				Case "1" : vBody = "인사 공지"
				Case "2" : vBody = "경영제도"
				Case "3" : vBody = "회사내규"
				Case "4" : vBody = "경조사"
				Case "5" : vBody = "업무 공지"
				Case "0" : vBody = "일반 안내"
				Case "6" : vBody = "보안 공지"
				Case "11" : vBody = "인사규정"
				Case "12" : vBody = "근태"
				Case "13" : vBody = "복리후생"
				Case "90" : vBody = "기타"
				Case Else : vBody = ""
			End Select
		End IF
		fnBrdType = vBody
	End Function

	Function getEvtoneWorkname(eC)
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT T.username "
		strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten as T "
		strSql = strSql & " JOIN db_event.dbo.tbl_event_display as D on T.userid = D.partMDid "
		strSql = strSql & " WHERE D.evt_code = '"&eC&"' and T.userid <> '' "
		rsget.Open strSql,dbget
		IF not rsget.eof THEN
			getEvtoneWorkname = rsget("username")
		End IF
		rsget.close
	End Function

	Sub sbEVTGetwork(ByVal selName, ByVal sIDValue, ByVal sScript)
		Dim strSql, arrList, intLoop
		strSql = " SELECT userid, username "
		strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten  "
		strSql = strSql & " WHERE userid = '"&sIDValue&"' and userid <> '' "
		rsget.Open strSql,dbget
		IF not rsget.eof THEN
			arrList = rsget.getRows()
		End IF
		rsget.close

		IF isArray(arrList) THEN
%>
			<input type="text" class="text" name="doc_workername" value="<%=arrList(1,0)%>" size="10" readonly>
			<input type="button" class="button" value="지정" onClick="evtworkerlist('<%=mdlist.FbrdList(i).Fevt_code%>')">
<%
		Else
%>
			<input type="text" class="text" name="doc_workername" value="" size="10" readonly>
			<input type="button" class="button" value="지정" onClick="evtworkerlist('<%=mdlist.FbrdList(i).Fevt_code%>')">
<%
		End IF
	End Sub

	Public Function fnJandiCall(iseq)
		Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
		istrParam = "?seq="&iseq&"&jandi=74f54dc45e253a9edf502bc2a1f61590"

		On Error Resume Next
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.open "GET", "http://110.93.128.100:8090/scmapi/boards/notiboard" & istrParam, false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send()


		Set objXML= nothing
	End Function
%>