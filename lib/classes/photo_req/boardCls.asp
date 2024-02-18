<%
Class board_item
	Public Fbrd_subject
	Public Fbrd_hit
	Public Fbrd_regdate
	Public Fbrd_fixed
	Public Fbrd_sn
	Public Fbrd_username
	Public Fbrd_team
	Public Fcnt
	
	Public Fevt_code
	Public FpartMDid
	Public Fevt_name
	Public Fevt_winner
	Public Fevt_regdate
	Public Fevt_laterdate	
End Class

Class board
	Public FbrdList()
	Public FcmtList()
	Public Fbrd_sn
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
	Public Fsearch_team
	Public Fdetail_search
	Public Fsearchstr

	Public Fcmt_sn
	Public FOnecmt
	Public Fid
	Public Fcmt_content
	Public Fcmt_regdate
	Public Fjob

	Public Sub fnBoardcontent
		Dim strSql
		strSql = "select max(bbs_no) as bbs_no from db_partner.dbo.tbl_photo_bbs"
		rsget.open strSql, dbget, 1

		If not rsget.EOF Then
			Fbrd_sn = rsget("bbs_no")
		End If

		If IsNull(rsget("bbs_no")) Then
			Fbrd_sn = "0"
		End If

		rsget.close
	End Sub


	Public Function fnBoardlist
		Dim strSql, i, where
		' 열람선택 검색 부 where절'
		'If Fsearch_team = "all" Then
		'	where = " and p.part_sn = 1 "
		'ElseIf Fsearch_team = "team" Then
		'	where = " and p.part_sn <> 1 "
		'ElseIf Fsearch_team = "" Then
		'	where = ""
		'End If

		' 상세 검색 부 where절'
		If Fdetail_search = "title" Then
			where = where & " and B.brd_title like '%"&Fsearchstr&"%' "
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
		
		'연람선택 관련 where절
		'If Fdetail_search = "" Then
		'	If FJob_sn <> "0" Then
		'		Fjob = " and (p.job_sn >= '" & FJob_sn & "' or P.job_sn = '0' ) and ('" & FAdminlsn & "' = '1' or '"&FPositsn&"' = '2') or (p.part_sn = '1' or p.part_sn = '" & FPartpsn & "') "
		'	Else
		'		Fjob = " and (p.job_sn <= '" & FJob_sn & "') and ('" & FAdminlsn & "' = '1' or '"&FPositsn&"' = '2') or (p.part_sn = '1' or p.part_sn = '" & FPartpsn & "') "
		'	End If
		'End If
		
		' 총 갯수 구하기 '
		strSql = "select count(distinct B.bbs_no) as cnt " + vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_photo_bbs as B " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on B.bbs_id = T.userid " & vbcrlf
		strSql = strSql & " where b.brd_isusing = 'N' " & vbcrlf
		'response.write strSql
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTotalCount = rsget("cnt")
		Else
			FTotalCount = 0
		End If
		rsget.Close
		' 총 갯수 구하기 끝'
		
		'게시글 리스트 구하기'
		strSql = ""
		strSql = strSql & " select top "& Cstr(FPageSize * FCurrPage) &" B.bbs_no, T.username, B.bbs_title, B.bbs_hit, B.bbs_regdate,  B.brd_fixed" & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_photo_bbs as B " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on B.bbs_id = T.userid " & vbcrlf
		strSql = strSql & " where b.brd_isusing = 'N'" & vbcrlf
		strSql = strSql & " group by T.username, B.bbs_no, B.bbs_title, B.bbs_hit, B.bbs_regdate,  B.brd_fixed" & vbcrlf
		strSql = strSql & " order by B.brd_fixed asc, B.bbs_no desc"
		'response.write strSql
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
				FbrdList(i).Fbrd_sn 		= rsget("bbs_no")
				FbrdList(i).Fbrd_username	= rsget("username")
				FbrdList(i).Fbrd_subject	= rsget("bbs_title")
				FbrdList(i).Fbrd_fixed 		= rsget("brd_fixed")
				FbrdList(i).Fbrd_hit 		= rsget("bbs_hit")
				FbrdList(i).Fbrd_regdate 	= rsget("bbs_regdate")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
		'게시글 리스트 구하기 끝'
	End Function

	Public Function fnBoardview
		Dim strSql, i
		strSql = ""
		strSql = strSql & " select Top 1 B.bbs_no, B.bbs_id, B.bbs_title, B.bbs_content, B.bbs_hit, B.bbs_regdate, T.username, B.brd_fixed" & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_photo_bbs as B " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on T.userid = B.bbs_id " & vbcrlf
		strSql = strSql & " where B.bbs_no = '"& Fbrd_sn &"'" & vbcrlf
		'response.write strSql
		rsget.Open strSql,dbget,1

		If not rsget.EOF Then
			Fbrd_sn 		= rsget("bbs_no")
			Fbid 			= rsget("bbs_id")
			Fbrd_subject	= rsget("bbs_title")
			Fbrd_content 	= rsget("bbs_content")
			Fbrd_hit 		= rsget("bbs_hit")
			Fbrd_regdate 	= rsget("bbs_regdate")
			Fbrd_fixed 		= rsget("brd_fixed")
			Fusername 		= rsget("username")
		End If
		rsget.Close
	End Function

	Public Function fnBoardmodify
		Dim strSql, i
		strSql = ""
		strSql = strSql & " select b.bbs_no, B.bbs_id, T.username, B.bbs_title, B.bbs_content, B.bbs_hit, B.bbs_regdate, B.brd_fixed" & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_photo_bbs as B " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on B.bbs_id = T.userid " & vbcrlf
		strSql = strSql & " where B.bbs_no = '"& Fbrd_sn &"'" & vbcrlf
		'response.write strSql
		rsget.Open strSql,dbget,1

		If not rsget.EOF Then
			Fbid 			= rsget("bbs_id")
			Fbrd_username	= rsget("username")
			Fbrd_subject	= rsget("bbs_title")
			Fbrd_content 	= rsget("bbs_content")
			Fbrd_fixed 		= rsget("brd_fixed")
			Fbrd_hit 		= rsget("bbs_hit")
			Fbrd_regdate 	= rsget("bbs_regdate")
		End If
		rsget.Close
	End Function

	Public Sub fnBoardreplylist
		Dim strSql

		strSql = ""
		strSql = strSql & " select C.cmt_sn, C.id, C.cmt_content, C.cmt_regdate, C.brd_sn, T.username " & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_cooperate_board_comment as C " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on C.id = T.userid   " & vbcrlf
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
		strSql = "	SELECT file_no, file_name, real_name" & _
				"		FROM [db_partner].[dbo].tbl_photo_file " & _
				"	WHERE bbs_no = '" & Fbrd_sn & "' " & _
				"	ORDER BY bbs_no ASC "
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
		strSql = strSql & " select top 5 B.brd_sn, T.username, B.brd_subject, B.brd_hit, B.brd_regdate, B.brd_fixed, B.brd_team, (select count(*) from db_partner.dbo.tbl_cooperate_board_comment where brd_sn = B.brd_sn) as cnt  " & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_cooperate_board as B " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_cooperate_board_part as P on B.brd_sn = P.brd_sn and B.brd_isusing = 'N' " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on B.id = T.userid " & vbcrlf
		strSql = strSql & " Left Join db_partner.dbo.tbl_cooperate_board_comment as C on B.brd_sn = C.brd_sn " & vbcrlf		
		strSql = strSql & " where  ('" & FAdminlsn & "' = '1' or '"&FPositsn&"' = '2') " & vbcrlf	'시스템팀일 때는 전부 보임
		strSql = strSql & " or (p.part_sn = '1' or p.part_sn = '" & FPartpsn & "') and  "&Fjob&"  " & vbcrlf
		strSql = strSql & " and p.posit_sn >= '"&FPositsn&"' and B.brd_isusing = 'N' " & vbcrlf
		strSql = strSql & " group by T.username, B.brd_sn, B.brd_subject, B.brd_hit, B.brd_regdate, B.brd_fixed, B.brd_team  " & vbcrlf
		strSql = strSql & " order by B.brd_regdate desc"
		rsget.Open strSql,dbget,1

		'response.write strSql
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
		Dim strSql, i
		
		strSql = ""
		strSql = strSql & " Select E.evt_code,D.partMDid, E.evt_name, P.evt_winner, E.evt_regdate, datediff(d,'"& date() &"', E.evt_prizedate) as laterdate " & vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_event as E " & vbcrlf
		strSql = strSql & " Inner Join db_event.dbo.tbl_event_display as D On E.evt_code = D.evt_code " & vbcrlf
		strSql = strSql & " Left Join db_event.dbo.tbl_event_prize as P On E.evt_code = P.evt_code " & vbcrlf
		strSql = strSql & " where D.partMDid = '" & session("ssBctId") & "' "  & vbcrlf
		'strSql = strSql & " where D.partMDid = 'monkeytn' "  & vbcrlf
		strSql = strSql & " and E.evt_using = 'Y' "  & vbcrlf
		strSql = strSql & " and E.evt_prizedate != '' and P.evt_winner IS NULL "  & vbcrlf
		strSql = strSql & " and E.evt_prizedate <= '"& date() &"' "  & vbcrlf
		strSql = strSql & " order by E.evt_code asc"
		rsget.Open strSql,dbget,1
		'response.write strSql
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

'####### 작업자리스트(촬영요청페이지). #######
Class CCoopUserList
	Public FUserList()
	Public FUser_no
	public FUserType
	public FUser_id
	public FUser_name
	public FUser_useyn
	Public FCodeType
	Public FMode
	Public FResultCount
	Public FTotalCount
	Public FPageCount
	Public FCurrPage
	
	'작업자 리스트
	public Sub fnGetCoopUserList
		Dim strSql, i
		strSql = ""
		strSql = strSql &  "SELECT user_no, user_type, user_id, user_name, user_useyn " & vbcrlf
		strSql = strSql &  "From [db_partner].[dbo].[tbl_photo_user] " & vbcrlf
	
		If FMode = "I" Then
			strSql = strSql &  "WHERE user_type = '"&FCodeType&"' ORDER BY user_no ASC "
		ElseIf FMode = "U" Then
			strSql = strSql &  "WHERE user_no = '"&FUser_no&"' "
		ElseIf FMode = "BB" Then
			strSql = strSql &  "WHERE user_id = '"&session("ssBctID")&"' "
		End If
	
		rsget.Open strSql,dbget,1
	
		FResultCount = rsget.recordcount
		FTotalCount = FResultCount
	
		Redim preserve FUserList(FResultCount)
		FPageCount = FCurrPage - 1
	
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				set FUserList(i) = new CCoopUserList
				FUserList(i).FUser_no		= rsget("user_no")
				FUserList(i).FUserType 		= rsget("user_type")
				FUserList(i).FUser_id 		= rsget("user_id")
				FUserList(i).FUser_name 	= rsget("user_name")
				FUserList(i).FUser_useyn 	= rsget("user_useyn")
				rsget.Movenext
				i = i + 1
			Loop
		End If
		rsget.Close
	
	End Sub

End Class
%>

<%
	Function DrawPartinfoCombo(checkBoxName,checkboxId, bsn)
	Dim tmp_str, strSql, i, Sn, k
		If checkboxId = "" Then
			strSql = "select part_sn, part_name from db_partner.dbo.tbl_partinfo where part_isDel = 'N' and part_sn not in (1,2,3,6)"
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
	<select name="<%=selectBoxName%>">
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
%>
