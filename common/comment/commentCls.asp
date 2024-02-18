<%
Class CSCMComment
	public FCPage	'Set 현재 페이지
	public FPSize	'Set 페이지 사이즈
	public FTotCnt
	public FGubun
	public FUserID
	
	public FBoardGubun
	public FParentIdx
	public FDeleteyn
	public FComment
	public FCIdx
	
	
	'####### 코멘트리스트 #######
	public Function fnGetSCMCommentList
		Dim strSql,iDelCnt, strSubSql

		strSubSql = " AND A.boardGubun = '" & FBoardGubun & "' AND A.parentIdx = '" & FParentIdx & "' "
		
		If FDeleteyn <> "" Then
			strSubSql = strSubSql & " AND A.deleteyn = '" & FDeleteyn & "' "
		End If
		
		strSql = " SELECT COUNT(A.cIdx) From " & _
				" 		[db_board].[dbo].[tbl_scm_comment] AS A " & _
				"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B ON A.regUserid = B.userid " & _
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
			
			strSql = "	SELECT TOP "&FPSize&" " & _
					"			 A.cIdx, A.comment, isNull(A.etc1,'') AS etc1, isNull(A.etc2,'') AS etc2, A.regUserid, A.deleteyn, A.regdate, B.username, A.boardGubun, A.parentIdx "&_
					"	FROM [db_board].[dbo].[tbl_scm_comment] AS A "&_
					"	INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B ON A.regUserid = B.userid "&_
					"	WHERE " & _
					"		1=1 " & strSubSql & " AND " & _
					"		A.cIdx NOT IN "&_
					"	( "&_
					"			SELECT TOP "&iDelCnt&" X.cIdx FROM [db_board].[dbo].[tbl_scm_comment] AS X "&_
					"			INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS Y ON X.regUserid = Y.userid "&_
					"			WHERE 1=1 " & Replace(strSubSql,"A.","X.") & " " & _
					"			ORDER BY X.cIdx DESC "&_
					"	) "&_
					"	ORDER BY A.cIdx DESC "
			rsget.Open strSql,dbget
			'response.write strSql
			IF not rsget.EOF THEN
				fnGetSCMCommentList = rsget.getRows() 
			END IF	
			rsget.close
		END IF	
		
	End Function
	
	
	'####### 코맨트보기 #######
	public Function fnGetSCMCommentView
		Dim strSql
		strSql = "	SELECT comment " & _
				"		FROM [db_board].[dbo].[tbl_scm_comment] " & _
				"	WHERE cIdx = '" & FCIdx & "' AND regUserid = '" & session("ssBctId") & "' "
		rsget.Open strSql,dbget,1
		'response.write strSql
		IF not rsget.EOF THEN
			 FComment	= db2html(rsget("comment"))
		END IF
		rsget.close

	End Function
End Class



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
	
'####### 게시판 이름 가져오기 #######
public Function fnBoardName(ttype, boardgubun)
	Dim strSql
	strSql = "SELECT boardName FROM [db_board].[dbo].[tbl_scm_commonBoard_list] WHERE type = '" & ttype & "' AND boardGubun <> '" & boardgubun & "' "
	rsget.Open strSql,dbget,1
	IF not rsget.EOF THEN
		fnBoardName = rsget("boardName")
	Else
		fnBoardName = ""
	END IF
	rsget.close
End Function
%>