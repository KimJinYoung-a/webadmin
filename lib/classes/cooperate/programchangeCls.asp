<%
'###########################################################
' Description : 프로그램변경내역
' Hieditor : 강준구 생성
'			 2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################

	Class CProgramChange
		public FCPage	'Set 현재 페이지
		public FPSize	'Set 페이지 사이즈
		public FTotCnt
		public FGubun
		public FUserID
		public FPIdx
		public FTitle
		public FContent
		public FReguserid
		public FUsername
		public FRegdate
		public FSign1
		public FSign2
		public FSign1date
		public FSign2date
		public FFileName
		public FDocIdx
		public FRectRegUserID
		public FRectTitle
		public FRect1Check
		public FRect2Check
		public FChkList
		public FSign1Chk
		public FSign2Chk
		

		
		public Function fnGetPrChList
			Dim strSql,iDelCnt, strSubSql

			If FRectRegUserID <> "" Then
				strSubSql = strSubSql & " AND A.reguserid = '" & FRectRegUserID & "' "
			End IF
			
			If FRectTitle <> "" Then
				strSubSql = strSubSql & " AND A.title like '%" & FRectTitle & "%' "
			End IF
			
			If FRect1Check <> "" Then
				If FRect1Check = "o" Then
					strSubSql = strSubSql & " AND A.sign1 <> '' "
				Else
					strSubSql = strSubSql & " AND A.sign1 = '' "
				End If
			End IF
			
			If FRect2Check <> "" Then
				If FRect2Check = "o" Then
					strSubSql = strSubSql & " AND A.sign2 <> '' "
				Else
					strSubSql = strSubSql & " AND A.sign2 = '' "
				End If
			End IF
			
			strSql = " SELECT COUNT(A.pidx) From " & _
					" 		[db_board].[dbo].[tbl_program_change] AS A " & _
					"		INNER JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS C ON A.reguserid = C.userid "&_
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
						"		A.pidx, A.title, A.contents, C.username, A.regdate, A.sign1, A.sign2, A.sign1date, A.sign2date, A.filename "&_
						"	FROM [db_board].[dbo].[tbl_program_change] AS A "&_
						"	INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS C ON A.reguserid = C.userid "&_
						"	WHERE " & _
						"		1=1 " & strSubSql & " AND " & _
						"		A.pidx NOT IN "&_
						"	( "&_
						"			SELECT TOP "&iDelCnt&" X.pidx FROM [db_board].[dbo].[tbl_program_change] AS X "&_
						"			INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS W ON X.reguserid = W.userid "&_
						"			WHERE 1=1 " & Replace(Replace(strSubSql,"A.","X."),"C.","W.") & " " & _
						"			ORDER BY X.pidx DESC "&_
						"	) "&_
						"	ORDER BY A.regdate DESC "
				rsget.Open strSql,dbget
				'response.write strSql
				IF not rsget.EOF THEN
					fnGetPrChList = rsget.getRows() 
				END IF	
				rsget.close
			END IF	
			
		End Function
		
		
		public Function fnGetPrChView
			Dim strSql
			strSql = "	SELECT A.pidx, A.title, A.contents, A.reguserid, B.username, A.regdate, A.sign1, A.sign2, A.sign1date, A.sign2date, A.filename, A.doc_idx, A.chklist, A.sign1chk, A.sign2chk " & _
					"		FROM [db_board].[dbo].[tbl_program_change] AS A " & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B ON A.reguserid = B.userid " & _
					"	WHERE A.pidx = '" & FPIdx & "' "
			rsget.Open strSql,dbget,1
			'response.write strSql
			IF not rsget.EOF THEN
				FTitle			= db2html(rsget("title"))
				FContent		= db2html(rsget("contents"))
				FReguserid		= rsget("reguserid")
				FUsername		= rsget("username")
				FRegdate		= rsget("regdate")
				FSign1			= rsget("sign1")
				FSign2			= rsget("sign2")
				FSign1date		= rsget("sign1date")
				FSign2date		= rsget("sign2date")
				FRegdate		= rsget("regdate")
				FFileName		= rsget("filename")
				FDocIdx			= rsget("doc_idx")
				FChkList		= rsget("chklist")
				FSign1Chk		= rsget("sign1chk")
				FSign2Chk		= rsget("sign2chk")

			END IF
			rsget.close

		End Function
		
		
		public Function fnGetPrChAnsList
			Dim strSql,iDelCnt, strSubSql

			strSubSql = " AND A.pidx = '" & FPIdx & "' "
			
			strSql = " SELECT COUNT(A.idx) From " & _
					" 		[db_board].[dbo].tbl_program_change_comment AS A with (nolock)" & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B with (nolock) ON A.userid = B.userid " & _
					"	WHERE " & _
					"		1=1 AND A.useyn = 'Y' " & strSubSql & " "

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
						"			 A.idx, A.userid, A.comment, A.regdate, B.username "&_
						"	FROM [db_board].[dbo].tbl_program_change_comment AS A "&_
						"	INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS B ON A.userid = B.userid "&_
						"	WHERE " & _
						"		1=1 AND A.useyn = 'Y' " & strSubSql & " AND " & _
						"		A.idx NOT IN "&_
						"	( "&_
						"			SELECT TOP "&iDelCnt&" X.idx FROM [db_board].[dbo].tbl_program_change_comment AS X "&_
						"			INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS Y ON X.userid = Y.userid "&_
						"			WHERE 1=1 AND X.useyn = 'Y' " & Replace(strSubSql,"B.","Y.") & " " & _
						"			ORDER BY X.idx DESC "&_
						"	) "&_
						"	ORDER BY A.idx DESC "

				'response.write strSql & "<Br>"
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

				IF not rsget.EOF THEN
					fnGetPrChAnsList = rsget.getRows() 
				END IF	
				rsget.close
			END IF	
			
		End Function
		
		
		public Function fnGetMemList
			Dim strSql,iDelCnt, strSubSql

			strSql = "	SELECT userid, username FROM [db_partner].[dbo].tbl_user_tenbyten WHERE userid <> '' AND part_sn in(7,30) AND isusing = 1" & vbcrlf

			' 퇴사예정자 처리	' 2018.10.16 한용민
			strSql = strSql & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
			strSql = strSql & "	ORDER BY posit_sn ASC "
	
			'response.write strSql &"<Br>"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetMemList = rsget.getRows() 
			END IF	
			rsget.close

		End Function
		
End Class


Function fnCheckBoxCheck(arr,v)
	Dim i, tmp
	tmp = ""
	If arr <> "" Then
		For i = 0 To UBound(Split(arr,","))
			If CStr(Split(arr,",")(i)) = CStr(v) Then
				tmp = "checked"
				Exit For
			End If
		Next
	End If
	fnCheckBoxCheck = tmp
End Function
%>
