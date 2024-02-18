<%
'#############################################################
'	PageName 	: /admin/hitchhiker/downHitchhiker.asp
'	Description : 히치하이커
'	History		: 2006.11.30 정윤정 생성
'				  2016.07.07 한용민 수정
'#############################################################

Class Chitchhiker_item
	public FiHvol
	public FiAvol
	public FiAvol2
	public Fuserid
	public Fregdate
	public FAdminId
	public FAdminNm
End Class

Class Chitchhiker
	public FHVol	'Set 히치하이커 발행회차
	public FAVol	'Set 발송신청회차
	public FisSend	'Set 발송여부
	public FCPage	'Set 현재 페이지
	public FPSize	'Set 페이지 사이즈
	public FTotCnt	'Set 전체 레코드 갯수
	public FSearch
	public FSearchTxt 'Set 검색어
	public FSDate	'Set 검색시작일
	public FEDate	'Set 검색종료일

	public FPageSize
	public FCurrPage
	public FTotalCount
	public FResultCount
	public FTotalPage
	public FPageCount
	Public FScrollCount
	public FHitchLogList()

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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

	public Function fnGetHVol()
		Dim strSql 

		strSql = "SELECT HVol FROM [db_user].[dbo].tbl_user_hitchhiker GROUP BY HVol ORDER BY HVol DESC "
		rsget.Open strSql,dbget,1	
		IF NOT rsget.EOF THEN
			fnGetHVol = rsget.getRows()
		END IF		
		rsget.Close		
	End Function

	public Function fnGetApplyVol()
		Dim strSql

		strSql =" SELECT ApplyVol "&_
				" FROM [db_user].[dbo].[tbl_user_hitchhiker] "&_
				" WHERE HVol = "&FHVol&_
				" Group by ApplyVol ORDER BY ApplyVol DESC "
		rsget.Open strSql,dbget,1	
		IF NOT rsget.EOF THEN
			fnGetApplyVol = rsget.getRows()
		END IF		
		rsget.Close
	End Function
	
	public Function fnGetSendApplyVol()
		Dim strSql

		strSql =" SELECT ApplyVol  "&_
				" FROM [db_user].[dbo].[tbl_user_hitchhiker] "&_
				" WHERE HVol = "&FHVol&" and ApplyVol <=  isnull(SendVol,0) "&_
				" Group by ApplyVol ORDER BY ApplyVol DESC "				
		rsget.Open strSql,dbget,1	
		IF NOT rsget.EOF THEN
			fnGetSendApplyVol = rsget.getRows()
		END IF		
		rsget.Close
	End Function
	
	public Function fnGetList()
		Dim strSql, strSqlCnt, strSqlAdd
		
		If FSearch <> "" THEN
			If FSearch = "userid" Then
				IF FSearchTxt <> "" THEN
					strSqlAdd = strSqlAdd & " and a.userid like '%"&FSearchTxt&"%' " 
				END IF			
			ElseIf FSearch = "username" Then
				IF FSearchTxt <> "" THEN
					strSqlAdd = strSqlAdd & " and b.username like '%"&FSearchTxt&"%' " 
				END IF			
			ElseIf FSearch = "receviename" Then
				IF FSearchTxt <> "" THEN
					strSqlAdd = strSqlAdd & " and a.recevieName like '%"&FSearchTxt&"%' " 
				END IF			
			End If
		End If
		
		IF FAVol <> "" THEN
			strSqlAdd = strSqlAdd & " and a.ApplyVol = "&FAVol
		END IF	
		
		IF FisSend = "1" THEN
			strSqlAdd = strSqlAdd & " and (isnull(a.ApplyVol,0) > isnull(a.SendVol,0)) "
		ELSEIF 	FisSend = "2" THEN
			strSqlAdd = strSqlAdd & " and (isnull(a.ApplyVol,0) <= isnull(a.SendVol,0)) "
		END IF	

		IF FSDate<>"" then
			strSqlAdd = strSqlAdd & " and a.ApplyDate >= '" & FSDate & " 00:00:00' "
		end if
		IF FEDate<>"" then
			strSqlAdd = strSqlAdd & " and a.ApplyDate <= '" & FEDate & " 23:59:59' "
		end if

		strSqlCnt = " SELECT COUNT(a.userid) FROM [db_user].[dbo].tbl_user_hitchhiker as a INNER JOIN [db_user].[dbo].tbl_user_n as b ON a. userid = b.userid " &_
					"	WHERE a.HVol = "&FHVol & strSqlAdd 	
	
		rsget.Open strSqlCnt,dbget,1	
		IF NOT rsget.EOF THEN
			FTotCnt = rsget(0)
		END IF		
		rsget.Close					

		IF FTotCnt > 0 THEN
			strSql = "SELECT distinct top "& FPSize*FCPage
			strSql = strSql & "	a.HVol,a.ApplyVol, a.ApplyDate, a.userid, b.username, a.zipcode, a.zipaddr" & vbcrlf
			strSql = strSql & "	,a.useraddr, a.userphone, a.usercell, a.SendDate, " & vbcrlf
			strSql = strSql & " (select userlevel from [db_user].[dbo].tbl_logindata where userid = a.userid) as userlevel, a.recevieName, a.idx" & vbcrlf
			strSql = strSql & " ,a.reqmsg" & vbcrlf
			strSql = strSql & "	FROM [db_user].[dbo].tbl_user_hitchhiker as a" & vbcrlf
			strSql = strSql & "	join [db_user].[dbo].tbl_user_n as b" & vbcrlf
			strSql = strSql & "		on a.userid = b.userid" & vbcrlf
			strSql = strSql & "	WHERE a.HVol = "&FHVol & strSqlAdd & "" & vbcrlf
			strSql = strSql & " ORDER BY ApplyDate DESC" & vbcrlf

			'response.write strSql & "<Br>"
			rsget.pagesize = FPSize					
			rsget.Open strSql,dbget,1

			IF NOT rsget.EOF   THEN
				rsget.absolutepage = FCPage	
				fnGetList = rsget.getRows()
			END IF		
			rsget.Close		
		END IF	
	End Function

	Public Function fnHitchLoglist
		Dim strSql, i, where

		strSql = "select count(*) as cnt " + vbcrlf
		strSql = strSql & " from [db_log].[dbo].[tbl_user_hitchhikerLog] " & vbcrlf
		strSql = strSql & " where iHVol = "&FHVol&" " & vbcrlf
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTotalCount = rsget("cnt")
		Else
			FTotalCount = 0
		End If
		rsget.Close

		strSql = " select top "& Cstr(FPageSize * FCurrPage) &" H.iHvol, H.iAvol, H.iAvol2, H.userid, H.regdate, H.AdminId, T.username " & vbcrlf
		strSql = strSql & " from [db_log].[dbo].[tbl_user_hitchhikerLog] as H" & vbcrlf
		strSql = strSql & "		left join [db_partner].dbo.tbl_user_tenbyten as T " & vbcrlf
		strSql = strSql & "		on H.AdminId=T.userid " & vbcrlf
		strSql = strSql & " where H.iHVol = "&FHVol&" " & vbcrlf
		strSql = strSql & " order by H.idx desc " & vbcrlf

		'response.write strSql & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1
		Redim preserve FHitchLogList(FResultCount)
		FPageCount = FCurrPage - 1

		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FHitchLogList(i) = new Chitchhiker_item
				FHitchLogList(i).FiHvol 	= rsget("iHvol")
				FHitchLogList(i).FiAvol		= rsget("iAvol")
				FHitchLogList(i).FiAvol2	= rsget("iAvol2")
				FHitchLogList(i).Fuserid	= rsget("userid")
				FHitchLogList(i).Fregdate	= rsget("regdate")
				FHitchLogList(i).FAdminId 	= rsget("AdminId")
				FHitchLogList(i).FAdminNm 	= rsget("username")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Function
End Class

Class CUserInfo
	public FUID 'set  아이디
	public FZipCode
	public FUsername
	public FAddress1
	public FAddress2
	public FPhone
	public FCell	
	public FChk

	Public FHVol
	Public FRecevieName
	Public FUserphone
	Public FUsercell
	Public FiHVol
	Public FiApplyVol
	Public FRegCount
	public frectidx

	public Function fnGetUserInfo()
		Dim strSql, userid, searchsql

		if frectidx <> "" Then
			searchsql = searchsql & " and idx= '"&frectidx&"'"
		end if
		if FUID <> "" Then
			searchsql = searchsql & " and userid= '"&FUID&"'"
		end if

		strSql = ""
		strSql = "select count(*) as cnt, Applyvol from [db_user].[dbo].tbl_user_hitchhiker"
		strSql = strSql & " where Hvol = '"&FiHVol&"' " & searchsql
		strSql = strSql & " group by Applyvol "

		'response.write strSql & "<Br>"
		'response.end
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FRegCount = rsget("cnt")
			FiApplyVol = rsget("Applyvol")
		Else
			FRegCount = 0
			FiApplyVol = 0
		End If
		rsget.Close

'		strSql = "select userid from [db_user].[dbo].tbl_user_hitchhiker where Hvol = '"&FiHVol&"' " & searchsql
'
'		rsget.Open strSql,dbget,1
'		If not rsget.EOF Then
'			userid = rsget("userid")
'		End If
'		rsget.Close

		strSql = " SELECT"
		strSql = strSql & " isnull(u.zipcode,'') as zipcode, u.zipaddr as address1, u.useraddr, isnull(u.userphone,'') as userphone" & vbcrlf
		strSql = strSql & " , isnull(u.usercell,'') as usercell, u.username, u.userid" & vbcrlf
		strSql = strSql & " from [db_user].[dbo].tbl_user_n as u" & vbcrlf
		strSql = strSql & " WHERE u.userid = '" & FUID & "'"

		'response.write strSql & "<Br>"
		rsget.Open strSql,dbget,1
		IF NOT rsget.EOF THEN
			FChk = 1
			FUsername = rsget("username")
			FZipCode = rsget("zipcode")
			FAddress1 = rsget("address1")
			FAddress2 = rsget("useraddr")
			FPhone = rsget("userphone")
			FCell = rsget("usercell")
			FUID = rsget("userid")
		ELSE
			FChk = 0	
	    END IF	
	End Function

	Public Function updateHitchAddr()
		Dim strSql

		strSql = "SELECT top 1 a.HVol,a.ApplyVol, a.ApplyDate, a.userid, a.zipcode, " & vbcrlf
		strSql = strSql & "a.zipaddr, a.useraddr, a.userphone, a.usercell, a.SendDate, a.recevieName, a.idx" & vbcrlf
		strSql = strSql & "FROM [db_user].[dbo].tbl_user_hitchhiker as a" & vbcrlf
		strSql = strSql & "where a.HVol = '"&FHVol&"' and a.idx ='"&frectidx&"'" & vbcrlf

		'response.write strSql & "<br>"
		rsget.Open strSql,dbget,1
		IF NOT rsget.EOF THEN
			updateHitchAddr = rsget.getRows()
	    END IF	
	End Function
End Class
%>
