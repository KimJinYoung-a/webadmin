<%
Class CWishList
	public FCPage	'Set 현재 페이지
	public FPSize	'Set 페이지 사이즈
	public FTotCnt
	
	public FUserID
	public FFidx
	public FItemID
	public FCate_Large
	public FOrgPrice
	
	public FPriceS
	public FPriceE
	public FInCnt
	public FGrade
	public FAllCate
	
	public FIdx
	'// 리스트
	public Function fnGetWishList
		Dim strSql,iDelCnt, strSubSql
		
		strSubSql = ""
		
		If FPriceS <> "" AND FPriceE <> "" Then
			strSubSql = strSubSql & " AND isNull(SUM(A.orgprice),0) Between '" & FPriceS & "' AND '" & FPriceE & "' "
		ElseIf FPriceS <> "" AND FPriceE = "" Then
			strSubSql = strSubSql & " AND isNull(SUM(A.orgprice),0) >= '" & FPriceS & "' "
		ElseIf FPriceS = "" AND FPriceE <> "" Then
			strSubSql = strSubSql & " AND isNull(SUM(A.orgprice),0) <= '" & FPriceE & "' "
		End IF
		
		If FInCnt <> "" Then
			strSubSql = strSubSql & " AND count(A.itemid) >= '" & FInCnt & "' "
		End If
		
		If FGrade <> "" Then
			strSubSql = strSubSql & " AND IsNULL(B.userlevel,5) = '" & FGrade & "' "
		End If
		
		If FAllCate <> "" Then
			strSubSql = strSubSql & " AND (select count(*) from (select distinct cate_large from [db_temp].[dbo].[tbl_wishlist_event] where userid = A.userid) AS aa) = 12 "
		End If
		
		strSql = " SELECT COUNT(*) From ( " & _
				" 	SELECT A.userid FROM [db_temp].[dbo].[tbl_wishlist_event] AS A " & _
				"	INNER JOIN [db_user].[dbo].[tbl_logindata] AS B ON A.userid = B.userid " & _
				"	GROUP BY A.userid, B.userlevel, A.fidx " & _
				"	HAVING 1=1 " & strSubSql & " " & _
				"	) AS X"
		rsget.Open strSql,dbget 
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close


		IF FTotCnt > 0 THEN
			iDelCnt =  ((FCPage - 1) * FPSize )	
			
			strSql = "	SELECT TOP "&FPSize&" A.userid, count(A.itemid) AS TOTALCNT, "&_
					"		isNull(SUM(A.orgprice),0) AS ORGPRICE, IsNULL(B.userlevel,5) as userlevel, A.fidx "&_
					"	FROM [db_temp].[dbo].tbl_wishlist_event AS A "&_
					"	INNER JOIN [db_user].[dbo].[tbl_logindata] AS B ON A.userid = B.userid "&_
					"	WHERE A.userid NOT IN "&_
					"	( "&_
					"		SELECT TOP "&iDelCnt&" X.userid FROM [db_temp].[dbo].tbl_wishlist_event AS X "&_
					"		INNER JOIN [db_user].[dbo].[tbl_logindata] AS Y ON X.userid = Y.userid "&_
					"		GROUP BY X.userid, Y.userlevel "&_
					"		HAVING 1=1 " & Replace(Replace(strSubSql,"A.","X."),"B.","Y.") & " " & _
					"		ORDER BY X.userid ASC "&_
					"	) "&_
					"	GROUP BY A.userid, B.userlevel, A.fidx "&_
					"	HAVING 1=1 " & strSubSql & " " & _
					"	ORDER BY A.userid ASC "
			rsget.Open strSql,dbget 
			IF not rsget.EOF THEN
				fnGetWishList = rsget.getRows() 
			END IF	
			rsget.close
		END IF	
	End Function
	
	
	public Function fnGetWishListExcel
		Dim strSql
'		If FUserID = "" AND FIdx = "" Then
'			strSql = "	SELECT	" & _
'					"		A.userid, IsNULL(B.userlevel,5) as userlevel, A.fidx, A.itemid, A.orgprice, C.itemname, D.socname, A.cate_large	" & _
'					"	FROM [db_temp].[dbo].[tbl_wishlist_event] AS A	" & _
'					"		INNER JOIN [db_user].[dbo].[tbl_logindata] AS B ON A.userid = B.userid "&_
'					"		INNER JOIN [db_item].[dbo].[tbl_item] AS C ON A.itemid = C.itemid "&_
'					"		INNER JOIN [db_user].[dbo].[tbl_user_c] AS D ON C.makerid = D.userid "&_
'					"	ORDER BY A.userid ASC, A.fidx ASC "
'		Else
			strSql = "	SELECT	" & _
					"		A.userid, IsNULL(B.userlevel,5) as userlevel, A.fidx, A.itemid, A.orgprice, C.itemname, D.socname, A.cate_large, C.smallimage	" & _
					"	FROM [db_temp].[dbo].[tbl_wishlist_event] AS A	" & _
					"		INNER JOIN [db_user].[dbo].[tbl_logindata] AS B ON A.userid = B.userid "&_
					"		INNER JOIN [db_item].[dbo].[tbl_item] AS C ON A.itemid = C.itemid "&_
					"		INNER JOIN [db_user].[dbo].[tbl_user_c] AS D ON C.makerid = D.userid "&_
					"	WHERE A.userid in ('apple906','minyuzoa','orol77') " & _
					"	ORDER BY A.userid ASC, A.fidx ASC "
'		End IF
'					"	WHERE A.userid = '" & FUserID & "' AND A.fidx = '" & FFIdx & "'	" & _
		rsget.Open strSql,dbget 
		IF not rsget.EOF THEN
			fnGetWishListExcel = rsget.getRows() 
		END IF	
		rsget.close
	End Function
	
	
	public FListImg
	public FMainImg
	public FUsing
	public FRegdate
	public FOpendate
	public FVolnum
	
	'// 내용보기
	public Function fnGetConts
	 	Dim strSql
	 	strSql = " SELECT idx, listimg, mainimg, isUsing, regdate, opendate, volnum "&_
	 			" FROM [db_event].[dbo].[tbl_event_wonderday] "&_
				" WHERE idx ="&FIdx				
		rsget.Open strSql,dbget 
			IF not rsget.EOF THEN
				FListImg	= rsget("listimg")	
				FMainImg	= rsget("mainimg")	
				FUsing		= rsget("isUsing")	
				FRegdate	= rsget("regdate")	
				FOpendate	= rsget("opendate")
				FVolnum		= rsget("volnum")
			END IF	
		rsget.close				
	END Function
	
End Class

Public Function UserGrade(data)
	Dim vGrade
	If data <> "" Then
		Select Case data
			Case 1
				vGrade = "<font color='#59a05b'><b>Green</b></font>"
			Case 2
				vGrade = "<font color='#5b6fc6'><b>Blue</b></font>"
			Case 3
				vGrade = "<font color='#af4eb0'><b>VIP</b></font>"
			Case 4
				vGrade = "<font color='#fa378e'><b>Mania</b></font>"
			Case 5
				vGrade = "<font color='#e36c14'><b>Orange</b></font>"
			Case 7
				vGrade = "<font color='#c40c0c'><b>Staff</b></font>"
			Case 8
				vGrade = "<font color='#fa378e'><b>Friends</b></font>"
			Case 0
				vGrade = "<font color='#e9b322'><b>Yellow</b></font>"
		End Select
	End If
	UserGrade = vGrade
END Function

Public Function CategoryName(data)
	Dim vCate
	If data <> "" Then
		Select Case data
			Case "010"
				vCate = "디자인문구"
			Case "020"
				vCate = "오피스/개인"
			Case "030"
				vCate = "키덜트/취미"
			Case "040"
				vCate = "가구/수납"
			Case "050"
				vCate = "조명/데코"
			Case "055"
				vCate = "페브릭"
			Case "060"
				vCate = "주방/욕실"
			Case "070"
				vCate = "가방/슈즈/쥬얼리"
			Case "080"
				vCate = "Women"
			Case "090"
				vCate = "Men"
			Case "100"
				vCate = "Baby"
			Case "110"
				vCate = "감성채널"
			Case Else
				vCate = " "
		End Select
	End If
	CategoryName = vCate
END Function
%>
