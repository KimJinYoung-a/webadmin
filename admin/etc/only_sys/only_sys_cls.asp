<%
'####################################################
' Description : 히스토리
' History : 강준구 생성
'			2023.10.04 한용민 수정(쿼리튜닝)
'####################################################

Class cOnlySys_oneitem
	public Fitemid
	public Fitemname
	public Fregdate
	public Fsmallimage
	public Fidx
	public Fuserid
	public Fusername
	public Fgubun
	public Fpk_idx
	public Fsub_idx
	public Fmenupos
	public Fmenuname
	public Fmenulink
	public Fcontents
	public Frefip

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class cOnlySys
	public FItemList()
	public FOneItem
	public FGubun
	public FPK_Idx
	public FPK_Idx_txt
	public FEvtcode
	public FEvtSubject
	public FEvtStatus
	public FEvtRealStatus
	public FEvtSDate
	public FEvtEDate
	public FEvtMDname
	public FEvtIsSale
	public FEvtIsGift
	public FEvtSaleCnt
	public FEvtGiftCnt
	public FUserID
	public FUserName
	public FUserJuminNO
	public FUserEnc_jumin2
	public FUserRealChk
	public FUserBirth
	public FItemID
	public FMakerID
	public FNewMakerID
	public FBrandName
	public FMoveItemCnt
	public FEvtPrizeCode
	public FDB
	public FOrderSerial

	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	
	public Function fnEventCont
		Dim strSql
		IF FEvtcode = "" THEN Exit Function
		strSql = " SELECT E.evt_code, E.evt_name, E.evt_startdate, E.evt_enddate, B.issale, B.isgift, E.evt_state AS real_state "&_
		 		 " ,evt_state = Case When DateDiff(day,getdate(),evt_enddate) < 0 Then 9 "&_
		 		 "				When E.evt_state = 7 and DateDiff(day,getdate(),evt_startdate) <= 0 Then 6 "&_
		 		 "				ELSE	E.evt_state end"&_
				 "		, (Case When isNull(B.partMDid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten with (nolock) WHERE userid = B.partMDid ) Else '' end) as mdname " & _
				 "		, (SELECT COUNT(sale_code) FROM [db_event].[dbo].[tbl_sale] with (nolock) WHERE evt_code = E.evt_code and sale_using =1) as sale_count "&_
				 "		, (SELECT COUNT(gift_code) FROM [db_event].[dbo].[tbl_gift] with (nolock) WHERE evt_code = E.evt_code and gift_using ='y') as gift_count "&_
				 " 		FROM [db_event].[dbo].[tbl_event] AS E with (nolock)"&_
				 "		LEFT OUTER JOIN [db_event].[dbo].[tbl_event_display] as B with (nolock) ON E.evt_code = B.evt_code "&_
				 " WHERE E.evt_code = '" & FEvtcode & "' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FEvtcode		= rsget("evt_code")
			FEvtSubject		= db2html(rsget("evt_name"))
			FEvtRealStatus	= rsget("real_state")
			FEvtStatus		= rsget("evt_state")
			FEvtSDate		= rsget("evt_startdate")
			FEvtEDate		= rsget("evt_enddate")
			FEvtMDname		= rsget("mdname")
			FEvtIsSale		= rsget("issale")
			FEvtIsGift		= rsget("isgift")
			FEvtSaleCnt 	= rsget("sale_count")
			FEvtGiftCnt		= rsget("gift_count")

		End IF
		rsget.Close
	End Function
	
	public Function fnUserInfo
		Dim strSql
		strSql = " SELECT U.userid, U.username, U.juminno, U.Enc_jumin2, U.realnamecheck, U.birthday "&_
				 " 		FROM [db_user].[dbo].[tbl_user_n] AS U with (nolock)"&_
				 " WHERE 1=1"
		If FUserID <> "" Then
			strSql = strSql & " AND U.userid = '" & FUserID & "' "
		End If
		
		If FUserName <> "" Then
			strSql = strSql & " AND U.username = '" & FUserName & "' "
		End If
		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FUserID 		= rsget("userid")
			FUserName 		= rsget("username")
			FUserJuminNO 	= rsget("juminno")
			FUserEnc_jumin2	= rsget("Enc_jumin2")
			FUserRealChk 	= rsget("realnamecheck")
			FUserBirth		= Left(rsget("birthday"),10)
		Else
			FGubun = "x"
		End IF
		rsget.Close
	End Function

	public Function fnUserCheckLog
		Dim strSql
		strSql = "	SELECT * " & _
				"		FROM [db_log].[dbo].[tbl_user_checkLog] with (nolock)" & _
				"	WHERE userid = '" & FUserID & "' " & _
				"	ORDER BY chkIdx DESC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		'response.write strSql
		IF not rsget.EOF THEN
			fnUserCheckLog = rsget.getRows() 
		END IF	
		rsget.close
	End Function

	public Function fnUserCoupon
		Dim strSql
		strSql = "	SELECT idx, userid, couponname, regdate, startdate, expiredate, isusing, deleteyn, orderserial " & _
				"		FROM [db_user].[dbo].[tbl_user_coupon] with (nolock)" & _
				"	WHERE userid = '" & FUserID & "' " & _
				"		and masteridx=126 " & _
				"	ORDER BY idx DESC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		'response.write strSql
		IF not rsget.EOF THEN
			fnUserCoupon = rsget.getRows() 
		END IF	
		rsget.close
	End Function

	public Function fnBrandCont
		Dim strSql
		strSql = " SELECT socname FROM [db_user].[dbo].[tbl_user_c] with (nolock)"&_
				 " WHERE userid = '" & FNewMakerID & "' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FBrandName = db2html(rsget("socname"))

		End IF
		rsget.Close

		If FMakerID <> "" Then
			strSql = "SELECT COUNT(itemid) FROM [db_item].[dbo].[tbl_item] WHERE makerid = '" & FMakerID & "'"
		End If		
		If FItemID <> "" Then
			strSql = "SELECT COUNT(itemid) FROM [db_item].[dbo].[tbl_item] WHERE itemid IN(" & FItemID & ")"
		End If
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		FMoveItemCnt = rsget(0)
		rsget.Close
	End Function

	public Function fnOffBrandCont
		Dim strSql
		strSql = " SELECT socname FROM [db_user].[dbo].[tbl_user_c] with (nolock)"&_
				 " WHERE userid = '" & FNewMakerID & "' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FBrandName = db2html(rsget("socname"))

		End IF
		rsget.Close

		If FMakerID <> "" Then
			strSql = "SELECT COUNT(itemid) FROM db_shop.dbo.tbl_shop_item with (nolock) WHERE makerid = '" & FMakerID & "' and itemgubun = '90' and itemoption = '0000'"
		End If		
		If FItemID <> "" Then
			strSql = "SELECT COUNT(itemid) FROM db_shop.dbo.tbl_shop_item with (nolock) WHERE itemid IN(" & FItemID & ") and itemgubun = '90' and itemoption = '0000'"
		End If
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		FMoveItemCnt = rsget(0)
		rsget.Close
	End Function

	public Function fnTesterList
		Dim strSql, sqlAdd
		If FUserID <> "" Then
			sqlAdd = sqlAdd & " AND evt_winner = '" & FUserID & "' "
		End IF
		If FItemID <> "" Then
			sqlAdd = sqlAdd & " AND itemid = '" & FItemID & "' "
		End IF
		If FEvtCode <> "" Then
			sqlAdd = sqlAdd & " AND evt_code = '" & FEvtCode & "' "
		End IF
		If FEvtPrizeCode <> "" Then
			sqlAdd = sqlAdd & " AND evtprize_code = '" & FEvtPrizeCode & "' "
		End IF
		strSql = "	SELECT * " & _
				"		FROM [db_event].[dbo].[tbl_tester_event_winner] with (nolock)" & _
				"	WHERE 1=1 " & sqlAdd & " " & _
				"	ORDER BY evtprize_code DESC "
		'response.write strSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		IF not rsget.EOF THEN
			fnTesterList = rsget.getRows() 
		END IF	
		rsget.close
	End Function

	public Function fnBrandOrderCont
		Dim strSql

		If FMakerID <> "" Then
			strSql = "SELECT COUNT(itemid) FROM [db_item].[dbo].[tbl_item] with (nolock) WHERE makerid = '" & FMakerID & "'"
		End If		
		If FItemID <> "" Then
			strSql = "SELECT COUNT(itemid) FROM [db_item].[dbo].[tbl_item] with (nolock) WHERE itemid IN(" & FItemID & ")"
		End If
		'response.write strSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		FMoveItemCnt = rsget(0)
		rsget.Close
	End Function

	public Function fnGoodUsingList
		Dim strSql, sqlAdd
		If FUserID <> "" Then
			sqlAdd = sqlAdd & " AND userid = '" & FUserID & "' "
		End IF
		If FItemID <> "" Then
			sqlAdd = sqlAdd & " AND itemid = '" & FItemID & "' "
		End IF

		strSql = "	SELECT Top 20 IDX, UserID, OrderSerial, ItemID, ItemOptionName, ItemOption, Left(Convert(varchar,Contents),30) AS Contents, IsUsing, RegDate " & _
				"		FROM [db_board].[dbo].[tbl_Item_Evaluate] with (nolock)" & _
				"	WHERE 1=1 " & sqlAdd & " " & _
				"	ORDER BY IDX DESC "
		'response.write strSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		IF not rsget.EOF THEN
			fnGoodUsingList = rsget.getRows() 
		END IF	
		rsget.close
	End Function

	public Function fnOrderList
		Dim strSql, sqlAdd
		If FUserID <> "" Then
			sqlAdd = sqlAdd & " AND userid = '" & FUserID & "' "
		End IF
		If FOrderSerial <> "" Then
			sqlAdd = sqlAdd & " AND orderserial = '" & FOrderSerial & "' "
		End If

		strSql = "	SELECT isNull(userDisplayYn,'Y') AS userDisplayYn, orderserial, userid, accountname, buyname, reqname, subtotalprice, regdate " & _
				"		FROM " & FDB & " " & _
				"	WHERE 1=1 " & sqlAdd & " " & _
				"	ORDER BY orderserial DESC "
		'response.write strSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		IF not rsget.EOF THEN
			fnOrderList = rsget.getRows() 
		END IF	
		rsget.close
	End Function

	public Sub fnAwardNotIncludeItemList
		Dim strSql, sqlAdd, i

		strSql = "SELECT COUNT(A.itemid), CEILING(CAST(Count(A.itemid) AS FLOAT)/" & FPageSize & ") AS totPg FROM [db_const].[dbo].[tbl_const_award_NotInclude_Item] AS A with (nolock)"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			ftotalcount = rsget(0)
			FTotalPage = rsget(1)
		END IF
		rsget.close
		
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
		
		IF ftotalcount > 0 THEN
			strSql = "SELECT TOP "&CStr(FPageSize*FCurrPage)&" A.itemid, I.itemname, A.regdate, I.smallimage " & _
					"		FROM [db_const].[dbo].[tbl_const_award_NotInclude_Item] AS A with (nolock)" & _
					"	INNER JOIN [db_item].[dbo].[tbl_item] AS I with (nolock) ON A.itemid = I.itemid " & _
					"	WHERE 1=1 " & _
					"	ORDER BY A.regdate DESC "
			'response.write strSql
			rsget.pagesize = FPageSize
			rsget.CursorLocation = adUseClient
			rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
			Redim preserve FItemList(FResultCount)
			i = 0
			If not rsget.EOF Then
				rsget.absolutepage = FCurrPage
				Do until rsget.EOF
					Set FItemList(i) = new cOnlySys_oneitem
						FItemList(i).Fitemid		= rsget("itemid")
						FItemList(i).Fitemname		= rsget("itemname")
						FItemList(i).Fregdate		= rsget("regdate")
						FItemList(i).Fsmallimage	= rsget("smallimage")
						
					i = i + 1
					rsget.moveNext
				Loop
			End If
			rsget.Close
		End If
	End Sub
	
	' /common/pop_itemhistory.asp
	public Sub fnSCMChangeList
		Dim strSql, sqlAdd, i
		
		If FGubun <> "" Then
			sqlAdd = sqlAdd & " and A.gubun = '" & FGubun & "' "
		End IF
		
		If FEvtSDate <> "" Then
			sqlAdd = sqlAdd & " and A.regdate >= '" & FEvtSDate & " 00:00:00' "
		End IF
		
		If FEvtEDate <> "" Then
			sqlAdd = sqlAdd & " and A.regdate <= '" & FEvtEDate & " 23:59:59' "
		End IF
		
		If FPK_Idx <> "" AND vS_PKIdx_txt <> "" Then
			sqlAdd = sqlAdd & " and A.pk_idx = '" & vS_PKIdx_txt & "' "
		End IF

		strSql = "SELECT COUNT(A.idx), CEILING(CAST(Count(A.idx) AS FLOAT)/" & FPageSize & ") AS totPg FROM [db_log].[dbo].[tbl_scm_change_log] AS A with (nolock)" & _
				 "	LEFT JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS u with (nolock) ON A.userid = u.userid " & _
				 "	LEFT JOIN [db_partner].[dbo].[tbl_partner_menu] AS m with (nolock) ON A.menupos = m.id " & _
				 "	WHERE 1=1 " & sqlAdd

		'response.write strSql & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			ftotalcount = rsget(0)
			FTotalPage = rsget(1)
		END IF
		rsget.close
		
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
		
		IF ftotalcount > 0 THEN
			strSql = "SELECT TOP "&CStr(FPageSize*FCurrPage)&" A.idx, A.userid, u.username, A.gubun, isNull(A.pk_idx,'') as pk_idx, isNull(A.sub_idx,'') as sub_idx, " & _
					"		A.menupos, m.menuname, m.linkurl, A.contents, A.refip, A.regdate " & _
					"		FROM [db_log].[dbo].[tbl_scm_change_log] AS A with (nolock)" & _
					"	LEFT JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS u with (nolock) ON A.userid = u.userid " & _
					"	LEFT JOIN [db_partner].[dbo].[tbl_partner_menu] AS m with (nolock) ON A.menupos = m.id " & _
					"	WHERE 1=1 " & sqlAdd & _
					"	ORDER BY A.idx DESC "

			'response.write strSql & "<br>"
			rsget.pagesize = FPageSize
			rsget.CursorLocation = adUseClient
			rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
			Redim preserve FItemList(FResultCount)
			i = 0
			If not rsget.EOF Then
				rsget.absolutepage = FCurrPage
				Do until rsget.EOF
					Set FItemList(i) = new cOnlySys_oneitem
						FItemList(i).Fidx = rsget("idx")
						FItemList(i).Fuserid = rsget("userid")
						FItemList(i).Fusername = rsget("username")
						FItemList(i).Fgubun = rsget("gubun")
						FItemList(i).Fpk_idx = rsget("pk_idx")
						FItemList(i).Fsub_idx = rsget("sub_idx")
						FItemList(i).Fmenupos = rsget("menupos")
						FItemList(i).Fmenuname = rsget("menuname")
						FItemList(i).Fmenulink = rsget("linkurl")
						FItemList(i).Fcontents = rsget("contents")
						FItemList(i).Frefip = rsget("refip")
						FItemList(i).Fregdate = rsget("regdate")
						
					i = i + 1
					rsget.moveNext
				Loop
			End If
			rsget.Close
		End If
	End Sub
	
	public Function fnItemDetail
		Dim strSql, sqlAdd
		If FItemID <> "" Then
			sqlAdd = sqlAdd & " AND itemid in(" & FItemID & ") "
		End IF

		strSql = "SELECT itemid, itemname, reserveItemTp, availPayType, regdate, lastupdate " & _
				"		FROM [db_item].[dbo].[tbl_Item] with (nolock)" & _
				"	WHERE 1=1 " & sqlAdd & " "
		'response.write strSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		IF not rsget.EOF THEN
			fnItemDetail = rsget.getRows() 
		END IF	
		rsget.close
	End Function

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
%>