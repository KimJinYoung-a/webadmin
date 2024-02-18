<%
Class outmallItem
	Public FMallID
	Public FIdx
	Public FRegdate
	Public FReguserid
	Public FCurrstat
	Public FWhyhope
	Public FAdminText
	Public FAdminRegdate
	Public FHopeidx

	Public FUseYN
	Public FHopeStr
	Public FMakerid
	Public FIsusing
	Public FLastupdate
	Public FRegid
	Public FUpdateid
End Class

Class cOutmall
	Public FRectMakerid
	Public FRectMallid
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public Sub getOutmallList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c c "
		strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid ni on c.userid=ni.mallGubun and ni.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE c.isusing='Y' and c.userdiv='50' "
		strSQL = strSQL & " and c.userid <> 'nvstorefarm' "
		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		strSQL = ""
		strSQL = strSQL & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		strSQL = strSQL & " c.userid as MallID, isnull(ni.idx, 0) as idx, ni.regdate, ni.reguserid, h.currstat, h.whyhope, h.adminText, h.adminregdate, h.idx as hopeidx "
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c c "
		strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid ni on c.userid=ni.mallGubun and ni.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE c.isusing='Y' and c.userdiv='50' "
		strSQL = strSQL & " and c.userid <> 'nvstorefarm' "
		strSQL = strSQL & " ORDER BY c.userid ASC "
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new outmallItem
					FItemList(i).FMallID			= rsget("MallID")
					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FReguserid			= rsget("reguserid")
					FItemList(i).FCurrstat			= rsget("currstat")
					FItemList(i).FWhyhope			= rsget("whyhope")
					FItemList(i).FAdminText			= rsget("adminText")
					FItemList(i).FAdminRegdate		= rsget("adminregdate")
					FItemList(i).FHopeidx			= rsget("hopeidx")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getOutmallLogList
		Dim strSQL, i, addSQL
'		If FRectMallid <> "all" Then
			addSQL = addSQL & " and mallgubun = '"&FRectMallid&"' "
'		End If

		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell_Log "
		strSQL = strSQL & " WHERE makerid = '"&FRectMakerid&"' "
		strSQL = strSQL & " and mallgubun <> '' "
 		strSQL = strSQL & addSQL
		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		strSQL = ""
		strSQL = strSQL & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		strSQL = strSQL & " mallgubun, makerid, hopeStr, useYN, reguserid, regdate"
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell_Log "
		strSQL = strSQL & " WHERE makerid = '"&FRectMakerid&"'" & addSQL
		strSQL = strSQL & " and mallgubun <> '' "
 		strSQL = strSQL & addSQL
		strSQL = strSQL & " ORDER BY regdate DESC "
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new outmallItem
					FItemList(i).FMallID		= rsget("mallgubun")
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).FHopeStr		= rsget("hopeStr")
					FItemList(i).FUseYN			= rsget("useYN")
					FItemList(i).FReguserid		= rsget("reguserid")
					FItemList(i).FRegdate		= rsget("regdate")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getPotalSiteList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_temp.dbo.tbl_Epshop as E  "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_Epshop_not_in_makerid as M on E.mallgubun = M.mallgubun and M.makerid = '"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on h.makerid = '"&FRectMakerid&"' and E.mallgubun = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' "
		strSQL = strSQL & " WHERE E.mallgubun <> 'shodocep' "		'쇼닥 제휴 종료로 리스트 삭제..2017-08-22 김진영
		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		strSQL = ""
		strSQL = strSQL & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		strSQL = strSQL & " E.mallgubun, M.makerid, isnull(M.isusing, 'Y') as isusing, M.regdate, M.lastupdate, M.regid, M.updateid, h.currstat, h.whyhope, h.adminText, h.adminregdate, h.idx as hopeidx "
		strSQL = strSQL & " FROM db_temp.dbo.tbl_Epshop as E  "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_Epshop_not_in_makerid as M on E.mallgubun = M.mallgubun and M.makerid = '"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on h.makerid = '"&FRectMakerid&"' and E.mallgubun = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' "
		strSQL = strSQL & " WHERE E.mallgubun <> 'shodocep' "		'쇼닥 제휴 종료로 리스트 삭제..2017-08-22 김진영
		strSQL = strSQL & " ORDER BY E.mallgubun ASC "
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new outmallItem
					FItemList(i).FMallID			= rsget("mallgubun")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FIsusing			= rsget("isusing")
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FLastupdate		= rsget("lastupdate")
					FItemList(i).FReguserid			= rsget("regid")
					FItemList(i).FUpdateid			= rsget("updateid")
					FItemList(i).FCurrstat			= rsget("currstat")
					FItemList(i).FWhyhope			= rsget("whyhope")
					FItemList(i).FAdminText			= rsget("adminText")
					FItemList(i).FAdminRegdate		= rsget("adminregdate")
					FItemList(i).FHopeidx		= rsget("hopeidx")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function fnGetIsExtusing(makerid, byref cisextusing, byref allHopeInsert, byref currstat, byref whyhope, byref adminText, byref adminRegdate, byref idx)
		Dim strSQL
		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 c.isextusing, isnull(h.mallgubun, '') as mallgubun, currstat, whyhope, adminText, adminregdate, h.idx "
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c as c "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.makerid and h.mallgubun = 'all' and isComplete = 'N' "
		strSQL = strSQL & " WHERE userid = '"& makerid &"' "
		rsget.Open strSQL, dbget
		If Not rsget.Eof then
			cisextusing = rsget("isextusing")
			If rsget("mallgubun") <> "" Then
				allHopeInsert = "Y"
			Else
				allHopeInsert = "N"
			End If
			currstat		= rsget("currstat")
			whyhope			= rsget("whyhope")
			adminText		= rsget("adminText")
			adminregdate	= rsget("adminregdate")
			idx				= rsget("idx")
		Else
			cisextusing = "N"
			allHopeInsert	= "N"
		End If
		rsget.close
	End Function

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
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

Public Function fnHoperegConfirm(makerid)
	Dim strSQL
	strSQL = ""
	strSQL = strSQL & " SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_jaehumall_hopeSell "
	strSQL = strSQL & " WHERE makerid = '"&makerid&"' and mallgubun <> 'all' and mallgubun <> 'naverep' and mallgubun <> 'daumep' and mallgubun <> 'shodocep' and iscomplete <> 'X' and currstat <> '3' "
	rsget.Open strSQL, dbget
	If rsget("cnt") > 0 Then
		fnHoperegConfirm = True
	Else
		fnHoperegConfirm = False
	End If
	rsget.Close
End Function

Public Function fnIsRegedHopeCnt(imallid, imakerid)
	Dim strSQL
	strSQL = ""
	strSQL = strSQL & " SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_jaehumall_hopeSell where makerid='"&imakerid&"' and mallgubun='"&vMallid&"' and currstat = 1 and isComplete = 'N' "
	rsget.Open strSQL,dbget,1
		fnIsRegedHopeCnt = rsget("cnt")
	rsget.Close
End Function
%>