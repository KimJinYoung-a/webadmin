<%
Class outmallItem
	Public FIdx
	Public FMakerid
	Public FMallgubun
	Public FRegdate
	Public FIsComplete
	Public FCurrstat
	Public FHopesellstat
	Public FWhyhope
	Public FAdminText
	Public FHoperegdate
	Public FAdminRegdate
	Public FReguserid
	Public FMallid
	Public FIsusing
	Public FLastupdate
	Public FUpdateid
	Public FHopeidx
	Public FItemid
	Public FItemname
	Public FBigo
	Public FOutmallstandardMargin
End Class

Class cOutmall
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount
	Public FRectmakerid
	Public FRectItemid
	Public FRectItemName
	Public FRectBigo
	Public FRectBigoText
	Public FRectsDt
	Public FRecteDt
	Public FRectCurrstat
	Public FRectMallgubun

	Public FRectTargetmall2
	Public FRectTmpMallList
	Public FRectIdx

	Public Sub getConfirmList
		Dim strSQL, i, addSql

		If FRectmakerid <> "" Then
			addSql = addSql & " and makerid = '"&FRectmakerid&"' "
		End If

		If FRectCurrstat <> "" Then
			addSql = addSql & " and currstat = '"&FRectCurrstat&"' "
		End If

		If FRectsDt <> "" and FRecteDt <> "" Then
			addSql = addSql & " and hoperegdate between '"&FRectsDt&"' and '"&FRecteDt&"' "
		End If

		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell "
		strSQL = strSQL & " WHERE 1=1 "
		strSQL = strSQL & " and isComplete <> 'X' " & addSql
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
		strSQL = strSQL & " idx, makerid, mallgubun, isComplete, currstat, hopesellstat, whyhope, adminText, hoperegdate, adminRegdate "
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell "
		strSQL = strSQL & " WHERE 1=1 "
		strSQL = strSQL & " and isComplete <> 'X' " & addSql
		strSQL = strSQL & " ORDER BY hoperegdate DESC "
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new outmallItem
					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FMallgubun			= rsget("mallgubun")
					FItemList(i).FIsComplete		= rsget("isComplete")
					FItemList(i).FCurrstat			= rsget("currstat")
					FItemList(i).FHopesellstat		= rsget("hopesellstat")
					FItemList(i).FWhyhope			= rsget("whyhope")
					FItemList(i).FAdminText			= rsget("adminText")
					FItemList(i).FHoperegdate		= rsget("hoperegdate")
					FItemList(i).FAdminRegdate		= rsget("adminRegdate")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function isUsingMakerid(makerid)
		Dim strSQL
		strSQL = ""
		strSQL = strSQL & " SELECT COUNT(*) as cnt "
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c as c "
		strSQL = strSQL & " WHERE userid = '"& makerid &"' "
		rsget.Open strSQL, dbget, 1
			isUsingMakerid = rsget("cnt")
		rsget.Close
	End Function

	Public Function fnGetIsExtusing(makerid, byref cisextusing, byref allHopeInsert, byref currstat, byref whyhope, byref adminText, byref adminRegdate, byref idx)
		Dim strSQL
		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 c.isextusing, isnull(h.mallgubun, '') as mallgubun, currstat, whyhope, adminText, adminregdate, h.idx "
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c as c "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.makerid and h.mallgubun = 'all' and isComplete = 'N' "
		strSQL = strSQL & " WHERE userid = '"& makerid &"' "
		rsget.Open strSQL, dbget, 1
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

	Public Sub getOutmallList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c c "
		strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid ni on c.userid=ni.mallGubun and ni.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE c.isusing='Y' and c.userdiv='50' "
		strSQL = strSQL & " and c.userid not in ('nvstorefarm', 'Mylittlewhoopee', 'nvstoregift') "
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
		strSQL = strSQL & " and c.userid not in ('nvstorefarm', 'Mylittlewhoopee', 'nvstoregift') "
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

	'기존에 스토어팜은 등록제외상품에서 제외했으나 수정
	Public Sub getNewOutmallList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c c "
		strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid ni on c.userid=ni.mallGubun and ni.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE c.isusing='Y' and c.userdiv='50' "
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

	Public Sub getAllOutmallList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c c "
		strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid ni on c.userid=ni.mallGubun and ni.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE c.isusing='Y' and c.userdiv='50' "
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

	Public Sub getSpecialOutmallList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c c "
		strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid ni on c.userid=ni.mallGubun and ni.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE c.isusing='Y' and c.userdiv='50' "
		strSQL = strSQL & " and c.userid in ('nvstorefarm', 'Mylittlewhoopee', 'nvstoregift') "
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
		strSQL = strSQL & " and c.userid in ('nvstorefarm', 'Mylittlewhoopee', 'nvstoregift') "
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

	Public Sub getPotalSiteList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_temp.dbo.tbl_Epshop as E  "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_Epshop_not_in_makerid as M on E.mallgubun = M.mallgubun and M.makerid = '"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on E.mallgubun = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid= '"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE E.mallgubun not in ('shodocep', 'wemakepriceep') "		'쇼닥 제휴 종료로 리스트 삭제..2017-08-22 김진영 / 위메프EP 5/31일로 서비스 종료..2023-06-02 김진영
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
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
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on E.mallgubun = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid= '"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE E.mallgubun not in ('shodocep', 'wemakepriceep') "		'쇼닥 제휴 종료로 리스트 삭제..2017-08-22 김진영 / 위메프EP 5/31일로 서비스 종료..2023-06-02 김진영
		strSQL = strSQL & " ORDER BY E.mallgubun ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
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
					FItemList(i).FHopeidx			= rsget("hopeidx")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getOverseasOutmallList
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM db_user.dbo.tbl_user_c c "
		strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=8 "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid ni on c.userid=ni.mallGubun and ni.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE c.isusing='Y' and c.userdiv='50' "
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
		strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=8 "
		strSQL = strSQL & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid ni on c.userid=ni.mallGubun and ni.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " LEFT JOIN db_etcmall.dbo.tbl_jaehumall_hopeSell as h on c.userid = h.mallgubun and h.iscomplete <> 'X' and h.iscomplete <> 'Y' and h.makerid='"&FRectMakerid&"' "
		strSQL = strSQL & " WHERE c.isusing='Y' and c.userdiv='50' "
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

	Public Sub getExtUseList
		Dim strSQL, i, addSQL

		If FRectMakerid <> "" Then
			addSQL = addSQL & " and A.makerid = '"&FRectMakerid&"' "
		End If

		If FRectMallgubun <> "" Then
			addSQL = addSQL & " and A.mallgubun = '"&FRectMallgubun&"' "
		End If

		strSQL = ""
		strSQL = strSQL & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		strSQL = strSQL & " FROM [db_temp].[dbo].[tbl_jaehyumall_not_in_makerid] as A  "
		strSQL = strSQL & " WHERE 1=1 "
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
		strSQL = strSQL & " A.makerid, A.mallgubun, A.regdate, A.idx, A.reguserid "
		strSQL = strSQL & " FROM [db_temp].[dbo].[tbl_jaehyumall_not_in_makerid] as A "
		strSQL = strSQL & " WHERE 1=1 "
		strSQL = strSQL & addSQL
		strSQL = strSQL & " ORDER BY A.mallgubun ASC "
		rsget.pagesize = FPageSize
		rsget.Open strSQL,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new outmallItem
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FMallID			= rsget("mallgubun")
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).FReguserid			= rsget("reguserid")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getJaehyuNotinItemList
		Dim strSql, addSql, tmpMallList, i

		Select Case FRectTargetmall2
			Case "sel"
				tmpMallList = FRectTmpMallList
				tmpMallList = Mid(tmpMallList, 2, 1000)
				tmpMallList = Replace(tmpMallList, ",", "','")
				addSql = addSql & " AND A.mallgubun in ('" & tmpMallList & "') "
			Case "all"
			Case Else
				addSql = addSql & " AND A.mallgubun = '" & FRectMallgubun & "' "
		End Select

		If FRectmakerid <> "" Then
			addSql = addSql & " AND I.makerid = '" & FRectmakerid & "' "
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and A.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and A.itemid in (" + FRectItemid + ")"
            End If
        End If

		If FRectItemName <> "" Then
			addSql = addSql & " AND I.itemname Like '%" & FRectItemName & "%' "
		End If

		Select Case FRectBigo
			Case "Y"
				addSql = addSql & " AND isNull(A.bigo, '') <> ''  "
			Case "N"
				addSql = addSql & " AND isNull(A.bigo, '') = ''  "
		End Select

		If FRectBigo = "Y" and FRectBigoText <> "" Then
			addSql = addSql & " AND isNull(A.bigo, '') like '%"& FRectBigoText &"%'  "
		End If

		strSql = ""
		strSql = strSql & " SELECT Count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		strSql = strSql & " FROM [db_temp].[dbo].[tbl_jaehyumall_not_in_itemid] AS A "
		strSql = strSql & " JOIN [db_item].[dbo].[tbl_item] AS I ON A.itemid = I.itemid "
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		strSql = ""
		strSql = strSql & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		strSql = strSql & " A.itemid, I.itemname, A.mallgubun, isNull(A.bigo, '') as bigo, A.idx "
		strSql = strSql & " FROM [db_temp].[dbo].[tbl_jaehyumall_not_in_itemid] AS A "
		strSql = strSql & "	JOIN [db_item].[dbo].[tbl_item] AS I ON A.itemid = I.itemid "
		strSql = strSql & " Where 1=1 "
		strSql = strSql & addSql
		strSql = strSql & " ORDER BY A.itemid DESC, A.mallgubun"
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				Set FItemList(i) = new outmallItem
					FItemList(i).FItemid		= rsget("itemid")
					FItemList(i).FItemname		= rsget("itemname")
					FItemList(i).FMallgubun		= rsget("mallgubun")
					FItemList(i).FBigo			= rsget("bigo")
					FItemList(i).FIdx			= rsget("idx")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

	Public Sub getOutmallSettingList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c c "
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 "
		sqlStr = sqlStr & " WHERE c.isusing='Y' and c.userdiv='50' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " c.userid as MallID, f.outmallstandardMargin "
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c c "
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 "
		sqlStr = sqlStr & " WHERE c.isusing='Y' and c.userdiv='50' "
		sqlStr = sqlStr & " ORDER BY c.userid ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new outmallItem
					FItemList(i).FMallID				= rsget("MallID")
					FItemList(i).FOutmallstandardMargin	= rsget("outmallstandardMargin")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function fnNotInKeywordList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT [idx], [keyword], [regdate] "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_outmall_not_in_keywords "
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			fnNotInKeywordList = rsget.getRows()
		End If
		rsget.Close
	End Function

	Public Function fnOutmallList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT c.userid as MallID, m.mallid as chkmallid "
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c c "
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid = f.partnerid "
		sqlStr = sqlStr & " left join db_etcmall.[dbo].[tbl_outmall_not_in_keywords_mallid] as m on f.partnerid = m.mallid and m.midx = '"& FRectIdx &"' "
		sqlStr = sqlStr & " WHERE c.isusing='Y' and c.userdiv='50' "
		sqlStr = sqlStr & " and ((f.pcomType = 1 and f.pmallSellType = 1 ) or c.userid = 'ezwel') "
		sqlStr = sqlStr & " ORDER BY c.userid ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			fnOutmallList = rsget.getRows()
		End If
		rsget.Close
	End Function

	Public Function fnOutmallList2
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT Case When c.userid = 'gseshop' THEN 'gsshop' Else c.userId End as MallID, m.mallid as chkmallid  "
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c c "
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid = f.partnerid "
		sqlStr = sqlStr & " left join db_etcmall.[dbo].[tbl_outmall_not_in_keywords_mallid] as m on Case When f.partnerid = 'gseshop' THEN 'gsshop' Else f.partnerid End = m.mallid and m.midx = '"& FRectIdx &"' "
		sqlStr = sqlStr & " WHERE c.isusing='Y' and c.userdiv='50' "
		sqlStr = sqlStr & " and ((f.pcomType = 1 and f.pmallSellType = 1 ) or c.userid = 'ezwel') "
		sqlStr = sqlStr & " ORDER BY c.userid ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			fnOutmallList2 = rsget.getRows()
		End If
		rsget.Close
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

Public Function fnIsRegedHopeCnt(imallid, imakerid)
	Dim strSQL
	strSQL = ""
	strSQL = strSQL & " SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_jaehumall_hopeSell where makerid='"&imakerid&"' and mallgubun='"&imallid&"' and currstat = 1 and isComplete = 'N' "
	rsget.Open strSQL,dbget,1
		fnIsRegedHopeCnt = rsget("cnt")
	rsget.Close
End Function

Public Function goodNoUpdateUser
	Select Case session("ssBctID")
		Case "kjy8517", "as2304", "sj100", "nys1006", "z0516"	goodNoUpdateUser = "Y"
		Case Else goodNoUpdateUser = "N"
	End Select
End Function
%>
