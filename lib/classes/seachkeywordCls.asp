<%
Class CSearchKeywordItem
    public FprevKeyword
	public FcurrKeyword
    public Fcount
	public FrankPrevDay
	public FrankPrevWeek

	public Fyyyymmdd
	public FkeywordRank

	public Fminmxrectcnt
	public Fmaxmxrectcnt
	public Favgmxrectcnt

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CRelatedKeywordItem
	public Fidx
    public ForgKeyword
	public FrelatedKeyword
    public FsearchCount
	public FkeywordRank

	public FmodiType
	public Freguserid
	public FuseYN
	public Fregdate

	public function GetModiTypeName()
		select case FmodiType
			case "A"
				GetModiTypeName = "추가"
			case "D"
				GetModiTypeName = "제외"
			case else
				GetModiTypeName = "ERR" & FmodiType
		end select
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CRelatedKeywordModiItem
    public FprevKeyword
	public FcurrKeyword
    public Fcount
	public FrankPrevDay
	public FrankPrevWeek

	public Fyyyymmdd
	public FkeywordRank

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CLowResultKeywordItem
    public Frect
	public Fsumsearchcnt
	public FmxrectCNT

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CTopKeywordItem
	public FtopKeyword
    public FsearchCount

	public Fidx
	public FmodiType
	public Freguserid
	public FuseYN
	public Fregdate

	public function GetModiTypeName()
		select case FmodiType
			case "A"
				GetModiTypeName = "추가"
			case "D"
				GetModiTypeName = "제외"
			case else
				GetModiTypeName = "ERR" & FmodiType
		end select
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CRelatedKeywordNewItem
	public FRowNum
	public Fprect
	public Frect
	public FacctCNT
	public FacctSearchCNT
	public FlastrectCNT
	public FlastpRectCNT
	public FisAutoType
	public FisUsingType
	public FUserAddCNT
	public Fregdate
	public Flastupdate
	public FenginAssign
	public FrecentAcctCNT
	public FrecentAcctCNT2

	public function GetIsAutoTypeName()
		select case FisAutoType
			case 1
				GetIsAutoTypeName = "N"
			case 2
				GetIsAutoTypeName = "Y"
			case else
				GetIsAutoTypeName = "ERR" & FisAutoType
		end select
	end function

	public function GetIsUsingTypeName()
		select case FIsUsingType
			case 1
				GetIsUsingTypeName = "Y"
			case 0
				GetIsUsingTypeName = "N"
			case else
				GetIsUsingTypeName = "ERR" & FIsUsingType
		end select
	end function

	'// FisAutoType

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

class CSearchKeyword
    public FItemList()
	public FOneItem
	public FResultArray

    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectStart
	public FRectEnd
	public FRectBaseDate
	public FRectGroupBy
	public FRectKeyword
	public FRectOrgKeyword
	public FRectRelatedKeyword
	public FRectUseYN
	public FRectModiType
	public FRectIsEnginMayAssign
	public FRectPlatform

	public FRectYYYYMMDD
	public FRectMxrectCNT
	public FRectSearchCNT

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FPageSize = 20
		FTotalPage = 0
		FPageCount = 0
		FResultCount = 0
		FScrollCount = 10
		FCurrPage = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	'// 인기검색어 - 위탁기간
	public function getReportByPopular()
		Dim strSql, i

		strSql = " exec [db_datamart].[dbo].[usp_Ten_Datamart_Get_Keyword_By_Popular] '" + CStr(FRectStart) + "', '" + CStr(FRectEnd) + "' "
		''response.write strSql & "<br>"

		db3_rsget.CursorLocation = 3
		db3_rsget.Open strSql, db3_dbget, 3, 1

		FTotalCount = db3_rsget.RecordCount
		redim FItemList(FTotalCount)

		if not db3_rsget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CSearchKeywordItem

					FItemList(i).FcurrKeyword	= db3_rsget("currKeyword")
					FItemList(i).Fcount			= db3_rsget("cnt")
				db3_rsget.movenext
			next
		end if
		db3_rsget.close

	End Function

	public function getReportByPopularEVT()
		Dim strSql, i

		strSql = " exec [db_EVT].[dbo].[usp_Ten_itemevent_Get_Keyword_By_Popular] '" + CStr(FRectStart) + " 00:00:00', '" + CStr(FRectEnd) + " 23:00:00', '" & FRectPlatform & "', '" & FRectKeyword & "' "
		''response.write strSql & "<br>"

        rsEVTget.CursorLocation = adUseClient
        rsEVTget.Open strSQL, dbEVTget, adOpenForwardOnly, adLockReadOnly

		'rsEVTget.CursorLocation = 3
		'rsEVTget.Open strSql, dbEVTget, 3, 1

		FTotalCount = rsEVTget.RecordCount
		redim FItemList(FTotalCount)

		if not rsEVTget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CSearchKeywordItem

					FItemList(i).FcurrKeyword	= rsEVTget("rect")
					FItemList(i).Fcount			= rsEVTget("searchcnt")
					FItemList(i).Fminmxrectcnt	= rsEVTget("minmxrectcnt")
					FItemList(i).Fmaxmxrectcnt	= rsEVTget("maxmxrectcnt")
					FItemList(i).Favgmxrectcnt	= rsEVTget("avgmxrectcnt")
				rsEVTget.movenext
			next
		end if
		rsEVTget.close

	End Function

	'// 인기검색어 - 위탁일
	public function getReportByPopularAndDay()
		Dim strSql, i

		strSql = " exec [db_datamart].[dbo].[usp_Ten_Datamart_Get_Keyword_By_Popular_And_Day] '" + CStr(FRectBaseDate) + "' "
		''response.write strSql & "<br>"

		db3_rsget.CursorLocation = 3
		db3_rsget.Open strSql, db3_dbget, 3, 1

		FTotalCount = db3_rsget.RecordCount
		redim FItemList(FTotalCount)

		if not db3_rsget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CSearchKeywordItem

					FItemList(i).FcurrKeyword	= db3_rsget("currKeyword")
					FItemList(i).Fcount			= db3_rsget("cnt")

					FItemList(i).FrankPrevDay	= db3_rsget("rankPrevDay")
					FItemList(i).FrankPrevWeek	= db3_rsget("rankPrevWeek")
				db3_rsget.movenext
			next
		end if
		db3_rsget.close

	End Function

	'// 위탁검색어 - 검색트랜드
	public function getReportByTrand()
		Dim strSql, i

		strSql = " exec [db_datamart].[dbo].[usp_Ten_Datamart_Get_Keyword_By_Trand] '" + CStr(FRectStart) + "', '" + CStr(FRectEnd) + "', '" + html2db(CStr(FRectKeyword)) + "' "
		''response.write strSql & "<br>"

		db3_rsget.CursorLocation = 3
		db3_rsget.Open strSql, db3_dbget, 3, 1

		FTotalCount = db3_rsget.RecordCount
		redim FItemList(FTotalCount)

		if not db3_rsget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CSearchKeywordItem

					FItemList(i).Fyyyymmdd		= db3_rsget("yyyymmdd")
					FItemList(i).FkeywordRank	= db3_rsget("keywordRank")
				db3_rsget.movenext
			next
		end if
		db3_rsget.close

	End Function

	'// 위탁검색어 - 연관검색어
	public function getReportByRelated()
		Dim strSql, i

		strSql = " exec [db_datamart].[dbo].[usp_Ten_Datamart_Get_Keyword_By_Related] '" + CStr(FRectStart) + "', '" + CStr(FRectEnd) + "', '" + html2db(CStr(FRectKeyword)) + "' "
		''response.write strSql & "<br>"

		db3_rsget.CursorLocation = 3
		db3_rsget.Open strSql, db3_dbget, 3, 1

		FTotalCount = db3_rsget.RecordCount
		redim FItemList(FTotalCount)

		if not db3_rsget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CSearchKeywordItem

					FItemList(i).FprevKeyword	= db3_rsget("prevKeyword")
					FItemList(i).FcurrKeyword	= db3_rsget("currKeyword")
					FItemList(i).Fcount			= db3_rsget("cnt")
				db3_rsget.movenext
			next
		end if
		db3_rsget.close

	End Function

	'// 연관검색어 : 서비스용
	public function getRelatedKeywordReal()
		Dim strSql, i

		strSql = " exec [db_log].[dbo].[sp_Ten_SearchKey_related_REAL] '" + CStr(FRectOrgKeyword) + "', '" + CStr(FRectRelatedKeyword) + "' "
		''response.write strSql & "<br>"

		rsget.CursorLocation = 3
		rsget.Open strSql, dbget, 3, 1

		FTotalCount = rsget.RecordCount
		redim FItemList(FTotalCount)

		if not rsget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CRelatedKeywordItem

				FItemList(i).ForgKeyword		= rsget("orgKeyword")
				FItemList(i).FrelatedKeyword	= rsget("relatedKeyword")
				FItemList(i).FsearchCount		= rsget("searchCount")
				FItemList(i).FkeywordRank		= rsget("keywordRank")

				rsget.movenext
			next
		end if
		rsget.close

	End Function

	'// 연관검색어 : 서비스용(추가/제외)
	public function getRelatedKeywordModi()
		Dim strSql, i

		strSql = " exec [db_log].[dbo].[sp_Ten_SearchKey_related_MODI] '" + CStr(FRectOrgKeyword) + "', '" + CStr(FRectRelatedKeyword) + "', '" + CStr(FRectModiType) + "', '" + CStr(FRectUseYN) + "' "
		''response.write strSql & "<br>"

		rsget.CursorLocation = 3
		rsget.Open strSql, dbget, 3, 1

		FTotalCount = rsget.RecordCount
		redim FItemList(FTotalCount)

		if not rsget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CRelatedKeywordItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).ForgKeyword		= rsget("orgKeyword")
				FItemList(i).FrelatedKeyword	= rsget("relatedKeyword")
				FItemList(i).FmodiType			= rsget("modiType")
				FItemList(i).Freguserid			= rsget("reguserid")
				FItemList(i).FuseYN				= rsget("useYN")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FsearchCount		= rsget("searchCount")

				rsget.movenext
			next
		end if
		rsget.close

	End Function

	'// 연관검색어 : 서비스용(추가/제외)
	public function getRelatedKeywordModi_Paging()
		Dim strSql, i

		strSql = " exec [db_log].[dbo].[sp_Ten_SearchKey_related_MODI_CNT] '" + CStr(FRectOrgKeyword) + "', '" + CStr(FRectRelatedKeyword) + "', '" + CStr(FRectModiType) + "', '" + CStr(FRectUseYN) + "' "
		''response.write strSql & "<br>"
		''response.end
		rsget.Open strSql,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		FTotalPage =  CLng(FTotalCount\FPageSize)
		If (FTotalCount\FPageSize) <> (FTotalCount/FPageSize) Then
			FTotalPage = FTotalPage + 1
		End If

		strSql = " exec [db_log].[dbo].[sp_Ten_SearchKey_related_MODI_LIST] '" + CStr(FRectOrgKeyword) + "', '" + CStr(FRectRelatedKeyword) + "', '" + CStr(FRectModiType) + "', '" + CStr(FRectUseYN) + "', " + CStr(FCurrPage) + ", " + CStr(FPageSize) + " "
		''response.write strSql & "<br>"
		''response.end
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount
		redim FItemList(FResultCount)

		if not rsget.eof then
			for i = 0 to FResultCount - 1
				set FItemList(i) = new CRelatedKeywordItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).ForgKeyword		= rsget("orgKeyword")
				FItemList(i).FrelatedKeyword	= rsget("relatedKeyword")
				FItemList(i).FmodiType			= rsget("modiType")
				FItemList(i).Freguserid			= rsget("reguserid")
				FItemList(i).FuseYN				= rsget("useYN")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FsearchCount		= rsget("searchCount")

				rsget.movenext
			next
		end if
		rsget.close

	End Function

	'// 인기검색어 : 서비스용
	public function getTopKeywordReal()
		Dim strSql, i

		strSql = " exec [db_log].[dbo].[sp_Ten_SearchKey_top_REAL] "
		''response.write strSql & "<br>"

		rsget.CursorLocation = 3
		rsget.Open strSql, dbget, 3, 1

		FTotalCount = rsget.RecordCount
		redim FItemList(FTotalCount)

		if not rsget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CTopKeywordItem

				FItemList(i).FtopKeyword		= rsget("topKeyword")
				FItemList(i).FsearchCount		= rsget("searchCount")

				rsget.movenext
			next
		end if
		rsget.close

	End Function

	'// 인기검색어 : 서비스용(추가/제외)
	public function getTopKeywordModi()
		Dim strSql, i

		strSql = " exec [db_log].[dbo].[sp_Ten_SearchKey_top_MODI] '" + CStr(FRectKeyword) + "', '" + CStr(FRectModiType) + "', '" + CStr(FRectUseYN) + "' "
		''response.write strSql & "<br>"

		rsget.CursorLocation = 3
		rsget.Open strSql, dbget, 3, 1

		FTotalCount = rsget.RecordCount
		redim FItemList(FTotalCount)

		if not rsget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CTopKeywordItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).FtopKeyword		= rsget("topKeyword")
				FItemList(i).FmodiType			= rsget("modiType")
				FItemList(i).Freguserid			= rsget("reguserid")
				FItemList(i).FuseYN				= rsget("useYN")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FsearchCount		= rsget("searchCount")

				rsget.movenext
			next
		end if
		rsget.close

	End Function

	public function GetRelatedKeywordList()
		Dim strSql, i
		dim IsUsing, IsEnginMayAssign
		dim cmd, rs

		IsEnginMayAssign = NULL
		if (FRectIsEnginMayAssign = CStr(1)) then
			IsEnginMayAssign = CInt(FRectIsEnginMayAssign)
		end if
		IsUsing = NULL
		if (FRectUseYN <> "") then
			select case FRectUseYN
				case "Y"
					IsUsing = 1
				case "N"
					IsUsing = 0
				case else
					'//
			end select
		end if

		'// https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/createparameter-method-ado
		set cmd = CreateObject("ADODB.Command")

		cmd.ActiveConnection = dbEVTget
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "[db_EVT].[dbo].[sp_Ten_Keyword_Relate_GetList]"
		cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
		cmd.Parameters.Append cmd.CreateParameter("@PageSize", adInteger, adParamInput, , FPageSize)
		cmd.Parameters.Append cmd.CreateParameter("@CurrPage", adInteger, adParamInput, , FCurrPage)
		cmd.Parameters.Append cmd.CreateParameter("@prect", adVarChar, adParamInput, 50, FRectOrgKeyword)
		cmd.Parameters.Append cmd.CreateParameter("@rect", adVarChar, adParamInput, 50, FRectRelatedKeyword)
		cmd.Parameters.Append cmd.CreateParameter("@isEnginMayAssign", adInteger, adParamInput, , IsEnginMayAssign)
		cmd.Parameters.Append cmd.CreateParameter("@isAutoType", adInteger, adParamInput, , NULL)
		cmd.Parameters.Append cmd.CreateParameter("@isUsingType", adInteger, adParamInput, , IsUsing)
		rsEVTget.CursorLocation = adUseClient
		rsEVTget.open cmd, , adOpenStatic, adLockReadOnly

		FTotalCount = cmd.Parameters("returnValue")

		FTotalPage =  CLng(FTotalCount\FPageSize)
		If (FTotalCount\FPageSize) <> (FTotalCount/FPageSize) Then
			FTotalPage = FTotalPage + 1
		End If

		FResultCount = rsEVTget.RecordCount
		redim FItemList(FResultCount)

		if not rsEVTget.eof then
			for i = 0 to FResultCount - 1
				set FItemList(i) = new CRelatedKeywordNewItem

				'// RowNum, prect, rect, acctCNT, acctSearchCNT, lastrectCNT, lastpRectCNT, isAutoType, isUsingType, UserAddCNT, regdate, lastupdate, enginAssign, recentAcctCNT, recentAcctCNT2

				FItemList(i).FRowNum			= rsEVTget("RowNum")
				FItemList(i).Fprect				= rsEVTget("prect")
				FItemList(i).Frect				= rsEVTget("rect")
				FItemList(i).FacctCNT			= rsEVTget("acctCNT")
				FItemList(i).FacctSearchCNT		= rsEVTget("acctSearchCNT")
				FItemList(i).FlastrectCNT		= rsEVTget("lastrectCNT")
				FItemList(i).FlastpRectCNT		= rsEVTget("lastpRectCNT")
				FItemList(i).FisAutoType		= rsEVTget("isAutoType")
				FItemList(i).FisUsingType		= rsEVTget("isUsingType")
				FItemList(i).FUserAddCNT		= rsEVTget("UserAddCNT")
				FItemList(i).Fregdate			= rsEVTget("regdate")
				FItemList(i).Flastupdate		= rsEVTget("lastupdate")
				FItemList(i).FenginAssign		= rsEVTget("enginAssign")
				FItemList(i).FrecentAcctCNT		= rsEVTget("recentAcctCNT")
				FItemList(i).FrecentAcctCNT2	= rsEVTget("recentAcctCNT2")

				rsEVTget.movenext
			next
		end if
		rsEVTget.close
		set cmd = Nothing
	end function

	public function GetCorrectKeywordList()
		Dim strSql, i
		dim IsUsing, IsEnginMayAssign
		dim cmd, rs

		IsEnginMayAssign = NULL
		if (FRectIsEnginMayAssign = CStr(1)) then
			IsEnginMayAssign = CInt(FRectIsEnginMayAssign)
		end if
		IsUsing = NULL
		if (FRectUseYN <> "") then
			select case FRectUseYN
				case "Y"
					IsUsing = 1
				case "N"
					IsUsing = 0
				case else
					'//
			end select
		end if

		'// https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/createparameter-method-ado
		set cmd = CreateObject("ADODB.Command")

		cmd.ActiveConnection = dbEVTget
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "[db_EVT].[dbo].[sp_Ten_Keyword_Correct_GetList]"
		cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
		cmd.Parameters.Append cmd.CreateParameter("@PageSize", adInteger, adParamInput, , FPageSize)
		cmd.Parameters.Append cmd.CreateParameter("@CurrPage", adInteger, adParamInput, , FCurrPage)
		cmd.Parameters.Append cmd.CreateParameter("@prect", adVarChar, adParamInput, 50, FRectOrgKeyword)
		cmd.Parameters.Append cmd.CreateParameter("@rect", adVarChar, adParamInput, 50, FRectRelatedKeyword)
		cmd.Parameters.Append cmd.CreateParameter("@isEnginMayAssign", adInteger, adParamInput, , IsEnginMayAssign)
		cmd.Parameters.Append cmd.CreateParameter("@isAutoType", adInteger, adParamInput, , NULL)
		cmd.Parameters.Append cmd.CreateParameter("@isUsingType", adInteger, adParamInput, , IsUsing)
		rsEVTget.CursorLocation = adUseClient
		rsEVTget.open cmd, , adOpenStatic, adLockReadOnly

		FTotalCount = cmd.Parameters("returnValue")

		rw "aa" & IsEnginMayAssign

		FTotalPage =  CLng(FTotalCount\FPageSize)
		If (FTotalCount\FPageSize) <> (FTotalCount/FPageSize) Then
			FTotalPage = FTotalPage + 1
		End If

		FResultCount = rsEVTget.RecordCount
		redim FItemList(FResultCount)

		if not rsEVTget.eof then
			for i = 0 to FResultCount - 1
				set FItemList(i) = new CRelatedKeywordNewItem

				'// RowNum, prect, rect, acctCNT, acctSearchCNT, lastrectCNT, lastpRectCNT, isAutoType, isUsingType, UserAddCNT, regdate, lastupdate, enginAssign, recentAcctCNT, recentAcctCNT2

				FItemList(i).FRowNum			= rsEVTget("RowNum")
				FItemList(i).Fprect				= rsEVTget("prect")
				FItemList(i).Frect				= rsEVTget("rect")
				FItemList(i).FacctCNT			= rsEVTget("acctCNT")
				FItemList(i).FacctSearchCNT		= rsEVTget("acctSearchCNT")
				FItemList(i).FlastrectCNT		= rsEVTget("lastrectCNT")
				FItemList(i).FlastpRectCNT		= rsEVTget("lastpRectCNT")
				FItemList(i).FisAutoType		= rsEVTget("isAutoType")
				FItemList(i).FisUsingType		= rsEVTget("isUsingType")
				FItemList(i).FUserAddCNT		= rsEVTget("UserAddCNT")
				FItemList(i).Fregdate			= rsEVTget("regdate")
				FItemList(i).Flastupdate		= rsEVTget("lastupdate")
				FItemList(i).FenginAssign		= rsEVTget("enginAssign")
				FItemList(i).FrecentAcctCNT		= rsEVTget("recentAcctCNT")
				FItemList(i).FrecentAcctCNT2	= rsEVTget("recentAcctCNT2")

				rsEVTget.movenext
			next
		end if
		rsEVTget.close
		set cmd = Nothing
	end function

	public function GetAutoCompleteKeywordList()
		Dim strSql, i
		dim IsUsing, IsEnginMayAssign
		dim cmd, rs

		IsEnginMayAssign = NULL
		if (FRectIsEnginMayAssign = CStr(1)) then
			IsEnginMayAssign = CInt(FRectIsEnginMayAssign)
		end if
		IsUsing = NULL
		if (FRectUseYN <> "") then
			select case FRectUseYN
				case "Y"
					IsUsing = 1
				case "N"
					IsUsing = 0
				case else
					'//
			end select
		end if

		'// https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/createparameter-method-ado
		set cmd = CreateObject("ADODB.Command")

		cmd.ActiveConnection = dbEVTget
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "[db_EVT].[dbo].[sp_Ten_Keyword_AutoComplete_GetList]"
		cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
		cmd.Parameters.Append cmd.CreateParameter("@PageSize", adInteger, adParamInput, , FPageSize)
		cmd.Parameters.Append cmd.CreateParameter("@CurrPage", adInteger, adParamInput, , FCurrPage)
		cmd.Parameters.Append cmd.CreateParameter("@rect", adVarChar, adParamInput, 50, FRectOrgKeyword)
		cmd.Parameters.Append cmd.CreateParameter("@isEnginMayAssign", adInteger, adParamInput, , IsEnginMayAssign)
		cmd.Parameters.Append cmd.CreateParameter("@isAutoType", adInteger, adParamInput, , NULL)
		cmd.Parameters.Append cmd.CreateParameter("@isUsingType", adInteger, adParamInput, , IsUsing)
		rsEVTget.CursorLocation = adUseClient
		rsEVTget.open cmd, , adOpenStatic, adLockReadOnly

		FTotalCount = cmd.Parameters("returnValue")

		FTotalPage =  CLng(FTotalCount\FPageSize)
		If (FTotalCount\FPageSize) <> (FTotalCount/FPageSize) Then
			FTotalPage = FTotalPage + 1
		End If

		FResultCount = rsEVTget.RecordCount
		redim FItemList(FResultCount)

		if not rsEVTget.eof then
			for i = 0 to FResultCount - 1
				set FItemList(i) = new CRelatedKeywordNewItem

				'// RowNum, rect, acctCNT, acctSearchCNT, lastrectCNT, lastpRectCNT, isAutoType, isUsingType, UserAddCNT, regdate, lastupdate, enginAssign, recentAcctCNT, recentAcctCNT2

				FItemList(i).FRowNum			= rsEVTget("RowNum")
				FItemList(i).Frect				= rsEVTget("rect")
				FItemList(i).FacctCNT			= rsEVTget("acctCNT")
				FItemList(i).FacctSearchCNT		= rsEVTget("acctSearchCNT")
				FItemList(i).FlastrectCNT		= rsEVTget("lastrectCNT")
				FItemList(i).FisAutoType		= rsEVTget("isAutoType")
				FItemList(i).FisUsingType		= rsEVTget("isUsingType")
				FItemList(i).FUserAddCNT		= rsEVTget("UserAddCNT")
				FItemList(i).Fregdate			= rsEVTget("regdate")
				FItemList(i).Flastupdate		= rsEVTget("lastupdate")
				FItemList(i).FenginAssign		= rsEVTget("enginAssign")
				FItemList(i).FrecentAcctCNT		= rsEVTget("recentAcctCNT")
				FItemList(i).FrecentAcctCNT2	= rsEVTget("recentAcctCNT2")

				rsEVTget.movenext
			next
		end if
		rsEVTget.close
		set cmd = Nothing
	end function

	public function GetLowResultKeywordList()
		Dim strSql, i
		dim cmd, rs

		'// https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/createparameter-method-ado
		set cmd = CreateObject("ADODB.Command")

		cmd.ActiveConnection = dbEVTget
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "[db_EVT].[dbo].[usp_Ten_Keyword_LowResult_getList]"
		cmd.Parameters.Append cmd.CreateParameter("@rect", adVarChar, adParamInput, 10, FRectYYYYMMDD)
		cmd.Parameters.Append cmd.CreateParameter("@mxrectCNT", adInteger, adParamInput, , FRectMxrectCNT)
		cmd.Parameters.Append cmd.CreateParameter("@searchcnt", adInteger, adParamInput, , FRectSearchCNT)
		rsEVTget.CursorLocation = adUseClient
		rsEVTget.open cmd, , adOpenStatic, adLockReadOnly

		FResultCount = rsEVTget.RecordCount
		redim FItemList(FResultCount)

		if not rsEVTget.eof then
			for i = 0 to FResultCount - 1
				set FItemList(i) = new CLowResultKeywordItem

				'// rect,sum(searchcnt) sumsearchcnt,avg(mxrectCNT) mxrectCNT

				FItemList(i).Frect				= db2html(rsEVTget("rect"))
				FItemList(i).Fsumsearchcnt		= rsEVTget("sumsearchcnt")
				FItemList(i).FmxrectCNT			= rsEVTget("mxrectCNT")

				rsEVTget.movenext
			next
		end if
		rsEVTget.close
		set cmd = Nothing
	end function

	public function GetRelatedKeywordList_API()
		Dim strSql, i
		dim cmd, rs

		'// https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/createparameter-method-ado
		set cmd = CreateObject("ADODB.Command")

		cmd.ActiveConnection = dbEVTget
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "[db_EVT].[dbo].[sp_Ten_Keyword_Relate_GetArray]"
		rsEVTget.CursorLocation = adUseClient
		rsEVTget.open cmd, , adOpenStatic, adLockReadOnly
		If Not rsEVTget.EOF Then
			FResultArray = rsEVTget.GetRows
		end if
		rsEVTget.close
		set cmd = Nothing
	end function

	public function GetCorrectKeywordList_API()
		Dim strSql, i
		dim cmd, rs

		'// https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/createparameter-method-ado
		set cmd = CreateObject("ADODB.Command")

		cmd.ActiveConnection = dbEVTget
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "[db_EVT].[dbo].[sp_Ten_Keyword_Correct_GetArray]"
		rsEVTget.CursorLocation = adUseClient
		rsEVTget.open cmd, , adOpenStatic, adLockReadOnly
		If Not rsEVTget.EOF Then
			FResultArray = rsEVTget.GetRows
		end if
		rsEVTget.close
		set cmd = Nothing
	end function

	public function GetPopularKeywordList_API()
		Dim strSql, i
		dim cmd, rs

		set cmd = CreateObject("ADODB.Command")

		cmd.ActiveConnection = dbEVTget
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "[db_EVT].[dbo].[sp_Ten_Keyword_Popular_GetList_UseSearchEngin]"
		rsEVTget.CursorLocation = adUseClient
		rsEVTget.open cmd, , adOpenStatic, adLockReadOnly
		If Not rsEVTget.EOF Then
			FResultArray = rsEVTget.GetRows
		end if
		rsEVTget.close
		set cmd = Nothing
	end function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end class

%>
