<%

class CEmergencyQuestionMasterItem
	Public Fidx
	Public FupcheGubun			'// 1=텐바이텐 고객센터, 2=텐바이텐 물류센터, 3=업체
	Public FupcheName
	Public Fmakerid
	Public FcategoryGubun
	Public FcategoryName
	Public FneedReplyYN
	Public Ftitle
	Public Fcontents
	Public Forderserial
	Public FbuyName
	Public Fitemids
	Public Fdeleteyn
	Public FcurrState
	Public FdeadlineDate
	Public Fregdate
	Public FregUserid
	Public FlastUpdate

	Public Function GetRegdateFormatString()
		'// 2019-02-14 15:15:01.380
		dim regdate
		regdate = CDate(Left(Fregdate, 19))
		if DateDiff("d", regdate, Now()) = 0 then
			GetRegdateFormatString = Mid(Fregdate, 12, 5)
		else
			GetRegdateFormatString = Mid(Fregdate, 1, 16)
		end if
	End Function
end class

class CEmergencyQuestionMaster
	Private EntityManager, ObjConn

    public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectRegStart
	public FRectRegEnd
	public FRectCategoryGubun
	public FRectCurrState
	public FRectOrdBy
	Public FRectSearchField
	Public FRectSearchString
	Public FRectMyCsOnly
	Public FRectShowUsingOnly

	Public Default Function Init(Conn, Rs)
        Set EntityManager = (new ClassEntityManager)("db_cs.dbo.tbl_emergency_question_master", "idx", me)

		Call EntityManager.Register("idx", "Fidx", "int")
		Call EntityManager.Register("upcheGubun", "FupcheGubun", "string")
		Call EntityManager.Register("upcheName", "FupcheName", "string")
		Call EntityManager.Register("makerid", "Fmakerid", "string")
		Call EntityManager.Register("categoryGubun", "FcategoryGubun", "string")
		Call EntityManager.Register("categoryName", "FcategoryName", "string")
		Call EntityManager.Register("needReplyYN", "FneedReplyYN", "string")
		Call EntityManager.Register("title", "Ftitle", "string")
		Call EntityManager.Register("contents", "Fcontents", "string")
		Call EntityManager.Register("orderserial", "Forderserial", "string")
		Call EntityManager.Register("buyName", "FbuyName", "string")
		Call EntityManager.Register("itemids", "Fitemids", "string")
		Call EntityManager.Register("deleteyn", "Fdeleteyn", "string")
		Call EntityManager.Register("currState", "FcurrState", "string")
		Call EntityManager.Register("deadlineDate", "FdeadlineDate", "string")
		Call EntityManager.Register("regdate", "Fregdate", "string")
		Call EntityManager.Register("regUserid", "FregUserid", "string")
		Call EntityManager.Register("lastUpdate", "FlastUpdate", "string")

		Set EntityManager.ObjConn = Conn
		Set EntityManager.ObjRs2 = Rs
		Set Init = me
	End Function

    Public Sub Save()
		'// idx 가 없으면 insert, 있으면 update
        Call EntityManager.Save()
    End Sub

    Public Sub Delete()
        ''Call EntityManager.Delete()
    End Sub

    Public Sub LoadOne(pIdx)
		me.FOneItem.Fidx = pIdx
        Call EntityManager.LoadOne()
    End Sub

    Public Sub LoadList()
	    Dim sqlStr, addSql, i
		Dim countQuery, selectQuery

		EntityManager.ResetDictionary()

		Call EntityManager.Register("idx", "Fidx", "int")
		Call EntityManager.Register("upcheGubun", "FupcheGubun", "string")
		Call EntityManager.Register("upcheName", "FupcheName", "string")
		Call EntityManager.Register("makerid", "Fmakerid", "string")
		Call EntityManager.Register("categoryGubun", "FcategoryGubun", "string")
		Call EntityManager.Register("categoryName", "FcategoryName", "string")
		Call EntityManager.Register("needReplyYN", "FneedReplyYN", "string")
		Call EntityManager.Register("title", "Ftitle", "string")
		Call EntityManager.Register("contents", "Fcontents", "string")
		Call EntityManager.Register("orderserial", "Forderserial", "string")
		Call EntityManager.Register("buyName", "FbuyName", "string")
		Call EntityManager.Register("itemids", "Fitemids", "string")
		Call EntityManager.Register("deleteyn", "Fdeleteyn", "string")
		Call EntityManager.Register("currState", "FcurrState", "string")
		Call EntityManager.Register("deadlineDate", "FdeadlineDate", "string")
		Call EntityManager.Register("regdate", "Fregdate", "string")
		Call EntityManager.Register("regUserid", "FregUserid", "string")
		Call EntityManager.Register("lastUpdate", "FlastUpdate", "string")

		addSql = addSql & " where 1=1 "

		if (FRectRegStart <> "") then
			addSql = addSql & " and regdate >= '" & FRectRegStart & "' "
		end if

		if (FRectRegStart <> "") then
			addSql = addSql & " and regdate < '" & FRectRegEnd & "' "
		end if

		if (FRectCategoryGubun <> "") then
			addSql = addSql & " and categoryGubun = '" & FRectCategoryGubun & "' "
		end if

		if (FRectCurrState <> "") then
			addSql = addSql & " and currState = '" & FRectCurrState & "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchString <> "") then
			select case FRectSearchField
				case "orderserial"
					addSql = addSql & " and orderserial = '" & FRectSearchString & "' "
				case "regUserid"
					addSql = addSql & " and regUserid = '" & FRectSearchString & "' "
				case "makerid"
					addSql = addSql & " and makerid = '" & FRectSearchString & "' "
				case else
					'//
			end select
		end if

		if (FRectMyCsOnly = "Y") then
			addSql = addSql & " and regUserid = '" & session("ssBctId") & "' "
		end if

		if (FRectShowUsingOnly = "Y") then
			addSql = addSql & " and deleteyn = 'N' "
		end if


		'// ====================================================================
		sqlStr = "select count(idx) as cnt "
		sqlStr = sqlStr + " from db_cs.dbo.tbl_emergency_question_master "
        sqlStr = sqlStr + addSql
		countQuery = sqlStr
		''response.write sqlStr

		'// ====================================================================
	    sqlStr = "select top " + CStr(FCurrPage * FPageSize) + " idx, upcheGubun, upcheName, makerid, categoryGubun, categoryName, needReplyYN, title, contents, orderserial, buyName, itemids, deleteyn, currState, deadlineDate, convert(varchar, regdate, 121) as regdate, regUserid, convert(varchar, lastUpdate, 121) as lastUpdate from db_cs.dbo.tbl_emergency_question_master "
	    sqlStr = sqlStr + addSql
		if (FRectOrdBy = "T") then
			sqlStr = sqlStr & " order by idx desc "
		elseif (FRectOrdBy = "U") then
			sqlStr = sqlStr & " order by (case when currState in ('1', '2', '4') then '1' else 'X' end), idx desc "
		else
			sqlStr = sqlStr & " order by (case when currState in ('3', '5') then '1'  when currState in ('1', '2', '4') then '2' else 'X' end), idx desc "
		end if
		selectQuery = sqlStr
		''response.write sqlStr

		Call EntityManager.LoadList(countQuery, selectQuery)
    End Sub

	Public Sub SetFItemListSize()
		redim preserve FItemList(FResultCount)
	End Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		Set FOneItem = New CEmergencyQuestionMasterItem

		FCurrPage         = 1
		FPageSize         = 20
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

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

class CEmergencyQuestionReply
	'
end class

Sub SelectBoxCsEmergencyQuestionCategoryGubun(gubunname, selectedgubun, showAll)
   dim tmp_str,query1
%>
<select class="select" name="<%= gubunname %>">
	<% if (showAll = "Y") then %><option value="" <% if (selectedgubun = "") then %>selected<% end if %>>전체</option><% end if %>
    <option value="1" <% if (selectedgubun = "1") then %>selected<% end if %>>상품</option>
	<option value="2" <% if (selectedgubun = "2") then %>selected<% end if %>>주문</option>
	<option value="3" <% if (selectedgubun = "3") then %>selected<% end if %>>배송</option>
	<option value="4" <% if (selectedgubun = "4") then %>selected<% end if %>>반품/취소</option>
	<option value="5" <% if (selectedgubun = "5") then %>selected<% end if %>>교환/변경</option>
	<option value="9" <% if (selectedgubun = "9") then %>selected<% end if %>>기타</option>
</select>
<%
End Sub

Sub RadioBoxCsEmergencyQuestionCategoryGubun(gubunname, selectedgubun, showAll)
   dim tmp_str,query1
%>
<% if (showAll = "Y") then %>
<input type="radio" name="<%= gubunname %>" value="" <% if (selectedgubun = "") then %>selected<% end if %> />전체
&nbsp;
<% end if %>
<input type="radio" name="<%= gubunname %>" value="1" <% if (selectedgubun = "1") then %>selected<% end if %>/> 상품
&nbsp;
<input type="radio" name="<%= gubunname %>" value="2" <% if (selectedgubun = "2") then %>selected<% end if %>/> 주문
&nbsp;
<input type="radio" name="<%= gubunname %>" value="3" <% if (selectedgubun = "3") then %>selected<% end if %>/> 배송
&nbsp;
<input type="radio" name="<%= gubunname %>" value="4" <% if (selectedgubun = "4") then %>selected<% end if %>/> 반품/취소
&nbsp;
<input type="radio" name="<%= gubunname %>" value="5" <% if (selectedgubun = "5") then %>selected<% end if %>/> 교환/변경
&nbsp;
<input type="radio" name="<%= gubunname %>" value="9" <% if (selectedgubun = "9") then %>selected<% end if %>/> 기타
<%
End Sub

Public Function CsEmergencyQuestionCategoryGubunToName(categoryGubun)
	select case categoryGubun
		case "1"
			CsEmergencyQuestionCategoryGubunToName = "상품"
		case "2"
			CsEmergencyQuestionCategoryGubunToName = "주문"
		case "3"
			CsEmergencyQuestionCategoryGubunToName = "배송"
		case "4"
			CsEmergencyQuestionCategoryGubunToName = "반품/취소"
		case "5"
			CsEmergencyQuestionCategoryGubunToName = "교환/변경"
		case "9"
			CsEmergencyQuestionCategoryGubunToName = "기타"
		case else
			CsEmergencyQuestionCategoryGubunToName = categoryGubun
	end select
End Function

Public Function CsEmergencyQuestionCurrStateToName(currState)
	select case currState
		case "1"
			CsEmergencyQuestionCurrStateToName = "미확인"
		case "2"
			CsEmergencyQuestionCurrStateToName = "답변대기"
		case "3"
			CsEmergencyQuestionCurrStateToName = "답변완료"
		case "4"
			CsEmergencyQuestionCurrStateToName = "재답변요청"
		case "5"
			CsEmergencyQuestionCurrStateToName = "재답변완료"
		case "9"
			CsEmergencyQuestionCurrStateToName = "완료처리"
		case else
			CsEmergencyQuestionCurrStateToName = currState
	end select
End Function

Public Function CsEmergencyQuestionCurrStateColor(currState)
	select case currState
		case "1"
			CsEmergencyQuestionCurrStateColor = "blue"
		case "2"
			CsEmergencyQuestionCurrStateColor = "blue"
		case "3"
			CsEmergencyQuestionCurrStateColor = "red"
		case "4"
			CsEmergencyQuestionCurrStateColor = "blue"
		case "5"
			CsEmergencyQuestionCurrStateColor = "red"
		case "9"
			CsEmergencyQuestionCurrStateColor = "#000000"
		case else
			CsEmergencyQuestionCurrStateColor = "#000000"
	end select
End Function

%>
