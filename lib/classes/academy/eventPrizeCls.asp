<%
Class CEventPrizeJoinItem
	public FeventGubun
	public Fevt_code
	public Fevt_name
	public Fevt_startdate
	public Fevt_enddate
	public Fevt_state
	public Fmaster_isusing
	public Fevt_prizedate
	public Fuserid
	public Fcomment
	public Fdetail_isusing
	public Fregdate
	public finvaliduserid

	public function GetEventGubunName()
		Select Case FeventGubun
			Case "designfingers"
				GetEventGubunName = "디자인핑거스"
			Case "culturestation"
				GetEventGubunName = "컬쳐스테이션"
			Case "tbl_event_etc"
				GetEventGubunName = "기타"
			Case Else
				GetEventGubunName = "일반"
		End Select
	end function

	public function GetIsUsingStr()
		if (Fmaster_isusing <> "Y") or (Fdetail_isusing <> "Y") then
			GetIsUsingStr = "<font color='red'>삭제</font>"
		else
			GetIsUsingStr = "정상"
		end if
	end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CEventPrize
	public FSUserid
	public FEPType
	public FEPStatus
	public FEKind
	public FEEventCode
	public FEEventName

	public FTotCnt
	public FCPage
	public FPSize
	public FGubun
	public FWinnerOX
	public FResultCount
	public FRectRegDate1
	public FRectRegDate2

    public FItemList()
    public FOneItem

    public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	''public FResultCount
	public FScrollCount
	public FPageCount

	public FRectEventGubun
	public FRectUserid
	public FRectUserName
	public FRectUserCell
	public FRectEventCode
	public FRectEventName
	public FRectStartDate
	public FRectEndDate
	public frectgubun
	public frectinvaliduseryn
	public FPrizeType
	public FStatus
	public FSongjangno
	public FStatusDesc
	'-----------------------------------------------------------------------
	' fnSetStatus : 이벤트 공통코드 가져오기
	'-----------------------------------------------------------------------
	public Function fnSetStatus
		FStatusDesc =""
	IF FPrizeType = 2 THEN
        FStatusDesc="쿠폰발급완료"
	ELSEIF FPrizeType = 3 THEN
        IF FStatus = 0 THEN
           FStatusDesc ="배송지입력대기"
        ELSEIF FStatus = 3 THEN
            IF  FSongjangno <> "" THEN
            	FStatusDesc="출고완료"
            ELSE
            	FStatusDesc="상품준비중"
            END IF
        END IF
    ELSEIF FPrizeType = 4 THEN
         IF FStatus = 0 THEN
         	FStatusDesc="티켓승인대기"
         ELSEIF FStatus = 3 THEN
         	FStatusDesc="티켓승인확정"
         END IF
    END IF
	End Function

	'//admin/eventmanage/event/eventjoin_list_new.asp
	public Sub GetEventJoinListNew
		dim sqlStr, addSqlStr, i
		dim tmpTable

		if (FRectStartdate <> "") then
			addSqlStr = addSqlStr + " 	and c.prtDate >= '" + CStr(FRectStartdate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " 	and c.prtDate < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectEventName <> "") then
			addSqlStr = addSqlStr + " 	and e.evtTitle like '%" + CStr(FRectEventName) + "%' "
		end if

		if (FRectEventCode <> "") then
			addSqlStr = addSqlStr + " 	and e.evtId = " + CStr(FRectEventCode) + " "
		end if

		if (FRectUserid <> "") then
			addSqlStr = addSqlStr + " 	and c.userid = '" + CStr(FRectUserid) + "' "
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_academy].[dbo].[tbl_eventInfo] e "
		sqlStr = sqlStr + " join [db_academy].[dbo].[tbl_eventSub] c "
		sqlStr = sqlStr + " on e.evtId = c.evtId "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr

		'response.write sqlStr &"<br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " e.evtId, e.evtTitle, e.evtSdate, e.evtEdate, e.evt_state, e.evt_using as master_isusing, e.prizeDate, c.userid, c.evtcom_txt as comment, c.evtcom_using as detail_isusing, c.prtDate as regdate"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_academy].[dbo].[tbl_eventInfo] e "
		sqlStr = sqlStr + " 	join [db_academy].[dbo].[tbl_eventSub] c "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		e.evtId = c.evtId "
		sqlStr = sqlStr + addSqlStr
		sqlStr = sqlStr + " order by c.regdate desc "

		''response.write sqlStr &"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new CEventPrizeJoinItem

				''eventGubun, evt_code, evt_name, evt_startdate, evt_enddate, evt_state, master_isusing, evt_prizedate, userid, comment, detail_isusing, regdate
				FItemList(i).finvaliduserid = rsACADEMYget("invaliduserid")
				FItemList(i).FeventGubun 		= rsACADEMYget("eventGubun")
				FItemList(i).Fevt_code 			= rsACADEMYget("evt_code")
				FItemList(i).Fevt_name 			= rsACADEMYget("evt_name")
				FItemList(i).Fevt_startdate 	= rsACADEMYget("evt_startdate")
				FItemList(i).Fevt_enddate 		= rsACADEMYget("evt_enddate")
				FItemList(i).Fevt_state 		= rsACADEMYget("evt_state")
				FItemList(i).Fmaster_isusing 	= rsACADEMYget("master_isusing")
				FItemList(i).Fevt_prizedate 	= rsACADEMYget("prizeDate")
				FItemList(i).Fuserid 			= rsACADEMYget("userid")
				FItemList(i).Fcomment 			= rsACADEMYget("comment")
				FItemList(i).Fdetail_isusing 	= rsACADEMYget("detail_isusing")
				FItemList(i).Fregdate 			= rsACADEMYget("regdate")

				if Left(FItemList(i).Fevt_startdate, 10) = "1900-01-01" or Left(FItemList(i).Fevt_startdate, 10) = "2001-10-10" then
					FItemList(i).Fevt_startdate = ""
				end if

				if Left(FItemList(i).Fevt_enddate, 10) = "1900-01-01" or Left(FItemList(i).Fevt_enddate, 10) = "2001-10-10" then
					FItemList(i).Fevt_enddate = ""
				end if

				if Left(FItemList(i).Fevt_prizedate, 10) = "1900-01-01" or Left(FItemList(i).Fevt_prizedate, 10) = "2001-10-10" then
					FItemList(i).Fevt_prizedate = ""
				end if

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
    end Sub

    Private Sub Class_Initialize()
		ReDim FItemList(0)

		FCurrPage		= 1
		FPageSize 		= 20
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
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
End Class
%>
