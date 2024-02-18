<%
Class CLecWaitUserItem
	public Fidx
	public Flec_idx
	public FlecOption
	public Fuserid
	public Fuser_name
	public Fuser_phone
	public Fuser_email
	public Fregcount
	public Fcurrstate
	public Fisusing
	public Fregdate

	public FRegRank
	public FRegEndDay

	public Flec_title
	public FLec_period
	public FLec_Date
	public FLec_Cost
	public Fmat_cost

	public FWait_count

	public FoptName
	public FoptRegSDate
	public FoptRegEDate
	public FoptLecSDate
	public FoptLecEDate
	public FoptLimitCnt
	public FoptLimitSold
	public FoptWaitCnt

	public FSmallImg


	public function getStateName()
		if Fcurrstate=0 then
			getStateName = "대기신청중"
		elseif Fcurrstate=3 then
			if (IsSettleExpired) then
				getStateName = "기간만료"
			else
				getStateName = "결제대기"
			end if
		elseif Fcurrstate=7 then
			getStateName = "결제완료"
		end if
	end function

	public function getStateNameColor()
		if Fcurrstate=0 then
			getStateNameColor = "#000000"
		elseif Fcurrstate=3 then
			if (IsSettleExpired) then
				getStateNameColor = "#CC3333"
			else
				getStateNameColor = "#33CC33"
			end if
		elseif Fcurrstate=7 then
			getStateNameColor = "#3333CC"
		end if
	end function

	public function IsSettleExpired()
		if IsNULL(FRegEndDay) then
			IsSettleExpired = true
		else
			IsSettleExpired = (now()>CDate(FRegEndDay))
		end if
	end function

	Private Sub Class_Initialize()
        '
    End Sub

    Private Sub Class_Terminate()
        '
    End Sub
end Class


Class CLecWaitUser
	public FOneItem
    public FItemList()

	public FTotalCount
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount

    public FRectIdx
	public FRectUserID
	public FRectLecIdx
	public FRectLecOpt

	public FRectOnlyusing
	public FRectYYYYMM
	public FRectOnlyNotStart

	public Sub getWaitingList()
		dim sqlStr, addSql, i

		addSql = ""
		if FRectLecIdx<>"" then
			addSql = addSql + " and w.lec_idx=" + CStr(FRectLecIdx)
		else
			addSql = addSql + " and L.lec_date='" + CStr(FRectYYYYMM) + "'"
		end if

		if FRectOnlyusing="on" then
			addSql = addSql + " and w.isusing='Y'"
		end if

		if FRectOnlyNotStart="on" then
			addSql = addSql + " and L.lec_startday1>getdate() "
		end if

		if FRectLecOpt<>"" then
			addSql = addSql + " and w.lecOption='" & FRectLecOpt & "'"
		end if

		'//카운트
		sqlStr = "select count(w.idx) as cnt from [db_academy].[dbo].tbl_lec_waiting_user w,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_lec_item L"
		sqlStr = sqlStr + " where w.lec_idx=L.idx " & addSql

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		'// 목록 접수
		sqlStr	=	"select top " + CStr(FPageSize*FCurrPage) + " w.idx, w.lec_idx,w.lecOption,W.user_phone,"
		sqlStr = sqlStr + " w.userid, w.user_name, w.user_phone, w.user_email, w.regcount,w.currstate,"
		sqlStr = sqlStr + " w.isusing, convert(varchar(10),w.regdate,21) as regdate, w.regrank, w.regendday,"
		sqlStr = sqlStr + " L.storyimg,L.lec_title,L.lec_period,L.lec_cost,"
		sqlStr = sqlStr + " L.mat_cost,L.lec_startday1, L.smallimg "
		sqlStr = sqlStr + " from db_academy.dbo.tbl_lec_waiting_user W "
		sqlStr = sqlStr + " ,[db_academy].[dbo].tbl_lec_item L"
		sqlStr = sqlStr + " where W.lec_idx=L.idx " & addSql
		sqlStr = sqlStr + " order by  L.idx desc, w.regrank"

		'response.write sqlStr
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
        if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
            FtotalPage = FtotalPage +1
        end if
        FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        redim preserve FItemList(FResultCount)

        if  not rsACADEMYget.EOF  then
            i = 0
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.eof
                set FItemList(i) = new CLecWaitUserItem
                FItemList(i).Fidx        = rsACADEMYget("idx")
				FItemList(i).Flec_idx    = rsACADEMYget("lec_idx")
				FItemList(i).FlecOption  = rsACADEMYget("lecOption")
				
				FItemList(i).Fuserid     = rsACADEMYget("userid")
				FItemList(i).Fuser_name  = db2html(rsACADEMYget("user_name"))
				FItemList(i).Fuser_phone = rsACADEMYget("user_phone")
				FItemList(i).Fuser_email = db2html(rsACADEMYget("user_email"))
				FItemList(i).Fregcount   = rsACADEMYget("regcount")
				FItemList(i).Fcurrstate     = rsACADEMYget("currstate")
				FItemList(i).Fisusing    = rsACADEMYget("isusing")
				FItemList(i).Fregdate    = rsACADEMYget("regdate")

				FItemList(i).Flec_title	 = db2html(rsACADEMYget("lec_title"))
				FItemList(i).FLec_Date	= rsACADEMYget("lec_startday1")

				FItemList(i).FRegRank = rsACADEMYget("regrank")
				FItemList(i).FRegEndDay = rsACADEMYget("regendday")

				FItemList(i).FSmallImg	= imgFingers & "/lectureitem/small/" & GetImageSubFolderByItemid(FItemList(i).Flec_idx) & "/" & rsACADEMYget("smallimg")

                rsACADEMYget.MoveNext
                i = i + 1
            loop
        end if
		rsACADEMYget.Close
	end sub

	public Sub GetMyWaitList()
		dim sqlStr, i

		''대기순번처리 검토

		sqlStr = "select count(idx) as cnt from [db_academy].[dbo].tbl_lec_waiting_user"
		sqlStr = sqlStr + " where userid='" + FRectUserID + "'"
		sqlStr = sqlStr + " and isusing='Y'"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sqlStr	=	"select top " + CStr(FPageSize*FCurrPage) + " w.idx, w.lec_idx,W.user_phone,"
		sqlStr = sqlStr + " w.userid, w.user_name, w.user_phone, w.user_email, w.regcount,w.currstate,w.isusing,w.regdate,w.regendday,"
		sqlStr = sqlStr + " L.storyimg,L.lec_title,L.lec_period,L.lec_cost,L.mat_cost,L.lec_startday1 "
		sqlStr = sqlStr + " from db_academy.dbo.tbl_lec_waiting_user W "
		sqlStr = sqlStr + " left join db_academy.dbo.tbl_lec_item L on W.lec_idx=L.idx "
		sqlStr = sqlStr + " where w.userid='" + FRectUserID + "'"
		sqlStr = sqlStr + " and w.isusing='Y'"
		sqlStr = sqlStr + " order by w.idx desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
        if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
            FtotalPage = FtotalPage +1
        end if
        FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        redim preserve FItemList(FResultCount)

        if  not rsACADEMYget.EOF  then
            i = 0
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.eof
                set FItemList(i) = new CLecWaitUserItem
                FItemList(i).Fidx        = rsACADEMYget("idx")
				FItemList(i).Flec_idx    = rsACADEMYget("lec_idx")
				FItemList(i).Fuserid     = rsACADEMYget("userid")
				FItemList(i).Fuser_name  = db2html(rsACADEMYget("user_name"))
				FItemList(i).Fuser_phone = rsACADEMYget("user_phone")
				FItemList(i).Fuser_email = db2html(rsACADEMYget("user_email"))
				FItemList(i).Fregcount   = rsACADEMYget("regcount")
				FItemList(i).Fcurrstate     = rsACADEMYget("currstate")
				FItemList(i).Fisusing    = rsACADEMYget("isusing")
				FItemList(i).Fregdate    = rsACADEMYget("regdate")

				FItemList(i).Flec_title	 = db2html(rsACADEMYget("lec_title"))
				FItemList(i).FLec_Date	= rsACADEMYget("lec_startday1")

				FItemList(i).FRegEndDay = rsACADEMYget("regendday")

                rsACADEMYget.MoveNext
                i = i + 1
            loop
        end if
		rsACADEMYget.Close
	end sub

	public sub GetOneWaitUser
		dim sqlStr
		sqlStr = "select top 1 w.idx, w.lec_idx,W.user_phone,"
		sqlStr = sqlStr + " w.userid, w.user_name, w.user_phone, w.user_email, w.regcount,w.currstate,w.isusing,w.regrank,w.regdate, convert(varchar(19),w.regendday,21) as regendday,"
		sqlStr = sqlStr + " L.lec_period,L.lec_title,L.lec_period,L.lec_cost,L.mat_cost,L.lec_startday1, L.lec_cost, L.smallimg, L.mat_cost, "
		sqlStr = sqlStr + " L.wait_count, O.LecOptionName, O.RegStartDate, O.RegEndDate, O.LecStartDate, O.LecEndDate,"
		sqlStr = sqlStr + " O.limit_count as optLimitCnt, O.limit_sold as optLimitSold, O.wait_count as optWaitCnt"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_lec_waiting_user w"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_item_option O on W.lec_idx=O.lecidx "
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_item L on W.lec_idx=L.idx "
		sqlStr = sqlStr + " where w.idx=" + CStr(FRectIdx)
		if FRectUserID<>"" then
			sqlStr = sqlStr + " and w.userid='" + FRectUserID + "'"
		end if

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount

		if Not rsACADEMYget.Eof then
			set FOneItem = new CLecWaitUserItem

			FOneItem.Fidx        = rsACADEMYget("idx")
			FOneItem.Flec_idx    = rsACADEMYget("lec_idx")
			FOneItem.Fuserid     = rsACADEMYget("userid")
			FOneItem.Fuser_name  = db2html(rsACADEMYget("user_name"))
			FOneItem.Fuser_phone = rsACADEMYget("user_phone")
			FOneItem.Fuser_email = db2html(rsACADEMYget("user_email"))
			FOneItem.Fregcount   = rsACADEMYget("regcount")
			FOneItem.Fcurrstate     = rsACADEMYget("currstate")
			FOneItem.Fisusing    = rsACADEMYget("isusing")
			FOneItem.Fregdate    = rsACADEMYget("regdate")

			FOneItem.FRegRank	= rsACADEMYget("regrank")
			FOneItem.FRegEndDay = rsACADEMYget("regendday")

			FOneItem.Flec_title	= rsACADEMYget("lec_title")
			FOneItem.FLec_Date	= rsACADEMYget("lec_startday1")
			FOneItem.FLec_period= rsACADEMYget("lec_period")
			FOneItem.FLec_Cost	= rsACADEMYget("lec_cost")
			FOneItem.Fmat_cost	= rsACADEMYget("mat_cost")
			FOneItem.FWait_count= rsACADEMYget("wait_count")

			FOneItem.FSmallImg	= imgFingers & "/lectureitem/small/" & GetImageSubFolderByItemid(FOneItem.Flec_idx) & "/" & rsACADEMYget("smallimg")
			
			'옵션 관련내용
			FOneItem.FoptName		= rsACADEMYget("LecOptionName")
			FOneItem.FoptRegSDate	= rsACADEMYget("RegStartDate")
			FOneItem.FoptRegEDate	= rsACADEMYget("RegEndDate")
			FOneItem.FoptLecSDate	= rsACADEMYget("LecStartDate")
			FOneItem.FoptLecEDate	= rsACADEMYget("LecEndDate")
			FOneItem.FoptLimitCnt	= rsACADEMYget("optLimitCnt")
			FOneItem.FoptLimitSold	= rsACADEMYget("optLimitSold")
			FOneItem.FoptWaitCnt	= rsACADEMYget("optWaitCnt")

		end if
		rsACADEMYget.close
	end sub

	public sub GetOneValidWaitingUser
		dim sqlStr
		sqlStr = "select top 1 * from [db_academy].[dbo].tbl_lec_waiting_user"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " and currstate=3"

		if FRectUserID<>"" then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount

		if Not rsACADEMYget.Eof then
			set FOneItem = new CLecWaitUserItem
			FOneItem.Fidx        = rsACADEMYget("idx")
			FOneItem.Flec_idx    = rsACADEMYget("lec_idx")
			FOneItem.Fuserid     = rsACADEMYget("userid")
			FOneItem.Fuser_name  = db2html(rsACADEMYget("user_name"))
			FOneItem.Fuser_phone = rsACADEMYget("user_phone")
			FOneItem.Fuser_email = db2html(rsACADEMYget("user_email"))
			FOneItem.Fregcount   = rsACADEMYget("regcount")
			FOneItem.Fcurrstate     = rsACADEMYget("currstate")
			FOneItem.Fisusing    = rsACADEMYget("isusing")
			FOneItem.Fregdate    = rsACADEMYget("regdate")

		end if
		rsACADEMYget.close
	end sub

	Private Sub Class_Initialize()
        redim  FItemList(0)
        FCurrPage =1
        FPageSize = 100
        FResultCount = 0
        FScrollCount = 10
        FTotalCount =0
    End Sub

    Private Sub Class_Terminate()
        '
    End Sub

    public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

%>