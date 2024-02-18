<%
Class MileageExtinctionObj
'태스크 멤버변수
    public task_id
    public task_jukyo
    public task_jukyocd
    public task_startdate
    public task_enddate
    public task_chkDays
    public task_useyn
    public task_taskStatus
    public task_regdate
    public task_lastupdate
    public task_regUser
    public task_updateUser

'로그
    public log_id
    public log_taskKey
    public log_jukyo
    public log_jukyocd
    public log_doneDate
    public log_chkPopulation
    public log_updatedUsersCnt

'기타
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class MileageExtinctionCls
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FUoption
	public FUkeyword	
	public FUsdt
	public FUedt
	
	Public FRectSubIdx	

    public Sub getTaskList()
        dim sqlStr, i, sqlWhere

		sqlwhere = ""
		if FUoption <> "" and FUkeyword <> "" then
			sqlWhere = sqlWhere + " and "& FUoption  &" like '%" & FUkeyword & "%'"
		end if
		if FUsdt <> "" then
			sqlWhere = sqlWhere + " and startdate >= '" & FUsdt & "'"
		end if
		if FUedt <> "" then
			sqlWhere = sqlWhere + " and enddate >= '" & FUedt & "'"
		end if

		sqlStr = " select count(1) as cnt from DB_USER.DBO.tbl_mileage_auto_extinction_master with(nolock) "
		sqlStr = sqlStr + " where 1=1 "
        sqlStr = sqlStr + "     and useyn =1 "
		sqlStr = sqlStr + sqlwhere

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		sqlStr = sqlStr + "  * "
        sqlStr = sqlStr + " FROM DB_USER.DBO.tbl_mileage_auto_extinction_master with(nolock) "
        sqlStr = sqlStr + " where 1=1 "
        sqlStr = sqlStr + "     and useyn =1 "
		sqlStr = sqlStr + sqlWhere

		sqlStr = sqlStr + " order by  id desc"

		'response.write sqlStr &"<br>"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new MileageExtinctionObj

                FItemList(i).task_id = rsget("id")
                FItemList(i).task_jukyo = rsget("jukyo")
                FItemList(i).task_jukyocd = rsget("jukyocd")
                FItemList(i).task_startdate = rsget("startdate")
                FItemList(i).task_enddate = rsget("enddate")
                FItemList(i).task_chkDays = rsget("chk_days")
                FItemList(i).task_useyn = rsget("useyn")
                FItemList(i).task_taskStatus = rsget("task_status")
                FItemList(i).task_regdate = rsget("regdate")
                FItemList(i).task_lastupdate = rsget("lastupdate")
                FItemList(i).task_regUser = rsget("reg_user")
                FItemList(i).task_updateUser   = rsget("update_user")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub getTaskLogList()
        dim sqlStr, i, sqlWhere

		sqlwhere = ""
		if FUoption <> "" and FUkeyword <> "" then
			sqlWhere = sqlWhere + " and "& FUoption  &" like '%" & FUkeyword & "%'"
		end if
		if FUsdt <> "" then
			sqlWhere = sqlWhere + " and done_date >= '" & FUsdt & "'"
		end if
		if FUedt <> "" then
			sqlWhere = sqlWhere + " and done_date <= '" & FUedt & "'"
		end if

		sqlStr = " SELECT count(1) as cnt FROM DB_USER.DBO.tbl_mileage_auto_extinction_log with(nolock) "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

		sqlStr = " select top " + CStr(FPageSize * FCurrPage) + " "
		sqlStr = sqlStr + "  *														"
		sqlStr = sqlStr + "  FROM DB_USER.DBO.tbl_mileage_auto_extinction_log with(nolock)	"
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere
		sqlStr = sqlStr + " order by id desc"

'		response.write sqlStr &"<br>"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new MileageExtinctionObj

                FItemList(i).log_id = rsget("id")
                FItemList(i).log_taskKey = rsget("task_key")
                FItemList(i).log_jukyo = rsget("jukyo")
                FItemList(i).log_jukyocd = rsget("jukyocd")
                FItemList(i).log_doneDate = rsget("done_date")
                FItemList(i).log_chkPopulation = rsget("chk_population")
                FItemList(i).log_updatedUsersCnt = rsget("updated_users_cnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = " SELECT top 1 *  "
        sqlStr = sqlStr & " FROM DB_USER.DBO.tbl_mileage_auto_extinction_master with(nolock) "
        SqlStr = SqlStr & " where id=" + CStr(FRectSubIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new MileageExtinctionObj

        if Not rsget.Eof then
            
            FOneItem.task_id = rsget("id")
            FOneItem.task_jukyo = rsget("jukyo")
            FOneItem.task_jukyocd = rsget("jukyocd")
            FOneItem.task_startdate = rsget("startdate")
            FOneItem.task_enddate = rsget("enddate")
            FOneItem.task_chkDays = rsget("chk_days")
            FOneItem.task_useyn = rsget("useyn")
            FOneItem.task_taskStatus = rsget("task_status")
            FOneItem.task_regdate = rsget("regdate")
            FOneItem.task_lastupdate = rsget("lastupdate")
            FOneItem.task_regUser = rsget("reg_user")
            FOneItem.task_updateUser   = rsget("update_user")

        end if
        rsget.close
	End Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
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
end Class
%>