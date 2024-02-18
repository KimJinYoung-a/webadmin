<%

'휴가

Class CTenByTenVacationMasterItem
	'tbl_vacation_master
	'idx, userid, divcd, startday, endday, totalvacationday, usedvacationday, requestedday, deleteyn, registerid, regdate

	public Fidx
	public Fuserid
	public Fempno
	public Fdivcd			'1:연차/2:월차/3:포상/4:위로/5:장기근속/6:경조사/7:휴일대체
	public Fstartday
	public Fendday
	public Ftotalvacationday
	public Fusedvacationday
	public Frequestedday	'승인대기일수
	public Fdeleteyn
	public Fregisterid
	public Fregdate

	public Fusername
	public Fpart_name
	public Fposit_name
	public Fjob_name

	public Fposit_sn

	public Fjoinday
	public Frealjoinday
	public Fretireday

	public FpromotionDay
	public FjungsanDay
	public FretireJungsanDay

	public FdepartmentNameFull
	public Fcomment

	public function GetDivCDStr()
		if (Fdivcd = "1") then
			GetDivCDStr = "연차"
		elseif (Fdivcd = "2") then
			GetDivCDStr = "월차"
		elseif (Fdivcd = "3") then
			GetDivCDStr = "포상"
		elseif (Fdivcd = "4") then
			GetDivCDStr = "위로"
		elseif (Fdivcd = "6") then
			GetDivCDStr = "경조사"
		elseif (Fdivcd = "7") then
			GetDivCDStr = "휴일대체"
		elseif (Fdivcd = "5") then
			GetDivCDStr = "장기"
		elseif (Fdivcd = "8") then
			GetDivCDStr = "기타"
		elseif (Fdivcd = "9") then
			GetDivCDStr = "보상"
		elseif (Fdivcd = "A") then
			GetDivCDStr = "생일"
		else
			GetDivCDStr = "===="
		end if
	end function

    public function IsExpiredVacation()
        IsExpiredVacation = ( now()> CDate(Fendday))
    end function

	public function GetRemainVacationDay()
        GetRemainVacationDay = (Ftotalvacationday - (Fusedvacationday + Frequestedday + FpromotionDay + FjungsanDay + FretireJungsanDay))
    end function

	public function IsAvailableVacation()
		dim today

		today = date()

		if (Fendday >= today) then
			if (Fdeleteyn = "Y") then
				IsAvailableVacation = "D"
			else
				'// 잔여일수가 모두 소진되었다면 사용여부 N (2011.03.18;허진원)
				'// 승인대기 일수 있는 경우 유효한 휴가(skyer9 2013-08-08) 
				'//정산일수 추가 2014.08.05 정윤정
				if (Ftotalvacationday - Fusedvacationday-FjungsanDay)<=0 then
					IsAvailableVacation = "N"
				else
					IsAvailableVacation = "Y"
				end if
			end if
		else
			IsAvailableVacation = "N"
		end if

	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CTenByTenVacationDetailItem
	'tbl_vacation_detail
	'idx, masteridx, startday, endday, totalday, approverid, approveday, statedivcd, deleteyn, registerid, regdate

	public Fidx
	public Fmasteridx
	public FmasterDivCD
	public Fstartday
	public Fendday
	public Ftotalday
	public Fapproverid
	public Fapproveday
	public Fstatedivcd	'R:신청/A:승인/D:거부
	public Fdeleteyn
	public Fregisterid
	public Fregdate
	public Freportidx
	public Freportstate
	public Fhalfgubun
	public Fapproverempno
	public Fregisterempno

	public Fapprovername
	public Fregistername
	public Fcomment

	public function GetStateDivCDStr()
		if (Fstatedivcd = "R") then
			GetStateDivCDStr = "신청"
		elseif (Fstatedivcd = "A") then
			GetStateDivCDStr = "승인"
		elseif (Fstatedivcd = "D") then
			GetStateDivCDStr = "거부"
		else
			GetStateDivCDStr = "===="
		end if
	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class


Class CTenByTenVacationCalendarItem
	public Fsolar_date
	public Fmasteridx
	public Fpart_sn
	public Fuserid
	public Fusername
	public Fpart_name
	public Fstartday
	public Fendday
	public Fstatedivcd	'R:신청/A:승인/D:거부
	public Ftotalday
	public Fhalfgubun
	public Fholiday
	public Fholiday_name
	public FworkAgent
	public FcallNum


	public function GetStateDivCDStr()
		if (Fstatedivcd = "R") then
			GetStateDivCDStr = "신청"
		elseif (Fstatedivcd = "A") then
			GetStateDivCDStr = "승인"
		elseif (Fstatedivcd = "D") then
			GetStateDivCDStr = "거부"
		else
			GetStateDivCDStr = "===="
		end if
	end function

	public function GetDay()
		GetDay = Day(Fsolar_date)
	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class



Class CTenByTenVacation
	public FItemList()
	public FCalendarItemList()
	public FItemOne

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectUserId
	public FRectEmpNO
	public FRectMasterIdx
	public FRectdetailIdx
	public FRectSearchKey
	public FRectSearchString
	public FRectIsDelete
	public FRectPart_sn
	public FRectposit_sn
	public FRectNeedApprove
	public FRectDivCd
	public FRectStateDiv
	public FRectShowOnlyAvail

	public FRectYYYY
	public FRectMM
	public FRectStartDate
	public FRectEndDate

	public Fidx
	public Fmasteridx
	public Fstartday
	public Fendday
	public Ftotalday
	public statedivcd
	public Fregdate
	public Fempno
	public Fuserid
	public Fdivcd
	public Ftotstartday
	public Ftotendday
	public Ftotalvacationday
	public Fusedvacationday
	public Frequestedday
	public FdivcdStr
	public Fhalfgubun
	public FworkAgent
	public FcallNum
	public Fdepartment_id
	public Finc_subdepartment

	Private Sub Class_Initialize()
		redim  FitemList(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Sub GetMasterList()
		dim sql, i

		'// 개수 파악 //
		sql = "select count(vm.idx), Ceiling(Cast(Count(vm.idx) as float)/" & FPageSize & ") " & vbCrlf
		sql = sql & "from [db_partner].[dbo].tbl_vacation_master vm, [db_partner].[dbo].tbl_user_tenbyten t " & vbCrlf
		sql = sql & "left join [db_partner].dbo.tbl_partInfo as pa on t.part_sn = pa.part_sn " & vbCrlf
		sql = sql & "left join [db_partner].dbo.tbl_positInfo as po on t.posit_sn = po.posit_sn " & vbCrlf
		sql = sql & "left join [db_partner].dbo.tbl_JobInfo as jo on t.job_sn = jo.job_sn " & vbCrlf
		sql = sql & "left join db_partner.dbo.vw_user_department dv " & vbCrLf
		sql = sql & "on " & vbCrLf
		sql = sql & "	1 = 1 " & vbCrLf
		sql = sql & "	and t.department_id = dv.cid " & vbCrLf
		sql = sql & "	and dv.useYN = 'Y' " & vbCrLf
		sql = sql & "where  vm.empno = t.empno " & vbCrlf
		sql = sql & "and t.isusing =1 " & vbCrlf

		if (FRectStateDiv <> "") then
			sql = sql & "and t.statediv = '" & FRectStateDiv & "' " & vbCrlf
		end if

		if (FRectposit_sn <> "") then
			if (FRectposit_sn = "99") then
				sql = sql & "and t.posit_sn <= '11' " & vbCrlf
			else
				sql = sql & "and t.posit_sn = '" & FRectposit_sn & "' " & vbCrlf
			end if
		end if

		if (FRectIsDelete <> "") then
			sql = sql & "and vm.deleteyn = '" & FRectIsDelete & "' " & vbCrlf
		end if

		if (FRectShowOnlyAvail = "Y") then
			sql = sql & " and vm.endday >= getdate() " & vbCrlf
			sql = sql & " and vm.deleteyn <> 'Y' " & vbCrlf
			sql = sql & " and (vm.totalvacationday - vm.usedvacationday) > 0 " & vbCrlf
		end if

		if ((FRectPart_sn <> "") and (FRectPart_sn <> "1")) then
			sql = sql & "and t.part_sn = " & CStr(FRectPart_sn) & " " & vbCrlf
		end if

		if (FRectNeedApprove <> "") then
			sql = sql & "and vm.requestedday > 0 " & vbCrlf
			sql = sql & "and vm.endday >= getdate() " & vbCrlf
			sql = sql & "and vm.deleteyn <> 'Y' " & vbCrlf
		end if

		'// 검색어 쿼리 //
		if FRectSearchKey<>"" and FRectSearchString<>"" then
			sql = sql & " and " & FRectSearchKey & " like '%" & FRectSearchString & "%' "
		end if

		if (FRectDivCd <> "") then
			sql = sql & "and vm.divcd = '" & CStr(FRectDivCd) & "' " & vbCrlf
		end if

		if (Fdepartment_id <> "") then
			if (Finc_subdepartment = "N") then
				sql = sql & " AND t.department_id = '" & Fdepartment_id & "' "
			else
				sql = sql & " AND (IsNull(dv.cid1, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid2, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid3, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid4, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid5, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid6, -1) = '" & Fdepartment_id & "') "
			end if
		end if

		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close



		'// 목록 //
		sql = "select top " & CStr(FPageSize*FCurrPage) & " " & vbCrlf
		sql = sql & "	vm.idx, vm.userid, vm.divcd, vm.startday, vm.endday, vm.totalvacationday, vm.usedvacationday, vm.requestedday, vm.deleteyn, vm.registerid, vm.regdate " & vbCrlf
		sql = sql & "	, t.username " & vbCrlf
		sql = sql & "	, pa.part_name " & vbCrlf
		sql = sql & "	, po.posit_name " & vbCrlf
		sql = sql & "	, jo.job_name " & vbCrlf
		sql = sql & "	, t.joinday, t.realjoinday, t.retireday, t.empno, t.posit_sn " & vbCrlf
		sql = sql & "	, vm.promotionDay, vm.jungsanDay, vm.retireJungsanDay, isNull(dv.departmentNameFull,'') AS departmentNameFull " & vbCrlf
		sql = sql & "from [db_partner].[dbo].tbl_vacation_master vm, [db_partner].[dbo].tbl_user_tenbyten t " & vbCrlf
		sql = sql & "left join [db_partner].dbo.tbl_partInfo as pa on t.part_sn = pa.part_sn " & vbCrlf
		sql = sql & "left join [db_partner].dbo.tbl_positInfo as po on t.posit_sn = po.posit_sn " & vbCrlf
		sql = sql & "left join [db_partner].dbo.tbl_JobInfo as jo on t.job_sn = jo.job_sn " & vbCrlf
		sql = sql & "left join db_partner.dbo.vw_user_department dv " & vbCrLf
		sql = sql & "on " & vbCrLf
		sql = sql & "	1 = 1 " & vbCrLf
		sql = sql & "	and t.department_id = dv.cid " & vbCrLf
		sql = sql & "	and dv.useYN = 'Y' " & vbCrLf
		sql = sql & "where vm.empno = t.empno " & vbCrlf
		sql = sql & "and t.isusing =1 " & vbCrlf

		if (FRectStateDiv <> "") then
			sql = sql & "and t.statediv = '" & FRectStateDiv & "' " & vbCrlf
		end if

		if (FRectposit_sn <> "") then
			if (FRectposit_sn = "99") then
				sql = sql & "and t.posit_sn <= '11' " & vbCrlf
			else
				sql = sql & "and t.posit_sn = '" & FRectposit_sn & "' " & vbCrlf
			end if
		end if

		if (FRectIsDelete <> "") then
			sql = sql & "and vm.deleteyn = '" & FRectIsDelete & "' " & vbCrlf
		end if

		if (FRectIsDelete <> "") then
			sql = sql & "and vm.deleteyn = '" & FRectIsDelete & "' " & vbCrlf
		end if

		if (FRectShowOnlyAvail = "Y") then
			sql = sql & " and vm.endday >= getdate() " & vbCrlf
			sql = sql & " and vm.deleteyn <> 'Y' " & vbCrlf
			sql = sql & " and (vm.totalvacationday - vm.usedvacationday) > 0 " & vbCrlf
		end if

		if ((FRectPart_sn <> "") and (FRectPart_sn <> "1")) then
			sql = sql & "and t.part_sn = " & CStr(FRectPart_sn) & " " & vbCrlf
		end if

		if (FRectNeedApprove <> "") then
			sql = sql & "and vm.requestedday > 0 " & vbCrlf
			sql = sql & "and vm.endday >= getdate() " & vbCrlf
			sql = sql & "and vm.deleteyn <> 'Y' " & vbCrlf
		end if

		'// 검색어 쿼리 //
		if FRectSearchKey<>"" and FRectSearchString<>"" then
			sql = sql & " and " & FRectSearchKey & " like '%" & FRectSearchString & "%' "
		end if

		if (FRectDivCd <> "") then
			sql = sql & "and vm.divcd = '" & CStr(FRectDivCd) & "' " & vbCrlf
		end if

		if (Fdepartment_id <> "") then
			if (Finc_subdepartment = "N") then
				sql = sql & " AND t.department_id = '" & Fdepartment_id & "' "
			else
				sql = sql & " AND (IsNull(dv.cid1, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid2, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid3, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid4, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid5, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid6, -1) = '" & Fdepartment_id & "') "
			end if
		end if

		sql = sql & "order by vm.idx desc " & vbCrlf

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly
		'response.write sql

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)



		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTenByTenVacationMasterItem

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fuserid		= rsget("userid")
				FItemList(i).Fdivcd			= rsget("divcd")
				FItemList(i).Fstartday		= rsget("startday")
				FItemList(i).Fendday		= rsget("endday")
				FItemList(i).Ftotalvacationday		= rsget("totalvacationday")
				FItemList(i).Fusedvacationday		= rsget("usedvacationday")
				FItemList(i).Frequestedday	= rsget("requestedday")
				FItemList(i).Fdeleteyn		= rsget("deleteyn")
				FItemList(i).Fregisterid	= rsget("registerid")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Fusername		= rsget("username")
				FItemList(i).Fpart_name		= rsget("part_name")
				FItemList(i).Fposit_name	= rsget("posit_name")
				FItemList(i).Fjob_name		= rsget("job_name")
				FItemList(i).Fjoinday		= rsget("joinday")
				FItemList(i).Frealjoinday	= rsget("realjoinday")
				FItemList(i).Fretireday		= rsget("retireday")
				FItemList(i).Fempno			= rsget("empno")
				FItemList(i).Fposit_sn		= rsget("posit_sn")

				FItemList(i).FpromotionDay		= rsget("promotionDay")
				FItemList(i).FjungsanDay		= rsget("jungsanDay")
				FItemList(i).FretireJungsanDay	= rsget("retireJungsanDay")

				FItemList(i).FdepartmentNameFull		= rsget("departmentNameFull")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	public Function fnGetPartList
		dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_VacationMonth_GetList('"&Fempno&"')"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetPartList = rsget.getRows()
			END IF
			rsget.close
	End Function

	public Sub GetMasterOne()
		dim sql, i

		if FRectMasterIdx="" or isnull(FRectMasterIdx) then exit Sub

		sql = "select top 1 " & vbCrlf
		sql = sql & "	vm.idx, vm.userid, vm.empno, vm.divcd, vm.startday, vm.endday, vm.totalvacationday, vm.usedvacationday, vm.requestedday, vm.deleteyn, vm.registerid, vm.regdate " & vbCrlf
		sql = sql & "	, t.username " & vbCrlf
		sql = sql & "	, pa.part_name " & vbCrlf
		sql = sql & "	, po.posit_name " & vbCrlf
		sql = sql & "	, jo.job_name " & vbCrlf
		sql = sql & "	, t.posit_sn " & vbCrlf
		sql = sql & "	, vm.promotionDay, vm.jungsanDay, vm.retireJungsanDay, isNull(vm.comment,'') as comment " & vbCrlf
		sql = sql & "from [db_partner].[dbo].tbl_vacation_master vm, [db_partner].[dbo].tbl_user_tenbyten t " & vbCrlf
		sql = sql & "left join [db_partner].dbo.tbl_partInfo as pa on t.part_sn = pa.part_sn " & vbCrlf
		sql = sql & "left join [db_partner].dbo.tbl_positInfo as po on t.posit_sn = po.posit_sn " & vbCrlf
		sql = sql & "left join [db_partner].dbo.tbl_JobInfo as jo on t.job_sn = jo.job_sn " & vbCrlf
		sql = sql & "where vm.empno = t.empno " & vbCrlf


		if ((FRectPart_sn <> "") and (FRectPart_sn <> "1")) then
			sql = sql & " and t.part_sn = " & CStr(FRectPart_sn) & " " & vbCrlf
		end if

		'// 검색어 쿼리 //
		if FRectSearchKey<>"" and FRectSearchString<>"" then
			sql = sql & " and " & FRectSearchKey & " = '" & FRectSearchString & "' "
		end if

		sql = sql & "and vm.idx = " & CStr(FRectMasterIdx) & " " & vbCrlf

		'response.write sql & "<Br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		set FItemOne = new CTenByTenVacationMasterItem

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				FItemOne.Fidx			= rsget("idx")
				FItemOne.Fuserid		= rsget("userid")
				FItemOne.Fempno			= rsget("empno")
				FItemOne.Fdivcd			= rsget("divcd")
				FItemOne.Fstartday		= rsget("startday")
				FItemOne.Fendday		= rsget("endday")
				FItemOne.Ftotalvacationday		= rsget("totalvacationday")
				FItemOne.Fusedvacationday		= rsget("usedvacationday")
				FItemOne.Frequestedday	= rsget("requestedday")
				FItemOne.Fdeleteyn		= rsget("deleteyn")
				FItemOne.Fregisterid	= rsget("registerid")
				FItemOne.Fregdate		= rsget("regdate")
				FItemOne.Fusername		= rsget("username")
				FItemOne.Fpart_name		= rsget("part_name")
				FItemOne.Fposit_name	= rsget("posit_name")
				FItemOne.Fjob_name		= rsget("job_name")
				FItemOne.Fposit_sn		= rsget("posit_sn")

				FItemOne.FpromotionDay		= rsget("promotionDay")
				FItemOne.FjungsanDay		= rsget("jungsanDay")
				FItemOne.FretireJungsanDay	= rsget("retireJungsanDay")
				FItemOne.Fcomment			= rsget("comment")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	'//2011.05.09 정윤정추가
	'//휴가 상세내역보기
	public Sub GetDetailOne()
		dim strSql
	 	strSql ="[db_partner].[dbo].[sp_Ten_vacation_detail_getData]( "&FRectdetailIdx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			Fidx				= rsget("idx")
			Fmasteridx          = rsget("masteridx")
			Fstartday           = rsget("startday")
			Fendday             = rsget("endday")
			Ftotalday           = rsget("totalday")
			statedivcd          = rsget("statedivcd")
			Fregdate           = rsget("regdate")
			Fempno              = rsget("empno")
			Fuserid             = rsget("userid")
			Fdivcd              = rsget("divcd")
			Ftotstartday        = rsget("totstartday")
			Ftotendday          = rsget("totendday")
			Ftotalvacationday   = rsget("totalvacationday")
			Fusedvacationday    = rsget("usedvacationday")
			Frequestedday       = rsget("requestedday")
			Fhalfgubun			= rsget("halfgubun")
			FworkAgent			= rsget("workAgent")
			FcallNum			= rsget("callNum")
		END IF
		rsget.Close

		set FItemOne = new CTenByTenVacationMasterItem
			FItemOne.Fdivcd = Fdivcd
			FdivcdStr  = FItemOne.GetDivCDStr
		set FItemOne = nothing

	end Sub

	public function fnGetMasterIdx
	dim strSql
	 	strSql ="[db_partner].[dbo].[sp_Ten_vacationdetail_getMasteridx]( "&FRectdetailIdx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
		fnGetMasterIdx = rsget(0)
		ELSE
		fnGetMasterIdx = 0
		END IF
		rsget.Close
	End Function
	public Sub GetMasterSummary()
		dim sql, i

	end Sub

	public Sub GetDetailList()
		dim sql, i
		dim monstartday, monendday

		if FRectEmpNO="" and FRectMasterIdx="" then exit Sub

		if FRectStartDate <> "" and FRectEndDate <> "" THEN
			monstartday = FRectStartDate
			monendday = FRectEndDate
		elseif (FRectYYYY <> "") and (FRectMM <> "") then
			monstartday = CStr(FRectYYYY) + "-" + CStr(FRectMM) + "-01"
			monendday = DateAdd("d", -1, DateAdd("m", 1, monstartday))		'한달을 더하고 하루를 빼주면 마지막날이 된다.
			monendday = Left(monendday, 10)
		end if


		'// 개수 파악 //
		sql = "select count(vm.idx), Ceiling(Cast(Count(vm.idx) as float)/" & FPageSize & ") " & vbCrlf
		sql = sql & "from [db_partner].[dbo].tbl_vacation_master vm, [db_partner].[dbo].tbl_vacation_detail vd, [db_partner].[dbo].tbl_user_tenbyten t " & vbCrlf
		sql = sql & "where 1 = 1 " & vbCrlf
		sql = sql & "and vm.idx = vd.masteridx " & vbCrlf
		sql = sql & "and vm.empno = t.empno " & vbCrlf

		if (FRectEmpNO <> "") then
			sql = sql & "and vm.empno = '" + CStr(FRectEmpNO) + "' " & vbCrlf
		else
			sql = sql & "and vm.idx = " & CStr(FRectMasterIdx) & " " & vbCrlf
		end if

		if (monstartday <> "") and (monendday <> "") then
			sql = sql & "and " & vbCrlf
			sql = sql & "	( " & vbCrlf
			sql = sql & "		((vd.startday <= '" + CStr(monstartday) + "') and (vd.endday >= '" + CStr(monstartday) + "')) " & vbCrlf
			sql = sql & "		or " & vbCrlf
			sql = sql & "		((vd.startday > '" + CStr(monstartday) + "') and (vd.startday <= '" + CStr(monendday) + "')) " & vbCrlf
			sql = sql & ") " & vbCrlf
		end if

		if (FRectIsDelete <> "") then
			'sql = sql & "and vm.deleteyn = '" & FRectIsDelete & "' " & vbCrlf
			sql = sql & "and vd.deleteyn = '" & FRectIsDelete & "' " & vbCrlf
		end if

		if (FRectPart_sn <> "") then
			sql = sql & "and t.part_sn = " & CStr(FRectPart_sn) & " " & vbCrlf
		end if

		'// 검색어 쿼리 //
		if FRectSearchKey<>"" and FRectSearchString<>"" then
			sql = sql & " and " & FRectSearchKey & " = '" & FRectSearchString & "' "
		end if

		'response.write sql & "<BR>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'// 목록 //
		sql = "select top " & CStr(FPageSize*FCurrPage) & " " & vbCrlf
		sql = sql & "	vm.divcd as masterdivcd, vd.idx, vd.masteridx, vd.startday, vd.endday, vd.totalday, vd.approverid, vd.approveday, vd.statedivcd, vd.deleteyn, vd.registerid, vd.regdate " & vbCrlf
		sql = sql & "	,p.reportidx, p.reportstate, vd.halfgubun " & vbCrlf
		sql = sql & "	, vd.approverempno, vd.registerempno, ta.username as approvername, tr.username as registername, isNull(vd.comment,'') as comment " & vbCrlf
		sql = sql & " from [db_partner].[dbo].tbl_vacation_master as  vm " & vbCrlf
		sql = sql & " inner join [db_partner].[dbo].tbl_vacation_detail as vd on vm.idx = vd.masteridx and vd.deleteyn = 'N'  " & vbCrlf
		sql = sql & " inner join [db_partner].[dbo].tbl_user_tenbyten as t  on  vm.empno = t.empno " & vbCrlf
		sql = sql & " left join [db_partner].[dbo].tbl_user_tenbyten as ta  on  vd.approverempno = ta.empno " & vbCrlf
		sql = sql & " left join [db_partner].[dbo].tbl_user_tenbyten as tr  on  vd.registerempno = tr.empno " & vbCrlf
		sql = sql & " left outer join db_partner.dbo.tbl_eappreport as p on vd.idx = p.scmlinkNo and p.isUsing =1   and p.edmsidx = 22 " & vbCrlf	 '휴가신청문서 번호 하드코딩처리
		sql = sql & "where 1 = 1 " & vbCrlf

		if (FRectEmpNO <> "") then
			sql = sql & "and vm.empno = '" + CStr(FRectEmpNO) + "' " & vbCrlf
		else
			sql = sql & "and vm.idx = " & CStr(FRectMasterIdx) & " " & vbCrlf
		end if

		if (monstartday <> "") and (monendday <> "") then
			sql = sql & "and " & vbCrlf
			sql = sql & "	( " & vbCrlf
			sql = sql & "		((convert(varchar(10),vd.startday,121) <= '" + CStr(monstartday) + "') and (convert(varchar(10),vd.endday,121) >= '" + CStr(monstartday) + "')) " & vbCrlf
			sql = sql & "		or " & vbCrlf
			sql = sql & "		((convert(varchar(10),vd.startday,121) > '" + CStr(monstartday) + "') and (convert(varchar(10),vd.startday,121) <= '" + CStr(monendday) + "')) " & vbCrlf
			sql = sql & ") " & vbCrlf
		end if

		if (FRectIsDelete <> "") then
			sql = sql & "and vm.deleteyn = '" & FRectIsDelete & "' " & vbCrlf
		end if

		if (FRectPart_sn <> "") then
			sql = sql & "and t.part_sn = " & CStr(FRectPart_sn) & " " & vbCrlf
		end if

		'// 검색어 쿼리 //
		if FRectSearchKey<>"" and FRectSearchString<>"" then
			sql = sql & " and " & FRectSearchKey & " = '" & FRectSearchString & "' "
		end if

		sql = sql & "order by vd.idx desc " & vbCrlf

		'response.write sql & "<Br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		'response.write FResultCount

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTenByTenVacationDetailItem

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fmasteridx		= rsget("masteridx")
				FItemList(i).FmasterDivCD	= rsget("masterdivcd")

				FItemList(i).Fstartday		= rsget("startday")
				FItemList(i).Fendday		= rsget("endday")
				FItemList(i).Ftotalday		= rsget("totalday")
				FItemList(i).Fapproverid	= rsget("approverid")
				FItemList(i).Fapproverempno	= rsget("approverempno")
				FItemList(i).Fapprovername	= rsget("approvername")

				FItemList(i).Fapproveday	= rsget("approveday")
				FItemList(i).Fstatedivcd	= rsget("statedivcd")
				FItemList(i).Fdeleteyn		= rsget("deleteyn")
				FItemList(i).Fregisterid	= rsget("registerid")
				FItemList(i).Fregisterempno	= rsget("registerempno")
				FItemList(i).Fregistername	= rsget("registername")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Freportidx		= rsget("reportidx")
				FItemList(i).Freportstate	= rsget("reportstate")
				FItemList(i).Fhalfgubun		= rsget("halfgubun")
				FItemList(i).Fcomment		= rsget("comment")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetVacationList()

		dim sql, i
		dim lastday, tmp, tmp2
		dim daybeforetwomonth, dayaftertwomonth, basemonth

		if ((FRectYYYY = "") or (FRectMM = "")) then
			FRectYYYY = Year(now)
			FRectMM = Month()
		end if

		if (CInt(FRectMM) < 10) then
			FRectMM = "0" + CStr(CInt(FRectMM))
		end if
		basemonth = CStr(FRectYYYY) + "-" + CStr(FRectMM) + "-01"

		tmp = CDate(basemonth)
		tmp2 = DateAdd("m", -2, tmp)
		daybeforetwomonth = CStr(Year(tmp2)) + "-" + CStr(Month(tmp2)) + "-01"
		if (Month(tmp2) < 10) then
			daybeforetwomonth = CStr(Year(tmp2)) + "-0" + CStr(Month(tmp2)) + "-01"
		end if


		tmp2 = DateAdd("m", 2, tmp)
		dayaftertwomonth = CStr(Year(tmp2)) + "-" + CStr(Month(tmp2)) + "-01"
		if (Month(tmp2) < 10) then
			dayaftertwomonth = CStr(Year(tmp2)) + "-0" + CStr(Month(tmp2)) + "-01"
		end if

		tmp2 = DateAdd("d", -1, DateAdd("m", 1, basemonth))		'한달을 더하고 하루를 빼주면 마지막날이 된다.
		lastday = Left(tmp2, 10)



		sql = "select count(s.solar_date) " & vbCrlf
		sql = sql & "from db_sitemaster.dbo.LunarToSolar s " & vbCrlf
		sql = sql & "	left join ( " & vbCrlf
		sql = sql & "		select pa.part_sn,m.userid,t.username, pa.part_name, d.startday, d.endday, d.statedivcd " & vbCrlf
		sql = sql & "		from  " & vbCrlf
		sql = sql & "			[db_partner].[dbo].tbl_vacation_master m " & vbCrlf
		sql = sql & "			, [db_partner].[dbo].tbl_vacation_detail d " & vbCrlf
		sql = sql & "			, [db_partner].[dbo].tbl_user_tenbyten t " & vbCrlf
		sql = sql & "			left join [db_partner].dbo.tbl_partInfo as pa on t.part_sn = pa.part_sn " & vbCrlf
		sql = sql & "		where m.idx = d.masteridx " & vbCrlf
		sql = sql & "		and m.empno = t.empno " & vbCrlf
		sql = sql & "		and m.deleteyn <> 'Y' " & vbCrlf
		sql = sql & "		and d.deleteyn <> 'Y' " & vbCrlf

		if (FRectPart_sn <> "") then
			sql = sql & "and t.part_sn = " & CStr(FRectPart_sn) & " " & vbCrlf
		end if

		if FRectSearchKey<>"" and FRectSearchString<>"" then
			sql = sql & " and " & FRectSearchKey & " = '" & FRectSearchString & "' "
		end if

		sql = sql & "		and d.statedivcd in ('R', 'A') " & vbCrlf
		sql = sql & "		and d.endday >= '" + daybeforetwomonth + "' and d.endday < '" + dayaftertwomonth + "' " & vbCrlf
		sql = sql & "		and d.startday >= '" + daybeforetwomonth + "' and d.startday < '" + dayaftertwomonth + "' " & vbCrlf
		sql = sql & "	) as v on datediff(d,v.startday,s.solar_date) >= 0 and datediff(d,v.endday,s.solar_date) <= 0 " & vbCrlf
		sql = sql & "where 1 = 1 " & vbCrlf
		sql = sql & "and s.solar_date >= '" + basemonth + "' and s.solar_date <= '" + lastday + "' " & vbCrlf

		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = 1
			FPageSize = FTotalCount
		rsget.Close



		sql = "select top " + CStr(FTotalCount) + " s.solar_date, s.holiday, s.holiday_name, v.* " & vbCrlf
		sql = sql & "from db_sitemaster.dbo.LunarToSolar s " & vbCrlf
		sql = sql & "	left join ( " & vbCrlf
		sql = sql & "		select m.idx, pa.part_sn, m.userid, t.username, pa.part_name, d.startday, d.endday, d.statedivcd, d.totalday, d.halfgubun, d.workAgent, d.callNum" & vbCrlf
		sql = sql & "		from  " & vbCrlf
		sql = sql & "			[db_partner].[dbo].tbl_vacation_master m " & vbCrlf
		sql = sql & "			, [db_partner].[dbo].tbl_vacation_detail d " & vbCrlf
		sql = sql & "			, [db_partner].[dbo].tbl_user_tenbyten t " & vbCrlf
		sql = sql & "			left join [db_partner].dbo.tbl_partInfo as pa on t.part_sn = pa.part_sn " & vbCrlf
		sql = sql & "		where  m.idx = d.masteridx " & vbCrlf
		sql = sql & "		and m.empno = t.empno " & vbCrlf
		sql = sql & "		and m.deleteyn <> 'Y' " & vbCrlf
		sql = sql & "		and d.deleteyn <> 'Y' " & vbCrlf

		if (FRectPart_sn <> "") then
			sql = sql & "and t.part_sn = " & CStr(FRectPart_sn) & " " & vbCrlf
		end if

		if FRectSearchKey<>"" and FRectSearchString<>"" then
			sql = sql & " and " & FRectSearchKey & " = '" & FRectSearchString & "' "
		end if

		sql = sql & "		and d.statedivcd in ('R', 'A') " & vbCrlf
		sql = sql & "		and d.endday >= '" + daybeforetwomonth + "' and d.endday < '" + dayaftertwomonth + "' " & vbCrlf
		sql = sql & "		and d.startday >= '" + daybeforetwomonth + "' and d.startday < '" + dayaftertwomonth + "' " & vbCrlf
		sql = sql & "	) as v on datediff(d,v.startday,s.solar_date) >= 0 and datediff(d,v.endday,s.solar_date) <= 0 " & vbCrlf
		sql = sql & "where 1 = 1 " & vbCrlf
		sql = sql & "and s.solar_date >= '" + basemonth + "' and s.solar_date <= '" + lastday + "' " & vbCrlf
		sql = sql & "order by s.solar_date, v.part_sn, v.username " & vbCrlf
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly
		'response.write sql

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		'response.write FResultCount

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTenByTenVacationCalendarItem

				FItemList(i).Fmasteridx		= rsget("idx")
				FItemList(i).Fsolar_date	= rsget("solar_date")
				FItemList(i).Fpart_sn		= rsget("part_sn")
				FItemList(i).Fuserid		= rsget("userid")
				FItemList(i).Fusername		= rsget("username")
				FItemList(i).Fpart_name		= rsget("part_name")
				FItemList(i).Fstartday		= rsget("startday")
				FItemList(i).Fendday		= rsget("endday")
				FItemList(i).Fstatedivcd	= rsget("statedivcd")
				FItemList(i).Ftotalday		= rsget("totalday")
				FItemList(i).Fhalfgubun		= rsget("halfgubun")
				FItemList(i).Fholiday		= rsget("holiday")
				FItemList(i).Fholiday_name	= rsget("holiday_name")
				FItemList(i).FworkAgent		= rsget("workAgent")
				FItemList(i).FcallNum		= rsget("callNum")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetVacationListSimple()

		dim sql, i
		dim basemonth, lastday

		if ((FRectYYYY = "") or (FRectMM = "")) then
			FRectYYYY = Year(now)
			FRectMM = Month()
		end if

		basemonth = DateSerial(FRectYYYY,FRectMM,1)
		lastday = DateAdd("d", -1, DateAdd("m", 1, basemonth))		'한달을 더하고 하루를 빼주면 마지막날이 된다.

		sql = "select m.idx, pa.part_sn, m.userid, t.username, '' as holiday_name, pa.part_name, d.startday, d.endday, d.statedivcd, d.totalday, d.halfgubun, d.workAgent, d.callNum, 0 as holiday " & vbCrlf
		sql = sql & "from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
		sql = sql & "	join [db_partner].[dbo].tbl_vacation_detail d on m.idx = d.masteridx " & vbCrlf
		sql = sql & "	join [db_partner].[dbo].tbl_user_tenbyten t on m.empno = t.empno " & vbCrlf
		sql = sql & "	left join [db_partner].dbo.tbl_partInfo as pa on t.part_sn = pa.part_sn " & vbCrlf
		if (Fdepartment_id<>"") then
			sql = sql & "	left join db_partner.[dbo].[vw_user_department_v2] as p on t.department_id = p.cid and p.useYN='Y' " & vbCrlf
		end if
		sql = sql & "where m.deleteyn <> 'Y' " & vbCrlf
		sql = sql & "	and d.deleteyn <> 'Y' " & vbCrlf
		sql = sql & "	and d.statedivcd in ('R', 'A') " & vbCrlf
		sql = sql & "	and (d.endday >= '" & basemonth & "' and d.endday < '" & lastday & " 23:59:59:999' " & vbCrlf
		sql = sql & "	or d.startday >= '" & basemonth & "' and d.startday < '" & lastday & "') " & vbCrlf
		if (FRectPart_sn <> "") then
			sql = sql & "and t.part_sn = " & CStr(FRectPart_sn) & " " & vbCrlf
		end if
		if (Fdepartment_id<>"") then
			sql = sql & "and p.cidArr like '%," & CStr(Fdepartment_id) & ",%' " & vbCrlf
		end if
		if FRectSearchKey<>"" and FRectSearchString<>"" then
			sql = sql & " and " & FRectSearchKey & " = '" & FRectSearchString & "' "
		end if
		sql = sql & "union " & vbCrlf
		sql = sql & "select 0,0,'','',holiday_name,'',solar_date as startday, solar_date as endday,'',1,'no','','',holiday " & vbCrlf
		sql = sql & "from db_sitemaster.dbo.LunarToSolar " & vbCrlf
		sql = sql & "where solar_date between '" & basemonth & "' and '" & lastday & "' " & vbCrlf
		sql = sql & "	and holiday>0 " & vbCrlf
		sql = sql & "	and isNull(holiday_name,'')<>'' " & vbCrlf
		sql = sql & "order by startday"
        rsget.CursorLocation = adUseClient
        rsget.Open sql,dbget,adOpenForwardOnly,adLockReadOnly
		'response.write sql

		FResultCount = rsget.RecordCount

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		'response.write FResultCount

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTenByTenVacationCalendarItem

				FItemList(i).Fmasteridx		= rsget("idx")
				FItemList(i).Fpart_sn		= rsget("part_sn")
				FItemList(i).Fuserid		= rsget("userid")
				FItemList(i).Fusername		= rsget("username")
				FItemList(i).Fpart_name		= rsget("part_name")
				FItemList(i).Fstartday		= rsget("startday")
				FItemList(i).Fendday		= rsget("endday")
				FItemList(i).Fstatedivcd	= rsget("statedivcd")
				FItemList(i).Ftotalday		= rsget("totalday")
				FItemList(i).Fhalfgubun		= rsget("halfgubun")
				FItemList(i).FworkAgent		= rsget("workAgent")
				FItemList(i).FcallNum		= rsget("callNum")
				FItemList(i).Fholiday		= rsget("holiday")
				FItemList(i).Fholiday_name		= rsget("holiday_name")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function


	'##### 사용자 내용 접수 #####
	public Sub GetMember()
		dim SQL

		SQL =	"select top 1 " & vbCrlf
		SQL =	SQL & "	p.id as userid, p.password, t.part_sn, t.posit_sn, p.level_sn, t.job_sn, p.isusing, p.userdiv, p.bigo, t.empno, " & vbCrlf
		SQL =	SQL & "	t.frontid, t.username, t.juminno, t.birthday, t.issolar, t.sexflag, IsNull(t.zipcode,'000-000') as zipcode, t.useraddr, t.userphone, t.usercell, t.usermail, t.msnmail, t.interphoneno, t.extension, t.direct070, t.jobdetail, t.statediv, t.joinday, t.realjoinday, t.retireday, t.regdate, " & vbCrlf
		SQL =	SQL & "	pa.part_name, " & vbCrlf
		SQL =	SQL & "	po.posit_name, " & vbCrlf
		SQL =	SQL & "	le.level_name, " & vbCrlf
		SQL =	SQL & "	jo.job_name " & vbCrlf
		SQL =	SQL & "from " & vbCrlf
		SQL =	SQL & "	[db_partner].[dbo].tbl_user_tenbyten as t  " & vbCrlf
		SQL =	SQL & "	inner join [db_partner].[dbo].tbl_partner as p on t.userid = p.id and p.userdiv < 999 and p.level_sn < 999 " & vbCrlf
		SQL =	SQL & "	left join [db_partner].dbo.tbl_partInfo as pa on t.part_sn = pa.part_sn " & vbCrlf
		SQL =	SQL & "	left join [db_partner].dbo.tbl_positInfo as po on t.posit_sn = po.posit_sn " & vbCrlf
		SQL =	SQL & "	left join [db_partner].dbo.tbl_level as le on p.level_sn = le.level_sn " & vbCrlf
		SQL =	SQL & "	left join [db_partner].dbo.tbl_JobInfo as jo on t.job_sn = jo.job_sn " & vbCrlf
		SQL =	SQL & "where " & vbCrlf
		SQL =	SQL & "	 t.part_sn <> 4 " & vbCrlf		'오프라인 - 직영점 제외
		SQL =	SQL & "	and t.part_sn <> 5 " & vbCrlf		'오프라인 - 가맹점 제외
		SQL =	SQL & "	and t.userid = '" & FRectUserId & "' " & vbCrlf
        rsget.CursorLocation = adUseClient
        rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly

		if Not(rsget.EOF or rsget.BOF) then

			FResultCount = 1
			redim preserve FItemList(1)
			set FItemList(1) = new CTenByTenMemberItem

			FItemList(1).Fuserid		= rsget("userid")
			FItemList(1).Fpart_sn		= rsget("part_sn")
			FItemList(1).Fposit_sn		= rsget("posit_sn")
			FItemList(1).Flevel_sn		= rsget("level_sn")
			FItemList(1).Fjob_sn		= rsget("job_sn")
			FItemList(1).Fisusing		= rsget("isusing")
			FItemList(1).Fuserdiv		= rsget("userdiv")
			FItemList(1).Fbigo			= rsget("bigo")

			'딱히 필요없다.
			'FItemList(1).Fuserpass		= rsget("password")

			FItemList(1).Fempno			= rsget("empno")
			FItemList(1).Ffrontid		= rsget("frontid")
			FItemList(1).Fusername		= rsget("username")
			FItemList(1).Fjuminno		= rsget("juminno")
			FItemList(1).Fbirthday		= rsget("birthday")
			FItemList(1).Fissolar		= rsget("issolar")
			FItemList(1).Fsexflag		= rsget("sexflag")
			FItemList(1).Fzipcode		= rsget("zipcode")
			FItemList(1).Fuseraddr		= rsget("useraddr")
			FItemList(1).Fuserphone		= rsget("userphone")
			FItemList(1).Fusercell		= rsget("usercell")
			FItemList(1).Fusermail		= rsget("usermail")
			FItemList(1).Fmsnmail		= rsget("msnmail")
			FItemList(1).Finterphoneno	= rsget("interphoneno")
			FItemList(1).Fextension		= rsget("extension")
		FItemList(1).Fdirect070		= rsget("direct070")
			FItemList(1).Fjobdetail		= rsget("jobdetail")
			FItemList(1).Fstatediv		= rsget("statediv")
			FItemList(1).Fjoinday		= rsget("joinday")
			FItemList(1).Frealjoinday	= rsget("realjoinday")
			FItemList(1).Fretireday		= rsget("retireday")
			FItemList(1).Fregdate		= rsget("regdate")

			FItemList(1).Fpart_name		= rsget("part_name")
			FItemList(1).Fposit_name	= rsget("posit_name")
			FItemList(1).Flevel_name	= rsget("level_name")
			FItemList(1).Fjob_name		= rsget("job_name")

				if IsNull(FItemList(1).Fbigo) or (FItemList(1).Fbigo = "0") then
					FItemList(1).Fbigo = ""
				end if

		else

			FResultCount = 0

		end if

		rsget.Close

	if (Len(FItemList(1).Fzipcode) = 7) then
			FItemList(1).Fzipaddr = GetZipAddress(FItemList(1).Fzipcode)
		end if
	end Sub
end Class



'/// 우편번호에서 주소 반환 함수 ///
public function GetZipAddress(zipcode)
	dim zip1, zip2, tmp, result
	dim sql

	tmp = Split(zipcode, "-")
	zip1 = tmp(0)
	zip2 = tmp(1)

	result = ""
	SQL =	"select top 1 (ADDR050_SI + ' ' + ADDR050_GU) as zipaddr " & vbCrlf
	SQL =	SQL & "	from [db_zipcode].[dbo].ADDR050TL " & vbCrlf
	SQL =	SQL & "	where 1 = 1 " & vbCrlf
	SQL =	SQL & "	and ADDR050_ZIP1 = '" & CStr(zip1) & "' " & vbCrlf
	SQL =	SQL & "	and ADDR050_ZIP2 = '" & CStr(zip2) & "' " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		result = rsget("zipaddr")
	end if
	rsget.Close

	GetZipAddress = result
end function



'==============================================================================
'작년 연차 총일수
public function GetPrevYearTotalVacationDay(userid)
	dim tmp, result
	dim sql

	result = ""
	SQL =	"select IsNull(sum(m.totalvacationday), 0) as prevyeartotalvacationday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where m.userid = '" & CStr(userid) & "' " & vbCrlf
	SQL =	SQL & "	and m.userid <> '' " & vbCrlf
	SQL =	SQL & "	and m.divcd = '1' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = (Year(getdate()) - 1) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		result = rsget("prevyeartotalvacationday")
	end if
	rsget.Close

	GetPrevYearTotalVacationDay = result
end function



'==============================================================================
'작년 연차 작년 사용일(전년이월년차 계산에 필요)
public function GetPrevYearUsedVacationDay(userid)
	dim tmp, result
	dim sql

	result = ""
	SQL =	"select IsNull(sum(d.totalday), 0) as prevyearusedvacationday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m, [db_partner].[dbo].tbl_vacation_detail d " & vbCrlf
	SQL =	SQL & "	where userid = '" & CStr(userid) & "' " & vbCrlf
	SQL =	SQL & "	and m.userid <> '' " & vbCrlf
	SQL =	SQL & "	and m.divcd = '1' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and d.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and d.statedivcd = 'A' " & vbCrlf
	SQL =	SQL & "	and d.masteridx = m.idx " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = (Year(getdate()) - 1) " & vbCrlf
	SQL =	SQL & "	and Year(d.startday) = (Year(getdate()) - 1) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		result = rsget("prevyearusedvacationday")
	end if
	rsget.Close

	GetPrevYearUsedVacationDay = result
end function


'==============================================================================
'작년 연차 금년 사용일
public function GetPrevYearCurrUsedVacationDay(userid)
	dim tmp, result
	dim sql

	result = ""
	SQL =	"select IsNull(sum(d.totalday), 0) as prevyearcurrusedvacationday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m, [db_partner].[dbo].tbl_vacation_detail d " & vbCrlf
	SQL =	SQL & "	where  userid = '" & CStr(userid) & "' " & vbCrlf
	SQL =	SQL & "	and m.userid <> '' " & vbCrlf
	SQL =	SQL & "	and m.divcd = '1' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and d.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and d.statedivcd = 'A' " & vbCrlf
	SQL =	SQL & "	and d.masteridx = m.idx " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = (Year(getdate()) - 1) " & vbCrlf
	SQL =	SQL & "	and Year(d.startday) = Year(getdate()) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		result = rsget("prevyearcurrusedvacationday")
	end if
	rsget.Close

	GetPrevYearCurrUsedVacationDay = result
end function


'==============================================================================
'작년 연차 작년 승인대기 - 연단위 연차 생성시 작년 연차승인대기 모두 삭제 및 requestedday = 0 필요
'작년 연차 금년 승인대기
public function GetPrevYearRequestedVacationDay(userid)
	dim tmp, result
	dim sql

	result = ""

	SQL =	"select " & vbCrlf
	SQL =	SQL & "		IsNull(sum(m.requestedday), 0) as prevyearrequestedday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where  m.userid = '" & CStr(userid) & "' " & vbCrlf
	SQL =	SQL & "	and m.userid <> '' " & vbCrlf
	SQL =	SQL & "	and m.divcd = '1' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = (Year(getdate()) - 1) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		result = rsget("prevyearrequestedday")
	end if
	rsget.Close

	GetPrevYearRequestedVacationDay = result
end function



'==============================================================================
'금년 연차 총일수 / 금년 연차 사용일 / 금년 연차 승인대기
public function GetCurrYearVacationDay(userid, byref curryeartotalvacationday, byref curryearusedvacationday, byref curryearrequestedday)
	dim tmp, result
	dim sql

	result = ""

	SQL =	"select " & vbCrlf
	SQL =	SQL & "		IsNull(sum(m.totalvacationday), 0) as curryeartotalvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.usedvacationday), 0) as curryearusedvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.requestedday), 0) as curryearrequestedday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where   m.userid = '" & CStr(userid) & "' " & vbCrlf
	SQL =	SQL & "	and m.userid <> '' " & vbCrlf
	SQL =	SQL & "	and m.divcd = '1' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = Year(getdate()) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		curryeartotalvacationday = rsget("curryeartotalvacationday")
		curryearusedvacationday = rsget("curryearusedvacationday")
		curryearrequestedday = rsget("curryearrequestedday")
	end if
	rsget.Close

	GetCurrYearVacationDay = result
end function


'==============================================================================
'금년 휴가 총일수 / 금년 휴가 사용일 / 금년 휴가 승인대기
public function GetCurrVacationDay(userid, byref currtotalvacationday, byref currusedvacationday, byref currrequestedday)
	dim tmp, result
	dim sql

	result = ""

	SQL =	"select " & vbCrlf
	SQL =	SQL & "		IsNull(sum(m.totalvacationday), 0) as currtotalvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.usedvacationday), 0) as currusedvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.requestedday), 0) as currrequestedday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where  m.userid = '" & CStr(userid) & "' " & vbCrlf
	SQL =	SQL & "	and m.userid <> '' " & vbCrlf
	SQL =	SQL & "	and m.divcd <> '1' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = Year(getdate()) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		currtotalvacationday = rsget("currtotalvacationday")
		currusedvacationday = rsget("currusedvacationday")
		currrequestedday = rsget("currrequestedday")
	end if
	rsget.Close

	GetCurrVacationDay = result
end function



'==============================================================================
'총일수 / 사용일 / 승인대기
public function GetVacationDay(userid, divcd, byref totalvacationday, byref usedvacationday, byref requestedday, byref expiredday)
	dim tmp, result
	dim sql

	result = ""

	SQL =	"select " & vbCrlf
	SQL =	SQL & "		IsNull(sum(m.totalvacationday), 0) as totalvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.usedvacationday), 0) as usedvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.requestedday), 0) as requestedday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(case when m.endday >= getdate() then 0 else (m.totalvacationday - m.usedvacationday) end), 0) as expiredday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where m.userid = '" & CStr(userid) & "' " & vbCrlf
	SQL =	SQL & "	and m.userid <> '' " & vbCrlf
	SQL =	SQL & "	and m.divcd = '" & divcd & "' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = Year(getdate()) " & vbCrlf
	'response.write sql

	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		totalvacationday = rsget("totalvacationday")
		usedvacationday = rsget("usedvacationday")
		requestedday = rsget("requestedday")
		expiredday = rsget("expiredday")
	end if
	rsget.Close

	GetVacationDay = result
end function



'==============================================================================
'작년 휴가 총일수 / 사용일 / 승인대기 / 만료일
public function GetPrevYearVacationDay(userid, byref totalvacationday, byref usedvacationday, byref requestedday, byref expiredday)
	dim tmp, result
	dim sql

	result = ""

	SQL =	"select " & vbCrlf
	SQL =	SQL & "		IsNull(sum(m.totalvacationday), 0) as totalvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.usedvacationday), 0) as usedvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.requestedday), 0) as requestedday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(case when m.endday >= getdate() then 0 else (m.totalvacationday - m.usedvacationday) end), 0) as expiredday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where m.userid = '" & CStr(userid) & "' " & vbCrlf
	SQL =	SQL & "	and m.userid <> '' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = (Year(getdate()) - 1) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		totalvacationday = rsget("totalvacationday")
		usedvacationday = rsget("usedvacationday")
		requestedday = rsget("requestedday")
		expiredday = rsget("expiredday")
	end if
	rsget.Close

	GetPrevYearVacationDay = result
end function


'==============================================================================
'작년 휴가 총일수 / 사용일 / 승인대기 / 만료일
public function GetPrevYearVacationDayByEmpno(empno, byref totalvacationday, byref usedvacationday, byref requestedday, byref expiredday)
	dim tmp, result
	dim sql

	result = ""

	SQL =	"select " & vbCrlf
	SQL =	SQL & "		IsNull(sum(m.totalvacationday), 0) as totalvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.usedvacationday), 0) as usedvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.requestedday), 0) as requestedday " & vbCrlf   '만료일수 승인대기도 제외처리 2014/10/02 정윤정수정
	SQL =	SQL & "		, IsNull(sum(case when m.endday >= getdate() then 0 else (m.totalvacationday - m.usedvacationday-m.requestedday) end), 0) as expiredday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where m.empno = '" & CStr(empno) & "' " & vbCrlf
	SQL =	SQL & "	and m.empno <> '' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = (Year(getdate()) - 1) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		totalvacationday = rsget("totalvacationday")
		usedvacationday = rsget("usedvacationday")
		requestedday = rsget("requestedday")
		expiredday = rsget("expiredday")
	end if
	rsget.Close

	GetPrevYearVacationDayByEmpno = result
end function


'==============================================================================
'금년 휴가 총일수 / 사용일 / 승인대기 / 만료일
public function GetCurrYearVacationDay(userid, byref totalvacationday, byref usedvacationday, byref requestedday, byref expiredday)
	dim tmp, result
	dim sql

	result = ""

	SQL =	"select " & vbCrlf
	SQL =	SQL & "		IsNull(sum(m.totalvacationday), 0) as totalvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.usedvacationday), 0) as usedvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.requestedday), 0) as requestedday " & vbCrlf     '만료일수 승인대기도 제외처리 2014/10/02 정윤정수정
	SQL =	SQL & "		, IsNull(sum(case when m.endday >= getdate() then 0 else (m.totalvacationday - m.usedvacationday) end), 0) as expiredday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where  m.userid = '" & CStr(userid) & "' " & vbCrlf
	SQL =	SQL & "	and m.userid <> '' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = Year(getdate()) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		totalvacationday = rsget("totalvacationday")
		usedvacationday = rsget("usedvacationday")
		requestedday = rsget("requestedday")
		expiredday = rsget("expiredday")
	end if
	rsget.Close

	GetCurrYearVacationDay = result
end function


'==============================================================================
'금년 휴가 총일수 / 사용일 / 승인대기 / 만료일
public function GetCurrYearVacationDayByEmpno(empno, byref totalvacationday, byref usedvacationday, byref requestedday, byref expiredday)
	dim tmp, result
	dim sql

	result = ""

	SQL =	"select " & vbCrlf
	SQL =	SQL & "		IsNull(sum(m.totalvacationday), 0) as totalvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.usedvacationday), 0) as usedvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.requestedday), 0) as requestedday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(case when m.endday >= getdate() then 0 else (m.totalvacationday - m.usedvacationday-m.requestedday) end), 0) as expiredday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where  m.empno = '" & CStr(empno) & "' " & vbCrlf
	SQL =	SQL & "	and m.empno <> '' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf
	SQL =	SQL & "	and Year(m.startday) = Year(getdate()) " & vbCrlf
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		totalvacationday = rsget("totalvacationday")
		usedvacationday = rsget("usedvacationday")
		requestedday = rsget("requestedday")
		expiredday = rsget("expiredday")
	end if
	rsget.Close

	GetCurrYearVacationDayByEmpno = result
end function
'==============================================================================
'계약직  휴가 총일수 / 사용일 / 승인대기 / 만료일
public function GetPartYearVacationDayByEmpno(empno, byref totalvacationday, byref usedvacationday, byref requestedday, byref expiredday)
	dim tmp, result
	dim sql

	result = ""

	SQL =	"select " & vbCrlf
	SQL =	SQL & "		IsNull(sum(m.totalvacationday), 0) as totalvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.usedvacationday), 0) as usedvacationday " & vbCrlf
	SQL =	SQL & "		, IsNull(sum(m.requestedday), 0) as requestedday " & vbCrlf   '만료일수 승인대기도 제외처리 2014/10/02 정윤정수정
	SQL =	SQL & "		, IsNull(sum(case when m.endday >= getdate() then 0 else (m.totalvacationday - m.usedvacationday-m.requestedday) end), 0) as expiredday " & vbCrlf
	SQL =	SQL & "	from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	SQL =	SQL & "	where m.empno = '" & CStr(empno) & "' " & vbCrlf
	SQL =	SQL & "	and m.empno <> '' " & vbCrlf
	SQL =	SQL & "	and m.deleteyn <> 'Y' " & vbCrlf 
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		totalvacationday = rsget("totalvacationday")
		usedvacationday = rsget("usedvacationday")
		requestedday = rsget("requestedday")
		expiredday = rsget("expiredday")
	end if
	rsget.Close

	GetPartYearVacationDayByEmpno = result
end function
'/// 부서 옵션 생성 함수 ///
public function printPartOption(fnm, psn)
	dim SQL, i, strOpt

strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::부서선택::</option>"

	SQL =	"Select part_sn, part_name " &_
			"From db_partner.dbo.tbl_partInfo " &_
			"Where part_isDel='N' " &_
			"Order by part_sort"
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("part_sn") & "'"
			if Cstr(rsget("part_sn"))=Cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("part_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printPartOption = strOpt
end function



'/// 직급 옵션 생성 함수 ///
public function printPositOption(fnm, psn)
	dim SQL, i, strOpt

	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::직급선택::</option>"

	SQL =	"Select posit_sn, posit_name " &_
			"From db_partner.dbo.tbl_positInfo " &_
			"Where posit_isDel='N' "
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("posit_sn") & "'"
			if rsget("posit_sn")=psn then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("posit_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printPositOption = strOpt
end function



'/// 등급 옵션 생성 함수 ///
public function printLevelOption(fnm, psn)
	dim SQL, i, strOpt

	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::등급선택::</option>"

	SQL =	"Select level_sn, level_name " &_
			"From db_partner.dbo.tbl_level " &_
			"Where level_isDel='N' " &_
			"Order by level_no"
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("level_sn") & "'"
			if rsget("level_sn")=psn then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("level_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printLevelOption = strOpt
end function



'/// 직책 옵션 생성 함수 ///
public function printJobOption(fnm, jsn)
	dim SQL, i, strOpt

	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::직책선택::</option>"

	SQL =	"Select job_sn, job_name " &_
			"From db_partner.dbo.tbl_JobInfo " &_
			"Where job_isDel='N' "
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("job_sn") & "'"
			if rsget("job_sn")=jsn then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("job_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printJobOption = strOpt
end function




'/// 담당샵 옵션 생성 함수 ///
public function printShopOption(fnm, shopid)
	dim SQL, i, strOpt

	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value='0'>::담당샵선택::</option>"

	SQL =	"select userid, shopname " &_
			"from [db_shop].[dbo].tbl_shop_user " &_
			"where 1 = 1 " &_
			"and isusing <> 'N' " &_
			"and shopdiv in ('1', '9') " &_
			"order by userid "
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("userid") & "'"
			if rsget("userid")=shopid then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("shopname") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printShopOption = strOpt
end function

public function GetDayOrHourWithPositSN(po_sn, d)
	dim tmpStr, s

	if (po_sn = 13) then
		'// 시급계약직(시간을 구한다.)
		'// 1일은 8시간, 1시간은 0.125(= 1/8) 
		  
		 tmpStr = round((d / 0.125),2)  'tmpStr= Cstr(d/0.125) 에서 round 함수 추가 2016.05.10 정윤정
		    
'		if (InStr(tmpStr, ".") >= 1) and (Len(tmpStr) > (InStr(tmpStr, ".") + 2)) then
'			'// 소수점 이하가 길어지면 잘라버린다.
'			tmpStr = Left(tmpStr, (InStr(tmpStr, ".") + 2))
'		end if
 
		GetDayOrHourWithPositSN = tmpStr * 1.0
	else
		GetDayOrHourWithPositSN = d
	end if
end function

public function GetDayOrHourNameWithPositSN(po_sn)
	if (po_sn = 13) then
		GetDayOrHourNameWithPositSN = " 시간"
	else
		GetDayOrHourNameWithPositSN = " 일"
	end if
end function

%>
