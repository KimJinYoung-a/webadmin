<%
'####################################################
' Description :  사원관리 클래스
' History : 2011.1.19 정윤정 생성
'			2018.04.11 한용민 수정
'####################################################

Class CPartnerAddLevelItem
    public Fuserid
    public Fpart_sn
    public Flevel_sn
    public FisDefault
    public Fpart_name
    public Flevel_name

    Private Sub Class_Initialize()

    End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CPartnerAddLevel
    public FItemList()
	public FOneItem

    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectUserID
	public FRectOnlyAdd
	public FRectOnlyDefault

	public Sub getUserAddLevelList()
	    Dim SqlStr, i
	    sqlStr = "select L.userid, L.part_sn, v.level_sn, P.part_name, v.level_name, L.isDefault"
        sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_AddLevel L"
        sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partInfo p"
        sqlStr = sqlStr & "     on L.part_sn=p.part_sn"
        sqlStr = sqlStr & "     left join db_partner.dbo.tbl_level V"
        sqlStr = sqlStr & "     on V.level_sn=L.level_sn"
        sqlStr = sqlStr & " where L.UserID='"&FRectUserID&"'"

        if (FRectOnlyDefault<>"") then
            sqlStr = sqlStr & " and L.isDefault='Y'"
        end if

        if (FRectOnlyAdd<>"") then
            sqlStr = sqlStr & " and L.isDefault<>'Y'"
        end if
'rw sqlStr
        rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount  = FResultCount
        if (FResultCount<1) then FResultCount=0

	    redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CPartnerAddLevelItem
				FItemList(i).Fuserid      = rsget("userid")
                FItemList(i).Fpart_sn     = rsget("part_sn")
                FItemList(i).Flevel_sn    = rsget("level_sn")
                FItemList(i).FisDefault   = rsget("isDefault")
                FItemList(i).Fpart_name   = rsget("part_name")
                FItemList(i).Flevel_name  = rsget("level_name")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    End sub

    Private Sub Class_Initialize()

    End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CTenByTenMemberitem
	public fidx
	public flogidx
	public Fempno
	public Fuserid
	public fmenuid
	public flogtype
	public flogmsg
	public fisusing
	public fadminid
	public fregdate
	public fmenu_id
	public fpart_sn
	public flevel_sn

    Private Sub Class_Initialize()
    End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CTenByTenMember
	public FPageSize
	public FCurrPage
	public FResultCount
	public FScrollCount
	public FPageCount
	public FTotalCount
	public FItemList()

	public Fempno
	public Fuserid
	public Ffrontid
	public Fusername
	public fuserNameEN
	public Fjuminno
	public Fbirthday
	public Fissolar
	public Fsexflag
	public Fzipcode
	public Fzipaddr
	public Fuseraddr
	public Fuserphone
	public Fusercell
	public Fusermail
	public Fmsnmail
	public Fmessenger
	public Finterphoneno
	public Fextension
	public Fdirect070
	public Fjobdetail
	public Fstatediv
	public Fjoinday
	public Frealjoinday
	public Fretireday
	public fshopid
	public Fuserimage
	public Fmywork
	public Fpart_sn
	public Fposit_sn
	public Fjob_sn
	public Flevel_sn
	public Fuserdiv
	public Frank_sn
	public Fdepartmentname
	public Fpart_name
	public Fposit_name
	public Fjob_name
	public Flevel_name
	public FSearchType
	public FSearchText
	public Fextparttime
	public Forderby
	public FTotCnt
	public FSPageNo
	public FEPageNo
	public FchkDate
	public FisIdentify
	public Fretirereason
	public FStartDate
	public FEndDate
	public FMaxInoOnly
	public Fyyyymm
	public FBizsection_cd
	public Fcriticinfouser
	public Flv1customerYN
	public Flv2partnerYN
	public Flv3InternalYN
	public Fdepartment_id
	public FdepartmentNameFull
	public Finc_subdepartment
	public FRectCriticInfoUser
	public FRectNoDepartOnly
	public Fcid1
	public Fcid2
	public Fcid3
	public Fcid4

    public Fpartnerusing
	public Fpersonalmail
	public Fgsshopuserid
	public frectempno
	public frectmenuid
	public Frectlv1customerYN
	public Frectlv2partnerYN
	public Frectlv3InternalYN
	public frectUserId

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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

	' 직원어드민권한 로그 가져오기 		' 2020.08.19 한용민 생성
	' /admin/member/tenbyten/popAdminAuth.asp
	public Sub getUserTenbytenAdminAuthLog()
	    Dim SqlStr, i

		if frectempno="" or isnull(frectempno) then exit Sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
	    sqlStr = sqlStr & " logidx, empno, userid, logtype, logmsg, adminid, regdate"
        sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_authlog with (nolock)"
        sqlStr = sqlStr & " where empno='"& frectempno &"'"
		sqlStr = sqlStr & " order by logidx desc"

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount  = FResultCount
        if (FResultCount<1) then FResultCount=0

	    redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CTenByTenMemberitem
				FItemList(i).flogidx      = rsget("logidx")
                FItemList(i).fempno     = rsget("empno")
                FItemList(i).fuserid    = rsget("userid")
				FItemList(i).flogtype    = rsget("logtype")
                FItemList(i).flogmsg   = db2html(rsget("logmsg"))
                FItemList(i).fadminid  = rsget("adminid")
				FItemList(i).fregdate  = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    End sub

	' /admin/menu/menu_edit.asp
	public Sub getpartner_menu_log()
	    Dim SqlStr, i

		if frectmenuid="" or isnull(frectmenuid) then exit Sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
	    sqlStr = sqlStr & " idx,menuid,logtype,logmsg,isusing,adminid,regdate"
        sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_menu_log with (nolock)"
        sqlStr = sqlStr & " where isusing=N'Y' and menuid='"& frectmenuid &"'"
		sqlStr = sqlStr & " order by idx desc"

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount  = FResultCount
        if (FResultCount<1) then FResultCount=0

	    redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CTenByTenMemberitem
				FItemList(i).fidx      = rsget("idx")
                FItemList(i).fmenuid     = rsget("menuid")
                FItemList(i).flogtype    = rsget("logtype")
                FItemList(i).flogmsg   = db2html(rsget("logmsg"))
                FItemList(i).fisusing   = rsget("isusing")
                FItemList(i).fadminid  = rsget("adminid")
				FItemList(i).fregdate  = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    End sub

	' /admin/menu/menu_edit.asp
	public Sub getpartner_menu_part()
	    Dim SqlStr, i

		if frectmenuid="" or isnull(frectmenuid) then exit Sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " menu_id, part_sn, level_sn"
        sqlStr = sqlStr & " from db_partner.[dbo].[tbl_menu_part] with (nolock)"
        sqlStr = sqlStr & " where menu_id='"& frectmenuid &"'"
		sqlStr = sqlStr & " order by menu_id asc, part_sn asc, level_sn asc"

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount  = FResultCount
        if (FResultCount<1) then FResultCount=0

	    redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CTenByTenMemberitem
				FItemList(i).fmenu_id      = rsget("menu_id")
                FItemList(i).fpart_sn     = rsget("part_sn")
                FItemList(i).flevel_sn    = rsget("level_sn")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    End sub

	'//New 직원상세정보 가져오기	-- 2010.12 정윤정
	public Function fnGetMemberData
		IF Fempno="" and frectUserId="" THEN Exit Function
		Dim strSql

		strSql ="exec db_partner.dbo.sp_Ten_user_tenbyten_getData '"&Fempno&"','"& trim(frectUserId) &"'"

		'response.write strSql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			IF Not (rsget.EOF OR rsget.BOF) THEN
				Fempno			= rsget("empno")
				Fuserid         = rsget("userid")
				Ffrontid        = rsget("frontid")
				Fusername       = rsget("username")
				fuserNameEN       = rsget("userNameEN")
				FJuminno		= rsget("juminno")
				Fbirthday		= rsget("birthday")
				Fissolar		= rsget("issolar")
				Fsexflag		= rsget("sexflag")
				Fzipcode        = rsget("zipcode")
				Fzipaddr		= rsget("zipaddr")
				Fuseraddr       = rsget("useraddr")
				Fuserphone		= rsget("userphone")
				Fusercell       = rsget("usercell")
				FisIdentify     = rsget("isIdentify")
				Fusermail       = rsget("usermail")
				Fmsnmail        = rsget("msnmail")
				Fmessenger      = rsget("messenger")
				Fmywork			= rsget("mywork")
				Finterphoneno   = rsget("interphoneno")
				Fextension      = rsget("extension")
				Fdirect070      = rsget("direct070")
				Fjobdetail      = rsget("jobdetail")
				Fstatediv       = rsget("statediv")
				Fjoinday        = rsget("joinday")
				Fretireday      = rsget("retireday")
				Fuserimage      = rsget("userimage")
				Fpart_sn        = rsget("part_sn")
				Fposit_sn       = rsget("posit_sn")
				Fjob_sn         = rsget("job_sn")
				Flevel_sn       = rsget("level_sn")
				Fuserdiv        = rsget("userdiv")
				Fpart_name      = rsget("part_name")
				Fposit_name     = rsget("posit_name")
				Fjob_name       = rsget("job_name")
				Fmywork			= rsget("mywork")
				Frealjoinday	= rsget("realjoinday")
				Fretirereason	= rsget("retirereason")				'// 1-6 : 퇴사사유, 99 : 정규직전환
				Fbizsection_cd	= rsget("bizsection_cd")
				Fcriticinfouser = rsget("criticinfouser")           '' 0-일반 , 1-개인정보취급자
				Flv1customerYN = rsget("lv1customerYN")
				Flv2partnerYN = rsget("lv2partnerYN")
				Flv3InternalYN = rsget("lv3InternalYN")

				'// 부서NEW (2014-10-21, skyer9)
				Fdepartment_id 		= rsget("department_id")
				FdepartmentNameFull = rsget("departmentNameFull")
				Flevel_name			= rsget("level_name")
				Frank_sn			= rsget("rank_sn")
				Fpersonalmail			= rsget("personalmail")
				Fgsshopuserid			= rsget("gsshopuserid")

			END IF
		rsget.close
	End Function

	'//사번 가져오기
	public Function fnGetEmpNo
	 Dim strSql
	 strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getEmpNo('"&Fuserid&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEmpNo = rsget(0)
		END IF
		rsget.close
	End Function

	'//발령정보
	public Function fnGetUserModLog
		dim strSql
		strSql = "db_partner.dbo.usp_Ten_user_tenbyten_getLogList('"&Fempno&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
		 	fnGetUserModLog = rsget.getRows()
		END IF
		rsget.close
	End Function

		'//개인별 부서업무비율 리스트
	public Function fnGetUserBizSection
	 Dim strSql
	 strSql ="[db_partner].[dbo].sp_Ten_user_Bizsection_getList('"&Fempno&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
		 	fnGetUserBizSection = rsget.getRows()
		END IF
		rsget.close
	End Function


	'//개인별 부서업무비율 월별 데이터
	public Function fnGetUserBizSectionData
	Dim strSql
	strSql = "[db_partner].[dbo].[sp_Ten_user_Bizsection_getData]('"&Fempno&"','"&Fyyyymm&"')"
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
		 fnGetUserBizSectionData = rsget.getRows()
		END IF
	rsget.close
	End Function

	'//New 직원리스트
	public Function fnGetMemberList
		Dim strSql

		IF Fpart_sn = "" THEN Fpart_sn = 0
		IF Fposit_sn = "" THEN Fposit_sn = 0
		IF Fjob_sn = "" THEN Fjob_sn = 0
		if Flevel_sn="" then Flevel_sn=0

		dim tmpDepartmentID : tmpDepartmentID = Fdepartment_id
		if (FRectNoDepartOnly = "Y") then
			'// 부서 미지정
			tmpDepartmentID = -999
		end if

		strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getListCount("&Fpart_sn&","&Fposit_sn&","&Fjob_sn&",'"&Fstatediv&"','"&FSearchType&"','"&FSearchText&"', '" + CStr(FStartDate) + "', '" + CStr(FEndDate) + "', '" + CStr(tmpDepartmentID) + "', '" + CStr(Finc_subdepartment) + "', '" + CStr(FRectCriticInfoUser) + "','"& Frank_sn &"',"& Flevel_sn &",'"& Frectlv1customerYN &"','"& Frectlv2partnerYN &"','"& Frectlv3InternalYN &"')"

		'response.write strSql & "<br>"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getList("&Fpart_sn&","&Fposit_sn&","&Fjob_sn&",'"&Fstatediv&"','"&FSearchType&"','"&FSearchText&"','"&Forderby&"',"&FSPageNo&","&FEPageNo&", '" + CStr(FStartDate) + "', '" + CStr(FEndDate) + "', '" + CStr(tmpDepartmentID) + "', '" + CStr(Finc_subdepartment) + "', '" + CStr(FRectCriticInfoUser) + "','"& Frank_sn &"',"& Flevel_sn &",'"& Frectlv1customerYN &"','"& Frectlv2partnerYN &"','"& Frectlv3InternalYN &"')"

		'response.write strSql & "<br>"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetMemberList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	'//New 직원리스트
	public Function fnGetMemberList_csv
		Dim strSql

		IF Fpart_sn = "" THEN Fpart_sn = 0
		IF Fposit_sn = "" THEN Fposit_sn = 0
		IF Fjob_sn = "" THEN Fjob_sn = 0
		if Flevel_sn="" then Flevel_sn=0

		dim tmpDepartmentID : tmpDepartmentID = Fdepartment_id
		if (FRectNoDepartOnly = "Y") then
			'// 부서 미지정
			tmpDepartmentID = -999
		end if

		strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getList_csv("&Fpart_sn&","&Fposit_sn&","&Fjob_sn&",'"&Fstatediv&"','"&FSearchType&"','"&FSearchText&"','"&Forderby&"', '" + CStr(FStartDate) + "', '" + CStr(FEndDate) + "', '" + CStr(tmpDepartmentID) + "', '" + CStr(Finc_subdepartment) + "', '" + CStr(FRectCriticInfoUser) + "','"& Frank_sn &"',"& Flevel_sn &",'"& Frectlv1customerYN &"','"& Frectlv2partnerYN &"','"& Frectlv3InternalYN &"')"

		if session("ssBctId")="tozzinet" then
		response.write strSql & "<br>"
		end if
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetMemberList_csv = rsget.getRows()
		END IF
		rsget.close

	End Function

	public Function fnGetMaxChangePosit
	Dim strSql
		IF Fpart_sn = "" THEN Fpart_sn = 0
		IF Fposit_sn = "" THEN Fposit_sn = 0
		IF Fjob_sn = "" THEN Fjob_sn = 0

		dim tmpDepartmentID : tmpDepartmentID = Fdepartment_id
		if (FRectNoDepartOnly = "Y") then
			'// 부서 미지정
			tmpDepartmentID = -999
		end if

		strSql ="[db_partner].[dbo].usp_Ten_user_tenbyten_GetCntModLog("&Fpart_sn&","&Fposit_sn&","&Fjob_sn&",'"&Fstatediv&"','"&FSearchType&"','"&FSearchText&"', '" + CStr(FStartDate) + "', '" + CStr(FEndDate) + "', '" + CStr(tmpDepartmentID) + "', '" + CStr(Finc_subdepartment) + "', '" + CStr(FRectCriticInfoUser) + "')"

		 rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetMaxChangePosit = rsget(0)
		END IF
		rsget.close

	End Function

	'//계약사원 리스트
	public Function fnGetContractMemberList
	Dim strSql

		IF Fposit_sn = "" THEN Fposit_sn = 0
		IF Fjob_sn = "" THEN Fjob_sn = 0
		IF FchkDate = "" THEN FchkDate = 0

		strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getContractUserListCount("&Fposit_sn&","&Fjob_sn&",'"&Fstatediv&"','"&FSearchType&"','"&FSearchText&"',"&FchkDate&",'"&fshopid&"', '" + CStr(FStartDate) + "', '" + CStr(FEndDate) + "', '" + CStr(FMaxInoOnly) + "', '" + CStr(Fdepartment_id) + "', '" + CStr(Finc_subdepartment) + "')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getContractUserList("&Fposit_sn&","&Fjob_sn&",'"&Fstatediv&"','"&FSearchType&"','"&FSearchText&"',"&FchkDate&",'"&fshopid&"','"&Forderby&"',"&FSPageNo&","&FEPageNo&", '" + CStr(FStartDate) + "', '" + CStr(FEndDate) + "', '" + CStr(FMaxInoOnly) + "', '" + CStr(Fdepartment_id) + "', '" + CStr(Finc_subdepartment) + "')"
	'response.write strSql
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetContractMemberList = rsget.getRows()
		END IF
		rsget.close
		END IF

	End Function


	'//scm 이용 사원정보
	public Function fnGetScmMyInfo
	IF Fuserid = "" THEN Exit Function
		Dim strSql
		strSql ="db_partner.dbo.sp_Ten_partner_GetMyInfo ('"&Fuserid&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				Fempno		= rsget("empno")
				Fuserid            = rsget("id")
				Ffrontid            = rsget("frontid")
				Fusername       = rsget("username")
				FJuminno		= rsget("juminno")
				Fbirthday		= rsget("birthday")
				Fissolar		= rsget("issolar")
				Fsexflag		= rsget("sexflag")
				Fzipcode        = rsget("zipcode")
				Fzipaddr		= rsget("zipaddr")
				Fuseraddr       = rsget("useraddr")
				Fuserphone	= rsget("userphone")
				Fusercell           = rsget("usercell")
				Fusermail           = rsget("usermail")
				Fmsnmail            = rsget("msnmail")
				Finterphoneno       = rsget("interphoneno")
				Fextension          = rsget("extension")
				Fdirect070          = rsget("direct070")
				Fjobdetail          = rsget("jobdetail")
				Fstatediv           = rsget("statediv")
				Fjoinday            = rsget("joinday")
				Fretireday          = rsget("retireday")
				Fuserimage          = rsget("userimage")
				Fmywork				= rsget("mywork")
				Fpart_sn            = rsget("part_sn")
				Fposit_sn           = rsget("posit_sn")
				Fjob_sn             = rsget("job_sn")
				Flevel_sn           = rsget("level_sn")
				Fuserdiv            = rsget("userdiv")
				Fpart_name            = rsget("part_name")
				Fposit_name           = rsget("posit_name")
				Fjob_name             = rsget("job_name")
				Flevel_name          = rsget("level_name")
				FisIdentify			= rsget("isIdentify")
				FBizsection_cd	= rsget("bizsection_cd")

				Fpartnerusing       = rsget("partnerusing")
			END IF
		rsget.close
	End Function

	'//부서명 가져오기
	public Function fnGetPartName
		Dim strSql
		strSql ="db_partner.dbo.sp_Ten_partInfo_getName ("&Fpart_sn&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				Fpart_name		= rsget("part_name")
			END IF
		rsget.close
	End Function

		'//New 부서명 가져오기
	public Function fnGetDepartmentName
		Dim strSql
		strSql ="db_partner.dbo.[sp_Ten_department_getName] ("&Fdepartment_id&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetDepartmentName		= rsget("departmentname")
			END IF
		rsget.close
	End Function

	'//부서별 직원 리스트
	public Function fnGetPartUserList
	Dim strSql
	strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getPartList("&Fpart_sn&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPartUserList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'//아이디 배열로 이름, 직책 가져오기
	public Function fnGetInIDOutName
	IF Fuserid = "" THEN Exit Function
	Dim strSql
	strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getNameJob('"&Fuserid&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetInIDOutName = rsget.getRows()
		END IF
		rsget.close
	End Function

	'//직원 부서별 트리구조리스트 //김진영 작업분
    public Function fnGetUserTreeList
    Dim strSql
    	strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getTreeList('"&Fusername&"')"
    		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			fnGetUserTreeList = rsget.getRows()
    		END IF
    		rsget.close
    End Function

    '//직원 New부서별 트리구조리스트 //정윤정
     public Function fnGetUserTreeListNew
    Dim strSql
    	strSql ="[db_partner].[dbo].sp_Ten_user_tenbyten_getTreeList_New('"&Fusername&"')"
    		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			fnGetUserTreeListNew = rsget.getRows()
    		END IF
    		rsget.close
    End Function

    '//아이디로 new 부서번호, 부서명 가져오기
    public Function fnGetDepartmentInfo
    Dim strSql
    	strSql ="[db_partner].[dbo].[sp_Ten_user_tenbyten_getDepartmentName]('"&Fuserid&"')"
    		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			Fdepartment_id = rsget("department_id")
    			FdepartmentNameFull = rsget("departmentNameFull")
    			Fcid1 = rsget("cid1")
    			Fcid2 = rsget("cid2")
    			Fcid3 = rsget("cid3")
    			Fcid4 = rsget("cid4")
    		END IF
    		rsget.close
  	End Function

  	  '//아이디 또는 부서번호로 new 부서번호, 부서명, 상위부서번호 가져오기
    public Function fnGetDepartmentInfoPID
    Dim strSql
    	IF Fdepartment_id = "" THEN Fdepartment_id = 0
    	strSql ="[db_partner].[dbo].[sp_Ten_user_tenbyten_getDepartmentPID]('"&Fuserid&"',"&Fdepartment_id&")"
    		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			Fdepartment_id = rsget("cid")
    			FdepartmentNameFull = rsget("departmentNameFull")
    			Fcid1 = rsget("cid1")
    			Fcid2 = rsget("cid2")
    			Fcid3 = rsget("cid3")
    			Fcid4 = rsget("cid4")
    		END IF
    		rsget.close
  	End Function

    '//부서별 리스트 + 인원수
     public Function fnGetTeamPartList2017
    Dim strSql
    	strSql ="[db_partner].[dbo].[sp_Ten_user_tenbyten_getTeamPartList_2017]"
    		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			fnGetTeamPartList2017 = rsget.getRows()
    		END IF
    		rsget.close
    End Function
end Class



'/// 우편번호에서 주소 반환 함수 ///
public function GetZipAddress(zipcode)
	dim zip1, zip2, tmp, result
	dim sql

	tmp = Split(zipcode, "-")
	zip1 = tmp(0)
	zip2 = tmp(1)

	result = ""
	SQL =	"select top 1 (ADDR_SI + ' ' + ADDR_GU) as zipaddr " & vbCrlf
	SQL =	SQL & "	from [db_zipcode].[dbo].ADDR080TL " & vbCrlf
	SQL =	SQL & "	and ADDR_ZIP1 = '" & CStr(zip1) & "' " & vbCrlf
	SQL =	SQL & "	and ADDR_ZIP2 = '" & CStr(zip2) & "' " & vbCrlf
	rsget.Open SQL,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		result = rsget("zipaddr")
	end if
	rsget.Close

	GetZipAddress = result
end function


'/// 부서 코드로 부서명 생성 함수 ///
public function getPartNameByPartSN(psn)
    dim SQL
    SQL =	"Select part_name " & vbCRLF
	SQL = SQL & "From db_partner.dbo.tbl_partInfo " & vbCRLF
	SQL = SQL & "Where part_sn="&psn& vbCRLF

	rsget.Open SQL,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    getPartNameByPartSN = rsget("part_name")
	end if
	rsget.Close
end function

'/// 부서 옵션 생성 함수 ///
public function printPartOption(fnm, psn)
	dim SQL, i, strOpt
	if isnull(psn) then psn = ""
strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::부서선택::</option>"

	SQL =	"Select part_sn, part_name " &_
			"From db_partner.dbo.tbl_partInfo with (nolock)" &_
			"Where part_isDel='N' " &_
			"Order by part_sort"

	'response.write SQL & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

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


'/// 부서 옵션 생성 함수 ///
public function printPartOptionAddEtc(fnm, psn, strEtc)
	dim SQL, i, strOpt
	if isnull(psn) then psn = ""
strOpt =	"<select name='" & fnm & "' "&strEtc&">" &_
				"<option value=''>::부서선택::</option>"

	SQL =	"Select part_sn, part_name " &_
			"From db_partner.dbo.tbl_partInfo " &_
			"Where part_isDel='N' " &_
			"Order by part_sort"
	rsget.Open SQL,dbget,1

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
	printPartOptionAddEtc = strOpt
end function

'/// 직급 옵션 생성 함수 ///
public function printPositOption(fnm, psn)
	dim SQL, i, strOpt
if isnull(psn) then psn = ""
	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::선택::</option>"

	SQL =	"Select posit_sn, posit_name " &_
			"From db_partner.dbo.tbl_positInfo " &_
			"Where posit_isDel='N' "
	rsget.Open SQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("posit_sn") & "'"
			if Cstr(rsget("posit_sn"))=cstr(psn) then
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

'/// 직급 옵션 생성 함수 ///
public function printPositOptionIN90(fnm, psn)
	dim SQL, i, strOpt
if isnull(psn) then psn = ""
	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::직위선택::</option>"

	SQL =	"Select posit_sn, posit_name " &_
			"From db_partner.dbo.tbl_positInfo " &_
			"Where posit_isDel='N' "
	rsget.Open SQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("posit_sn") & "'"
			if Cstr(rsget("posit_sn"))=cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("posit_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close
		strOpt = strOpt & "<option value='99' "
		 IF cstr(psn) = "99" THEN
		strOpt = strOpt & "selected"
		 END IF
		 strOpt = strOpt & ">사원이상</option>"
	strOpt = strOpt & "</select>"

	'값 반환
	printPositOptionIN90 = strOpt
end function

'/// 직급 옵션 생성 함수 ///
public function printPositOptionPartTime(fnm, psn)
	dim SQL, i, strOpt
if isnull(psn) then psn = ""
	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::직급선택::</option>"

	SQL =	"Select posit_sn, posit_name " &_
			"From db_partner.dbo.tbl_positInfo " &_
			"Where posit_isDel='N' and posit_sn in (12,13,14,15) "
	rsget.Open SQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("posit_sn") & "'"
			if Cstr(rsget("posit_sn"))=cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("posit_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printPositOptionPartTime = strOpt
end function

'/// 직급 옵션 생성 함수 - 옵션만 ///
public function printPositOptionOnlyOption(psn)
	dim SQL, i, strOpt
	if isnull(psn) then psn = ""

	SQL =	"Select posit_sn, posit_name From db_partner.dbo.tbl_positInfo Where posit_isDel='N' and posit_sn not in('14','15') "
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("posit_sn") & "'"
			if Cstr(rsget("posit_sn"))=cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("posit_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	'값 반환
	printPositOptionOnlyOption = strOpt
end function

'/// 등급 옵션 생성 함수 ///
public function printLevelOption(fnm, psn)
	dim SQL, i, strOpt
if isnull(psn) then psn = ""
	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::등급선택::</option>"

	SQL =	"Select level_sn, level_name " &_
			"From db_partner.dbo.tbl_level " &_
			"Where level_isDel='N' " &_
			"Order by level_no"
	rsget.Open SQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("level_sn") & "'"
			if Cstr(rsget("level_sn"))=cstr(psn) then
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
if isnull(jsn) then jsn = ""
	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value=''>::직책선택::</option>"

	SQL =	"Select job_sn, job_name " &_
			"From db_partner.dbo.tbl_JobInfo " &_
			"Where job_isDel='N' "
	rsget.Open SQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("job_sn") & "'"
			if Cstr(rsget("job_sn"))=cstr(jsn) then
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

'/// 직책 옵션 생성 함수 - 옵션만 ///
public function printJobOptionOnlyOption(jsn)
	dim SQL, i, strOpt
	if isnull(jsn) then jsn = ""

	SQL =	"Select job_sn, job_name From db_partner.dbo.tbl_JobInfo Where job_isDel='N' "
	rsget.CursorLocation = adUseClient
	rsget.Open SQL,dbget,adOpenForwardOnly,adLockReadOnly

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("job_sn") & "'"
			if Cstr(rsget("job_sn"))=cstr(jsn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("job_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	'값 반환
	printJobOptionOnlyOption = strOpt
end function


'/// 담당샵 옵션 생성 함수 ///
public function printShopOption(fnm, shopid)
	dim SQL, i, strOpt
if isnull(shopid) then shopid = ""
	strOpt =	"<select name='" & fnm & "'>" &_
				"<option value='0'>::담당샵선택::</option>"

	SQL =	"select userid, shopname " &_
			"from [db_shop].[dbo].tbl_shop_user " &_
			"where 1 = 1 " &_
			"and isusing <> 'N' " &_
			"and shopdiv in ('1', '9') " &_
			"order by userid "
	rsget.Open SQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("userid") & "'"
			if rsget("userid")=cstr(shopid) then
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

public function GetYearDiff(Fjoinday)
	dim yyyy, mm, today

	today = now()
	if (IsNull(Fjoinday) or (Fjoinday = "")) then
		GetYearDiff = ""
		exit function
	end if

	yyyy = Year(today) - Year(Fjoinday)
	 if (Month(Fjoinday) > Month(today)) then
	 	yyyy = yyyy - 1
	 end if

	GetYearDiff = yyyy
end function

public function GetMonthDiff(Fjoinday)
	dim yyyy, mm, today
	dim realjoinday

	today = now()
	if (IsNull(Fjoinday) or (Fjoinday = "")) then
		GetMonthDiff = ""
		exit function
	end if

	mm = DateDiff("m", Fjoinday, today) mod 12
	'if (mm < 1) then
	'	mm = 1
	'end if

	GetMonthDiff = mm
end function

'// 사용자 권한과 메뉴등급을 일치시켜야 함
Sub DrawSelectBoxCriticInfoUser(selectedname, selectedId)
	dim tmp_str, query1
%>
<select class='select' name="<%= selectedname %>" >
    <option value='' <%if selectedId="" then response.write " selected"%> >선택</option>
	<option value='500' <%if selectedId="500" then response.write " selected"%> >LV1(개인정보)</option>
	<option value='100' <%if selectedId="100" then response.write " selected"%> >LV2(배송정보)</option>
	<option value='1' <%if selectedId="1" then response.write " selected"%> >LV3(주문정보)</option>
	<option value='200' <%if selectedId="200" then response.write " selected"%> >LV4(인사정보)</option>
	<option value='0' <%if selectedId="0" then response.write " selected"%> >권한없음</option>
</select>
<%
End Sub

'// 사용자 권한과 메뉴등급을 일치시켜야 함
Function GetCriticInfoUserLevelName(selectedId)
	Select Case selectedId
		Case "500"
			GetCriticInfoUserLevelName = "LV1(개인정보)"
		Case "100"
			GetCriticInfoUserLevelName = "LV2(배송정보)"
		Case "1"
			GetCriticInfoUserLevelName = "LV3(주문정보)"
		Case "200"
			GetCriticInfoUserLevelName = "LV4(인사정보)"
		Case "0"
			GetCriticInfoUserLevelName = "권한없음"
		Case Else
			GetCriticInfoUserLevelName = selectedId
	End Select
End Function

Function fnRankInfoSelectBox(selectedId)
	Dim vBody, sql
	vBody = "<option value=""0"" " & CHKIIF(CStr(selectedId)=CStr("0"),"selected","") & ">직급선택(없음)</option>" &  vbCrLf
	sql = "select rank_sn, rank_name from [db_partner].[dbo].[tbl_rankInfo] order by rank_sort asc"
	rsget.Open SQL,dbget,1
	if not rsget.eof then
		Do Until rsget.eof
			vBody = vBody & "<option value=""" & rsget("rank_sn") & """ " & CHKIIF(CStr(rsget("rank_sn"))=CStr(selectedId),"selected","") & ">" & rsget("rank_name") & "</option>" &  vbCrLf
			rsget.movenext
		Loop
	end if
	rsget.close
	fnRankInfoSelectBox = vBody
End Function

Function myDepartmentId(iuserid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 department_id FROM db_partner.dbo.tbl_user_tenbyten WHERE userid = '"&iuserid&"' "
	rsget.Open strSql,dbget,1
	if not rsget.eof then
		myDepartmentId = rsget("department_id")
	End If
	rsget.Close
End Function

Function fnContractWorkerCount(part_sn,job_sn,statediv,SearchType,SearchText,StartDate,EndDate,department_id,inc_subdepartment,CriticInfoUser)
	Dim strSql, cnt
	IF part_sn = "" THEN part_sn = 0
	IF job_sn = "" THEN job_sn = 0

	strSql ="[db_partner].[dbo].[sp_Ten_user_tenbyten_ContractWorkerCount] "&part_sn&","&job_sn&",'"&statediv&"','"&SearchType&"','"&SearchText&"', '" + CStr(StartDate) + "', '" + CStr(EndDate) + "', '" + CStr(department_id) + "', '" + CStr(inc_subdepartment) + "', '" + CStr(CriticInfoUser) + "'"
	'response.write strSql
	rsget.CursorLocation = adUseClient
	rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly

	If Not rsget.Eof Then
		cnt = rsget(0)
	End IF
	rsget.Close
	fnContractWorkerCount = cnt
End Function

Function fnTeamPartBossUserID(part_sn)
	Dim strSql, empno
	IF part_sn = "" THEN part_sn = 0
	IF job_sn = "" THEN job_sn = 0

	strSql = "select top 1 empno from db_partner.dbo.tbl_user_tenbyten as a "
	strSql = strSql & "where department_id = '" & part_sn & "' and isusing = '1' and statediv = 'Y' and job_sn > 0 order by a.job_sn asc"
	'response.write strSql
	rsget.CursorLocation = adUseClient
	rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly

	If Not rsget.Eof Then
		empno = rsget(0)
	End IF
	rsget.Close
	fnTeamPartBossUserID = empno
End Function

Sub sbOrganizationChartOne(empno)
	Dim strSql
	'strSql = "EXEC [db_partner].[dbo].sp_Ten_user_tenbyten_getList 0,0,0,'Y','1','"&userid&"','',1,1, '', '', '1', '', ''"
	strSql = "EXEC [db_partner].[dbo].sp_Ten_user_tenbyten_getList 0,0,0,'Y','3','"&empno&"','',1,1, '', '', '', '', ''"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly

	If Not rsget.Eof Then
		vStaffImage		= rsget(16)
		vStaffName		= rsget(1)
		vStaffID			= rsget(2)
		vStaffPartName	= Replace(rsget(27),"텐바이텐 - ","")
		vStaffPosit		= rsget(13)
		vStaffJob			= rsget(14)
		vStaffEmail		= rsget(8)
		vStaffHP			= rsget(17)
		vStaffPhone		= rsget(9)
		vStaffDirect		= rsget(11)
		vStaffExt			= rsget(10)
		vStaffMyWork		= rsget(20)
	End IF
	rsget.Close
End Sub

function checkValidPart(iuserid,ipart_sn)
    Dim SqlStr
    sqlStr = "select L.userid, L.part_sn, L.isDefault"
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_AddLevel L"
    sqlStr = sqlStr & " where L.UserID='"&iuserid&"'"
    sqlStr = sqlStr & " and L.part_sn="&ipart_sn&""

    rsget.Open sqlStr,dbget,1
    if  not rsget.EOF  then
		checkValidPart = rsget("part_sn")
	else
	    checkValidPart = -999
	end if
	rsget.Close

end function

function drawValidPartCombo(iuserid,compname,compval)
    dim oaddlevel, i, bufStr
    set oaddlevel = new CPartnerAddLevel
    oaddlevel.FRectUserID = iuserid
    oaddlevel.getUserAddLevelList
    bufStr = "<select name='"&compname&"'>"
    if (oaddlevel.FResultCount>0) then
        for i=0 to oaddlevel.FResultCount-1
             bufStr = bufStr&"<option value='"&oaddlevel.FItemList(i).Fpart_sn&"'"&chkIIF(CStr(compval)=CStr(oaddlevel.FItemList(i).Fpart_sn),"selected","")&">"&oaddlevel.FItemList(i).Fpart_name
        next
    End if
    bufStr = bufStr&"</select>"
    set oaddlevel = Nothing

    response.write bufStr
end function

%>