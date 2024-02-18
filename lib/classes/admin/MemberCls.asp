<%
Class CMemberItem
	public Fid
	public Fpassword
	public Fempno
	public Fusername
	public Fusermail
	public Fpart_name
	public Fposit_name
	public Fjob_name
	public Flevel_name
	public Fpart_sn
	public Fposit_sn
	public Fjob_sn
	public Flevel_sn
	public Fuserdiv
	public FisUsing
	public FAddLevelCnt
    public Fcriticinfouser

    public function getPartnerUserDivName()
        select CASE Fuserdiv
            CASE "9999" : getPartnerUserDivName="업체"
            CASE "999"  : getPartnerUserDivName="제휴사"
            CASE "9000" : getPartnerUserDivName="강사"
            CASE "9"    : getPartnerUserDivName="관리자"
            CASE "7"    : getPartnerUserDivName="마스타"
            CASE "5"    : getPartnerUserDivName="LV4"
            CASE "4"    : getPartnerUserDivName="LV3"
            CASE "2"    : getPartnerUserDivName="LV2"
            CASE "1"    : getPartnerUserDivName="LV1"
            CASE "500"  : getPartnerUserDivName="매장공통"
            CASE "501"  : getPartnerUserDivName="직영매장"
            CASE "502"  : getPartnerUserDivName="수수료매장"
            CASE "503"  : getPartnerUserDivName="대리점"
            CASE "101"  : getPartnerUserDivName="오프샾"
            CASE "111"  : getPartnerUserDivName="오프샾점장"
            CASE "112"  : getPartnerUserDivName="오프샾부점장"
            CASE "509"  : getPartnerUserDivName="오프매출조회"
            CASE "201"  : getPartnerUserDivName="Zoom"
            CASE "301"  : getPartnerUserDivName="College"
            CASE ELSE : getPartnerUserDivName="?"
        END select
    end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class


Class CMember
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectId
	public FRectsearchKey
	public FRectsearchString
	public FRectisUsing
	public FRectpart_sn
	public FRectuserdiv
	public FRectLevelsn
    public FRectCriticinfouser

    public FRectPositsn
    public FRectJobsn

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

	'##### 사용자 목록 접수 ##### //업체 제외
	public Sub GetMemberList()
		dim SQL, AddSQL, i, strTemp

		'// 검색어 쿼리 //
		if FRectsearchKey<>"" and FRectsearchString<>"" then
			AddSQL = AddSQL & " and u1." & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectisUsing<>"" then
			AddSQL = AddSQL & " and t1.isUsing = '" & isUsing & "' "
		end if

		if FRectpart_sn<>"" then
			AddSQL = AddSQL & " and u1.part_sn = " & FRectpart_sn & " "
		end if

        if FRectuserdiv<>"" then
            if (FRectuserdiv="T") then
                AddSQL = AddSQL & " and t1.userdiv <=10 "
            elseif (FRectuserdiv="L") then
                AddSQL = AddSQL & " and t1.userdiv <=5 "
            else
    			AddSQL = AddSQL & " and t1.userdiv = '" & FRectuserdiv & "' "
    		end if
		end if

		if FRectLevelsn<>"" then
		    AddSQL = AddSQL & " and t1.level_sn = '" & FRectLevelsn & "' "
		end if

		if FRectCriticinfouser<>"" then
		    AddSQL = AddSQL & " and u1.criticinfouser>0 "
		end if

		if FRectPositsn<>"" then
		    AddSQL = AddSQL & " and u1.posit_sn = '" & FRectPositsn & "' "
		end if

		if FRectJobsn<>"" then
		    AddSQL = AddSQL & " and u1.job_sn = '" & FRectJobsn & "' "
		end if

		'// 개수 파악 //
		SQL =	"Select count(id), CEILING(CAST(Count(id) AS FLOAT)/" & FPageSize & ") "
		SQL = SQL & "From db_partner.[dbo].tbl_partner as t1 "
		SQL = SQL & "	left outer  join db_partner.dbo.tbl_user_tenbyten as u1 "
		SQL = SQL & "		on t1.id = u1.userid and u1.statediv ='Y' and u1.isusing=1 "
		SQL = SQL & "where t1.id<>'' and t1.userdiv<500" & AddSQL	''' 999=>500
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'// 목록 접수 //
		SQL =	    "select top " & CStr(FPageSize*FCurrPage)
		SQL = SQL & "	t1.id, t1.password, u1.empno,u1.username ,u1.usermail , t1.company_name, t1.email "
		SQL = SQL & "	,t2.part_name "
		SQL = SQL & "	,t3.posit_name "
		SQL = SQL & "	,t4.level_name "
		SQL = SQL & "	,t5.job_name "
		SQL = SQL & "	,t1.userdiv, t1.isUsing "
		SQL = SQL & "	,u1.empno, u1.criticinfouser"
		SQL = SQL & "	,(select count(*) from db_partner.dbo.tbl_partner_AddLevel L where t1.id=L.userid and L.isDefault='N') as AddLevelCnt "
		SQL = SQL & " from db_partner.[dbo].tbl_partner as t1 "
		SQL = SQL & "	left outer join db_partner.dbo.tbl_user_tenbyten as u1 "
		SQL = SQL & "		on t1.id = u1.userid and u1.statediv ='Y' and u1.isusing=1 "
		SQL = SQL & "	left join db_partner.dbo.tbl_partInfo as t2 "
		SQL = SQL & "		on u1.part_sn=t2.part_sn "
		SQL = SQL & "	left join db_partner.dbo.tbl_positInfo as t3 "
		SQL = SQL & "		on u1.posit_sn=t3.posit_sn "
		SQL = SQL & "	left join db_partner.dbo.tbl_level as t4 "
		SQL = SQL & "		on t1.level_sn=t4.level_sn "
	    SQL = SQL & "	left join db_partner.dbo.tbl_jobInfo as t5 "
		SQL = SQL & "		on t1.job_sn=t5.job_sn "
		SQL = SQL & "where t1.id<>'' and t1.userdiv<500" & AddSQL
		SQL = SQL & "Order by t1.id "

		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMemberItem

				FItemList(i).Fid			= rsget("id")
				FItemList(i).Fpassword	= rsget("password")
				FItemList(i).Fempno		= rsget("empno")
				'if rsget("userdiv") <= 9 then 2011.05.09 정윤정 수정
				if not isNull(rsget("username")) then
				FItemList(i).Fusername	= rsget("username")
				FItemList(i).Fusermail	= rsget("usermail")
				else
				FItemList(i).Fusername	= rsget("company_name")
				FItemList(i).Fusermail	= rsget("email")
				end if
				FItemList(i).Fpart_name	= rsget("part_name")
				FItemList(i).Fposit_name	= rsget("posit_name")
				FItemList(i).Fjob_name      = rsget("job_name")
				FItemList(i).Flevel_name	= rsget("level_name")
				FItemList(i).Fuserdiv	= rsget("userdiv")
				FItemList(i).FisUsing	= rsget("isUsing")

                FItemList(i).FAddLevelCnt = rsget("AddLevelCnt")
                FItemList(i).Fempno = rsget("empno")
                FItemList(i).Fcriticinfouser = rsget("criticinfouser")

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

		'// 목록 접수 //
		SQL =	"Select " &_
				"	t1.id, t1.password, u1.empno, u1.username ,u1.usermail " &_
				"	,u1.part_sn, u1.posit_sn, u1.job_sn, t1.level_sn " &_
				"	,t1.userdiv, t1.isUsing " &_
				"from db_partner.[dbo].tbl_partner as t1 " &_
				"	left outer join db_partner.dbo.tbl_user_tenbyten as u1 on t1.id = u1.userid and u1.statediv ='Y' and u1.isusing =1 "&_
				"Where t1.id='" & FRectId & "'"
		rsget.Open SQL,dbget,1

		if Not(rsget.EOF or rsget.BOF) then

			FResultCount = 1
			redim preserve FItemList(1)
			set FItemList(1) = new CMemberItem

			FItemList(1).Fpassword		= rsget("password")
			FItemList(1).Fempno		= rsget("empno")
			FItemList(1).Fusername		= rsget("username")
			FItemList(1).Fusermail		= rsget("usermail")
			FItemList(1).Fpart_sn		= rsget("part_sn")
			FItemList(1).Fposit_sn		= rsget("posit_sn")
			FItemList(1).Fjob_sn		= rsget("job_sn")
			FItemList(1).Flevel_sn		= rsget("level_sn")
			FItemList(1).Fuserdiv		= rsget("userdiv")
			FItemList(1).FisUsing		= rsget("isUsing")
		else
			FResultCount = 0
		end if

		rsget.Close

	end Sub
end Class

'/// 부서 옵션 생성 함수 ///
public function printPartOption(fnm, psn)
	dim SQL, i, strOpt

	strOpt =	"<select class='select' name='" & fnm & "'>" &_
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
	printPartOption = strOpt
end function

'/// 직급 옵션 생성 함수 ///
public function printPositOption(fnm, psn)
	dim SQL, i, strOpt

	strOpt =	"<select class='select' name='" & fnm & "'>" &_
				"<option value=''>::직급선택::</option>"

	SQL =	"Select posit_sn, posit_name " &_
			"From db_partner.dbo.tbl_positInfo " &_
			"Where posit_isDel='N' "
	rsget.Open SQL,dbget,1

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

	strOpt =	"<select class='select' name='" & fnm & "'>" &_
				"<option value=''>::등급선택::</option>"

	SQL =	"Select level_sn, level_name " &_
			"From db_partner.dbo.tbl_level " &_
			"Where level_isDel='N' " &_
			"Order by level_no"
	rsget.Open SQL,dbget,1

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



'/////// 탈퇴회원 정보 검색 ////////
Class CwithDrawItem
	public Fuid
	public Fjumin1
	public Fregdate
	public FcomplainDiv
	public FcomplainText

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class


Class CwithDraw
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectUserId
	public FRectChkInit
	public FRectChkCmt
	public FRectCplDiv

	Private Sub Class_Initialize()
		redim  FitemList(0)

		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'##### 목록 접수 #####
	public Sub GetUserList()
		dim SQL, AddSQL, i, strTemp

		'// 검색 쿼리 //
		if FRectUserId<>"" then
			if FRectChkInit="on" then
				AddSQL = " Where userid like '" & FRectUserId & "%' "
			else
				AddSQL = " Where userid='" & FRectUserId & "' "
			end if
		end if

		if FRectChkCmt<>"" then
			AddSQL = AddSQL & chkIIF(AddSQL="","Where ","and ")
			AddSQL = AddSQL & " len(Replace(Replace(Replace(convert(varchar(10),complaintext),'.',''),' ',''),CHAR(13)+CHAR(10),''))>1 "
		end if

		if FRectCplDiv<>"" then
			AddSQL = AddSQL & chkIIF(AddSQL="","Where ","and ")
			if FRectCplDiv="not" then
				AddSQL = AddSQL & " complaindiv='' "
			else
				AddSQL = AddSQL & " complaindiv='" & FRectCplDiv & "' "
			end if
		end if

		'// 개수 파악 //
		SQL =	"Select count(id), CEILING(CAST(Count(id) AS FLOAT)/" & FPageSize & ") " &_
				"From db_user.[dbo].tbl_deluser " & AddSQL
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'// 목록 접수 //
		SQL =	"select top " & CStr(FPageSize*FCurrPage) &_
				"	userid, left(juminno,6) as juminno " &_
				"	,Case complaindiv " &_
				"		When '01' Then '상품품질불만' " &_
				"		When '02' Then '이용빈도낮음' " &_
				"		When '03' Then '배송지연' " &_
				"		When '04' Then '개인정보유출우려' " &_
				"		When '05' Then '교환/환불/품질불만' " &_
				"		When '06' Then '기타' " &_
				"		When '07' Then 'A/S불만' " &_
				"		Else '미지정' " &_
				"	End as complaindiv " &_
				"	,regdate, complaintext " &_
				"From db_user.[dbo].tbl_deluser " & AddSQL &_
				"Order by id desc "
		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CwithDrawItem

				FItemList(i).Fuid			= rsget("userid")
				FItemList(i).Fjumin1		= rsget("juminno")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FcomplainDiv	= rsget("complaindiv")
				FItemList(i).FcomplainText	= rsget("complaintext")

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
end Class

Class CLoginIPItem
	public Fidx
	public Fipaddress
	public Fdepartment_id
	public FdepartmentnameFull
	public Fuserid
	public Fmanagername
	public Fcomment
	public Fusescmyn
	public Fuselogicsyn
	public Fusecustomerinfoyn
	public Freguserid
	public Fmodiuserid
	public Fuseyn
	public Fregdate
	public Flastupdate

	Private Sub Class_Initialize()
		''
	End Sub

	Private Sub Class_Terminate()
		''
	End Sub
end Class

Class CLoginIP
	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectIdx
	public FRectUserId
	public FRectIPAddress
	public FRectDepartment_id
	public FRectSearchRect
	public FRectSearchStr
	public FRectuseyn

	Private Sub Class_Initialize()
		redim  FitemList(0)

		FCurrPage 		= 1
		FPageSize 		= 20
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
	End Sub

	Private Sub Class_Terminate()
		''
	End Sub

	public Sub GetIPList()
		dim SQL, AddSQL, i

		AddSQL = " where 1 = 1"
		if FRectDepartment_id <> "" then
			AddSQL = AddSQL + " and i.department_id = '" & FRectDepartment_id & "' "
		end if

		if FRectSearchRect <> "" and FRectSearchStr <> "" then
			AddSQL = AddSQL + " and i." & FRectSearchRect & " = '" & FRectSearchStr & "' "
		end if

		if FRectuseyn <> "" then
			AddSQL = AddSQL & " and i.useyn='"& FRectuseyn &"'"
		end if

		SQL =	"Select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") " &_
				"From db_partner.dbo.tbl_user_loginIP i " & AddSQL
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		SQL =	"select top " & CStr(FPageSize*FCurrPage) &_
				"	idx, ipaddress, department_id, userid, managername, comment, usescmyn " &_
				"	, uselogicsyn, usecustomerinfoyn, reguserid, modiuserid, i.useyn, i.regdate, i.lastupdate, d.departmentnameFull " &_
				"From db_partner.dbo.tbl_user_loginIP i " &_
				" left join db_partner.dbo.vw_user_department as d on i.department_id = d.cid "&_
				AddSQL &_

				"Order by idx desc "
		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLoginIPItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fipaddress			= rsget("ipaddress")
				FItemList(i).Fdepartment_id		= rsget("department_id")
				FItemList(i).FdepartmentnameFull		= rsget("departmentnameFull")
				FItemList(i).Fuserid			= db2html(rsget("userid"))
				FItemList(i).Fmanagername		= db2html(rsget("managername"))
				FItemList(i).Fcomment			= db2html(rsget("comment"))
				FItemList(i).Fusescmyn			= rsget("usescmyn")
				FItemList(i).Fuselogicsyn		= rsget("uselogicsyn")
				FItemList(i).Fusecustomerinfoyn	= rsget("usecustomerinfoyn")
				FItemList(i).Freguserid			= rsget("reguserid")
				FItemList(i).Fmodiuserid		= rsget("modiuserid")
				FItemList(i).Fuseyn				= rsget("useyn")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close
	end sub

	public Sub GetIPOne()
		dim SQL, AddSQL, i

		AddSQL = " where 1 = 1"
		if (FRectIdx <> "") then
			AddSQL = AddSQL + " and idx = " & FRectIdx
		end if

		SQL =	"select top 1 " &_
				"	idx, ipaddress, department_id, userid, managername, comment, usescmyn " &_
				"	, uselogicsyn, usecustomerinfoyn, reguserid, modiuserid, i.useyn, i.regdate, i.lastupdate, d.departmentnameFull " &_
				"From db_partner.dbo.tbl_user_loginIP i " &_
				" left join db_partner.dbo.vw_user_department as d on i.department_id = d.cid "&_
				AddSQL &_
				"Order by idx desc "
		rsget.Open SQL,dbget,1

		Set FOneItem = new CLoginIPItem
		if Not(rsget.EOF or rsget.BOF) then
			FOneItem.Fidx				= rsget("idx")
			FOneItem.Fipaddress			= rsget("ipaddress")
			FOneItem.Fdepartment_id		= rsget("department_id")
			FOneItem.FdepartmentnameFull		= rsget("departmentnameFull")
			FOneItem.Fuserid			= db2html(rsget("userid"))
			FOneItem.Fmanagername		= db2html(rsget("managername"))
			FOneItem.Fcomment			= db2html(rsget("comment"))
			FOneItem.Fusescmyn			= rsget("usescmyn")
			FOneItem.Fuselogicsyn		= rsget("uselogicsyn")
			FOneItem.Fusecustomerinfoyn	= rsget("usecustomerinfoyn")
			FOneItem.Freguserid			= rsget("reguserid")
			FOneItem.Fmodiuserid		= rsget("modiuserid")
			FOneItem.Fuseyn				= rsget("useyn")
			FOneItem.Fregdate			= rsget("regdate")
			FOneItem.Flastupdate		= rsget("lastupdate")
		end if
		rsget.Close
	end sub

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

Class CUserNotificationItem
	public fUserid
	public fUsername
	public fEmpno
	public fUserCount
	public fIsusing
	public fstatediv
	public fidx
	public fnotificationType
	public fregdate
	public flastupdate
	public freguserid
	public flastuserid
	public fnotificationTypeName

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CUserNotification
	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectIdx
	public FRectUserId
	public FRectDepartment_id
	public FRectSearchRect
	public FRectSearchStr
	public fRectIsusing
	public fRectstatediv

	Private Sub Class_Initialize()
		redim  FitemList(0)

		FCurrPage 		= 1
		FPageSize 		= 20
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	' /admin/member/notification/userList.asp
	public Sub GetUserList()
		dim sqlStr, sqlsearch, i

		if FRectDepartment_id <> "" then
			sqlsearch = sqlsearch & " and ut.department_id = '" & FRectDepartment_id & "'"
		end if
		if FRectSearchRect <> "" and FRectSearchStr <> "" then
			sqlsearch = sqlsearch & " and ut." & FRectSearchRect & " = '" & FRectSearchStr & "' "
		end if
		if fRectIsusing <> "" then
			if fRectIsusing = "Y" then
				sqlsearch = sqlsearch & " and ut.isusing=1"
			else
				sqlsearch = sqlsearch & " and ut.isusing=0"
			end if
		end if
		if fRectstatediv <> "" then
			sqlsearch = sqlsearch & " and ut.statediv='"& fRectstatediv &"'"
		end if

		sqlStr = "Select count(ut.empno), CEILING(CAST(Count(ut.empno) AS FLOAT)/" & FPageSize & ")"
		sqlStr = sqlStr & " FROM db_partner.dbo.tbl_user_tenbyten ut with (nolock)"
		sqlStr = sqlStr & " WHERE ut.isusing = 1 " & sqlsearch
		sqlStr = sqlStr & " and ut.part_sn in (7,30)"		' 개발팀

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		if FTotalCount < 1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = "select top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " ut.userid ,ut.username ,ut.empno, ut.isusing, ut.statediv"
		sqlStr = sqlStr & " ,(select count(*) from db_partner.dbo.notificationUser as nu with (nolock)"
		sqlStr = sqlStr & " 	where nu.userid = ut.userid and nu.isusing='Y') as userCount"
		sqlStr = sqlStr & " FROM db_partner.dbo.tbl_user_tenbyten ut with (nolock)"
		sqlStr = sqlStr & " WHERE ut.isusing = 1 " & sqlsearch
		sqlStr = sqlStr & " and ut.part_sn in (7,30)"		' 개발팀
		sqlStr = sqlStr & " order by ut.statediv asc, ut.posit_sn asc, ut.job_sn asc, ut.empno asc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if FResultCount<1 then FResultCount=0
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CUserNotificationItem

				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fusername = rsget("username")
				FItemList(i).fempno = rsget("empno")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fstatediv = rsget("statediv")
				FItemList(i).fuserCount = rsget("userCount")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close
	end sub

	'/admin/member/notification/NotificationUser.asp?userId=coolhas&menupos=9216
	public Sub GetNotificationUserList()
	    dim sqlStr, i , sqlsearch
	    
	    if frectuserid <> "" then
	    	sqlsearch = sqlsearch & " and nu.userid = '"&frectuserid&"'"
	    end if
	    
		sqlStr = "Select count(nu.idx), CEILING(CAST(Count(nu.idx) AS FLOAT)/" & FPageSize & ")"
		sqlStr = sqlStr & " FROM db_partner.dbo.notificationUser nu with (nolock)"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten ut with (nolock)"
		sqlStr = sqlStr & " 	on nu.userid=ut.userid"
		sqlStr = sqlStr & " left join db_partner.dbo.notificationType nt with (nolock)"
		sqlStr = sqlStr & " 	on nu.notificationType = nt.notificationType"
		sqlStr = sqlStr & " WHERE nu.isusing = 'Y' " & sqlsearch

	    'response.write sqlStr &"<Br>"	        
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		if FTotalCount < 1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = "select top " & CStr(FPageSize*FCurrpage)
		sqlStr = sqlStr & " nu.idx, nu.userid, nu.notificationType, nu.isusing, nu.regdate, nu.lastupdate, nu.reguserid, nu.lastuserid"
		sqlStr = sqlStr & " , ut.userid ,ut.username ,ut.empno, ut.isusing, ut.statediv, nt.notificationTypeName"
		sqlStr = sqlStr & " FROM db_partner.dbo.notificationUser nu with (nolock)"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten ut with (nolock)"
		sqlStr = sqlStr & " 	on nu.userid=ut.userid"
		sqlStr = sqlStr & " left join db_partner.dbo.notificationType nt with (nolock)"
		sqlStr = sqlStr & " 	on nu.notificationType = nt.notificationType"
		sqlStr = sqlStr & " WHERE nu.isusing = 'Y' " & sqlsearch
		sqlStr = sqlStr & " order by nu.idx desc"
	    
	    'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if FResultCount<1 then FResultCount=0
        
		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CUserNotificationItem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fnotificationType = rsget("notificationType")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastupdate = rsget("lastupdate")
				FItemList(i).freguserid = rsget("reguserid")
				FItemList(i).flastuserid = rsget("lastuserid")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fusername = rsget("username")
				FItemList(i).fempno = rsget("empno")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fstatediv = rsget("statediv")
				FItemList(i).fnotificationTypeName = rsget("notificationTypeName")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
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
end Class

%>
