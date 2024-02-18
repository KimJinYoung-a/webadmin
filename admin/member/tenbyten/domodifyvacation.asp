<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 휴가관리
' History : 2011.01.19 정윤정 생성
'			2022.09.21 한용민 수정(오류수정)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<html>
<head>
<title>연차(휴가) 등록</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body>
<%

dim login_empno
login_empno = session("ssBctSn")

dim login_userid
login_userid = session("ssBctId")

'//approverempno
'//registerempno


dim mode
dim userid, divcd, startday, endday, totalvacationday, ishalfvacation, vHalfGubun
dim empno,posit_sn
Dim authstate
dim totalday
dim masteridx
dim detailidx
dim insDivcode

dim promotionDay, jungsanDay, retireJungsanDay, comment

dim registerid, registerempno
dim approverid, approverempno

dim oVacation

dim i, sql
Dim sReturnUrl

mode 				= requestCheckvar(request("mode"),32)
userid 				= requestCheckvar(request("userid"),32)
empno 				= requestCheckvar(request("empno"),32)
divcd 				= requestCheckvar(request("divcd"),1)
startday 			= requestCheckvar(request("startday"),10)
endday 				= requestCheckvar(request("endday"),10)
totalvacationday	= requestCheckvar(request("totalvacationday"),32)
ishalfvacation		= requestCheckvar(request("ishalfvacation"),1)
vHalfGubun 			= requestCheckvar(request("halfgubun"),2)
totalday 			= requestCheckvar(request("totalday"),32)
masteridx 			= requestCheckvar(Request("masteridx"),32)
detailidx 			= requestCheckvar(Request("detailidx"),32)
insDivcode 			= requestCheckvar(Request("insDivcode"),8)

promotionDay 		= requestCheckvar(Request("promotionDay"),16)
jungsanDay 			= requestCheckvar(Request("jungsanDay"),16)
retireJungsanDay 	= requestCheckvar(Request("retireJungsanDay"),16)

authstate 			= requestCheckvar(Request("ias"),10)
sReturnUrl 			= requestCheckvar(request("hidRU"),100)

posit_sn			= requestCheckvar(request("posit_sn"),4)
comment				= html2db(request("comment"))

if authstate = "5" then mode = "denydetail"  '전자결재를 통해서 반려처리된 경우 추가 (2011.05.12 정윤정)
userid = Replace(userid, " ", "")

dim oMember


if (mode = "chkemploytype") then
	Set oMember = new CTenByTenMember
 
	if (userid <> "") then
		oMember.Fuserid = userid
		oMember.fnGetScmMyInfo
	elseif (empno <> "") then
		oMember.Fempno = empno
		oMember.fnGetMemberData
	else
		response.write "<script type='text/javascript'>alert('잘못된 접근입니다.');</script>"
		response.end
	end if

	if (isNull(oMember.Fempno) or oMember.Fempno="") then
		response.write "<script type='text/javascript'>alert('잘못된 사원번호 입니다.');</script>"
		response.end
	end if

	if (Left(oMember.Fempno, 2) = "90") then
		response.write "<script type='text/javascript'>parent.ReActEmployType(2, '" & oMember.Fempno & "', '"& oMember.Fuserid &"','"& oMember.Fposit_sn &"')</script>"
	else
		response.write "<script type='text/javascript'>parent.ReActEmployType(1, '" & oMember.Fempno & "', '"& oMember.Fuserid &"','"& oMember.Fposit_sn &"')</script>"
	end if

elseif(mode ="calYV") then
	dim icalyv
	icalyv = 0
	sql = " select    Ceiling(1.0*sum(d.wholidaytime)/(select count(wholidaytime) from db_partner.dbo.tbl_user_dailypay where empno = u.empno and left(yyyymmdd,7) = m.yyyymm and wholidaytime > 0) ) "&_
		  	" from  db_partner.dbo.tbl_user_tenbyten as u "&_
		  	"		inner join db_partner.dbo.tbl_user_monthlypay as m 	on u.empno= m.empno  "&_
		  	"		inner join db_partner.dbo.tbl_user_dailypay as d on m.empno = d.empno and left(d.yyyymmdd,7) = m.yyyymm	" &_
		  	" where  u.isusing = 1 "&_
			"	and (u.statediv ='Y' or (u.statediv ='N' and datediff(dd,u.retireday,getdate())<=0))	"	&_							 
			"	and u.posit_sn =13		" &_
			"	and m.paystate >= 5  and u.empno = '"&empno&"'"&_
			"	and m.yyyymm = convert(varchar(7),dateadd(m,-1,getdate()),121) "&_
			"	 group by u.empno, m.yyyymm	having  sum(d.wholidaytime) > 0 "  
	rsget.Open sql,dbget,1		 
	if not (rsget.EOF OR rsget.BOF) then
		icalyv = rsget(0) 
	end if
	rsget.close		
	 
	if cint(icalyv) > 0 then icalyv =  Cint(icalyv)/60
%>
	<script type="text/javascript"> 
		parent.document.frm.totalvacationday.value = "<%=icalyv%>";
	</script>
<%	
			
 			
elseif (mode = "add") then
	Set oMember = new CTenByTenMember

	if (userid <> "") then
		oMember.Fuserid = userid
		oMember.fnGetScmMyInfo
	else
		oMember.Fempno = empno
		oMember.fnGetMemberData
	end if

	if (isNull(oMember.Fempno) or oMember.Fempno="") then
		response.write "<script>alert('잘못된 아이디 또는 사번입니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	registerid = session("ssBctId")
	if posit_sn = "13" then
		totalvacationday = totalvacationday*0.125
	end if
	sql = "insert into [db_partner].[dbo].tbl_vacation_master(empno, userid, divcd, startday, endday, totalvacationday, registerid) " & vbCrlf
	sql = sql & "values('"&oMember.Fempno&"','" & oMember.Fuserid & "', '" & divcd & "', '" & startday & " 00:00:01', '" & endday & " 23:59:59', " & CStr(totalvacationday) & ", '" & registerid & "') " & vbCrlf
	dbget.Execute(sql)
	
	if (divcd = "1" or divcd = "7") and Left(oMember.Fempno, 2) = "90" then '계약직 연차일때 (또는 휴일대체)
		if posit_sn = "13" then
		totalvacationday = (totalvacationday/0.125)*60
		end if
		sql = "insert into db_partner.dbo.tbl_vacation_month(empno,posit_sn,yyyymm,yearvacationday, adminid) "& vbCrlf
		sql = sql & " values ('"&oMember.Fempno&"','"&oMember.Fposit_sn&"','"&left(date(),7)&"' ,'" & CStr(totalvacationday)  & "','"&registerid&"') " & vbCrlf
		 dbget.Execute(sql)
  end if

	response.write "<script>alert('저장 되었습니다.');</script>"
	response.write "<script>opener.location.reload(); opener.focus();window.close()</script>"
elseif (mode = "addallyearvacation") then

	'// 사용안함
	response.end

'당해입사 반올림 ( 12일 * (당해 일할 일수 / 365) )				TODO : 올림으로 변경필요!!
'// >>>>>> XXXXXXXXXXXXXXXX'-- 2년차까지 12일, 3년차 이상은 15일
'// 입사 익년부터 15일
'-- 임시직 제외
'// XXXXXXXXXXXXXXXXXXXXXXX'-- 이상구 제외
'1-6월 입사자는 연차계산에서 1년을 더해준다. -->> 변경 : 입사일자 기준
	sql = "insert into [db_partner].[dbo].tbl_vacation_master(empno, userid, divcd, startday, endday, totalvacationday, registerid) " & vbCrlf
	sql = sql & "select " & vbCrlf
	sql = sql & "	T.empno, T.userid " & vbCrlf
	sql = sql & "	, '1' " & vbCrlf
	sql = sql & "	, '' + Cast(Year(getdate()) as varchar) + '-01-01 00:00:01' " & vbCrlf
	sql = sql & "	, '' + Cast((Year(getdate()) + 1) as varchar) + '-03-31 23:59:59' " & vbCrlf
	sql = sql & "	, ( " & vbCrlf
	sql = sql & "		case " & vbCrlf
	sql = sql & "			when (Year(getdate()) = Year(joinday)) then Round(12 * ( DateDiff(D, joinday,  Cast(Year(joinday) as varchar) + '-12-31') / 365), 0) " & vbCrlf
	''sql = sql & "			when T.yeardiff >= 3 then 15 " & vbCrlf
	sql = sql & "			else 15 " & vbCrlf
	sql = sql & "		end " & vbCrlf
	sql = sql & "	) as regularvacation " & vbCrlf
	sql = sql & "	,'system' " & vbCrlf
	sql = sql & "from ( " & vbCrlf
	sql = sql & "	select " & vbCrlf
	sql = sql & "		t.empno,userid " & vbCrlf
	sql = sql & "		, joinday " & vbCrlf
	''sql = sql & "		, ( " & vbCrlf
	''sql = sql & "			case " & vbCrlf
	''sql = sql & "				when (Month(joinday) <= 6) then (Year(getdate()) - Year(joinday)) + 1 " & vbCrlf
	''sql = sql & "				else (Year(getdate()) - Year(joinday)) " & vbCrlf
	''sql = sql & "			end " & vbCrlf
	''sql = sql & "		) as yeardiff " & vbCrlf
	sql = sql & "	from [db_partner].[dbo].tbl_user_tenbyten t, db_partner.[dbo].tbl_partner p " & vbCrlf
	sql = sql & "	where  t.userid = p.id " & vbCrlf
	sql = sql & "	and t.isusing = 1 " & vbCrlf
	sql = sql & "	and p.isusing = 'Y' " & vbCrlf
	sql = sql & "	and p.userdiv < 999 " & vbCrlf
	sql = sql & "	and p.level_sn < 999 " & vbCrlf
	sql = sql & "	and t.part_sn <> 4 " & vbCrlf
	sql = sql & "	and t.part_sn <> 5 " & vbCrlf
	sql = sql & "	and t.posit_sn < 12 " & vbCrlf
	'sql = sql & "	and Year(joinday) < Year(getdate()) " & vbCrlf
	'sql = sql & "	and Month(joinday) <= 6 " & vbCrlf
	sql = sql & ") as T " & vbCrlf

	'// 시스템팀:7, 30
	if (session("ssAdminPsn") = 7) or (session("ssAdminPsn") = 30) then
		dbget.Execute(sql)
		response.write "<script>alert('저장 되었습니다.');</script>"
	end if

	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
elseif (mode = "addallyearvacationNew") then

	'// 전체연차등록(정규직, 계약직)
	sql = " exec [db_partner].[dbo].[usp_Ten_user_tenbyten_InsertAllYearVacation] '" + CStr(insDivcode) + "', '" + CStr(login_userid) + "' "
	''response.write sql
	dbget.Execute(sql)
	response.write "<script>alert('저장 되었습니다.');</script>"

	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
elseif (mode = "addalllongmonthvacation") then

	'// 장기근속 휴가 (만 3년에 한번씩, 입사일자 기준 해당월에 발급, 5일, 1년 유효기간)
	sql = " exec [db_partner].[dbo].[sp_Ten_vacation_insert] '" + CStr(login_userid) + "' "    
	dbget.Execute(sql)
	response.write "<script>alert('저장 되었습니다.');</script>"

	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
elseif (mode = "addalllongyearvacation") then

'-- 4년차 이후 매3년차에 장기근속 휴가 발생(4년차, 7년차 10년차...)
'-- 5일
'-- 임시직 제외
'-- 이상구 제외
'1-6월 입사자는 연차계산에서 1년을 더해준다. -->> 변경 : 입사일자 기준
	sql = "insert into [db_partner].[dbo].tbl_vacation_master(empno, userid, divcd, startday, endday, totalvacationday, registerid) " & vbCrlf
	sql = sql & "select " & vbCrlf
	sql = sql & "	T.empno,T.userid " & vbCrlf
	sql = sql & "	, '5' " & vbCrlf
	sql = sql & "	, DateAdd(yy, (yeardiff - 1), joinday) " & vbCrlf
	sql = sql & "	, DateAdd(dd, -1, DateAdd(yy, yeardiff, joinday)) " & vbCrlf
	'''sql = sql & "	, '' + Cast(Year(getdate()) as varchar) + '-01-01 00:00:01' " & vbCrlf
	'''sql = sql & "	, '' + Cast(Year(getdate()) as varchar) + '-12-31 23:59:59' " & vbCrlf
	sql = sql & "	, 5 as totalvacationday " & vbCrlf
	sql = sql & "	,'system' " & vbCrlf
	sql = sql & "from ( " & vbCrlf
	sql = sql & "	select " & vbCrlf
	sql = sql & "		empno,userid " & vbCrlf
	sql = sql & "		, joinday " & vbCrlf
	sql = sql & "		, (Year(getdate()) - Year(joinday)) + 1 as yeardiff " & vbCrlf
	'''sql = sql & "		, ( " & vbCrlf
	'''sql = sql & "			case " & vbCrlf
	'''sql = sql & "				when (Month(joinday) <= 6) then (Year(getdate()) - Year(joinday)) + 1 " & vbCrlf
	'''sql = sql & "				else (Year(getdate()) - Year(joinday)) " & vbCrlf
	'''sql = sql & "			end " & vbCrlf
	'''sql = sql & "		) as yeardiff " & vbCrlf
	sql = sql & "	from [db_partner].[dbo].tbl_user_tenbyten t, db_partner.[dbo].tbl_partner p " & vbCrlf
	sql = sql & "	where t.userid = p.id " & vbCrlf
	sql = sql & "	and t.isusing = 1 " & vbCrlf
	sql = sql & "	and p.isusing = 'Y' " & vbCrlf
	sql = sql & "	and p.userdiv < 999 " & vbCrlf
	sql = sql & "	and p.level_sn < 999 " & vbCrlf
	sql = sql & "	and t.part_sn <> 4 " & vbCrlf
	sql = sql & "	and t.part_sn <> 5 " & vbCrlf
	sql = sql & "	and t.posit_sn < 12 " & vbCrlf
	'sql = sql & "	and Year(joinday) < Year(getdate()) " & vbCrlf
	'sql = sql & "	and Month(joinday) <= 6 " & vbCrlf
	sql = sql & ") as T " & vbCrlf
	sql = sql & "where T.yeardiff in (4, 7, 10, 13, 16) " & vbCrlf
	sql = sql & "	and t.userid not in ( " & vbCrlf
	sql = sql & "		select userid " & vbCrlf
	sql = sql & "		from " & vbCrlf
	sql = sql & "		[db_partner].[dbo].tbl_vacation_master " & vbCrlf
	sql = sql & "		where divcd = '5' and deleteyn = 'N' and regdate >= CAST(Year(getdate()) AS VARCHAR) + '-01-01' " & vbCrlf
	sql = sql & "		group by userid " & vbCrlf
	sql = sql & "		having count(idx) > 0 " & vbCrlf
	sql = sql & "	) " & vbCrlf
 
	'// 시스템팀:7, 30
	if (session("ssAdminPsn") = 7) or (session("ssAdminPsn") = 30) then
		dbget.Execute(sql)
		response.write "<script>alert('저장 되었습니다.');</script>"
	end if

	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
	
elseif (mode = "addrecalvacation") then

	'// 시급계약직 퇴사자 연차 재정산
	sql = " exec [db_partner].[dbo].[usp_ten_user_tenbyten_RetireCalVacation] '" + CStr(login_userid) + "' "    
	dbget.Execute(sql)
	response.write "<script>alert('저장 되었습니다.');</script>"

	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"	
elseif (mode = "adddetail") then

	registerid = login_userid
	registerempno = login_empno

	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx

	oVacation.GetMasterOne

	'// ========================================================================
	if (oVacation.FItemOne.IsAvailableVacation <> "Y") then
		response.write "<script>alert('사용할 수 없는 휴가입니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	if ((ishalfvacation = "Y") and (CLng(totalday) > 1)) then
		response.write "<script>alert('잘못된 반차등록입니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	if (ishalfvacation = "Y") then
		totalday = 0.5
	end if

	i = (oVacation.FItemOne.Ftotalvacationday - (oVacation.FItemOne.Fusedvacationday + oVacation.FItemOne.Frequestedday))
	if (CDBl(totalday) > i) then
		response.write "<script>alert('잔여일 수 이상의 휴가를 신청하셨습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	if ((Left(oVacation.FItemOne.Fstartday,10) > startday) or (Left(oVacation.FItemOne.Fendday,10) < endday)) then
		'response.write "<script>alert('" & (oVacation.FItemOne.Fstartday < startday) & "');</script>"
		response.write "<script>alert('신청할 수 없는 휴가 기간입니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	'휴가는 해를 넘어 신청할 수 없다
	if (Left(startday,4) <> Left(endday,4)) then
		response.write "<script>alert('시작연도과 종료연도는 같아야 합니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	'// ========================================================================
	sql = "insert into [db_partner].[dbo].tbl_vacation_detail(masteridx, startday, endday, totalday, statedivcd, deleteyn, registerid, registerempno, halfgubun) " & vbCrlf
	sql = sql & "values(" & CStr(masteridx) & ", '" & startday & " 00:00:01', '" & endday & " 23:59:59', " & CStr(totalday) & ", 'R', 'N', '" & registerid & "', '" + CStr(registerempno) + "', '" & vHalfGubun & "') " & vbCrlf
	dbget.Execute(sql)


	'' sql = "update [db_partner].[dbo].tbl_vacation_master " & vbCrlf
	'' sql = sql & "set requestedday = requestedday + " & CStr(totalday) & " " & vbCrlf
	'' sql = sql & "where userid = '" & userid & "' " & vbCrlf
	'' sql = sql & "and idx = " & CStr(masteridx) & " " & vbCrlf
	'' dbget.Execute(sql)

	sql = "update m " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	m.requestedday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'R' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "	, m.usedvacationday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'A' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	sql = sql & "where 1 = 1 " & vbCrlf
	sql = sql & "and m.deleteyn <> 'Y' " & vbCrlf
	sql = sql & "and m.idx = " & CStr(masteridx) & " " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('등록 되었습니다.');</script>"
	response.write "<script>opener.location.reload(); opener.focus();window.close()</script>"
elseif (mode = "allowdetail") then

	approverid = login_userid
	approverempno = login_empno

	Set oVacation = new CTenByTenVacation
	oVacation.FrectDetailIdx = detailidx
	IF masteridx = "" or masteridx = "0" THEN masteridx = oVacation.fnGetMasterIdx
	oVacation.FRectMasterIdx = masteridx
	oVacation.GetMasterOne

	if (oVacation.FItemOne.Fdeleteyn = "Y") then
		response.write "<script>alert('삭제된 데이타는 수정할 수 없습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	sql = "update [db_partner].[dbo].tbl_vacation_detail " & vbCrlf
	sql = sql & "set statedivcd = 'A' " & vbCrlf
	sql = sql & ", approverid = '" & CStr(approverid) & "' " & vbCrlf
	sql = sql & ", approverempno = '" & CStr(approverempno) & "' " & vbCrlf
	sql = sql & ", approveday = getdate() " & vbCrlf
	sql = sql & ", comment = '" & CStr(comment) & "' " & vbCrlf
	sql = sql & "where idx = " & CStr(detailidx) & " " & vbCrlf
	sql = sql & "and deleteyn <> 'Y' " & vbCrlf
	dbget.Execute(sql)

	sql = "update m " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	m.requestedday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'R' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "	, m.usedvacationday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'A' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	sql = sql & "where 1 = 1 " & vbCrlf
	sql = sql & "and m.deleteyn <> 'Y' " & vbCrlf
	sql = sql & "and m.idx = " & CStr(masteridx) & " " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('승인되었습니다.');</script>"
	IF sReturnUrl <> "" THEN
		response.write "<script>location.href = '" & sReturnUrl & "';</script>"
	ELSE
		response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
	END IF
	'response.write "<script>location.href = '/admin/member/tenbyten/tenbyten_vacation_detail_list.asp?masteridx=" & CStr(masteridx) & "';</script>"
elseif (mode = "denydetail") then

	approverid = login_userid
	approverempno = login_empno

	Set oVacation = new CTenByTenVacation
	oVacation.FrectDetailIdx = detailidx
	IF masteridx = "" or masteridx = "0" THEN masteridx = oVacation.fnGetMasterIdx
	oVacation.FRectMasterIdx = masteridx

	oVacation.GetMasterOne

	if (oVacation.FItemOne.Fdeleteyn = "Y") then
		response.write "<script>alert('삭제된 데이타는 수정할 수 없습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	sql = "update [db_partner].[dbo].tbl_vacation_detail " & vbCrlf
	sql = sql & "set statedivcd = 'D' " & vbCrlf
	sql = sql & ", approverid = '" & CStr(approverid) & "' " & vbCrlf
	sql = sql & ", approverempno = '" & CStr(approverempno) & "' " & vbCrlf
	sql = sql & ", approveday = getdate() " & vbCrlf
	sql = sql & ", comment = '" & CStr(comment) & "' " & vbCrlf
	sql = sql & "where idx = " & CStr(detailidx) & " " & vbCrlf
	sql = sql & "and deleteyn <> 'Y' " & vbCrlf
	dbget.Execute(sql)

	sql = "update m " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	m.requestedday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'R' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "	, m.usedvacationday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'A' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	sql = sql & "where 1 = 1 " & vbCrlf
	sql = sql & "and m.deleteyn <> 'Y' " & vbCrlf
	sql = sql & "and m.idx = " & CStr(masteridx) & " " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('거부되었습니다.');</script>"
	IF sReturnUrl <> "" THEN
		response.write "<script>location.href = '" & sReturnUrl & "';</script>"
	ELSE
		response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
	END IF

	'response.write "<script>location.href = '/admin/member/tenbyten/tenbyten_vacation_detail_list.asp?masteridx=" & CStr(masteridx) & "';</script>"
elseif (mode = "deletedetail") then

	approverid = login_userid
	approverempno = login_empno

	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx

	oVacation.GetMasterOne

	if (oVacation.FItemOne.Fdeleteyn = "Y") then
		response.write "<script>alert('삭제된 데이타는 수정할 수 없습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	sql = "update [db_partner].[dbo].tbl_vacation_detail " & vbCrlf
	sql = sql & "set deleteyn = 'Y' " & vbCrlf
	sql = sql & ", approverid = '" & CStr(approverid) & "' " & vbCrlf
	sql = sql & ", approverempno = '" & CStr(approverempno) & "' " & vbCrlf
	sql = sql & ", approveday = getdate() " & vbCrlf
	sql = sql & "where idx = " & CStr(detailidx) & " " & vbCrlf
	sql = sql & "and deleteyn <> 'Y' " & vbCrlf
	dbget.Execute(sql)

	sql = "update m " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	m.requestedday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'R' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "	, m.usedvacationday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'A' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	sql = sql & "where 1 = 1 " & vbCrlf
	sql = sql & "and m.deleteyn <> 'Y' " & vbCrlf
	sql = sql & "and m.idx = " & CStr(masteridx) & " " & vbCrlf
	dbget.Execute(sql)

	Dim objCmd,returnValue
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_DeleteWith]( 1,"&detailidx&",'"&userid&"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
	End With
	returnValue = objCmd(0).Value
	Set objCmd = nothing
	IF 	returnValue =0 THEN
		response.write "<script>alert('품의서 삭제에 문제가 발생했습니다.관리자에게 문의해주세요');</script>"
	END IF

	response.write "<script>alert('삭제되었습니다.');</script>"
	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
elseif (mode = "deletemaster") then

	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx

	oVacation.GetMasterOne

	if (oVacation.FItemOne.Fdeleteyn = "Y") then
		response.write "<script>alert('삭제된 데이타는 수정할 수 없습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	sql = "update [db_partner].[dbo].tbl_vacation_master " & vbCrlf
	sql = sql & "set deleteyn = 'Y' " & vbCrlf
	sql = sql & "where idx = " & CStr(masteridx) & " " & vbCrlf
	sql = sql & "and deleteyn <> 'Y' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('삭제되었습니다.');</script>"
	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"

elseif (mode = "modifymaster") then

	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx

	oVacation.GetMasterOne

	if (oVacation.FItemOne.Fdeleteyn = "Y") then
		response.write "<script>alert('삭제된 데이타는 수정할 수 없습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	if (promotionDay = "") then
		promotionDay = "0"
		jungsanDay = "0"
		retireJungsanDay = "0"
	end if

	sql = "update [db_partner].[dbo].tbl_vacation_master " & vbCrlf
	sql = sql & "set divcd = '" + CStr(divcd) + "' " & vbCrlf
	sql = sql & " , startday = '" + CStr(startday) + "' " & vbCrlf
	sql = sql & " , endday = '" + CStr(endday) + "' " & vbCrlf
	sql = sql & " , totalvacationday = '" + CStr(totalvacationday) + "' " & vbCrlf
	sql = sql & " , promotionDay = '" + CStr(promotionDay) + "' " & vbCrlf
	sql = sql & " , jungsanDay = '" + CStr(jungsanDay) + "' " & vbCrlf
	sql = sql & " , retireJungsanDay = '" + CStr(retireJungsanDay) + "' " & vbCrlf
	sql = sql & " , comment = '" + CStr(comment) + "' " & vbCrlf
	sql = sql & "where idx = " & CStr(masteridx) & " " & vbCrlf
	sql = sql & "and deleteyn <> 'Y' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('변경되었습니다.');</script>"
	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"

end if

%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
