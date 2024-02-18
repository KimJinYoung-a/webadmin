<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사원관리 엑셀다운로드
' History : 2011.1.19 정윤정 생성
'           2018.03.30 허진원 - 직급 선택 표시
'			2023.08.23 한용민 수정(휴대폰번호 추가)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
response.write "사용안함"
response.end
Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby, ilevel_sn, iTotCnt,iPageSize, iTotalPage
Dim job_sn, posit_sn, continuous_service_year, employeeonly, nodepartonly, criticinfouser, workdaycheck
dim fromDate, toDate, department_id, inc_subdepartment, rank_sn, lv1customerYN, lv2partnerYN, lv3InternalYN
dim imaxchposit, yyyy1, yyyy2, mm1, mm2, dd1, dd2, oMember, arrList,intLoop, i, j, k,p, sTitle, bufStr
dim oldempno
	workdaycheck = requestcheckvar(request("workdaycheck"),1)
	lv1customerYN 	= requestCheckvar(request("lv1customerYN"),1)
	lv2partnerYN 	= requestCheckvar(request("lv2partnerYN"),1)
	lv3InternalYN 	= requestCheckvar(request("lv3InternalYN"),1)
	yyyy1 = requestcheckvar(request("yyyy1"),4)
	yyyy2 = requestcheckvar(request("yyyy2"),4)
	mm1	  = requestcheckvar(request("mm1"),2)
	mm2	  = requestcheckvar(request("mm2"),2)
	dd1	  = requestcheckvar(request("dd1"),2)
	dd2	  = requestcheckvar(request("dd2"),2)
	iPageSize	  = requestcheckvar(request("pagesize"),10)
	page = requestCheckvar(Request("page"),10)
	isUsing = requestCheckvar(Request("isUsing"),1)
	SearchKey = requestCheckvar(Request("SearchKey"),1)
	SearchString = requestCheckvar(Request("SearchString"),32)
	part_sn = requestCheckvar(Request("part_sn"),10)
	job_sn = requestCheckvar(Request("job_sn"),10)
	posit_sn = requestCheckvar(Request("posit_sn"),10)
	research = requestCheckvar(Request("research"),2)
	orderby = requestCheckvar(Request("orderby"),1)
	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)
	nodepartonly = requestCheckvar(Request("nodepartonly"),1)
	criticinfouser = requestCheckvar(Request("criticinfouser"),10)
	rank_sn = requestCheckvar(Request("rank_sn"),2)
	ilevel_sn = requestCheckvar(Request("ilevel_sn"),10)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(1)

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()) + 1)
	dd2 = Cstr(1)

	toDate = CStr(DateSerial(yyyy2, mm2, 0))

	yyyy2 = CStr(Year(toDate))
	mm2 = CStr(Month(toDate))
	dd2 = CStr(Day(toDate))
end if

toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if (iPageSize = "") then
	iPageSize = 20
end if

if isUsing="" and research="" then isUsing="Y"
if page="" then page=1

Set oMember = new CTenByTenMember
	oMember.FPagesize 	= iPageSize
	oMember.FCurrPage 	= page
	oMember.FSearchType 	= searchKey
	oMember.FSearchText 	= searchString
	oMember.Fstatediv 		= isUsing
	oMember.Fpart_sn 		= part_sn
	oMember.Fjob_sn 		= job_sn
	oMember.Fposit_sn 		= posit_sn
	oMember.Frank_sn		= rank_sn
	oMember.Fdepartment_id 		= department_id
	oMember.Finc_subdepartment 	= inc_subdepartment
	oMember.FRectNoDepartOnly 	= nodepartonly
	oMember.FRectCriticInfoUser 	= criticinfouser
	oMember.Flevel_sn = ilevel_sn
	oMember.Forderby 		= orderby

	if (workdaycheck = "Y") then
		oMember.FStartDate		= fromDate
		oMember.FEndDate		= toDate
	end if
	oMember.Frectlv1customerYN = lv1customerYN
	oMember.Frectlv2partnerYN = lv2partnerYN
	oMember.Frectlv3InternalYN = lv3InternalYN
	imaxchposit= oMember.fnGetMaxChangePosit
	arrList = oMember.fnGetMemberList_csv
	iTotCnt = oMember.FTotCnt
set oMember = nothing

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=EmployeeManagementList.csv"
Response.CacheControl = "public"

sTitle = "직책,사번,이름,영문이름,아이디,입사일(정규직),실제입사일,퇴사일,연차,부서,직위,LV1(고객정보),LV2(파트너정보),LV3(내부정보),계약전환여부,GSSHOP아이디,핸드폰번호"
for p =1 To imaxchposit
sTitle=sTitle& ",전환일, 전환전부서, 전환전직위"
next

response.write sTitle& VbCrlf

if isArray(arrList) then
	for intLoop=0 to ubound(arrList,2)

	if oldempno <> arrList(0,intLoop) then
		if intLoop <>0 then
			bufStr =bufStr & VbCrlf
			response.write bufStr
		end if

		bufStr = ""
		bufStr = bufStr & arrList(14,intLoop)
		bufStr = bufStr & "," &arrList(0,intLoop)
		bufStr = bufStr & "," &arrList(1,intLoop)
		bufStr = bufStr & "," &arrList(44,intLoop)
		bufStr = bufStr & "," &arrList(2,intLoop)
		bufStr = bufStr & "," &Left(arrList(3,intLoop), 10)

		bufStr = bufStr & ","

		if Not IsNull(arrList(24,intLoop)) then
			bufStr = bufStr &Left(arrList(24,intLoop), 10)
		end if

		bufStr = bufStr & ","

		if Not IsNull(arrList(4,intLoop)) then
			if arrList(15,intLoop) <> "N" then
				bufStr = bufStr & Left(arrList(4,intLoop), 10)
			else
				if (arrList(26,intLoop) = 99) then
					bufStr = bufStr   &  Left(arrList(4,intLoop), 10)
				else
					bufStr = bufStr   & Left(arrList(4,intLoop), 10)
				end if
			end if
		end if

		bufStr = bufStr & ","
		IF Not isNull(arrList(3,intLoop)) and Left(arrList(0,intLoop), 1) = "1" THEN
			if Not IsNull(arrList(24,intLoop)) then
				if GetYearDiff(arrList(24,intLoop)) >= 1 then
					bufStr = bufStr &   GetYearDiff(arrList(24,intLoop))  &"년"
				end if
				if GetMonthDiff(arrList(24,intLoop)) > 0 THEN
					bufStr = bufStr & GetMonthDiff(arrList(24,intLoop)) &"개월"
				end if
			else
				bufStr = bufStr & arrList(3,intLoop)
				if GetYearDiff(arrList(3,intLoop)) >= 1 then
					bufStr = bufStr & GetYearDiff(arrList(3,intLoop))&"년"
				end if
				if GetMonthDiff(arrList(3,intLoop)) > 0 THEN
					bufStr = bufStr & GetMonthDiff(arrList(3,intLoop))&"개월"
				end if
			end if
		END IF
		bufStr = bufStr & "," &arrList(27,intLoop)
		bufStr = bufStr & "," &arrList(13,intLoop)
		'bufStr = bufStr & "," &GetCriticInfoUserLevelName(arrList(30,intLoop))
		bufStr = bufStr & "," &arrList(41,intLoop)
		bufStr = bufStr & "," &arrList(42,intLoop)
		bufStr = bufStr & "," &arrList(43,intLoop)

		if arrList(33,intLoop) >0 then
			bufStr = bufStr & ",Y"
		else
			bufStr = bufStr & ",N "
		end if
		bufStr = bufStr & "," & db2html(arrList(40,intLoop))
		bufStr = bufStr & "," &arrList(17,intLoop)
	END IF

	bufStr = bufStr & "," &arrList(39,intLoop)
	bufStr = bufStr & "," &arrList(38,intLoop)
	bufStr = bufStr & "," &arrList(36,intLoop)

	oldempno = arrList(0,intLoop)
	next
	bufStr =bufStr & VbCrlf
	response.write bufStr
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
