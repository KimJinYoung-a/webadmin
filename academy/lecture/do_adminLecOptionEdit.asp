<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// 변수 선언
dim lp, menupos
dim mode, lec_idx
dim RegStartDate, RegEndDate, LecSDay, LecSTime, LecETime, limit_count, limit_sold, isUsing
dim tmpRegSDt, tmpRegEDt, tmpLecOptName, tmpLecSDate, tmpLecEDate
dim OptCnt, lecOption, lecOptionName, min_count, SqlStr, vOrderSerial
dim msg, retURL

'// 내용 접수 및 처리
mode			= RequestCheckvar(Request("mode"),16)
lec_idx			= RequestCheckvar(Request("lec_idx"),10)
lecOption		= Request("lecOption")
lecOptionName	= Request("lecOptionName")
RegStartDate	= Request("RegStartDate")
RegEndDate		= Request("RegEndDate")
LecSDay			= Request("LecSDay")
LecSTime		= Request("LecSTime")
LecETime		= Request("LecETime")
limit_count		= Request("limit_count")
limit_sold		= Request("limit_sold")
isusing			= Request("isusing")
tmpRegSDt		= Request("tmpRegSDt")
tmpRegEDt		= Request("tmpRegEDt")
tmpLecOptName	= Request("tmpLecOptName")
tmpLecSDate		= Request("tmpLecSDt") & " " & Request("tmpLecStime")
tmpLecEDate		= Request("tmpLecSDt") & " " & Request("tmpLecEtime")
if lecOption <> "" then
	if checkNotValidHTML(lecOption) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if lecOptionName <> "" then
	if checkNotValidHTML(lecOptionName) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if RegStartDate <> "" then
	if checkNotValidHTML(RegStartDate) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if RegEndDate <> "" then
	if checkNotValidHTML(RegEndDate) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if LecSDay <> "" then
	if checkNotValidHTML(LecSDay) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if LecSTime <> "" then
	if checkNotValidHTML(LecSTime) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if LecETime <> "" then
	if checkNotValidHTML(LecETime) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if limit_count <> "" then
	if checkNotValidHTML(limit_count) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if limit_sold <> "" then
	if checkNotValidHTML(limit_sold) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if isusing <> "" then
	if checkNotValidHTML(isusing) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
'배열처리
lecOption		=	split(lecOption,",")
lecOptionName	=	split(lecOptionName,",")
RegStartDate	=	split(RegStartDate,",")
RegEndDate		=	split(RegEndDate,",")
LecSDay			=	split(LecSDay,",")
LecSTime		=	split(LecSTime,",")
LecETime		=	split(LecETime,",")
limit_count		=	split(limit_count,",")
limit_sold		=	split(limit_sold,",")
isusing			=	split(isusing,",")


OptCnt = ubound(lecOption)

vOrderSerial = RequestCheckvar(Request("orderserial"),16)
'==============================================================================
'## 내용 저장/수정 처리

'트랜젝션 시작
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ 신규등록

				'옵션코드를 위해 가장 큰번호 및 강좌테이블의 정원 및 최소인원 접수
				SqlStr= "Select top 1 t1.limit_count, t1.min_count, t1.optionCnt, Max(t2.lecOption) as optCd " &_
						" from [db_academy].[dbo].tbl_lec_item as t1 " &_
						" 	join [db_academy].[dbo].tbl_lec_item_option as t2 " &_
						" 		on t1.idx=t2.lecIdx " &_
						" Where t1.idx='" & CStr(lec_idx) & "' " &_
						" group by t1.limit_count, t1.min_count, t1.optionCnt " &_
						" order by optCd desc "
				rsACADEMYget.open sqlStr,dbACADEMYget,1

					if not rsACADEMYget.eof then
						lecOption	= Cint(rsACADEMYget("optCd"))+1
						limit_count	= rsACADEMYget("limit_count")
						min_count	= rsACADEMYget("min_count")
						OptCnt		= rsACADEMYget("optionCnt")
						if OptCnt<=0 then OptCnt=1
					else
						response.write	"<script language='javascript'>" &_
										"	alert('잘못된 강좌정보 입니다.');" &_
										"	history.back();" &_
										"</script>"
						rsACADEMYget.close(): dbACADEMYget.close():	response.End
					end if
				rsACADEMYget.close

				'옵션 저장
				lecOption = Num2Str(lecOption,4,"0","R")
				SqlStr= " insert into [db_academy].[dbo].tbl_lec_item_option (lecIdx,lecOption,lecOptionName,regStartDate, regEndDate, lecStartDate, lecEndDate, limit_count, min_count)" &_
						" values ('" + CStr(lec_idx) + "','" & lecOption & "','" & tmpLecOptName & "','" & tmpRegSDt & "','" & tmpRegEDt & "'" &_
						"	,'"& tmpLecSDate & "','" & tmpLecEDate & "','" & CInt(limit_count/OptCnt) & "','" & min_count & "')"
				dbACADEMYget.Execute(SqlStr)

				'스케줄 저장
				SqlStr= " insert into [db_academy].[dbo].tbl_lec_schedule (lec_idx,lecOption,StartDate, EndDate)" &_
						" values ('" + CStr(lec_idx) + "','" & lecOption & "','" & tmpLecSDate & "','" & tmpLecEDate & "')"
				dbACADEMYget.Execute(SqlStr)

				'강의 테이블 옵션수 수정
				SqlStr= " update db_academy.dbo.tbl_lec_item " &_
						" set lec_period='" & trim(lecOptionName(0)) & "',optionCnt=A.cnt " &_
						"	,limit_count=A.ttLmtCnt, limit_sold=A.ttSold " &_
						"	,reg_StartDay=A.rSdt, reg_EndDay=A.rEdt " &_
						" from ( " &_
						" 	select lecidx, count(*) as cnt, sum(limit_count) as ttLmtCnt, sum(limit_sold) as ttSold " &_
						"		,min(regStartDate) rSdt, max(regEndDate) rEdt" &_
						" 	from db_academy.dbo.tbl_lec_item_option " &_
						" 	where isusing='Y' " &_
						" 	group by lecidx) as A " &_
						" where db_academy.dbo.tbl_lec_item.Idx=A.lecIdx " &_
						" 	and idx='" & CStr(lec_idx) & "'"
				dbACADEMYget.Execute(SqlStr)
				
		msg = "등록하였습니다."

		'돌아갈 페이지
		retURL = "popLecOptionEdit.asp?lec_idx=" & lec_idx

	Case "modi"
		'@@ 수정처리
		
		for lp=0 to OptCnt
			if trim(lecOption(lp))<>"" then
				'@ 옵션 수정
				SqlStr= " update [db_academy].[dbo].tbl_lec_item_option " &_
						" Set regStartDate='"& trim(RegStartDate(lp)) & "'" &_
						" ,regEndDate='" & trim(RegEndDate(lp)) & "'" &_
						" ,lecOptionName='" & trim(lecOptionName(lp)) & "'" &_
						" ,lecStartDate='" & trim(LecSDay(lp)) & " " & trim(LecSTime(lp)) & "'" &_
						" ,lecEndDate='" & trim(LecSDay(lp)) & " " & trim(LecETime(lp)) & "'" &_
						" ,limit_count='" & Trim(limit_count(lp)) & "'" &_
						" ,limit_sold='" & Trim(limit_sold(lp)) & "'" &_
						" ,isusing='" & Trim(isusing(lp)) & "'" &_
						" Where lecIdx='" & CStr(lec_idx) & "'" &_
						"	and lecOption='" & trim(lecOption(lp)) & "'"
				dbACADEMYget.Execute(SqlStr)
			end if
		next

		'강의 테이블 옵션수 수정
		SqlStr= " update db_academy.dbo.tbl_lec_item " &_
				" set lec_period='" & trim(lecOptionName(0)) & "',optionCnt=A.cnt " &_
				"	,limit_count=A.ttLmtCnt, limit_sold=A.ttSold " &_
				"	,reg_StartDay=A.rSdt, reg_EndDay=A.rEdt " &_
				" from ( " &_
				" 	select lecidx, count(*) as cnt, sum(limit_count) as ttLmtCnt, sum(limit_sold) as ttSold " &_
				"		,min(regStartDate) rSdt, max(regEndDate) rEdt" &_
				" 	from db_academy.dbo.tbl_lec_item_option " &_
				" 	where isusing='Y' " &_
				" 	group by lecidx) as A " &_
				" where db_academy.dbo.tbl_lec_item.Idx=A.lecIdx " &_
				" 	and idx='" & CStr(lec_idx) & "'"
		dbACADEMYget.Execute(SqlStr)

		msg = "수정하였습니다."
		'돌아갈 페이지
		retURL = "popLecOptionEdit.asp?lec_idx=" & lec_idx

End Select

If vOrderSerial <> "" Then
	retURL = retURL & "&orderserial="&vOrderSerial&""
End IF

'오류검사 및 반영
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'커밋(정상)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbACADEMYget.RollBackTrans				'롤백(에러발생시)

	response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->