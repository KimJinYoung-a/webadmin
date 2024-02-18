<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	dim menupos, iLp, sYYYY, sMM, sCDL, sLevel, arrLecIdx, arrSn, mode
	dim strSql, rstMsg

	menupos = RequestCheckvar(request("menupos"),10)
	mode = RequestCheckvar(request("mode"),16)
	sYYYY = RequestCheckvar(request("yyyy"),4)
	sMM = RequestCheckvar(request("mm"),2)
	sCDL = RequestCheckvar(request("cdl"),3)
	sLevel = RequestCheckvar(request("level"),1)
	arrLecIdx = request("arrLecIdx")
	arrSn = request("arrSn")
  	if arrLecIdx <> "" then
		if checkNotValidHTML(arrLecIdx) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if arrSn <> "" then
		if checkNotValidHTML(arrSn) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if	
'회차가 필요없다 하셔서 진영 하단 2줄 추가(2012-09-17)
	if sYYYY="" then sYYYY=year(date)
	if sMM="" then sMM=Num2Str(Month(date),2,"0","R")

	if mode="" then
		Call Alert_return("잘못된 접근입니다.")
		response.End
	end if

	Select Case mode
		Case "add"
			if arrLecIdx="" then 
				Call Alert_return("등록할 강좌번호가 없습니다.")
				response.End
			end if
			if sYYYY="" or sMM="" then 
				Call Alert_return("등록할 회차가 지정되지 않았습니다.")
				response.End
			end if
			if sCDL="" then 
				Call Alert_return("등록할 카테고리가 지정되지 않았습니다.")
				response.End
			end if
			if sLevel="" then 
				Call Alert_return("등록할 강좌의 난이도가 지정되지 않았습니다.")
				response.End
			end if

			'중복 체크
			strSql = "Select lecIdx From [db_academy].[dbo].tbl_lec_pickInfo " &_
					" Where YYYYMM='" & sYYYY & sMM & "'" &_
					"	and lecIdx in (" & arrLecIdx & ")"
			rsACADEMYget.Open strSql, dbACADEMYget, 1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				rstMsg = "지정하신 회차에 이미 등록된 ["
				do until rsACADEMYget.EOF
					if rstMsg="지정하신 회차에 이미 등록된 [" then
						rstMsg = rstMsg & rsACADEMYget(0)
					else
						rstMsg = rstMsg & "," & rsACADEMYget(0)
					end if
					rsACADEMYget.MoveNext
				loop
				rstMsg = rstMsg & "]를 제외하고\n"
			end if
			rsACADEMYget.Close

			'저장 처리
			strSql = "Insert Into [db_academy].[dbo].tbl_lec_pickInfo (YYYYMM,lecLevel,code_large,lecIdx) " &_
					" Select '" & sYYYY & sMM & "', '" & sLevel & "', '" & sCDL & "', idx " &_
					" From [db_academy].[dbo].tbl_lec_item " &_
					" Where idx in (" & arrLecIdx & ")" &_
					"	and idx not in (" &_
					"		Select lecIdx From [db_academy].[dbo].tbl_lec_pickInfo " &_
					" 		Where YYYYMM='" & sYYYY & sMM & "'" &_
					"			and lecIdx in (" & arrLecIdx & "))"
			dbACADEMYget.execute(strSql)

			rstMsg = rstMsg & "강좌가 등록되었습니다."
		Case "del"
			if arrSn="" then 
				Call Alert_return("삭제할 강좌번호가 없습니다.")
				response.End
			end if

			'선택 삭제
			strSql = "delete from [db_academy].[dbo].tbl_lec_pickInfo " &_
					" Where pickSn in (" & arrSn & ")"
			dbACADEMYget.execute(strSql)

			rstMsg = "선택하신 강좌가 삭제되었습니다."
	end Select

	Call Alert_Move(rstMsg,"lec_pickInfo.asp?menupos="&menupos&"&yyyy="&sYYYY&"&mm="&sMM&"&cdl="&sCDL&"&level="&sLevel)
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyClose.asp" -->