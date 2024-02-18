<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim vQuery, vOrderSerial, vIsNextStatus, vWantStudyName, vWantStudyYear, vWantStudyMonth, vWantStudyDay, vWantStudyAmPm, vWantStudyHour, vWantStudyMin, vWantStudyPlace, vWantStudyWho
	vOrderSerial = RequestCheckvar(Request("orderserial"),16)
	vWantStudyName	= Trim(request.Form("wantstudyName"))
	vWantStudyName = Replace(vWantStudyName,chr(34),"")
	vWantStudyName = Replace(vWantStudyName,"'","")
	vWantStudyName = Replace(vWantStudyName,chr(34),"")
	vWantStudyYear	= Trim(RequestCheckvar(request.Form("wantstudyYear"),4))
	vWantStudyMonth	= Trim(RequestCheckvar(request.Form("wantstudyMonth"),2))
	vWantStudyDay	= Trim(RequestCheckvar(request.Form("wantstudyDay"),2))
	vWantStudyAmPm	= Trim(RequestCheckvar(request.Form("wantstudyAmPm"),4))
	vWantStudyHour	= Trim(RequestCheckvar(request.Form("wantstudyHour"),2))
	vWantStudyMin	= Trim(RequestCheckvar(request.Form("wantstudyMin"),2))
	vWantStudyPlace	= Trim(request.Form("wantstudyPlace"))
	vWantStudyPlace = Replace(vWantStudyPlace,"'","")
	vWantStudyPlace = Replace(vWantStudyPlace,chr(34),"")
	vWantStudyWho	= Trim(RequestCheckvar(request.Form("wantstudyWho"),6))
	vIsNextStatus	= RequestCheckvar(request.Form("gopay"),1)

  	if vWantStudyName <> "" then
		if checkNotValidHTML(vWantStudyName) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if

	If vOrderSerial <> "" Then
		vQuery = "UPDATE [db_academy].[dbo].[tbl_academy_order_weclass] SET "
		vQuery = vQuery & "		wantstudyName = '" & vWantStudyName & "', "
		vQuery = vQuery & "		wantstudyYear = '" & vWantStudyYear & "', "
		vQuery = vQuery & "		wantstudyMonth = '" & vWantStudyMonth & "', "
		vQuery = vQuery & "		wantstudyDay = '" & vWantStudyDay & "', "
		vQuery = vQuery & "		wantstudyAmPm = '" & vWantStudyAmPm & "', "
		vQuery = vQuery & "		wantstudyHour = '" & vWantStudyHour & "', "
		vQuery = vQuery & "		wantstudyMin = '" & vWantStudyMin & "', "
		vQuery = vQuery & "		wantstudyPlace = '" & vWantStudyPlace & "', "
		vQuery = vQuery & "		wantstudyWho = '" & vWantStudyWho & "' "
		vQuery = vQuery & "WHERE orderserial = '" & vOrderSerial & "'"
		dbACADEMYget.Execute(vQuery)
		
		If vIsNextStatus = "o" Then
			vQuery = "UPDATE [db_academy].[dbo].[tbl_academy_order_master] SET ipkumdiv = '3' WHERE orderserial = '" & vOrderSerial & "'"
			dbACADEMYget.Execute(vQuery)
		End IF
		
		rw "<script language='javascript'>alert('저장되었습니다.');opener.location.reload();window.close();</script>"
	Else
		rw "<script language='javascript'>alert('잘못된 경로입니다.');window.close();</script>"
	End IF
%>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->