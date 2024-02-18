<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 안전인증품목관리
' History : 2018.01.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim strSql, i, lastuserid, menupos, mode
dim infoDiv, infoDivName, infoValidCnt, SafetyTargetYN, SafetyCertYN, SafetyConfirmYN, SafetySupplyYN, SafetyComply
	lastuserid=session("ssBctId")
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	mode = requestcheckvar(request("mode"),32)

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

if (InStr(referer,"10x10.co.kr")<1) then
	response.write "not valid Referer"
    response.end
end if

if mode="safetylistedit" then
	for i=1 to request.form("infoDiv").count
		if request.form("infoDiv")="" then
			response.write "품목번호가 없습니다."
			dbget.close()	:	response.end
		else
			infoDiv = requestcheckvar(request.form("infoDiv")(i),2)
		end if

		if request.form("SafetyTargetYN_"&infoDiv)="" then
			response.write "안전인증대상여부가 없습니다[0]."
			dbget.close()	:	response.end
		else
			SafetyTargetYN = requestcheckvar(request.form("SafetyTargetYN_"&infoDiv),1)
		end if

		if request.form("SafetyCertYN_"&infoDiv)="" then
			response.write "안전인증여부가 없습니다."
			dbget.close()	:	response.end
		else
			SafetyCertYN = requestcheckvar(request.form("SafetyCertYN_"&infoDiv),1)
		end if

		if request.form("SafetyConfirmYN_"&infoDiv)="" then
			response.write "안전확인여부가 없습니다."
			dbget.close()	:	response.end
		else
			SafetyConfirmYN = requestcheckvar(request.form("SafetyConfirmYN_"&infoDiv),1)
		end if

		if request.form("SafetySupplyYN_"&infoDiv)="" then
			response.write "공급자적합성여부가 없습니다."
			dbget.close()	:	response.end
		else
			SafetySupplyYN = requestcheckvar(request.form("SafetySupplyYN_"&infoDiv),1)
		end if

		if request.form("SafetyComply_"&infoDiv)="" then
			response.write "안전기준준수여부가 없습니다."
			dbget.close()	:	response.end
		else
			SafetyComply = requestcheckvar(request.form("SafetyComply_"&infoDiv),1)
		end if

		strSql = "Update db_item.dbo.tbl_item_infoDiv" & vbcrlf
		strSql = strSql & " Set SafetyTargetYN='" & trim(SafetyTargetYN) & "'" & vbcrlf
		strSql = strSql & " ,SafetyCertYN='" & trim(SafetyCertYN) & "'" & vbcrlf
		strSql = strSql & " ,SafetyConfirmYN='" & trim(SafetyConfirmYN) & "'" & vbcrlf
		strSql = strSql & " ,SafetySupplyYN='" & trim(SafetySupplyYN) & "'" & vbcrlf
		strSql = strSql & " ,SafetyComply='" & trim(SafetyComply) & "'" & vbcrlf
		strSql = strSql & " ,lastupdate=getdate()" & vbcrlf
		strSql = strSql & " ,lastadminid='" & trim(lastuserid) & "' Where " & vbcrlf
		strSql = strSql & " infoDiv="& trim(infoDiv) &"" & vbcrlf

		'response.write strSql & "<br>"
		dbget.Execute strSql
	next

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.replace('/admin/itemmaster/safetycert/safetycert.asp?menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('구분자가 없습니다.');"
	response.write "</script>"
	dbget.close()	:	response.end
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->