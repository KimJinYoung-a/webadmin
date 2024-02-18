<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim siteDiv, pageDiv, menupos, sqlStr
Dim tplIdx, mainStartDate, mainEndDate, mainTitle
Dim mainRegUserId

siteDiv				= request("site")
pageDiv				= request("pDiv")
menupos				= request("menupos")

tplIdx				= getNumeric(request("tplIdx"))
mainStartDate		= request("StartDate") & " " & request("sTm")
mainEndDate			= request("EndDate") & " " & request("eTm")
mainTitle			= request("mainTitle")
mainRegUserId		= session("ssBctId")

if tplIdx<>"" then
	sqlStr = " insert into [db_sitemaster].[dbo].tbl_cms_mainInfo" + VbCrlf
	sqlStr = sqlStr + " (tplIdx, mainStartDate, mainEndDate, mainTitle, mainTitleYn, mainSortNo, mainTimeYN, mainIcon, mainSubNum, mainExtDataCd"+ VbCrlf
	sqlStr = sqlStr + " , mainIsPreOpen, mainIsUsing, mainRegUserId, mainRegDate, mainWorkRequest, mainStat)"+ VbCrlf

	sqlStr = sqlStr + "select tplIdx, '" & mainStartDate & "', '" & mainEndDate & "', '" & mainTitle & "', 'Y', '50'"+ VbCrlf
	sqlStr = sqlStr + "	, isTimeUse, '', 1, '', 'N', 'Y', '" & mainRegUserId & "',getdate() , '', '0'"+ VbCrlf
	sqlStr = sqlStr + "from db_sitemaster.dbo.tbl_cms_template"+ VbCrlf
	sqlStr = sqlStr + "where tplIdx=" & tplIdx + VbCrlf
	sqlStr = sqlStr + "	and siteDiv='" & siteDiv & "'"+ VbCrlf
	sqlStr = sqlStr + "	and pageDiv='" & pageDiv & "'"+ VbCrlf
	dbget.Execute sqlStr
end if

dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->