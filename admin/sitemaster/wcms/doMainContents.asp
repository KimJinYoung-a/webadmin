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
Dim mainIdx, tplIdx, mainStartDate, mainEndDate, mainTitle, mainTitleYn, mainSortNo, mainTimeYN, mainIcon, mainSubNum, mainExtDataCd
Dim mainIsPreOpen, mainIsUsing, mainRegUserId, mainRegDate, mainLastModiUserid, mainLastModiDate, mainWorkRequest, mainStat
Dim srcSDT, srcEDT

siteDiv				= request("site")
pageDiv				= request("pDiv")
menupos				= request("menupos")

mainIdx				= request("mainIdx")
tplIdx				= request("tplIdx")
mainStartDate		= request("StartDate") & " " & request("sTm")
mainEndDate			= request("EndDate") & " " & request("eTm")
mainTitle			= request("mainTitle")
mainTitleYn			= request("mainTitleYn")
mainSortNo			= request("mainSortNo")
mainTimeYN			= request("mainTimeYN")
mainIcon			= request("mainIcon")
mainSubNum			= request("mainSubNum")
mainExtDataCd		= request("mainExtDataCd")
mainIsPreOpen		= request("mainIsPreOpen")
mainIsUsing			= request("mainIsUsing")
mainRegUserId		= session("ssBctId")
mainLastModiUserid	= session("ssBctId")
mainWorkRequest		= request("mainWorkRequest")
mainStat			= request("mainStat")
srcSDT				= request("srcSDT")
srcEDT				= request("srcEDT")

if (mainIdx<>"") then
    sqlStr = " update [db_sitemaster].[dbo].tbl_cms_mainInfo" + VbCrlf

    sqlStr = sqlStr + " Set tplIdx='" + tplIdx + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainStartDate='" + html2db(mainStartDate) + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainEndDate='" + html2db(mainEndDate) + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainTitle='" + html2db(mainTitle) + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainTitleYn='" + mainTitleYn + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainSortNo='" + html2db(mainSortNo) + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainTimeYN='" + mainTimeYN + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainIcon='" + mainIcon + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainSubNum='" + html2db(mainSubNum) + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainExtDataCd='" + mainExtDataCd + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainIsPreOpen='" + mainIsPreOpen + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainIsUsing='" + mainIsUsing + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainLastModiUserid='" + mainLastModiUserid + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainLastModiDate= getdate()" + VbCrlf
    sqlStr = sqlStr + " ,mainWorkRequest='" + html2db(mainWorkRequest) + "'" + VbCrlf
    sqlStr = sqlStr + " ,mainStat='" + mainStat + "'" + VbCrlf
    sqlStr = sqlStr + " where mainIdx=" + CStr(mainIdx) + VbCrlf
    
    dbget.Execute sqlStr
else
    sqlStr = " insert into [db_sitemaster].[dbo].tbl_cms_mainInfo" + VbCrlf
    sqlStr = sqlStr + " (tplIdx, mainStartDate, mainEndDate, mainTitle, mainTitleYn, mainSortNo, mainTimeYN, mainIcon, mainSubNum, mainExtDataCd"+ VbCrlf
	sqlStr = sqlStr + " , mainIsPreOpen, mainIsUsing, mainRegUserId, mainRegDate, mainWorkRequest, mainStat)"+ VbCrlf
    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " '" + tplIdx + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + html2db(mainStartDate) + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + html2db(mainEndDate) + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + html2db(mainTitle) + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + mainTitleYn + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + html2db(mainSortNo) + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + mainTimeYN + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + mainIcon + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + html2db(mainSubNum) + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + mainExtDataCd + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + mainIsPreOpen + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + mainIsUsing + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + mainRegUserId + "'" + VbCrlf
    sqlStr = sqlStr + ",getdate()" + VbCrlf
    sqlStr = sqlStr + ",'" + html2db(mainWorkRequest) + "'" + VbCrlf
    sqlStr = sqlStr + ",'" + mainStat + "'" + VbCrlf
    sqlStr = sqlStr + " )" + VbCrlf
    
    dbget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('[db_sitemaster].[dbo].tbl_cms_mainInfo') as idx"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
		mainIdx = rsget("idx")
	end if
	rsget.close
end if

dim retUrl
retUrl = "mainPageManage.asp?site=" & siteDiv & "&pDiv=" & pageDiv & "&menupos=" & menupos & "&mainIdx=" & mainIdx & "&sDt=" & srcSDT & "&eDt=" & srcEDT
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + retUrl + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->