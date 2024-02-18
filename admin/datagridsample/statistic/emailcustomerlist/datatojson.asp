<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' PageName : /datagridsample/statistic/emailcustomerlist/datajson.asp
' Discription : datagridsample - emailcustomerlist
' Response : response > 결과
' History : 2019.07.29 
'###############################################

'//헤더 출력
Response.ContentType = "application/json"
response.charset = "utf-8"

Dim sData : sData = Request("json")
Dim oJson
dim omd , i
dim page :  page = 1
'// 전송결과 파징
on Error Resume Next

'// json객체 선언
SET oJson = jsArray()
Dim contents_json , contents_object

IF (Err) then
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다.1"
else
    dim sendName , mailTitle , totalSendUserCount , successSendCount , emailOpenCount , emailClickCount
    dim emailSendDate , emailSendComplateDate , totalCount , idx
    set omd = New CMailzine
        omd.FCurrPage = page
        omd.FPageSize=100
        omd.GetMailingList

        totalcount = omd.FResultCount
	
        if omd.FResultCount > 0 then
            ReDim contents_object(omd.FResultCount-1)
            FOR i=0 to omd.FResultCount-1
                sendName                = omd.FItemList(i).Ftitle
                mailTitle               = omd.FItemList(i).fsubject
                totalSendUserCount      = omd.FItemList(i).Ftotalcnt
                successSendCount        = omd.FItemList(i).fsuccesscnt
                emailOpenCount          = omd.FItemList(i).fopencnt
                emailClickCount         = omd.FItemList(i).fclickcnt
                emailSendDate           = omd.FItemList(i).Fstartdate
                emailSendComplateDate   = omd.FItemList(i).Fenddate
                idx                     = omd.FItemList(i).Fidx

                Set oJson(null) = jsObject()
                    oJson(null)("sendName")              = sendName
                    oJson(null)("mailTitle")	         = mailTitle
                    oJson(null)("totalSendUserCount")	 = totalSendUserCount
                    oJson(null)("successSendCount")	     = successSendCount
                    oJson(null)("emailOpenCount")	     = emailOpenCount
                    oJson(null)("emailClickCount")		 = emailClickCount
                    oJson(null)("emailSendDate")         = FormatDate(emailSendDate,"0000-00-00 00:00:00")
                    oJson(null)("emailSendCompleteDate") = FormatDate(emailSendComplateDate,"0000-00-00 00:00:00")
                    oJson(null)("idx")                   = idx
            next
        end if 

    set omd = nothing

	'// 결과 출력
	IF (Err) then
		oJson("response") = getErrMsg("9999",sFDesc)
		oJson("faildesc") = "처리중 오류가 발생했습니다.2"
	end if
end if

'Json 출력(JSON)
oJson.flush
Set oJson = Nothing

if ERR then Call OnErrNoti()
On Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->