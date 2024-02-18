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
Dim sqlStr
Dim Idx, appName, startDate, endDate, eventName, sortNo, bannerType, bannerImg, bannerLink, isUsing, regUserid, lastUpdateUser, workComment

Idx				= request("idx")
appName			= request("appName")
startDate		= request("startDate") & " " & request("startTime")
endDate			= request("endDate") & " " & request("endTime")
eventName		= html2db(request("eventName"))
sortNo			= request("sortNo")
bannerType		= request("bannerType")
bannerImg		= request("bannerImg")
bannerLink		= html2db(request("bannerLink"))
isUsing			= request("isusing")
regUserid		= session("ssBctId")
lastUpdateUser	= session("ssBctId")
workComment		= html2db(request("workcomment"))

if (Idx<>"") then
    sqlStr = " update [db_contents].[dbo].tbl_app_eventBanner" + VbCrlf
    sqlStr = sqlStr + " Set appName='" + appName + "'" + VbCrlf
    sqlStr = sqlStr + " ,startDate='" + startDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,endDate='" + endDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,eventName='" + eventName + "'" + VbCrlf
    sqlStr = sqlStr + " ,sortNo='" + sortNo + "'" + VbCrlf
    sqlStr = sqlStr + " ,bannerType='" + bannerType + "'" + VbCrlf
    sqlStr = sqlStr + " ,bannerImg='" + bannerImg + "'" + VbCrlf
    sqlStr = sqlStr + " ,bannerLink='" + bannerLink + "'" + VbCrlf
    sqlStr = sqlStr + " ,isUsing='" + isUsing + "'" + VbCrlf
    sqlStr = sqlStr + " ,lastUpdateUser='" + lastUpdateUser + "'" + VbCrlf
    sqlStr = sqlStr + " ,lastUpdate=getdate()" + VbCrlf
    sqlStr = sqlStr + " ,workComment='" + workComment + "'" + VbCrlf
    sqlStr = sqlStr + " where idx=" + CStr(idx) + VbCrlf
    
    dbget.Execute sqlStr
else
    sqlStr = " insert into [db_contents].[dbo].tbl_app_eventBanner" + VbCrlf
    sqlStr = sqlStr + " (appName, startDate, endDate, eventName, sortNo, bannerType, bannerImg, bannerLink, isUsing, regUserid, regdate, workComment)"+ VbCrlf
    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " '" + appName + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + startDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + endDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + eventName + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + sortNo + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + bannerType + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + bannerImg + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + bannerLink + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isUsing + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + regUserid + "'" + VbCrlf
    sqlStr = sqlStr + " ,getdate()" + VbCrlf
    sqlStr = sqlStr + " ,'" + workComment + "'" + VbCrlf
    sqlStr = sqlStr + " )" + VbCrlf
    
    dbget.Execute sqlStr
end if

response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>opener.history.go(0); self.close();</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->