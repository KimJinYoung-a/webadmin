<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : nb_proc.asp
' Discription : 모바일 사이트 알림배너 처리페이지
' History : 2013.04.02 이종화
'###############################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
dim iidx , utArr , utnArr
Dim startday , endday , sthh , edhh
Dim title , sorting , text , texturl , isusing , mode , infourl
Dim writer 
Dim tempSdate , tempEdate

writer  = session("ssBctCname")

iidx		= Trim(RequestCheckVar(request("iidx"),10))
utArr		= Trim(RequestCheckVar(request("utArr"),20))
utnArr		= Trim(RequestCheckVar(request("utnArr"),50))
startday	= Trim(RequestCheckVar(request("startday"),10))
endday		= Trim(RequestCheckVar(request("endday"),10))
sthh		= Trim(RequestCheckVar(request("sthh"),8))
edhh		= Trim(RequestCheckVar(request("edhh"),8))
title		= Trim(RequestCheckVar(request("title"),30))
sorting		= Trim(RequestCheckVar(request("sorting"),3))
text		= Trim(RequestCheckVar(request("text"),30))
texturl		= Trim(RequestCheckVar(request("texturl"),200))
infourl		= Trim(RequestCheckVar(request("infourl"),20))
isusing		= Trim(RequestCheckVar(request("isusing"),1))

mode		= Trim(RequestCheckVar(request("mode"),10))

''입력용 sdate , edate
tempSdate = startday & " " & sthh
tempEdate = endday & " " & edhh

dim sqlStr, referer

If mode = "add" Then

	sqlStr = " insert into db_sitemaster.dbo.tbl_mobile_noticebanner" + VbCrlf
	sqlStr = sqlStr + " (usertype,usertypename,startdate,enddate,title,sorting,textcontents,texturl,isusing,writer,textcopy)"+ VbCrlf
	sqlStr = sqlStr + " values("
	sqlStr = sqlStr + " '" + utArr +"'" + VbCrlf
	sqlStr = sqlStr + " ,'" + utnArr + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + tempSdate + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + tempEdate + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + title + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + sorting + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + text + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + texturl + "'" + VbCrlf
	sqlStr = sqlStr + " ," + isusing + VbCrlf
	sqlStr = sqlStr + " ,'" + writer + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + infourl + "'" + VbCrlf
	sqlStr = sqlStr + " )" + VbCrlf

	'rw sqlStr
	'response.end

	dbget.Execute sqlStr
ElseIf mode = "chg" Then
	If isusing = "0" Then
		isusing = "1"
	Else
		isusing = "0"
	End If 
	
	sqlStr = " update db_sitemaster.dbo.tbl_mobile_noticebanner " + VbCrlf
    sqlStr = sqlStr + " set  isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " where idx=" + CStr(iidx) + VbCrlf

	dbget.Execute sqlStr
Else 
	writer  = session("ssBctCname")

    sqlStr = " update db_sitemaster.dbo.tbl_mobile_noticebanner " + VbCrlf
    sqlStr = sqlStr + " set usertype='" + utArr + "'" + VbCrlf
    sqlStr = sqlStr + " ,usertypename='" + utnArr + "'" + VbCrlf
    sqlStr = sqlStr + " ,startdate='" + tempSdate + "'" + VbCrlf
    sqlStr = sqlStr + " ,enddate='" + tempEdate + "'" + VbCrlf
    sqlStr = sqlStr + " ,title='" + title + "'" + VbCrlf
    sqlStr = sqlStr + " ,sorting='" + sorting + "'" + VbCrlf
    sqlStr = sqlStr + " ,textcontents='" + text + "'"+ VbCrlf
    sqlStr = sqlStr + " ,texturl='" + texturl + "'" + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ,lastwriter='" + writer + "'" + VbCrlf
    sqlStr = sqlStr + " ,textcopy='" + infourl + "'" + VbCrlf
    sqlStr = sqlStr + " where idx=" + CStr(iidx) + VbCrlf
    
    dbget.Execute sqlStr
End If 

referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->