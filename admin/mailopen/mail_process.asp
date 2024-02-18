<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메일진 통계
' History : 2008.01.21 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

<%
dim mode, idx,gubun ,title,startdate,enddate,reenddate,totalcnt, mailergubun
Dim opencnt,openpct,noopencnt,noopenpct ,isusing ,i,cnt,sqlStr
Dim realcnt,realpct,filteringcnt,filteringpct ,successcnt,successpct,failcnt,failpct
	mailergubun = request("mailergubun")
	isusing = request("isusing")
	mode = request("mode")
	idx = request("idx")
	gubun = request("gubun")
	title = request("title")
	startdate = request("startdate")
	enddate = request("enddate")
	reenddate = request("reenddate")
	totalcnt = request("totalcnt")
	realcnt = request("realcnt")
	realpct = trim(request("realpct"))
	filteringcnt = request("filteringcnt")
	filteringpct = trim(request("filteringpct"))
	successcnt = request("successcnt")
	successpct = trim(request("successpct"))
	failcnt = request("failcnt")
	failpct = trim(request("failpct"))
	opencnt = request("opencnt")
	openpct = trim(request("openpct"))
	noopencnt = request("noopencnt")
	noopenpct = trim(request("noopenpct"))

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if mode="add" then

	sqlStr = "insert into [db_log].[dbo].tbl_mailing_data(gubun,title,startdate,enddate,reenddate,totalcnt,realcnt,realpct,"
	sqlStr = sqlStr + "filteringcnt,filteringpct,successcnt,successpct,failcnt,failpct,opencnt,openpct,noopencnt,noopenpct, mailergubun)" + vbCrlf
	sqlStr = sqlStr + " values("  + vbCrlf
	sqlStr = sqlStr + "'" + CStr(gubun) + "'," + vbCrlf
	sqlStr = sqlStr + "'" + html2db(title) + "'," + vbCrlf
	sqlStr = sqlStr + "'" + html2db(left(startdate,10) & " " & mid(startdate,14,9)) + "'," + vbCrlf
	sqlStr = sqlStr + "'" + html2db(left(enddate,10) & " " & mid(enddate,14,9)) + "'," + vbCrlf
	sqlStr = sqlStr + "'" + html2db(left(reenddate,10) & " " & mid(reenddate,14,9)) + "'," + vbCrlf
	sqlStr = sqlStr + "" + CStr(totalcnt) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(realcnt) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(realpct) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(filteringcnt) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(filteringpct) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(successcnt) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(successpct) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(failcnt) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(failpct) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(opencnt) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(openpct) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(noopencnt) + "," + vbCrlf
	sqlStr = sqlStr + "" + CStr(noopenpct) + "," + vbCrlf
	sqlStr = sqlStr + "'" + CStr(mailergubun) + "')"
	
	'response.write sqlStr
	dbget.execute sqlStr
	
elseif mode="edit" then

	sqlStr = "update [db_log].[dbo].tbl_mailing_data" + vbCrlf
	sqlStr = sqlStr + "set title='" + html2db(title) + "'," + vbCrlf
	sqlStr = sqlStr + " gubun='" + CStr(gubun) + "'," + vbCrlf
	sqlStr = sqlStr + " startdate='" + html2db(left(startdate,10) & " " & mid(startdate,14,9)) + "'," + vbCrlf
	sqlStr = sqlStr + " enddate='" + html2db(left(enddate,10) & " " & mid(enddate,14,9)) + "'," + vbCrlf
	sqlStr = sqlStr + " reenddate='" + html2db(left(reenddate,10) & " " & mid(reenddate,14,9)) + "'," + vbCrlf
	sqlStr = sqlStr + " totalcnt=" + CStr(totalcnt) + "," + vbCrlf
	sqlStr = sqlStr + " realcnt=" + CStr(realcnt) + "," + vbCrlf
	sqlStr = sqlStr + " realpct=" + CStr(realpct) + "," + vbCrlf
	sqlStr = sqlStr + " filteringcnt=" + CStr(filteringcnt) + "," + vbCrlf
	sqlStr = sqlStr + " filteringpct=" + CStr(filteringpct) + "," + vbCrlf
	sqlStr = sqlStr + " successcnt=" + CStr(successcnt) + "," + vbCrlf
	sqlStr = sqlStr + " successpct=" + CStr(successpct) + "," + vbCrlf
	sqlStr = sqlStr + " failcnt=" + CStr(failcnt) + "," + vbCrlf
	sqlStr = sqlStr + " failpct=" + CStr(failpct) + "," + vbCrlf
	sqlStr = sqlStr + " opencnt=" + CStr(opencnt) + "," + vbCrlf
	sqlStr = sqlStr + " openpct=" + CStr(openpct) + "," + vbCrlf
	sqlStr = sqlStr + " noopencnt=" + CStr(noopencnt) + "," + vbCrlf
	sqlStr = sqlStr + " noopenpct=" + CStr(noopenpct) + "," + vbCrlf
	sqlStr = sqlStr + " isusing='" + CStr(isusing) + "'," + vbCrlf
	sqlStr = sqlStr + " mailergubun='" + CStr(mailergubun) + "'" + vbCrlf	
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	
	'response.write sqlStr
	dbget.execute sqlStr
end If
%>

<script language="javascript">
	alert('OK');
	opener.location.reload();
	self.close();
	//location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->