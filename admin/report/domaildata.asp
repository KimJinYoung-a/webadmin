<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메일 통계
' History : 2007.08.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode, idx,gubun
Dim title,startdate,enddate,reenddate,totalcnt
Dim realcnt,realpct,filteringcnt,filteringpct
Dim successcnt,successpct,failcnt,failpct
Dim opencnt,openpct,noopencnt,noopenpct

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

dim i,cnt,sqlStr
dim refer
refer = request.ServerVariables("HTTP_REFERER")

if mode="add" then

	sqlStr = "insert into [db_log].[dbo].tbl_mailing_data(gubun,title,startdate,enddate,reenddate,totalcnt,realcnt,realpct,"
	sqlStr = sqlStr + "filteringcnt,filteringpct,successcnt,successpct,failcnt,failpct,opencnt,openpct,noopencnt,noopenpct)" + vbCrlf
	sqlStr = sqlStr + " values("  + vbCrlf
	sqlStr = sqlStr + "'" + CStr(gubun) + "'," + vbCrlf
	sqlStr = sqlStr + "'" + CStr(title) + "'," + vbCrlf
	sqlStr = sqlStr + "'" + CStr(startdate) + "'," + vbCrlf
	sqlStr = sqlStr + "'" + CStr(enddate) + "'," + vbCrlf
	sqlStr = sqlStr + "'" + CStr(reenddate) + "'," + vbCrlf
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
	sqlStr = sqlStr + "" + CStr(noopenpct) + ")"
	rsget.Open sqlStr, dbget, 1

elseif mode="edit" then

	sqlStr = "update [db_log].[dbo].tbl_mailing_data" + vbCrlf
	sqlStr = sqlStr + "set title='" + CStr(title) + "'," + vbCrlf
	sqlStr = sqlStr + " gubun='" + CStr(gubun) + "'," + vbCrlf
	sqlStr = sqlStr + " startdate='" + CStr(startdate) + "'," + vbCrlf
	sqlStr = sqlStr + " enddate='" + CStr(enddate) + "'," + vbCrlf
	sqlStr = sqlStr + " reenddate='" + CStr(reenddate) + "'," + vbCrlf
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
	sqlStr = sqlStr + " noopenpct=" + CStr(noopenpct) + "" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1

end If

%>

<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->