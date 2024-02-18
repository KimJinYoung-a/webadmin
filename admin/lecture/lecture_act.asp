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

dim idx,mode
dim linkitemid,lectitle,lecturer,lecsum,matinclude,matsum
dim leccount,lectime,tottime,matdesc,properperson,minperson
dim reservestart,reserveend,lecdate01,lecdate02,lecdate03
dim lecdate04,lecdate05,lecdate06,lecdate07,lecdate08,lecdate01_end
dim lecdate02_end,lecdate03_end,lecdate04_end,lecdate05_end,lecdate06_end
dim lecdate07_end,lecdate08_end,leccontents,lecetc
dim lecturerid,lecperiod,leccurry,lecspace
dim yyyymm, regfinish
dim isusing

idx = request("idx")
mode = request("mode")

lecturerid = request("lecturerid")
lecperiod = request("lecperiod")
leccurry = request("leccurry")
linkitemid = request("linkitemid")
lectitle = request("lectitle")
lecturer = request("lecturer")
lecsum = request("lecsum")
matinclude = request("matinclude")
matsum = request("matsum")
lecspace = request("lecspace")
leccount = request("leccount")
lectime = request("lectime")
tottime = request("tottime")
matdesc = request("matdesc")
properperson = request("properperson")
minperson = request("minperson")
reservestart = request("reservestart")
reserveend = request("reserveend")
lecdate01 = request("lecdate01")
lecdate02 = request("lecdate02")
lecdate03 = request("lecdate03")
lecdate04 = request("lecdate04")
lecdate05 = request("lecdate05")
lecdate06 = request("lecdate06")
lecdate07 = request("lecdate07")
lecdate08 = request("lecdate08")
lecdate01_end = request("lecdate01_end")
lecdate02_end = request("lecdate02_end")
lecdate03_end = request("lecdate03_end")
lecdate04_end = request("lecdate04_end")
lecdate05_end = request("lecdate05_end")
lecdate06_end = request("lecdate06_end")
lecdate07_end = request("lecdate07_end")
lecdate08_end = request("lecdate08_end")
leccontents = request("leccontents")
lecetc = request("lecetc")
yyyymm = request("yyyymm")
regfinish = request("regfinish")
isusing = request("isusing")

if matinclude = "on" then
 matinclude = "Y"
else
 matinclude = "N"
end if
'==============================================================================

dim sqlStr,iid

if (mode = "add") then

	sqlStr = " insert into [db_contents].[dbo].tbl_lecture_item "
	sqlStr = sqlStr + " (linkitemid,lectitle,lecturerid,lecturer,lecsum,matinclude,matsum,lecspace,leccount,lecperiod,lectime,tottime,matdesc," + vbcrlf
	sqlStr = sqlStr + "properperson,minperson,reservestart,reserveend,lecdate01,lecdate02,lecdate03,lecdate04,lecdate05," + vbcrlf
	sqlStr = sqlStr + "lecdate06,lecdate07,lecdate08,lecdate01_end,lecdate02_end,lecdate03_end,lecdate04_end,lecdate05_end," + vbcrlf
	sqlStr = sqlStr + "lecdate06_end,lecdate07_end,lecdate08_end,leccontents,leccurry,lecetc,isusing,mastercode,regfinish) " + vbcrlf
	sqlStr = sqlStr + " values("
	sqlStr = sqlStr + "" + Cstr(linkitemid) + ","
	sqlStr = sqlStr + "'" + html2db(lectitle) + "',"
	sqlStr = sqlStr + "'" + html2db(lecturerid) + "',"
	sqlStr = sqlStr + "'" + html2db(lecturer) + "',"
	sqlStr = sqlStr + "" + Cstr(lecsum) + ","
	sqlStr = sqlStr + "'" + Cstr(matinclude) + "',"
	sqlStr = sqlStr + "" + Cstr(matsum) + ","
	sqlStr = sqlStr + "'" + html2db(lecspace) + "',"
	sqlStr = sqlStr + "" + Cstr(leccount) + ","
	sqlStr = sqlStr + "'" + html2db(lecperiod) + "',"
	sqlStr = sqlStr + "'" + Cstr(lectime) + "',"
	sqlStr = sqlStr + "'" + Cstr(tottime) + "',"
	sqlStr = sqlStr + "'" + html2db(matdesc) + "',"
	sqlStr = sqlStr + "" + Cstr(properperson) + ","
	sqlStr = sqlStr + "" + Cstr(minperson) + ","
	sqlStr = sqlStr + "'" + Cstr(reservestart) + "',"
	sqlStr = sqlStr + "'" + Cstr(reserveend) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate01) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate02) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate03) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate04) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate05) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate06) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate07) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate08) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate01_end) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate02_end) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate03_end) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate04_end) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate05_end) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate06_end) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate07_end) + "',"
	sqlStr = sqlStr + "'" + Cstr(lecdate08_end) + "',"
	sqlStr = sqlStr + "'" + html2db(leccontents) + "',"
	sqlStr = sqlStr + "'" + html2db(leccurry) + "',"
	sqlStr = sqlStr + "'" + html2db(lecetc) + "',"
	sqlStr = sqlStr + "'" + html2db(isusing) + "',"
	
	sqlStr = sqlStr + "'" + Cstr(yyyymm) + "',"
	sqlStr = sqlStr + "'" + Cstr(regfinish) + "'"
	sqlStr = sqlStr + ")"
'response.write sqlStr
	rsget.Open sqlStr, dbget, 1


	sqlStr = "select IDENT_CURRENT('[db_contents].[dbo].tbl_lecture_item') as iid"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
	iid = rsget("iid")
	end if
	rsget.close
	idx = iid
else

	sqlStr = " update [db_contents].[dbo].tbl_lecture_item " + vbcrlf
	sqlStr = sqlStr + " set linkitemid = " + Cstr(linkitemid) + "," + vbcrlf
	sqlStr = sqlStr + " lectitle = '" + html2db(lectitle) + "'," + vbcrlf
	sqlStr = sqlStr + " lecturerid = '" + html2db(lecturerid) + "'," + vbcrlf
	sqlStr = sqlStr + " lecturer = '" + html2db(lecturer) + "'," + vbcrlf
	sqlStr = sqlStr + " lecsum = " + Cstr(lecsum) + "," + vbcrlf
	sqlStr = sqlStr + " matinclude = '" + Cstr(matinclude) + "'," + vbcrlf
	sqlStr = sqlStr + " matsum = " + Cstr(matsum) + "," + vbcrlf
	sqlStr = sqlStr + " lecspace = '" + html2db(lecspace) + "'," + vbcrlf
	sqlStr = sqlStr + " leccount = " + Cstr(leccount) + "," + vbcrlf
	sqlStr = sqlStr + " lecperiod = '" + html2db(lecperiod) + "'," + vbcrlf
	sqlStr = sqlStr + " lectime = '" + Cstr(lectime) + "'," + vbcrlf
	sqlStr = sqlStr + " tottime = '" + Cstr(tottime) + "'," + vbcrlf
	sqlStr = sqlStr + " matdesc = '" + html2db(matdesc) + "'," + vbcrlf
	sqlStr = sqlStr + " properperson = " + Cstr(properperson) + "," + vbcrlf
	sqlStr = sqlStr + " minperson = " + Cstr(minperson) + "," + vbcrlf
	sqlStr = sqlStr + " reservestart = '" + Cstr(reservestart) + "'," + vbcrlf
	sqlStr = sqlStr + " reserveend = '" + Cstr(reserveend) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate01 = '" + Cstr(lecdate01) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate02 = '" + Cstr(lecdate02) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate03 = '" + Cstr(lecdate03) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate04 = '" + Cstr(lecdate04) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate05 ='" + Cstr(lecdate05) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate06 = '" + Cstr(lecdate06) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate07 = '" + Cstr(lecdate07) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate08 = '" + Cstr(lecdate08) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate01_end = '" + html2db(lecdate01_end) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate02_end = '" + html2db(lecdate02_end) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate03_end = '" + html2db(lecdate03_end) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate04_end = '" + html2db(lecdate04_end) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate05_end = '" + html2db(lecdate05_end) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate06_end = '" + html2db(lecdate06_end) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate07_end = '" + html2db(lecdate07_end) + "'," + vbcrlf
	sqlStr = sqlStr + " lecdate08_end = '" + html2db(lecdate08_end) + "'," + vbcrlf
	sqlStr = sqlStr + " leccontents = '" + html2db(leccontents) + "'," + vbcrlf
	sqlStr = sqlStr + " leccurry = '" + html2db(leccurry) + "'," + vbcrlf
	sqlStr = sqlStr + " lecetc = '" + html2db(lecetc) + "'," + vbcrlf
	sqlStr = sqlStr + " isusing = '" + html2db(isusing) + "'," + vbcrlf
	sqlStr = sqlStr + " mastercode = '" + (yyyymm) + "'," + vbcrlf
	sqlStr = sqlStr + " regfinish = '" + (regfinish) + "'" + vbcrlf

	sqlStr = sqlStr + " where idx = " + CStr(idx) + " "
''response.write sqlStr
	rsget.Open sqlStr, dbget, 1

end if

response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('lecreg.asp?idx=" + Cstr(idx) + "&mode=edit');</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->