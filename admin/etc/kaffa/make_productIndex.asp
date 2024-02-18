<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 900
 
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/kaffa/kaffaCls.asp"-->
<%

Dim delim, manageUrl
delim = VbCrlf

IF application("Svr_Info") = "Dev" THEN
	manageUrl = "http://testwebadmin.10x10.co.kr"
Else
	manageUrl = "http://webadmin.10x10.co.kr"
End If


dim oKaffatotalpage, oKaffaitem,i, k, buf, optbuf, optstr, vTemp, arrList, intLoop
dim keywordsStr, keywordsBuf

dim j,totalpage
dim maxpage
dim fso, FileName,tFile,appPath
dim readtextfile

dim nowdate
dim adate,bdate
nowdate = now()
adate = CDate(Left(nowdate,10) + " 09:00:00")
bdate = CDate(Left(nowdate,10) + " 23:00:00")

maxpage = 300

appPath = server.mappath("/admin/etc/kaffa/xml") + "\"
FileName = "ProductIndex.xml"

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

dim IsTheLastOption

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "ATS1-" + ref + "')"
'dbget.execute sqlStr

if (TRUE) then 
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )
	tFile.WriteLine "<?xml version=""1.0"" encoding=""utf-8""?>"
	tFile.WriteLine "<data>"
	tFile.WriteLine "	<version>1.0</version>"
	tFile.WriteLine "	<modified>" & date() & " " & TwoNumber(hour(now)) & ":" & TwoNumber(minute(now)) & ":" & TwoNumber(second(now)) & "</modified>"
	tFile.WriteLine "	<api_key>$2a$08$ik.RQbF9tGCZibk7JnPueuG/8AIeuTDd.lgCP/fYuuZX7dnNuJRe6</api_key>"
	tFile.WriteLine "	<product_dir>" & manageUrl & "/admin/etc/kaffa/item_xml.asp?itemid=</product_dir>"
	tFile.WriteLine "	<extension></extension>"
	tFile.WriteLine "	<insert>"

	set oKaffaitem = new cKaffaItem
	oKaffaitem.FRectUseYN = "n"
	arrList = oKaffaitem.GetMakeProducIndexItemList
	set oKaffaitem = Nothing
	
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
			tFile.WriteLine "		<product_id>" & arrList(0,intLoop) & "</product_id>"
		Next
	End If

	tFile.WriteLine "	</insert>"
	tFile.WriteLine "</data>"

	tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
end if
%>


<%
sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "ATS2-" + ref + "')"
'dbget.execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
response.redirect "/admin/etc/kaffa/xml/" & FileName
%>