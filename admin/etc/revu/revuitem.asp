<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 300

%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/revuitemcls.asp"-->
<%


Dim delim 
delim = VbCrlf



dim oRevutotalpage, oRevuitem,i, buf, optbuf, optstr
dim keywordsStr, keywordsBuf

dim j,totalpage,k
dim maxpage
dim fso, FileName,tFile,appPath
dim readtextfile

dim nowdate
dim adate,bdate
nowdate = now()
adate = CDate(Left(nowdate,10) + " 09:00:00")
bdate = CDate(Left(nowdate,10) + " 23:00:00")

maxpage = 300

appPath = server.mappath("/admin/etc/revu/") + "\"
FileName = "revuitem.xml"

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "REV1-" + ref + "')"
dbget.execute sqlStr

dim TotalCount, versionDate


if ((nowdate<adate) or (nowdate>bdate)) then
    
    sqlStr = "select convert(varchar(19),getdate(),21) as versionDate"
    rsget.Open sqlStr,dbget,1
        versionDate = rsget("versionDate")
    rsget.close

	set oRevutotalpage = new CRevuItem
	oRevutotalpage.FPageSize = 500
	oRevutotalpage.GetAllRevuItemTotalPageRecent

	totalpage = oRevutotalpage.FtotalPage
	TotalCount = oRevutotalpage.FTotalCount
	
	if totalpage>maxpage then totalpage = maxpage
	set oRevutotalpage = Nothing

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )
	tFile.WriteLine "<?xml version='1.0' encoding='EUC-KR' ?>"
	tFile.WriteLine "<root>"
	tFile.WriteLine "<Totalcount>" & CStr(TotalCount) & "</Totalcount>"
	tFile.WriteLine "<versionDate>" & CStr(versionDate) & "</versionDate>"

	for j=0 to totalpage - 1
		set oRevuitem = new CRevuItem
		oRevuitem.FCurrPage = j+1
		oRevuitem.FPageSize = 500
		oRevuitem.GetAllRevuItemListRecent

		optbuf = ""
		optstr = ""
		i =0
		for i=0 to oRevuitem.FResultCount-1
			buf = ""
		
			buf = buf + "<product>" + delim
			buf = buf + "<itemid>" + CStr(oRevuitem.FItemList(i).FItemID) + "</itemid>" + delim
			buf = buf + "<itemname><![CDATA[" + stripHTML(oRevuitem.FItemList(i).FItemName) + "]]></itemname>" + delim
			buf = buf + "<keyWords><![CDATA[" +oRevuitem.FItemList(i).Fkeywords + "]]></keyWords>" + delim
			''buf = buf + "<saleCost>" + CStr(oRevuitem.FItemList(i).FSellcash) + "</saleCost>" + delim
			buf = buf + "<cate1>" + oRevuitem.FItemList(i).Fitemserial_large + "</cate1>" + delim
			buf = buf + "<cate2>" + oRevuitem.FItemList(i).Fitemserial_mid + "</cate2>" + delim
			buf = buf + "<cate3>" + oRevuitem.FItemList(i).Fitemserial_small + "</cate3>" + delim
			buf = buf + "<cate1Nm><![CDATA[" + oRevuitem.FItemList(i).Fitemserial_largeNm + "]]></cate1Nm>" + delim
			buf = buf + "<cate2Nm><![CDATA[" + oRevuitem.FItemList(i).Fitemserial_midNm + "]]></cate2Nm>" + delim
			buf = buf + "<cate3Nm><![CDATA[" + oRevuitem.FItemList(i).Fitemserial_smallNm + "]]></cate3Nm>" + delim
			buf = buf + "<itemLink><![CDATA[" + oRevuitem.FItemList(i).getItemLink + "]]></itemLink>" + delim
			buf = buf + "<imgSrc>" + oRevuitem.FItemList(i).get400Image + "</imgSrc>" + delim
			''buf = buf + "<brand><![CDATA[" + oRevuitem.FItemList(i).Fbrandname + "]]></brand>" + delim
			buf = buf + "<desc><![CDATA[" + oRevuitem.FItemList(i).getItemPreInfodataHTML + oRevuitem.FItemList(i).FItemContent + oRevuitem.FItemList(i).getItemInfoImageHTML + "]]></desc>" + delim
			
			'buf = buf + "<origin><![CDATA[" + oRevuitem.FItemList(i).Fsourcearea + "]]></origin>" + delim
			'buf = buf + "<maker><![CDATA[" + oRevuitem.FItemList(i).Fmakername + "]]></maker>" + delim
			
'				if oRevuitem.FItemList(i).Fvatinclude="N" then
'				    buf = buf + "<freeTaxYn>Y</freeTaxYn>" + delim
'				else
'				    buf = buf + "<freeTaxYn>N</freeTaxYn>" + delim
'			    end if
		
'				buf = buf + "<stdnCost>" + CStr(oRevuitem.FItemList(i).Forgsellcash) + "</stdnCost>" + delim
			
'				if oRevuitem.FItemList(i).IsSoldOut then
'    				buf = buf + "<regDt>20061120</regDt>" + delim
'    				buf = buf + "<vldEnddt>20061120</vldEnddt>" + delim
'				else
'				    buf = buf + "<regDt>20061120</regDt>" + delim
'    				buf = buf + "<vldEnddt>99991231</vldEnddt>" + delim
'				end if
			
            buf = buf + "</product>" + delim

			if buf<>"" then
				tFile.WriteLine buf
			end if
		next
		set oRevuitem = Nothing
	next

	tFile.WriteLine "</root>"

	tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
end if
%>
 

<%
sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "REV2-" + ref + "')"
dbget.execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
response.redirect "/admin/etc/revu/revuitem.xml"
%>