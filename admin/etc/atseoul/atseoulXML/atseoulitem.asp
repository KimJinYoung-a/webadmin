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
<!-- #include virtual="/lib/classes/items/atseoul_extsiteitemcls.asp"-->
<%

''墨抛绊府 沥狼 刚历 秦具窃
'response.redirect "/admin/dnshop/dnshopitem.xml"
'response.end




Dim delim 
delim = VbCrlf



dim oAtSeoultotalpage, oAtSeoulitem,i, k, buf, optbuf, optstr, vTemp
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

appPath = server.mappath("/admin/etc/atseoul/atseoulXML/") + "\"
FileName = "atseoulitem.xml"

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

''可记包访
dim IsOptionExists, NotSoldOutOptionExists
dim IsTheLastOption

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "ATS1-" + ref + "')"
dbget.execute sqlStr

if (TRUE) then 
'if ((nowdate<adate) or (nowdate>bdate)) then
	set oAtSeoultotalpage = new CTTLItem
	oAtSeoultotalpage.FPageSize = 1000
	oAtSeoultotalpage.GetAllAtSeoulItemTotalPage

	totalpage = oAtSeoultotalpage.FtotalPage
	if totalpage>maxpage then totalpage = maxpage
	set oAtSeoultotalpage = Nothing

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )
	tFile.WriteLine "<?xml version='1.0' encoding='EUC-KR' ?>"
	tFile.WriteLine "<goods>"
	for j=0 to totalpage - 1
		set oAtSeoulitem = new CTTLItem
		oAtSeoulitem.FCurrPage = j+1
		oAtSeoulitem.FPageSize = 1000
		oAtSeoulitem.GetAllAtSeoulItemList4

		optbuf = ""
		optstr = ""
		i = 0
		k = 1
		
		
		for i=0 to oAtSeoulitem.FResultCount-1
		
			''可记List ---
			IsTheLastOption = false
			
			If vTemp = oAtSeoulitem.FItemList(i).FItemID Then
				k = k + 1
			Else
				k = 1
			End If
			
			if (oAtSeoulitem.FItemList(i).Fitemoption="0000") then
			    IsOptionExists = false
    			IsTheLastOption = true
    			optstr = ""
			else
			    IsOptionExists = true
			    
				if (i+1<=oAtSeoulitem.FResultCount-1) then
					if (oAtSeoulitem.FItemList(i).FItemID=oAtSeoulitem.FItemList(i+1).FItemID) then
					    if (Not oAtSeoulitem.FItemList(i).IsOptionSoldOut) then
    						'optbuf = optbuf + "<it_opt"&k&"_subject><![CDATA[" + oAtSeoulitem.FItemList(i).FOptionTypeName + "]]></it_opt"&k&"_subject>" 
    						'optbuf = optbuf + "<it_opt"&k&"><![CDATA[" + (oAtSeoulitem.FItemList(i).FItemOptionName) + "]]></it_opt"&k&">" 
    						optbuf = optbuf + "" + (oAtSeoulitem.FItemList(i).FItemOptionName) + "" + delim
    						NotSoldOutOptionExists = true
						else
							k = k - 1
    					end if
					else
					    if (Not oAtSeoulitem.FItemList(i).IsOptionSoldOut) then
    					    'optbuf = optbuf + "<it_opt"&k&"_subject><![CDATA[" + oAtSeoulitem.FItemList(i).FOptionTypeName + "]]></it_opt"&k&"_subject>" 
    						'optbuf = optbuf + "<it_opt"&k&"><![CDATA[" + (oAtSeoulitem.FItemList(i).FItemOptionName) + "]]></it_opt"&k&">" 
    						optbuf = optbuf + "" + (oAtSeoulitem.FItemList(i).FItemOptionName) + "" + delim
    						NotSoldOutOptionExists = true
						else
							k = k - 1
    					end if
    					IsTheLastOption = true
						optstr = optbuf
						
						optbuf = ""
					end if
				elseif (i=oAtSeoulitem.FResultCount-1) then
				    if (Not oAtSeoulitem.FItemList(i).IsOptionSoldOut) then
    				    'optbuf = optbuf + "<it_opt"&k&"_subject><![CDATA[" + oAtSeoulitem.FItemList(i).FOptionTypeName + "]]></it_opt"&k&"_subject>" 
    					'optbuf = optbuf + "<it_opt"&k&"><![CDATA[" + (oAtSeoulitem.FItemList(i).FItemOptionName) + "]]></it_opt"&k&">" 
    					optbuf = optbuf + "" + (oAtSeoulitem.FItemList(i).FItemOptionName) + "" + delim
                        NotSoldOutOptionExists = true
                        
                        IsTheLastOption = true
					else
						k = k - 1
    				end if
    				
                    optstr = optbuf
                    
					optbuf = ""
				end if
			end if

			buf = ""
            keywordsStr = ""
            
            if (optstr<>"") and (optstr<>delim) then
                IsTheLastOption = true
            end if
            
            
			if (Not IsOptionExists) or (IsTheLastOption) then
				
				buf = buf + "<item>" + delim
				buf = buf + "<it_ot_id>" + CStr(oAtSeoulitem.FItemList(i).FItemID) + "</it_ot_id>" + delim
				buf = buf + "<ca_id>" + CStr(oAtSeoulitem.FItemList(i).Fatseoulcategory) + "</ca_id>" + delim
				buf = buf + "<it_name><![CDATA[" + CStr(oAtSeoulitem.FItemList(i).FItemName) + "]]></it_name>" + delim
				buf = buf + "<it_model><![CDATA[]]></it_model>" + delim
				buf = buf + "<it_origin><![CDATA[" + oAtSeoulitem.FItemList(i).Fsourcearea + "]]></it_origin>" + delim
				buf = buf + "<it_material><![CDATA[" + oAtSeoulitem.FItemList(i).Fitemsource + "]]></it_material>" + delim
				buf = buf + "<it_mini_weight>" + CStr(oAtSeoulitem.FItemList(i).Fitemweight) + "</it_mini_weight>" + delim
				buf = buf + "<it_amount>" + CStr(oAtSeoulitem.FItemList(i).FSellcash) + "</it_amount>" + delim
				buf = buf + "<it_stock_qty>" + CStr(oAtSeoulitem.FItemList(i).Fstockqty) + "</it_stock_qty>" + delim
				
				If optstr = "" Then
	    			buf = buf + "<it_opt1_subject><![CDATA[]]></it_opt1_subject>"  + delim
	    			buf = buf + "<it_opt1><![CDATA[]]></it_opt1>" + delim
				Else
					buf = buf + "<it_opt1_subject><![CDATA[" + CStr(oAtSeoulitem.FItemList(i).FOptionTypeName) + "]]></it_opt1_subject>" + delim
					buf = buf + "<it_opt1><![CDATA[" + delim
					buf = buf + optstr
					buf = buf + "]]></it_opt1>" + delim
				End If
			    
				buf = buf + "<it_limg1>" + oAtSeoulitem.FItemList(i).Fbasicimage + "</it_limg1>" + delim
				buf = buf + "<it_limg2>" + oAtSeoulitem.FItemList(i).FInfoImage1 + "</it_limg2>" + delim
				buf = buf + "<it_limg3>" + oAtSeoulitem.FItemList(i).FInfoImage2 + "</it_limg3>" + delim
				buf = buf + "<it_limg4>" + oAtSeoulitem.FItemList(i).FInfoImage3 + "</it_limg4>" + delim
				buf = buf + "<it_limg5>" + oAtSeoulitem.FItemList(i).FInfoImage4 + "</it_limg5>" + delim
				
				buf = buf + "<it_explan><![CDATA[" + oAtSeoulitem.FItemList(i).getItemPreInfodataHTML + oAtSeoulitem.FItemList(i).FItemContent + oAtSeoulitem.FItemList(i).getItemInfoImageHTML + "]]></it_explan>" + delim

				If oAtSeoulitem.FItemList(i).Flimityn = "Y" AND optstr = "" Then
					buf = buf + "<it_use>0</it_use>" + delim
				Else
					buf = buf + "<it_use>" + oAtSeoulitem.FItemList(i).Fsellyn + "</it_use>" + delim
				End IF
                buf = buf + "</item>" + delim
				optstr = ""
				
				NotSoldOutOptionExists = false
			end if

			if buf<>"" then
				tFile.WriteLine buf
			end if
			
			vTemp = oAtSeoulitem.FItemList(i).FItemID
		next
		set oAtSeoulitem = Nothing
	next

	tFile.WriteLine "</goods>"

	tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
'end if
end if
%>
 

<%
sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "ATS2-" + ref + "')"
dbget.execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
'response.redirect "/admin/etc/atseoul/atseoulXML/" & FileName
%>
xml 积己 肯丰.