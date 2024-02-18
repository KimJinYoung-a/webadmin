<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<% Response.contentType = "text/xml; charset=euc-kr" %>
<%
	Dim strSql, itemid, objXML, objXMLv, webImgUrl
	itemid = request("itemid")

	if Not(isNumeric(itemid)) Then
		itemid = ""
		Response.end
	End If 

	IF application("Svr_Info")="Dev" THEN
		webImgUrl		= "http://testwebimage.10x10.co.kr"
	else
		webImgUrl		= "http://webimage.10x10.co.kr"
	end if

	if itemid="" then
		dbget.close(): response.End
	end if

	strSql = "Select top 1 itemname, smallImage From db_item.dbo.tbl_item Where itemid=" & itemid
	rsget.Open strSql, dbget, 1

	if Not(rsget.EOF or rsget.BOF) then
		Set objXML = server.CreateObject("Microsoft.XMLDOM")
		objXML.async = False
	
		objXML.appendChild(objXML.createProcessingInstruction("xml","version=""1.0"""))
		objXML.appendChild(objXML.createElement("itemInfo"))

		Set objXMLv = objXML.createElement("item")
			objXMLv.appendChild(objXML.createElement("itemname"))
			objXMLv.appendChild(objXML.createElement("smallImage"))
	
			objXMLv.childNodes(0).appendChild(objXML.createCDATASection("itemname_Cdata"))
			objXMLv.childNodes(1).appendChild(objXML.createCDATASection("smallImage_Cdata"))
	
			objXMLv.childNodes(0).childNodes(0).text = rsget("itemname")
			objXMLv.childNodes(1).childNodes(0).text = chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid) & "/" & rsget("smallImage"),"")

			objXML.documentElement.appendChild(objXMLv.cloneNode(True))
		Set objXMLv = Nothing

		Response.Write objXML.xml

		Set objXML = Nothing
	end if

	rsget.Close()
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->