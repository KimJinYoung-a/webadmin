<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim sMode, strSql
Dim pdiscountKey, pdiscountTitle, pstDT, pedDT, pdiscountPro, pdiscountbuyRule, pdiscountbuyPro, popenDate, pdiscountStatus
sMode				= request("sMode")
pdiscountKey		= request("discountKey")
pdiscountTitle		= request("discountTitle")
pstDT				= request("stDT")
pedDT				= request("edDT")
pdiscountPro		= request("discountPro")
pdiscountbuyRule	= request("discountbuyRule")
pdiscountbuyPro		= request("discountbuyPro")
pdiscountStatus		= request("discountStatus")

If pdiscountStatus = "" Then pdiscountStatus = 0
Select Case sMode
	Case "I"
		If pdiscountStatus = "7" Then
			If pOpenDate = "" Then
				pOpenDate = "getdate()"
			End If
		End If

		If pOpenDate = "" Then pOpenDate = "NULL"

		strSql = ""
		strSql = strSql & " INSERT INTO db_item.dbo.tbl_kaffa_Discount_List " & VBCRLF
		strSql = strSql & " (discountTitle, promotionType, stDT, edDT, discountPro, discountbuyRule, discountbuyPro, regdate, openDate, regUserID) Values " & VBCRLF
		strSql = strSql & " ('"&pdiscountTitle&"', 0, '"&pstDT&"', '"&pedDT&" 23:59:59', '"&pdiscountPro&"', '"&pdiscountbuyRule&"', '"&pdiscountbuyPro&"', getdate(), "&popendate&", '"&session("ssBctId")&"')	"
		dbget.execute strSql
		response.redirect("saleList.asp?menupos="&menupos)
		dbget.close()	:	response.End
	Case "U"
		strSql = ""
		strSql = strSql & " UPDATE db_item.dbo.tbl_kaffa_Discount_List "
		strSql = strSql & " SET discountTitle='"&pdiscountTitle&"'" & VBCRLF
		strSql = strSql & " , stDT='"&pstDT&"'" & VBCRLF
		strSql = strSql & " , edDT= '"&pedDT&" 23:59:59' " & VBCRLF
		strSql = strSql & " , discountPro='"&pdiscountPro&"'" & VBCRLF
		strSql = strSql & " , discountbuyRule= '"&pdiscountbuyRule&"'" & VBCRLF
		strSql = strSql & " , discountbuyPro='"&pdiscountbuyPro&"' " & VBCRLF
		strSql = strSql & " , lastupdate=getdate() , lastUpUserID = '"&session("ssBctId")&"' " & VBCRLF
		if (pdiscountStatus = "9") Then
		    strSql = strSql & " , expireddate=isNULL(expireddate,getdate())"
		elseif (pdiscountStatus = "7") Then
		    strSql = strSql & " , opendate=isNULL(opendate,getdate())"
		end if
		strSql = strSql & " WHERE discountKey = "&pdiscountKey
		dbget.execute strSql
		response.redirect("saleList.asp?menupos="&menupos)
		dbget.close()	:	response.End
End Select
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

