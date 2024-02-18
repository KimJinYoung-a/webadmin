<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim imod        : imod = RequestCheckVar(request("imod"),10)
Dim OGtitle     : OGtitle = RequestCheckVar(request("OGtitle"),100)
Dim frontopen   : frontopen = RequestCheckVar(request("frontopen"),10)
Dim eCode       : eCode = RequestCheckVar(request("eCode"),10)
Dim menupos     : menupos = RequestCheckVar(request("menupos"),10)
Dim openHtml	: openHtml = RequestCheckVar(request("openHtml"),1500)
Dim openHtmlWeb	: openHtmlWeb = RequestCheckVar(request("openHtmlWeb"),2000)
Dim opengiftType: opengiftType = RequestCheckVar(request("opengiftType"),2)
Dim opengiftScope: opengiftScope = RequestCheckVar(request("opengiftScope"),2)
Dim sqlStr

SELECT CASE imod
    CASE "I" :
         sqlStr = "Insert Into db_event.dbo.tbl_openGift"
         sqlStr = sqlStr & " (event_code, frontOpen, openImage1, openHtml, openHtmlWeb, reguser, opengiftType, opengiftScope)"
         sqlStr = sqlStr & " Values(" & eCode
         sqlStr = sqlStr & " ,'" & frontOpen & "'"
         sqlStr = sqlStr & " ,'" & OGtitle & "'"
         sqlStr = sqlStr & " ,'" & Html2Db(openHtml) & "'"
         sqlStr = sqlStr & " ,'" & Html2Db(openHtmlWeb) & "'"
         sqlStr = sqlStr & " ,'" & session("SsbctId") & "'"
         sqlStr = sqlStr & " ,"&opengiftType&""& VbCrlf
         sqlStr = sqlStr & " ,"&opengiftScope&""& VbCrlf
         sqlStr = sqlStr & " )"

         dbget.Execute sqlStr

         Call sbAlertMsg ("저장되었습니다..", "openGift.asp?menupos="&menupos, "self")
    CASE "E" :
        sqlStr = "update db_event.dbo.tbl_openGift" & VbCrlf
        sqlStr = sqlStr & " Set frontOpen='" & frontOpen & "'" & VbCrlf
        sqlStr = sqlStr & " , openImage1='" & OGtitle & "'" & VbCrlf
        sqlStr = sqlStr & " , openHtml='" & Html2Db(openHtml) & "'" & VbCrlf
        sqlStr = sqlStr & " , openHtmlWeb='" & Html2Db(openHtmlWeb) & "'" & VbCrlf
        sqlStr = sqlStr & " , opengiftType="&opengiftType&""& VbCrlf
        sqlStr = sqlStr & " , opengiftScope="&opengiftScope&""& VbCrlf
        sqlStr = sqlStr & " Where event_code=" & eCode

        dbget.Execute sqlStr

        Call sbAlertMsg ("저장되었습니다..", "openGift.asp?menupos="&menupos, "self")
    CASE "chgScope" :
        sqlStr = "update db_event.dbo.tbl_Gift" & VbCrlf
        sqlStr = sqlStr & " set gift_scope="&opengiftType&""& VbCrlf
        sqlStr = sqlStr & " Where evt_code=" & eCode
        dbget.Execute sqlStr

        Call sbAlertMsg ("저장되었습니다.", "openGift.asp?menupos="&menupos, "self")
END SELECT



%>

<!-- #include virtual="/lib/db/dbClose.asp" -->