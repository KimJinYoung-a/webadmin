<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
Response.AddHeader "Content-Disposition","attachment;filename=딜_상품리스트_" & date & hour(now) & minute(now) & ".xls"
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
dim idx : idx = requestCheckvar(request("idx"),10)
dim itemsort : itemsort = requestCheckvar(request("itemsort"),32)
dim strG : strG = requestCheckvar(Request("selG"),10)
dim makerid : makerid = requestCheckvar(request("makerid"),32)
dim itemid : itemid = request("itemid")
dim itemname : itemname = requestCheckvar(request("itemname"),64)
dim sellyn : sellyn = requestCheckvar(request("sellyn"),2)
dim dispCate : dispCate = requestCheckvar(request("disp"),16)
dim iCurrpage : iCurrpage = Request("iC")	'현재 페이지 번호
    
dim cdealGroup, arrGroup
Dim iTotCnt, arrList, intLoop
Dim iPageSize, iDelCnt, oDealItem
Dim iStartPage, iEndPage, iTotalPage, ix, iPerCnt

set cdealGroup = new CDealSelect
cdealGroup.FRectDealCode = idx
arrGroup = cdealGroup.fnGetRootGroup
set cdealGroup = nothing

if itemsort = "" then itemsort = 1
if iCurrpage = "" then iCurrpage = 1
iPageSize = 10000		'한 페이지의 보여지는 열의 수
iPerCnt = 10		'보여지는 페이지 간격

set oDealItem = new CDealItem
oDealItem.FPSize = iPageSize
oDealItem.FRectMasterIDX = idx
oDealItem.FRectMakerid = makerid
oDealItem.FRectItemid = itemid
oDealItem.FRectItemName = itemname
oDealItem.FRectDispCate = dispCate
oDealItem.FCPage = iCurrpage
oDealItem.FESGroup = strG
oDealItem.FESSort = itemsort
oDealItem.FRectSellYN = sellyn
arrList = oDealItem.fnGetDealEventItemNew
iTotCnt = oDealItem.FTotCnt	'전체 데이터  수
iTotalPage = int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<style type="text/css">
br { mso-data-placement:same-cell; }
</style>
</head>
<body>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td>상품ID</td>
        <td>상품명</td>
        <td>판매가</td>
        <td>매입가</td>
        <td>할인율</td>
    </tr>
<%IF isArray(arrList) THEN 
    For intLoop = 0 To UBound(arrList,2)
%>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%=arrList(0,intLoop)%></td>
        <td align="left">&nbsp;<%=db2html(arrList(1,intLoop))%></td>
        <td>
            <%
            Response.Write FormatNumber(arrList(4,intLoop),0)
            '할인가
            if arrList(8,intLoop)="Y" then
                Response.Write "<br><font color=#F08050>(할)" & FormatNumber(arrList(6,intLoop),0) & "</font>"
            end if
            '쿠폰가
            if arrList(9,intLoop)="Y" then
                Select Case arrList(15,intLoop)
                    Case "1"
                        Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(3,intLoop)*((100-arrList(16,intLoop))/100),0) & "</font>"
                    Case "2"
                        Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(3,intLoop)-arrList(16,intLoop),0) & "</font>"
                end Select
            end if
            %>
        </td>
        <td>
            <%
            Response.Write FormatNumber(arrList(5,intLoop),0)
            '할인가
            if arrList(8,intLoop)="Y" then
                Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(7,intLoop),0) & "</font>"
            end if
            '쿠폰가
            if arrList(9,intLoop)="Y" then
                if arrList(15,intLoop)="1" or arrList(15,intLoop)="2" then
                    if arrList(21,intLoop)=0 or isNull(arrList(21,intLoop)) then
                        Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(5,intLoop),0) & "</font>"
                    else
                        Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(21,intLoop),0) & "</font>"
                    end if
                end if
            end if
            %>
        </td>
        <td>
            <%if arrList(8,intLoop)="Y" then%>
            <font color=#F08050><%=CLng(((arrList(4,intLoop)-arrList(6,intLoop))/arrList(4,intLoop))*100)%>%</font>		
            <%end if%>
            <%
            if arrList(9,intLoop)="Y" then 
                if arrList(15,intLoop)="1" or arrList(15,intLoop)="2" then
                    if arrList(21,intLoop)=0 or isNull(arrList(21,intLoop)) then
                        Response.Write "<br><font color=#5080F0>" & FormatNumber( arrList(5,intLoop),0) & "</font>"
                    else
                        Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(16,intLoop),0) 
                        if arrList(15,intLoop)="1" then 
                        Response.Write "%"
                        else
                        Response.Write "원"
                        end if
                        Response.Write "</font>"
                    end if
                end if
            end if
            %>
        </td>
    </tr>
<% Next %>
<% end if %>
</table>
</body>
</html>
<% session.codePage = 949 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->