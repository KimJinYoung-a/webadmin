<%@ Language=VBScript %>
<%	Option Explicit %>
<%	Response.Expires = -1440 %>
<%	Response.CharSet = "euc-kr" %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim outmallitemname : outmallitemname = requestCheckvar(request("outmallitemname"),100)
Dim tenitemid : tenitemid= requestCheckvar(request("tenitemid"),10)
Dim mallid : mallid = requestCheckvar(request("mallid"),32)

Dim sqlStr
Dim retitemid, retCNT, retOptCnt, optARR
retOptCnt=0

IF (tenitemid<>"") then
    sqlStr = "select top 2 itemid from db_item.dbo.tbl_item where itemid='"&(tenitemid)&"'"
ELSE
    sqlStr = "select top 2 itemid from db_item.dbo.tbl_item where itemname='"&html2DB(outmallitemname)&"'"
    ''sqlStr = sqlStr& " and itemid not in (select itemid from db_temp.dbo.tbl_xSite_EtcItemLink where mallid='"&mallid&"')"
END IF
rsget.Open sqlStr,dbget,1
   if Not rsget.Eof then
        retitemid = rsget("itemid")
        retCNT = rsget.RecordCount
   end if
rsget.close

if (retCNT=1) then
    sqlStr = "select itemoption,optionname from db_item.dbo.tbl_item_option where itemid="&retitemid&" and isusing='Y'"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        optARR = rsget.getRows()
    end if
    rsget.close
    
    if IsArray(optARR) then
        retOptCnt = UBound(optARR,2)+1
    end if
end if

dim bufRet, i

if (retCNT=0) then
    bufRet = "<font color=blue>검색상품 없음.</font>"
ELSE
    IF (retOptCnt<1) then
        bufRet = "상품코드:"&retitemid
    ELSE
        bufRet = "상품코드:"&retitemid
        bufRet = bufRet & " <select name='opt' id='vOpt'>"
        bufRet = bufRet & "<option value=''>옵션선택"
        for i=0 to retOptCnt-1
            bufRet = bufRet & "<option value='"&optARR(0,i)&"'>"&optARR(1,i)
        next
        bufRet = bufRet & "</select> "
    END IF
    
    IF (bufRet<>"") then
        bufRet = bufRet & "<input type=button value='선택' onClick='selThisItem("""&retitemid&""");'>"        
    end if
END IF

response.write bufRet
%>


<!-- #include virtual="/lib/db/dbclose.asp" -->