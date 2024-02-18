<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/etc/wmp/wmpCls.asp"-->
<%
Response.CharSet = "euc-kr"
Dim itemid, sqlStr, itemname, optioncnt, limityn, limitno, limitsold
Dim availCnt, optArrRows, i
itemid			= requestCheckVar(request("itemid"),10)
If isNumeric(itemid) = False Then
    rw "상품코드는 숫자만 입력하세요"
    response.write "<script> $(""#itemid"").val('');</script>"
    response.end
End If

sqlStr = ""
sqlStr = sqlStr & " SELECT TOP 1 itemname, optioncnt, limityn, limitno, limitsold "
sqlStr = sqlStr & " FROM db_item.dbo.tbl_item "
sqlStr = sqlStr & " WHERE itemid = '"& itemid &"' "
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
If not rsget.EOF Then
    itemname = rsget("itemname")
    optioncnt = rsget("optioncnt")
    limityn = rsget("limityn")
    limitno = rsget("limitno")
    limitsold = rsget("limitsold")
End If
rsget.Close
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
    if ("<%= itemname %>" != "") {
        $("#itemnameTr").show();
        $("#itemname").val('<%= itemname %>');
        if ("<%= optioncnt %>" == "0") {
             $("#limitCount").val('<%= availCnt %>');
            $("#limitCountTr").show();
        }else{
            $("#limitCountTr").hide();
        }
    }
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->