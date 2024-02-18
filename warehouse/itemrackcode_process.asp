<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dblogicsopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
dim itemgubun, itemid, itemoption, itemrackcode
dim mode

itemgubun 	= trim(request("itemgubun"))
itemid		= trim(request("itemid"))
itemoption	= trim(request("itemoption"))
itemrackcode = trim(request("itemrackcode"))
mode	= trim(request("mode"))


dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr

if (mode="ByRackCodeProc") then
    if (itemgubun="10") then
    	if (Len(itemrackcode)<>4) or (itemid="") then
    		response.write "<script>alert('상품코드나 랙코드가 입력되지 않았습니다.');</script>"
    		response.write "<script>history.back();</script>"
    		dbget.close()	:	response.End
    	end if

    	sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
    	sqlStr = sqlStr + " set itemrackcode='" + itemrackcode + "'" + VbCrlf
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
    	dbget.execute(sqlStr)


    	sqlStr = "update [db_logics].[dbo].tbl_logics_item" + VbCrlf
    	sqlStr = sqlStr + " set itemrackcode='" + itemrackcode + "'" + VbCrlf
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
    	dblogicsget.execute(sqlStr)
    else
        response.write "<script >alert('현재 오프 상품은 랙코드가 지원되지 않습니다..');</script>"
        response.write "<script >location.replace('" + refer + "');</script>"
        dbget.close()	:	response.End
    end if

end if
%>
<script language='javascript'>
alert('등록 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dblogicsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->