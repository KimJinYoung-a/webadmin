<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 900 %>
<%
'###########################################################
' Description : 매출로그 재작성큐 등록
' Hieditor : 2021.09.01 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim sqlStr, i, j, k
dim mode, orderserial, orderserialArr

dim refer
	refer = request.ServerVariables("HTTP_REFERER")


mode = requestCheckvar(request("mode"),32)
orderserial = requestCheckvar(request("orderserial"),6400)


if (mode = "reMakeOrdrQueOn") then
    orderserialArr = Split(orderserial, vbCrLf)

    for i = 0 to UBound(orderserialArr)
        if Trim(orderserialArr(i)) <> "" then
            sqlStr = " if not exists(select top 1 1 from [db_datamart].[dbo].[tbl_order_log_remakeQue] where orderserial = '" & Trim(orderserialArr(i)) & "') "
            sqlStr = sqlStr & " begin "
            sqlStr = sqlStr & "     insert into [db_datamart].[dbo].[tbl_order_log_remakeQue](orderserial, chktype) values('" & Trim(orderserialArr(i)) & "', 999) "
            sqlStr = sqlStr & "     insert into [tendb].db_temp.dbo.tbl_orderSerial_change(orderserial,lastupdate,gubun) values('" & Trim(orderserialArr(i)) & "',getdate(),'MAKEQUE') "
            sqlStr = sqlStr & " end "
            db3_dbget.Execute sqlStr
        end if
    next
else
    response.write "잘못된 접근입니다." : dbget.close : response.end
end if

%>
<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script language='javascript'>
alert('저장되었습니다.');
location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
