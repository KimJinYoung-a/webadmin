<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 180
%>
<%
'###########################################################
' Description : 원가정산리스트 세금계산서 발급
' History : 2022.08.03 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode, yyyy,mm, menupos, yyyymm, sheetidx, i, finishflag, orgFinishflag, addsql
    mode = requestcheckvar(request.Form("mode"),32)
    yyyy = requestcheckvar(request.Form("yyyy"),4)
    mm = requestcheckvar(request.Form("mm"),2)
    menupos = requestcheckvar(getNumeric(request.Form("menupos")),10)
    sheetidx = requestCheckVar(getNumeric(request("sheetidx")),10)
    finishflag = requestCheckVar(getNumeric(request("finishflag")),10)

i=0
yyyymm = yyyy & "-" & mm

dim refer
    refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, objCmd, returnValue, retErrText

if mode="finishflagone" then
    if sheetidx="" or isnull(sheetidx) then
%>
        <script type='text/javascript'>
            alert('상세idx값이 없습니다.');
            location.replace('<%= refer %>');
        </script>
<%
	    dbget.close() : response.end
    end if
    if finishflag="" or isnull(finishflag) then
%>
        <script type='text/javascript'>
            alert('변경할 세금계산서 상태값이 없습니다.');
            location.replace('<%= refer %>');
        </script>
<%
	    dbget.close() : response.end
    end if

    ' orgFinishflag=""
    ' sqlStr = "select"
    ' sqlStr = sqlStr & " sm.idx, sm.ppMasterIdx, sm.yyyymm, sm.codeList, sm.ppGubun, sm.groupCode"
    ' sqlStr = sqlStr & " , sm.finishflag, sm.taxtype, sm.taxregdate, sm.taxinputdate, sm.taxlinkidx, sm.neotaxno, sm.billsiteCode, sm.eseroEvalSeq"
    ' sqlStr = sqlStr & " from [db_storage].[dbo].[tbl_pp_product_sheet_master] sm with (nolock)"
    ' sqlStr = sqlStr & " where sm.idx="& sheetidx &""

    ' 'response.write sqlStr & "<br>"
    ' rsget.CursorLocation = adUseClient
    ' rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    ' if not rsget.EOF  then
    '     orgFinishflag=rsget("finishflag")
    ' end if
    ' rsget.Close

    sqlStr = "update [db_storage].[dbo].[tbl_pp_product_sheet_master]" & VbCrlf
    sqlStr = sqlStr & " set finishflag="& finishflag &","
    sqlStr = sqlStr & " taxtype='01' where"
    sqlStr = sqlStr & " idx="& sheetidx &""

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlstr

elseif mode="finishflagarr" then
    if finishflag="" or isnull(finishflag) then
%>
        <script type='text/javascript'>
            alert('변경할 세금계산서 상태값이 없습니다.');
            location.replace('<%= refer %>');
        </script>
<%
	    dbget.close() : response.end
    end if

    for i=1 to request.form("check").count
        sheetidx=""
        sheetidx = request.form("check")(i)

        sqlStr = "update [db_storage].[dbo].[tbl_pp_product_sheet_master]" & VbCrlf
        sqlStr = sqlStr & " set finishflag="& finishflag &","
        sqlStr = sqlStr & " taxtype='01' where"
        sqlStr = sqlStr & " idx="& sheetidx &""

        'response.write sqlStr & "<Br>"
        dbget.Execute sqlstr
    next

elseif mode="finishflagall" then
    if finishflag="" or isnull(finishflag) then
%>
        <script type='text/javascript'>
            alert('변경할 세금계산서 상태값이 없습니다.');
            location.replace('<%= refer %>');
        </script>
<%
	    dbget.close() : response.end
    end if

	sqlStr = "update [db_storage].[dbo].[tbl_pp_product_sheet_master]" & VbCrlf
	sqlStr = sqlStr & " set finishflag="& finishflag &","
    sqlStr = sqlStr & " taxtype='01' where"
	sqlStr = sqlStr & " yyyymm='" & yyyymm & "'" & VbCrlf

    'response.write sqlStr & "<Br>"
	dbget.Execute sqlstr

else
%>
    <script type='text/javascript'>
        alert('구분자가 없습니다.');
        location.replace('<%= refer %>');
    </script>
<%
	dbget.close() : response.end
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script type='text/javascript'>
    alert('저장 되었습니다.');
    location.replace('<%= refer %>');
</script>
