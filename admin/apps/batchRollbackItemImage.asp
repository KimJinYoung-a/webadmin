<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���δ���ǰ �̹����� �ǻ�ǰ �̹����� �ϰ� ����
' History : 2021.11.27 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%
dim procCnt, sqlStr, arrItemid, i, sCnt, retVal, strRst
sCnt = 0: strRst=""

procCnt = requestCheckVar(request("cnt"),8)
if procCnt="" then procCnt=50

    '' ���ť���� ��ǰ ���̵� ����
    sqlStr = " select top " & procCnt & " itemid "
    sqlStr = sqlStr + " from db_temp.dbo.temp_rollbackItems "
    sqlStr = sqlStr + " where status=0 "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        arrItemid = rsget.getRows()
    end if
    rsget.close

    if isArray(arrItemid) then

        for i=0 to ubound(arrItemid,2)
            retVal = SendReq(ItemUploadUrl & "/linkweb/items/rollbackItemimageFromWaitItem.asp","itemid=" & arrItemid(0,i) & "&adid=system&sell=Y")
            'retVal = SendReq("https://upload.10x10.co.kr/linkweb/items/rollbackItemimageFromWaitItem.asp","itemid=" & arrItemid(0,i) & "&adid=system&sell=Y")
            if retVal="OK" then
                sCnt = sCnt+1
                sqlStr = "Update db_temp.dbo.temp_rollbackItems set status=7 where itemid=" & arrItemid(0,i)
                dbget.Execute sqlStr
            else
                strRst = strRst & chkIIF(strRst="",""," | ") & retVal & "(" & arrItemid(0,i) & ")"
                sqlStr = "Update db_temp.dbo.temp_rollbackItems set status=9 where itemid=" & arrItemid(0,i)
                dbget.Execute sqlStr
            end if
            response.Write i &"."& arrItemid(0,i) & "<br />"
            response.flush
        next

        Response.Write (ubound(arrItemid,2)+1) & "�� �� " & sCnt  & "�� ����"
        if strRst<>"" then
            Response.Write "<br />" & strRst
        end if

    else
        Response.Write "No More Data" & strRst
    end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
