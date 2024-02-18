<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
	dim eCode, menudiv, menuidx, strSql
    eCode = requestCheckVar(Request.Form("eC"),10)
    menudiv = requestCheckVar(Request.Form("menudiv"),2)

    strSql = "select idx FROM [db_event].[dbo].[tbl_event_multi_contents_master]" & vbCrlf
    strSql = strSql + " where evt_code=" & eCode
    strSql = strSql + " and menudiv=6"
    rsget.CursorLocation = adUseClient
    rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
    IF not rsget.EOF THEN
        menuidx = rsget("idx")
    End IF
    rsget.Close

if menuidx<>"" then
	response.write menuidx
	dbget.close()	:	response.End
else
	dbget.beginTrans
        '===========================================================
        '--2.disply 수정
        strSql = "INSERT INTO [db_event].[dbo].[tbl_event_multi_contents_master]"
        strSql = strSql + "(evt_code, menudiv, isusing, BGColorLeft)" & vbCrlf
        strSql = strSql + " values(" & eCode & "," & menudiv & ",'Y','#FFFFFF')" & vbCrlf
        dbget.execute strSql

        strSql = "select SCOPE_IDENTITY()"
        rsget.Open strSql, dbget, 0
        menuidx = rsget(0)
        rsget.Close

        strSql = "IF EXISTS(SELECT evt_code FROM [db_event].[dbo].[tbl_event_top_slide_addimage] WHERE evt_code=" & eCode & ")"
        strSql = strSql + "    BEGIN"
        strSql = strSql + "        UPDATE [db_event].[dbo].[tbl_event_top_slide_addimage]"
        strSql = strSql + "        SET menuidx=" & menuidx
        strSql = strSql + "        WHERE evt_code=" & eCode
        strSql = strSql + "    END"
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans
    
	response.write menuidx
	dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->