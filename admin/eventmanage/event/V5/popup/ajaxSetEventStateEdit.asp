<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %>
<%
'###############################################
' PageName : ajaxSetEventStateEdit.asp
' Discription : �̺�Ʈ ��ϴ�� ���� ���� ���μ���
' History : 2021.12.08 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
	dim eCode, arrSale, igiftcnt, strSql, giftNsaleYN, oJson
    eCode = requestCheckVar(Request.Form("eC"),10)
    giftNsaleYN = False

    'object �ʱ�ȭ
    Set oJson = jsObject()

	dbget.beginTrans

        strSql =" SELECT count(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code=" & eCode & " AND gift_using='Y'"
        rsget.Open strSql, dbget
        IF not (rsget.EOF or rsget.BOF) THEN
            igiftcnt = rsget(0)
        END IF
        rsget.close

        strSql = " SELECT sale_code, sale_status FROM [db_event].[dbo].[tbl_sale] WHERE evt_code=" & eCode & " AND sale_using=1"
        rsget.Open strSql, dbget
        IF not (rsget.EOF or rsget.BOF) THEN
            arrSale = rsget.getRows()
        END IF
        rsget.close

        IF isarray(arrSale) or igiftcnt>0 THEN
            giftNsaleYN = True
        END IF

        strSql = "UPDATE [db_event].[dbo].[tbl_event]"
        strSql = strSql + " SET evt_startdate=CONVERT(VARCHAR(10),GETDATE(),21)"
        strSql = strSql + " , evt_enddate=DATETIMEFROMPARTS(YEAR(DATEADD(DD,1,GETDATE())),MONTH(DATEADD(DD,1,GETDATE())),DAY(DATEADD(DD,1,GETDATE())),23,59,29,0)"
        strSql = strSql + " , evt_state=0"
        strSql = strSql + " , closedate=NULL"
        strSql = strSql + " WHERE evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            oJson("response") = "err"
            oJson("message") = "������ ó���� ������ �߻��Ͽ����ϴ�"
            oJson.flush
            Set oJson = Nothing
            dbget.close() : Response.End
        end if
    '===========================================================
	dbget.CommitTrans
    if giftNsaleYN then
        oJson("response") = "OK"
        oJson("message") = "��ϴ�� ���� �ݿ��߽��ϴ�. ����ǰ �� ���� ���� �̺�Ʈ �Դϴ�. Ȯ�� �� �������ּ���."
    else
        oJson("response") = "OK"
        oJson("message") = "��ϴ�� ���� �ݿ��߽��ϴ�."
    end if
    oJson.flush
    Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->