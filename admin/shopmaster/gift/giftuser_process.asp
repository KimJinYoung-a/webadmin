<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ����ǰ���ۼ�,���񽺹߼�
' History : 2019.11.07 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim mode, menupos, i, gift_code, sqlStr
	gift_code = requestCheckVar(getNumeric(Request("gift_code")),10)
    menupos = requestCheckVar(getNumeric(Request("menupos")),10)
    mode = requestCheckVar(Request("mode"),32)

dim ref
    ref = request.ServerVariables("HTTP_REFERER")

if gift_code="" then
    response.write "<script type='text/javascript'>"
    response.write "    alert('����ǰ �ڵ尡 �����ϴ�.');"
    response.write "    history.back();"
    response.write "</script>"
    dbget.close()	:	response.End
end if

dim ievt_code : ievt_code=0
sqlStr = " select top 1 event_code from db_event.dbo.tbl_openGift O" & VbCRLF
sqlStr = sqlStr & " Join db_event.dbo.tbl_Event E" & VbCRLF
sqlStr = sqlStr & " on O.event_code=E.evt_code" & VbCRLF
sqlStr = sqlStr & " Join db_event.dbo.tbl_Gift G" & VbCRLF
sqlStr = sqlStr & " on E.evt_code=G.evt_code" & VbCRLF
sqlStr = sqlStr & " where G.gift_code="&gift_code

rsget.Open sqlStr,dbget,1
if not rsget.Eof then
    ievt_code = rsget("event_code")
end if
rsget.Close

' ��ü���� �̺�Ʈ�ΰ�� �����Ұ�
if (ievt_code<>0) then
    response.write "<script>alert('��ü �����̺�Ʈ "&ievt_code&" ����ǰ�ڵ� ("&gift_code&") ���ۼ� �Ұ�');</script>"
    response.write "��ü �����̺�Ʈ "&ievt_code&" ����ǰ�ڵ� ("&gift_code&") ���ۼ� �Ұ�"
    dbget.close()	:	response.End
end if

if mode="giftremakebefore" then
    sqlStr = "exec [db_order].[dbo].[sp_Ten_order_giftuser_list] 1000, 1, "& gift_code &", '', '', '', '', '', 'B'"

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

    response.write "<script type='text/javascript'>"
    response.write "    alert('��� ���� ����ǰ ���ۼ��� �Ϸ� �Ǿ����ϴ�.');"
    response.write "    location.replace('" & ref & "');"
    response.write "</script>"
    dbget.close()	:	response.End

elseif mode="giftremakeafter" then
    sqlStr = "exec [db_order].[dbo].[sp_Ten_order_giftuser_list] 1000, 1, "& gift_code &", '', '', '', '', '', 'A'"

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

    response.write "<script type='text/javascript'>"
    response.write "    alert('��� ���� ����ǰ ���񽺹߼��� �Ϸ� �Ǿ����ϴ�.');"
    response.write "    location.replace('" & ref & "');"
    response.write "</script>"
    dbget.close()	:	response.End

else
    response.write "<script type='text/javascript'>"
    response.write "    alert('�������� ��ΰ� �ƴմϴ�.');"
    response.write "    history.back();"
    response.write "</script>"
    dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->