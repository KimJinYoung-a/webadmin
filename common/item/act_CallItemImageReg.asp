<%@language="VBScript" %>
<% option explicit %>
<%
response.charset = "euc-kr"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' Discription : ��ǰ �̹��� �ڵ���� API ȣ��
' History : 2019.08.29 ������ : �ű� ����
'###############################################

dim oXML, itemid, sRst, pRst

itemid = requestCheckVar(request("itemid"),10)

Set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")

oXML.open "POST", "http://xapi.10x10.co.kr:8080/scmapi/image/imageprocessing", false
oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
oXML.send "itemid=" & itemid

sRst = BinaryToText(oXML.ResponseBody,"utf-8")  '��� ���� �� TEXT ��ȯ

Set oXML = Nothing

if sRst<>"" then
    set pRst = JSON.parse(sRst)

    if pRst.success then
        response.Write "OK"
    else
        'response.Write sRst
        response.Write pRst.message
    end if
end if
%>