<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "Cache-Control","no-cache,must-revalidate"

'###############################################
' PageName : ajaxGetItemInfo.asp
' Discription : ��ǰ ���� �Ѱ� ��ȸ
' History : 2021.02.04 ������
'###############################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->

<%
Dim return_object : Set return_object = jsObject()
Dim itemid : itemid = requestCheckVar(Request("itemid"),10)

If itemid="" Then
    return_object("result") = False
    return_object("message") = "��ǰID�� �Է� �� �ּ���"
    return_object.Flush
	dbget.close()	:	Response.End
End If

Dim strSql, resultData
    strSql =    "SELECT" + vbcrlf _
                    & "  itemid" + vbcrlf _
                    & ", itemname" + vbcrlf _
                    & ", basicimage" + vbcrlf _
                & "FROM db_item.dbo.tbl_item" + vbcrlf _
                & "WHERE itemid = '" & itemid & "'"
    rsget.Open strSql,dbget
        IF not rsget.EOF THEN
            return_object("result") = True
            return_object("itemid") = rsget("itemid")
            return_object("itemname") = rsget("itemname")
            return_object("itemimage") = rsget("basicimage")
        Else
            return_object("result") = False
            return_object("message") = "�Է��Ͻ� ID�� �ش��ϴ� ��ǰ�� �����ϴ�."
        End IF
    return_object.Flush
    rsget.Close

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->