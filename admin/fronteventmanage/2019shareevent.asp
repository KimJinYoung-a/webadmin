<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 4�� ���⼼�� - �ѽ�! ������2 ���� ������������
' History : 2019-04-11 ����ȭ
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim sqlStr , mktArr , intLoop

sqlStr = "EXEC db_temp.dbo.usp_WWW_snsshare_mktdata"
rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly	
if not rsget.EOF then
    mktArr = rsget.getRows()	
end if
rsget.close
%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jqueryui/css/jquery-ui.css"/>
<div class="content scrl" style="top:40px;">
	<div class="pad20">
        <table class="tbType1 listTb" style="width:800px;margin-left:50;">
            <colgroup>
                <col width="16%" />
                <col width="16%" />
                <col width="16%" />
                <col width="16%" />
                <col width="16%" />
                <col width="16%" />
            </colgroup>
            <tr align="center" bgcolor="#E6E6E6" height="20">
                <th><strong>��¥</strong></th>
                <th><strong>�ð�</strong></th>
                <th><strong>�̹��� ���� ����</strong></th>
                <th><strong>�̹��� ���� Ƚ��</strong></th>
                <th><strong>�����ڼ�</strong></th>
                <th><strong>����Ƚ��</strong></th>
            </tr>
            <% IF isArray(mktArr) THEN %>
            <% For intLoop = 0 To UBound(mktArr,2) %>
            <tr>
                <td><%=mktArr(0,intLoop)%></td>
                <td><%=mktArr(1,intLoop)%></td>
                <td><%=mktArr(2,intLoop)%></td>
                <td><%=mktArr(3,intLoop)%></td>
                <td><%=mktArr(4,intLoop)%></td>
                <td><%=mktArr(5,intLoop)%></td>
            </tr>
            <% Next %>
            <% End If %>
        </table>
    </div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->