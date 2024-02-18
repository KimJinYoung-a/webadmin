<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%

dim id, orderserial, buyhp
dim enc_link, validdate
dim title, message, itemname
dim sqlStr

id  		= requestCheckvar(request("id"),32)
orderserial = requestCheckvar(request("orderserial"),32)
buyhp  		= requestCheckvar(request("buyhp"),32)


validdate = Left(DateAdd("d", 8, Now()), 10)
enc_link = TBTEncryptUrl("01," & id & "," & orderserial & "," & validdate)


sqlStr = " select max(d.itemname) as itemname, count(d.itemid) as cnt "
sqlStr = sqlStr & " from "
sqlStr = sqlStr & " 	[db_cs].[dbo].[tbl_new_as_list] a "
sqlStr = sqlStr & " 	join [db_cs].[dbo].[tbl_new_as_list] r on a.refasid = r.id "
sqlStr = sqlStr & " 	join [db_cs].[dbo].[tbl_new_as_detail] d on r.id = d.masterid "
sqlStr = sqlStr & " where "
sqlStr = sqlStr & " 	1 = 1 "
sqlStr = sqlStr & " 	and a.id = " & id
sqlStr = sqlStr & " 	and d.itemid <> 0 "

rsget.CursorLocation = adUseClient
rsget.CursorType = adOpenStatic
rsget.LockType = adLockOptimistic
rsget.Open sqlStr,dbget,1

if  not rsget.EOF  then
    if (rsget("cnt") = 0) then
        itemname = "��ۺ�"
    elseif (rsget("cnt") = 1) then
        itemname = db2html(rsget("itemname"))
    else
        itemname = db2html(rsget("itemname")) & " �� " & (rsget("cnt") - 1) & "��"
    end if
end if
rsget.close


title = "[�ٹ�����] ȯ�Ұ��� �Է¾ȳ�"

message = "�ȳ��ϼ��� ����" & vbCrLf
message = message & "[" & itemname & "] ��ǰ ȯ�� ������" & vbCrLf
message = message & "�������� ����(�Ǵ� ����)�� ȯ���� �����Ǿ� �ȳ��帳�ϴ�." & vbCrLf
message = message & "�Ʒ� ��ũ�� �����Ͻþ� ȯ������(���¹�ȣ,������, �����) �Է� �Ǵ� ���� ��Ź�帳�ϴ�." & vbCrLf
message = message & "�����մϴ�." & vbCrLf
message = message & "" & vbCrLf
message = message & "https://m.10x10.co.kr/my10x10/login.asp?k=" & enc_link & vbCrLf
message = message & "" & vbCrLf
message = message & "�ظ�ũ ��ȿ�Ⱓ�� 1�����Դϴ�." & vbCrLf

%>
<body style="margin:10 10 10 10" bgcolor="#FFFFFF">
<script language='javascript'>

function jsSubmit() {
    var frm = document.frm;
    frm.submit();
}

</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a" bgcolor="#FFFFFF">

<form name="frm" method="post" action="pop_cs_RequestRefundAccountLMS_process.asp">
<input type="hidden" name="mode" value="sendlms">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<tr>
    <td>
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#FFFFFF">
            <td ><img src="/images/icon_star.gif" align="absbottom">&nbsp; <b>ȯ�Ұ��� �Է¾ȳ� : <%= buyhp %></b></td>
        </tr>
        </table>
    </td>
</tr>
</tr>
    <td>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#FFFFFF">
            <td height=25 align=center bgcolor="<%= adminColor("topbar") %>">������ �ڵ���</td>
            <td align=left>
                &nbsp;
                <input type="text" class="text_ro" readonly name="buyhp" value="<%= buyhp %>" size="13">
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td height=25 align=center bgcolor="<%= adminColor("topbar") %>">����</td>
            <td align=left>
                &nbsp;
                <input type="text" class="text_ro" readonly name="title" value="<%= title %>" size="30">
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td height=25 align=center bgcolor="<%= adminColor("topbar") %>">�޽���</td>
            <td align=left>
                &nbsp;
                <textarea class="textarea" name="message" cols="80" rows="15"><%= message %></textarea>
            </td>
        </tr>
        </table>
    </td>
</tr>
</form>
</table>

<p />

<div align="center">
    <input type="button" class="button" value=" �����ϱ� " onClick="jsSubmit()">
</div>

</body>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
