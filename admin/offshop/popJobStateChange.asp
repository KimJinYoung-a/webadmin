<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->
<%
dim idx : idx = requestCheckvar(request("idx"),9)
dim mode : mode = requestCheckvar(request("mode"),16)
dim jobkey : jobkey = requestCheckvar(request("jobkey"),9)
dim nextState : nextState = requestCheckvar(request("nextState"),9)

dim sqlStr
if (mode="editState") then
    if (nextState<>"") and (idx<>"") and (jobkey<>"") then
        sqlStr = "update [db_shop].[dbo].tbl_shop_tempstock_master" &VbCrlf
        sqlStr = sqlStr & " set jobstate='" & nextState & "'" &VbCrlf
        sqlStr = sqlStr & " where jobkey=" & jobkey
        
        dbget.Execute sqlStr
        
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"
        response.end
    end if
end if


dim oshopBatch
set oshopBatch = new CShopOrder
oshopBatch.FRectIdx = idx
oshopBatch.GetOneShopBatchOrder


%>
<script language='javascript'>
function frmSubmit(){
    if (confirm('���¸� ���� �Ͻðڽ��ϱ�?')){
        frmAct.submit();
    }
}
</script>
<% if (oshopBatch.FREsultCount>0) then %>
<table width="360" border="0" cellspacing="1" cellpadding="3" bgcolor="#3d3d3d" class=a>
<form name="frmAct" method="post" action="">
<input type="hidden" name="mode" value="editState">

<tr  bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" width="100">�������</td>
    <td ><%= oshopBatch.FOneItem.GetJobStateName %></td>
</tr>
<tr  bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" width="100">�������</td>
    <td >
        <% if oshopBatch.FOneItem.FjobState="3" then %>
        <input type="hidden" name="jobkey" value="<%= oshopBatch.FOneItem.Fjobkey %>">
        <input type="radio" name="nextState" value="7" checked >
        ����ľǿϷ�
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center" colspan="2">
        <% if oshopBatch.FOneItem.FjobState="3" then %>
        <input type="button" value="����" class="button" onclick="frmSubmit();">
        <% else %>
        <input type="button" value="�ݱ�" class="button" onclick="window.close();">
        <% end if %>
    </td>
</tr>
</form>
</table>
<% end if %>
<%
set oshopBatch = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->