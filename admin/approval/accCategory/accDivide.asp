<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ī�װ� ����Ʈ
' History : 2012.08.07 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/accCategoryCls.asp"-->
<%
Dim clsAcc
Dim acccd
dim accnm, issale10x10, issalepartner, sdivide, sdividedesc
acccd =  requestCheckvar(Request("acccd"),15)
Set clsAcc = new CAccCategory
clsAcc.FACCCD = acccd
clsAcc.fnGetACCDivData
accnm = clsAcc.FACCNM
issale10x10 = clsAcc.FSale10x10
issalepartner = clsAcc.FSalePartner
sdivide = clsAcc.FDivide
sdividedesc = clsAcc.FDividedesc
Set clsAcc = nothing
%>
<form name="frmDiv" method="post" action="procCategory.asp">
<input type="hidden" name="hidM" value="C">
<input type="hidden" name="hidacc" value="<%=acccd%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td bgcolor="<%= adminColor("tabletop") %>">��������</td>
    <td bgcolor="#ffffff"><%=accnm%></td>
</tr>
<tr>
    <td bgcolor="<%= adminColor("tabletop") %>">����ó</td>
    <td bgcolor="#ffffff">
    <input type="checkbox" name="isS10" value="1" <%if issale10x10 then%>checked<%end if%>> 10x10
    <input type="checkbox" name="isSP" value="1" <%if issalepartner then%>checked<%end if%>> ����
    </td>
</tr>
<tr>
    <td bgcolor="<%= adminColor("tabletop") %>">�Ⱥб���</td>
    <td bgcolor="#ffffff">
        <select class="select" name="sdivide">
        <option value="�ֹ���ȣ" <%if sdivide ="�ֹ���ȣ" then%>selected<%end if%>>�ֹ���ȣ (��ǰ�ڵ� ���� ī�װ��� �Ⱥ�)</option>
        <option value="��۰Ǽ�" <%if sdivide ="��۰Ǽ�" then%>selected<%end if%>>��۰Ǽ�</option>
        </select>
    </td>
</tr>
<tr>
    <td bgcolor="<%= adminColor("tabletop") %>">�Ⱥб��� ����</td>
    <td bgcolor="#ffffff">
    <textarea name="sdividedesc" style="width:400px;height:50px;" > <%=sdividedesc%></textarea>
   
    </td>
</tr>
</table>
<div style="text-align:center;padding:10px"><input type="button" class="button" value="���" onClick="javascript:document.frmDiv.submit();" style="width:200px;height:30px;"></div>
</form>