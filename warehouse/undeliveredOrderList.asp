<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [LOG]��������>>[CENTER]���� > �̹�� �ֹ� ���
' History : 2020.11.20 ������
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/undeliveredOrderCls.asp"-->
<%
Dim yyyymmdd, yyyymmdd2, mode, ordstat, danjong, i

yyyymmdd	= requestCheckvar(request("yyyymmdd"),10)
yyyymmdd2	= requestCheckvar(request("yyyymmdd2"),10)
mode	    = requestCheckvar(request("mode"),2)
ordstat     = requestCheckvar(request("ordstat"),1)
danjong     = requestCheckvar(request("danjong"),1)
if yyyymmdd 	= "" then yyyymmdd=dateadd("d",-3,date())     'D+3
if yyyymmdd2 	= "" then yyyymmdd2=dateadd("d",-3,date())     'D+3
if mode 		= "" then mode="OD"   'OD:�ֹ��ϱ���, IT:��ǰ����

dim oOrder
set oOrder = new COrder
    oOrder.FRectDate = yyyymmdd
    oOrder.FRectDate2 = yyyymmdd2
    oOrder.FRectMode = mode
    oOrder.FRectStat = ordstat
    oOrder.FRectDanjong = danjong
    oOrder.OrderList()
%>
<script type="text/javascript" src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
th {
    position: sticky; top: 0; background:<%= adminColor("tabletop") %>;
    border-bottom:1px solid <%= adminColor("tablebg") %>;
}
.txtct {text-align:center;}
.txtrt {text-align:right;}
</style>
<script type="text/javascript">
function SubmitFrm() {
    var frm = document.frm;
    frm.submit();
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
    	�������� :
        <input id="baseDate" name="yyyymmdd" value="<%=yyyymmdd%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="baseDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
        ~
        <input id="baseDate2" name="yyyymmdd2" value="<%=yyyymmdd2%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="baseDate2_trigger" border="0" style="cursor:pointer" align="absmiddle" />
        <script type="text/javascript">
            var CAL_baseDate = new Calendar({
                inputField : "baseDate",
                trigger    : "baseDate_trigger",
                bottomBar: true,
                dateFormat: "%Y-%m-%d",
                onSelect: function() {
                    this.hide();
                }
            });
            var CAL_baseDate2 = new Calendar({
                inputField : "baseDate2",
                trigger    : "baseDate2_trigger",
                bottomBar: true,
                dateFormat: "%Y-%m-%d",
                onSelect: function() {
                    this.hide();
                }
            });
        </script>
        /
        ǥ�ù�� :
        <label><input type="radio" name="mode" value="OD" <%=chkIIF(mode="OD","checked","")%> /> �ֹ���ȣ ����</label>
        <label><input type="radio" name="mode" value="IT" <%=chkIIF(mode="IT","checked","")%> /> ��ǰ ����</label>
        /
        ������ :
        <select name="ordstat">
            <option value="" <%=chkIIF(ordstat="","selected","")%>>��ü</option>
            <option value="B" <%=chkIIF(ordstat="B","selected","")%>>���������</option>
            <option value="C" <%=chkIIF(ordstat="C","selected","")%>>�̹��</option>
        </select>
        /
        �������� :
        <select name="danjong">
            <option value="" <%=chkIIF(danjong="","selected","")%>>��ü</option>
            <option value="N" <%=chkIIF(danjong="N","selected","")%>>������</option>
            <option value="S" <%=chkIIF(danjong="S","selected","")%>>�Ͻ�ǰ��</option>
            <option value="M" <%=chkIIF(danjong="M","selected","")%>>MDǰ��</option>
            <option value="Y" <%=chkIIF(danjong="Y","selected","")%>>����</option>
        </select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitFrm();">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br />

* �ִ� 3õ�Ǳ����� �˻��˴ϴ�.

<br />

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<thead>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="<%=chkIIF(mode="OD","14","12")%>">
		�˻���� : <b><%= oOrder.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if mode="OD" then %>
    <th>�ֹ���ȣ</th>
    <th>������</th>
    <% end if %>
    <th>�귣��ID</th>
    <th>��ǰ�ڵ�</th>
    <th>�ɼ��ڵ�</th>
    <th>��ǰ��</th>
    <th>�ֹ�����</th>
    <th>���</th>
    <th>�ǸŰ�</th>
    <th>���԰�</th>
    <th>����ּ���</th>
    <th>��������</th>
    <th>���ڵ�</th>
    <th>������</th>
</tr>
</thead>
<tbody>
<%
    if oOrder.FResultCount>0 then
        for i=0 to oOrder.FResultCount - 1
%>
<tr bgcolor="#FFFFFF">
    <% if mode="OD" then %>
    <td class="txtct"><%= oOrder.FItemList(i).Forderserial %></td>
    <td class="txtct"><%= left(oOrder.FItemList(i).Fipkumdate,10) %></td>
    <% end if %>
    <td><%= oOrder.FItemList(i).Fmakerid %></td>
    <td class="txtct"><%= oOrder.FItemList(i).Fitemid %></td>
    <td class="txtct"><%= oOrder.FItemList(i).Fitemoption %></td>
    <td><%= oOrder.FItemList(i).Fitemname & chkIIF(oOrder.FItemList(i).Foptionname<>""," ("&oOrder.FItemList(i).Foptionname&")","") %></td>
    <td class="txtrt"><%= oOrder.FItemList(i).Ficnt %></td>
    <td class="txtrt"><%= oOrder.FItemList(i).Frealstock %></td>
    <td class="txtrt"><%= FormatNumber(oOrder.FItemList(i).Fsellcash,0) %></td>
    <td class="txtrt"><%= FormatNumber(oOrder.FItemList(i).Fbuycash,0) %></td>
    <td class="txtrt"><%= oOrder.FItemList(i).Fpreordernofix %></td>
    <td class="txtct"><%= oOrder.FItemList(i).FisDanjong %></td>
    <td class="txtct"><%= oOrder.FItemList(i).FrackcodeByOption & chkIIF(oOrder.FItemList(i).FsubRackcodeByOption<>"","<br />("&oOrder.FItemList(i).FsubRackcodeByOption&")","") %></td>
    <td class="txtct"><%= oOrder.FItemList(i).Fgubun %></td>
</tr>
<%
            '���� �÷���
            if ((i+1) mod 500)=0 then
                response.Flush()
            end if
        next
    else
%>
<tr bgcolor="#FFFFFF">
    <td colspan="<%=chkIIF(mode="OD","14","12")%>" style="text-align:center;">�˻������ �����ϴ�.</td>
</tr>
<%  end if %>
</tbody>
</table>
<%
    set oOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
