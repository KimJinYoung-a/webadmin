<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ä�����-����
' History : 2016.07.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_analisys.asp" -->
<%
dim i, yyyy, tmpyyyymm, mm_MaechulProfit, tot_MaechulProfit, mm_beforeitemcostsum
dim mm_itemcostsum, mm_buycashsum, mm_ordercnt, mm_itemnosum, tot_itemcnt
	yyyy = requestcheckvar(request("yyyy"),4)

if yyyy="" then yyyy = year(date)

dim cStatistic
Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectyyyy = yyyy
	cStatistic.fStatistic_monthly_channel()
%>

<script type='text/javascript'>

function searchSubmit(){
    frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="30">
				* ��¥ : <% DrawyearBoxdynamic "yyyy", yyyy, " onchange='searchSubmit();'" %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<p>

* �ֹ��Ǽ��� ������ ���� ��ǰ�� �ִ� �ֹ��� �Ǽ��Դϴ�.(2������ ���� ��ǰ�� ���Ǹ� ���� 1�Ǿ�.)

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>�Ǹſ�</td>
	<td>����ä��</td>
    <td>�����Ѿ�[��ǰ]<br>(��ǰ��������)</td>
    <td>�ֹ��Ǽ�</td>
    <td>���ܰ�<br>(�ֹ�)</td>
    <td>��ǰ����</td>
    <td>���ܰ�<br>(��ǰ)</td>
    <td>�������</td>
	<td>������ͷ�</td>
    <td>ä�κ���</td>
    <td>�������<br>�����</td>
</tr>
<%
if cStatistic.FTotalCount > 0 then

For i = 0 To cStatistic.FTotalCount -1

if tmpyyyymm <> cStatistic.flist(i).Fyyyymm and i <> 0 then
%>
	<tr align="center" bgcolor="#f1f1f1">
		<td colspan=2><%= tmpyyyymm %> �հ�</td>
		<td align="right">
			<% '/����� %>
			<%= FormatNumber(mm_itemcostsum,0) %>
		</td>
		<td align="right">
			<% '/�ֹ��Ǽ� %>
			<%= FormatNumber(mm_ordercnt,0) %>
		</td>
		<td align="right">
			<%
			'/���ܰ�(�ֹ�)
			if mm_itemcostsum<>0 and mm_ordercnt<>0 then
			%>
				<%= FormatNumber(mm_itemcostsum/mm_ordercnt,0) %>
			<% else %>
				0
			<% end if %>
		</td>
		<td align="right">
			<% '/��ǰ���� %>
			<%= FormatNumber(mm_itemnosum,0) %>
		</td>
		<td align="right">
			<%
			'/���ܰ�(��ǰ)
			if mm_itemcostsum<>0 and mm_itemnosum<>0 then
			%>
				<%= FormatNumber(mm_itemcostsum/mm_itemnosum,0) %>
			<% else %>
				0
			<% end if %>
		</td>
		<td align="right">
			<% '/������� %>
			<%= FormatNumber(mm_MaechulProfit,0) %>
		</td>
		<td>
			<%
			'/������ͷ�
			if mm_itemcostsum<>0 then
			%>
				<%= round( ((( mm_itemcostsum-mm_buycashsum ) / mm_itemcostsum )*100) ,2) %>%
			<% else %>
				<%= round( ((( mm_itemcostsum-mm_buycashsum ) / 1 )*100) ,2) %>%
			<% end if %>
		</td>
		<td></td>
		<td>
			<%
			'/������� �����
			if mm_itemcostsum<>0 and mm_beforeitemcostsum<>0 then
			%>
				<%= round( (( mm_itemcostsum/mm_beforeitemcostsum )*100) -100 ,2) %>%
			<% else %>
				0%
			<% end if %>
		</td>
	</tr>
<%
	mm_itemcostsum = 0
	mm_beforeitemcostsum = 0
	mm_buycashsum = 0
	mm_ordercnt = 0
	mm_itemnosum = 0
	mm_MaechulProfit = 0
end if

tmpyyyymm = cStatistic.flist(i).Fyyyymm
mm_itemcostsum = mm_itemcostsum + cStatistic.flist(i).Fitemcostsum
mm_beforeitemcostsum = mm_beforeitemcostsum + cStatistic.flist(i).Fbeforeitemcostsum
mm_buycashsum = mm_buycashsum + cStatistic.flist(i).fbuycashsum
mm_ordercnt = mm_ordercnt + cStatistic.flist(i).fordercnt
mm_itemnosum = mm_itemnosum + cStatistic.flist(i).fitemnosum
mm_MaechulProfit = mm_MaechulProfit + cStatistic.flist(i).fMaechulProfit
tot_MaechulProfit = tot_MaechulProfit + mm_MaechulProfit
%>

<tr bgcolor="#FFFFFF" align="center">
	<td>
		<%= cStatistic.flist(i).Fyyyymm %>
	</td>
	<td>
		<%= getchannelname(cStatistic.flist(i).Fchannel) %>
	</td>
	<td align="right">
		<% '/����� %>
		<%= FormatNumber(cStatistic.flist(i).Fitemcostsum,0) %>
	</td>
	<td align="right">
		<% '/�ֹ��Ǽ� %>
		<%= FormatNumber(cStatistic.flist(i).Fordercnt,0) %>
	</td>
	<td align="right">
		<% '/���ܰ�(�ֹ�) %>
		<%= FormatNumber(cStatistic.flist(i).forderunit,0) %>
	</td>
	<td align="right">
		<% '/��ǰ���� %>
		<%= FormatNumber(cStatistic.flist(i).Fitemnosum,0) %>
	</td>
	<td align="right">
		<% '/���ܰ�(��ǰ) %>
		<%= FormatNumber(cStatistic.flist(i).fitemunit,0) %>
	</td>
	<td align="right">
		<% '������� %>
		<%= FormatNumber(cStatistic.flist(i).fMaechulProfit,0) %>
	</td>
	<td>
		<% '������ͷ� %>
		<%= cStatistic.flist(i).FMaechulProfitPer %>%
	</td>
	<td>
		<%
		'ä�κ���
		if cStatistic.flist(i).fchannelitemcostsum<>0 and cStatistic.flist(i).Fitemcostsum<>0 then
		%>
			<%= round((cStatistic.flist(i).Fitemcostsum/cStatistic.flist(i).fchannelitemcostsum)*100,2) %>%
		<% else %>
			0%
		<% end if %>
	</td>
	<td>
		<% '/�������<br>����� %>
		<%= cStatistic.flist(i).fbeforemmper %>%
	</td>
</tr>
<%
Next
%>

<tr align="center" bgcolor="#f1f1f1">
	<td colspan=2><%= tmpyyyymm %> �հ�</td>
	<td align="right">
		<% '/����� %>
		<%= FormatNumber(mm_itemcostsum,0) %>
	</td>
	<td align="right">
		<% '/�ֹ��Ǽ� %>
		<%= FormatNumber(mm_ordercnt,0) %>
	</td>
	<td align="right">
		<%
		'/���ܰ�(�ֹ�)
		if mm_itemcostsum<>0 and mm_ordercnt<>0 then
		%>
			<%= FormatNumber(mm_itemcostsum/mm_ordercnt,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right">
		<% '/��ǰ���� %>
		<%= FormatNumber(mm_itemnosum,0) %>
	</td>
	<td align="right">
		<%
		'/���ܰ�(��ǰ)
		if mm_itemcostsum<>0 and mm_itemnosum<>0 then
		%>
			<%= FormatNumber(mm_itemcostsum/mm_itemnosum,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right">
		<% '/������� %>
		<%= FormatNumber(mm_MaechulProfit,0) %>
	</td>
	<td>
		<%
		'/������ͷ�
		if mm_itemcostsum<>0 then
		%>
			<%= round( ((( mm_itemcostsum-mm_buycashsum ) / mm_itemcostsum )*100) ,2) %>%
		<% else %>
			<%= round( ((( mm_itemcostsum-mm_buycashsum ) / 1 )*100) ,2) %>%
		<% end if %>
	</td>
	<td></td>
	<td>
		<%
		'/������� �����
		if mm_itemcostsum<>0 and mm_beforeitemcostsum<>0 then
		%>
			<%= round( (( mm_itemcostsum/mm_beforeitemcostsum )*100) -100 ,2) %>%
		<% else %>
			0%
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#f1f1f1">
	<td colspan=2>���հ�</td>
	<td align="right">
		<% '/����� %>
		<%= FormatNumber(cStatistic.totitemcostsum,0) %>
	</td>
	<td align="right">
		<% '/�ֹ��Ǽ� %>
		<%= FormatNumber(cStatistic.totordercnt,0) %>
	</td>
	<td align="right">
		<%
		'/���ܰ�(�ֹ�)
		if cStatistic.totitemcostsum<>0 and cStatistic.totordercnt<>0 then
		%>
			<%= FormatNumber(cStatistic.totitemcostsum/cStatistic.totordercnt,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right">
		<% '/��ǰ���� %>
		<%= FormatNumber(cStatistic.totitemnosum,0) %>
	</td>
	<td align="right">
		<%
		'/���ܰ�(��ǰ)
		if cStatistic.totitemcostsum<>0 and cStatistic.totitemnosum<>0 then
		%>
			<%= FormatNumber(cStatistic.totitemcostsum/cStatistic.totitemnosum,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right">
		<% '/������� %>
		<%= FormatNumber(tot_MaechulProfit,0) %>
	</td>
	<td>
		<%
		'/������ͷ�
		if mm_itemcostsum<>0 then
		%>
			<%= round( ((( cStatistic.totitemcostsum-cStatistic.totbuycashsum ) / cStatistic.totitemcostsum )*100) ,2) %>%
		<% else %>
			<%= round( ((( cStatistic.totitemcostsum-cStatistic.totbuycashsum ) / 1 )*100) ,2) %>%
		<% end if %>
	</td>
	<td></td>
	<td>
		<%
		'/������� �����
		if cStatistic.totitemcostsum<>0 and cStatistic.totbeforeitemcostsum<>0 then
		%>
			<%= round( (( cStatistic.totitemcostsum/cStatistic.totbeforeitemcostsum )*100) -100 ,2) %>%
		<% else %>
			0%
		<% end if %>
	</td>
</tr>

<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">������ �����ϴ�.</td>
	</tr>
<% end if %>

</table>

<%
set cStatistic = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
