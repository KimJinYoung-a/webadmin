<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� ����Ȯ���� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim i
dim research : research = requestCheckvar(request("research"),10)
dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)
dim page : page = requestCheckvar(request("page"),10)

dim difftp : difftp = requestCheckvar(request("difftp"),10)
dim chkerritemno : chkerritemno = requestCheckvar(request("chkerritemno"),10)

dim yyyy1, mm1
''dim fromDate, toDate, dlvyyyy, dlvmm
yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)
'dlvyyyy = requestCheckvar(request("dlvyyyy"),4)
'dlvmm = requestCheckvar(request("dlvmm"),2)

if (yyyy1="") then yyyy1=LEFT(NOW(),4)
if (mm1="") then mm1=MID(NOW(),6,2)
if (page="") then page=1

dim oCExtJungsanDiff
SET oCExtJungsanDiff = new CExtJungsan
	oCExtJungsanDiff.FPageSize = 100
	oCExtJungsanDiff.FCurrPage = page
	oCExtJungsanDiff.FRectSellSite = sellsite
	oCExtJungsanDiff.FRectDlvMonth = yyyy1+"-"+mm1
	oCExtJungsanDiff.getExtOrderJungsanFixdate

dim FormatDotNo : FormatDotNo=0
%>
<script language='javascript'>

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���޸�:
		<%= getJungsanXsiteComboHTML("sellsite",sellsite,"") %>
		&nbsp;
		
		* ����:
		<% DrawYMBox yyyy1,mm1 %>
        &nbsp;
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p  >
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oCExtJungsanDiff.FTotalcount %></b>
		&nbsp;
		<% if oCExtJungsanDiff.FTotalcount>=oCExtJungsanDiff.FPageSize then %>
        (�ִ� <%=FormatNumber(oCExtJungsanDiff.FPageSize,0)%> ��)
        <% end if %>
	</td>
</tr>

<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">���޸�</td>
	<td width="90">�ֹ���ȣ</td>
	<td width="140">���� �ֹ���ȣ</td>
    <td width="80">��ǰ�ڵ�</td>
    <td width="70">�ɼ��ڵ�</td>

	<td width="40">����(��)</td>
	<td width="70">�ǸŰ�(��)</td>
    <td width="70">���ǸŰ�(��)</td>

	<td width="70">����</td>
	<td width="70">��������</td>
	<td width="70">�����ݾ�</td>

	<td width="80">�ù��</td>
	<td width="90">�����ȣ</td>
    
	<td width="70">��ۿϷ���</td>
	<td width="70">����Ϸ���</td>
</tr>

<% if oCExtJungsanDiff.FresultCount<1 then %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <td colspan="20">
        <% if (sellsite="") then %>
        [���� ���޸��� ���� �ϼ���.]
        <% else %>
        [�˻������ �����ϴ�.]
        <% end if %>
    </td>
</tr>
<% else %>
<% for i=0 to oCExtJungsanDiff.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCExtJungsanDiff.FItemList(i).Fsitename %></td>
	<td><a href="#" onClick="popDeliveryTrackingSummaryOne(<%= oCExtJungsanDiff.FItemList(i).ForgOrderserial %>,'<%= oCExtJungsanDiff.FItemList(i).Forgsongjangno %>',<%= oCExtJungsanDiff.FItemList(i).Forgsongjangdiv %>);return false;"><%= oCExtJungsanDiff.FItemList(i).ForgOrderserial %></a></td>
    <td><a href="#" onClick="popByExtorderserial('<%= oCExtJungsanDiff.FItemList(i).Fauthcode %>');return false;"><%= oCExtJungsanDiff.FItemList(i).Fauthcode %></a></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Fitemid %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Fitemoption %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Fitemno+oCExtJungsanDiff.FItemList(i).FMinus_itemno %></td>
    <td align="right"><%= FormatNumber(oCExtJungsanDiff.FItemList(i).FitemcostSum+oCExtJungsanDiff.FItemList(i).FMinus_itemcostSum,0) %></td>
    <td align="right"><%= FormatNumber(oCExtJungsanDiff.FItemList(i).FreducedpriceSum+oCExtJungsanDiff.FItemList(i).FMinus_reducedpriceSum,0) %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).FbeasongMonth %></td>
	<td>
		<% if isNULL(oCExtJungsanDiff.FItemList(i).Fjorgorderserial) then %>

		<% else %>
			<% if (oCExtJungsanDiff.FItemList(i).Fdiffitemno<>0) then %>
			<strong><%= FormatNumber(oCExtJungsanDiff.FItemList(i).Fdiffitemno,0) %></strong>
			<% else %>
			<%= FormatNumber(oCExtJungsanDiff.FItemList(i).Fdiffitemno,0) %>
			<% end if %>
		<% end if %>
	</td>
	<td>
		<% if isNULL(oCExtJungsanDiff.FItemList(i).Fjorgorderserial) then %>
		
		<% else %>
		<%= FormatNumber(oCExtJungsanDiff.FItemList(i).FdiffSum,0) %>
		<% end if %>
	</td>
    <td><%=getSongjangDiv2Val(oCExtJungsanDiff.FItemList(i).Forgsongjangdiv,1) %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Forgsongjangno %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Forgdlvfinishdt %></td>
    <td><%= oCExtJungsanDiff.FItemList(i).Forgjungsanfixdate %></td>
</tr>
<% next %>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
    <% if (FALSE) then %>
		<% if oCExtJungsanDiff.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCExtJungsanDiff.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCExtJungsanDiff.StartScrollPage to oCExtJungsanDiff.FScrollCount + oCExtJungsanDiff.StartScrollPage - 1 %>
			<% if i>oCExtJungsanDiff.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCExtJungsanDiff.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    <% end if %>
	</td>
</tr>

</table>

<p>
<form name="frmcmt" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="addcmt">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="addcomment" value="">
</form>

<%
set oCExtJungsanDiff = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->