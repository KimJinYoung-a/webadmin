<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ����
' Hieditor : 2020/03/30 eastone
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/jungsan/jungsanCheckCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
Dim i
dim research : research = requestCheckvar(request("research"),10)
dim orderserial : orderserial = requestCheckvar(request("orderserial"),16)

dim oJungsanCheck
SET oJungsanCheck = new CJungsanCheck
oJungsanCheck.FRectOrderserial = orderserial
oJungsanCheck.getLogDiffByOrderserial


Dim pitemgubun, pitemid, pitemoption
pitemid = -1

%>
<script language='javascript'>
function chgCancelOrderJFixdtNULL(iorderserial){
    var frm = document.frmXsiteOrderVal;
    frm.mode.value="chgCancelOrderJFixdtNULL";
    frm.orderserial.value=iorderserial;

    if (confirm("�ֹ� ���� ��ۿϷ��� �������� NULL �� �����Ͻðڽ��ϱ�?")){
        frm.submit();
    }
}

function chgCancelOrderDetailRealsellprice(iorderserial,iitemid,iitemoption,orgrealsellprice){
    var chgrealsellprice = "";
    chgrealsellprice = prompt("�����ұݾ�", "");
    if (chgrealsellprice == null) return;

    if (chgrealsellprice.length<1){
        alert("���ǸŰ��� �Է��ϼ���.");
        return;
    }

    if (!IsDigit(chgrealsellprice)){
        alert('���ڸ� �Է��ϼ���.');
        return;
    }

    var frm = document.frmXsiteOrderVal;
    frm.mode.value="chgRealOrderRealsellprice";
    frm.orderserial.value=iorderserial;
    frm.itemid.value=iitemid;
    frm.itemoption.value=iitemoption;
    frm.chgval.value=chgrealsellprice;

    if (confirm("�ֹ����� ���ǸŰ� ���� "+orgrealsellprice+" => "+chgrealsellprice+" �� �����Ͻðڽ��ϱ�?")){
        frm.submit();
    }
}
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
	
		
		* �ֹ���ȣ : <input type="text" name="orderserial" value="<%=orderserial%>" size="11" maxlength="16">
        &nbsp;

	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF" >
	<td>

	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p  >
<% if oJungsanCheck.FresultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100">�ֹ���ȣ</td>
    <td width="120">����Ʈ</td>
    <td width="70">��ҿ���</td>
    <td width="70">CHK</td>
    <td>���</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
    <td><%= oJungsanCheck.FItemList(0).Forderserial %></td>
    <td><%= oJungsanCheck.FItemList(0).Fsitename %></td>
    <td><%= oJungsanCheck.FItemList(0).Fcancelyn %></td>
    <td><%= oJungsanCheck.FItemList(0).getLogCheckTypeName %></td>
    <td align="left">
    <% if  (NOT isNULL(oJungsanCheck.FItemList(0).Fchktype)) then %>
    <% if  oJungsanCheck.FItemList(0).Fchktype=8 then %>
        <input type="button" value="���/������ NULLó��" onClick="chgCancelOrderJFixdtNULL('<%= oJungsanCheck.FItemList(0).Forderserial %>'); return false;">
    <% end if %>
    <% end if %>
    </td>
</tr>
</table>
<% end if %>
<p>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="27">
		
		
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="30">��ǰ<br>����</td>
	<td width="60">��ǰ�ڵ�</td>
	<td width="60">�ɼ��ڵ�</td>
	<td width="90">�귣��ID</td>
    <td width="50">��ǰ<br>����</td>
	<td width="50">���<br>����</td>

    <td width="60">�ǸŰ�</td>
    <td width="60">���Ŵܰ�</td>
    <td width="60">����ܰ�</td>
    <td width="60">���԰�</td>
    <td width="60">����<br>����</td>
	<td width="60">�����</td>
    <td width="60">�����</td>
    <td width="60">������</td>
    <td width="40">����<br>����</td>
    <td width="10"></td>

    <td width="30">�α�<br>Sub</td>
    <td width="60">�α�<br>����</td>
    <td width="60">�α�<br>�ǸŰ�</td>
    <td width="60">�α�<br>���Ŵܰ�</td>
    <td width="60">�α�<br>����ܰ�</td>
    <td width="60">�α�<br>���԰�</td>
    <td width="60">�α�<br>���Ա���</td>
	<td width="60">�α�<br>�����</td>
    <td width="60">�α�<br>������</td>
    <td width="40">�α�<br>����</td>
	<td>���</td>

   
</tr>

<% if oJungsanCheck.FresultCount<1 then %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <td colspan="27">
       
        [�˻������ �����ϴ�.]
    </td>
</tr>
<% else %>
<% for i=0 to oJungsanCheck.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <% if (oJungsanCheck.FItemList(i).Fitemgubun=pitemgubun and oJungsanCheck.FItemList(i).Fitemid=pitemid and oJungsanCheck.FItemList(i).Fitemoption=pitemoption) then %>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
    <% else %>
        <td><%= oJungsanCheck.FItemList(i).Fitemgubun %></td>
        <td><%= oJungsanCheck.FItemList(i).Fitemid %></td>
        <td><%= oJungsanCheck.FItemList(i).Fitemoption %></td>
        <td><%= oJungsanCheck.FItemList(i).Fmakerid %></td>
        <td><%= oJungsanCheck.FItemList(i).Fitemno %></td>
        <td><%= oJungsanCheck.FItemList(i).getDCancelynName %></td>
        
        <td align="right"><%= FormatNumber(oJungsanCheck.FItemList(i).FitemcostCouponNotApplied, 0) %></td>
        <td align="right"><%= FormatNumber(oJungsanCheck.FItemList(i).Fitemcost, 0) %></td>
        <td align="right">
            <% if (oJungsanCheck.FItemList(i).Fdcancelyn="Y") and (oJungsanCheck.FItemList(i).Fsitename<>"10x10") and (LEN(oJungsanCheck.FItemList(i).Forderserial)=11) and (oJungsanCheck.FItemList(i).Fitemid<>0)  then %>
            <a href="#" onClick="chgCancelOrderDetailRealsellprice('<%= oJungsanCheck.FItemList(i).Forderserial %>','<%= oJungsanCheck.FItemList(i).Fitemid %>','<%= oJungsanCheck.FItemList(i).Fitemoption %>',<%= oJungsanCheck.FItemList(i).Freducedprice %>); return false;"><%= FormatNumber(oJungsanCheck.FItemList(i).Freducedprice, 0) %></a>
            <% else %>
            <%= FormatNumber(oJungsanCheck.FItemList(i).Freducedprice, 0) %>
            <% end if %>
        </td>
        <td align="right"><%= FormatNumber(oJungsanCheck.FItemList(i).Fbuycash, 0) %></td>
        <td><%= oJungsanCheck.FItemList(i).Fomwdiv %></td>
        <td><%= oJungsanCheck.FItemList(i).Fbeasongdate %></td>
        <td><%= oJungsanCheck.FItemList(i).Fdlvfinishdt %></td>
        <td><%= oJungsanCheck.FItemList(i).Fjungsanfixdate %></td>
        <td><%= oJungsanCheck.FItemList(i).Fvatinclude %></td>
    <% end if %>
    <%
    pitemgubun  = oJungsanCheck.FItemList(i).Fitemgubun
    pitemid     = oJungsanCheck.FItemList(i).Fitemid
    pitemoption = oJungsanCheck.FItemList(i).Fitemoption
    %>
    <td width="10"></td>
    
    <td><%= oJungsanCheck.FItemList(i).Fsuborderserial %></td>
    <td><%= oJungsanCheck.FItemList(i).Flgitemno %></td>
    <td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).FlgitemcostCouponNotApplied) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).FlgitemcostCouponNotApplied, 0) %>
        <% end if %>
    </td>
    <td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).Flgitemcost) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).Flgitemcost, 0) %>
        <% end if %>
    </td>
    <td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).FlgreducedPrice) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).FlgreducedPrice, 0) %>
        <% end if %>
    </td>
    <td align="right">
        <% if NOT isNULL(oJungsanCheck.FItemList(i).Flgbuycash) then %>
        <%= FormatNumber(oJungsanCheck.FItemList(i).Flgbuycash, 0) %>
        <% end if %>
    </td>
    <td><%= oJungsanCheck.FItemList(i).Flgomwdiv %></td>
	<td><%= oJungsanCheck.FItemList(i).Flgbeasongdate %></td>
    <td><%= oJungsanCheck.FItemList(i).FDTLjFixedDt %></td>
    <td><%= oJungsanCheck.FItemList(i).Flgvatinclude %></td>
    <td></td>
</tr>
<% next %>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="27" align="center">
		
	</td>
</tr>
</table>

<%
set oJungsanCheck = Nothing
%>
<form name="frmXsiteOrderVal" method="post" action="/admin/maechul/extjungsandata/extJungsan_process.asp">
<input type="hidden" name="mode" value="chgRealOrderRealsellprice">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="chgval" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
