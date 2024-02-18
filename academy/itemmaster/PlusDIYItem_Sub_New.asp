<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/PlusDIYItemCls.asp"-->
<%
dim itemid
itemid = requestCheckvar(request("itemid"),9)
itemid = CStr(itemid)
'itemid = Cint(itemid)

dim oitem
set oitem = new CItem
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItem
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if


dim oPlusSaleItem
set oPlusSaleItem = new CPlusSaleItem
oPlusSaleItem.FRectItemID = itemid

if itemid<>"" then
	oPlusSaleItem.GetOnePlusSaleSubItem
end if

dim i
dim IsNewReg        '' �űԵ������
IsNewReg = (oPlusSaleItem.FResultCount<1)

'' ���� IsLinkedItem �ΰ��
dim IsLinkedItem
if itemid<>"" then
    IsLinkedItem = oPlusSaleItem.IsPlusSaleLinkItem
end if
%>

<script language='javascript'>
function CalcuMargin(frm){
    var vSalePro = frm.PlusSalePro.value;
    var vMarginFlag = frm.PlusSaleMaginFlag.value;
    var vOrgMargin = 0;
    var vSaleMargin = 0;

    if (vSalePro.length<1){
        alert('�������� �Է��ϼ���.');
        frm.PlusSalePro.focus();
    }

    if (!IsDigit(vSalePro)){
        alert('�������� ���ڷ� �Է��ϼ���.');

        frm.PlusSalePro.focus();
        frm.PlusSalePro.select();
    }

    frm.tmpSellCash.value = parseInt(frm.osellcash.value-frm.osellcash.value*vSalePro/100);
    vOrgMargin = 100-parseInt(frm.obuycash.value*1/frm.osellcash.value*1*100*100)/100;

    frm.PlusSaleMargin.readOnly = true;
    frm.PlusSaleMargin.className = "text_ro";

    if (vMarginFlag=="1"){      //���ϸ���
        vSaleMargin = vOrgMargin;
        frm.PlusSaleMargin.value = vSaleMargin;
        frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-frm.tmpSellCash.value*vSaleMargin/100);

    }else if(vMarginFlag=="2"){  //��ü�δ� : ������
        frm.tmpBuyCash.value = frm.tmpSellCash.value*1-parseInt(frm.osellcash.value*1-frm.obuycash.value*1);
        vSaleMargin = 100-parseInt(frm.tmpBuyCash.value*1/frm.tmpSellCash.value*1*100*100)/100;
        frm.PlusSaleMargin.value = vSaleMargin;
        frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-(frm.osellcash.value*1-frm.obuycash.value*1));

    }else if(vMarginFlag=="3"){  //�ݹݺδ� : ������
        frm.tmpBuyCash.value = frm.obuycash.value*1-parseInt((frm.osellcash.value*1-frm.tmpSellCash.value*1)/2);
        vSaleMargin = 100-parseInt(frm.tmpBuyCash.value*1/frm.tmpSellCash.value*1*100*100)/100;
        frm.PlusSaleMargin.value = vSaleMargin;
        frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-frm.tmpSellCash.value*vSaleMargin/100);

    }else if(vMarginFlag=="4"){  //�ٹ����ٺδ� : ���� ��� ���� 0 ����.(+-1�� ���� ������?.)
        frm.tmpBuyCash.value = frm.obuycash.value*1;
        vSaleMargin = 100-parseInt(frm.tmpBuyCash.value*1/frm.tmpSellCash.value*1*100*100)/100;
        frm.PlusSaleMargin.value = vSaleMargin;
        //frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-frm.tmpSellCash.value*vSaleMargin/100);
        frm.tmpBuyCash.value = frm.obuycash.value;

    }else if(vMarginFlag=="5"){  //��������
        frm.PlusSaleMargin.readOnly = false;
        frm.PlusSaleMargin.className = "text";
        frm.PlusSaleMargin.focus();

        //vSaleMargin = vOrgMargin;
        //frm.PlusSaleMargin.value = vSaleMargin;
        vSaleMargin = frm.PlusSaleMargin.value;
        frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-frm.tmpSellCash.value*vSaleMargin/100);


    }



}

function setComp(comp){
    if (comp.name=="termsGubun"){
        if (comp.value=="A"){
            comp.form.PlusSaleStartDate.value = "1901-01-01";
            comp.form.PlusSaleEndDate.value = "9999-12-31";
        }else if (comp.value=="S"){
            comp.form.PlusSaleStartDate.value = "";
            comp.form.PlusSaleEndDate.value = "";
        }
    }
}

function RegPLusSale(frm){
    if (frm.PlusSalePro.value.length<1){
        alert('�������� �Է��ϼ���.');
        frm.PlusSalePro.focus();
        return;
    }

    if (!IsDigit(frm.PlusSalePro.value)){
        alert('�������� �Է��ϼ���.');
        frm.PlusSalePro.focus();
        return;
    }

    if (frm.PlusSalePro.value*1>50){
        alert('�������� Ȯ���� �ּ���.');
        frm.PlusSalePro.focus();
        return;
    }

    if ((frm.omwdiv.value=="M")&&(frm.PlusSaleMaginFlag.value!="4")){
        alert('��ǰ ���Ա����� ������ ��� ���� ���� ������ �ٹ����� �δ����� �����ϼ���.');
        frm.PlusSaleMaginFlag.focus();
        return;
    }

    if ((frm.PlusSaleMargin.value*1>100)||(frm.PlusSaleMargin.value*1<1)){
        alert('���ν� �������� Ȯ���� �ּ���.');
        //frm.PlusSaleMargin.focus();
        return;
    }

    if  (!IsDouble(frm.PlusSaleMargin.value)){
        alert('���ν� �������� Ȯ���� �ּ���.');
        //frm.PlusSaleMargin.focus();
        return;
    }

    if (frm.tmpBuyCash.value*1<0){
        alert('���ν� ���԰��� Ȯ���� �ּ���.');
        return;
    }

    if (frm.tmpSellCash.value*1<frm.tmpBuyCash.value*1){
        alert('���ν� ���԰��� Ȯ���� �ּ���. ���ν� �ǸŰ����� Ŭ �� �����ϴ�.');
        return;
    }

    //PlusSaleMaginFlag

    if ((!frm.termsGubun[0].checked)&&(!frm.termsGubun[1].checked)){
        alert('�Ⱓ ���� ���θ� ������ �ּ���.');
        frm.termsGubun[1].focus();
        return;
    }


    if (frm.PlusSaleStartDate.value.length<1){
        alert('�������� ���� �ϼ���.');
        return;
    }

    if (frm.PlusSaleEndDate.value.length<1){
        alert('�������� ���� �ϼ���.');
        return;
    }

    if (frm.PlusSaleStartDate.value>frm.PlusSaleEndDate.value){
        alert('�������� ������ �������� ���� �� �� �����ϴ�.');
        return;
    }

    <% if IsNewReg then %>
    if (confirm('��� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    <% else %>
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    <% end if %>
}

function showLinkedItemList(iitemid){
    var popwin = window.open('PlusDIYItem_Edit.asp?itemid=' + iitemid,'PlusDIYItem_Edit','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}


function DelPLusSale(frm){
    if (confirm('Plus Sale  �߰����� ��ǰ �� ���� �Ͻðڽ��ϱ�? - ���λ�ǰ ��ũ�� ���� �����˴ϴ�.')){
        frm.mode.value = "delPlusSale";
        frm.submit();
    }
}
</script>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method="get" >
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<img src="/images/icon_star.gif" border="0" align="absbottom">
			<b>PlusSale �߰����� ��ǰ ���</b>
		</td>
	</tr>
	<% if (oitem.FResultCount<1) then %>
	<tr height="25" bgcolor="FFFFFF">
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
		<td>
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">
			<input type="button" class="button" value="�˻�" onClick="document.frm.submit();">
		</td>
	</tr>
	<tr bgcolor="FFFFFF">
	    <td colspan="3" align="center">[�˻� ����� �����ϴ�.]</td>
	</tr>
	<% else %>
	<tr height="25" bgcolor="FFFFFF">
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
		<td>
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">
			<input type="button" class="button" value="�˻�" onClick="document.frm.submit();">
		</td>
		<td rowspan="4" width="100" align="right">
		    <img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100">
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
		<td><%= oitem.FOneItem.FItemName %></td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
		<td><%= oitem.FOneItem.FMakerid %></td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�Һ��ڰ�/���԰�</td>
		<td>
		    <% if (oitem.FOneItem.FsaleYn="Y") then %>
    			<%= FormatNumber(oitem.FOneItem.FOrgPrice,0) %> / <%= FormatNumber(oitem.FOneItem.Forgsuplycash,0) %>
    			&nbsp;
    			<%= fnPercent(oitem.FOneItem.Forgsuplycash,oitem.FOneItem.FOrgPrice,1) %>
    			&nbsp;&nbsp;
    			<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>

    			<br>

    			<font color=#F08050>(��)<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %></font>
    			&nbsp;
    			<%= fnPercent(oitem.FOneItem.FBuycash,oitem.FOneItem.FSellcash,1) %>
    			&nbsp;&nbsp;
    			<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>

    			<% if (oitem.FOneItem.IsCouponItem) then %>
    			<br><font color=#10F050>(��) <%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %></font>
    			<% end if %>
			<% else %>
    			<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
    			&nbsp;
    			<%= fnPercent(oitem.FOneItem.FBuycash,oitem.FOneItem.FSellcash,1) %>
    			&nbsp;&nbsp;
    			<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>

    			<% if (oitem.FOneItem.IsCouponItem) then %>
    			<br><font color=#10F050>(��) <%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %> <!-- / <%= FormatNumber(oitem.FOneItem.Fcouponbuyprice) %> --> &nbsp;<%= oitem.FOneItem.GetCouponDiscountStr %> ���� </font>
    			<% end if %>
			<% end if %>
		</td>
	</tr>
	<% end if %>
	</form>
</table>

<p>

<% if (oitem.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmPlusSale" method="post" action="PlusDIYItem_Process.asp">
    <input type="hidden" name="osellcash" value="<%= oitem.FOneItem.FSellcash %>">
    <input type="hidden" name="obuycash" value="<%= oitem.FOneItem.FBuycash %>">
    <input type="hidden" name="omwdiv" value="<%= oitem.FOneItem.FMwDiv %>">
    <input type="hidden" name="itemid" value="<%= itemid %>">
    <% if (IsNewReg) then %>
    <input type="hidden" name="mode" value="regPlusSale">
    <% else %>
    <input type="hidden" name="mode" value="editPlusSale">
    <% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�÷���������</td>
		<td>
		    <% if IsNewReg then %>
		    ������ :<input type="text" name="PlusSalePro" value="" class="text" size="5" maxlength="3" onKeyUp="CalcuMargin(frmPlusSale)">%
		    <% else %>
		    ������ :<input type="text" name="PlusSalePro" value="<%= oPlusSaleItem.FOneItem.FPlusSalePro %>" class="text" size="5" maxlength="3" onKeyUp="CalcuMargin(frmPlusSale)">%
		    <% end if %>
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">���νð�����</td>
	    <td>
	        <% if IsNewReg then %>
    			<!-- ���ް����� : -->
    			<select class="select" name="PlusSaleMaginFlag" onChange="CalcuMargin(frmPlusSale)">
    			    <option value="1" >���ϸ���</option>
                	<option value="2" >��ü�δ�</option>
                	<!-- <option value="3" >�ݹݺδ�</option> -->
                	<option value="4" >�ٹ����ٺδ�</option>
                	<option value="5" >��������</option>
    			</select>

    			<input type="text" name="PlusSaleMargin" class="text_ro" size="4" maxlength="4" onKeyUp="CalcuMargin(frmPlusSale)">%
    			&nbsp;&nbsp;
    			<input type="text" name="tmpSellCash" class="text_ro" size="10" maxlength="10" ReadOnly > / <input type="text" name="tmpBuyCash" value="" class="text_ro" size="10" maxlength="10" ReadOnly >
			<% else %>
			    <table border="0" cellspacing="0" cellpadding="0" class="a" >
			    <tr>
			        <td width="100" >
			            <select class="select" name="PlusSaleMaginFlag"  onChange="CalcuMargin(frmPlusSale)">
            			    <option value="1" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="1","selected","") %> >���ϸ���</option>
                        	<option value="2" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="2","selected","") %> >��ü�δ�</option>
                        	<!-- <option value="3" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="3","selected","") %> >�ݹݺδ�</option> -->
                        	<option value="4" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="4","selected","") %> >�ٹ����ٺδ�</option>
                        	<option value="5" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="5","selected","") %> >��������</option>
            			</select>
			        </td>
			        <td width="70">
			            <input type="text" name="PlusSaleMargin" class="text_ro" value="<%= oPlusSaleItem.FOneItem.FPlusSaleMargin %>" size="4" maxlength="4" onKeyUp="CalcuMargin(frmPlusSale)">%
			        </td>
			        <td>
			            <input type="text" name="tmpSellCash" value="<%= oPlusSaleItem.FOneItem.getPlusSalePrice %>" class="text_ro" size="10" maxlength="10"> / <input type="text" name="tmpBuyCash" value="<%= oPlusSaleItem.FOneItem.getPlusSaleBuycash %>" class="text_ro" size="10" maxlength="10">
			        </td>
			    </tr>
			    <tr>
			        <td><%= oPlusSaleItem.FOneItem.getMaginFlagName %></td>
			        <td><%= oPlusSaleItem.FOneItem.FPlusSaleMargin %>%</td>
			        <td>
			            <%= oPlusSaleItem.FOneItem.getPlusSalePrice %>
			            /
			            <%= oPlusSaleItem.FOneItem.getPlusSaleBuycash %>
			        </td>
			    </tr>
			    </table>
			<% end if %>
	    </td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�Ⱓ���࿩��</td>
		<td>
		    <% if IsNewReg then %>
			<input type="radio" name="termsGubun" value="A" onClick="setComp(this);">�������
			<input type="radio" name="termsGubun" value="S" checked onClick="setComp(this);">�Ⱓ����
			<% else %>
			<input type="radio" name="termsGubun" value="A" <%= ChkIIF(oPlusSaleItem.FOneItem.IsAlwaysTerms,"checked","") %> onClick="setComp(this);">�������
			<input type="radio" name="termsGubun" value="S" <%= ChkIIF(oPlusSaleItem.FOneItem.IsAlwaysTerms,"","checked") %> onClick="setComp(this);">�Ⱓ����
			<% end if %>
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td>
		    <% if IsNewReg then %>
		    <input type="text" class="text" name="PlusSaleStartDate" value="" size="11" maxlength="10"> <a href="javascript:calendarOpen(frmPlusSale.PlusSaleStartDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (�Ⱓ������ ���, ��������)
		    <% else %>
		    <input type="text" class="text" name="PlusSaleStartDate" value="<%= ChkIIF(Not IsNewReg,Left(oPlusSaleItem.FOneItem.FPlusSaleStartDate,10),"") %>" size="11" maxlength="10"> <a href="javascript:calendarOpen(frmPlusSale.PlusSaleStartDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (�Ⱓ������ ���, ��������)
		    <% end if %>
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td>
		    <% if IsNewReg then %>
		    <input type="text" class="text" name="PlusSaleEndDate" value="" size="11" maxlength="10"> <a href="javascript:calendarOpen(frmPlusSale.PlusSaleEndDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (�Ⱓ������ ���, ��������)
		    <% else %>
		    <input type="text" class="text" name="PlusSaleEndDate" value="<%= ChkIIF(Not IsNewReg,Left(oPlusSaleItem.FOneItem.FPlusSaleEndDate,10),"") %>" size="11" maxlength="10"> <a href="javascript:calendarOpen(frmPlusSale.PlusSaleEndDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (�Ⱓ������ ���, ��������)
		    <% end if %>
		</td>
	</tr>



	<!-- DB�� �ִ� ��ǰ�� ���̴� �޴� -->
	<% if (IsNewReg) then %>

	<% else %>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<td>
			 <!-- ���࿹�� / ������ / �Ⱓ���� (�Ⱓ���࿩�� �� �Ⱓ���� �Ǵ�) -->
			 <%= oPlusSaleItem.FOneItem.getCurrstateName %>
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">��ϵ� ���λ�ǰ</td>
		<td>
			<%= oPlusSaleItem.FOneItem.FLinkedItemCount %> ��
			<input type="button" class="button" value="��ũ��ǰ����Ʈ" onclick="showLinkedItemList('<%= itemid %>');">
		</td>
	</tr>
	<% end if %>
	<!-- DB�� �ִ� ��ǰ�� ���̴� �޴� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2" align="center">
		    <% if (IsNewReg) then %>
			<input type="button" class="button" value="�űԵ��" <%= ChkIIF(IsLinkedItem,"disabled","") %> onClick="RegPLusSale(frmPlusSale)";>
			<% else %>
			<input type="button" class="button" value=" ��  �� " onClick="RegPLusSale(frmPlusSale)";>
			&nbsp;&nbsp;&nbsp;
			<input type="button" class="button" value=" ��  �� " onClick="DelPLusSale(frmPlusSale)";>
			<% end if %>
		</td>
	</tr>
	</form>
</table>
<% end if %>
<!--
<p>

��ǰ�ڵ� �˻��� �÷������ϻ�ǰ DB�� �������, �űԵ��<br>
DB�� �������, ������ư ǥ��<br>
��ǰ�˻��� �� ��ǰ�ڵ尡 ���λ�ǰ���� ���ǰ� �������, �Ʒ��� ���� ��� ��ϺҰ����� ǥ��<br>
<br>
  ���ϸ���: �ǸŰ� ��� ���� ������ ����<br>
  ��ü�δ�: ���ǸŰ��� �����ݾ׸�ŭ �����ǸŰ����� ���� <br>
  �ݹݺδ�: ���αݾ��� 1/2�ݾ��� �����ް����� ����<br>
  �ٹ����ٺδ�: �����ް��� �����ǸŰ��ް��� ���� <br>
-->
<script language='javascript'>
function getOnLoad(){
    <% if (oitem.FResultCount>0) then %>
    <% if (oitem.FOneItem.FsaleYn="Y") then %>
    alert('�̹� �������� ��ǰ�Դϴ�.');
    <% end if %>

    <% if (IsLinkedItem) then %>
    alert('�̹� ���� ��ũ�� ��ϵ� ��ǰ�Դϴ�. - �÷��� ���� ��ǰ���� ��� �Ұ�.');
    <% end if %>
    <% end if %>
}
window.onload = getOnLoad;
</script>
<%
set oitem = Nothing
set oitemoption = Nothing
set oPlusSaleItem = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
