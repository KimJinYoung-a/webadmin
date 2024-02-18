<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<%

dim itemid, oitem
dim makerid

itemid = requestCheckvar(request("itemid"),10)
makerid = requestCheckvar(request("makerid"),32)
menupos = RequestCheckvar(request("menupos"),10)
if (itemid = "") then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbACADEMYget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

'==============================================================================
''��ü �⺻��� ����
dim defaultmargin, defaultmaeipdiv, defaultFreeBeasongLimit, defaultDeliverPay, defaultDeliveryType
Dim npartner, i
set npartner = new CPartnerUser
npartner.FRectDesignerID = oitem.FOneItem.Fmakerid
npartner.GetAcademyPartnerList
	defaultmargin			= npartner.FPartnerList(0).Fdiy_margin
    defaultmaeipdiv         = npartner.FPartnerList(0).Fmaeipdiv
    defaultFreeBeasongLimit = npartner.FPartnerList(0).FdefaultFreeBeasongLimit
    defaultDeliverPay       = npartner.FPartnerList(0).FdefaultDeliverPay
    defaultDeliveryType     = npartner.FPartnerList(0).FdefaultDeliveryType
set npartner = Nothing

'==============================================================================
'���ϸ���
dim sailmargine, orgmargine, margine

''����
if oitem.FOneItem.Fsailprice<>0 then
	sailmargine = 100-CLng(oitem.FOneItem.Fsailsuplycash/oitem.FOneItem.Fsailprice*100*100)/100
else
	sailmargine = 0
end if

if oitem.FOneItem.Forgprice<>0 then
	orgmargine = 100-CLng(oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice*100*100)/100
else
	orgmargine = 0
end if

if oitem.FOneItem.Fsellcash<>0 then
	margine = 100-CLng(oitem.FOneItem.Fbuycash/oitem.FOneItem.Fsellcash*100*100)/100
else
	margine = 0
end if

'==============================================================================
Sub SelectBoxDesignerItem(selectedId)
   dim query1,tmp_str
   %><select name="designer" onchange="TnDesignerNMargineAppl(this.value);">
     <option value='' <%if selectedId="" then response.write " selected"%>>-- ��ü���� --</option><%
   query1 = " select userid,socname_kor,defaultmargine from [db_user].[dbo].tbl_user_c order by userid"
'   query1 = query1 + " where isusing='Y' order by userid desc"
   rsACADEMYget.Open query1,dbACADEMYget,1

   if  not rsACADEMYget.EOF  then
       rsACADEMYget.Movefirst

       do until rsACADEMYget.EOF
           if Lcase(selectedId) = Lcase(rsACADEMYget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsACADEMYget("userid")& "," & rsACADEMYget("defaultmargine") & "' "&tmp_str&">" & rsACADEMYget("userid") & "  [" & replace(db2html(rsACADEMYget("socname_kor")),"'","") & "]" & "</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   response.write("</select>")
End Sub


%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language="javascript">

function UseTemplate() {
	window.open("/academy/comm/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

// ============================================================================
// ��ü�����ڵ��Է�
function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.margin.value = varArray[1];

}

function CalcuAuto(frm){
	var isvatYn, imileage;
	var isellcash, ibuycash, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatYn = frm.vatYn[0].checked;

	if (frm.saleYn[0].checked == true) {
	    // ���󰡰�
	    isellcash = frm.sellcash.value;
	    imargin = frm.margin.value;

    	if (imargin.length<1){
    		alert('������ �Է��ϼ���.');
    		frm.margin.focus();
    		return;
    	}

    	if (isellcash.length<1){
    		alert('�ǸŰ��� �Է��ϼ���.');
    		frm.sellcash.focus();
    		return;
    	}

    	if (!IsDouble(imargin)){
    		alert('������ ���ڷ� �Է��ϼ���.');
    		frm.margin.focus();
    		return;
    	}

    	if (!IsDigit(isellcash)){
    		alert('�ǸŰ��� ���ڷ� �Է��ϼ���.');
    		frm.sellcash.focus();
    		return;
    	}

    	if (isvatYn==true){
    		ibuycash = isellcash - parseInt(isellcash*imargin/100);
    		imileage = parseInt(isellcash*0.01) ;
    	}else{
    		ibuycash = isellcash - parseInt(isellcash*imargin/100);
    		imileage = parseInt(isellcash*0.01) ;
    	}

    	frm.buycash.value = ibuycash;
    	frm.mileage.value = imileage;
	} else {
	    // ���ϰ���
	    isailprice = frm.sailprice.value;
	    isailmargin = frm.sailmargin.value;

    	if (isailmargin.length<1){
    		alert('���ϸ����� �Է��ϼ���.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (isailprice.length<1){
    		alert('�����ǸŰ��� �Է��ϼ���.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (!IsDouble(isailmargin)){
    		alert('���ϸ����� ���ڷ� �Է��ϼ���.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (!IsDigit(isailprice)){
    		alert('�����ǸŰ��� ���ڷ� �Է��ϼ���.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (isvatYn==true){
    		isailpricevat = parseInt(parseInt(1/11 * parseInt(isailprice)));
    		isailsuplycash = isailprice - parseInt(isailprice*isailmargin/100);
    		isailsuplycashvat = parseInt(parseInt(1/11 * parseInt(isailsuplycash)));
    		imileage = parseInt(isailprice*0.01) ;
    	}else{
    		isailpricevat = 0;
    		isailsuplycash = isailprice - parseInt(isailprice*isailmargin/100);
    		isailsuplycashvat = 0;
    		imileage = parseInt(isailprice*0.01) ;
    	}

    	frm.sailpricevat.value = isailpricevat;
    	frm.sailsuplycash.value = isailsuplycash;
    	frm.sailsuplycashvat.value = isailsuplycashvat;
    	frm.mileage.value = imileage;
    }

	//������ ���
	if (frm.saleYn[0].checked == true) {
		document.getElementById("lyrPct").innerHTML = "";
	} else {
		isellcash = frm.sellcash.value;
		isailprice = frm.sailprice.value;
		var isalePercent = parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10;
		document.getElementById("lyrPct").innerHTML = "������: <font color='#EE0000'><strong>" + isalePercent + "%</strong></font>";
	}
}

// ============================================================================
// �����ϱ�
function SubmitSave() {
	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

    if (itemreg.saleYn[0].checked == true) {
        // ���󰡰�
        if (parseInt((itemreg.sellcash.value*1) * (itemreg.margin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {
    		alert("���ް��� �߸��ԷµǾ����ϴ�.[�Һ��ڰ�*���� = ���ް�]");
    		itemreg.sellcash.focus();

    		if (!confirm('�������� ��� �� �� ������ ���ް��� �Է��ϸ� �������� ���ް��� ���� ���˴ϴ�. \n��� ���� �Ͻðڽ��ϱ�?')){
				return;
			}
        }

        if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
            alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
            itemreg.mileage.focus();
            return;
        }

        if (itemreg.sellcash.value*1 < 300 || itemreg.sellcash.value*1 >= 20000000){
			alert("�Ǹ� ������ 300�� �̻� 20,000,000���� �̸����� ��� �����մϴ�.");
			itemreg.sellcash.focus();
			return;
		}

    } else {
        // ���ΰ���
        if (parseInt((itemreg.sailprice.value*1) * (itemreg.sailmargin.value*1) / 100) != ((itemreg.sailprice.value*1) - (itemreg.sailsuplycash.value*1))) {
    		alert("���ް��� �߸��ԷµǾ����ϴ�.[���μҺ��ڰ�*���θ��� = ���ΰ��ް�]");
    		itemreg.sailprice.focus();

    		if (!confirm('��� ���� �Ͻðڽ��ϱ�?')){
				return;
			}
        }

        if (itemreg.mileage.value*1 > itemreg.sailprice.value*1){
            alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
            itemreg.mileage.focus();
            return;
        }

        if (itemreg.sailprice.value*1 < 300 || itemreg.sailprice.value*1 >= 20000000){
			alert("�Ǹ� ������ 300�� �̻� 20,000,000���� �̸����� ��� �����մϴ�.");
			itemreg.sailprice.focus();
			return;
		}
    }


    //���ϰ����� ���󰡰� ���� Ŭ �� ����.
    if (itemreg.sailprice.value*1>itemreg.sellcash.value*1){
        alert('���ϰ����� ���󰡺��� Ŭ �� �����ϴ�.');
        return;
    }

    if (itemreg.sailsuplycash.value*1>itemreg.buycash.value*1){
        alert('���ϸ��԰��� ���� ���԰����� Ŭ �� �����ϴ�.');
        return;
    }

    //��۱��� üũ =======================================
    //��ü ���ǹ��
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[0].checked){
            alert('��� ������ Ȯ�����ּ���. ������� ��ü�� �ƴմϴ�.');
            return;
        }
    }

    //��ü���ҹ�� : ���ǹ�۵� ���Ҽ�������
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[2].checked)){
        alert('��� ������ Ȯ�����ּ���. [��ü ���ҹ��,��ü ���ǹ��] ��ü�� �ƴմϴ�.');
        return;
    }

    //==================================================================================

    if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
        itemreg.deliverytype[0].disabled=false;
		itemreg.deliverytype[1].disabled=false;
		itemreg.deliverytype[2].disabled=false;
        itemreg.submit();
    }

}


function TnChecksaleYn(frm){
	CheckSailEnDisabled(frm);
    CalcuAuto(frm);
}

function CheckSailEnDisabled(frm){
	if (frm.saleYn[0].checked == true) {
	    // ���󰡰�
        frm.sellcash.readonly = false;
        frm.margin.readonly = false;

        frm.sellcash.style.background = '#FFFFFF';
        frm.buycash.style.background = '#FFFFFF';
        frm.margin.style.background = '#FFFFFF';

        frm.sailprice.readonly = true;
        frm.sailmargin.readonly = true;

        frm.sailprice.style.background = '#E6E6E6';
        frm.sailsuplycash.style.background = '#E6E6E6';
        frm.sailmargin.style.background = '#E6E6E6';
	} else {
	    // ���ϰ���
        frm.sellcash.readonly = true;
        frm.margin.readonly = true;

        frm.sellcash.style.background = '#E6E6E6';
        frm.buycash.style.background = '#E6E6E6';
        frm.margin.style.background = '#E6E6E6';

        frm.sailprice.readonly = false;
        frm.sailmargin.readonly = false;

        frm.sailprice.style.background = '#FFFFFF';
        frm.sailsuplycash.style.background = '#FFFFFF';
        frm.sailmargin.style.background = '#FFFFFF';
    }
}

function ClearVal(comp){
    comp.value = "";
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>��ǰ ����/�Ǹ� ���� ����</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>��ϵ� ��ǰ�� ���� �� �Ǹ� ������ �����մϴ�.</b>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>

</table>
<!-- ǥ ��ܹ� ��-->


<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>�⺻����
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <form name="itemreg" method="post" action="itemmodify_Process.asp" onsubmit="return false;">
  <input type="hidden" name="mode" value="ItemPriceInfo">
  <input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
  <input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">

  <!-- ��ü �⺻ ��� ���� -->
  <input type="hidden" name="defaultmargin" value="<%= defaultmargin %>">
  <input type="hidden" name="defaultmaeipdiv" value="<%= defaultmaeipdiv %>">
  <input type="hidden" name="defaultFreeBeasongLimit" value="<%= defaultFreeBeasongLimit %>">
  <input type="hidden" name="defaultDeliverPay" value="<%= defaultDeliverPay %>">
  <input type="hidden" name="defaultDeliveryType" value="<%= defaultDeliveryType %>">

  <input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
  <input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
  <input type="hidden" name="availPayType" value="<%= oitem.FOneItem.FavailPayType %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�ڵ� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <%= oitem.FOneItem.Fitemid %>
  	  &nbsp;&nbsp;&nbsp;&nbsp;
  	  <input type="button" value="�̸�����" onclick="window.open('<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>');">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��üID :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><%=oitem.FOneItem.FMakerid %>&nbsp;&nbsp;(���� : <%= defaultmargin %>%)</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><%= oitem.FOneItem.Fitemname %></td>
  </tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>��������
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">���ݼ��� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
        <tr align="center">
          <td height="25" width="90" bgcolor="#DDDDFF">����</td>
          <td width="100" bgcolor="#DDDDFF">�Һ��ڰ�</td>
          <td width="100" bgcolor="#DDDDFF">���ް�</td>
          <td width="100" bgcolor="#DDDDFF">����</td>
          <td bgcolor="#DDDDFF">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" bgcolor="#FFFFFF"><input type="radio" name="saleYn" onClick="TnChecksaleYn(itemreg)" value="N" <% if oitem.FOneItem.FsaleYn = "N" then response.write "checked" %>> ���󰡰�</td>
          <td bgcolor="#FFFFFF" align="center">
            <% if oitem.FOneItem.FsaleYn = "N" then %>
            <input type="text" name="sellcash" maxlength="16" size="8" id="[on,on,off,off][�Һ��ڰ�]" value="<%= oitem.FOneItem.Fsellcash %>" onkeyup="CalcuAuto(itemreg);">��
            <% else %>
            <input type="text" name="sellcash" maxlength="16" size="8" id="[on,on,off,off][�Һ��ڰ�]" value="<%= oitem.FOneItem.Forgprice %>" onkeyup="CalcuAuto(itemreg);">��
            <% end if %>
          </td>
          <td bgcolor="#FFFFFF" align="center">
            <% if oitem.FOneItem.FsaleYn = "N" then %>
            <input type="text" name="buycash" maxlength="16" size="8" id="[on,on,off,off][���ް�]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fbuycash %>">��
            <% else %>
            <input type="text" name="buycash" maxlength="16" size="8" id="[on,on,off,off][���ް�]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Forgsuplycash %>">��
            <% end if %>
          </td>
          <% if oitem.FOneItem.FsaleYn = "N" then %>
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="margin" maxlength="32" size="5" id="[on,off,off,off][����]" value="<%= margine %>">%
          </td>
          <% else %>
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="margin" maxlength="32" size="5" id="[on,off,off,off][����]" value="<%= orgmargine %>">%
          </td>
          <% end if %>
          <td bgcolor="#FFFFFF" align="left">
            <input type="button" value="���ް� �ڵ����" onclick="CalcuAuto(itemreg);">
          </td>
        </tr>
        <tr>
          <td height="25" bgcolor="#FFFFFF"><input type="radio" name="saleYn" onClick="TnChecksaleYn(itemreg)" value="Y" <% if oitem.FOneItem.FsaleYn = "Y" then response.write "checked" %>> ���ΰ���</td>
          <input type="hidden" name="sailpricevat">
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="sailprice" maxlength="16" size="8" id="[on,on,off,off][���μҺ��ڰ�]" value="<%= oitem.FOneItem.Fsailprice %>"  onkeyup="CalcuAuto(itemreg);">��
          </td>
          <input type="hidden" name="sailsuplycashvat">
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="sailsuplycash" maxlength="16" size="8" id="[on,on,off,off][���ΰ��ް�]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fsailsuplycash %>">��
          </td>
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="sailmargin" maxlength="32" size="5" id="[on,off,off,off][���θ���]" value="<%= sailmargine %>">%
          </td>
          <td bgcolor="#FFFFFF" align="left">
            <input type="button" value="���ް� �ڵ����" onclick="CalcuAuto(itemreg);">
			<%
				dim itemSalePer : itemSalePer=0
				if oitem.FOneItem.FsaleYn="Y" then
					itemSalePer = oitem.FOneItem.Forgprice - oitem.FOneItem.Fsailprice
					itemSalePer = itemSalePer/oitem.FOneItem.Forgprice*100
				end if
			%>
			<span id="lyrPct" style="white-space:nowrap;"><% if itemSalePer>0 then %>������: <font color="#EE0000"><strong><%=formatNumber(itemSalePer,1)%>%</strong></font><% end if %></span>
          </td>
        </tr>
      </table>
      <br>
      - ���ް��� <b>�ΰ��� ���԰�</b>�Դϴ�.<br>
      - �Һ��ڰ�(���ΰ�)�� ����(���θ���)�� �Է��ϰ� [���ް��ڵ����] ��ư�� ������ ���ް��� ���ϸ����� �ڵ����˴ϴ�.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">���ϸ��� :</td>
  	<td width="35%" bgcolor="#FFFFFF"><input type="text" name="mileage" maxlength="32" size="10" id="[on,on,off,off][���ϸ���]" value="<%= oitem.FOneItem.Fmileage %>">point</td>
  	<td width="15%" bgcolor="#DDDDFF">����, �鼼 ���� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="vatYn" value="Y" <% if oitem.FOneItem.FvatYn = "Y" then response.write "checked" %>>����
      <input type="radio" name="vatYn" value="N" <% if oitem.FOneItem.FvatYn = "N" then response.write "checked" %>>�鼼
  	</td>
  </tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>�Ǹ�����
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">����Ư������ :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<%= oitem.FOneItem.Fmwdiv %> <font color="red">**����Ұ�</font>
		<input type="hidden" name="mwdiv" value="<%= oitem.FOneItem.Fmwdiv %>">
  	</td>
</tr>
<tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�����å���� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<input type="radio" name="deliverytype" value="1" <% if oitem.FOneItem.Fdeliverytype = "1" then response.write "checked" %>>�ٹ����ٹ��&nbsp;
		<input type="radio" name="deliverytype" value="2" <% if oitem.FOneItem.Fdeliverytype = "2" then response.write "checked" %>>��ü(����)���&nbsp;
		<input type="radio" name="deliverytype" value="4" <% if oitem.FOneItem.Fdeliverytype = "4" then response.write "checked" %>>�ٹ����ٹ�����
      	<input type="radio" name="deliverytype" value="9" <% if oitem.FOneItem.Fdeliverytype = "9" then response.write "checked" %>>��ü���ǹ��(���� ��ۺ�ΰ�)
  	  	<input type="radio" name="deliverytype" value="7" <% if oitem.FOneItem.Fdeliverytype = "7" then response.write "checked" %>>��ü���ҹ��
  	</td>
</tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�Ǹſ��� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="sellyn" value="Y" <% if oitem.FOneItem.Fsellyn = "Y" then response.write "checked" %>>�Ǹ���&nbsp;&nbsp;
  	  <input type="radio" name="sellyn" value="S" <% if oitem.FOneItem.Fsellyn = "S" then response.write "checked" %>>�Ͻ�ǰ��&nbsp;&nbsp;
  	  <input type="radio" name="sellyn" value="N" <% if oitem.FOneItem.Fsellyn = "N" then response.write "checked" %>>�Ǹž���
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">��뿩�� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	    <input type="radio" name="isusing" value="Y" <% if oitem.FOneItem.Fisusing = "Y" then response.write "checked" %>>�����&nbsp;&nbsp;
  	    <input type="radio" name="isusing" value="N" <% if oitem.FOneItem.Fisusing = "N" then response.write "checked" %>>������
  	</td>
  </tr>
</table>



<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="�����ϱ�" onClick="SubmitSave()">
          <input type="button" value="����ϱ�" onClick="self.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<p>
<script language='javascript'>
// ����Ư������ �� ��۱��м���
for (var i = 0; i < itemreg.elements.length; i++) {
    if (itemreg.elements(i).name == "deliverytype") {
        if (itemreg.elements(i).value == "<%= oitem.FOneItem.Fdeliverytype %>") {
            itemreg.elements(i).checked = true;
        }
    }
}

// ����
CheckSailEnDisabled(itemreg);
</script>
<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
