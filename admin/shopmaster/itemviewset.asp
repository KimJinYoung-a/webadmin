<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%

dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, deliverytype
dim cdl, cdm, cds
dim page, pageSize, dispCate, isDeal, dealYn, itemDiv

itemid      = html2db(request("itemid"))
itemname    = requestCheckvar(request("itemname"),100)
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),10)
usingyn     = requestCheckvar(request("usingyn"),10)
danjongyn   = requestCheckvar(request("danjongyn"),10)
mwdiv       = requestCheckvar(request("mwdiv"),10)
limityn     = requestCheckvar(request("limityn"),10)
vatyn       = requestCheckvar(request("vatyn"),10)
sailyn      = requestCheckvar(request("sailyn"),10)
deliverytype= requestCheckvar(request("deliverytype"),10)
dispCate = requestCheckvar(request("disp"),16)
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
isDeal = requestCheckvar(request("isDeal"),2)

page = requestCheckvar(request("page"),10)
pageSize = requestCheckvar(request("pagesize"),10)

if (page="") then page=1
if (pageSize="") then pageSize=100
if (pageSize>10000) then pageSize=100
if isDeal="Y" then
	dealYn="N"	'Y:����ǰ����, N:����ǰ����
	itemDiv="21"
end if
 
if itemid<>"" then
	dim iA ,arrTemp,arrItemid 
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
 
	iA = 0
	do while iA <= ubound(arrTemp) 	
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then 
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if 
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if
 
dim oitem

set oitem = new CItem

oitem.FPageSize         = pageSize
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectDanjongyn    = danjongyn
oitem.FRectLimityn      = limityn
oitem.FRectMWDiv        = mwdiv
oitem.FRectVatYn        = vatyn
oitem.FRectSailYn       = sailyn

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectSellReserve	= "Y"
oitem.FRectdeliverytype = deliverytype
oitem.FRectDealYn		= dealYn
oitem.FRectItemDiv		= itemDiv
oitem.GetItemList

dim i
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function ItemViewsetSave(){
	var frm = document.frmSvArr;
	var itemKey = "";
	var pass = false;
    var val_mwdiv, val_deliveryTypePolicy, val_defaultDeliveryType;
    var val_sellyn,val_danjongyn, val_limityn;


    if (!frm.cksel.length){
        if (frm.cksel.checked){
            //��ۺ� ��å ����
            itemKey = frm.cksel.value;
            val_mwdiv = eval("frm.mwdiv_" + itemKey ).value;
            val_deliveryTypePolicy = eval("frm.deliveryTypePolicy_" + itemKey ).value;
            val_defaultDeliveryType = eval("frm.defaultDeliveryType_" + itemKey ).value;

            val_sellyn    = getFieldValue(eval("frm.sellyn_" + itemKey ));
            val_danjongyn = eval("frm.danjongyn_" + itemKey ).value;
            val_limityn   = eval("frm.limityn_" + itemKey ).value;
            if ((val_mwdiv!="U")&&((val_deliveryTypePolicy=="2")||(val_deliveryTypePolicy=="9")||(val_deliveryTypePolicy=="7"))){
                alert('[' + itemKey + '] - ���� ��Ź�� ��ü ��ۺ� ��å���� ������ �� �����ϴ�.');
                return;
            }

            if ((val_mwdiv=="U")&&((val_deliveryTypePolicy=="1")||(val_deliveryTypePolicy=="4"))){
                alert('[' + itemKey + '] - ��ü����� �ٹ����� ��ۺ� ��å���� ������ �� �����ϴ�.');
                return;
            }

            if (((val_defaultDeliveryType=="9")||(val_defaultDeliveryType=="7"))&&((val_mwdiv=="W")||(val_mwdiv=="M"))){
                alert('[' + itemKey + '] - ��ü ���� �� ���� ��� �귣��� �ٹ����� ��ۺ� ��å���� ������ �� �����ϴ�.');
                return;
            }

            if ((val_defaultDeliveryType=="")&&((val_deliveryTypePolicy=="9")||(val_deliveryTypePolicy=="7"))){
                alert('[' + itemKey + '] - ��ü����/��ü���� ��� �귣�尡 �ƴմϴ�.');
                return;
            }
            if ((val_defaultDeliveryType=="9")&&(val_deliveryTypePolicy=="7")){
                alert('[' + itemKey + '] - ��ü ���ǹ�� �귣��� ��ü ���ҹ�ۺ�� ������ �� �����ϴ�.');
                return;
            }
            if ((val_defaultDeliveryType=="7")&&(val_deliveryTypePolicy=="9")){
                alert('[' + itemKey + '] - ��ü ���ҹ�� �귣��� ��ü ���ǹ�ۺ�� ������ �� �����ϴ�.');
                return;
            }
		/*
            if ((val_defaultDeliveryType=="9")&&(val_deliveryTypePolicy!="9")){
                alert('[' + itemKey + '] - ��ü ���� ��ۺ� ��å�� �ƴ� �������� �ϰ� ���� �Ұ��� �մϴ�.');
                return;
            }

            if ((val_defaultDeliveryType=="7")&&(val_deliveryTypePolicy!="7")){
                alert('[' + itemKey + '] - ��ü ���� ��ۺ� ��å�� �ƴ� �������� �ϰ� ���� �Ұ��� �մϴ�.');
                return;
            }
		*/

            //�ٹ��̰�, �Ǹ������� ��� ���� ���� �Ͽ��� ��.
            if ((val_mwdiv!="U")&&(val_sellyn=="N")&&(val_danjongyn=="N")){
                alert('[' + itemKey + '] - �Ǹű����� N �ΰ�� ������,����ǰ�� �Ǵ� MDǰ���� �����ϼž� �մϴ�.(�ٹ��)');
                return;
            }

            //�ٹ��̰�, �Ǹ����̸�, ���������� ��� �����Ǹſ�����.
            if ((val_mwdiv!="U")&&(val_sellyn=="Y")&&(val_danjongyn!="N")&&(val_limityn=="N")){
                alert('[' + itemKey + '] - �Ǹű����� Y �ΰ�� ������,����ǰ�� �Ǵ� MDǰ���� ���� �Ϸ��� ���� ������ �ϼž� �մϴ�.(�ٹ��)\n���������� �� ���������� �����Ͻñ�ٶ�');
                return;
            }

            pass = true;
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemKey = frm.cksel[i].value;
                val_mwdiv = eval("frm.mwdiv_" + itemKey ).value;
                val_deliveryTypePolicy = eval("frm.deliveryTypePolicy_" + itemKey ).value;
                val_defaultDeliveryType = eval("frm.defaultDeliveryType_" + itemKey ).value;

                val_sellyn    = getFieldValue(eval("frm.sellyn_" + itemKey ));
                val_danjongyn = eval("frm.danjongyn_" + itemKey ).value;
                val_limityn   = eval("frm.limityn_" + itemKey ).value;

                if ((val_mwdiv!="U")&&((val_deliveryTypePolicy=="2")||(val_deliveryTypePolicy=="9")||(val_deliveryTypePolicy=="7"))){
                    alert('[' + itemKey + '] - ���� ��Ź�� ��ü ��ۺ� ��å���� ������ �� �����ϴ�.');
                    frm.cksel[i].focus();
                    return;
                }

                if ((val_mwdiv=="U")&&((val_deliveryTypePolicy=="1")||(val_deliveryTypePolicy=="4"))){
                    alert('[' + itemKey + '] - ��ü����� �ٹ����� ��ۺ� ��å���� ������ �� �����ϴ�.');
                    frm.cksel[i].focus();
                    return;
                }

	            if (((val_defaultDeliveryType=="9")||(val_defaultDeliveryType=="7"))&&((val_mwdiv=="W")||(val_mwdiv=="M"))){
	                alert('[' + itemKey + '] - ��ü ���� �� ���� ��� �귣��� �ٹ����� ��ۺ� ��å���� ������ �� �����ϴ�.');
	                return;
	            }

                if ((val_defaultDeliveryType=="")&&((val_deliveryTypePolicy=="9")||(val_deliveryTypePolicy=="7"))){
                    alert('[' + itemKey + '] - ��ü����/��ü���� ��� �귣�尡 �ƴմϴ�.');
                    return;
                }
                if ((val_defaultDeliveryType=="9")&&(val_deliveryTypePolicy=="7")){
                    alert('[' + itemKey + '] - ��ü ���ǹ�� �귣��� ��ü ���ҹ�ۺ�� ������ �� �����ϴ�.');
                    return;
                }

                if ((val_defaultDeliveryType=="7")&&(val_deliveryTypePolicy=="9")){
                    alert('[' + itemKey + '] - ��ü ���ҹ�� �귣��� ��ü ���ǹ�ۺ�� ������ �� �����ϴ�.');
                    return;
                }

			/*
                if ((val_defaultDeliveryType=="9")&&(val_deliveryTypePolicy!="9")){
                    alert('[' + itemKey + '] - ��ü ���� ��ۺ� ��å�� �ƴ� �������� �ϰ� ���� �Ұ��� �մϴ�.');
                    return;
                }

                if ((val_defaultDeliveryType=="7")&&(val_deliveryTypePolicy!="7")){
                    alert('[' + itemKey + '] - ��ü ���� ��ۺ� ��å�� �ƴ� �������� �ϰ� ���� �Ұ��� �մϴ�.');
                    return;
                }
			*/

            //�ٹ��̰�, �Ǹ������� ��� ���� ���� �Ͽ��� ��.
            if ((val_mwdiv!="U")&&(val_sellyn=="N")&&(val_danjongyn=="N")){
                alert('[' + itemKey + '] - �Ǹű����� N �ΰ�� ������,����ǰ�� �Ǵ� MDǰ���� �����ϼž� �մϴ�.(�ٹ��)');
                return;
            }

            //�ٹ��̰�, �Ǹ����̸�, ���������� ��� �����Ǹſ�����.
            if ((val_mwdiv!="U")&&(val_sellyn=="Y")&&(val_danjongyn!="N")&&(val_limityn=="N")){
                alert('[' + itemKey + '] - �Ǹű����� Y �ΰ�� ������,����ǰ�� �Ǵ� MDǰ���� ���� �Ϸ��� ���� ������ �ϼž� �մϴ�.(�ٹ��)\n���������� �� ���������� �����Ͻñ�ٶ�');
                return;
            }

                pass = true;
            }
        }
    }

	if (!pass) {
		alert("���� �������� �����ϴ�.");
		return;
	}

	var schFrm = document.frm;
	schFrm.page.value="<%=page%>";
	frm.preparam.value=$(schFrm).serialize();

	if (confirm('���� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?')){
	    frm.submit();
	}

}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=800,scrollbars=yes')
}

function CheckComboOBJChange(comp,flag,objName){
	var frm = document.frmSvArr;
	var itemKey = "";
	var pass = false;
	var comp;

    if (!frm.cksel.length){
        if (frm.cksel.checked){
            itemKey = frm.itemid.value;
            comp = eval("frm." + objName + "_" + itemKey );
            comp.value=flag;

            pass = true;
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemKey = frm.itemid[i].value;
                comp = eval("frm." + objName + "_" + itemKey);
                comp.value=flag;
                pass = true;
            }
        }
    }

	if (!pass) {
		alert("���� �������� �����ϴ�. ���� ���� �Ϸ��� ��ǰ�� ���� �ϼ���");
		comp.value = '';
		return;
	}
}


function CheckRadioOBJChange(comp,flag,objName){
	var frm = document.frmSvArr;
	var itemKey = "";
	var pass = false;
	var comp;
  var icount;icount =0 ;
  frm.dSR.value = "";
  document.frmReserve.Ritemid.value ="";

    if (!frm.cksel.length){
        if (frm.cksel.checked){
            itemKey = frm.itemid.value;
            document.frmReserve.Ritemid.value =  itemKey; 
            comp = eval("frm." + objName + "_" + itemKey);
             eval("document.all.sellreserve_"+itemKey).innerHTML = "";

            for (var j=0;j<comp.length;j++){
                if (comp[j].value==flag){
                    comp[j].checked = true;
                }
            }

            pass = true;
            icount = 1;
            
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemKey = frm.itemid[i].value;
                if(document.frmReserve.Ritemid.value ==""){
                	  document.frmReserve.Ritemid.value =  itemKey; 
                }else{
                  document.frmReserve.Ritemid.value =  document.frmReserve.Ritemid.value+","+ itemKey; 
               }
                comp = eval("frm." + objName + "_" + itemKey);
            	eval("document.all.sellreserve_"+itemKey).innerHTML = "";

                for (var j=0;j<comp.length;j++){
                    if (comp[j].value==flag){
                        comp[j].checked = true;
                    }
                }
                pass = true;
                icount = icount + 1;
            }
        }
    }

	if (!pass) {
		alert("���� �������� �����ϴ�. ���� ���� �Ϸ��� ��ǰ�� ���� �ϼ���");
		comp.value = '';
		return;
	}

//���¿���
	if(frm.sellynChange.value == "R"){
		winSR = window.open("","popSR","width=400, height=200");
		document.frmReserve.action="popSellReserve.asp";
		document.frmReserve.target="popSR"; 
		document.frmReserve.iCnt.value = icount;
		document.frmReserve.submit();
		winSR.focus();
	}
}


function ChkThisRow(itemid){ 
	if (eval("document.frmSvArr.usingyn_"+itemid)[1].checked){
		if(!eval("document.frmSvArr.sellyn_"+itemid)[2].checked){
			alert("��뿩�ΰ� N�϶� �Ǹſ��ε� N���� ����˴ϴ�.");
		}
		eval("document.frmSvArr.sellyn_"+itemid)[0].checked= false;
		eval("document.frmSvArr.sellyn_"+itemid)[1].checked= false;
		eval("document.frmSvArr.sellyn_"+itemid)[2].checked= true;
	} 
}


function IsDigit(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("0123456789".indexOf(v.charAt(j)) < 0) {
			return false;
		}

		//if ((v.charAt(j) * 0 == 0) == false){
		//	return false;
		//}
	}
	return true;
}

//�˻�
function jsSearch(){   
	//��ǰ�ڵ� ����&���͸� �Է°����ϵ��� üũ-----------------------------
	var itemid = document.frm.itemid.value;  
	 itemid =  itemid.replace(",","\r");    //�޸��� �ٹٲ�ó�� 
		 for(i=0;i<itemid.length;i++){ 
			if ( itemid.charCodeAt(i) != "13" && itemid.charCodeAt(i) != "10" && "0123456789".indexOf(itemid.charAt(i)) < 0){ 
					alert("��ǰ�ڵ�� ���ڸ� �Է°����մϴ�.");
					return;
			}
		}  
	//---------------------------------------------------------------------
	
	document.frm.submit();
}
</script>

<form name="frmReserve" method="post">
	<input type="hidden" name="Ritemid" value=""> 
	<input type="hidden" name="iCnt" value="">
</form>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="POST" action="itemviewset.asp">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<table border="0" cellpadding="5" cellspacing="0" class="a">
				<tr>
					<td>�귣��: <%	drawSelectBoxDesignerWithName "makerid", makerid %> </td> 
					<td>��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"></td>
				  <td>��ǰ�ڵ�:</td>
					<td rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>  </td> 
				</tr>
				<tr>
					<td colspan="3">>����<!-- #include virtual="/common/module/categoryselectbox.asp"--> ���� ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></td>
				</tr>
			</table>
		</td> 
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			�Ǹ�:
			   <select class="select" name="sellyn">
   <option value="">��ü</option>
   <option value="Y" <% if sellyn="Y" then response.write "selected" %> >�Ǹ�</option>
   <option value="S" <% if sellyn="S" then response.write "selected" %> >�Ͻ�ǰ��</option>
   <option value="N" <% if sellyn="N" then response.write "selected" %> >ǰ��</option>
   <option value="YS" <% if sellyn="YS" then response.write "selected" %> >�Ǹ�+�Ͻ�ǰ��</option>
   <option value="SR" <% if sellyn="SR" then response.write "selected" %> >���¿���</option>
   </select>
	     	&nbsp;
	     	���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;
	     	����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	     	&nbsp;
	     	����:<% drawSelectBoxLimitYN "limityn", limityn %>
	     	&nbsp;
	     	�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
	     	����: <% drawSelectBoxVatYN "vatyn", vatyn %>
	     	&nbsp;
	     	����: <% drawSelectBoxSailYN "sailyn", sailyn %>
	     	&nbsp;
	     	���: <% drawBeadalDiv "deliverytype", deliverytype %>
			&nbsp;
			<label><input type="checkbox" name="isDeal" value="Y" <%=chkIIF(isDeal="Y","checked","")%> /> ����ǰ ����</label>
			&nbsp;
			<select name="pagesize">
				<option value="100" <%=chkIIF(pageSize=100,"selected","")%>>100</option>
				<option value="200" <%=chkIIF(pageSize=200,"selected","")%>>200</option>
			</select>
			���� ����
		</td>
	</tr>
    </form>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="��ü����" onClick="fnCheckAll(true,frmSvArr.cksel);">
			&nbsp;
			<input type="button" class="button" value="���û�ǰ����" onClick="ItemViewsetSave()">
			&nbsp;
		</td>
		<td align="right">
		<!--
			<input type="button" class="button" value="���û�ǰ �Ǹ�Y" onClick="CheckRadioOBJChange(this,'Y','sellyn')">
			&nbsp;<input type="button" class="button" value="���û�ǰ �Ͻ�ǰ��S" onClick="CheckRadioOBJChange(this,'S','sellyn')">
			&nbsp;<input type="button" class="button" value="���û�ǰ �Ǹ�N" onClick="CheckRadioOBJChange(this,'N','sellyn')">
			&nbsp;<input type="button" class="button" value="���û�ǰ ���Y" onClick="CheckRadioOBJChange(this,'Y','usingyn')">
			&nbsp;<input type="button" class="button" value="���û�ǰ ���N" onClick="CheckRadioOBJChange(this,'N','usingyn')">
		-->
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="itemSet_Process.asp">
	<input type="hidden" name="mode" value="ModiSellArr">
	<input type="hidden" name="dSR" value="">
	<input type="hidden" name="preparam" value="">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= FormatNumber(oitem.FTotalCount,0) %></b>
			&nbsp;
			������ : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(oitem.FTotalPage,0) %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
	    <td colspan="5" align="right">��ǰ ���� �� ���� ���� ���ý� �ϰ� ���� =&gt</td>
	    <td >
			<select name="limitChange" onChange="CheckComboOBJChange(this,this.value,'limityn')">
	        <option value="" >��������</option>
            <option value="Y" >����</option>
            <option value="N" >������</option>
            </select>
	    </td>
	    <td > <!--2014.06.10 �ŷ����� �뷮���� ���ϵ��� ����.�̹����̻��/������ ó��
	     <!--   <select name="mwdivChange" onChange="CheckComboOBJChange(this,this.value,'mwdiv')">
	        <option value="" >�ŷ�����</option>
            <option value="W" >��Ź</option>
            <option value="M" >����</option>
            <option value="U" >��ü</option>
            </select> -->
	    </td>
	    <td >
	        <select name="deliveryTypePolicyChange" onChange="CheckComboOBJChange(this,this.value,'deliveryTypePolicy')">
	        <option value="" >��ۺ񱸺�</option>
            <option value="1" >�ٹ����ٹ��</option>
            <option value="4" >�ٹ����ٹ�����</option>
            <option value="2" >��ü������</option>
            <option value="9" >��ü���ǹ��</option>
            <option value="7" >��ü���ҹ��</option>
            </select>
	    </td>
	    <td >
	        <select name="sellynChange" onChange="CheckRadioOBJChange(this,this.value,'sellyn')">
	        <option value="" >�Ǹű���</option>
            <option value="Y" >�Ǹ�</option>
            <option value="S" >�Ͻ�ǰ��</option>
            <option value="N" >ǰ��</option>
            <option value="R">���¿���</option>
            </select>
	    </td>
	    <td >
	        <select name="danjongynChange" onChange="CheckComboOBJChange(this,this.value,'danjongyn')">
	        <option value="" >��������</option>
            <option value="N" >������</option>
            <option value="S" >������</option>
            <option value="Y" >����ǰ��</option>
            <option value="M">MDǰ��</option>
            </select>
	    </td>
	    <td >
	        <select name="usingynChange" onChange="CheckRadioOBJChange(this,this.value,'usingyn')">
	        <option value="" >��뱸��</option>
            <option value="Y" >�����</option>
            <option value="N" >������</option>
            </select>
	    </td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td><input type="checkbox" name="ckall" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
		<td width="50">�̹���</td>
		<td width="50">��ǰ�ڵ�</td>
		<td>��ǰ��</td>
		<td>�귣��ID</td>
		<td width="80">��������</td>
		<td width="100">�ŷ�����</td>
		<td width="100">��۱���</td>
		<td width="90">�Ǹſ���</td>
		<td width="90">��������</td>
		<td width="70">��뿩��</td>
	</tr>
	<% for i=0 to oitem.FresultCount-1 %>
	<input type="hidden" name="itemid" value="<%= oitem.FItemList(i).FItemID %>">
	<tr align="center" bgcolor="FFFFFF">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" value="<%= oitem.FItemList(i).FItemID %>"></td>
		<td><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50"></td>
		<td><a href="javascript:PopItemSellEdit('<%= oitem.FItemList(i).FItemID %>');"><%= oitem.FItemList(i).FItemID %></a></td>
		<td align="left"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemID %>" target="_blank"><%= oitem.FItemList(i).FItemName %></a></td>
		<td><%= oitem.FItemList(i).FMakerID %></td>
		<td>
		    <% if (oitem.FItemList(i).Flimityn="Y") then %>
            ���� (<%=oitem.FItemList(i).GetLimitEa%>)<br />
			<label><input type="radio" name="limityn_<%= oitem.FItemList(i).FItemID %>" value="Y" <%=chkIIF(oitem.FItemList(i).Flimityn="Y","checked","") %> />Y</label>
			<label><input type="radio" name="limityn_<%= oitem.FItemList(i).FItemID %>" value="N" <%=chkIIF(oitem.FItemList(i).Flimityn="N","checked","") %> />N</label>
		    <% else %>
			<input type="hidden" name="limityn_<%= oitem.FItemList(i).FItemID %>" value="<%=oitem.FItemList(i).Flimityn%>" />
			<% end if %>
			<input type="hidden" name="orgLimityn_<%= oitem.FItemList(i).FItemID %>" value="<%=oitem.FItemList(i).Flimityn%>" />
		</td>
		<td>
		    <font color="<%= MwDivColor(oitem.FItemList(i).FMwDiv) %>"><%= oitem.FItemList(i).GetMwDivName %></font>
		    <input type="hidden" name="mwdiv_<%= oitem.FItemList(i).FItemID %>" value="<%=oitem.FItemList(i).FMwDiv%>">
		   <!-- &nbsp;
		    <select class="select" name="mwdiv_<%= oitem.FItemList(i).FItemID %>">
            <option value="W" <%= ChkIIF(oitem.FItemList(i).FMwDiv="W","selected","") %> >��Ź</option>
            <option value="M" <%= ChkIIF(oitem.FItemList(i).FMwDiv="M","selected","") %> >����</option>
            <option value="U" <%= ChkIIF(oitem.FItemList(i).FMwDiv="U","selected","") %> >��ü</option>
            </select> -->
		</td>
		<td>
		    <select class="select" name="deliveryTypePolicy_<%= oitem.FItemList(i).FItemID %>">
            <option value="1" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="1","selected","") %> >�ٹ����ٹ��</option>
            <option value="4" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="4","selected","") %> >�ٹ����ٹ�����</option>
            <option value="2" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="2","selected","") %> >��ü������</option>
            <option value="9" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="9","selected","") %> >��ü���ǹ��</option>
            <option value="7" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="7","selected","") %> >��ü���ҹ��</option>
            </select>
            <input type="hidden" name="defaultDeliveryType_<%= oitem.FItemList(i).FItemID %>" value="<%= oitem.FItemList(i).FdefaultDeliveryType %>">
            <!--input type="text" name="realstock_<%= oitem.FItemList(i).FItemID %>" value="<%= oitem.FItemList(i).Frealstock %>"-->
		</td>
		<td>
			<label><input type="radio" name="sellyn_<%= oitem.FItemList(i).FItemID %>" value="Y" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).FSellYn="Y" then response.write "checked" %>>Y</label>
			<label><input type="radio" name="sellyn_<%= oitem.FItemList(i).FItemID %>" value="S" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).FSellYn="S" then response.write "checked ><font color=blue>S</font>" else response.write ">S" %></label>
			<label><input type="radio" name="sellyn_<%= oitem.FItemList(i).FItemID %>" value="N" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).FSellYn="N" then response.write "checked ><font color=red>N</font>" else response.write ">N" %></label>
			<input type="hidden" name="defaultsellyn_<%= oitem.FItemList(i).FItemID %>" value="<%= oitem.FItemList(i).FSellYn %>">
			<div id="sellreserve_<%= oitem.FItemList(i).FItemID %>"  style="padding:3"><%IF not isNull(oitem.FItemList(i).Fsellreservedate) THEN %><font color="blue">���¿���: <%=oitem.FItemList(i).Fsellreservedate%></font><%END IF%></div>
		</td>
		<td>
		    <select class="select" name="danjongyn_<%= oitem.FItemList(i).FItemID %>">
            <option value="N" <%= ChkIIF(oitem.FItemList(i).Fdanjongyn="N","selected","") %> >������</option>
            <option value="S" <%= ChkIIF(oitem.FItemList(i).Fdanjongyn="S","selected","") %> >������</option>
            <option value="Y" <%= ChkIIF(oitem.FItemList(i).Fdanjongyn="Y","selected","") %> >����ǰ��</option>
            <option value="M" <%= ChkIIF(oitem.FItemList(i).Fdanjongyn="M","selected","") %> >MDǰ��</option>
            </select>
		</td>
		<td>
			<label><input type="radio" name="usingyn_<%= oitem.FItemList(i).FItemID %>" value="Y" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).Fisusing="Y" then response.write "checked" %>>Y</label>
			<label><input type="radio" name="usingyn_<%= oitem.FItemList(i).FItemID %>" value="N" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).Fisusing="N" then response.write "checked ><font color=red>N</font>" else response.write ">N" %></label>
		</td>
	</tr>
	<%
			if i mod 250 = 0 then
				Response.Flush		' ���۸��÷���
			end if
		next
	%>
</form>
	<tr bgcolor="FFFFFF">
		<td colspan="11" align="center">
		<% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>


</table>

<form name="frmArrupdate" method="post" action="doItemSellSet.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="sellyn" value="">
<input type="hidden" name="usingyn" value="">
<input type="hidden" name="packyn" value="">
</form>
<br>
<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
