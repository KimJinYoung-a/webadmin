<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸�
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->

<%
'Matchitemoption �� ������Ʈ
Dim mallid, itemid, vIsOption, mode, itemoption, outmallitemid, outmallitemname, outmallPrice, outmallSellYn
Dim itemOptionname, outmallitemOptionname, itemname, sellyn, sellcash, i
	mallid	= requestCheckvar(request("mallid"),32)
	itemid	= requestCheckvar(request("itemid"),10)
	itemoption= requestCheckvar(request("itemoption"),4)

Dim oxItem
set oxItem = new CxSiteTempLinkItem
	oxItem.FRectItemID = itemid
	oxItem.FRectSellSite = mallid
	oxItem.FRectItemOption= itemoption
	
	If itemid <> "" Then
		oxItem.getOnexSiteTempLinkItem
	End If

if oxItem.fresultCount>0 then 
    mode="edit"
    outmallitemid   = oxItem.FOneItem.Foutmallitemid
    outmallitemname = oxItem.FOneItem.Foutmallitemname
    outmallitemOptionname = oxItem.FOneItem.FoutmallitemOptionname
    outmallPrice    = oxItem.FOneItem.FoutmallPrice
    outmallSellYn   = oxItem.FOneItem.FoutmallSellYn
    itemname        = oxItem.FOneItem.Fitemname
    itemOptionname  = oxItem.FOneItem.FitemOptionname
    sellyn          = oxItem.FOneItem.Fsellyn
    sellcash        = oxItem.FOneItem.Fsellcash
end if

if (itemid="") then
    mode="add"
end if
%>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<script language="javascript" src="/js/jquery-1.6.2.min.js"></script>
<script language="javascript">

function selThisItem(iitemid){
    var frm = document.frmXItem;
    var selopt ='';

	selopt = $("#vOpt").val();
    if (selopt){
        if (selopt.length!=4){
            alert('�ɼ��� ���� �ϼ���.');
            return;
        }
    }else{
        selopt = '0000';
    }
    
    frm.itemid.value=iitemid;
    frm.itemoption.value=selopt;
    
}

function searchItem(frm){
    var xitemname = escape(frm.outmallitemname.value);
    //var xitemname = encodeURI(frm.outmallitemname.value);
    //var xitemname = frm.outmallitemname.value;
    if (xitemname.length<1){
        alert('��ǰ���� �Է��� �˻��ϼ���.');
        frm.outmallitemname.focus();
        return;
    }
    
    $("#divView").html('');
    
    $.ajax({
		type: "POST",
		url: "/admin/etc/orderInput/ajxMatchXsiteItem.asp",
		data: "outmallitemname="+xitemname,
		dataType: "text",
		//timeout : 1000,
		error: function(){
			html = "/admin/etc/orderInput/ajxMatchXsiteItem.asp?"+"outmallitemname="+xitemname;
			$("#divView").html(html);
		},
		success: function(html){
			$("#divView").html(html);
			
		}
	});
}


function searchItem2(frm){
    var tenitemid = escape(frm.itemid.value);
    if (tenitemid.length<1){
        alert('��ǰ�ڵ� �Է��� �˻��ϼ���.');
        frm.itemid.focus();
        return;
    }
    
    $("#divView").html('');
    
    $.ajax({
		type: "POST",
		url: "/admin/etc/orderInput/ajxMatchXsiteItem.asp",
		data: "tenitemid="+tenitemid,
		dataType: "text",
		//timeout : 1000,
		error: function(){
			html = "/admin/etc/orderInput/ajxMatchXsiteItem.asp?"+"tenitemid="+tenitemid;
			$("#divView").html(html);
		},
		success: function(html){
			$("#divView").html(html);
			
		}
	});
}
function ModiXItem(){
	var frm = document.frmXItem;
	
	if (frm.itemid.value.length<1){
	    alert('TEN ��ǰ��ȣ �ʼ� �Դϴ�.')
	    frm.itemid.focus();
	    return;
	}
	
	if ((frm.outmallitemid.value.length<1)&&(frm.outmallitemname.value.length<1)){
	    alert('���� ��ǰ��ȣ �Ǵ� ���� ��ǰ�� �� �ϳ��� �ʼ� �Է� ���Դϴ�.')
	    frm.outmallitemid.focus();
	    return;
	}
	
	//�ݵ�ط��̽� ���� �ʼ� ��.
	if (frm.itemoption.value.length!=4){
	    if (frm.mallid.value!="bandinlunis11111111" ){			//&& frm.mallid.value!="mintstore"
	        alert('bandinlunis, mintstore �� ���� �ɼ��ڵ� �ʼ���.')
    	    frm.itemoption.focus();
    	    return;
	    }
	}
	
	if (frm.outmallPrice.value.length<1){
	    alert('���� �ǸŰ��� �ʼ� �Դϴ�.')
	    frm.outmallPrice.focus();
	    return;
	}
	
	if ((!frm.outmallSellYn[0].checked)&&(!frm.outmallSellYn[1].checked)&&(!frm.outmallSellYn[2].checked)){
	    alert('���� �Ǹ� ���θ� �����ϼ���.')
	    frm.outmallSellYn[0].focus();
	    return;
	}
	
	if (confirm('���� �Ͻðڽ��ϱ�?')){
	    frm.submit();
	}
}

function DelXItem(){
    var frm = document.frmXItem;
	
	if (frm.itemid.value.length<1){
	    alert('TEN ��ǰ��ȣ �ʼ� �Դϴ�.')
	    frm.itemid.focus();
	    return;
	}
	
	if (confirm('���� �Ͻðڽ��ϱ�?')){
	    frm.mode.value="del";
	    frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl">
<form name="frmXItem" method="post" action="partneritemlink_process.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="mallid" value="<%=mallid%>">
<tr>
	<td width="120" align="right" class="td_br_tablebar">�� ����:</td>
	<td class="td_br" colspan="2"><%= mallid %></td>
</tr>
<tr>
	<td width="120" align="right" class="td_br_tablebar">TEN ��ǰ��ȣ:</td>
	<td class="td_br" colspan="2">
	<% if mode="add" then %>
	    <input type="text" name="itemid" value="" size="6" maxlength="9">(�ʼ�) <input type="button" value="Search" onClick="searchItem2(document.frmXItem);">
	<% else %>
	    <input type="text" name="itemid" value="<%= oxItem.FOneItem.FItemId %>" size="6" maxlength="9" readonly class="text_ro">
	<% end if %>
	</td>
</tr>
<% 
if mallid="bandinlunis11111111"  then 		'/or mallid="mintstore"
%>
    <input type="hidden" name="itemoption">
    <input type="hidden" name="p_itemoption">
<% else %>
<tr>
	<td width="120" align="right" class="td_br_tablebar">TEN �ɼǹ�ȣ:</td>
	<td class="td_br" colspan="2">
	<% if mode="add" then %>
	    <input type="text" name="itemoption" value="" size="6" maxlength="9">
	<% else %>
	    <input type="hidden" name="p_itemoption" value="<%= oxItem.FOneItem.Fitemoption %>">
	    <input type="text" name="itemoption" value="<%= oxItem.FOneItem.Fitemoption %>" size="4" maxlength="4" >
	<% end if %>
	( �ʼ� )
	</td>
</tr>
<% end if %>
<tr>
	<td width="120" align="right" class="td_br_tablebar">���� ��ǰ��ȣ:</td>
	<td class="td_br" colspan="2">
	    <input type="text" name="outmallitemid" value="<%= outmallitemid %>" size="20" maxlength="20">
	    (�ֹ� ����Ʈ ������ ���� ��ǰ��ȣ�� �ִ°�� �ʼ� �Է�)
	</td>
</tr>
<tr>
	<td width="120" align="right" class="td_br_tablebar">���� ��ǰ��:</td>
	<td class="td_br">
	    <input type="text" name="outmallitemname" value="<%= outmallitemname %>" size="40" maxlength="50">
	    (�ֹ� ����Ʈ ������ ���� ��ǰ��ȣ�� ���°�� �ʼ� �Է�)
	    <% if mode="add" then %>
	    <br>
	    <table border=0 cellspacing=2 cellpadding=2>
	    <tr>
	        <td><input type="button" value="Search" onClick="searchItem(document.frmXItem);"></td>
	        <td><div id="divView"></div></td>
	    </tr>
	    </table>
	    <% end if %>
	</td>
	<% if mode<>"add" then %>
	<td width="200" class="td_br"><%= oxItem.FOneItem.Fitemname %></td>
	<% end if %>
</tr>
<% 
if mallid="bandinlunis11111111" then 		'/ or mallid="mintstore"
%>
    <input type="hidden" name="outmallitemOptionname" >
<% else %>
<tr>
	<td width="120" align="right" class="td_br_tablebar">���� �ɼǸ�:</td>
	<td class="td_br">
	    <input type="text" name="outmallitemOptionname" value="<%= outmallitemOptionname %>" size="40" maxlength="50">
	   
	    <% if mallid="hottracks" then %>
	    	<br>������Ʈ������ ��� ���� ��ǰ�� �ִ� ��ǰ��� �ɼǸ��� ����� ���ο� �Է����ּ���
	    <% end if %>
	</td>
	<% if mode<>"add" then %>
	<td width="200" class="td_br"><%= oxItem.FOneItem.FitemOptionname %></td>
	<% end if %>
</tr>	
<% end if %>
<tr>
	<td width="120" align="right" class="td_br_tablebar">���� �ǸŰ�:</td>
	<td class="td_br">
	    <input type="text" name="outmallPrice" value="<%= outmallPrice %>" size="10" maxlength="10">
	    (������ �ʿ� - �ʼ�)
	</td>
	<% if mode<>"add" then %>
	<td width="200" class="td_br"><%= oxItem.FOneItem.Fsellcash %> &nbsp;</td>
	<% end if %>
</tr>		
<tr>
	<td width="120" align="right" class="td_br_tablebar">���� �Ǹſ���:</td>
	<td class="td_br">
	    <input type="radio" name="outmallSellYn" value="Y" <%= CHKIIF(outmallSellYn="Y","checked","") %> >�Ǹ���
	    <input type="radio" name="outmallSellYn" value="N" <%= CHKIIF(outmallSellYn="N","checked","") %> >�Ǹž���
	    <input type="radio" name="outmallSellYn" value="X" <%= CHKIIF(outmallSellYn="X","checked","") %> >�Ǹ�����
	    (������ �ʿ� - �ʼ�)
	</td>
	<% if mode<>"add" then %>
	<td width="200" class="td_br"><%= oxItem.FOneItem.Fsellyn %> &nbsp;</td>
	<% end if %>
</tr>	
<tr>
	<td align="center" colspan="3" class="td_br">
	<% If mode = "add" Then %>
	    <input type="button" class="button" value="�߰�" onClick="ModiXItem();">
	<% Else %>
		<input type="button" class="button" value="����" onClick="ModiXItem()">
		&nbsp;
		<input type="button" class="button" value="����" onClick="DelXItem()">
		&nbsp;
		<input type="button" class="button" value="�ݱ�" onClick="self.close()">
	<% End If %>
	</td>
</tr>
</form>	
</table>

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td align="center" width=200>
		���޻�ǰ��� ���޿ɼǸ�����<br>��Ī�ϴ� ���޸� : 
	</td>
	<td align="left">
		<% GetItemMaeching_itemname_itemoptionname_list() %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<%
SET oxItem= Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->