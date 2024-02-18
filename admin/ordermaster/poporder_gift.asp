<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<%

dim baljuid, page, isupchebeasong
dim research

baljuid         = request("baljuid")
page            = request("page")
isupchebeasong  = request("isupchebeasong")
research        = request("research")

if page="" then page=1
if research="" and isupchebeasong="" then isupchebeasong="N"

dim oOrderGift
set oOrderGift = new COrderGift
oOrderGift.FPageSize =1000
oOrderGift.FCurrPage = page
oOrderGift.FRectisupchebeasong = isupchebeasong
oOrderGift.FRectBaljuid = baljuid
oOrderGift.GetOrderGiftList

dim i
%>
<script language='javascript'>
function setGiftCode(gift_code){
    frmAct.gift_code.value = gift_code;
}

function ActReGiftMaker(frm){
    var gift_code = frm.gift_code.value;
    
    if (gift_code.length<1){
        alert('Gift �ڵ带 �־��ּ���.');
        frm.gift_code.focus();
        return;
    }else{
        if (confirm('Gift(' + gift_code + ') ��ü ����ǰ ������ ���ۼ� �Ͻðڽ��ϱ�? \n\n������ ��ü �ֹ��ǿ� ���� ���ۼ� �˴ϴ�.')){
            frm.submit();
        }
    }
}

function ActBaljuGift(frm){
    return;
    
    var baljuid = frm.baljuid.value;
    var evt_code = frm.evt_code.value;
    
    
    if (baljuid.length<1){
        alert('������� ��ȣ�� �Է� �� �˻� �ϼ���..');
        document.frm.baljuid.focus();
        return;
    }
    
    
    if (evt_code.length<1){
        //if (confirm('�ش� �������(' + baljuid + ')�� ��ü ����ǰ ������ ���ۼ� �Ͻðڽ��ϱ�?')){
        //    frm.submit();
        //}
        alert('�̺�Ʈ �ڵ带 �־��ּ���.');
        frm.evt_code.focus();
        return;
    }else{
        if (confirm('�ش� �������(' + baljuid + ')�� �̺�Ʈ(' + evt_code + ') ����ǰ ������ ���ۼ� �Ͻðڽ��ϱ�?')){
            frm.submit();
        }
    }
}
</script>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			������ù�ȣ : <input type="text" class="text" name="baljuid" value="<%= baljuid %>" size="10" maxlength="10">
        	&nbsp;
        	��۱���:
        	<input type="radio" name="isupchebeasong" value=""  <% if isupchebeasong="" then response.write "checked" %> >��ü
        	<input type="radio" name="isupchebeasong" value="N" <% if isupchebeasong="N" then response.write "checked" %> >�ٹ�
        	<input type="radio" name="isupchebeasong" value="Y" <% if isupchebeasong="Y" then response.write "checked" %> >����
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<form name="frmAct" method="post" action="poporder_gift_process.asp" onSubmit="ActReGiftMaker(frmAct); return false;">
	<input type="hidden" name="baljuid" class="text_ro" size="6" value="<%= baljuid %>">
	<tr>
		<td align="left">
			������ù�ȣ : <b><%= baljuid %></b>
			&nbsp;
            ����Ʈ��ȣ : <input type="text" name="gift_code" class="text" size="6" value="" >
            &nbsp;
            <% if (session("ssBctID")="icommang") or (session("ssBctID")="tozzinet") or (session("ssBctID")="coolhas") or (session("ssBctID")="kobula") then %>
            <input type="button" class="button" value="����ǰ������ۼ�" onclick="ActReGiftMaker(frmAct);">
        	(tbl_order_gift)
        	<% end if %>
		</td>
		<td align="right">
		
		</td>
	</tr>
	</form>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">�������ID</td>
		<td width="80">�ֹ���ȣ</td>
		<td width="50">DAS</td>
		<td width="50">EVENT<br>ID</td>
		<td>�̺�Ʈ��</td>
		<td width="50">GiftID</td>
		<td>����ǰ��</td>
		<td width="50">���<br>����</td>
		<td width="100">�Ⱓ</td>
		<td>����</td>
	</tr>
	<% for i=0 to oOrderGift.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td><%= oOrderGift.FItemList(i).FBaljuID %></td>
	    <td><%= oOrderGift.FItemList(i).Forderserial %></td>
	    <td><%= oOrderGift.FItemList(i).Fdasindex %></td>
	    <td><%= oOrderGift.FItemList(i).Fevt_code %></td>
	    <td align="left"><%= oOrderGift.FItemList(i).Fevt_name %></td>
	    <td><a href="javascript:setGiftCode('<%= oOrderGift.FItemList(i).Fgift_code %>');"><%= oOrderGift.FItemList(i).Fgift_code %></a></td>
	    <td align="left"><%= oOrderGift.FItemList(i).Fgiftkind_name %></td>
	    <td>
	    <% if oOrderGift.FItemList(i).Fisupchebeasong="Y" then %>  
	    ��ü
	    <% else %>
	    �ٹ�
	    <% end if %>  
	    </td>
	    
	    <td>
	        <%= oOrderGift.FItemList(i).Fevt_startdate %>
	        ~ <br>
	        <%= oOrderGift.FItemList(i).Fevt_enddate %>
	    </td>
	    <td>
	        <%= oOrderGift.FItemList(i).GetEventConditionStr %>
	    </td>
	</tr>
	<% next %>
</table>

<%
set oOrderGift = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->