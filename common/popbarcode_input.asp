<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itembarcodecls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
rw "��������޴�-�����ڹ��� ���"
response.end

dim itembarcode
itembarcode = requestCheckVar(request("itembarcode"),20)

dim itemgubun, itemid, itemoption

'if (Len(itembarcode)=12) then
'	itemgubun 	= left(itembarcode,2)
'	itemid		= CLng(mid(itembarcode,3,6))
'	itemoption	= right(itembarcode,4)
'else
'	itemgubun = "10"
'	itemid = itembarcode
'	itemoption  = "0000"
'end if

if BF_IsMaybeTenBarcode(itembarcode) then
    itemgubun 	= BF_GetItemGubun(itembarcode)
	itemid 		= BF_GetItemId(itembarcode)
	itemoption 	= BF_GetItemOption(itembarcode)
else
	itemgubun = "10"
	itemid = itembarcode
	itemoption  = "0000"
end if


dim oitembar
set oitembar = new CItemBarCode
oitembar.FRectItemGubun = itemgubun
oitembar.FRectItemID = itemid
'''oitembar.FRectItemoption = itemoption
if itemid<>"" then
	oitembar.getItemBarcodeInfo
end if


dim i
%>
<script language='javascript'>
function InputRackcode(frm){
	if (frm.itemrackcode.value.length!=4){
		alert('��ǰ ���ڵ带 ��Ȯ�� �Է��ϼ���. 4�ڸ�');
		frm.itemrackcode.focus();
		return;
	}

	if (confirm('��ǰ ���ڵ带 �����Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function Research(frm){
	frm.submit();
}

function InputBarcode(comp){
	var barcode = comp.value;
	var frm = document.frmsavebar;

	if ((barcode.substr(0,2)=='10')&&(barcode.length==12)){
		alert('����� �� ���� ���ڵ� �����Դϴ�.');
		comp.focus();
		return;
	}

	if (barcode.length<8){
		alert('���ڵ带 ��Ȯ�� �Է��ϼ���.');
		comp.focus();
		return;
	}

	if ((frm.itemgubun.value!="10")&&(frm.itemgubun.value!="90")&&(frm.itemgubun.value!="70")){
		alert('��ǰ���п��� - ������ ���ǿ��');
		return;
	}

	if (frm.itemid.value.length<1){
		alert('��ǰ�ڵ尡 ���ǵ��� �ʾҽ��ϴ�. - ��ǰ�˻��� ����ϼ���.');
		document.frmbar.itembarcode.focus();
		return;
	}

	frm.itemoption.value = comp.id;
	frm.publicbarcode.value = barcode;

    if (confirm('���� ���ڵ带 �����Ͻðڽ��ϱ�?')){
		frm.submit();
	}

/*	??
	if (frm.confirmbarcode.value.length<1){
	    alert('Ȯ���� ���� �ٽ��ѹ� ���ڵ带 �Է����ּ���.');
	    frm.confirmbarcode.value = barcode;
	    comp.value ='';
	    comp.focus();

	}else{
	    if (frm.confirmbarcode.value==frm.publicbarcode.value){
        	if (confirm('���� ���ڵ带 �����Ͻðڽ��ϱ�?')){
        		frm.submit();
        	}
        }else{
            frm.confirmbarcode.value = "";
            frm.publicbarcode.value = "";

            alert('���ڵ尡 ��ġ���� �ʽ��ϴ�. ó������ �ٽ� �õ��� �ּ���.');
            comp.value ='';
            comp.focus();
        }
    }
*/
}

function GetOnLoad(){
	<% if oitembar.FResultCount>0 then %>
	    if (document.frmbar.publicbar_<%= itemoption %>) {
	        document.frmbar.publicbar_<%= itemoption %>.select();
    	    document.frmbar.publicbar_<%= itemoption %>.focus();
        }

	<% else %>
	    document.frmbar.itembarcode.select();
	    document.frmbar.itembarcode.focus();
	<% end if %>
}
window.onload=GetOnLoad;
</script>

  <table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- ��ܹ� ���� -->
	<form name="frmbar" method=get>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red">&nbsp;<strong>��ǰ���ڵ��Է�</strong></font>
				    </td>
				    <td align="right">
						<input type="text" class="text"  name="itembarcode" value="<%= itembarcode %>" size=17 maxlength=14 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">
        				<input type="button" class="button" value="�˻�" onclick="Research(frmbar)" >
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->

<% if oitembar.FResultCount>0 then %>
  	<tr bgcolor="#FFFFFF">
    	<td width="80" bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
   	<td colspan="2"><%= oitembar.FItemList(0).FbrandName %>(<%= oitembar.FItemList(0).Fmakerid %>)</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
    	<td colspan="2"><%= oitembar.FItemList(0).FItemName %></td>
    </tr>
	<% for i=0 to oitembar.FResultCount-1 %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�ɼǸ�(<%= oitembar.FItemList(i).Fitemoption %>)</td>
		<% if oitembar.FItemList(i).FitemoptionName="" then %>
		<td>�ɼǾ���</td>
		<% else %>
			<% if itemoption=oitembar.FItemList(i).Fitemoption then %>
			<td><b><%= oitembar.FItemList(i).FitemoptionName %></b></td>
			<% else %>
			<td><%= oitembar.FItemList(i).FitemoptionName %></td>
			<% end if %>
		<% end if %>

		<td align="right">
		<% if oitembar.FItemList(i).Fitemoption=itemoption then %>
			<input type="text" class="text" id="<%= oitembar.FItemList(i).Fitemoption %>" name="publicbar_<%= oitembar.FItemList(i).Fitemoption %>" value="<%= oitembar.FItemList(i).FPublicBarcode %>" size=20 maxlength=20 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ InputBarcode(frmbar.publicbar_<%= oitembar.FItemList(i).Fitemoption %>); return false;}">
			<input type="button" class="button" value="���" onclick="InputBarcode(frmbar.publicbar_<%= oitembar.FItemList(i).Fitemoption %>)">
		<% else %>
		    <input type="text" class="text" id="<%= oitembar.FItemList(i).Fitemoption %>" name="publicbar_<%= oitembar.FItemList(i).Fitemoption %>" value="<%= oitembar.FItemList(i).FPublicBarcode %>" size=20 maxlength=20 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ InputBarcode(frmbar.publicbar_<%= oitembar.FItemList(i).Fitemoption %>); return false;}" disabled >
			<input type="button" class="button" value="���" onclick="InputBarcode(frmbar.publicbar_<%= oitembar.FItemList(i).Fitemoption %>)" disabled >
		<% end if %>
		</td>
	<% next %>
	</tr>
	<tr bgcolor="#FFFFFF">
    	<td bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
    	<td colspan="2"><img src="<%= oitembar.FItemList(0).FImageList %>" width="100" height="100" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
    </tr>
	</form>
	<!--
	<form name="frmitemrackcode" method=post  action="/warehouse/itemrackcode_process.asp">
	<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
	<input type="hidden" name="itemid" value="<%= itemid %>">
	<input type="hidden" name="mode" value="byitem">
    <tr bgcolor="#FFFFFF">
    	<td bgcolor="<%= adminColor("tabletop") %>">�ǸŰ�</td>
    	<td colspan="2"><%= FormatNumber(oitembar.FItemList(0).FSellcash,0) %></td>
    </tr>

    <tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ���ڵ�</td>
    	<td colspan="2">
    		<input type="text" class="text" name="itemrackcode" value="<%= oitembar.FItemList(0).Fitemrackcode %>" size="4" maxlength="4" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ InputRackcode(frmitemrackcode); return false;}">
    		<input type="button" class="button" value="����" onclick="InputRackcode(frmitemrackcode);">
    		&nbsp;
    		(�귣�巢�ڵ� : <%= oitembar.FItemList(0).Fprtidx %>)
    	</td>
    </tr>
    </form>
    -->
<% else %>
	<tr bgcolor="#FFFFFF">
    	<td colspan="3" align="center">
    		�˻������ �����ϴ�

    		<!-- <br>
    		���� 10�ڵ�(�¶��ε�ϻ�ǰ)�� ����� �����մϴ�.
    		<br>90�ڵ��� ��� ������ǰ������ �̿��ϼ���. -->
    	</td>
    </tr>
<% end if %>


</table>


<%
set oitembar = Nothing
%>
<form name="frmsavebar" method=post action="barcode_input_process.asp">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="publicbarcode" value="">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->