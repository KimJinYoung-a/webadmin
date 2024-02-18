<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ����
' History : �̻� ����
'			2022.01.19 �ѿ�� ����(�����ù�,����� �߰�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/checknoticls.asp"-->
<%
dim BasicMonth
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)
%>

<script type='text/javascript'>

function publicbarreg(barcode){
	var popwin = window.open('/common/popbarcode_input.asp?itembarcode=' + barcode,'popbarcode_input','width=500,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popBrandRackCodeEdit(imakerid){
    var popwin = window.open('pop_BrandRackCodeEdit.asp?makerid=' + imakerid,'popBrandRackCodeEdit','width=500,height=200,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function itemrackcodereg(itemrackcode){
	var popwin = window.open('popitemrackcode_input.asp?itemrackcode=' + itemrackcode,'popitemrackcode_input','width=300,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popItemRackCodeEdit(itemrackcode){
    var popwin = window.open('pop_ItemRackCodeEdit.asp?itemrackcode=' + itemrackcode,'pop_ItemRackCodeEdit','width=400,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popitemsearch(barcode){
	var popwin = window.open('pop_item_search.asp?barcode=' + barcode,'popitemsearch','width=500,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popRealErrInput(itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poprealerrinput.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&BasicMonth=<%= BasicMonth %>','poprealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popRealStockTaking(iitemid){
    var popwin = window.open('/admin/stock/jaegoadd.asp?itemid='+ iitemid,'poprealstockTaking','width=900,height=460,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function PopBadItemInput(){
	var popwin = window.open('/common/pop_baditem_input.asp','pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'simpleitemedit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemWeightEdit(iitemid){
	var popwin = window.open('pop_ItemWeightEdit.asp?itemid=' + iitemid,'itemWeightEdit','width=500,height=300,scrollbars=yes,resizable=yes')
}

function PopSongjangList(songjangdiv){
	var popwin = window.open('/warehouse/pop_SongjangList.asp?songjangdiv='+songjangdiv,'PopSongjangList','width=500,height=300,scrollbars=yes,resizable=yes')
}

function reSearchExcelDown(select_type){
    alert('��ٷ� �ּ���. �ۼ����Դϴ�.');
	frm.target = "exceldown";
	frm.action = "/admin/ordermaster/chulgobogo_ExcelDown.asp"
    frm.select_type.value=select_type;
    frm.submit();
	frm.target = "";
	frm.action = "";
    frm.select_type.value='';
}

</script>


<a href="javascript:publicbarreg('');">������ڵ���</a>
<br>
<a href="javascript:popBrandRackCodeEdit('');">�귣�巢�ڵ���</a>
<br>
<a href="javascript:itemrackcodereg('');">���ڵ庰��ǰ�Է�</a>
<br>
<a href="javascript:popItemRackCodeEdit('');">��ǰ�����ڵ��Է�</a>
<br>
<a href="javascript:popRealStockTaking('');">�������</a>
<br>
<a href="javascript:popRealErrInput('','','');">���(����)�Է�</a>

<br>
<a href="javascript:PopBadItemInput();">�ҷ����</a>
<br>
<a href="javascript:popitemsearch('');">��ǰ�˻�</a>
<br>
<a href="javascript:PopItemSellEdit('');">��ǰ�Ӽ�����</a>
<br>
<a href="javascript:PopItemWeightEdit('');">��ǰ�����Է�</a>
<br>
������ �����ޱ� : 
<a href="#" onclick="PopSongjangList('2'); return false;">�Ե��ù�</a>
/ <a href="#" onclick="PopSongjangList('1'); return false;">�����ù�</a>
/ <a href="#" onclick="PopSongjangList('4'); return false;">CJ�������</a>

<hr>
<a href="undeliveredOrderList.asp">�̹�� �ֹ� ���</a>
<hr>
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="research" value="on">
<input type="hidden" name="select_type" value="">
<br>
* �����(��¥����:����)
<%
'<input type="button" class="button_s" value="�����ٿ�ε�(���Ϲ�����ֹ�)" onclick="reSearchExcelDown('samedaymichulgo');">
'<input type="button" class="button_s" value="�����ٿ�ε�(��������ֹ�)" onclick="reSearchExcelDown('delaychulgo');">
'<input type="button" class="button_s" value="�����ٿ�ε�(�������_�����Ϻ�����¥)" onclick="reSearchExcelDown('delaychulgodate');">
'<input type="button" class="button_s" value="�����ٿ�ε�(�������_�����Ϻ����ֹ�)" onclick="reSearchExcelDown('delaychulgocnt');">
%>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="300"></iframe>
<% else %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="0"></iframe>
<% end if %>



















<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->