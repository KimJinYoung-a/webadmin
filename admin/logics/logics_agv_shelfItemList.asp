<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : agv
' History : �̻� ����
'           2020.05.12 ������ ����
'           2020.05.20 �ѿ�� ����
'####################################################
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/barcodefunction.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<%
dim page, i
dim itembarcode, found
dim itemgubun, itemid, itemoption, isUsing, genBarcode
itembarcode 	= requestCheckVar(request("itembarcode"),32)
itemgubun 	= requestCheckVar(request("itemgubun"),2)
itemid 	= requestCheckVar(request("itemid"),10)
itemoption = requestCheckVar(request("itemoption"),4)
isUsing = requestCheckVar(request("isusing"),1)
page = chkIIF(page="",1,page)
isUsing = chkIIF(isUsing="","Y",isUsing)

if (itembarcode <> "") then
	if Len(itembarcode) > 8 or Not IsNumeric(itembarcode) then
		'// ���̰� 8���� ũ�ų� ���ڰ� �ƴѰ�� ���ڵ����� ���� Ȯ��
		found = fnGetItemCodeByPublicBarcode(itembarcode, itemgubun, itemid, itemoption)
	end if

	if Not found and BF_IsMaybeTenBarcode(itembarcode) = True then
		'// �ٹ����� : �����ڵ� �˻��� ���(10 111111 0000 �Ǵ� 10 01000000 0000)
		itemgubun 	= BF_GetItemGubun(itembarcode)
		itemid 		= BF_GetItemId(itembarcode)
		itemoption 	= BF_GetItemOption(itembarcode)
		found = True
	end if

	if Not found and Len(itembarcode) <= 8 and IsNumeric(itembarcode) then
		'��ǰ�ڵ�� �˻�(111111 �Ǵ� 1000000)
		itemgubun = "10"
		itemid = itembarcode
		itemoption  = "0000"
		itembarcode = BF_MakeTenBarcode(itemgubun, itemid, itemoption)
	end if

end if

dim oAGV
Set oAGV = new CAGVItems
    oAGV.FPageSize = 10000
    oAGV.FCurrPage = page
	oAGV.FRectItemGubun = itemgubun
	oAGV.FRectItemID  =itemid
	oAGV.FRectItemoption = itemoption
    oAGV.FRectIsUsing = isUsing
    oAGV.GetShelfItemList
%>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/19.1.4/css/dx.common.css" />
<link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/19.1.4/css/dx.light.compact.css" />
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.2/jszip.min.js"></script>
<script src="https://cdn3.devexpress.com/jslib/19.1.4/js/dx.all.js"></script>
<style type="text/css">
.dx-widget {font-size:12px;}
</style>
<script type="text/javascript">
function goSearch(frm) {
    frm.submit();
}

function resetSearchForm(frm) {
    frm.itembarcode.value="";
    frm.itemgubun.value="";
    frm.itemid.value="";
    frm.itemoption.value="";
    frm.isusing.value="Y";
    goSearch(frm);
}
</script>
<!-- �˻� ���� ���� -->
<form name="frm" method="get" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td>
        <table width="100%" class="a">
        <tr>
            <td align="left">
                <label>�� ���ڵ� : <input type="text" name="itembarcode" size="16" value="<%=itembarcode%>" class="text" /> </label> &nbsp;
                <label>�� ���� : <input type="text" name="itemgubun" size="2" value="<%=itemgubun%>" class="text" /> </label> &nbsp;
                <label>�� ��ǰ�ڵ� : <input type="text" name="itemid" size="10" value="<%=itemid%>" class="text" /> </label> &nbsp;
                <label>�� �ɼ��ڵ� : <input type="text" name="itemoption" size="4" value="<%=itemoption%>" class="text" /> </label> &nbsp;
                <label>�� ��뿩�� :
                    <select name="isusing" class="select">
                    <option value="A" <%=chkIIF(isUsing="A","selected","")%>>���</option>
                    <option value="Y" <%=chkIIF(isUsing="Y","selected","")%>>���</option>
                    <option value="N" <%=chkIIF(isUsing="N","selected","")%>>����</option>
                    </select>
                </label>
            </td>
            <td align="right">
                <a href="#" onClick="resetSearchForm(document.frm);" title="�˻� ������ �ʱ�ȭ�մϴ�.">Reset</a>
            </td>
        </tr>
        </table>
    </td>
    <td width="80" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goSearch(document.frm);">
	</td>
</tr>
</table>
</form>
<!-- �˻� ���� �� -->
<!-- ������ �׸��� ���� -->
<div class="dx-viewport" style="margin-top:5px;padding:5px;">
    <div class="demo-container">
        <div id="gridContainer"><center>No Data...</center></div>
    </div>
</div>
<!-- ������ �׸��� �� -->
<script type="text/javascript">
$(function(){
    $("#gridContainer").dxDataGrid({
        showColumnLines: true, // �÷� ����
        showRowLines: true, // �ο� ����
        rowAlternationEnabled: true, // �ο캰 ȸ�� ����
        showBorders: true, // ��ü ����
        columnChooser: { // ȭ�鿡 �����ִ� �÷� ����
            enabled: true,
            mode: "select" // or "dragAndDrop"
        },
        allowColumnReordering: true, // �÷� ���� ����
        "export": { // ���� �ٿ�ε� ����
            enabled: true,
            fileName: "EmailCustomerList",
            allowExportSelectedData: true
        },
        headerFilter: { // �÷��� ���� �˻� 
            visible: true
        },
        columnAutoWidth: true,
        columns: [
            {dataField : "�Ϸù�ȣ",alignment : "center",dataType: "number",fixed: true},
            {dataField : "���ڵ�",alignment : "center",dataType: "string"},
            {dataField : "��������",alignment : "center",dataType: "number",format: "fixedPoint"},
            {dataField : "�԰����",alignment : "center",dataType: "number",format: "fixedPoint"},
            {dataField : "�����",alignment : "center",dataType: "date"},
            {dataField : "������",alignment : "center",dataType: "date"},
            {dataField : "���ڵ�",alignment : "center",dataType: "string"},
            {dataField : "�����ڵ�",alignment : "center",dataType: "string"},
            {dataField : "��뿩��",alignment : "center",dataType: "string"},
            {dataField : "����",alignment : "center",dataType: "string"},
        ],
        dataSource: [
        <% For i = 0 To oAGV.FResultCount-1 %>
            {"�Ϸù�ȣ":<%=oAGV.FItemList(i).FIdx%>,
            "���ڵ�":"<%=oAGV.FItemList(i).FItemGubun &"-"&Num2Str(oAGV.FItemList(i).FItemid,6,"0","R")&"-"&oAGV.FItemList(i).FItemOption%>",
            "��������":<%=oAGV.FItemList(i).FRealStock%>,
            "�԰����":<%=oAGV.FItemList(i).FfixedStock%>,
            "�����":"<%=oAGV.FItemList(i).FRegdate%>",
            "������":"<%=oAGV.FItemList(i).Flastupdate%>",
            "���ڵ�":"<%=oAGV.FItemList(i).FRackCode%>",
            "�����ڵ�":"<%=oAGV.FItemList(i).FShelfCode%>",
            "��뿩��":"<%=oAGV.FItemList(i).getIsUsing%>",
            "����":"<%=oAGV.FItemList(i).getStatus%>"},
        <% Next %>
        ]
    });
});
</script>
<% Set oAGV = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->