<%@  language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : ���Ի�ǰ��������
' History : 2022.01.17 �̻� ����
'           2022.08.30 �ѿ�� ����(�ֹ���ǰ������ �������� ������� Ʋ����� �Ѹ��� ����)
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/tenmember/lib/header.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim idx, codeList, i, sumTotalPrice, sumOrderNo, yyyymmArray, yyyymmArrayCount
dim makerid, purchaseNm, formidx, tContents
    idx = requestCheckVar(getNumeric(request("idx")),10)
	codeList	=  requestCheckvar(Request("codeList"),4000)

sumTotalPrice = 0
sumOrderNo = 0

Call GetPurchaseTypeList(codeList, makerid, purchaseNm, formidx)

purchaseNm = "<span style='color:red;'>" & purchaseNm & "</span>"

dim clseapp
set clseapp = new CEApproval
	clseapp.Fedmsidx = formidx
	clseapp.fnGetEAppForm
	tContents		= clseapp.FedmsForm
set clseapp = nothing

'// ��ǰ����
dim oCPurchasedProductItem
set oCPurchasedProductItem = new CPurchasedProduct
	oCPurchasedProductItem.FRectIdx = CHKIIF(idx="", "-1", idx)
	oCPurchasedProductItem.FPageSize = 500
	oCPurchasedProductItem.GetPurchasedProductItemList

'// ��������
dim oCPurchasedProductSheet
set oCPurchasedProductSheet = new CPurchasedProduct
	oCPurchasedProductSheet.FRectMasterIdx = CHKIIF(idx="", "-1", idx)
	oCPurchasedProductSheet.FPageSize = 500
	oCPurchasedProductSheet.FRectExcDel = "Y"
	oCPurchasedProductSheet.GetPurchasedProductSheetMasterList

if oCPurchasedProductSheet.FResultCount>0 then
    for i=0 to oCPurchasedProductSheet.FResultCount-1
        if instr(yyyymmArray,oCPurchasedProductSheet.FItemList(i).Fyyyymm)<1 then
            yyyymmArray = yyyymmArray & oCPurchasedProductSheet.FItemList(i).Fyyyymm & ","
        end if
    next
    if right(yyyymmArray,1)="," then yyyymmArray = left(yyyymmArray,len(yyyymmArray)-1)
end if

%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
	<input type="hidden" name="tC" value="">
	<input type="hidden" name="ieidx" value="<%=formidx%>">
	<input type="hidden" name="iSL" value="<%=idx%>">
</form>
<div id="divEapp" style="display:none;">
	<p style="padding-bottom:30px;">������ ���� <%=purchaseNm%>�� �����ϰ��� �Ͽ��� ���� �� �簡 �ٶ��ϴ�.</p>
	<p style="padding-bottom:30px;text-align:center;">- �� �� -</p>
	<p style="padding-bottom:10px;"><strong>�� ���� </strong>: <%= makerid %>&nbsp;<%=purchaseNm%></p>
	<p><strong>�� �ֹ����� </strong></p>
	<p style="padding-bottom:10px;">
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
            <td width="60">�԰���</td>
            <td>�귣��ID</td>
		    <td width="120">��ǰ�ڵ�</td>
		    <td>��ǰ��</td>
		    <td>�ɼǸ�</td>
            <td width="70">�Һ��ڰ�</td>
            <td width="70">����</td>
            <td width="70">�����Ѿ�<br />(VAT����)</td>
            <td width="50">����</td>
			<td>���</td>
		</tr>
        <%
		yyyymmArrayCount=0
        for i=0 to oCPurchasedProductItem.FResultCount-1

		if instr(yyyymmArray,oCPurchasedProductItem.FItemList(i).Fyyyymm)>0 or yyyymmArray="" then
            sumTotalPrice = sumTotalPrice + oCPurchasedProductItem.FItemList(i).FtotalPrice
            sumOrderNo = sumOrderNo + oCPurchasedProductItem.FItemList(i).ForderNo
    		yyyymmArrayCount = yyyymmArrayCount + 1
        %>
		<tr bgcolor="#FFFFFF"  align="center">
            <td align="center">
                <%= oCPurchasedProductItem.FItemList(i).Fyyyymm %>
            </td>
            <td align="center"><%= oCPurchasedProductItem.FItemList(i).Fmakerid %></td>
		    <td align="center">
			    <%= oCPurchasedProductItem.FItemList(i).FItemGubun %>-<%= BF_GetFormattedItemId(oCPurchasedProductItem.FItemList(i).FItemID) %>-<%= oCPurchasedProductItem.FItemList(i).Fitemoption %>
		    </td>
		    <td><%= oCPurchasedProductItem.FItemList(i).Fitemname %></td>
		    <td><%= oCPurchasedProductItem.FItemList(i).Fitemoptionname %></td>
            <td align="right"><%= FormatNumber(oCPurchasedProductItem.FItemList(i).Forgprice, 0) %></td>
            <td align="right"><%= FormatNumber(oCPurchasedProductItem.FItemList(i).Fcogs, 0) %></td>
            <td align="right"><%= FormatNumber(oCPurchasedProductItem.FItemList(i).FtotalPrice, 0) %></td>
            <td align="right"><%= FormatNumber(oCPurchasedProductItem.FItemList(i).ForderNo, 0) %></td>
            <td></td>
		</tr>
		<% end if %>
        <% next %>
		<tr bgcolor="#FFFFFF"  align="center">
            <td colspan="7">

            </td>
            <td align="right"><%= FormatNumber(sumTotalPrice, 0) %></td>
            <td align="right"><%= FormatNumber(sumOrderNo, 0) %></td>
            <td></td>
		</tr>
	</table>
	</p>

    <p style="padding-bottom:10px;"></p>
	<p><strong>�� �������� </strong></p>
	<p style="padding-bottom:10px;">
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
            <td width="150">��ü��</td>
            <td width="120">��뱸��</td>
            <td width="80">���԰�</td>
			<td>��������</td>
		</tr>
        <%
        sumTotalPrice = 0
        for i=0 to oCPurchasedProductSheet.FResultCount-1
            sumTotalPrice = sumTotalPrice + oCPurchasedProductSheet.FItemList(i).FbuyPrice
        %>
		<tr bgcolor="#FFFFFF"  align="center">
            <td align="center">
                <%= oCPurchasedProductSheet.FItemList(i).Fcompany_name %>
            </td>
            <td align="center"><%= oCPurchasedProductSheet.FItemList(i).FppGubunName %></td>
		    <td align="right">
			    <%= FormatNumber(oCPurchasedProductSheet.FItemList(i).FbuyPrice, 0) %>
		    </td>
            <td></td>
		</tr>
        <% next %>
		<tr bgcolor="#FFFFFF"  align="center">
            <td colspan="2">

            </td>
            <td align="right"><%= FormatNumber(sumTotalPrice, 0) %></td>
            <td></td>
		</tr>
	</table>
	</p>

	<%=tContents%>
	</div>


	<script type="text/javascript">
	document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");
	document.frmEapp.submit();
	</script>
