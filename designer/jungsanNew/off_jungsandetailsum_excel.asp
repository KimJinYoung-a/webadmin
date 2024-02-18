<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesignerNoCache.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%
dim idx, gubuncd, shopid, makerid, groupid
idx     = requestCheckvar(request("idx"),10)
gubuncd = requestCheckvar(request("gubuncd"),32)
shopid  = requestCheckvar(request("shopid"),32)

makerid = session("ssBctId")
groupid = getPartnerId2GroupID(makerid)

if (NOT chkAvailViewJungsanOF(idx,makerid,groupid)) then
    response.write "��ȸ ������ �����ϴ�"
    dbget.close()	:	response.End
end if

dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FRectIdx = idx
'ooffjungsan.FRectMakerid = makerid
'if (makerid<>"") then
'    ooffjungsan.GetOneOffJungsanMaster
'end if

'if (groupid<>"") then
'    ooffjungsan.FRectGroupid = groupid
'end if

ooffjungsan.FRectGroupid = groupid

ooffjungsan.GetOneOffJungsanMaster

if (ooffjungsan.FResultCount<1) then
    response.write "<script >alert('�˻� ����� �����ϴ�.');</script>"
    dbget.close():response.End
end if

Dim IsCommissionTax : IsCommissionTax = ooffjungsan.FOneItem.IsCommissionTax

dim ooffjungsandetail
set ooffjungsandetail = new COffJungsan
ooffjungsandetail.FPageSize   = 1000
ooffjungsandetail.FRectIDX = idx
ooffjungsandetail.FRectMakerid = ooffjungsan.FOneItem.FMakerid
ooffjungsandetail.GetOffJungsanDetailSummaryList

''��ǰ���հ�
''IsShowDtlSum = TRUE or (groupid="G04798")

dim ooffjungsandetaillist
set ooffjungsandetaillist = new COffJungsan
ooffjungsandetaillist.FPageSize  = 3000
ooffjungsandetaillist.FRectIDX = idx
ooffjungsandetaillist.FRectGubunCD = gubuncd
ooffjungsandetaillist.FRectShopid  = shopid

ooffjungsandetaillist.GetOffJungsanDetailList

dim i
dim ttlitemno, ttlorgsellprice, ttlrealsellprice, ttlsuplyprice, ttlcommission
ttlitemno       = 0
ttlorgsellprice = 0
ttlrealsellprice= 0
ttlsuplyprice   = 0
ttlcommission   = 0

dim subitemno, subtotal
subitemno       = 0
subtotal        = 0

dim orgsellmargin, realsellmargin, selecteddefaultmargin
orgsellmargin   = 0
realsellmargin  = 0

dim codestr, shopname
if ooffjungsandetail.FResultCount>0 then
    for i=0 to ooffjungsandetail.FResultCount - 1
	    if (shopid=ooffjungsandetail.FItemList(i).Fshopid) and (gubuncd=ooffjungsandetail.FItemList(i).Fgubuncd) then
			shopname = ooffjungsandetail.FItemList(i).Fshopname
			codestr = ooffjungsandetail.FItemList(i).Fcomm_name
	    end if
    next
end if

%>

<!-- �������Ϸ� ���� ��� �κ� -->
<%
Response.ContentType = "application/unknown"
Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=" & ooffjungsan.FOneItem.FTitle & " " & shopname & " " & codestr & ".xls"


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
/* ���� �ٿ�ε�� ����� ���ڷ� ǥ�õ� ��� ���� */
.txt {mso-number-format:'\@'}
</style>
</head>
<body>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="100">���걸��</td>
      <td width="100">���� ��</td>
      <td width="100">���걸��</td>
      <td width="80">����<br>����</td>
      <td width="80">����</td>
      <% if (IsCommissionTax) then %>
        <td width="90">�����Ѿ�</td>
	    <td width="80">�⺻�Ǹ�<br>������</td>
        <td width="50">&nbsp;</td>
        <td width="80">�������ξ�<br>(�ٹ����ٺδ�)</td>
        <td width="80">�����ֹ���<br>(���»�����)</td>
		<td width="100">������</td>
		<td width="100">���޴���<br>(����Ȯ����)</td>
      <% else %>
        <td width="150">�ǸŰ��Ѿ�</td>
	    <td width="150">���ް��Ѿ�</td>
	    <td width="100">���޸�����</td>
      <% end if %>
    </tr>
    <% if ooffjungsandetail.FResultCount>0 then %>
    <% for i=0 to ooffjungsandetail.FResultCount - 1 %>

    <% if (shopid=ooffjungsandetail.FItemList(i).Fshopid) and (gubuncd=ooffjungsandetail.FItemList(i).Fgubuncd) then %>
    <tr align="center" bgcolor="#FFFFFF">
      <td><%= ooffjungsandetail.FItemList(i).getJSummaryGugunName %></td>
      <td><%= ooffjungsandetail.FItemList(i).Fshopname %></td>
      <td><%= ooffjungsandetail.FItemList(i).Fcomm_name %></td>
      <td><%= ooffjungsandetail.FItemList(i).GetItemVatTypeName %></td>
      <td><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_itemno,0) %></td>
      <% if (IsCommissionTax) then %>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice-ooffjungsandetail.FItemList(i).Ftot_realsellprice+ooffjungsandetail.FItemList(i).Ftot_commission,0) %></td>
      <td align="center">
            <% if (ooffjungsandetail.FItemList(i).Ftot_orgsellprice<>0) then %>
            <%= CLNG((ooffjungsandetail.FItemList(i).Ftot_orgsellprice-ooffjungsandetail.FItemList(i).Ftot_realsellprice+ooffjungsandetail.FItemList(i).Ftot_commission)/ooffjungsandetail.FItemList(i).Ftot_orgsellprice*100*100)/100 %> %
            <% end if %>
      </td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice-ooffjungsandetail.FItemList(i).Ftot_realsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_realsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_commission,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_jungsanprice,0) %></td>
      <% else %>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_jungsanprice,0) %></td>
      <td align="center">
        <% if ooffjungsandetail.FItemList(i).Ftot_orgsellprice<>0 then %>
        <%= CLng((1-ooffjungsandetail.FItemList(i).Ftot_jungsanprice/ooffjungsandetail.FItemList(i).Ftot_orgsellprice)*10000)/100 %> %
        <% end if %>
      </td>
      <% end if %>
    </tr>
   <% end if %>
   <% next %>
   <% end if %>
</table>
<br>


<%
subitemno = 0
subtotal  = 0
%>
<br>
<% if ooffjungsandetaillist.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="70">�����ڵ�</td>
      <td width="70">��ǰ�ڵ�</td>
      <td width="100">��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <% if (IsCommissionTax) then %>
      <td width="60">�����Ѿ�</td>
      <td width="60">�⺻�Ǹ�<br>������</td>
      <td width="50">&nbsp;</td>
      <td width="70">�������ξ�<br>(�ٹ����ٺδ�)</td>
      <td width="80">�����ֹ���<br>(���»�����)</td>
      <td width="60">������</td>
      <td width="60">�����</td>
      <td width="65">�����հ�</td>
      <% else %>
      <td width="50">�ǸŰ�</td>
      <td width="50">���ް�</td>
      <td width="60">���޸�����</td>
      <td width="65">���ް���</td>
      <% end if %>

    </tr>
    <% for i=0 to ooffjungsandetaillist.FResultCount-1 %>
    <%
        subitemno   = subitemno + ooffjungsandetaillist.FItemList(i).FItemNo
        subtotal    = subtotal + ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo

    %>
    <tr  bgcolor="#FFFFFF">
      <td class="txt"><%= ooffjungsandetaillist.FItemList(i).Forderno %></td>
      <td class="txt"><%= ooffjungsandetaillist.FItemList(i).GetBarCode %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).FItemName %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).FItemNo,0) %></td>
      <% if (IsCommissionTax) then %>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Frealsellprice+ooffjungsandetaillist.FItemList(i).Fcommission,0) %></td>
      <td align="center">
            <% if (ooffjungsandetaillist.FItemList(i).Forgsellprice<>0) then %>
            <%= CLNG((ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Frealsellprice+ooffjungsandetaillist.FItemList(i).Fcommission)/ooffjungsandetaillist.FItemList(i).Forgsellprice*100*100)/100 %> %
            <% end if %>
      </td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Frealsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Frealsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fcommission,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice,0) %></td>
      <td align="right">
          <% if ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo<1 then %>
          <font color="red"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %></font>
          <% else %>
          <%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %>
          <% end if %>
      </td>
      <% else %>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice,0) %></td>
      <td align="center">
      <% if ooffjungsandetaillist.FItemList(i).Forgsellprice<>0 then %>
      <%= 100-CLNG((ooffjungsandetaillist.FItemList(i).Fsuplyprice)/ooffjungsandetaillist.FItemList(i).Forgsellprice*100) %> %
      <% end if %>
      </td>
      <td align="right">
          <% if ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo<1 then %>
          <font color="red"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %></font>
          <% else %>
          <%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %>
          <% end if %>
      </td>
      <% end if %>
    </tr>
   <% next %>
    <tr bgcolor="#FFFFFF">
        <td align="center">�հ�</td>
        <td colspan="3"></td>
        <td align="center">
            <% if (ooffjungsan.FOneItem.Ftot_itemno<>subitemno) then %>
            <b><%= FormatNumber(subitemno,0) %></b>
            <% else %>
            <%= FormatNumber(subitemno,0) %>
            <% end if %>
        </td>
        <td colspan="<%=CHKIIF(IsCommissionTax,7,3)%>"></td>
        <td align="right">
            <% if (ooffjungsan.FOneItem.Ftot_jungsanprice<>subtotal) then %>
            <b><%= FormatNumber(subtotal,0) %></b>
            <% else %>
            <%= FormatNumber(subtotal,0) %>
            <% end if %>
        </td>
    </tr>
</table>
<% end if %>
<%
set ooffjungsan = Nothing
set ooffjungsandetail = Nothing
set ooffjungsandetaillist = Nothing
%>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

