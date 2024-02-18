<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/offjungsancls.asp"-->

<%
dim idx, gubuncd, shopid
idx = request("idx")
gubuncd = request("gubuncd")
shopid  = request("shopid")

dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FRectIdx = idx
ooffjungsan.FRectMakerid = session("ssBctId")
ooffjungsan.GetOneOffJungsanMaster

if (ooffjungsan.FResultCount<1) then
    response.write "<script >alert('�˻� ����� �����ϴ�.');</script>"
    dbget.close()	:	response.End
end if

dim ooffjungsandetail
set ooffjungsandetail = new COffJungsan
ooffjungsandetail.FPageSize   = 1000
ooffjungsandetail.FRectIDX = idx
ooffjungsandetail.FRectMakerid = ooffjungsan.FOneItem.FMakerid
ooffjungsandetail.GetOffJungsanDetailSummaryList


dim ooffjungsandetaillist
set ooffjungsandetaillist = new COffJungsan
ooffjungsandetaillist.FPageSize  = 3000
ooffjungsandetaillist.FRectIDX = idx
ooffjungsandetaillist.FRectGubunCD = gubuncd
ooffjungsandetaillist.FRectShopid  = shopid

ooffjungsandetaillist.GetOffJungsanDetailList

dim i
dim ttlitemno, ttlorgsellprice, ttlrealsellprice, ttlsuplyprice
ttlitemno       = 0
ttlorgsellprice = 0
ttlrealsellprice= 0
ttlsuplyprice   = 0

dim subitemno, subtotal
subitemno       = 0
subtotal        = 0

dim orgsellmargin, realsellmargin, selecteddefaultmargin
orgsellmargin   = 0
realsellmargin  = 0




dim codestr, shopname
if ooffjungsandetail.FResultCount>0 then
    for i=0 to ooffjungsandetail.FResultCount - 1
	    ttlitemno           = ttlitemno + ooffjungsandetail.FItemList(i).Ftot_itemno
	    ttlorgsellprice     = ttlorgsellprice + ooffjungsandetail.FItemList(i).Ftot_orgsellprice
	    ttlrealsellprice    = ttlrealsellprice + ooffjungsandetail.FItemList(i).Ftot_realsellprice
	    ttlsuplyprice       = ttlsuplyprice + ooffjungsandetail.FItemList(i).Ftot_jungsanprice

	    if ooffjungsandetail.FItemList(i).Ftot_orgsellprice<>0 then
	        orgsellmargin = CLng((ooffjungsandetail.FItemList(i).Ftot_orgsellprice-ooffjungsandetail.FItemList(i).Ftot_jungsanprice)/ooffjungsandetail.FItemList(i).Ftot_orgsellprice*100*100)/100
	    else
	        orgsellmargin = 0
	    end if

	    if ooffjungsandetail.FItemList(i).Ftot_realsellprice<>0 then
	        realsellmargin = CLng((ooffjungsandetail.FItemList(i).Ftot_realsellprice-ooffjungsandetail.FItemList(i).Ftot_jungsanprice)/ooffjungsandetail.FItemList(i).Ftot_realsellprice*100*100)/100
	    else
	        realsellmargin = 0
	    end if


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
<!--ǥ ������-->
<!--
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="25" valign="top">
		<td colspan=8>
			<font color="red"><%= ooffjungsan.FOneItem.FTitle %>&nbsp;<%= ooffjungsan.FOneItem.Fmakerid %>&nbsp;&nbsp;
            <%= ooffjungsan.FOneItem.Fdifferencekey %> �� &nbsp;&nbsp;
            <font color="<%= ooffjungsan.FOneItem.GetTaxtypeNameColor %>"><%= ooffjungsan.FOneItem.GetSimpleTaxtypeName %></font> &nbsp;&nbsp;
            �� ����� : <%= FormatNumber(ooffjungsan.FOneItem.Ftot_jungsanprice,0) %>&nbsp;&nbsp;
            �� �ǸŻ�ǰ�� : <%= FormatNumber(ooffjungsan.FOneItem.Ftot_itemno,0) %></font>
		</td>
	</tr>
	<tr height="25" valign="top">
		<td>
		</td>
	</tr>
	<tr height="25" valign="top">
		<td>
		</td>
	</tr>
</table>
-->
<!--ǥ ��峡-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1">
    <tr align="center" >
      <td width="100">�������ڵ�</td>
      <td width="100" colspan=3>������ ��</td>
      <td width="100">���걸��</td>
      <td width="80">�ѻ�ǰ�Ǽ�</td>
      <td width="90">�����</td>
      <td width="90">�����</td>
    </tr>
    <% if ooffjungsandetail.FResultCount>0 then %>
    <% for i=0 to ooffjungsandetail.FResultCount - 1 %>
    <%
    ttlitemno           = ttlitemno + ooffjungsandetail.FItemList(i).Ftot_itemno
    ttlorgsellprice     = ttlorgsellprice + ooffjungsandetail.FItemList(i).Ftot_orgsellprice
    ttlrealsellprice    = ttlrealsellprice + ooffjungsandetail.FItemList(i).Ftot_realsellprice
    ttlsuplyprice       = ttlsuplyprice + ooffjungsandetail.FItemList(i).Ftot_jungsanprice

    if ooffjungsandetail.FItemList(i).Ftot_orgsellprice<>0 then
        orgsellmargin = CLng((ooffjungsandetail.FItemList(i).Ftot_orgsellprice-ooffjungsandetail.FItemList(i).Ftot_jungsanprice)/ooffjungsandetail.FItemList(i).Ftot_orgsellprice*100*100)/100
    else
        orgsellmargin = 0
    end if

    if ooffjungsandetail.FItemList(i).Ftot_realsellprice<>0 then
        realsellmargin = CLng((ooffjungsandetail.FItemList(i).Ftot_realsellprice-ooffjungsandetail.FItemList(i).Ftot_jungsanprice)/ooffjungsandetail.FItemList(i).Ftot_realsellprice*100*100)/100
    else
        realsellmargin = 0
    end if

    %>
    <% if (shopid=ooffjungsandetail.FItemList(i).Fshopid) and (gubuncd=ooffjungsandetail.FItemList(i).Fgubuncd) then %>
    <% selecteddefaultmargin = ooffjungsandetail.FItemList(i).Fdefaultmargin %>
    <tr align="center" bgcolor="#BBBBDD">
      <td><%= ooffjungsandetail.FItemList(i).Fshopid %></td>
      <td colspan=3><%= ooffjungsandetail.FItemList(i).Fshopname %></td>
      <td><%= ooffjungsandetail.FItemList(i).Fcomm_name %></td>
      <td><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_itemno,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_realsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_jungsanprice,0) %></td>
    </tr>
    <% end if %>
    <% next %>
    <% end if %>
	<tr height="25" valign="top">
		<td>
		</td>
	</tr>
	<tr height="25" valign="top">
		<td>
		</td>
	</tr>
</table>
<% if ooffjungsandetaillist.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" >
      <td width="70">�����ڵ�</td>
      <td width="70">��ǰ�ڵ�</td>
      <td width="100">��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="60">�ǸŰ�</td>
      <td width="60">���԰�</td>
      <td width="40">����</td>
      <td width="64">�����</td>
    </tr>
    <% for i=0 to ooffjungsandetaillist.FResultCount-1 %>
    <%
        subitemno   = subitemno + ooffjungsandetaillist.FItemList(i).FItemNo
        subtotal    = subtotal + ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo

        if ooffjungsandetaillist.FItemList(i).Forgsellprice<>0 then
            orgsellmargin = CLng((ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Fsuplyprice)/ooffjungsandetaillist.FItemList(i).Forgsellprice*100*100)/100
        else
            orgsellmargin = 0
        end if

        if ooffjungsandetaillist.FItemList(i).Frealsellprice<>0 then
            realsellmargin = CLng((ooffjungsandetaillist.FItemList(i).Frealsellprice-ooffjungsandetaillist.FItemList(i).Fsuplyprice)/ooffjungsandetaillist.FItemList(i).Frealsellprice*100*100)/100
        else
            realsellmargin = 0
        end if
    %>
    <tr  bgcolor="#FFFFFF">
      <td class="txt"><%= ooffjungsandetaillist.FItemList(i).Forderno %></td>
      <td class="txt"><%= ooffjungsandetaillist.FItemList(i).GetBarCode %></td>
      <td class="txt"><%= ooffjungsandetaillist.FItemList(i).FItemName %></td>
      <td class="txt"><%= ooffjungsandetaillist.FItemList(i).FItemOptionName %></td>
      <td align="right">
        <%= FormatNumber(ooffjungsandetaillist.FItemList(i).Frealsellprice,0) %>
      </td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice,0) %></td>
      <td align="center"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).FItemNo,0) %></td>
      <td align="right">
      <% if ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo<1 then %>
      <font color="red"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %></font>
      <% else %>
      <%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %>
      <% end if %>
      </td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
        <td align="center">�հ�</td>
        <td colspan="5"></td>
        <td align="center">
            <% if (ooffjungsan.FOneItem.Ftot_itemno<>subitemno) then %>
            <b><%= FormatNumber(subitemno,0) %></b>
            <% else %>
            <%= FormatNumber(subitemno,0) %>
            <% end if %>
        </td>
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

