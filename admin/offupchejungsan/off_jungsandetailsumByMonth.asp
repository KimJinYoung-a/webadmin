<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%

dim yyyy1, mm1, gubuncd, shopid, makerid
dim dSc, commcd

gubuncd = request("gubuncd")
shopid  = request("shopid")
makerid = request("makerid")
yyyy1   = request("yyyy1")
mm1   = request("mm1")
dSc   = request("dSc")
commcd= request("commcd")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim ooffjungsandetail
set ooffjungsandetail = new COffJungsan
ooffjungsandetail.FPageSize   = 1000
ooffjungsandetail.FRectYYYYMM  = yyyy1+"-"+mm1
ooffjungsandetail.FRectMakerid = makerid
ooffjungsandetail.FRectShopid  = shopid
ooffjungsandetail.FRectGubunCD = commcd
ooffjungsandetail.GetOffJungsanDetailSummaryListByMonth


dim ooffjungsandetaillist
set ooffjungsandetaillist = new COffJungsan
ooffjungsandetaillist.FPageSize  = 3000
ooffjungsandetaillist.FRectYYYYMM  = yyyy1+"-"+mm1
ooffjungsandetaillist.FRectGubunCD = gubuncd
ooffjungsandetaillist.FRectShopid  = shopid
ooffjungsandetaillist.FRectMakerid = makerid

if (shopid<>"") and (gubuncd<>"") and (makerid<>"") and (ooffjungsandetaillist.FRectYYYYMM<>"") and (dSc<>"") then
    ooffjungsandetaillist.GetOffJungsanDetailListByMonth
end if

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

%>
<script language='javascript'>
function PopDetailList(shopid,gubuncd){
    location.href = '?shopid=' + shopid + '&gubuncd=' + gubuncd +'&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&makerid=<%=makerid%>&commcd=<%=commcd%>&dSc=on';
}


</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1"  class="a" bgcolor="#999999">
    <form name="frm" method="get" action="">
    
    <tr align="center" bgcolor="#FFFFFF" >
        <td rowspan="2" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
        <td align="left">
            �������� : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;
			�귣��ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>&nbsp;&nbsp;
            ���� <% drawSelectBoxOffShopAll "shopid",shopid %>
        </td>
        <td rowspan="2" width="50" bgcolor="#EEEEEE">
            <a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
    </tr>
    <tr height="25" bgcolor="#FFFFFF">
        <td valign="top" width="530">
            ���걸�� : <% drawSelectBoxJungsanCommCombo "commcd",commcd,"Z002" %>
        </td>
    </tr>
    </form>
</table>
<!-- ǥ ��ܹ� ��-->
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="100">�������ڵ�</td>
      <td width="100">������ ��</td>
      <td width="100">�⺻���걸��</td>
      <td width="100">���걸��</td>
      <td width="80">�ѻ�ǰ�Ǽ�</td>
      <td width="90">�Һ���</td>
      <td width="90">�����</td>
      <td width="90">�����</td>
      <td width="50">�Һ�<br>����</td>
      <td width="50">�����<br>����</td>
      <td width="40">��<br>����</td>
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
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    <% end if %>
      <td><%= ooffjungsandetail.FItemList(i).Fshopid %></td>
      <td><%= ooffjungsandetail.FItemList(i).Fshopname %></td>
      <td>
        <%= ooffjungsandetail.FItemList(i).GetChargeDivName & " " & ooffjungsandetail.FItemList(i).Fdefaultmargin %>
      </td>
      <td><%= ooffjungsandetail.FItemList(i).Fcomm_name %></td>
      <td><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_itemno,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_realsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_jungsanprice,0) %></td>
      <td align="center">
      <% if ooffjungsandetail.FItemList(i).Fdefaultmargin<>orgsellmargin then %>
        <font color="red"><%= orgsellmargin %></font> %
      <% else %>
        <%= orgsellmargin %> %
      <% end if %>
      </td>
      <td align="center">
      <% if ooffjungsandetail.FItemList(i).Fdefaultmargin<>realsellmargin then %>
        <font color="blue"><%= realsellmargin %></font> %
      <% else %>
        <%= realsellmargin %> %
      <% end if %>
        
      </td>
      <td><a href="javascript:PopDetailList('<%= ooffjungsandetail.FItemList(i).Fshopid %>','<%= ooffjungsandetail.FItemList(i).FGubuncd %>')"><img src="/images/icon_search.jpg" width="16" border="0"></a></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td align="center">�հ�</td>
      <td colspan="3"></td>
      <td align="center"><%= FormatNumber(ttlitemno,0) %></td>
      <td align="right"><%= FormatNumber(ttlorgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ttlrealsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ttlsuplyprice,0) %></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <% else %>
    <tr bgcolor="#FFFFFF">
      <td colspan="12" align="center">[�˻� ����� �����ϴ�.]</td>
    </tr>
    <% end if %>
</table>
<br>

<% if ooffjungsandetaillist.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="70">�����ڵ�</td>
      <td width="70">��ǰ�ڵ�</td>
      <td width="100">��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="60">�Һ�</td>
      <td width="60">�ǸŰ�</td>
      <td width="60">���԰�</td>
      <td width="40">�Һ�<br>����</td>
      <td width="40">����<br>����</td>
      <td width="40">����</td>
      <td width="36">����<br>����</td>
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
      <td><%= ooffjungsandetaillist.FItemList(i).Forderno %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).GetBarCode %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).FItemName %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice,0) %></td>
      <td align="right">
        <%= FormatNumber(ooffjungsandetaillist.FItemList(i).Frealsellprice,0) %>
        <% if (ooffjungsandetaillist.FItemList(i).Frealsellprice<>ooffjungsandetaillist.FItemList(i).Forgsellprice) then %>
            <% if ooffjungsandetaillist.FItemList(i).Forgsellprice<>0 then %>
                <br><font color="red"><%= Clng((ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Frealsellprice)/ooffjungsandetaillist.FItemList(i).Forgsellprice*100*100)/100 %></font> %
            <% end if %> 
        <% end if %>
      </td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice,0) %></td>
      <td align="center">
      <% if selecteddefaultmargin<>orgsellmargin then %> 
        <font color="red"><%= orgsellmargin %></font> %
      <% else %>
        <%= orgsellmargin %> %
      <% end if %>
      </td>
      <td align="center"> 
      <% if orgsellmargin<>realsellmargin then %> 
        <font color="blue"><%= realsellmargin %></font> %
      <% else %>
        <%= realsellmargin %> %
      <% end if %>
      </td>
      <td align="center"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).FItemNo,0) %></td>
      <td align="center">
      <% if ooffjungsandetaillist.FItemList(i).Fcentermwdiv="M" then %>
      <b><%= ooffjungsandetaillist.FItemList(i).Fcentermwdiv %></b>
      <% else %>
      <%= ooffjungsandetaillist.FItemList(i).Fcentermwdiv %>
      <% end if %>
      </td>
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
        <td colspan="8"></td>
        <td align="center">
            <%= FormatNumber(subitemno,0) %>
        </td>
        <td></td>
        <td align="right">
            <%= FormatNumber(subtotal,0) %>
        </td>
    </tr>
</table>    

<% end if %>
<%
set ooffjungsandetail = Nothing
set ooffjungsandetaillist = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->