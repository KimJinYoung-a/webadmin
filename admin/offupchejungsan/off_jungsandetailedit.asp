<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->

<%
dim idx, gubuncd, shopid

idx     = requestCheckvar(request("idx"),10)
gubuncd = requestCheckvar(request("gubuncd"),16)
shopid = requestCheckvar(request("shopid"),32)

dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FRectIdx = idx
''ooffjungsan.FRectMakerid = ��ü�ϰ�� session ��ü���̵�
ooffjungsan.GetOneOffJungsanMaster

if (ooffjungsan.FResultCount<1) then
    response.write "<script >alert('�˻� ����� �����ϴ�.');</script>"
    dbget.close()	:	response.End
end if

dim ooffjungsandetail
set ooffjungsandetail = new COffJungsan
ooffjungsandetail.FRectIdx        = idx
ooffjungsandetail.FRectGubunCD    = gubuncd
ooffjungsandetail.FRectShopid     = shopid
ooffjungsandetail.FRectMakerid = ooffjungsan.FOneItem.FMakerid
ooffjungsandetail.GetOffJungsanDetailSummaryList

dim ooffjungsandetaillist
set ooffjungsandetaillist = new COffJungsan
ooffjungsandetaillist.FPageSize       = 3000
ooffjungsandetaillist.FRectIdx        = idx
ooffjungsandetaillist.FRectGubunCD    = gubuncd
ooffjungsandetaillist.FRectShopid     = shopid
ooffjungsandetaillist.GetOffJungsanDetailList


dim ochargeuser
set ochargeuser = new COffShopChargeUser
ochargeuser.FRectShopID = shopid
ochargeuser.FRectDesigner = ooffjungsan.FOneItem.Fmakerid
ochargeuser.GetOffShopDesignerList

dim defaultmargine
defaultmargine = ochargeuser.FItemList(0).FDefaultMargin




dim ttlitemno, ttlorgsellprice, ttlrealsellprice, ttlsuplyprice, ttlcommission
ttlitemno   = 0
ttlorgsellprice     = 0
ttlrealsellprice    = 0
ttlsuplyprice       = 0
ttlcommission       = 0

dim subitemno, subtotal
subitemno       = 0
subtotal        = 0

dim orgsellmargin, realsellmargin, selecteddefaultmargin
orgsellmargin   = 0
realsellmargin  = 0

dim i, orderdate, Is20ProSale
dim IsEditEnabled
IsEditEnabled = false

IsEditEnabled = ooffjungsan.FOneItem.IsEditenable


''IsEditEnabled = IsEditEnabled and (shopid<>"") and (gubuncd<>"") and (idx<>"")
IsEditEnabled = IsEditEnabled and (gubuncd<>"") and (idx<>"")
%>
<script language='javascript'>
<% if Not IsEditEnabled then %>
    alert('���� ���� ���°� �ƴմϴ�.(������ ���¸� ���� ����)');
<% end if %>

function PopDetailEdit(idx,gubuncd,shopid){
    location.href = 'off_jungsandetailedit.asp?idx=' + idx + '&gubuncd=' + gubuncd + '&shopid=' + shopid;
}

function delDetail(iidx){
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frmarr.mode.value = "deldetail";
		frmarr.detailidx.value = iidx;

		frmarr.submit();
	}

}

function AddEtc(iidx){
	var popwin = window.open("pop_off_adddetailetc.asp?shopid=<%= shopid %>&makerid=<%= ooffjungsan.FOneItem.Fmakerid %>&gubuncd=<%= gubuncd %>&masteridx=" + iidx,"pop_off_adddetailetc","width=640,height=340,scrollbars=yes,resizble=yes");
	popwin.focus();
}

function SaveArr(){
	var frm = document.frmList;
	var idxarr = "";
	var suplypricearr  = "";
	var itemnoarr = "";

	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if (e.name=="idx"){
			idxarr = idxarr + e.value + "|";
		}

		if (e.name=="buycash"){
			if (!IsInteger(e.value)){
				alert('������ ������ �����մϴ�.');
				e.focus();
				return;
			}
			suplypricearr = suplypricearr + e.value + "|";
		}

		if (e.name=="itemno"){
			if (!IsInteger(e.value)){
				alert('������ ������ �����մϴ�.');
				e.focus();
				return;
			}
			itemnoarr = itemnoarr + e.value + "|";
		}

	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frmarr.idxarr.value = idxarr;
		frmarr.suplyprice.value = suplypricearr;
		frmarr.itemno.value = itemnoarr;

		frmarr.submit();
	}
}

function ReMargin(){
	var frm = document.frm;
	var defaultmargine = <%= defaultmargine %>;
	var frm_buycash;

	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if (e.name=="idx"){
			i_sellprice = frm.elements[i+1].value;
			i_realsellprice = frm.elements[i+2].value;
			i_suplyprice = frm.elements[i+3].value;
			i_orderno = frm.elements[i+4].value;
			frm_buycash = frm.elements[i+5];

			if (i_orderno=="true"){
				frm_buycash.value = parseInt(i_sellprice*(1-defaultmargine/100)-(i_sellprice-i_realsellprice)/2);
			}
		}
	}

}

function calcu20(i){
    var frm = document.frmList;
    var orgsellprice  = frm.i_sellprice.value;
    var realsellprice = frm.i_realsellprice[i].value;
    var buycash       = frm.buycash[i].value;
}
function reMargin(i,reValue){
    var frm = document.frmList;
    frm.buycash[i].value = reValue;
}
</script>



<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <form name="frm" method="get" action="">
    <input type="hidden" name="idx" value="<%= idx %>">
    <tr height="10" valign="bottom" bgcolor="F4F4F4">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="25" valign="bottom" bgcolor="F4F4F4">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top" bgcolor="F4F4F4" width="530">
            <%= ooffjungsan.FOneItem.FTitle %>&nbsp;<%= ooffjungsan.FOneItem.Fmakerid %>&nbsp;&nbsp;
            <%= ooffjungsan.FOneItem.Fdifferencekey %> �� &nbsp;&nbsp;
            <font color="<%= ooffjungsan.FOneItem.GetTaxtypeNameColor %>"><%= ooffjungsan.FOneItem.GetSimpleTaxtypeName %></font> &nbsp;&nbsp;
            �� ����� : <%= FormatNumber(ooffjungsan.FOneItem.Ftot_jungsanprice,0) %>&nbsp;&nbsp;
            �� �ǸŻ�ǰ�� : <%= FormatNumber(ooffjungsan.FOneItem.Ftot_itemno,0) %>
            <% if (ooffjungsan.FOneItem.IsCommissionTax) then %>
            �� ������ : <%= FormatNumber(ooffjungsan.FOneItem.Ftotalcommission,0) %>
            <% end if %>
            <br><br>
            ���� �⺻ ���� : <b><%= ochargeuser.FItemList(0).FDefaultMargin %><b> %

            <br>
            <textarea cols="100" rows="3"><%= ochargeuser.FItemList(0).FEtcjunsandetail %></textarea>
        </td>
        <td valign="top" bgcolor="F4F4F4" align="right">
        &nbsp;
        <!--
            <a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        -->
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ ��ܹ� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="100">�������ڵ�</td>
      <td width="100">������ ��</td>
      <td width="100">�⺻���걸��</td>
      <td width="100">���걸��</td>
      <td width="80">�ѻ�ǰ�Ǽ�</td>
      <td width="90">�ǸŰ���</td>
      <td width="90">�����</td>
      <td width="80">������</td>
      <td width="90">�����</td>
      <td width="50">�Һ�<br>����</td>
      <td width="50">�����<br>����</td>
      <td width="50">����</td>
    </tr>
    <% if ooffjungsandetail.FResultCount>0 then %>
    <% for i=0 to ooffjungsandetail.FResultCount - 1 %>
    <%
    ttlitemno           = ttlitemno + ooffjungsandetail.FItemList(i).Ftot_itemno
    ttlorgsellprice     = ttlorgsellprice + ooffjungsandetail.FItemList(i).Ftot_orgsellprice
    ttlrealsellprice    = ttlrealsellprice + ooffjungsandetail.FItemList(i).Ftot_realsellprice
    ttlsuplyprice       = ttlsuplyprice + ooffjungsandetail.FItemList(i).Ftot_jungsanprice
    ttlcommission       = ttlcommission + ooffjungsandetail.FItemList(i).Ftot_commission

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
        <%= ooffjungsandetail.FItemList(i).GetChargeDivName %>,
        <%= ooffjungsandetail.FItemList(i).Fdefaultmargin %>,
        <% if ooffjungsandetail.FItemList(i).Fautojungsan="N" then response.write "<font color='blue'>��</font>" else response.write "��" %>,
        <%= ooffjungsandetail.FItemList(i).Fautojungsandiv %>
      </td>
      <td><%= ooffjungsandetail.FItemList(i).Fcomm_name %></td>
      <td><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_itemno,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_realsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_commission,0) %></td>
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
      <td><a href="javascript:PopDetailEdit('<%= idx %>','<%= ooffjungsandetail.FItemList(i).FGubuncd %>','<%= ooffjungsandetail.FItemList(i).FShopid %>')"><img src="/images/icon_modify.gif" border="0" width="45"></a></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td align="center">�հ�</td>
      <td colspan="3"></td>
      <td align="center"><%= FormatNumber(ttlitemno,0) %></td>
      <td align="right"><%= FormatNumber(ttlorgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ttlrealsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ttlcommission,0) %></td>
      <td align="right"><%= FormatNumber(ttlsuplyprice,0) %></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <% else %>
    <tr bgcolor="#FFFFFF">
      <td colspan="11" align="center">[�˻� ����� �����ϴ�.]</td>
    </tr>
    <% end if %>
</table>

<br>
<table width="100%" border=0 cellspacing="1" class="a"  width=800 bgcolor=#FFFFFF>
<tr>
	<td><input type=button value="��Ÿ�����߰�" onclick="AddEtc('<%= idx %>')" <% if Not IsEditEnabled then response.write "disabled" %> ></td>
	<td align=right><input type=button value="��ü ����" onclick="SaveArr()" <% if Not IsEditEnabled then response.write "disabled" %> ></td>
</tr>
</table>
<br>
<table width="100%" border=0 cellspacing="1" class="a"  width=800 bgcolor="<%= adminColor("tablebg") %>">
<form name=frmList method=post >
<tr align="center"  bgcolor="<%= adminColor("tabletop") %>" >
	<td width="100">�����ڵ�</td>
	<td width="80">���ڵ�</td>
	<td width="100">��ǰ��</td>
	<td width="70">�ɼǸ�</td>
	<td width="60">�Һ��ڰ�</td>
	<td width="60">���ǸŰ�</td>
	<td width="40">����<br>��</td>
	<td width="40">�Һ�<br>����</td>
	<td width="40">����<br>����</td>
	<td width="60">������<br></td>
	<td width="60">�����<br></td>
	<td width="60">����</td>

	<td width="60">�հ�</td>
	<td width="40">����<br>����</td>
	<td width="30">����</td>
</tr>
<% for i=0 to ooffjungsandetaillist.FResultCount-1 %>
<input type=hidden name=idx value="<%= ooffjungsandetaillist.FItemList(i).Fdetailidx %>">
<input type=hidden name=i_sellprice value="<%= ooffjungsandetaillist.FItemList(i).FOrgSellPrice %>">
<input type=hidden name=i_realsellprice value="<%= ooffjungsandetaillist.FItemList(i).FRealSellPrice %>">
<input type=hidden name=i_suplyprice value="<%= ooffjungsandetaillist.FItemList(i).FSuplyPrice %>">

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

orderdate = "20" & Left(ooffjungsandetaillist.FItemList(i).Forderno,2) +"-" + mid(ooffjungsandetaillist.FItemList(i).Forderno,3,2) + "-" + mid(ooffjungsandetaillist.FItemList(i).Forderno,5,2)

Is20ProSale = ((orderdate>="2009-12-28") and (orderdate<"2009-12-31"))
%>
<tr bgcolor="#FFFFFF">

	<td><% if (Is20ProSale) then %><strong><a href="javascript:calcu20('<%= i %>')"><%= ooffjungsandetaillist.FItemList(i).Forderno %></a></strong><% else %><%= ooffjungsandetaillist.FItemList(i).Forderno %><% end if %></td>
	<td><%= ooffjungsandetaillist.FItemList(i).Fitemgubun & CHKIIF(ooffjungsandetaillist.FItemList(i).Fitemid>=1000000,Format00(8,ooffjungsandetaillist.FItemList(i).Fitemid),Format00(6,ooffjungsandetaillist.FItemList(i).Fitemid)) & ooffjungsandetaillist.FItemList(i).Fitemoption %></td>
	<td><%= ooffjungsandetaillist.FItemList(i).FItemName %></td>
	<td><%= ooffjungsandetaillist.FItemList(i).FItemOptionName %></td>
	<td align=right ><%= ForMatNumber(ooffjungsandetaillist.FItemList(i).FOrgSellPrice,0) %></td>
	<% if ooffjungsandetaillist.FItemList(i).FOrgSellPrice<>ooffjungsandetaillist.FItemList(i).FRealSellPrice then %>
	<td align=right ><font color=blue><%= ForMatNumber(ooffjungsandetaillist.FItemList(i).FRealSellPrice,0) %></font></td>
	<% else %>
	<td align=right ><%= ForMatNumber(ooffjungsandetaillist.FItemList(i).FRealSellPrice,0) %></td>
	<% end if %>
	<td align=center >
	<% if (ooffjungsandetaillist.FItemList(i).FRealSellPrice<>ooffjungsandetaillist.FItemList(i).FOrgSellPrice) and (ooffjungsandetaillist.FItemList(i).FOrgSellPrice<>0) then %>
	    <% if session("ssBctId")="icommang" then %><a href="javascript:reMargin(<%= i %>,<%= CLNG(ooffjungsandetaillist.FItemList(i).FRealSellPrice*(100-orgsellmargin)/100) %>)"><% end if %>
	    <%= 100-CLNG(ooffjungsandetaillist.FItemList(i).FRealSellPrice/ooffjungsandetaillist.FItemList(i).FOrgSellPrice*100*100)/100 %>
	    <% if session("ssBctId")="icommang" then %></a><% end if %>
	<% end if %>
	</td>
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
    <td align=right ><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fcommission,0) %></td>
	<td align=right ><input type=text name=buycash value="<%= ooffjungsandetaillist.FItemList(i).FSuplyPrice %>" size=7 maxlength=9 style="border:1px #999999 solid; text-align=right"></td>

	<% if ooffjungsandetaillist.FItemList(i).FItemNo<0 then %>
	<td align=center ><input type=text name=itemno value="<%= ooffjungsandetaillist.FItemList(i).FItemNo %>" size=3 maxlength=8 style="border:1px #999999 solid; color:#FF0000; text-align=center"></td>
	<td align=right ><font color=red><%= ForMatNumber(ooffjungsandetaillist.FItemList(i).FSuplyPrice * ooffjungsandetaillist.FItemList(i).FItemNo,0) %></font></td>
	<% else %>
	<td align=center ><input type=text name=itemno value="<%= ooffjungsandetaillist.FItemList(i).FItemNo %>" size=3 maxlength=8 style="border:1px #999999 solid; text-align=center"></td>
	<td align=right ><%= ForMatNumber(ooffjungsandetaillist.FItemList(i).FSuplyPrice * ooffjungsandetaillist.FItemList(i).FItemNo,0) %></td>
	<% end if %>

	<td align="center">
    	<% if ooffjungsandetaillist.FItemList(i).Fcentermwdiv="M" then %>
    	<b><%= ooffjungsandetaillist.FItemList(i).Fcentermwdiv %></b>
    	<% else %>
    	<%= ooffjungsandetaillist.FItemList(i).Fcentermwdiv %>
    	<% end if %>

    	<% if ooffjungsandetaillist.FItemList(i).Fvatinclude="N" then %>
    	<font color="red">��</font>
    	<% end if %>
	</td>
	<td align="center"><a href="javascript:delDetail('<%= ooffjungsandetaillist.FItemList(i).Fdetailidx %>');">X</a></td>
</tr>

<% next %>
</form>
<tr bgcolor="#FFFFFF">
	<td colspan=11></td>
	<td align=center ><%= ForMatNumber(subitemno,0) %></td>

	<td align=right ><%= ForMatNumber(subtotal,0) %></td>
	<td ></td>
	<td ></td>
</tr>
</table>
<%
set ooffjungsan     = Nothing
set ooffjungsandetail  = Nothing
set ooffjungsandetaillist         = Nothing
set ochargeuser     = Nothing
%>
<form name=frmarr method=post action="off_jungsan_process.asp">
<input type=hidden name=mode value="modiedtailarr">
<input type=hidden name=masteridx value="<%= idx %>">
<input type=hidden name=idxarr value="">
<input type=hidden name=suplyprice value="">
<input type=hidden name=itemno value="">
<input type=hidden name=detailidx value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->