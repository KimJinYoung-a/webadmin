<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : �Ǹų���[�귣��]
' History	:  2017.04.07 ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate,searchnextdate, designer,itemid,dateback, ix,iy
dim datetype,delivertype,vPurchaseType, sitename, dispCate, maxDepth, sellchnl, chkMinus
dim xl, inc3pl, chknotcash, isPlusSaleItem, isSendGift
	nowdate = Left(CStr(now()),10)
	designer = requestcheckvar(request("designer"),32)
	itemid = requestcheckvar(getNumeric(trim(request("itemid"))),10)
	datetype   = requestcheckvar(request("datetype"),16)
	delivertype = requestcheckvar(request("delivertype"),16)
	yyyy1 = requestcheckvar(request("yyyy1"),4)
	mm1 = requestcheckvar(request("mm1"),2)
	dd1 = requestcheckvar(request("dd1"),2)
	yyyy2 = requestcheckvar(request("yyyy2"),4)
	mm2 = requestcheckvar(request("mm2"),2)
	dd2 = requestcheckvar(request("dd2"),2)
	vPurchaseType = requestCheckVar(request("purchasetype"),16)
	sitename = requestCheckVar(request("sitename"),32)
	dispCate	= requestCheckVar(request("disp"),20)
	maxDepth = "3"
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	xl 			= request("xl")
	chkMinus	= requestCheckVar(Request("chkMinus"),1)
	inc3pl      = request("inc3pl")
	chknotcash	= requestCheckVar(Request("chknotcash"),1)
	isPlusSaleItem = request("isPlusSaleItem")
	isSendGift	= requestCheckvar(request("isSendGift"),1)

if (datetype="") then datetype="jumunil"
if (delivertype="") then delivertype="upche"

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

''dateback = DateSerial(yyyy1,mm2, dd2-7) ''D-1�� ���� 2015/09/21
dateback = DateSerial(yyyy1,mm2, dd2-1)

yyyy1 = Left(dateback,4)
mm1   = Mid(dateback,6,2)
dd1   = Mid(dateback,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim cknodate
cknodate = request("cknodate")

dim page
dim ojumun

page = request("page")
if (page="") then page=1

set ojumun = new CBaljuMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

ojumun.FRectDesignerID = designer
ojumun.FRectItemid = itemid
ojumun.FPageSize = 500
ojumun.FCurrPage = page
ojumun.FRectDateType = datetype
ojumun.FRectDeliverType = delivertype
ojumun.FRectBrandPurchaseType = vPurchaseType
ojumun.FRectSiteName = sitename
ojumun.FRectDispCate = dispCate
ojumun.FRectSellChannelDiv = sellchnl
ojumun.FRectCheckMinus = chkMinus
ojumun.FRectInc3pl = inc3pl
ojumun.FRectchknotcash = chknotcash
ojumun.FRectIsPlusSaleItem = isPlusSaleItem
ojumun.FRectIsSendGift = isSendGift
ojumun.DesignerDateSellList

if (xl = "Y") then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=outmallbranddateselllist_xl.xls"
else
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function ViewOrderDetail(iorderserial){
	var popwin;
    popwin = window.open('/admin/ordermaster/viewordermaster.asp?orderserial=' + iorderserial,'orderdetail','scrollbars=yes,resizable=yes,width=800,height=600');
    popwin.focus();
}

function popOrderDetailEdit(idx){
	var popwin = window.open('/common/orderdetailedit_UTF8.asp?idx=' + idx,'orderdetailedit','width=600,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ViewItem(itemid){
window.open("http://www.10x10.co.kr/street/designershop.asp?itemid=" + itemid,"sample");
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function SubmitForm(frm) {
	if ((CheckDateValid(frm.yyyy1.value, frm.mm1.value, frm.dd1.value) == true) && (CheckDateValid(frm.yyyy2.value, frm.mm2.value, frm.dd2.value) == true)) {
		frm.submit();
	}
}
function popXL()
{
    frmXL.submit();
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� : <% drawSelectBoxDesigner "designer", designer %>
			&nbsp;
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="6">
			&nbsp;
			�˻��Ⱓ :
			<select class="select" name="datetype">
		     	<option value='jumunil' <% if (datetype = "jumunil") then %>selected<% end if %> >�ֹ��ϱ���</option>
		     	<option value='ipkumil' <% if (datetype = "ipkumil") then %>selected<% end if %> >�����ϱ���</option>
		     	<option value='chulgoil' <% if (datetype = "chulgoil") then %>selected<% end if %> >����ϱ���</option>
				 <option value='baesongil' <% if (datetype = "baesongil") then %>selected<% end if %> >����ϱ���</option>
				 <option value='jungsanil' <% if (datetype = "jungsanil") then %>selected<% end if %> >�����ϱ���</option>
	     	</select>
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:SubmitForm(document.frm)">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	     	��۱���
			<select class="select" name="delivertype">
		     	<option value="all" <% if delivertype="all" then response.write "selected" %> >��ü</option>
		     	<option value="ten" <% if (delivertype="ten") then response.write "selected" %> >��ü���</option>
		     	<option value="upche" <% if (delivertype="upche") then response.write "selected" %> >��ü���</option>
	     	</select>
	     	&nbsp;|&nbsp;
    		�������� : 
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
			&nbsp;|&nbsp;
			�߰����� : 
			<% drawSelectBoxIsPlusSaleItem "isPlusSaleItem", isPlusSaleItem %>
    		&nbsp;|&nbsp;
    		����Ʈ :
    		<% 'drawSelectBoxOnIpjumShop "sitename",sitename %>
    		<% Drawsitename "sitename",sitename %>
    		&nbsp;|&nbsp;
			����ī�װ� :
			<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
			&nbsp;|&nbsp;
		    ä�α��� :
		    <% drawSellChannelComboBox "sellchnl",sellchnl %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<b>����ó:</b> <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	     	&nbsp;&nbsp;
	     	<label><input type="checkbox" name="chkMinus" value="Y" <%=chkIIF(chkMinus="Y","checked ","")%>/> ������</label>
			&nbsp;&nbsp;
			<label><input type="checkbox" name="chknotcash" value="Y" <%=chkIIF(chknotcash="Y","checked ","")%>/>���������ֹ�����</label>
			&nbsp;&nbsp;
			<label><input type="checkbox" name="isSendGift" value="Y" <%=CHKIIF(isSendGift<>"","checked","")%>>�����ֹ��� ����</label>
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->

<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#EEEEEE">
	<tr>
		<td align="left">
			* <font color="red">�ǽð�</font> ������ �̸�, �ֱ� <font color="red">6����</font>�ֹ��� ǥ�õ˴ϴ�.
		</td>
		<td align="right">
		<% If sellchnl = "OUT" Then %>
			<input type="button" class="button" value="�����ޱ�" onClick="popXL()">
		<% End If %>
		</td>
	</tr>
</table>
<% end if %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21">
			�˻���� : <b><%= ojumun.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= ojumun.FTotalpage %></b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			���Ǹż��� : <%= FormatNumber(ojumun.FSumItemNo,0) %>
			&nbsp;
			/ �����Ѿ�[��ǰ](��ǰ��������) : <%= FormatNumber(ojumun.FSumItemCost,0) %>
			&nbsp;
			/ �����Ѿ�[��ǰ](��ǰ��������) : <%= FormatNumber(ojumun.FSumBuyCash,0) %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�귣��</td>
		<td width="70">�ֹ���ȣ</td>
		<td>����Ʈ</td>
		<td width="50">�ֹ��ڸ�</td>
<!--		<td width="50">�����θ�</td>	-->
		<td width="50">��ǰ�ڵ�</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="30">����</td>
		<td width="40">����<br>����</td>
		<% if (C_InspectorUser = False) then %>
		<td width="65">�Һ��ڰ�</td>
		<td width="65">�ǸŰ�<br>(��������)</td>
		<% end if %>
		<td width="65">���Ű�<% if (C_InspectorUser = False) then %><br>(��������)<% end if %></td>
		<td width="65">���԰�<% if (C_InspectorUser = False) then %><br>(��������)<% end if %></td>
		<td width="65">��������</td>
		<td width="65">�ֹ���</td>
		<td width="65">������</td>
		<td width="65">�����</td>
		<td width="65">�����</td>
		<td width="65">������</td>
		<td width="40">���<br>����</td>
		<td width="40">���<br>����</td>
		<td width="60">��ǰ����</td>
	</tr>
	<% if ojumun.FresultCount<1 then %>
    <tr align="center" bgcolor="#FFFFFF">
		<td colspan="21" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
	<% else %>

	<%
	dim itemCntSum, itemcostCouponNotAppliedSum, orgItemCostSum, sellCashSum, buyCashSum
		itemCntSum = 0
		orgItemCostSum = 0
		itemcostCouponNotAppliedSum=0
		sellCashSum = 0
		buyCashSum = 0

	for ix=0 to ojumun.FresultCount-1
	itemCntSum = itemcntsum + ojumun.FMasterItemList(ix).FItemcnt
	orgItemCostSum = OrgitemCostsum + ojumun.FMasterItemList(ix).FOrgitemCost
	itemcostCouponNotAppliedSum = itemcostCouponNotAppliedsum + ojumun.FMasterItemList(ix).FitemcostCouponNotApplied
	sellCashSum = sellCashsum + ojumun.FMasterItemList(ix).FSellCash
	buyCashSum = buyCashSum + ojumun.FMasterItemList(ix).FBuyCash
	%>
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="#FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ojumun.FMasterItemList(ix).FMakerid %></td>
		<td><a href="javascript:ViewOrderDetail('<%= ojumun.FMasterItemList(ix).FOrderSerial %>');"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td><%= ojumun.FMasterItemList(ix).Fsitename %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
<!--		<td><%= ojumun.FMasterItemList(ix).FReqname %></td>	-->
		<td><a href="javascript:popOrderDetailEdit(<%= ojumun.FMasterItemList(ix).Fdetailidx %>);"><%= ojumun.FMasterItemList(ix).FItemid %></a></td>
		<td align="left">
			<a href="javascript:popOrderDetailEdit(<%= ojumun.FMasterItemList(ix).Fdetailidx %>);"><%= ojumun.FMasterItemList(ix).FItemname %></a>
			<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
			<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
			<% end if %>
		</td>
		<td>
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
			<font color="red"><%= ojumun.FMasterItemList(ix).FItemcnt %></font>
			<% else %>
			<%= ojumun.FMasterItemList(ix).FItemcnt %>
			<% end if %>
		</td>
		<td>
		    <%= mwdivName(ojumun.FMasterItemList(ix).Fomwdiv) %>
		</td>

		<% if (C_InspectorUser = False) then %>
			<td align="right">
				<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
					<font color="red"><%= FormatNumber(ojumun.FMasterItemList(ix).FOrgitemCost,0) %></font>
				<% else %>
					<%= FormatNumber(ojumun.FMasterItemList(ix).FOrgitemCost,0) %>
				<% end if %>
			</td>
			<td align="right">
				<% if ojumun.FMasterItemList(ix).FOrgitemCost <> ojumun.FMasterItemList(ix).FitemcostCouponNotApplied then %>
					<font color="red"><%= FormatNumber(ojumun.FMasterItemList(ix).FitemcostCouponNotApplied,0) %></font>
				<% else %>
					<%= FormatNumber(ojumun.FMasterItemList(ix).FitemcostCouponNotApplied,0) %>
				<% end if %>
			</td>
		<% end if %>

		<td align="right">
			<% if ojumun.FMasterItemList(ix).FitemcostCouponNotApplied <> ojumun.FMasterItemList(ix).FSellCash then %>
			<font color="green"><%= FormatNumber(ojumun.FMasterItemList(ix).FSellCash,0) %></font>
			<% else %>
			<%= FormatNumber(ojumun.FMasterItemList(ix).FSellCash,0) %>
			<% end if %>
		</td>
		<td align="right">
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
			<font color="red"><%= FormatNumber(ojumun.FMasterItemList(ix).FBuyCash,0) %></font>
			<% else %>
			<%= FormatNumber(ojumun.FMasterItemList(ix).FBuyCash,0) %>
			<% end if %>
		</td>
		<td align="right">
		    <% if (ojumun.FMasterItemList(ix).FSellCash<>0) then %>
				<%= FormatNumber(100.0*((1.0 - ojumun.FMasterItemList(ix).FBuyCash / ojumun.FMasterItemList(ix).FSellCash)),2) %>
		    <% end if %>
		</td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).FOrderdate %>"><%= Left(CStr(ojumun.FMasterItemList(ix).FOrderdate),10) %></acronym></td>
		<td>
			<% if Not IsNull(ojumun.FMasterItemList(ix).FIpkumdate) then %>
			<acronym title="<%= ojumun.FMasterItemList(ix).FIpkumdate %>"><%= Left(CStr(ojumun.FMasterItemList(ix).FIpkumdate),10) %></acronym>
			<% end if %>
		</td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).FUpcheBeasongDate %>"><%= Left(CStr(ojumun.FMasterItemList(ix).FUpcheBeasongDate),10) %></acronym></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).Fdlvfinishdt %>"><%= Left(CStr(ojumun.FMasterItemList(ix).Fdlvfinishdt),10) %></acronym></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).Fjungsanfixdate %>"><%= Left(CStr(ojumun.FMasterItemList(ix).Fjungsanfixdate),10) %></acronym></td>
		<td>
			<% if ojumun.FMasterItemList(ix).FDeliveryType<>"Y" then %>
			��ü
			<% else %>
			<font color="#22AA22">��ü</font>
			<% end if %>
		</td>
		<td>
			<% if ojumun.FMasterItemList(ix).FMasterCancel = "Y" then %>
				<font color="red">���</font>
			<% else %>
				<% if ojumun.FMasterItemList(ix).FCancelYn = "Y" then %>
				<font color="red">���</font>
				<% elseif  ojumun.FMasterItemList(ix).FCancelYn = "A" then %>
				<font color="blue">�߰�</font>
				<% else %>
				&nbsp;
				<% end if %>
			<% end if %>
		</td>
		<td>
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
				<font color="red">���̳ʽ�</font>
			<% else %>
				<% if IsNull(ojumun.FMasterItemList(ix).FIpkumdate) then %>
				�ֹ�����
				<% elseif ojumun.FMasterItemList(ix).FCurrstate = 2 then %>
				<font color="<%= IpkumDivColor(ojumun.FMasterItemList(ix).FCurrstate) %>">�ֹ��뺸</font>
				<% elseif ojumun.FMasterItemList(ix).FCurrstate = 3 then %>
				<font color="<%= IpkumDivColor(ojumun.FMasterItemList(ix).FCurrstate) %>">�ֹ�Ȯ��</font>
				<% elseif ojumun.FMasterItemList(ix).FCurrstate = 7 then %>
				<font color="<%= IpkumDivColor(ojumun.FMasterItemList(ix).FCurrstate) %>">���Ϸ�</font>
				<% end if %>
			<% end if %>
		</td>
	</tr>
	<% next %>

	<tr class="a" align="center" bgcolor="#FFFFFF">
		<td colspan=6>�հ�</td>
		<td><%= FormatNumber(itemcntsum,0) %></td>
		<td>&nbsp;</td>

		<% if (C_InspectorUser = False) then %>
			<td align="right"><%= FormatNumber(orgItemCostSum,0) %></td>
			<td align="right"><%= FormatNumber(itemcostCouponNotAppliedSum,0) %></td>
		<% end if %>

		<td align="right"><%= FormatNumber(sellCashSum,0) %></td>
		<td align="right"><%= FormatNumber(buyCashSum,0) %></td>
		<td align="right">
		    <% if sellCashSum<>0 then %>
				<%= FormatNumber(100.0*((1.0 - buyCashSum / sellCashSum)),2) %>
		    <% end if %>
		</td>
		<td colspan=8>&nbsp;</td>
	</tr>
<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21" align="center">
			<% if ojumun.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
				<% if ix>ojumun.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>

			<% if ojumun.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>
<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="designer" value= <%= designer %>>
	<input type="hidden" name="itemid" value= <%= itemid %>>
	<input type="hidden" name="delivertype" value= <%= delivertype %>>
	<input type="hidden" name="yyyy1" value= <%= yyyy1 %>>
	<input type="hidden" name="mm1" value= <%= mm1 %>>
	<input type="hidden" name="dd1" value= <%= dd1 %>>
	<input type="hidden" name="yyyy2" value= <%= yyyy2 %>>
	<input type="hidden" name="mm2" value= <%= mm2 %>>
	<input type="hidden" name="dd2" value= <%= dd2 %>>
	<input type="hidden" name="purchasetype" value= <%= vPurchaseType %>>
	<input type="hidden" name="sitename" value= <%= sitename %>>
	<input type="hidden" name="disp" value= <%= dispCate %>>
	<input type="hidden" name="sellchnl" value= <%= sellchnl %>>
</form>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
