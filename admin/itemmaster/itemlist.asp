<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ����
' History : ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, itemcouponyn, overSeaYn, itemdiv
dim cdl, cdm, cds, showminusmagin, marginup, margindown, dispCate, showerrbuycash, pojangok
dim page, sDt, eDt, reserveItemTp
dim infodivYn, infodiv, deliverytype, deliverfixday, vPurchasetype, sortDiv
itemid      = requestCheckvar(request("itemid"),1500)
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),10)
usingyn     = requestCheckvar(request("usingyn"),10)
danjongyn   = requestCheckvar(request("danjongyn"),10)
mwdiv       = requestCheckvar(request("mwdiv"),10)
limityn     = requestCheckvar(request("limityn"),10)
vatyn       = requestCheckvar(request("vatyn"),10)
sailyn      = requestCheckvar(request("sailyn"),10)
itemcouponyn = requestCheckvar(request("itemcouponyn"),10)
overSeaYn   = requestCheckvar(request("overSeaYn"),10)
itemdiv     = requestCheckvar(request("itemdiv"),10)
deliverytype= requestCheckvar(request("deliverytype"),10)
deliverfixday= requestCheckvar(request("deliverfixday"),10)
pojangok	= requestCheckvar(request("pojangok"),10)
vPurchasetype = request("purchasetype")
reserveItemTp	= requestCheckvar(request("reserveItemTp"),1)
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)

showminusmagin = request("showminusmagin")
showerrbuycash = request("showerrbuycash")
marginup = request("marginup")
margindown = request("margindown")

sDt     = requestCheckvar(request("sDt"),10)
eDt     = requestCheckvar(request("eDt"),10)
sortDiv	= requestCheckvar(request("sortDiv"),5)
if sortDiv="" then sortDiv="new"

infodiv  = request("infodiv")
infodivYn  = requestCheckvar(request("infodivYn"),10)

If infodiv <> "" Then
	infodivYn = "Y"
End If

If marginup <> "" AND IsNumeric(marginup) = False Then
	rw "<script>alert('������(�̻�)�� �߸��Ǿ����ϴ�. - "&marginup&"');history.back();</script>"
	dbget.close()
	Response.End
End If

If margindown <> "" AND IsNumeric(margindown) = False Then
	rw "<script>alert('������(����)�� �߸��Ǿ����ϴ�. - "&margindown&"');history.back();</script>"
	dbget.close()
	Response.End
End If

page = requestCheckvar(request("page"),10)

if (page="") then page=1

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if


'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 100
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectDanjongyn    = danjongyn
oitem.FRectLimityn      = limityn
oitem.FRectMWDiv        = mwdiv
oitem.FRectVatYn        = vatyn
oitem.FRectSailYn       = sailyn
oitem.FRectCouponYN		= itemcouponyn
oitem.FRectIsOversea	= overSeaYn
oitem.FRectpojangok		= pojangok

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectItemDiv      = itemdiv

oitem.FRectMinusMigin = showminusmagin
oitem.FRectCheckBuycash = showerrbuycash
oitem.FRectMarginUP = marginup
oitem.FRectMarginDown = margindown
oitem.FRectInfodivYn    = infodivYn
oitem.FRectInfodiv    = infodiv
oitem.FRectDeliverytype = deliverytype
oitem.FRectStartDate = sDt
oitem.FRectEndDate = eDt
oitem.FRectdeliverfixday = deliverfixday
oitem.FRectPurchasetype = vPurchasetype
oitem.FRectSortDiv		= sortDiv
oitem.FRectreserveItemTp		= reserveItemTp
oitem.GetItemList

dim i

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function NextPage(ipage){
	document.frm.page.value= ipage;
	if ((document.frm.itemname.value.length>0)&&(ipage*1==1)){
	    alert('��ǰ�� �˻��� ����� �ִ� 1000���� ���ѵ˴ϴ�.');  // 2������ fulltext �˻��� ���ι������ ����.
	}
	document.frm.target="_self";
	document.frm.action="itemlist.asp";
	document.frm.submit();
}

// ============================================================================
// �ɼǼ��� -��ü
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//�Ǹż���
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?menupos=<%=request("menupos")%>&itemid=' + iitemid,'itemselledit','width=1400,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}

// ============================================================================
// �̹�������
function editItemImage(itemid, makerid) {
	var param = "itemid=" + itemid;

	//if(makerid =="ithinkso"){
		//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}else{
		popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=1000,height=900,scrollbars=yes,resizable=yes');
	//}
	popwin.focus();
}

// ��ǰ���� �̹��� ���/����
function popItemContImage(itemid)
{
	var popwin = window.open("/admin/shopmaster/item_imgcontents_write.asp?mode=edit&itemid=" + itemid + "&menupos=423","popitemContImage","width=600 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// �����Ȳ �˾�
function PopItemStock(itemid){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// �⺻���� ����
function editItemBasicInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// �ǸŰ� �� ���ް� ����
function editItemPriceInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('pop_ItemPriceInfo.asp?' + param ,'editItemPrice','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//Ƽ�� ��ǰ ���� ����
function editTicketItemInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('/admin/itemmaster/pop_ticketIteminfo.asp?' + param ,'pop_ticketIteminfo','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//����,�鼼 ���� �˾�
function vatedit(itemid,vat){
	var param = "itemid=" + itemid + "&vat="+vat+"";
	popwin = window.open('/admin/itemmaster/pop_vatEdit.asp?' + param ,'pop_vatEdit','width=300,height=150');
	popwin.focus();
}

function jsPopEditBuyItemInfo(itemgubun, itemid, itemoption){
	var param = "itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption;
	popwin = window.open('/admin/itemmaster/pop_BuyItemEdit.asp?' + param ,'jsPopEditBuyItemInfo','width=1500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//��ǰ����Ʈ �ٿ�
function jsItemDown(){
  document.frm.page.value = $('#selDCnt').val();
	document.frm.target="hidifr";
	//document.frm.action="itemlist_csv.asp";
	document.frm.action="/admin/itemmaster/itemlist_exceldownload.asp";
	document.frm.submit();
}

// �ɼǺ� ��ǰ����Ʈ �ٿ�
function jsItemoptionDown(){
  document.frm.page.value = $('#selODCnt').val();
	document.frm.target="hidifr";
	document.frm.action="/admin/itemmaster/itemlistoption_exceldownload.asp";
	document.frm.submit();
}

// ��ǰ����(����)
function itemedituploadexcel(){
	document.domain = "10x10.co.kr";
	var popwin = window.open('/admin/itemmaster/pop_itemlist_excel_upload_edit.asp','addreg','width=1400,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ��ǰ�űԵ��(����)
function itemreguploadexcel(){
	document.domain = "10x10.co.kr";
	var popwin = window.open('/admin/itemmaster/pop_itemlist_excel_upload_reg.asp','addreg','width=1400,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>


<style>
	p {margin:0; padding:0; border:0; font-size:100%;}
	i, em, address {font-style:normal; font-weight:normal;}
 .xls, .down {background-image:url(/images/partner/admin_element.png); background-repeat:no-repeat;}
.btn2 {display:inline-block; font-size:11px !important; letter-spacing:-0.025em; line-height:110%; border-left:1px solid #f0f0f0; border-top:1px solid #f0f0f0; border-right:1px solid #cdcdcd; border-bottom:1px solid #cdcdcd; background-color:#f2f2f2; background-image:-webkit-linear-gradient(#fff, #e1e1e1); background-image:-moz-linear-gradient(#fff, #e1e1e1); background-image:-ms-linear-gradient(#fff, #e1e1e1); background-image:linear-gradient(#fff, #e1e1e1); text-align:center; cursor:pointer;}
.btn2 a {display:block; font-size:11px !important; text-decoration:none !important;}
.btn2 span {display:block;}
.btn2 span em {display:block; padding-top:7px; padding-bottom:4px; text-align:center;}

.fIcon {padding-left:33px;}
.eIcon {padding-right:25px;}

.btn2 .xls {background-position:-125px -135px;}
.btn2 .down {background-position:right -231px;}
.cBk1, .cBk1 a {color:#000 !important;}
	</style>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=post>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<input type="hidden" name="sortDiv" value="<%=sortDiv%>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<table border="0" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td style="white-space:nowrap;">�귣��: <%	drawSelectBoxDesignerWithName "makerid", makerid %> </td>
					<td style="white-space:nowrap;padding-left:5px;">��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"> </td>
					<td style="white-space:nowrap;padding-left:5px;">��ǰ�ڵ�:</td>
					<td style="white-space:nowrap;" rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
			</tr>
			<tr>
				<td  style="white-space:nowrap;">����<!-- #include virtual="/common/module/categoryselectbox.asp"--> </td>
				<td  style="white-space:nowrap;"  colspan="2">&nbsp;&nbsp;����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"--> </td>
				<td ></td>
			</tr>
		</table>

		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="NextPage(1);">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			<span style="white-space:nowrap;">�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">���:<% drawSelectBoxUsingYN "usingyn", usingyn %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">����:<% drawSelectBoxLimitYN "limityn", limityn %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">����: <% drawSelectBoxVatYN "vatyn", vatyn %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">���� <% drawSelectBoxSailYN "sailyn", sailyn %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">���� <% drawSelectBoxCouponYN "itemcouponyn", itemcouponyn %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">�ؿܹ�� <% drawSelectBoxIsOverSeaYN "overSeaYn", overSeaYn %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">��۱��� <% drawBeadalDiv "deliverytype", deliverytype %></span>
	     	&nbsp;
	     	<span style="white-space:nowrap;">��۹�� <% drawdeliverfixday "deliverfixday", deliverfixday, "" %></span>
            &nbsp;
	     	<span style="white-space:nowrap;">��ǰ���� <% drawSelectBoxItemDiv "itemdiv", itemdiv %></span>
            &nbsp;
	     	<span style="white-space:nowrap;">���忩�� <% drawSelectBoxPackYN "pojangok", pojangok %></span>
	     	<br>
	     	<span style="white-space:nowrap;"><font color="red">ǰ�������Է¿���</font>
	     	<select class="select" name="infodivYn">
            <option value="">��ü</option>
            <option value="N" <%= CHKIIF(infodivYn="N","selected","") %> >�Է�����</option>
            <option value="Y" <%= CHKIIF(infodivYn="Y","selected","") %> >�Է¿Ϸ�</option>
            </select></span>&nbsp;
			<span style="white-space:nowrap;">ǰ��:
			<select name="infodiv" class="select">
				<option value="" >===��ü====</option>
				<option value="01" <%=chkIIF(infodiv="01","selected","")%>>01.�Ƿ�</option>
				<option value="02" <%=chkIIF(infodiv="02","selected","")%>>02.����/�Ź�</option>
				<option value="03" <%=chkIIF(infodiv="03","selected","")%>>03.����</option>
				<option value="04" <%=chkIIF(infodiv="04","selected","")%>>04.�м���ȭ</option>
				<option value="05" <%=chkIIF(infodiv="05","selected","")%>>05.ħ����/Ŀư</option>
				<option value="06" <%=chkIIF(infodiv="06","selected","")%>>06.����</option>
				<option value="07" <%=chkIIF(infodiv="07","selected","")%>>07.������</option>
				<option value="08" <%=chkIIF(infodiv="08","selected","")%>>08.������ ������ǰ</option>
				<option value="09" <%=chkIIF(infodiv="09","selected","")%>>09.��������</option>
				<option value="10" <%=chkIIF(infodiv="10","selected","")%>>10.�繫����</option>
				<option value="11" <%=chkIIF(infodiv="11","selected","")%>>11.���б��</option>
				<option value="12" <%=chkIIF(infodiv="12","selected","")%>>12.��������</option>
				<option value="13" <%=chkIIF(infodiv="13","selected","")%>>13.�޴���</option>
				<option value="14" <%=chkIIF(infodiv="14","selected","")%>>14.������̼�</option>
				<option value="15" <%=chkIIF(infodiv="15","selected","")%>>15.�ڵ�����ǰ</option>
				<option value="16" <%=chkIIF(infodiv="16","selected","")%>>16.�Ƿ���</option>
				<option value="17" <%=chkIIF(infodiv="17","selected","")%>>17.�ֹ��ǰ</option>
				<option value="18" <%=chkIIF(infodiv="18","selected","")%>>18.ȭ��ǰ</option>
				<option value="19" <%=chkIIF(infodiv="19","selected","")%>>19.�ͱݼ�/����/�ð��</option>
				<option value="20" <%=chkIIF(infodiv="20","selected","")%>>20.��ǰ</option>
				<option value="21" <%=chkIIF(infodiv="21","selected","")%>>21.������ǰ</option>
				<option value="22" <%=chkIIF(infodiv="22","selected","")%>>22.�ǰ���ɽ�ǰ</option>
				<option value="23" <%=chkIIF(infodiv="23","selected","")%>>23.�����ƿ�ǰ</option>
				<option value="24" <%=chkIIF(infodiv="24","selected","")%>>24.�Ǳ�</option>
				<option value="25" <%=chkIIF(infodiv="25","selected","")%>>25.��������ǰ</option>
				<option value="26" <%=chkIIF(infodiv="26","selected","")%>>26.����</option>
				<option value="27" <%=chkIIF(infodiv="27","selected","")%>>27.ȣ��/��� ����</option>
				<option value="28" <%=chkIIF(infodiv="28","selected","")%>>28.������Ű��</option>
				<option value="29" <%=chkIIF(infodiv="29","selected","")%>>29.�װ���</option>
				<option value="30" <%=chkIIF(infodiv="30","selected","")%>>30.�ڵ��� �뿩 ����</option>
				<option value="31" <%=chkIIF(infodiv="31","selected","")%>>31.��ǰ�뿩 ����</option>
				<option value="32" <%=chkIIF(infodiv="32","selected","")%>>32.��ǰ�뿩 ����</option>
				<option value="33" <%=chkIIF(infodiv="33","selected","")%>>33.������ ������</option>
				<option value="34" <%=chkIIF(infodiv="34","selected","")%>>34.��ǰ��/����</option>
				<option value="35" <%=chkIIF(infodiv="35","selected","")%>>35.��Ÿ</option>
			</select></span>&nbsp;&nbsp;
			��������: 
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
		</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			<span style="white-space:nowrap;">
				����<input type="text" class="text" name="marginup" value="<%=marginup%>" size="4">%�̻�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����<input type="text" class="text" name="margindown" value="<%=margindown%>" size="4">%����&nbsp;&nbsp;&nbsp;
				<input type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>��������</font>��ǰ����
				&nbsp;|&nbsp;
				<input type="checkbox" name="showerrbuycash" <%= ChkIIF(showerrbuycash="on","checked","") %> ><font color=red>���԰�����</font>��ǰ����
			</span>
			&nbsp;&nbsp;
			<span style="white-space:nowrap;">
				�ǸŽ�����
				<input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
				<input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "sDt", trigger    : "sDt_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "eDt", trigger    : "eDt_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</span>
			&nbsp;
			<input type="checkbox" name="reserveItemTp" <%= ChkIIF(reserveItemTp="on","checked","") %> >�ܵ�(����)���Ż�ǰ
		</td>
	</tr>
   </form>
</table>
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% If cdl = "110" and cdm = "010" and cds = "968" Then %>
			<input type="button" value="����� ���ø��ڵ� ���" class="button" onClick="window.open('pop_photobook.asp','popPhotobook','width=600,height=650,scrollbars=yes');">
		<% End If %>
	</td>
	<td align="right">
	</td>
</tr>
<tr>
	<td align="left">
		<input type="button" class="button" value="��ǰ�űԵ��(����)" onclick="itemreguploadexcel()">
		<input type="button" class="button" value="��ǰ����(����)" onclick="itemedituploadexcel()">
	</td>
	<td align="right">
		<select name="tsrtDv" class="select" style="height:25px;vertical-align:top;" onchange="document.frm.sortDiv.value=this.value;">
			<option value="new" <%=chkIIF(sortDiv="new","selected","")%>>�Ż��</option>
			<option value="best" <%=chkIIF(sortDiv="best","selected","")%>>�α��</option>
			<option value="cashH" <%=chkIIF(sortDiv="cashH","selected","")%>>����</option>
			<option value="cashL" <%=chkIIF(sortDiv="cashL","selected","")%>>������</option>
		</select>
		<%dim   imax, imin%>
		<select name="selDCnt" id="selDCnt" class="select" style="height:25px;vertical-align:top;">
			<%for i =1 To Int(oitem.FTotalCount/5000)+1
					imin = ((i-1)*5000)+1
					if i <  Int(oitem.FTotalCount/5000)+1 then
					imax = i*5000
					else
					imax = oitem.FTotalCount
					end if
			%>
			<option value="<%=i%>"><%=imin%>~<%=imax%></option>
			<%Next%>
		</select>
		<input type="button" class="button" value="��ǰ�ٿ�ε�(����)" onclick="jsItemDown();">
		&nbsp;&nbsp;
		<select name="selODCnt" id="selODCnt" class="select" style="height:25px;vertical-align:top;">
			<%for i =1 To Int(oitem.FTotalCount/2000)+1
					imin = ((i-1)*2000)+1
					if i <  Int(oitem.FTotalCount/2000)+1 then
					imax = i*2000
					else
					imax = oitem.FTotalCount
					end if
			%>
			<option value="<%=i%>"><%=imin%>~<%=imax%></option>
			<%Next%>
		</select>
		<input type="button" class="button" value="��ǰ�ɼǺ��ٿ�ε�(����)" onclick="jsItemoptionDown();">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><%= oitem.FTotalCount%></b>
			&nbsp;
			������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" class="sticky_top">
		<td width="60">��ǰ��ȣ</td>
		<td width=50> �̹���</td>
		<td width="100">�귣��ID<br><font color=darkblue>[ǥ�ú귣��]</font></td>
		<td> ��ǰ��</td>
		<td width="60">�ǸŰ�</td>
		<td width="60">���԰�</td>
		<td width="40">����</td>
		<td width="30">���<br>����</td>
		<td width="30">�Ǹ�<br>����</td>
		<td width="30">���<br>����</td>
		<td width="30">����<br>����</td>
		<td width="36">����<br>�鼼</td>
		<td width="40">���<br>��Ȳ</td>
		<td width="40">����<br />����</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="16" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center">
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="center"><a href="javascript:editItemImage('<%= oitem.FItemList(i).FItemId %>','<%= oitem.FItemList(i).Fmakerid %>')" title="�̹��� ����"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></a></td>
		<td align="left">
			<a href="javascript:PopBrandInfoEdit('<%= oitem.FItemList(i).Fmakerid %>')" title="�귣�� ���� ����"><%= oitem.FItemList(i).Fmakerid %></a>
			<%
				if oitem.FItemList(i).FfrontMakerid<>"" and trim(oitem.FItemList(i).Fmakerid)<>trim(oitem.FItemList(i).FfrontMakerid) then
					Response.Write "<br /><font color=darkblue>[" & oitem.FItemList(i).FfrontMakerid & "]</font>"
				end if
			%>
		</td>
		<td align="left">
			<% '������ ��ǰ���� ���� ��ǰ�� ����� �Ǿ��µ� MD�� ������ �Ұ�����. [B] �߰���. ' 2023.04.25 �ѿ�� %>
			<a href="javascript:editItemBasicInfo('<% =oitem.FItemList(i).Fitemid %>')" title="��ǰ �⺻���� ����"><% =oitem.FItemList(i).Fitemname %>[B]</a>

			<% if ((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or (session("ssAdminPOsn") = "5") or session("ssAdminPsn")=7 or C_ADMIN_AUTH) then %>
				<a href="itemmodify.asp?itemid=<% =oitem.FItemList(i).Fitemid %>&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>" target="_blank" title="��ü���� ����"><font color="#8090F0">[E]</font></a>
			<% end if %>

			<a href="pop_itemContentEdit.asp?itemid=<% =oitem.FItemList(i).Fitemid %>" target="_blank" onclick="window.open(this.getAttribute('href'),'popEditItemCont','width=1400,height=800');return false;" title="��ǰ���� ����"><font color="#FF7F50">[C]</font></a>

			<% if oitem.FItemList(i).FitemDiv="08" then %>
				<a href="javascript:editTicketItemInfo('<% =oitem.FItemList(i).Fitemid %>')" title="Ticket ���� ����"><font color="#F89020">[Ticket]</font></a>
			<% end if %>

			<% if oitem.FItemList(i).FitemDiv="18" then %>
				<a href="javascript:editTicketItemInfo('<% =oitem.FItemList(i).Fitemid %>')" title="travel ���� ����"><font color="#F89020">[travel]</font></a>
			<% end if %>

			<% if oitem.FItemList(i).FitemDiv="75" then %>
				<font color="#F12353">[���ⱸ��]</font>
			<% end if %>
		</td>
		<td align="right">
		<%
			Response.Write "<a href=""javascript:editItemPriceInfo('" & oitem.FItemList(i).Fitemid & "')"" title='�ǸŰ� �� ���ް� ����'>" & FormatNumber(oitem.FItemList(i).Forgprice,0) & "</a>"
			'���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) & "%��)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'������
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
				end Select
'[ON]��ǰ����>>��ǰ���� �� ���� ���� ���� �����ΰ�� ���� & ����%�� ���� �ʴٰ� �Ͽ� ����. 2011-04-20 ���ر�.
'				Select Case oitem.FItemList(i).FitemCouponType
'					Case "1"
'						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
'					Case "2"
'						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
'				end Select
			end if
		%>
		</td>
		<td align="right">
		<%

			'���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
			    if (oitem.FItemList(i).Fsailsuplycash>oitem.FItemList(i).Forgsuplycash) then
			        Response.Write "<strong>"&FormatNumber(oitem.FItemList(i).Forgsuplycash,0)&"</strong>"
			        Response.Write "<br><strong><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font></strong>"
			    else
			        Response.Write FormatNumber(oitem.FItemList(i).Forgsuplycash,0)
    				Response.Write "<br><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font>"
    			end if
    		else
    		    Response.Write FormatNumber(oitem.FItemList(i).Forgsuplycash,0)
			end if
			'������
			if oitem.FItemList(i).FitemCouponYn="Y" then
				if oitem.FItemList(i).FitemCouponType="1" or oitem.FItemList(i).FitemCouponType="2" then
					if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Forgsuplycash,0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Fcouponbuyprice,0) & "</font>"
					end if
				end if
			end if
		%>
		</td>
		<td align="right">
		<%
			Response.Write fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice,1)
			'���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>" & fnPercent(oitem.FItemList(i).Fsailsuplycash,oitem.FItemList(i).Fsailprice,1) & "</font>"
			end if
			'������
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fbuycash,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						end if
					Case "2"
						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fbuycash,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
						end if
				end Select
'[ON]��ǰ����>>��ǰ���� �� ���� ���� ���� �����ΰ�� ���� & ����%�� ���� �ʴٰ� �Ͽ� ����. 2011-04-20 ���ر�.
'				Select Case oitem.FItemList(i).FitemCouponType
'					Case "1"
'						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
'							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),1) & "</font>"
'						else
'							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),1) & "</font>"
'						end if
'					Case "2"
'						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
'							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,1) & "</font>"
'						else
'							Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,1) & "</font>"
'						end if
'				end Select
			end if
		%>
		</td>
		<td align="center"><a href="javascript:PopItemSellEdit('<%= oitem.FItemList(i).FItemId %>')" title="�Ǹ�����/�ɼ� ����"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></a>
			<br>
			<%
				If oitem.FItemList(i).Fdeliverytype = "1" Then
					response.write "�ٹ�"
				ElseIf oitem.FItemList(i).Fdeliverytype = "2" Then
					response.write "����"
				ElseIf oitem.FItemList(i).Fdeliverytype = "4" Then
					response.write "�ٹ�"
				ElseIf oitem.FItemList(i).Fdeliverytype = "9" Then
					response.write "����"
				ElseIf oitem.FItemList(i).Fdeliverytype = "7" Then
					response.write "����"
				End If
			%>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fvatinclude,"tx") %><br><a href="javascript:vatedit('<%= oitem.FItemList(i).FItemId %>','<%=oitem.FItemList(i).Fvatinclude%>');">[����]</a></td>
	    <td align="center"><a href="javascript:PopItemStock('<%= oitem.FItemList(i).FItemId %>')" title="�����Ȳ �˾�">[����]</a></td>
		<td align="center"><a href="javascript:jsPopEditBuyItemInfo('10','<%= oitem.FItemList(i).FItemId %>','0000');">[����]</a></td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
			<% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
    			<% if i>oitem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oitem.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>

</table>
<% end if %>

<iframe id="hidifr" src="" width="0" height="0" frameborder="0"></iframe>

<%
SET oitem = Nothing
%>
<div style="padding:5px;text-align:right;font-size:8pt">Ver1.0  lastupdate: 2013.12.18 </div>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
