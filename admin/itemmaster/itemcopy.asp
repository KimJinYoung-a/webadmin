<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, overSeaYn, itemdiv
Dim cdl, cdm, cds, showminusmagin, marginup, margindown, dispCate, showerrbuycash
Dim page, sDt, eDt
Dim infodivYn, infodiv, deliverytype, deliverfixday, vPurchasetype, sortDiv, showCopyitem
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
overSeaYn   = requestCheckvar(request("overSeaYn"),10)
itemdiv     = requestCheckvar(request("itemdiv"),10)
deliverytype= requestCheckvar(request("deliverytype"),10)
deliverfixday= requestCheckvar(request("deliverfixday"),10)
vPurchasetype = request("purchasetype")

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
showCopyitem = requestCheckvar(request("showCopyitem"),2)

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
	oitem.FRectIsOversea	= overSeaYn
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
	oitem.FRectShowCopyitem	= showCopyitem
	oitem.GetItemCopyList
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
	document.frm.action="itemcopy.asp";
	document.frm.submit();
}

// ���õ� ��ǰ ����
function itemCopyProcess() {
	var chkSel=0;
	var itemarr = document.getElementsByName('itemarr');
	var brandarr = document.getElementsByName('brandarr');
	var itemdivarr = document.getElementsByName('itemdivarr');

	var v1, v2, v3;
	v1 = "";
	v2 = "";
	v3 = "";
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) {
					chkSel++;
					v1 = v1 + frmSvArr.cksel[i].value + '||';
					if (frmSvArr.changemakerid[i].value == ""){
						alert('������ �귣�带 �Է��ϼ���');
						frmSvArr.changemakerid[i].focus();
						return;
					}else{
						v2 = v2 + frmSvArr.changemakerid[i].value + '||';
					}
					if (frmSvArr.changeitemdiv[i].value == ""){
						alert('������ ��ǰ������ �����ϼ���');
						frmSvArr.changeitemdiv[i].focus();
						return;
					}else{
						v3 = v3 + frmSvArr.changeitemdiv[i].value + '||';
					}
				}
			}
		}else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert(e);
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
		$("#itemarr").val(v1);
		$("#brandarr").val(v2);
		$("#itemdivarr").val(v3);
		$("#cmdparam").val("itemcopy");
		document.getElementById("btnCopy").disabled=true;
		document.frmArr.target = "xLink";
		document.frmArr.action = "<%=apiURL%>/itemcopy/actItemReq.asp"
		document.frmArr.submit();
    }
}
function btnOk(){
	$("input[name=changemakerid]").val( $("#copyid").val() );
}
function btnOk2(){
	$("select[name=changeitemdiv]").val( $("select[name=copyitemdiv]").val() );
}
function popHistory(){
	var pCM = window.open("/admin/itemmaster/popCopyHistory.asp","popHistory","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
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

<form name="frmArr" method=post>
	<input type="hidden" id= "itemarr" name="itemarr" value="" />
	<input type="hidden" id= "brandarr" name="brandarr" value="" />
	<input type="hidden" id= "itemdivarr" name="itemdivarr" value="" />
	<input type="hidden" id= "cmdparam" name="cmdparam" value="" />
</form>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=post>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<input type="hidden" name="sortDiv" value="<%=sortDiv%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
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
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
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
		<span style="white-space:nowrap;">�ؿܹ�� <% drawSelectBoxIsOverSeaYN "overSeaYn", overSeaYn %></span>
		&nbsp;
		<span style="white-space:nowrap;">��۱��� <% drawBeadalDiv "deliverytype", deliverytype %></span>
		&nbsp;
		<span style="white-space:nowrap;">��۹�� <% drawdeliverfixday "deliverfixday", deliverfixday, "" %></span>
		&nbsp;
		<span style="white-space:nowrap;">��ǰ���� <% drawSelectBoxItemDiv "itemdiv", itemdiv %></span>
		<br>
		<span style="white-space:nowrap;"><font color="red">ǰ�������Է¿���</font>
			<select class="select" name="infodivYn">
				<option value="">��ü</option>
				<option value="N" <%= CHKIIF(infodivYn="N","selected","") %> >�Է�����</option>
				<option value="Y" <%= CHKIIF(infodivYn="Y","selected","") %> >�Է¿Ϸ�</option>
			</select>
		</span>&nbsp;
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
			</select>
		</span>&nbsp;&nbsp;
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
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		<label>
			<input type="checkbox" name="showCopyitem" <%= ChkIIF(showCopyitem="on","checked","") %> ><font color=red>������</font>��ǰ����
		</label>

	</td>
</tr>
</form>
</table>
<br />
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
	    <table width="100%" class="a">
	    <tr>
	    	<td>
				��ǰ ���� : <input class="button" type="button" id="btnCopy" value="����" onClick="itemCopyProcess();">
			</td>
	    	<td align="right">
				<input class="button" type="button" value="�̷�" onClick="popHistory();">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
<br />
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�˻���� : <b><%= oitem.FTotalCount%></b>
		&nbsp;
		������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="60">No.</td>
	<td width=50> �̹���</td>
	<td width="100">�귣��ID</td>
	<td width="100">
		������ �귣��ID <br />
		<input type="text" class="text" name="copyid" id="copyid" value="" size="20" >
		<input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'copyid');" >
		<input type="button" class="button" value="�ϰ�����" onclick="btnOk();">
	</td>
	<td width="100">
		������ ��ǰ���� <br />
		<select class="select" name="copyitemdiv" onchange="btnOk2();">
			<option value="01">�Ϲ�</option>
			<option value="08">Ƽ��/Ŭ���� ��ǰ</option>
			<option value="09">Present��ǰ</option>
			<option value="18">�����ǰ</option>
			<option value="75">���ⱸ����ǰ</option>
			<option value="30">�̴Ϸ�Ż��ǰ</option>
			<option value="23">B2B��ǰ</option>
			<option value="16">�ֹ�����</option>
			<option value="06">�ֹ�����(����)</option>
		</select>
	</td>
	<td>��ǰ��</td>
	<td width="60">�ǸŰ�</td>
	<td width="60">���԰�</td>
	<td width="40">����</td>
	<td width="30">���<br>����</td>
	<td width="30">�Ǹ�<br>����</td>
	<td width="30">���<br>����</td>
	<td width="30">����<br>����</td>
	<td width="36">����<br>�鼼</td>
</tr>
<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oitem.FItemList(i).Fitemid %>"></td>
	<td align="center">
		<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">
		<%= oitem.FItemList(i).Fitemid %></a>
		</td>
	<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
	<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
	<td align="left">
		<input type="text" class="text" size="20" name="changemakerid" id="changemakerid<%= i %>" value="">
	</td>
	<td align="left">
		<select class="select" name="changeitemdiv">
			<option value="01" <%= chkiif(oitem.FItemList(i).Fitemdiv="01", "selected", "") %>>�Ϲ�</option>
			<option value="08" <%= chkiif(oitem.FItemList(i).Fitemdiv="08", "selected", "") %>>Ƽ��/Ŭ���� ��ǰ</option>
			<option value="09" <%= chkiif(oitem.FItemList(i).Fitemdiv="09", "selected", "") %>>Present��ǰ</option>
			<option value="11" <%= chkiif(oitem.FItemList(i).Fitemdiv="11", "selected", "") %>>��ǰ�ǻ�ǰ</option>
			<option value="18" <%= chkiif(oitem.FItemList(i).Fitemdiv="18", "selected", "") %>>�����ǰ</option>
			<option value="75" <%= chkiif(oitem.FItemList(i).Fitemdiv="75", "selected", "") %>>���ⱸ����ǰ</option>
			<option value="30" <%= chkiif(oitem.FItemList(i).Fitemdiv="30", "selected", "") %>>�̴Ϸ�Ż��ǰ</option>
			<option value="23" <%= chkiif(oitem.FItemList(i).Fitemdiv="23", "selected", "") %>>B2B��ǰ</option>
			<option value="16" <%= chkiif(oitem.FItemList(i).Fitemdiv="16", "selected", "") %>>�ֹ�����</option>
			<option value="06" <%= chkiif(oitem.FItemList(i).Fitemdiv="06", "selected", "") %>>�ֹ�����(����)</option>
		</select>
	</td>
	<td align="left">
		<% =oitem.FItemList(i).Fitemname %>
		<% if oitem.FItemList(i).FitemDiv="75" then %>
			<font color="#F12353">[���ⱸ��]</font>
		<% end if %>
	</td>
	<td align="right">
	<%
		Response.Write FormatNumber(oitem.FItemList(i).Forgprice,0)
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
		end if
	%>
	</td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %>
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
	<td align="center"><%= fnColor(oitem.FItemList(i).Fvatinclude,"tx") %></td>
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
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<% end if %>
<% SET oitem = Nothing %>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
