<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ ���ڵ� ���
' Hieditor : 2009.04.07 ������ ����
'			 2012.04.23 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim itemgubun, isusingyn, research ,designer, iitemid, barcode ,obarcode ,i ,makeriddispyn ,printpriceyn
	makeriddispyn 			= requestCheckVar(request("makeriddispyn"),1)
	printpriceyn 	= requestCheckVar(request("printpriceyn"),1)
	iitemid = request("iitemid")
	barcode = request("barcode")
	research = request("research")
	itemgubun = request("itemgubun")
	isusingyn = request("isusingyn")

designer = session("ssBctID")
if makeriddispyn = "" then makeriddispyn = "Y"
if (research="") and (isusingyn="") then isusingyn="Y"

'''REquire Paging

set obarcode = new COffShopItem
	obarcode.FPageSize= 500
	obarcode.FRectItemgubun = itemgubun
	obarcode.FRectDesigner = designer
	obarcode.FRectBarCode = barcode
	obarcode.FRectItemId = iitemid
	obarcode.FRectOnlyUsing = ChkIIF(isusingyn="Y","on","")

	if (designer<>"") or (iitemid<>"") then
		obarcode.GetBarCodeList
	end if

%>

<script language='javascript'>

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function AddData(itemid, itemoption, itemname, itemoptionname, brand, itemprice, itemtype, itemno){
	iaxobject.AddData(itemid, itemoption, itemname, itemoptionname, brand, itemprice, itemtype, itemno);
}

//AddData(v,'0000','�����۸�','�ɼǸ�','�귣��',3000,'T','5')
function AddArr(){
	var frmmaster = document.frm;
	var frm;
	var pass = false;
	var makeriddisp;
	var printprice;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}
	iaxobject.ClearItem();
	//iaxobject.setTitleVisible(true);
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				//�귣��ǥ��
				if (frmmaster.makeriddispyn.value != 'N'){
					makeriddisp = frm.brand.value;
				}else{
					makeriddisp = '';
				}

				//�ݾ�ǥ��
				//if (frmmaster.printpriceyn.value != 'N'){
					printprice = frm.sellprice.value;
				//}else{
				//	printprice = '';
				//}

                if (frm.itemid.value*1>=1000000){
                    AddData(frm.itemid.value,frm.itemoption.value,frm.itemname.value,frm.itemoptionname.value,makeriddisp ,printprice ,frm.itemgubun.value*10,frm.itemno.value);
                }else{
				    AddData(frm.itemid.value,frm.itemoption.value,frm.itemname.value,frm.itemoptionname.value,makeriddisp ,printprice ,frm.itemgubun.value,frm.itemno.value);
				}
			}
		}
	}
	iaxobject.ShowFrm();
}

function onlyNumberInput()
{
	var code = window.event.keyCode;
	if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) {
		window.event.returnValue = true;
		return;
	}
	window.event.returnValue = false;
}

function reg(){

	frm.submit();
}

</script>

<OBJECT
	  id=iaxobject
	  classid="clsid:5D776FEA-8C6B-4C53-8EC3-3585FC040BDB"
	  codebase="http://webadmin.10x10.co.kr/common/cab/tenbarPrint.cab#version=1,0,0,29"
	  width=0
	  height=0
	  align=center
	  hspace=0
	  vspace=0
>
</OBJECT>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on" %>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">

		��ǰ�ڵ� : <input type="text" class="text" name="iitemid" value="<%= iitemid %>" maxlength="7" size="7" onKeyDown = "javascript:onlyNumberInput()" style="IME-MODE: disabled" />
		&nbsp;
		���ڵ� : <input type="text" class="text" name="barcode" value="<%= barcode %>" maxlength="14" size="14">
	<!--	&nbsp;
		�ֹ��ڵ� : <input type="text" class="text" name="" value="" maxlength="8" size="9">(�ڵ��ؾ���)
    -->
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		��ǰ����:<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
		&nbsp;
		��뿩��:
		<select class="select" name="isusingyn">
			<option value="">��ü</option>
			<option value="Y" <%= CHKIIF(isusingyn="Y","selected","") %> >�����</option>
		</select>
	 </td>
</tr>
</table>

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		�� ������ ���� :
		<!--�ݾ�ǥ�ÿ��� :
		<select name="printpriceyn">
			<option value="Y" <%' if (printpriceyn = "Y") then %>selected<%' end if %>>�ݾ�ǥ��</option>
			<option value="N" <%' if (printpriceyn = "N") then %>selected<%' end if %>>�ݾ�ǥ�þ���</option>
		</select>-->
		�귣��ǥ�� :
		<select name="makeriddispyn">
			<option value="Y" <% if (makeriddispyn = "Y") then %>selected<% end if %>>�귣��ǥ��</option>
			<option value="N" <% if (makeriddispyn = "N") then %>selected<% end if %>>�귣��ǥ�þ���</option>
		</select>


	</td>
	<td align="right">
	    ���� ���� 65ĭ �� : LA-3100,LB-3100 ��
		<% if obarcode.FResultCount>0 then %>
			<input type="button" class="button" value="���û�ǰ ���ڵ����" onclick="AddArr()">
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= FormatNumber(obarcode.FTotalCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% if obarcode.FResultCount > 0 then %>
	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<td width="100">��ǰ�ڵ�</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width="80">�ǸŰ�</td>
	<td width="80">��¼���</td>
</tr>
<% for i=0 to obarcode.FResultCount-1 %>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="itemid" value="<%= obarcode.FItemList(i).Fshopitemid %>">
<input type="hidden" name="itemoption" value="<%= obarcode.FItemList(i).Fitemoption %>">
<input type="hidden" name="itemname" value="<%= Replace(obarcode.FItemList(i).Fshopitemname,Chr(34),"") %>">
<input type="hidden" name="itemoptionname" value="<%= obarcode.FItemList(i).Fshopitemoptionname %>">
<input type="hidden" name="brand" value="<%= obarcode.FItemList(i).FSocName %>">
<input type="hidden" name="sellprice" value="<%= obarcode.FItemList(i).Fshopitemprice %>">
<input type="hidden" name="itemgubun" value="<%= obarcode.FItemList(i).Fitemgubun %>">
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><%= obarcode.FItemList(i).Fitemgubun %>-<%= CHKIIF(obarcode.FItemList(i).Fshopitemid>=1000000,Format00(8,obarcode.FItemList(i).Fshopitemid),Format00(6,obarcode.FItemList(i).Fshopitemid)) %>-<%= obarcode.FItemList(i).Fitemoption %></td>
	<td align="left"><%= obarcode.FItemList(i).Fshopitemname %></td>
	<td><%= obarcode.FItemList(i).Fshopitemoptionname %></td>
	<td align="right"><%= FormatNumber(obarcode.FItemList(i).Fshopitemprice,0) %></td>
<td><input type="text" class="text" name="itemno" value="1" maxlength="3" size="3"></td>
</tr>
</form>
<% next %>

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="10" align="center">�˻� ����� �����ϴ�.</td>
</tr>

<% end if %>
</table>

<%
set obarcode = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->