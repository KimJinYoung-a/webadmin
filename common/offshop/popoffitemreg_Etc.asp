<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : pos ��ǰ����
' Hieditor : 2011.01.13 ������ ����
'			 2011.03.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopLocaleItemcls.asp"-->
<%
dim shopid ,makerid ,i ,oexchangerate, IsCommaValid ,oitem , itemid , itemoption,itemgubun
dim shopitemname,shopitemoptionname,orgsellprice,shopitemprice,	shopbuyprice,discountsellprice
dim centermwdiv,extbarcode,vatinclude ,isusing ,shopsuplycash ,exchangeRate ,currencyUnit
	makerid = requestCheckVar(request("makerid"),32)
	shopid  = requestCheckVar(request("shopid"),32)
	itemid = requestCheckVar(request("itemid"),10)
	itemoption = requestCheckVar(request("itemoption"),4)
	itemgubun = requestCheckVar(request("itemgubun"),2)

if (C_IS_SHOP) then
    if (LCASE(shopid)<>LCASE(C_STREETSHOPID)) or (shopid="") then
    
    ''�������� ��� �� ���̵� �ݵ�� �����Ǿ���.
        response.write "���� ���̵� �������� �ʾҽ��ϴ�. ������ ���ǿ��."
        dbget.close() : response.end
    end if
end if

IsCommaValid = false

set oexchangerate = new COffShopLocale
	oexchangerate.frectuserid = shopid
	
		if shopid <> "" then
			oexchangerate.fexchangeratecheck()

			if oexchangerate.ftotalcount > 0 then
				exchangeRate = oexchangerate.FOneItem.fexchangeRate
				currencyUnit = oexchangerate.foneitem.fcurrencyUnit
			end if					
			IsCommaValid = oexchangerate.foneitem.fcurrencyUnit<>"WON" and oexchangerate.foneitem.fcurrencyUnit<>"KRW" and oexchangerate.foneitem.fcurrencyUnit<>""
		end if
set oexchangerate = Nothing

set oitem = new COffShopItem
	oitem.frectitemid = itemid
	oitem.frectitemoption = itemoption
	oitem.frectitemgubun = itemgubun
	
	if itemid <> "" and itemoption <> "" and itemgubun <> "" then
		oitem.GetOffNOnLineShoponeItem
	end if

	if oitem.ftotalcount > 0 then
		itemgubun = oitem.FOneItem.Fitemgubun
		itemid = oitem.FOneItem.Fshopitemid
		itemoption = oitem.FOneItem.Fitemoption
		makerid = oitem.FOneItem.Fmakerid
		shopitemname = oitem.FOneItem.Fshopitemname
		shopitemoptionname = oitem.FOneItem.Fshopitemoptionname
		orgsellprice = oitem.FOneItem.FShopItemOrgprice
		shopitemprice = oitem.FOneItem.Fshopitemprice
		shopbuyprice = oitem.FOneItem.fshopbuyprice
		discountsellprice = oitem.FOneItem.fdiscountsellprice
		centermwdiv = oitem.FOneItem.fcentermwdiv
		extbarcode = oitem.FOneItem.fextbarcode
		vatinclude = oitem.FOneItem.fvatinclude
	end if

if centermwdiv = "" then centermwdiv = "M"
if shopbuyprice = "" then shopbuyprice = "0"
if itemgubun = "" then itemgubun = "00"
if isusing = "" then isusing = "Y"
if vatinclude = "" then vatinclude = "Y"
if shopsuplycash = "" then shopsuplycash = "0"	
%>

<script language='javascript'>

function CheckAddItem(frm ,mode){
	
	if (frm.makerid.value.length<1){
		alert('�귣�带 �����ϼ���.');
		return;
	}

	if (frm.shopitemname.value.length<1){
		alert('��ǰ���� �Է��ϼ���.');
		frm.shopitemname.focus();
		return;
	}

	if ((frm.extbarcode.value.length>0) && (frm.extbarcode.value.length<10)){
		alert('���ڵ� ���̰� �ʹ� ª���ϴ�. ���� ���ڵ尡 �ִ°�츸 �Է��� �ּ���' );
		frm.extbarcode.focus();
		return;
	}

	if (!<%= CHKIIF(IsCommaValid,"IsDouble","IsDigit") %>(frm.shopitemprice.value)){
		alert('�ǸŰ��� ���ڸ� �����մϴ�.');
		frm.shopitemprice.focus();
		return;
	}

	if (!<%= CHKIIF(IsCommaValid,"IsDouble","IsDigit") %>(frm.shopsuplycash.value)){
		alert('��ü ���԰��� ���ڸ� �����մϴ�.');
		frm.shopsuplycash.focus();
		return;
	}

	if (!<%= CHKIIF(IsCommaValid,"IsDouble","IsDigit") %>(frm.shopbuyprice.value)){
		alert('�� ���ް��� ���ڸ� �����մϴ�.');
		frm.shopbuyprice.focus();
		return;
	}

	if (((frm.shopsuplycash.value!=0)||(frm.shopbuyprice.value!=0))){
		if (!confirm('!! �⺻ ��� ������ �ٸ� ��쿡�� ���԰� ���ް��� �Է� �ϼž� �մϴ�. \n\n��� �Ͻðڽ��ϱ�?')){
			return;
		}
	}

	var ret = confirm('�߰��Ͻðڽ��ϱ�?');

	if (ret) {
		frm.mode.value=mode;
		frm.submit();
	}
}

</script>

<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#FFFFFF>
<tr>
	<td>&gt;&gt;���� POS�����ǰ ���</td>
</tr>
</table>
		
<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#3d3d3d>
<form name="frmedit" method="post" action="/common/offshop/popoffitemreg_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="isforeignshop" value="<%= chkIIF(IsCommaValid,"on","") %>">
<input type="hidden" name="centermwdiv" value="<%= centermwdiv %>">
<input type="hidden" name="shopbuyprice" value="<%= shopbuyprice %>">
<tr bgcolor="#FFDDDD">
	<td width=100>�귣�� ID</td>
	<td bgcolor="#FFFFFF" colspan=5><% FnDrawOptPosBrand shopid,"makerid",makerid %>
	</td>
</tr>
<% if makerid<>"" then %>
<tr bgcolor="#DDDDFF">
	<td width=100>��ǰ����</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="radio" name="itemgubun" value="<%= itemgubun %>" checked >POS�����ǰ &nbsp;
	<input type="hidden" name="itemid" value="<%= itemid %>">
	<input type="hidden" name="itemoption" value="<%= itemoption %>">
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>��ǰ��</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=text name="shopitemname" value="<%=shopitemname%>" size=40 maxlength=40 class="input_01" >
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>������ڵ�</td>
	<td bgcolor="#FFFFFF" colspan=5><input type="text" name="extbarcode" value="<%= extbarcode %>" size=20 maxlength=20 class="input_01" >(�ִ� ��츸 ���)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>�������</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="radio" name="isusing" value="Y" <% if isusing = "Y" then response.write " checked" %>>�����
	<input type="radio" name="isusing" value="N" <% if isusing = "N" then response.write " checked" %>>������
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td width=100 >��������</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<input type="radio" name="vatinclude" value="Y" <% if vatinclude = "Y" then response.write " checked" %>>����
		<input type="radio" name="vatinclude" value="N" <% if vatinclude = "N" then response.write " checked" %>>�鼼
	</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td width=100 align="left" rowspan="3">
		���ݼ���
		<% if currencyUnit <> "" then %>
			(<%= currencyUnit %>)
		<% end if %>
	</td>
	<td bgcolor="#FFFFFF" >�ǸŰ�</td>
	<td bgcolor="#FFFFFF" >���԰�</td>	
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td bgcolor="#FFFFFF"><input type="text" name="shopitemprice" value="<%=shopitemprice%>" size=8 maxlength=9 class="input_right" ></td>
	<td bgcolor="#FFFFFF"><input type="text" name="shopsuplycash" value="<%= shopsuplycash %>" size=8 maxlength=9 class="input_right" ></td>	
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td bgcolor="#FFFFFF" ></td>
	<td bgcolor="#FFFFFF" colspan="2">(0�ΰ�� �⺻���� ���� ������)</td>
</tr>
</form>
<tr bgcolor="#FFFFFF">
	<td colspan=6 align="center">
		<% if itemid <> "" and itemoption <> "" and itemgubun <> "" then %>
			<input type=button value="����" onclick="CheckAddItem(frmedit,'editetcoffitem')" class="input_01">
		<% else %>	
			<input type=button value="�ű�����" onclick="CheckAddItem(frmedit,'addetcoffitem')" class="input_01">
		<% end if %>
	</td>
</tr>
<% end if %>
</table>

<%
set oitem = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
