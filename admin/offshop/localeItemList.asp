<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������ ������ ��ǰ ����
' History : 2010.08.03 ������ ����
'			2010.08.05 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopLocaleItemcls.asp"-->

<%
dim designer,page,usingyn ,research,pricediff,imageview ,itemgubun, itemid, itemname , gubun
dim cdl, cdm, cds ,shopid , i ,PriceDiffExists , arrexchangerate, currencyUnit ,multipleRate , exchangeRate
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),9)
	usingyn     = RequestCheckVar(request("usingyn"),1)
	research    = RequestCheckVar(request("research"),9)
	pricediff   = RequestCheckVar(request("pricediff"),9)
	imageview   = RequestCheckVar(request("imageview"),9)
	
	itemgubun   = RequestCheckVar(request("itemgubun"),2)
	itemid      = RequestCheckVar(request("itemid"),9)
	itemname    = RequestCheckVar(request("itemname"),32)
	
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
	shopid      = RequestCheckVar(request("shopid"),32)
	gubun      = RequestCheckVar(request("gubun"),10)

''���� session("ssAdminPsn")="6" : �μ���ȣ�� ����Ұ�.
if session("ssBctDiv")="201" or session("ssAdminPsn")="6" then
	shopid = "cafe002"
elseif session("ssBctDiv")="301" or session("ssAdminPsn")="16" then
	shopid = "cafe003"
else    
end if

''���������� �������ΰ�� �ھ� �ִ´�
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

if page="" then page=1
if research<>"on" then usingyn="Y"

if shopid = "" then response.write "<script>alert('������ �����ϼ���');</script>"

dim ioffitem
set ioffitem  = new COffShopLocale
	ioffitem.FPageSize = 50
	ioffitem.FCurrPage = page
	ioffitem.FRectShopId = shopid
	ioffitem.FRectDesigner = designer
	ioffitem.FRectOnlyUsing = usingyn
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemID = itemid
	ioffitem.FRectItemName = html2db(itemname)
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.frectgubun = gubun
	'ioffitem.FRectOnlineExpiredItem = onexpire
	
	if (shopid<>"") then
	    ioffitem.GetLocaleItemList
	end if

dim oexchangerate
set oexchangerate = new COffShopLocale
	oexchangerate.frectuserid = shopid

if shopid <> "" then
	oexchangerate.fexchangeratecheck()

	if oexchangerate.foneitem.fcurrencyUnit = "" or isnull(oexchangerate.foneitem.fcurrencyUnit) then response.write "<script>alert('�ش���忡 ȭ������� ��ϵǾ� ���� �ʽ��ϴ�');</script>"
	if oexchangerate.foneitem.fmultipleRate = "" or isnull(oexchangerate.foneitem.fmultipleRate) then response.write "<script>alert('�ش���忡 ��������� ��ϵǾ� ���� �ʽ��ϴ�');</script>"

	currencyUnit = oexchangerate.foneitem.fcurrencyUnit
	multipleRate = oexchangerate.foneitem.fmultipleRate
	exchangeRate = oexchangerate.foneitem.fexchangeRate
end if
%>

<script language='javascript'>

// ����ǰ �߰� �˾�
function addnewItem(){
	var popup_item;
	popup_item = window.open("pop_localeItem_input.asp", "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popup_item.focus();
}

//ȯ������ ��� & ����
function popexchangerate(){
    var popexchangerate = window.open('/common/offshop/exchangerate/exchangerate.asp','popexchangerate','width=1024,height=768,scrollbars=yes,resizable=yes');
    popexchangerate.focus();
}

// ȯ�� ��� �ϰ�����
function automulti(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					
					//if (frm.lcprice.value==''){
					//	alert('�����ǸŰ� �������� �ʾҽ��ϴ�');
					//	frm.lcprice.focus;
					//	return;
					//}

					frm.lcprice.value = Math.round(((frm.ShopItemprice.value / upfrm.exchangeRate.value)* upfrm.multipleRate.value) * 1000) / 1000;								
				}
			}
		}
}

//ȯ���ϰ�����
function autoexchangeRate(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					
					if (frm.lcprice.value==''){
						alert('�����ǸŰ� �������� �ʾҽ��ϴ�');
						frm.lcprice.focus;
						return;
					}

					frm.lcprice.value = frm.lcprice.value / frm.exchangeRate.value;
					
						
				}
			}
		}
}

//��������ϰ�����
function automultipleRate(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					
					if (frm.lcprice.value==''){
						alert('�����ǸŰ� �������� �ʾҽ��ϴ�');
						frm.lcprice.focus;
						return;
					}

					frm.lcprice.value = frm.lcprice.value * upfrm.multipleRate.value;
					
						
				}
			}
		}
}

//�⺻�ǸŰ��ϰ�����
function autoShopItemprice(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					frm.lcprice.value = frm.ShopItemprice.value
						
				}
			}
		}
}

//�⺻��ǰ���ϰ�����
function autoShopItemName(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					frm.lcitemname.value = frm.ShopItemName.value						
				}
			}
		}
}

//�⺻�ɼǸ��ϰ�����
function autoshopitemoptionname(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					frm.lcitemoptionname.value = frm.shopitemoptionname.value						
				}
			}
		}
}

function ModiArr(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
				
					if (frm.lcitemname.value == ''){
						alert('��ǰ���� �Է����ּ���');
						frm.lcitemname.focus();
						return;
					}
					if (frm.lcprice.value == ''){
						alert('�ǸŰ��� �Է����ּ���');
						frm.lcprice.focus();
						return;
					}					
					upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "," ;
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "," ;
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "," ;
					upfrm.lcitemnamearr.value = upfrm.lcitemnamearr.value + frm.lcitemname.value + "," ;
					upfrm.lcitemoptionnamearr.value = upfrm.lcitemoptionnamearr.value + frm.lcitemoptionname.value + "," ;
					upfrm.lcpricearr.value = upfrm.lcpricearr.value + frm.lcprice.value + "," ;						
				}
			}
		}

		var itemidarr;
		var itemoptionarr;
		var itemgubunarr;
		var lcitemnamearr;
		var lcitemoptionnamearr;
		var lcpricearr;			
		itemidarr = upfrm.itemidarr.value;
		itemoptionarr = upfrm.itemoptionarr.value;
		itemgubunarr = upfrm.itemgubunarr.value;
		lcitemnamearr = upfrm.lcitemnamearr.value;
		lcitemoptionnamearr = upfrm.lcitemoptionnamearr.value;
		lcpricearr = upfrm.lcpricearr.value;
		upfrm.itemidarr.value = ""
		upfrm.itemoptionarr.value = ""
		upfrm.itemgubunarr.value = ""
		upfrm.lcitemnamearr.value = ""
		upfrm.lcitemoptionnamearr.value = ""
		upfrm.lcpricearr.value = ""
		
		var ModiArr;
		ModiArr = window.open('localeitem_process.asp?itemidarr='+itemidarr+'&itemoptionarr='+itemoptionarr+'&itemgubunarr='+itemgubunarr+'&lcitemnamearr='+lcitemnamearr+'&lcitemoptionnamearr='+lcitemoptionnamearr+'&lcpricearr='+lcpricearr+'&mode=localeitemadd&shopid=<%=shopid%>', "ModiArr","width=400,height=300,scrollbars=yes,resizable=yes");
		ModiArr.focus();
}

function reg(page){

	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('��ǰ�ڵ�� ���ڸ� �����մϴ�.');
			frm.itemid.focus();
			return;
		}
	}

	frm.page.value=page;
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="currencyUnit" value="<%= currencyUnit %>">
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemoptionarr">
<input type="hidden" name="itemgubunarr">
<input type="hidden" name="lcitemnamearr">
<input type="hidden" name="lcitemoptionnamearr">
<input type="hidden" name="lcpricearr">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	    ���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
	    �������:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:reg(<%=page%>);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�������� : <% drawlocaleitemgubun "gubun" , gubun , "" %>
		�귣�� : <% drawSelectBoxDesignerwithName "designer",designer %>
		��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9">		
		��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">     	     	
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >�̹�������
	</td>
</tr>
</table>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
�� ���庰 ȭ������� ����� [OFF]����_�������>>�����޸���Ʈ ���� �Է����ּ���.
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	    <% if currencyUnit <> "" and multipleRate <> "" then %>
	    	�ǸŰ� X ȯ��<input type="text" name="exchangeRate" value="<%= exchangeRate %>" size=5 maxlength=5>
	    	X ���<input type="text" name="multipleRate" value="<%= multipleRate %>" size=3 maxlength=3>
	    	<input type="button" class="button" value="���" onclick="automulti(frm)">
			&nbsp;<input type="button" class="button" value="�⺻��ǰ��" onclick="autoShopItemName(frm)">
			<input type="button" class="button" value="�⺻�ɼǸ�" onclick="autoshopitemoptionname(frm)">			
			&nbsp;<input type="button" class="button" value="�����ϰ�����" onclick="ModiArr(frm)">
			<!--<input type="button" class="button" value="�⺻�ǸŰ�����" onclick="autoShopItemprice(frm)">
			<input type="button" class="button" value="ȯ������" onclick="autoexchangeRate(frm)">
			<input type="button" class="button" value="�������(X<%= multipleRate %>)" onclick="automultipleRate(frm)">-->			
		<% end if %>
	</td>
	<td align="right">
		<!--<input type="button" value="����ǰ �߰�" onclick="addnewItem();" class="button">-->
		<input type="button" value="ȯ�� ����" class="button" onClick="popexchangerate();">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">		
		�˻���� : <b><%= ioffitem.FTotalcount %></b>
		&nbsp;
		<% if ioffitem.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>
		
		<b><%= page %> / <%= ioffitem.FTotalpage %></b>
		
		<% if (ioffitem.FTotalpage - ioffitem.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>
	</td>
</tr>	
<% if ioffitem.FresultCount > 0 then %>	
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<% if (imageview<>"") then %>
	<td>�̹���</td>
	<% end if %>
	<td>����<br>����</td>
	<td>�귣��ID<br>������ڵ�</td>
	<td>��ǰ�ڵ�<br>�ɼ��߰��ݾ�</td>
	<td>��ǰ��</font><br>������ǰ��</td>
	<td>�ɼǸ�</font><br>�����ɼǸ�</td>	
	<!--<td>���԰�</td>-->
	<td>�Һ��ڰ�(��)<br>�ǸŰ�(��)</td>
	<td>�����ǸŰ�<br>(<%= currencyUnit %>)</td>
	<!--<td>����ȯ��</td>
	<td>������<br>(%)</td>	
	<td>����<br>����</td>
	<td>����<br>����</td>
	<td>ON<br>�Ǹ�</td>
	<td>ON<br>����</td>-->
</tr>

<% for i=0 to ioffitem.FresultCount -1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">

<% if ioffitem.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>
	<td >
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		<input type="hidden" name="shopid" value="<%=shopid%>">
		<input type="hidden" name="itemid" value="<%=ioffitem.FItemlist(i).FShopitemid%>">
		<input type="hidden" name="itemoption" value="<%=ioffitem.FItemlist(i).Fitemoption%>">
		<input type="hidden" name="itemgubun" value="<%=ioffitem.FItemlist(i).fitemgubun%>">
	</td>
	<% if (imageview<>"") then %>
	<td><img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></td>
	<% end if %>
	<td>
		<%= ioffitem.FItemlist(i).fstatus %>
	</td>	
	<td>
		<%= ioffitem.FItemlist(i).FMakerID %>
		<br><%= ioffitem.FItemlist(i).FextBarcode %>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).Fitemgubun %>-<%=  Format00(6,ioffitem.FItemlist(i).Fshopitemid) %>-<%= ioffitem.FItemlist(i).Fitemoption %>
		<br>
		<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
		    <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>    
		<% end if %>				
	</td>
	<td>
		<%= ioffitem.FItemlist(i).FShopItemName %><input type="hidden" name="ShopItemName" value="<%= ioffitem.FItemlist(i).FShopItemName %>">
		<br><input type="text" name="lcitemname" value="<%= ioffitem.FItemlist(i).flcitemname %>" maxlength=123 size=30>
	</td>
	<td>
		<input type="hidden" name="shopitemoptionname" value="<%= ioffitem.FItemlist(i).fshopitemoptionname %>">
		<% if ioffitem.FItemlist(i).fshopitemoptionname <> "" then %>
			<%= ioffitem.FItemlist(i).FShopitemOptionname %>
			<br><input type="text" name="lcitemoptionname" value="<%= ioffitem.FItemlist(i).flcitemoptionname %>" maxlength=95 size=15>
		<% else %>
			<input type="hidden" name="lcitemoptionname" value="<%= ioffitem.FItemlist(i).flcitemoptionname %>" maxlength=95 size=15>
		<% end if %>		
	</td>	
    <% PriceDiffExists = false %>
	<!--<td><%'= FormatNumber(ioffitem.FItemlist(i).Fshopsuplycash,0) %></td>-->    
    <td>
        <%= FormatNumber(ioffitem.FItemlist(i).FShopItemOrgprice,0) %>
        <!--<%' if (ioffitem.FItemlist(i).FItemGubun="10") then %>
	        <%' if (ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemOrgprice)  then %>
	            <font color="red"><strong><%'= ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
	        <%' else %>
	            <%' if (PriceDiffExists) then %>
	            <%'= ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice %>
	            <%' end if %>
	        <%' end if %>
        <%' end if %>-->
		<br>
	    <%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %><input type="hidden" name="ShopItemprice" value="<%=ioffitem.FItemlist(i).FShopItemprice%>">
	    <!--<%' if (ioffitem.FItemlist(i).FItemGubun="10") then %>
	        <%' if (ioffitem.FItemlist(i).FOnLineItemprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemprice)  then %>
		        <font color="red"><strong><%'= ioffitem.FItemlist(i).FOnLineItemprice + ioffitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
		    <%' else %>
		        <%' if (PriceDiffExists) then %>
		        <%'= ioffitem.FItemlist(i).FOnLineItemprice + ioffitem.FItemlist(i).FOnlineOptaddprice %>
		        <%' end if %>
	        <%' end if %>
        <%' end if %>-->
    </td>
    <td>
		<input type="text" name="lcprice" value="<%= ioffitem.FItemlist(i).flcprice %>" size=5 maxlength=5>
    </td>    
    <!--<td>
		<input type="text" name="exchangeRate" value="<%'= ioffitem.FItemlist(i).fexchangeRate %>" size=5>
    </td>
	<td> 
        <%' if (ioffitem.FItemlist(i).FShopItemOrgprice<>0) then %>
            <%' if ioffitem.FItemlist(i).FShopItemOrgprice<>ioffitem.FItemlist(i).FShopItemprice then %>
            OFF:<font color="#FF3333"><%'= CLng((ioffitem.FItemlist(i).FShopItemOrgprice-ioffitem.FItemlist(i).FShopItemprice)/ioffitem.FItemlist(i).FShopItemOrgprice*100*100)/100 %></font>
            <%' end if %>
	    <%' end if %>
	    
	    <%' if (ioffitem.FItemlist(i).FOnlineitemorgprice<>0) then %>
	        <%' if ioffitem.FItemlist(i).FOnlineitemorgprice<>ioffitem.FItemlist(i).FOnLineItemprice then %>
            ON:<font color="#FF3333"><%'= CLng((ioffitem.FItemlist(i).FOnlineitemorgprice-ioffitem.FItemlist(i).FOnLineItemprice)/ioffitem.FItemlist(i).FOnlineitemorgprice*100*100)/100 %></font>
            <%' end if %>
	    <%' end if %>
	</td>
	<td>
	<%' if (ioffitem.FItemlist(i).FShopItemprice<>0) and (ioffitem.FItemlist(i).Fshopsuplycash<>0) then %>
		<font color="blue"><%'= CLng((ioffitem.FItemlist(i).FShopItemprice-ioffitem.FItemlist(i).Fshopsuplycash)/ioffitem.FItemlist(i).FShopItemprice*100) %>%</font>
	<%' end if %>
	</td>
	<td>
	<%' if (ioffitem.FItemlist(i).FShopItemprice<>0) and (ioffitem.FItemlist(i).Fshopbuyprice<>0) then %>
		<font color="blue"><%'= CLng((ioffitem.FItemlist(i).FShopItemprice-ioffitem.FItemlist(i).Fshopbuyprice)/ioffitem.FItemlist(i).FShopItemprice*100) %>%</font>
	<%' end if %>
    </td>
    <td><%'= fnColor(ioffitem.FItemlist(i).Fsellyn,"sellyn") %></td>
    <td><%'= fnColor(ioffitem.FItemlist(i).FonLineDanjongyn,"dj") %></td>-->
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if ioffitem.HasPreScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=ioffitem.StartScrollPage-1%>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ioffitem.StartScrollPage to ioffitem.StartScrollPage + ioffitem.FScrollCount - 1 %>
			<% if (i > ioffitem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ioffitem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ioffitem.HasNextScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% else %>

<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>

<%
	set ioffitem = nothing
	set oexchangerate = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->