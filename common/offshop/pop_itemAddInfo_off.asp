<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2010.12.02 �ѿ�� ����
' Description : ��ǰ �߰�
'				input - actionURL(db ó���� �ʿ��� �Ķ���ͱ��� ����) ex.acURL = "/admin/eventmanage/event/eventitem_process.asp?eC=1234"
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%

'// deprecated, 2018-02-14, skyer9, pop_itemAddInfoOFF.asp �� ������ ��!!!

dim designer,page,usingyn ,research,pricediff,imageview, pricelow , defaultmargin
dim itemgubun, itemid, itemname ,cdl, cdm, cds ,onexpire ,shopid , strparm
dim i, PriceDiffExists , actionURL , saleflg
	designer    = RequestCheckVar(request.form("designer"),32)
	page        = RequestCheckVar(request.form("page"),10)
	usingyn     = RequestCheckVar(request.form("usingyn"),1)
	research    = RequestCheckVar(request.form("research"),9)
	pricediff   = RequestCheckVar(request.form("pricediff"),9)
	pricelow    = RequestCheckVar(request.form("pricelow"),9)
	imageview   = RequestCheckVar(request.form("imageview"),9)
	onexpire    = RequestCheckVar(request.form("onexpire"),9)
	itemgubun   = RequestCheckVar(request.form("itemgubun"),2)
	'itemid      = RequestCheckVar(request.form("itemid"),9)
	itemid      = request.form("itemid")
	itemname    = RequestCheckVar(request.form("itemname"),32)
	cdl         = RequestCheckVar(request.form("cdl"),3)
	cdm         = RequestCheckVar(request.form("cdm"),3)
	cds         = RequestCheckVar(request.form("cds"),3)
	shopid    = RequestCheckVar(request("shopid"),32)
	saleflg    = RequestCheckVar(request("saleflg"),10)
	actionURL	= request("acURL")
	defaultmargin = RequestCheckVar(request("defaultmargin"),20)
	'response.write actionURL

	if shopid = "" then
		response.write "<script type='text/javascript'>alert('��ID �� �����ϴ�'); self.close();</script>"
		response.end
	end if

	'if sellyn = "" then sellyn ="Y"
	If itemid<>"" then
		Dim iA, arrTemp, arrItemid
		itemid = replace(itemid,",",chr(10))
		itemid = replace(itemid,chr(13),"")
		arrTemp = Split(itemid,chr(10))
		iA = 0
		Do While iA <= ubound(arrTemp)
			If Trim(arrTemp(iA))<>"" then
				If Not(isNumeric(trim(arrTemp(iA)))) then
					Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
					dbget.close()	:	response.End
				Else
					arrItemid = arrItemid & trim(arrTemp(iA)) & ","
				End If
			End If
			iA = iA + 1
		Loop
		itemid = left(arrItemid,len(arrItemid)-1)
	End If

	if page="" then page=1
	if research<>"on" then usingyn="Y"
	strparm = "designer="&designer&"&usingyn="&usingyn&""
	strparm = strparm & "&research="&research&"&pricediff="&pricediff&"&pricelow="&pricelow&"&imageview="&imageview&"&onexpire="&onexpire&""
	strparm = strparm & "&itemgubun="&itemgubun&"&itemid="&itemid&"&itemname="&itemname&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&saleflg="&saleflg

dim oitem
set oitem  = new COffShopItem
	oitem.FPageSize = 50
	oitem.FCurrPage = page
	oitem.FRectDesigner = designer
	oitem.frectsaleflg = saleflg
	oitem.frectshopid = shopid
	oitem.FRectOnlyUsing = usingyn
	oitem.FRectItemgubun = itemgubun
	oitem.FRectItemID = itemid
	oitem.FRectItemName = html2db(itemname)
	oitem.FRectCDL = cdl
	oitem.FRectCDM = cdm
	oitem.FRectCDS = cds
	oitem.FRectOnlineExpiredItem = onexpire

	if pricediff="on" then
	    oitem.FRectPriceRow = pricelow
		oitem.GetcontractOffShopPriceDiffItemList()
	else
		oitem.GetcontractShopItemList()
	end if
%>

<script type='text/javascript'>

function jsSerach(page){
	var frm;
	frm = document.frm;

	frm.target = "";
	frm.action = "";
	frm.page.value=page;
	frm.submit();
}


function SelectItems(){
var frm;
var frmA;
var itemcount = 0;
frm = document.frmItem;
frmA = document.frmAct;
	if(typeof(frm.chkitem) !="undefined"){
		if(!frm.chkitem.length){
			if(!frm.chkitem.checked){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}

			frmA.itemidarr.value = frmA.itemidarr.value + frm.chkitem.value + ",";
			frmA.itemgubunarr.value = frmA.itemgubunarr.value + frm.chkitemgubun.value + ",";
			frmA.itemoptionarr.value = frmA.itemoptionarr.value + frm.chkitemoption.value + ",";
			frmA.itemnoarr.value = frmA.itemnoarr.value + frm.chkitemno.value + ",";
			frmA.orgsellpricearr.value = frmA.orgsellpricearr.value + frm.chkorgsellprice.value + ",";
			frmA.sellcasharr.value = frmA.sellcasharr.value + frm.chksellcash.value + ",";
			frmA.shopsuplycasharr.value = frmA.shopsuplycasharr.value + frm.chkshopsuplycash.value + ",";
			frmA.shopbuypricearr.value = frmA.shopbuypricearr.value + frm.chkshopbuyprice.value + ",";
			//frm.itemnamearr.value = frm.itemnamearr.value + frm.chkitemname.value + ",";
			//frm.itemoptionnamearr.value = frm.itemoptionnamearr.value + frm.chkitemoptionname.value + ",";
			//frm.makeridarr.value = frm.makeridarr.value + frm.chkmakerid.value + ",";
			itemcount = 1;
		}else{
			for(i=0;i<frm.chkitem.length;i++){
				if(frm.chkitem[i].checked) {
					frmA.itemidarr.value = frmA.itemidarr.value + frm.chkitem[i].value + ",";
					frmA.itemgubunarr.value = frmA.itemgubunarr.value + frm.chkitemgubun[i].value + ",";
					frmA.itemoptionarr.value = frmA.itemoptionarr.value + frm.chkitemoption[i].value + ",";
					frmA.orgsellpricearr.value = frmA.orgsellpricearr.value + frm.chkorgsellprice[i].value + ",";
					frmA.sellcasharr.value = frmA.sellcasharr.value + frm.chksellcash[i].value + ",";
					frmA.itemnoarr.value = frmA.itemnoarr.value + frm.chkitemno[i].value + ",";
					frmA.shopsuplycasharr.value = frmA.shopsuplycasharr.value + frm.chkshopsuplycash[i].value + ",";
					frmA.shopbuypricearr.value = frmA.shopbuypricearr.value + frm.chkshopbuyprice[i].value + ",";
				//	frm.itemnamearr.value = frm.itemnamearr.value + frm.chkitemname[i].value + ",";
				//	frm.itemoptionnamearr.value = frm.itemoptionnamearr.value + frm.chkitemoptionname[i].value + ",";
				//	frm.makeridarr.value = frm.makeridarr.value + frm.chkmakerid[i].value + ",";
				}
				itemcount = frm.chkitem.length;
			}

			if (frmA.itemidarr.value == ""){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
		}
	}else{
		alert("�߰��� ��ǰ�� �����ϴ�.");
	return;
	}

	//frm.target = opener.name;
	frmA.target = "FrameCKP";
	frmA.action = "<%=actionURL%>";
	frmA.itemcount.value = itemcount;
	frmA.submit();
	frmA.itemnoarr.value = "";
	frmA.itemidarr.value = "";
	frmA.itemgubunarr.value = "";
	frmA.itemoptionarr.value = "";
	frmA.orgsellpricearr.value = "";
	frmA.sellcasharr.value = "";
	frmA.shopsuplycasharr.value = "";
	frmA.shopbuypricearr.value = "";
//	frm.itemnamearr.value = "";
//	frm.itemoptionnamearr.value = "";
//	frm.makeridarr.value = "";
	frmA.itemcount.value = 0;
	opener.history.go(0);
	//window.close();
}

//��ü ����
function jsChkAll(){
var frm;
frm = document.frmItem;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
		   	 	frm.chkitem.checked = true;
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					frm.chkitem[i].checked = true;
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}
		}
	  }

	}

}

</script>

<!-- �˻� ���� -->
<form name="frm" method="post">
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="page">
<input type="hidden" name="acURL" value="<%=actionURL%>">
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣�� : <% drawSelectBoxDesignerwithName "designer",designer  %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
	</td>

	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach('');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		��ǰ����:<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
     	&nbsp;
     	�������:<% drawSelectBoxUsingYN "usingyn", usingyn %>
     	<br>
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >�̹�������
		&nbsp;
		<input type="checkbox" name="pricediff" value="on" <% if pricediff="on" then response.write "checked" %> >���ݻ��̸� ����
		&nbsp;
		<input type="checkbox" name="pricelow" value="on" <% if pricelow="on" then response.write "checked" %> >�¶��κ��� ��������
		&nbsp;
		<input type="checkbox" name="onexpire" value="on" <% if onexpire="on" then response.write "checked" %> >ONǰ��+����+������(�Ż�ǰ����)
		<input type="checkbox" name="saleflg" value="on" <% if saleflg="on" then response.write "checked" %> >���λ�ǰ���ܾ���
	</td>
</tr>
</table>
</form>
<form name="frmAct" method="post">
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="page">
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemoptionarr">
<input type="hidden" name="itemgubunarr">
<input type="hidden" name="itemnoarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="orgsellpricearr">
<input type="hidden" name="sellcasharr">
<input type="hidden" name="shopsuplycasharr">
<input type="hidden" name="shopbuypricearr">
<!--<input type="hidden" name="itemnamearr">
<input type="hidden" name="itemoptionnamearr">
<input type="hidden" name="makeridarr">-->
<input type="hidden" name="mode" value="I">
<input type="hidden" name="acURL" value="<%=actionURL%>">
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
</form>
<form name="frmItem" method="post">
<input type="hidden" name="research" value="on">
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="page">
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
<tr height="40" valign="bottom">
    <td align="left">
    	�ؘ�(<%=shopid%>)�� ���� ��ǰ�� ���� �Ǹ�, ���԰� & �ް��ް� ������ ���°��, �ش� ���� ��� �⺻���� ���� ���˴ϴ�.
    </td>
    <td align="right">
    	<input type="button" value="���û�ǰ �߰�" onClick="SelectItems()" class="button">
    </td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr  bgcolor="#FFFFFF">
	<td colspan="20">
	�˻���� : <b><%= oitem.FTotalCount%></b>
	&nbsp;
	������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<% if (imageview<>"") then %>
	<td>�̹���</td>
	<% end if %>
	<td>�귣��ID</td>
	<td>
		��������<br>
		<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
		�⺻���Ը���<br>
		<% end if %>
		�⺻�ް��޸���
	</td>
	<td>��ǰ�ڵ�<br>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	<td>�Һ��ڰ�</td>
	<td>�ǸŰ�</td>
	<td>������<br>(%)</td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
    	<td>���԰�</td>
    <% end if %>
	<td>�ް��ް�</td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
	<td>����<br>����</td>
	<% end if %>
	<td>����<br>����</td>
	<td>����<br>����<br>����</td>
	<td>ON<br>�Ǹ�</td>
	<td>ON<br>����</td>
	<td>������ڵ�</td>
</tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF" >
    	<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF" align="center">
	<td>
		<input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).Fshopitemid %>">
		<input type="hidden" name="chkitemoption" value="<%= oitem.FItemList(i).Fitemoption %>">
		<input type="hidden" name="chkitemgubun" value="<%= oitem.FItemList(i).Fitemgubun %>">
		<input type="hidden" name="chkitemno" value="0">
		<input type="hidden" name="chkorgsellprice" value="<%= oitem.FItemlist(i).FShopItemOrgprice %>">
		<input type="hidden" name="chksellcash" value="<%= oitem.FItemlist(i).FShopItemprice %>">
		<input type="hidden" name="chkshopsuplycash" value="<%= oitem.FItemList(i).Fshopsuplycash %>">
		<input type="hidden" name="chkshopbuyprice" value="<%= oitem.FItemList(i).Fshopbuyprice %>">
		<input type="hidden" name="chkitemname" value="<%= oitem.FItemList(i).fshopitemname %>">
		<input type="hidden" name="chkitemoptionname" value="<%= oitem.FItemList(i).fshopitemoptionname %>">
		<input type="hidden" name="chkmakerid" value="<%= oitem.FItemList(i).FMakerID %>">

	<% if (imageview<>"") then %>
	<td><img src="<%= oitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></td>
	<% end if %>
	<td>
		<%= oitem.FItemlist(i).FMakerID %>
	</td>
	<td>
		<%= oitem.FItemList(i).getJungsanDivName %>
		<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
		<br><%= oitem.FItemlist(i).fdefaultmargin %>%
		<% end if %>
		<br><%= oitem.FItemlist(i).fdefaultsuplymargin %>%
	</td>
	<td>
		<%= oitem.FItemlist(i).Fitemgubun %><%= CHKIIF(oitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,oitem.FItemlist(i).Fshopitemid),Format00(6,oitem.FItemlist(i).Fshopitemid)) %><%= oitem.FItemlist(i).Fitemoption %>
		<br><%= oitem.FItemlist(i).FShopItemName %>
		<% if oitem.FItemlist(i).Fitemoption<>"0000" then %>
			<font color="blue">[<%= oitem.FItemlist(i).FShopitemOptionname %>]</font>
		<% end if %>
		<% if oitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
		    <br>�ɼ��߰��ݾ�: <%= FormatNumber(oitem.FItemlist(i).FOnlineOptaddprice,0) %>
		<% end if %>
	</td>
    <% PriceDiffExists = false %>
    <td align="right" >
        <%= FormatNumber(oitem.FItemlist(i).FShopItemOrgprice,0) %>
        <% if (oitem.FItemlist(i).FItemGubun="10") then %>
        <% if (oitem.FItemlist(i).FOnlineitemorgprice + oitem.FItemlist(i).FOnlineOptaddprice<>oitem.FItemlist(i).FShopItemOrgprice)  then %>
            <font color="red"><strong><%= oitem.FItemlist(i).FOnlineitemorgprice + oitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
        <% else %>
            <% if (PriceDiffExists) then %>
            <%= oitem.FItemlist(i).FOnlineitemorgprice + oitem.FItemlist(i).FOnlineOptaddprice %>
            <% end if %>
        <% end if %>
        <% end if %>
    </td>
	<td align="right" >
	    <%= FormatNumber(oitem.FItemlist(i).FShopItemprice,0) %>
	    <% if (oitem.FItemlist(i).FItemGubun="10") then %>
        <% if (oitem.FItemlist(i).FOnLineItemprice+ oitem.FItemlist(i).FOnlineOptaddprice<>oitem.FItemlist(i).FShopItemprice)  then %>
	        <font color="red"><strong><%= oitem.FItemlist(i).FOnLineItemprice + oitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
	    <% else %>
	        <% if (PriceDiffExists) then %>
	        <%= oitem.FItemlist(i).FOnLineItemprice + oitem.FItemlist(i).FOnlineOptaddprice %>
	        <% end if %>
        <% end if %>
        <% end if %>
	</td>
	<td align="center" >
        <% if (oitem.FItemlist(i).FShopItemOrgprice<>0) then %>
            <% if oitem.FItemlist(i).FShopItemOrgprice<>oitem.FItemlist(i).FShopItemprice then %>
            OFF:<font color="#FF3333"><%= CLng((oitem.FItemlist(i).FShopItemOrgprice-oitem.FItemlist(i).FShopItemprice)/oitem.FItemlist(i).FShopItemOrgprice*100*100)/100 %></font>
            <% end if %>
	    <% end if %>

	    <% if (oitem.FItemlist(i).FOnlineitemorgprice<>0) then %>
	        <% if oitem.FItemlist(i).FOnlineitemorgprice<>oitem.FItemlist(i).FOnLineItemprice then %>
            ON:<font color="#FF3333"><%= CLng((oitem.FItemlist(i).FOnlineitemorgprice-oitem.FItemlist(i).FOnLineItemprice)/oitem.FItemlist(i).FOnlineitemorgprice*100*100)/100 %></font>
            <% end if %>
	    <% end if %>
	</td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
		<td align="right" ><%= FormatNumber(oitem.FItemlist(i).Fshopsuplycash,0) %></td>
	<% end if %>
	<td align="right" ><%= FormatNumber(oitem.FItemlist(i).Fshopbuyprice,0) %></td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
	<td align="right" >
	<% if (oitem.FItemlist(i).FShopItemprice<>0) and (oitem.FItemlist(i).Fshopsuplycash<>0) then %>
		<font color="blue"><%= CLng((oitem.FItemlist(i).FShopItemprice-oitem.FItemlist(i).Fshopsuplycash)/oitem.FItemlist(i).FShopItemprice*100) %>%</font>
	<% end if %>
	</td>
	<% end if %>
	<td align="right" >
	<% if (oitem.FItemlist(i).FShopItemprice<>0) and (oitem.FItemlist(i).Fshopbuyprice<>0) then %>
		<font color="blue"><%= CLng((oitem.FItemlist(i).FShopItemprice-oitem.FItemlist(i).Fshopbuyprice)/oitem.FItemlist(i).FShopItemprice*100) %>%</font>
	<% end if %>
    </td>
    <td align="center" ><%= oitem.FItemlist(i).FCenterMwDiv %></td>
    <td align="center" ><%= fnColor(oitem.FItemlist(i).Fsellyn,"sellyn") %></td>
    <td align="center" ><%= fnColor(oitem.FItemlist(i).FonLineDanjongyn,"dj") %></td>
	<td align="right" ><%= oitem.FItemlist(i).FextBarcode %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
       	<% if oitem.HasPreScroll then %>
			<span class="list_link"><a href="javascript:jsSerach(<%=oitem.StartScrollPage-1%>);">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oitem.StartScrollPage to oitem.StartScrollPage + oitem.FScrollCount - 1 %>
			<% if (i > oitem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oitem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:jsSerach(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oitem.HasNextScroll then %>
			<span class="list_link"><a href="javascript:jsSerach(<%=i%>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% end if %>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
<tr height="40" valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    	<input type="button" value="���û�ǰ �߰�" onClick="SelectItems()" class="button">
    </td>
</tr>
</table>
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="800" height="100"></iframe>
<%
 set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
