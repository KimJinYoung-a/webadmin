<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/kaffa/itemsalecls.asp"-->
<%
Dim clsSale, page, i
Dim discountKey, discountTitle, promotionType, stDT, edDT, discountPro, discountbuyRule, discountbuyPro, regdate, lastupdate, openDate, expiredDate, regUserID, lastUpUserID, discountStatus
Dim mSPrice, mSBPrice, iSaleMargin, iOrgMargin, iSalePercent, smargin
Dim discountbuyRuleStr
Dim preref
if (InStr(LCASE(request.ServerVariables("HTTP_REFERER")),"admin/etc/kaffa/sale/salelist.asp")>0) then
    preref = request.ServerVariables("HTTP_REFERER")
else
    preref = request("preref")
end if
if (preref<>"") and (preref<>session("preref")) then
    session("preref")=preref
end if

discountKey = request("discountKey")

Dim acURL
acURL =Server.HTMLEncode("/admin/etc/kaffa/sale/saleitemProc.asp?discountKey="&discountKey)

'(��) �������¿� ���� ���԰� ����-------------------------------------------------------
Function fnSetSaleSupplyPrice(ByVal MarginType, ByVal MarginValue, ByVal orgPrice, ByVal orgSupplyPrice, ByVal salePrice)
	Dim orgMRate
	if orgPrice <>0 then '�� ������
		orgMRate = 100-fix(orgSupplyPrice/orgPrice*10000)/100
	end if

	SELECT CASE MarginType
		Case 1	'���ϸ���
			fnSetSaleSupplyPrice = salePrice- fix(salePrice*(orgMRate/100))
		Case 2	'��ü�δ�
			fnSetSaleSupplyPrice = salePrice-(orgPrice-orgSupplyPrice)
		Case 3	'�ݹݺδ�
			fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
		Case 4	'10x10�δ�
			fnSetSaleSupplyPrice = orgSupplyPrice
		Case 5	'��������
			fnSetSaleSupplyPrice = salePrice - fix(salePrice*(MarginValue/100))
	END SELECT
End Function
'-----------------------------------------------------------------------------------


Set clsSale = new CSale
	clsSale.FRectDiscountKey = discountKey
	clsSale.fnGetSaleConts

    discountTitle		= clsSale.FoneItem.FDiscountTitle
	promotionType		= clsSale.FoneItem.FPromotionType
	stDT				= clsSale.FoneItem.FStDT
	edDT				= clsSale.FoneItem.FEdDT
	discountPro			= clsSale.FoneItem.FDiscountPro
	discountbuyRule		= clsSale.FoneItem.FDiscountbuyRule
	discountbuyPro		= clsSale.FoneItem.FDiscountbuyPro
	regdate				= clsSale.FoneItem.FRegdate
	lastupdate			= clsSale.FoneItem.FLastupdate
	openDate			= clsSale.FoneItem.FOpenDate
	expiredDate			= clsSale.FoneItem.FExpiredDate
	regUserID			= clsSale.FoneItem.FRegUserID
	lastUpUserID		= clsSale.FoneItem.FLastUpUserID
	discountStatus		= clsSale.FoneItem.getDiscountStatus
	discountbuyRuleStr  = clsSale.FoneItem.getRuleStr
Set clsSale = nothing

page 		= request("page")
If page = "" Then page = 1

Set clsSale = new CSale
	clsSale.FRectDiscountKey = discountKey
	clsSale.FCurrPage	= page
	clsSale.FPageSize	= 30
	clsSale.fnGetSaleItemList
%>
<script language='javascript'>
function goPage(p){
    location.href="?discountKey=<%=discountKey%>&page="+p+"&menupos=<%=menupos%>";
}

function addnewItem(){
	var popwin;
		popwin = window.open("/admin/etc/kaffa/sale/itemlist.asp?sitename=CHNWEB&discountKey=<%=discountKey%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		popwin.focus();
}

function CkDisPrice(){
	CkDisOrOrg(true);
}

function CkOrgPrice(){
	CkDisOrOrg(false);
}

//���� ���ΰ� ����
function CkDisOrOrg(isDisc){
	var frm;
	var pass = false;

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


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if(isDisc==true){
					frm.iDSPrice.value = frm.saleprice.value;
					frm.iDBPrice.value = frm.salesupplyprice.value;
					frm.iDSMargin.value= frm.salemargin.value;
					frm.saleItemStatus.value = 7;
			}else{
					frm.iDSPrice.value = frm.orgPrice.value;
					frm.iDBPrice.value = frm.orgSupplyPrice.value;
					frm.iDSMargin.value = Math.round(((frm.iDSPrice.value-frm.iDBPrice.value)/frm.iDSPrice.value*1.0)*100*100)/100;
					frm.saleItemStatus.value = 9;
				}
			}
		}
	}
}
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

//���û�ǰ ����(��������)
function delArr(){
	var frm;
	var pass = false;
	var ovPer = 0;

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

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
			frmdel.itemid.value = frmdel.itemid.value + frm.itemid.value + ","
			}
		}
	}

	var ret = confirm('���� ��ǰ�� ����(��������) �Ͻðڽ��ϱ�?');

	if (ret){
		frmdel.submit();
	}
}

//���û�ǰ ����
function saveArr(){
	var frm;
	var pass = false;
	var ovPer = 0;

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

	frmarr.itemid.value = "";
	frmarr.sailyn.value = "";
	frmarr.iDSPrice.value ="";
	frmarr.iDBPrice.value ="";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				//check Not AvaliValue
				if (!IsDigit(frm.iDSPrice.value)){
					alert('���ڸ� �����մϴ�.');
					frm.iDSPrice.focus();
					return;
				}

				if (frm.iDSPrice.value<1){
					alert('�ݾ��� ��Ȯ�� �Է��ϼ���.');
					frm.iDSPrice.focus();
					return;
				}

				if (!IsDigit(frm.iDBPrice.value)){
					alert('���ڸ� �����մϴ�.');
					frm.iDBPrice.focus();
					return;
				}

				if (frm.iDBPrice.value<1){
					alert('�ݾ��� ��Ȯ�� �Է��ϼ���.');
					frm.iDBPrice.focus();
					return;
				}

				if(Math.round((frm.orgPrice.value-frm.iDSPrice.value)/frm.orgPrice.value*100)>=50) {
					ovPer++;
				}
				frmarr.itemid.value = frmarr.itemid.value + frm.itemid.value + ","
				frmarr.iDSPrice.value = frmarr.iDSPrice.value + frm.iDSPrice.value + ","
				frmarr.iDBPrice.value = frmarr.iDBPrice.value + frm.iDBPrice.value + ","
				frmarr.saleItemStatus.value = frmarr.saleItemStatus.value + frm.saleItemStatus.value+","
			}
		}
	}

	if(ovPer>0) {
		if(!confirm('!!!\n\n\n���� ��ǰ�߿� �������� �ſ� ���� ��ǰ(50%+)�� �ֽ��ϴ�!\n\n�Է��Ͻ� ������ �½��ϱ�?\n\n')) {
			return;
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frmarr.submit();
	}
}
// ������ ����
function reCALbyPrice(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSPrice.value>0) {
		frm.iDSMargin.value = Math.round(((frm.iDSPrice.value-frm.iDBPrice.value)/frm.iDSPrice.value*1.0)*100*100)/100;
	} else {
		frm.iDSMargin.value = 0;
	}

	//������ ǥ��
	var iorgPrice = frm.orgPrice.value;
	var isailprice = frm.iDSPrice.value;
	var isalePercent = Math.round((iorgPrice-isailprice)/iorgPrice*100);

	if(isalePercent>=50) {
		document.getElementById("lyrSpct"+fid).style.color="#EE0000";
		document.getElementById("lyrSpct"+fid).style.fontWeight="bold";
	} else {
		document.getElementById("lyrSpct"+fid).style.color="#000000";
		document.getElementById("lyrSpct"+fid).style.fontWeight="normal";
	}
	document.getElementById("lyrSpct"+fid).innerHTML = isalePercent + "%";
}
// ���԰� ����
function reCALbyMargin(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSMargin.value>0) {
		frm.iDBPrice.value = Math.round(frm.iDSPrice.value*(1-(frm.iDSMargin.value/100)));
	} else {
		frm.iDBPrice.value = frm.iDSPrice.value;
	}
}
</script>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="a">
<tr>
	<td width="100%">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="80">�����ڵ�</td>
			<td bgcolor="#FFFFFF" width="60"><%=discountKey%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"  width="80">���θ�</td>
			<td bgcolor="#FFFFFF"  width="150"><%=discountTitle%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="80">����</td>
			<td bgcolor="#FFFFFF"  width="60"><%%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="80">�Ⱓ</td>
			<td bgcolor="#FFFFFF" ><%=stDT%> ~ <%=edDT%></td>
			<td bgcolor="#FFFFFF" align="right" width="100"><input type="button" class="button" value=" List " onClick="location.href='<%=session("preref")%>'"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border=0>
		<form name=frmdummi>
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<tr height="40" valign="bottom">
			<td align="left"><input type=button value="���û�ǰ����" onClick="saveArr()" class="button">
			<input type=button value="���û�ǰ����(��������)" onClick="delArr()" class="button">
			</td>
			<td align="right">
			������: <font color="blue"><%=discountPro%>%</font>, ��������:<%=discountbuyRuleStr%>
			<input type="button" value="��������" onClick="CkDisPrice();" class="button">
			<!-- <input type="button" value="��������" onClick="CkOrgPrice();" class="button">-->
			&nbsp;&nbsp;
			<input type="button" value="����ǰ �߰�" onclick="addnewItem();" class="button">
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan="17" align="left">�˻���� : <b><%= FormatNumber(clsSale.FTotalCount,0) %></b>&nbsp;&nbsp;������ : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(clsSale.FTotalPage,0) %></b></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td><input type="checkbox" name="ck_all" onclick="SelectCk(this)"></td>
			<td align="center">��ǰID</td>
			<td align="center" >�̹���</td>
			<td align="center">�귣��</td>
			<td align="center">��ǰ��</td>
			<td align="center">���<br>����</td>
			<td align="center">On<br>���λ���</td>
			<td align="center">On����<br>�ǸŰ�</td>
			<td align="center">On����<br>���԰�</td>
			<td align="center">On����<br>������</td>

			<td align="center">�ؿ�<br>�ǸŰ�</td>

			<td align="center">������</td>
			<td align="center">�ؿ� ����<br>�ǸŰ�</td>
			<td align="center">�ؿ� ����<br>���԰�</td>
			<td align="center">����<br>������</td>
		</tr>
		<%
			For i = 0 To clsSale.FResultCount - 1
				mSPrice = clsSale.FItemList(i).FOrgprice - (clsSale.FItemList(i).FOrgprice*(discountPro/100))
				iSalePercent = ((clsSale.FItemList(i).FOrgprice-clsSale.FItemList(i).FDiscountPrice)/clsSale.FItemList(i).FOrgprice)*100

				if (discountbuyRule=0) then
				    mSBPrice = clsSale.FItemList(i).FOnBuycash
				elseif (discountbuyRule=1) then
				    mSBPrice = clsSale.FItemList(i).FOnBuycash ''FOrgprice*clsSale.FItemList(i).discountbuyPro/100
				else
                    mSBPrice = 0
			    end if
		%>
			<form name="frmBuyPrc_<%=clsSale.FItemList(i).FItemid%>" >
			<input type=hidden name="itemid" value="<%=clsSale.FItemList(i).FItemid%>">
			<input type=hidden name="saleprice" value="<%=mSPrice%>">
		    <input type=hidden name="salesupplyprice" value="<%=mSBPrice%>">
			<input type=hidden name="salemargin" value="<%=iSaleMargin%>">
			<input type=hidden name="orgPrice" value="<%=clsSale.FItemList(i).FOrgprice%>">
    		<input type=hidden name="orgSupplyPrice" value="<%=clsSale.FItemList(i).FOnBuycash%>">
			<input type=hidden name="saleItemStatus" value="<%=discountStatus%>">
		 <tr align="center" bgcolor=<%IF cint(discountStatus) = 8 or clsSale.FItemList(i).isSaleExpired THEN%>"#B3B3B3"<%ELSE%>"#FFFFFF"<%END IF%>>
		    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		    <td><%=clsSale.FItemList(i).FItemid%></td>
		    <td><%IF clsSale.FItemList(i).FSmallimage <> "" THEN%><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(clsSale.FItemList(i).FItemid)%>/<%=clsSale.FItemList(i).FSmallimage%>"><%END IF%></td>
		    <td><%=db2html(clsSale.FItemList(i).FMakerid)%></td>
		    <td align="left">&nbsp;<%=db2html(clsSale.FItemList(i).FItemname)%></td>
		    <td><%= fnColor(clsSale.FItemList(i).FMwdiv,"mw") %></td>
		    <td><%= clsSale.FItemList(i).getOnSaleStateStr() %></td>
		    <td>
		        <% if clsSale.FItemList(i).FOnOrgPrice>clsSale.FItemList(i).FOnSellcash then %>
		        <strike><%=formatnumber(clsSale.FItemList(i).FOnOrgPrice,0)%></strike><br>
		        <% end if %>
		        <%=formatnumber(clsSale.FItemList(i).FOnSellcash,0)%></td>
		    <td><%=formatnumber(clsSale.FItemList(i).FOnBuycash,0)%></td>
		    <td><% if clsSale.FItemList(i).FOnSellcash<>0 then %>
				<%= 100-CLNG(clsSale.FItemList(i).FOnBuycash/clsSale.FItemList(i).FOnSellcash*10000)/100 %>%
				<% end if %>
			</td>
		    <td>
		        <% if clsSale.FItemList(i).FOnOrgPrice<>clsSale.FItemList(i).FOrgprice then %>
		        <strong><%=formatnumber(clsSale.FItemList(i).FOrgprice,0)%></strong>
		        <% else %>
		        <%=formatnumber(clsSale.FItemList(i).FOrgprice,0)%>
		        <% end if %>
		    </td>
			<td id="lyrSpct<%=clsSale.FItemList(i).FItemid%>" style="<%=chkIIF(iSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=formatnumber(iSalePercent,0)%>%</td>
		<%IF cint(discountStatus) = 8 or  cint(discountStatus) = 9 THEN%>
			<td><input type="text" name="iDSPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyPrice('<%=clsSale.FItemList(i).FItemid%>')"></td>
		    <td><input type="text" name="iDBPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyPrice('<%=clsSale.FItemList(i).FItemid%>')"></td>
		    <td><input type="text" name="iDSMargin" value="0" style="text-align:right;" size="4" onkeyup="reCALbyMargin('<%=clsSale.FItemList(i).FItemid%>')">%</td>
		<%ELSE%>
		    <td><input type="text" name="iDSPrice" size="6" maxlength="9" value="<%=clsSale.FItemList(i).FDiscountPrice%>" style="text-align:right;" onkeyup="reCALbyPrice('<%=clsSale.FItemList(i).FItemid%>')"></td>
		    <td><input type="text" name="iDBPrice" size="6" maxlength="9" value="<%=clsSale.FItemList(i).FDiscountbuyMoney%>" style="text-align:right;" onkeyup="reCALbyPrice('<%=clsSale.FItemList(i).FItemid%>')"></td>
		    <td><% if clsSale.FItemList(i).FDiscountPrice<>0 then smargin= 100-CLNG(clsSale.FItemList(i).FDiscountbuyMoney/clsSale.FItemList(i).FDiscountPrice*10000)/100 	%>
				<input type="text" name="iDSMargin" value="<%=smargin%>" style=text-align:right;" size="6" onkeyup="reCALbyMargin('<%=clsSale.FItemList(i).FItemid%>')">%
		    </td>
		<%END IF%>
		</tr>
		</form>
		<% Next %>
		<tr height="20">
			<td colspan="17" align="center" bgcolor="#FFFFFF">
			<% If clsSale.HasPreScroll Then %>
				<a href="javascript:goPage('<%= clsSale.StartScrollPage-1 %>');">[pre]</a>
			<% Else %>
				[pre]
			<% End If %>
			<% For i=0 + clsSale.StartScrollPage To clsSale.FScrollCount + clsSale.StartScrollPage - 1 %>
				<% If i>clsSale.FTotalpage Then Exit For %>
				<% If CStr(page)=CStr(i) Then %>
				<font color="red">[<%= i %>]</font>
				<% Else %>
				<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
				<% End If %>
			<% Next %>
			<% If clsSale.HasNextScroll Then %>
				<a href="javascript:goPage('<%= i %>');">[next]</a>
			<% Else %>
			[next]
			<% End If %>
			</td>
		</tr>
	</td>
</tr>
</table>
<form name=frmarr method=post action="saleItemPRoc.asp">
<input type=hidden name=mode value="U">
<input type=hidden name=menupos value="<%=menupos%>">
<input type=hidden name=discountKey value="<%=discountKey%>">
<input type=hidden name=page value="<%=page%>">
<input type=hidden name=itemid value="">
<input type=hidden name=sailyn value="">
<input type=hidden name=iDSPrice value="">
<input type=hidden name=iDBPrice value="">
<input type=hidden name=saleItemStatus value="">
<input type=hidden name=saleStatus value="<%=discountStatus%>">
</form>
<form name=frmdel method=post action="saleItemPRoc.asp">
<input type=hidden name=mode value="D">
<input type=hidden name=menupos value="<%=menupos%>">
<input type=hidden name=page value="<%=page%>">
<input type=hidden name=discountKey value="<%=discountKey%>">
<input type=hidden name=itemid value="">
</form>
<%
set clsSale = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->