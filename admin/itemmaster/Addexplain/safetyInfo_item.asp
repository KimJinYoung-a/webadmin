<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ ������������ ���� - ��ǰ���
' Hieditor : 2015.05.28 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%

dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, overSeaYn, itemdiv
dim cdl, cdm, cds, showminusmagin, marginup, margindown
dim page
dim infodivYn, saftyYn, saftyInfoYn

itemid      = requestCheckvar(request("itemid"),255)
itemname    = request("itemname")
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

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)

showminusmagin = request("showminusmagin")
marginup = request("marginup")
margindown = request("margindown")
infodivYn  = requestCheckvar(request("infodivYn"),10)
saftyYn  = requestCheckvar(request("saftyYn"),1)
saftyInfoYn  = requestCheckvar(request("saftyInfoYn"),1)
If saftyInfoYn="" Then saftyInfoYn = "N"
''If infodivYn = "K" Then sellyn = "Y"

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

	arrTemp = Split(itemid,",")

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

oitem.FPageSize         = 30
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
oitem.FRectItemDiv      = itemdiv

oitem.FRectMinusMigin = showminusmagin
oitem.FRectMarginUP = marginup
oitem.FRectMarginDown = margindown
oitem.FRectInfodivYn    = infodivYn
oitem.FRectsaftyYn		= saftyYn
oitem.FRectsaftyInfoYn  = saftyInfoYn
oitem.FRectShowInfoDiv  = "on"
oitem.FRectSortDiv="best"               ''����Ʈ��.

oitem.getSafetyInfoItemList

dim i
%>
<script>
function NextPage(ipage){
	document.frm.page.value= ipage;
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
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

// ============================================================================
// �̹�������
function editItemImage(itemid, makerid) {
	var param = "itemid=" + itemid;
	
	//if(makerid =="ithinkso"){
		//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}else{
		popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
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
	popwin = window.open('/admin/itemmaster/pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=1100,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// �ǸŰ� �� ���ް� ����
function editItemPriceInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('/admin/itemmaster/pop_ItemPriceInfo.asp?' + param ,'editItemPrice','width=750,height=600,scrollbars=yes,resizable=yes');
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
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			<br>
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
			<br>
		</td>
		
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	&nbsp;
	     	���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;     	
	     	����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	     	&nbsp;
	     	����:<% drawSelectBoxLimitYN "limityn", limityn %>
	     	&nbsp;
	     	�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
	     	����: <% drawSelectBoxVatYN "vatyn", vatyn %>
	     	&nbsp;
	     	���� <% drawSelectBoxSailYN "sailyn", sailyn %>
	     	
	     	&nbsp;
	     	�ؿܹ�� <% drawSelectBoxIsOverSeaYN "overSeaYn", overSeaYn %>
            &nbsp;
	     	��ǰ���� <% drawSelectBoxItemDiv "itemdiv", itemdiv %>
	     	&nbsp;
	     	ǰ�������Է¿���
			<select class="select" name="infodivYn">
			<option value="">��ü</option>
			<option value="N" <%= CHKIIF(infodivYn="N","selected","") %> >�Է�����</option>
			<option value="Y" <%= CHKIIF(infodivYn="Y","selected","") %> >�Է¿Ϸ�</option>
			<option value="K" <%= CHKIIF(infodivYn="K","selected","") %> >�׸񴩶�</option>
			</select>
			&nbsp;
			<font color="red">����������󿩺�</font>
			<select class="select" name="saftyYn">
			<option value="">��ü</option>
			<option value="N" <%= CHKIIF(saftyYn="N","selected","") %> >���ƴ�</option>
			<option value="Y" <%= CHKIIF(saftyYn="Y","selected","") %> >�������</option>
			</select>
			&nbsp;
			<font color="red">KC��ũ�Է¿���</font>
			<select class="select" name="saftyInfoYn">
			<option value="A" <%= CHKIIF(saftyInfoYn="A","selected","") %> >��ü</option>
			<option value="N" <%= CHKIIF(saftyInfoYn="N","selected","") %> >�Է�����</option>
			<option value="Y" <%= CHKIIF(saftyInfoYn="Y","selected","") %> >�Է¿Ϸ�</option>
			</select>
    </tr>
    </form>
</table>

<p>
<% If cdl = "110" and cdm = "010" and cds = "968" Then %>
<input type="button" value="����� ���ø��ڵ� ���" class="button" onClick="window.open('pop_photobook.asp','popPhotobook','width=600,height=650,scrollbars=yes');"><p>
<% End If %>

<!-- ����Ʈ ���� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><%= oitem.FTotalCount%></b>
			&nbsp;
			������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">No.</td>
		<td width=50> �̹���</td>
		<td width="100">�귣��ID</td>
		<td> ��ǰ��</td>
		<td width="60">�ǸŰ�</td>
		<td width="30">�Ǹ�<br>����</td>
		<td width="30">���<br>����</td>
		<td width="60" bgcolor="#FFF0F0">��������<br>����</td>
		<td width="130" bgcolor="#FFF0F0">��������<br>����</td>
		<td width="170" bgcolor="#FFF0F0">��������<br>��ȣ</td>
		<td width="150">ǰ��</td>
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
		<td align="left"><a href="javascript:PopBrandInfoEdit('<%= oitem.FItemList(i).Fmakerid %>')" title="�귣�� ���� ����"><%= oitem.FItemList(i).Fmakerid %></a></td>
		<td align="left">
			<a href="javascript:editItemBasicInfo('<% =oitem.FItemList(i).Fitemid %>')" title="��ǰ �⺻���� ����"><% =oitem.FItemList(i).Fitemname %></a>
			<% if oitem.FItemList(i).FitemDiv="08" then %>
            <a href="javascript:editTicketItemInfo('<% =oitem.FItemList(i).Fitemid %>')" title="Ticket ���� ����"><font color="#F89020">[Ticket]</font></a>	
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
			end if
		%>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>

		<td align="center"><%= fnColor(oitem.FItemList(i).FsafetyYn,"yn") %></td>
	    <td align="center"><%= getSaftyDivName(oitem.FItemList(i).FsafetyYn,oitem.FItemList(i).FsafetyDiv) %></td>
	    <td align="center"><%= oitem.FItemList(i).FsafetyNum %></td>
	    <td align="center"><%= getAddExpInfoDivName(oitem.FItemList(i).FinfoDiv) %></td>
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

<%
SET oitem = Nothing
%>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->