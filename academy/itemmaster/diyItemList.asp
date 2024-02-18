<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���̘� ��ǰ ����
' History : 2010.08.25 ������ ����
'			2010.11.09 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<%
dim itemid, itemname, makerid, sellyn, usingyn, mwdiv, limityn, vatyn, saleyn ,cdl, cdm, cds , page , i , oitem, itemid_s
dim dispCate
	itemid      = requestCheckvar(request("itemid"),200)
	itemname    = requestCheckvar(request("itemname"),64)
	makerid     = requestCheckvar(request("makerid"),32)
	sellyn      = requestCheckvar(request("sellyn"),10)
	usingyn     = requestCheckvar(request("usingyn"),10)
	mwdiv       = requestCheckvar(request("mwdiv"),10)
	limityn     = requestCheckvar(request("limityn"),10)
	vatyn       = requestCheckvar(request("vatyn"),10)
	saleyn      = requestCheckvar(request("saleyn"),10)
	cdl = requestCheckvar(request("cdl"),10)
	cdm = requestCheckvar(request("cdm"),10)
	cds = requestCheckvar(request("cds"),10)
	page = requestCheckvar(request("page"),10)
    dispCate = requestCheckvar(request("disp"),18)
	itemid_s = request("itemid_s")

	if (page="") then page=1

	if itemid_s<>"" then
	dim iA ,arrTemp,arrItemid
	itemid_s = replace(itemid_s,",",chr(10))
	itemid_s = replace(itemid_s,chr(13),"")
	arrTemp = Split(itemid_s,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then
			'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.05;������)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop
	itemid_s = left(arrItemid,len(arrItemid)-1)
	end if

set oitem = new CItem
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.FRectMakerid      = makerid
	oitem.FRectItemid       = itemid_s
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectVatYn        = vatyn
	oitem.FRectsaleyn       = saleyn
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectDispCate     = dispCate
	oitem.GetItemList()
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

// ============================================================================
// �ɼǼ��� -��ü
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/academy/comm/pop_diyitemoptionedit.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//�Ǹż���
function PopItemSellEdit(iitemid){
	var popwin = window.open('/academy/comm/pop_diy_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

// ============================================================================
// �̹�������
function editItemImage(itemid, makerid) {
	var param = "itemid=" + itemid;
	popwin = window.open('/academy/comm/pop_diy_itemImage.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ��ǰ���� �̹��� ���/����
function popItemContImage(itemid)
{
	var popwin = window.open("/academy/itemmaster/pop_diyItem_imgcontents_write.asp?mode=edit&itemid=" + itemid + "&menupos=423","popitemContImage","width=600 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// �⺻���� ����
function editItemBasicInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>&fingerson=on";
	popwin = window.open('pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// �ǸŰ� �� ���ް� ����
function editItemPriceInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('pop_ItemPriceInfo.asp?' + param ,'editItemPrice','width=780,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//�÷��� ���̻�ǰ
function pop_plusdiyitem(itemid){	
	var pop_plusdiyitem = window.open('/academy/itemmaster/PlusDIYItem/PlusDIYItem_list.asp?itemid='+itemid,'pop_plusdiyitem','width=1024,height=768,scrollbars=yes,resizable=yes')
	pop_plusdiyitem.focus();
}

function Check_All()
{
	var chk = document.frm.itemid; 
	var cnt = 0;
	var ischecked = ""
	if(document.getElementById("chkall").checked){
		ischecked = "checked"
	}else{
		ischecked = ""
	}
	if(cnt == 0 && chk.length != 0){
		for(i = 0; i < chk.length; i++){ chk.item(i).checked = ischecked; }
		cnt++;
	}
}

function fnSoldOutItems(){
	var i = "";
	$("input:checkbox[name='itemid']").each(
		function(){
			if (this.checked)
			{
				i = i + this.value + ",";
			}
		}
	)
	
	if(i == ""){
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}else{
		if(confirm("�����Ͻ� ��ǰ���� ǰ�� ó�� �Ͻðڽ��ϱ�?") == true) {
			$('input[name="allitemid"]').val(i);
			$('input[name="action"]').val('soldout');
			frmallitem.submit();
		}else{
			return;
		}
	}
}
function fnNotSale(){
	var i = "";
	$("input:checkbox[name='itemid']").each(
		function(){
			if (this.checked)
			{
				i = i + this.value + ",";
			}
		}
	)
	
	if(i == ""){
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}else{
		if(confirm("�����Ͻ� ��ǰ���� �Ǹ� �� �� ó�� �Ͻðڽ��ϱ�?") == true) {
			$('input[name="allitemid"]').val(i);
			$('input[name="action"]').val('notsale');
			frmallitem.submit();
		}else{
			return;
		}
	}
}

function fnSellYNIsusingEditEnd(){
	document.frm.itemid_s.value="";
	document.frm.submit();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣�� :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;&nbsp;
		<!-- #include virtual="/academy/comm/CategorySelectBox.asp"-->
		&nbsp;&nbsp;
		����ī�װ� : 
		<script type="text/javascript">
		$(function(){
			chgDispCate('<%=dispCate%>');
		});
		
		function chgDispCate(dc) {
			$.ajax({
				url: "/academy/comm/dispCateSelectBox_response.asp?disp="+dc,
				cache: false,
				async: false,
				success: function(message) {
		       		// ���� �ֱ� 
		       		$("#lyrDispCtBox").empty().html(message);
		       		$("#oDispCate").val(dc);
				}
			});
		}
		</script>
		<span id="lyrDispCtBox"></span>
		<input type="hidden" name="disp" id="oDispCate" value="<%=dispCate%>">
		
	</td>
	
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		��ǰ�ڵ� :<textarea rows="3" cols="10" name="itemid_s" id="itemid_s"><%=replace(itemid_s,",",chr(10))%></textarea>
		&nbsp;
		��ǰ�� :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
     	&nbsp;
     	���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
     	&nbsp;     	
     	����:<% drawSelectBoxLimitYN "limityn", limityn %>
     	&nbsp;
     	�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
     	&nbsp;
     	����: <% drawSelectBoxVatYN "vatyn", vatyn %>
     	&nbsp;
     	���� <% drawSelectBoxsailyn "saleyn", saleyn %>
	</td>
</tr>
</table>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" value="���û�ǰ �Ͻ�ǰ��ó��" onClick="fnSoldOutItems()">&nbsp;<input type="button" value="���û�ǰ �Ǹž���ó��" onClick="fnNotSale()">
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- �׼� �� -->

<br>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oitem.FTotalCount%></b>
		&nbsp;
		������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkall" id="chkall" value="" onClick="Check_All()"></td>
	<td>No.</td>
	<td>�̹���</td>
	<td>�귣��ID</td>
	<td>��ǰ��</td>
	<td>�ǸŰ�</td>
	<td>���԰�</td>
	<td>����</td>
	<td>���<br>����</td>
	<td>�Ǹſ���<br>��뿩��</td>	
	<td>����<br>����</td>
	<td>����<br>�鼼</td>
	<td>����<br>�̹���</td>
	<td>�߰�����</td>
	<td>�����</td>
</tr>
<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="itemid" value="<%=oitem.FItemList(i).Fitemid%>"></td>
	<td align="center">
		<a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">				
		<%= oitem.FItemList(i).Fitemid %></a>
		</td>
	<td align="center"><a href="javascript:editItemImage('<%= oitem.FItemList(i).FItemId %>','<%= oitem.FItemList(i).Fmakerid %>')" title="�̹��� ����"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="60" border="0"></a></td>
	<td align="left"><a href="javascript:PopBrandInfoEdit('<%= oitem.FItemList(i).Fmakerid %>')" title="�귣�� ���� ����"><%= oitem.FItemList(i).Fmakerid %></a></td>
	<td align="left">
		<a href="javascript:editItemBasicInfo('<% =oitem.FItemList(i).Fitemid %>')" title="��ǰ �⺻���� ����"><% =oitem.FItemList(i).Fitemname %></a>
	</td>
	<td align="right">
	<%
		Response.Write "<a href=""javascript:editItemPriceInfo('" & oitem.FItemList(i).Fitemid & "')"" title='�ǸŰ� �� ���ް� ����'>" & FormatNumber(oitem.FItemList(i).Forgprice,0) & "</a>"
		'���ΰ�
		if oitem.FItemList(i).Fsaleyn="Y" then
			Response.Write "<br><font color=#F08050>("&CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) & "%��)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
		end if
		'������
		if oitem.FItemList(i).FitemCouponYn="Y" then
			Select Case oitem.FItemList(i).FitemCouponType
				Case "1"
					Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
				Case "2"
					Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
			end Select
		end if
	%>
	</td>
	<td align="right">
	<%
		Response.Write FormatNumber(oitem.FItemList(i).Forgsuplycash,0)
		'���ΰ�
		if oitem.FItemList(i).Fsaleyn="Y" then
			Response.Write "<br><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font>"
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
		if oitem.FItemList(i).Fsaleyn="Y" then
			Response.Write "<br><font color=#F08050>" & fnPercent(oitem.FItemList(i).Fsailsuplycash,oitem.FItemList(i).Fsailprice,1) & "</font>"
		end if
		'������
		if oitem.FItemList(i).FitemCouponYn="Y" then
			Select Case oitem.FItemList(i).FitemCouponType
				Case "1"
					if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
						Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),1) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),1) & "</font>"
					end if
				Case "2"
					if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
						Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,1) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,1) & "</font>"
					end if
			end Select
		end if
	%>
	</td>
	<td align="center">
		<a href="javascript:PopItemSellEdit('<%= oitem.FItemList(i).FItemId %>')" title="�Ǹ�����/�ɼ� ����"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></a>
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
	<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %><br><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>	
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fvatyn,"tx") %></td>
	<td align="center">
    	<% if Not (oitem.FItemList(i).FinfoimageExists) then %>
    	<a href="javascript:popItemContImage('<%= oitem.FItemList(i).FItemId %>')" title="��ǰ���� �̹��� ���"><font color="#F08050">N [���]</font></a>
    	<% else %>
    	<a href="javascript:popItemContImage('<%= oitem.FItemList(i).FItemId %>')" title="��ǰ���� �̹��� ����"><font color="#5080F0">Y [����]</font></a>
    	<% end if %>
    </td>
    <td align="center">
    	<% if oitem.FItemList(i).Fitemdiv <> "20" and oitem.FItemList(i).fPlusdiyItemregCount = "0" then %>
	    	<input type="button" onclick="pop_plusdiyitem(<%=oitem.FItemList(i).fitemid%>)" value="�÷�����ǰ[<%=oitem.FItemList(i).fPlusdiyItemCount%>]" class="button">
    	<% end if %>
    	<% if oitem.FItemList(i).Fitemdiv = "20" then %>
	    	��ǰ���� : �߰������ǰ
    	<% end if %>
		    	
    	<% if oitem.FItemList(i).fPlusdiyItemregCount > 0 then %>
	    	<br>�÷����߰����� : Y
    	<% end if %>	    
    </td>
	<td align="center"><%= FormatDate(oitem.FItemList(i).Fregdate,"0000.00.00") %></td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
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
</form>
</table>
<% end if %>

<%
	set oitem = nothing
%>
<form name="frmallitem" method="post" action="dodiyItemsoldoutnosale.asp" target="cateitemproc">
<input type="hidden" name="action" value="">
<input type="hidden" name="allitemid" value="">
</form>
<iframe src="" id="cateitemproc" name="cateitemproc" width="0" height="0" frameborder="0"></iframe>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->