<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��ǰ ����
' History : 2010.09.30 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/lib/util/base64.asp"-->
<!-- #include virtual="/academy/lib/classes/diyshopitem/itemcouponcls.asp" -->
<%
dim itemcouponidx ,sRectItemidArr ,page, makerid,sailyn, invalidmargin ,oitemcouponmaster, ocouponitemlist ,i
	itemcouponidx   = RequestCheckvar(request("itemcouponidx"),10)
	makerid         = RequestCheckvar(request("makerid"),32)
	page            = RequestCheckvar(request("page"),10)
	sailyn          = RequestCheckvar(request("sailyn"),1)
	invalidmargin   = RequestCheckvar(request("invalidmargin"),1)
	sRectItemidArr  = Trim(request("sRectItemidArr"))
  	if sRectItemidArr <> "" then
		if checkNotValidHTML(sRectItemidArr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if Right(sRectItemidArr,1)="," then sRectItemidArr=Left(sRectItemidArr,Len(sRectItemidArr)-1)
	
	if itemcouponidx="" then itemcouponidx=0
	if page="" then page=1

set oitemcouponmaster = new CItemCouponMaster
	oitemcouponmaster.FRectItemCouponIdx = itemcouponidx
	oitemcouponmaster.GetOneItemCouponMaster()

set ocouponitemlist = new CItemCouponMaster
	ocouponitemlist.FPageSize=50
	ocouponitemlist.FCurrPage=page
	ocouponitemlist.FRectItemCouponIdx = itemcouponidx
	ocouponitemlist.FRectMakerid = makerid
	ocouponitemlist.FRectSailYn = sailyn
	ocouponitemlist.FRectInvalidMargin = invalidmargin
	ocouponitemlist.FRectsRectItemidArr = sRectItemidArr
	ocouponitemlist.GetItemCouponItemList()
%>

<script language='javascript'>

function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function AddIttems(){
	frmbuf.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function EditArr(){
	var upfrm = document.frmbuf;
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

	upfrm.itemidarr.value = "";
	upfrm.couponbuypricearr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsDigit(frm.couponbuyprice.value)){
					alert('���԰��� ���ڸ� �����մϴ�.');
					frm.couponbuyprice.focus();
					return;
				}

				upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + ",";
				upfrm.couponbuypricearr.value = upfrm.couponbuypricearr.value + frm.couponbuyprice.value + ",";

			}
		}
	}

	if (confirm('���� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
		frmbuf.mode.value="modicouponitemarr"
		frmbuf.submit();
	}
}

function DelArr(){
	var upfrm = document.frmbuf;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.itemidarr.value = "";
	upfrm.couponbuypricearr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + ",";
			}
		}
	}

	if (confirm('���� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
		upfrm.mode.value="delcouponitemarr"
		frmbuf.submit();
	}
}

// Old
function AddNewCouponItem(targetcomp){
	var popwin;
	popwin = window.open("/admin/pop/viewitemlist3.asp?dispyn=Y&sellyn=Y&sailyn=N&target=" + targetcomp, "AddNewCouponItem", "width=800,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// ����ǰ �߰� �˾�
function addnewItem(couponCD,evtCd){
		var popwin;
		if ( evtCd > 0 ){
			popwin = window.open("/academy/event/common/pop_eventitem_addinfo.asp?defaultmargin=<%=oitemcouponmaster.FOneItem.FDefaultMargin%>&acURL=/academy/shopmaster/itemcoupon_Process.asp?itemcouponidx=" + couponCD, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}else{
			popwin = window.open("/academy/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&sailyn=N&defaultmargin=<%=oitemcouponmaster.FOneItem.FDefaultMargin%>&acURL=/academy/shopmaster/itemcoupon_Process.asp?itemcouponidx=" + couponCD, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}
		popwin.focus();
}

// Ŭ������� ����
function fnCBCopy(iid) {
	var doc = "<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=" + iid + "&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>";
	clipboardData.setData("Text", doc);
	alert('�����Ͻ� ��ǰ�� ��ũ�� ����Ǿ����ϴ�. ����Ͻ� ���� Ctrl+V �Ͻø�˴ϴ�.');
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#DDDDFF">
	<td width="100">������</td>
	<td bgcolor="#FFFFFF"><%= oitemcouponmaster.FOneItem.Fitemcouponname %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >������</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.GetDiscountStr %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >����Ⱓ</td>
	<td bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.Fitemcouponstartdate %> ~ <%= oitemcouponmaster.FOneItem.Fitemcouponexpiredate %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >��������</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.GetMargintypeName %> <% if oitemcouponmaster.FOneItem.FDefaultMargin<>0 then %>- (<%= oitemcouponmaster.FOneItem.FDefaultMargin %>%) <% End IF %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >�߱� ����</td>
	<td bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.GetOpenStateName %>
	</td>
</tr>
</table>

<!---- �˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
    	<input type="hidden" name="itemcouponidx" value="<%= itemcouponidx %>" >
    	�귣�� : <% drawSelectBoxLecturer "makerid",makerid %>
    	<input type="checkbox" name="sailyn" value="Y" <% if sailyn="Y" then response.write "checked" %> >�������� ��ǰ �˻�
        &nbsp;<input type="checkbox" name="invalidmargin" value="Y" <% if invalidmargin="Y" then response.write "checked" %> >��������(or ������) ��ǰ �˻�
        <br>
        ��ǰ�ڵ�:<input type="text" name="sRectItemidArr" value="<%= sRectItemidArr %>" size="50" maxlength="50">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!---- /�˻� ---->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<span>* <font color="red">��������� ���԰� 0</font>�� ���� ������ ���԰��� �����˴ϴ�. (���԰� ������ ���°��� 0���� �����Ұ�!)</span><br>
			<input type="button" class="button" value="���û�ǰ����" onclick="EditArr();">
			<input type="button" class="button" value="���û�ǰ����" onclick="DelArr();">				
		</td>			
		<td align="right">
			<input type="button" class="button" value="�űԵ��" onclick="addnewItem('<%= itemcouponidx %>');">
		</td>				
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ocouponitemlist.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ocouponitemlist.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= ocouponitemlist.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="AnSelectAllFrame(this.checked)"></td>
	<td>�̹���</td>
	<td>�귣��</td>
	<td>��ǰ��ȣ</td>
	<td >��ǰ��</td>
	<td>���� �ǸŰ�</td>
	<td>���� ���԰�</td>
	<td>����<br>����</td>
	<td>���� ����</td>
	<td>���������<br>�ǸŰ�</td>
	<td>���������<br>���԰�</td>
	<td>���������<br>����</td>
</tr>
<% for i=0 to ocouponitemlist.FResultCount - 1 %>
<form name="frmBuyPrc_<%= ocouponitemlist.FitemList(i).FItemID %>" method="post" onSubmit="return false;" action="do_itemcoupon.asp">
<input type="hidden" name="itemid" value="<%= ocouponitemlist.FitemList(i).FItemID %>">
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td ><img src="<%= ocouponitemlist.FitemList(i).FSmallimage %>"width="50"></td>
	<td><%= ocouponitemlist.FitemList(i).FMakerid %></td>
	<td align="center"><%= ocouponitemlist.FitemList(i).FItemID %>
    	<% if oitemcouponmaster.FOneItem.FcouponGubun="T" then %>
    	<input type="button" class="button" value="URL����" onClick="fnCBCopy('<%=ocouponitemlist.FitemList(i).FItemID%>')">
    	<% end if %>
	</td>
	<td ><%= ocouponitemlist.FitemList(i).FItemName %></td>
	<td align="right"><%= FormatNumber(ocouponitemlist.FitemList(i).FSellcash,0) %></td>
	<td align="right"><%= FormatNumber(ocouponitemlist.FitemList(i).FBuycash,0) %></td>
	<td align="center"><font color="<%= ocouponitemlist.FitemList(i).GetMwDivColor %>"><%= ocouponitemlist.FitemList(i).GetMwDivName %></font>
	<td align="center"><%= ocouponitemlist.FitemList(i).GetCurrentMargin %>%</td>
	<td align="right"><%= FormatNumber(ocouponitemlist.FitemList(i).GetCouponSellcash,0) %>
	<% if ocouponitemlist.FitemList(i).Fitemcoupontype="3" then %>
	<br><font color="red">(������)</font>
	<% end if %>
	</td>
	<td align="right"><input type="text" name="couponbuyprice" value="<%= ocouponitemlist.FitemList(i).Fcouponbuyprice %>" size="7" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyDown="CheckThis(this.form);"></td>
	<td align="center"><font color="<%= ocouponitemlist.FitemList(i).GetCouponMarginColor %>"><%= ocouponitemlist.FitemList(i).GetCouponMargin %></font>%
	<% if ocouponitemlist.FitemList(i).Fitemcoupontype="3" then %>
	    <br><font color="red"><%= ocouponitemlist.FitemList(i).GetFreeBeasongCouponMargin %></font>%
	<% end if %>
	</td>
</tr>
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if ocouponitemlist.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ocouponitemlist.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ocouponitemlist.StarScrollPage to ocouponitemlist.FScrollCount + ocouponitemlist.StarScrollPage - 1 %>
			<% if i>ocouponitemlist.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ocouponitemlist.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<form name="frmbuf" method="post" action="/academy/shopmaster/itemcoupon_process.asp">
	<input type="hidden" name="mode" value="addcouponitemarr">
	<input type="hidden" name="itemcouponidx" value="<%= itemcouponidx %>">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="couponbuypricearr" value="">
	<input type="hidden" name="makerid" value="<%= makerid %>">
	<input type="hidden" name="sailyn" value="<%= sailyn %>">
	<input type="hidden" name="defaultmargin" value="">
</form>

<%
	set ocouponitemlist = Nothing
	set oitemcouponmaster = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
