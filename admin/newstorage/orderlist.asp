<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [LOG]��������>>�ֹ�������
' History : �̻� ����
'			2018.08.07 ������ �˻��� ��� �۾��ð� ǥ��
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim page, shopid,yyyy1,yyyy2,mm1,mm2,dd1,dd2,fromDate,toDate, vPurchaseType, designer, statecd, baljucode, blinkcode, i
dim barcode, searchfield, searchtext, minusjumun
dim rstate
dim itemgubun, itemid, itemoption
	barcode   = request("barcode")
	designer = request("designer")
	statecd  = request("statecd")
	page = request("page")
	if page="" then page=1
	shopid = request("shopid")
	baljucode = request("baljucode")
	blinkcode = request("blinkcode")
	yyyy1 = request("yyyy1")
	yyyy2 = request("yyyy2")
	mm1	  = request("mm1")
	mm2	  = request("mm2")
	dd1	  = request("dd1")
	dd2	  = request("dd2")
	searchfield	  = requestCheckvar(request("searchfield"),16)
	searchtext	  = requestCheckvar(request("searchtext"),32)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	rstate= requestCheckvar(request("rstate"),32)
	minusjumun	  = requestCheckvar(request("minusjumun"),2)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now())-1)
if (dd1="") then dd1 = 1'Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

itemid = ""
barcode = Replace(barcode, "-", "")
if (Len(barcode) = 12) then
    itemgubun   = Mid(getNumeric(barcode), 1, 2)
    itemid      = Mid(getNumeric(barcode), 3, 6)
    itemoption  = Mid(barcode, 9, 4)
elseif (Len(barcode) = 14) then
    itemgubun   = Mid(getNumeric(barcode), 1, 2)
    itemid      = Mid(getNumeric(barcode), 3, 8)
    itemoption  = Mid(barcode, 11, 4)
elseif (Len(barcode)>6) then
    '''���ڵ��ΰ�� �˻��� ��ǰ�ڵ� ������.
    call fnGetItemCodeByPublicBarcode(barcode, itemgubun, itemid, itemoption)
end if

if trim(barcode)<>"" then
    if itemid = "" and IsNumeric(barcode) then
        itemid = barcode
    elseif Not IsNumeric(itemid) then
        response.write "<script>alert('�߸��� ���ڵ��Դϴ�.[" & barcode & "]');</script>"
        itemgubun = ""
        itemid = ""
        itemoption = ""
    end if
end if

'barcode = itemgubun & Format00(8,itemid) & itemoption

baljucode = Trim(baljucode)
blinkcode = Trim(blinkcode)

dim osheet
set osheet = new COrderSheet
	osheet.FCurrPage = page
	osheet.Fpagesize=20
	osheet.FRectBaljuid = shopid
	osheet.FRectStatecd = statecd
	osheet.FRecttargetid = designer
	osheet.FRectDivCodeArr = "('301','302')"
	osheet.FRectStartDate = fromDate
	osheet.FRectEndDate = toDate
	osheet.FRectBrandPurchaseType = vPurchaseType
	osheet.FRectBaljuCode = baljucode
	osheet.FRectBLinkCode = blinkcode
	osheet.FRectSearchField = searchfield
	osheet.FRectSearchText = searchtext
	osheet.FRectitemgubun = itemgubun
	osheet.FRectitemid = itemid
	osheet.FRectitemoption = itemoption
	osheet.FRectReportState = rstate
	osheet.FRectMinusOnly = minusjumun
	osheet.GetOrderSheetList
%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script src="/js/jquery.placeholder.min.js"></script>
<script type='text/javascript'>

function PopIpgoSheet(v,itype){
	var popwin;
	popwin = window.open('/admin/newstorage/popjumunsheet.asp?idx=' + v + '&itype=' + itype,'popjumunsheet','width=760,height=600,scrollbars=yes,resizabled=yes');
	popwin.focus();
}

function PopIpgoSheetforeign(v,itype){
	var popwin;
	popwin = window.open('/admin/newstorage/popjumunsheet_foreign.asp?idx=' + v + '&itype=' + itype,'popjumunsheet','width=760,height=600,scrollbars=yes,resizabled=yes');
	popwin.focus();
}

function ExcelSheet(v,itype){
	window.open('/admin/newstorage/popjumunsheet.asp?idx=' + v + '&itype=' + itype + '&xl=on');
}

function ExcelSheetforeign(v,itype){
	window.open('/admin/newstorage/popjumunsheet_foreign.asp?idx=' + v + '&itype=' + itype + '&xl=on');
}

function MakeOrder(){
	location.href="orderinput.asp";
}

function PopUpcheBrandInfoEdit(v){
	var popwin = window.open("/admin/member/popupchebrandinfo.asp?designer=" + v,"PopUpcheBrandInfoEdit","width=640,height=580,scrollbars=yes,resizabled=yes");
    popwin.focus();
}

function sendSMSEmail(idesigner,iidx){
	var popwin = window.open("/admin/lib/popupchejumunsms.asp?designer=" + idesigner + "&idx=" + iidx,"popupchejumunsms","width=600 height=500,scrollbars=yes,resizabled=yes");
	popwin.focus();
}

function IpgoProc(iidx){
	var popwin = window.open("popipgoproc.asp?idx=" + iidx ,"popipgoproc","width=800,height=550,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function SubmitFrm() {
	ClearPlaceHolder();
	document.frm.submit();
}

function NextPage(page){
	ClearPlaceHolder();
    document.frm.page.value = page;
    document.frm.submit();
}

function PopOpenorder(idx,loginsite, cunit, tpl) {
	var popwin;

	popwin = window.open('/admin/newstorage/ordersheet.asp?idx=' + idx+'&ls='+ loginsite+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenorder','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopOpenorderUTF8(idx,loginsite, cunit, tpl) {
	var popwin;

	popwin = window.open('/admin/newstorage/ordersheet_UTF8.asp?idx=' + idx+'&ls='+ loginsite+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenorderUTF8','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//���ڰ��� ǰ�Ǽ� ���뺸��
function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}

function popOpenPPMaster(idx) {
	var popwin;

	popwin = window.open('/admin/newstorage/PurchasedProductModify.asp?menupos=9175&idx=' + idx ,'popOpenPPMaster','width=1400,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ClearPlaceHolder() {
	var frm = document.frm;
	frm.baljucode.value = $('#baljucode').val();
	frm.blinkcode.value = $('#blinkcode').val();
}

$( document ).ready(function() {
    $('textarea').placeholder();
});

</script>

<style>
textarea:-webkit-input-placeholder {color:#acacac;}
textarea:-moz-placeholder {color:#acacac;}
textarea:-ms-input-placeholder {color:#acacac;}
.placeholder { color: #acacac; }
</style>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
    	* �귣�� : <% drawSelectBoxDesignerwithName "designer", designer %>
		&nbsp;
    	* �ֹ��ڵ� :
		<textarea class="textarea" id="baljucode" name="baljucode" cols="12" rows="1" placeholder="�ִ�50��"><%= baljucode %></textarea>
		&nbsp;
    	* �԰��ڵ� :
		<textarea class="textarea" id="blinkcode" name="blinkcode" cols="12" rows="1" placeholder="�ִ�50��"><%= blinkcode %></textarea>
		&nbsp;
		* �ֹ��� : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitFrm();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �ֹ����� :
		<select class="select" name="statecd" >
			<option value="">��ü
			<option value="0" <% if statecd="0" then response.write "selected" %> >�ֹ�����
			<option value="1" <% if statecd="1" then response.write "selected" %> >�ֹ�Ȯ��
			<option value="2" <% if statecd="2" then response.write "selected" %> >�Աݴ��
			<option value="5" <% if statecd="5" then response.write "selected" %> >����غ�
			<option value="7" <% if statecd="7" then response.write "selected" %> >���Ϸ�
			<option value="8" <% if statecd="8" then response.write "selected" %> >��ǰ�Ϸ�
			<option value="9" <% if statecd="9" then response.write "selected" %> >�԰�Ϸ�
			<option value="preorder" <% if statecd="preorder" then response.write "selected" %> >�԰�����(���ֹ�)
     	</select>
     	&nbsp;
		* ǰ�ǻ��� :
		<select class="select" name="rstate" >
			<option value="">��ü</option>
			<option value="0" <% if rstate="0" then response.write "selected" %> >ǰ���ۼ���</option>
			<option value="1" <% if rstate="1" then response.write "selected" %> >ǰ�������� </option>
			<option value="5" <% if rstate="5" then response.write "selected" %> >ǰ�ǹݷ� </option>
			<option value="7" <% if rstate="7" then response.write "selected" %> >ǰ�ǿϷ�</option>
		</select>
     	&nbsp;
		* �������� :
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
		&nbsp;
		* ��ǰ�ڵ�,�����ڵ�,������ڵ� :
		<input type="text" name="barcode" value="<%= barcode %>" size="20" maxlength="20" class="text" >
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �˻����� :
		<select class="select" name="searchfield">
			<option value="" >��ü</option>
			<option value="p1.username" <% if (searchfield = "p1.username") then %>selected<% end if %> >��ǰ��</option>
			<option value="m.finishname" <% if (searchfield = "m.finishname") then %>selected<% end if %> >�԰�ó����</option>
			<option value="socname" <% if (searchfield = "socname") then %>selected<% end if %> >��ü��</option>
			<option value="socno" <% if (searchfield = "socno") then %>selected<% end if %> >����ڹ�ȣ</option>
		</select>
		<input type="text" class="text" name="searchtext" value="<%= searchtext %>">
		&nbsp;
		<input type="checkbox" name="minusjumun" <% if minusjumun="on" then  response.write "checked" %>>���̳ʽ��ֹ���
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="�ֹ����ۼ�" onclick="MakeOrder();">
	</td>
	<td align="right">
	<% if osheet.FResultCount > 0 and statecd="9" then %>
			�۾��ð� ��� : <%= osheet.FAverageWorkSecond %>
	<% end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		�˻���� : <b><%= osheet.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= osheet.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=60>�ֹ��ڵ�</td>
	<td>����ó</td>
	<td width=60>�ֹ���</td>
	<td width=70>����</td>
	<td width=70>�ֹ�����</td>
	<td width="60">����IDX</td>
    <td width="60">ǰ�ǻ���</td>
	<td width=80>�ֹ���/<br>�԰��û��</td>
	<td width=100>���ֹ���<br>Ȯ����(�Һ��ڰ�)</td>
	<td width=90>���ֹ���<br>Ȯ����(���԰�)</td>
	<td width=70>������/<br>�߼���</td>
	<td width=100>�ù��<br>�����ȣ</td>
	<td width=50>��ǰ��</td>
	<td width="90">�۾��ð�<br />(��ǰ)</td>
	<td width=50>�԰�<br>ó����</td>
	<td width=50>������</td>
	<td width=50>����<br>������</td>
</tr>
<%
dim totaljumunsuply, totalfixsuply, totaljumunsellcash, totaltotalsellcash
if osheet.FResultCount >0 then
%>
	<% for i=0 to osheet.FResultcount-1 %>
	<%
	totaljumunsellcash = totaljumunsellcash + osheet.FItemList(i).Fjumunsellcash
	totaltotalsellcash = totaltotalsellcash + osheet.FItemList(i).Ftotalsellcash

	if osheet.FItemList(i).Ftargetid="10x10" then
		totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunsuplycash
		totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalsuplycash
	else
		totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunbuycash
		totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalbuycash
	end if

	if IsNull(osheet.FItemList(i).Fsongjangno) then
		osheet.FItemList(i).Fsongjangno = ""
	end if

	%>
	<tr bgcolor="#FFFFFF">
		<td rowspan=2 align=center><a href="jumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>&opage=<%= page %>&menupos=<%= menupos %>&odesigner=<%= designer %>&ostatecd=<%= statecd %>"><%= osheet.FItemList(i).Fbaljucode %></a></td>
		<td rowspan=2 align=center><b><a href="javascript:PopUpcheBrandInfoEdit('<%= osheet.FItemList(i).Ftargetid %>');"><%= osheet.FItemList(i).Ftargetid %></a></b><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
		<td rowspan=2 align=center><%= osheet.FItemList(i).Fregname %></td>
		<td rowspan=2 align=center><%= osheet.FItemList(i).GetDivCodeName %></td>
		<td rowspan=2 align=center><font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font></td>
        <td rowspan=2 align=center>
            <% if (osheet.FItemList(i).FppMasterIdx <> "" and not(isnull(osheet.FItemList(i).FppMasterIdx))) then %>
				<a href="#" onclick="popOpenPPMaster(<%= osheet.FItemList(i).FppMasterIdx %>); return false;"><%= osheet.FItemList(i).FppMasterIdx %></a>
            <% end if %>
        </td>
		<td rowspan=2 align=center>
			<%if osheet.FItemList(i).Freportidx <> "" and not isNUll( osheet.FItemList(i).Freportidx ) then%>
				<a href="javascript:jsViewEapp('<%=osheet.FItemList(i).Freportidx%>','<%= osheet.FItemList(i).Freportstate %>');">
				<%if osheet.FItemList(i).Freportstate = "7" or   osheet.FItemList(i).Freportstate ="8" or   osheet.FItemList(i).Freportstate ="9"  then %>
					ǰ�ǿϷ�
				<%elseif osheet.FItemList(i).Freportstate = "5" then %>
					ǰ�ǹݷ�
				<%else%>
					������
				<%end if%>
				</a>
			<% end if%>
		</td>
		<td align=center><font color="#777777"><%= Left(osheet.FItemList(i).FRegdate,10) %></font></td>
		<td align=right><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
		<td align=right><%= FormatNumber(osheet.FItemList(i).Fjumunbuycash,0) %></td>
		<td rowspan=2 align=center>
			<% if (Not osheet.FItemList(i).IsSendedSMS) and (osheet.FItemList(i).getScheduleIpgodate="") and (osheet.FItemList(i).Fstatecd="0") then %>
				<input type=button class="button" value="SMS" onclick="sendSMSEmail('<%= osheet.FItemList(i).Ftargetid %>','<%= osheet.FItemList(i).Fidx %>')">
			<% end if %>

			<%= Left(osheet.FItemList(i).getScheduleIpgodate,10) %><br><%= Left(osheet.FItemList(i).Fbeasongdate,10) %>
		</td>
		<td rowspan=2 align=center>
			<a href="<%= DeliverDivTrace(Trim(osheet.FItemList(i).Fsongjangdiv)) %><%= Replace(osheet.FItemList(i).Fsongjangno, "-", "") %>" target=_blank>
				<%= DeliverDivCd2Nm(Trim(osheet.FItemList(i).Fsongjangdiv)) %><br><%= osheet.FItemList(i).Fsongjangno %>
			</a>
		</td>
		<td rowspan=2 align=center>
			<%= osheet.FItemList(i).Fcheckusername %>
		</td>
		<td rowspan=2 align=center>
			<%= osheet.FItemList(i).FworkSecond %>
		</td>
		<td rowspan=2 align=center>
			<% if osheet.FItemList(i).Fstatecd="8" then %>
				<!-- <input type="button" class="button" value="�԰�ó��" onClick="IpgoProc('<%= osheet.FItemList(i).Fidx %>')"> -->
			<% elseif osheet.FItemList(i).Fstatecd="9" then %>
				<%= osheet.FItemList(i).Ffinishname %>
			<% end if %>
		</td>
		<td rowspan=2 width=50 align=center>
			<!--<a href="javascript:PopIpgoSheetforeign('<%'= osheet.FItemList(i).FIdx %>','');"><img src="/images/iexplorer.gif" width=21 border=0></a>
			<a href="javascript:ExcelSheetforeign('<%'= osheet.FItemList(i).FIdx %>','');"><img src="/images/iexcel.gif" width=21 border=0></a>-->
			<a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>','');"><img src="/images/iexplorer.gif" width=21 border=0></a>
			<a href="javascript:ExcelSheet('<%= osheet.FItemList(i).FIdx %>','');"><img src="/images/iexcel.gif" width=21 border=0></a>
		</td>
		<td rowspan=2 width=50 align=center>
			<%
			'/���������� ����(6),�귣�����(7) �ϰ�� ����
			if osheet.FItemList(i).fpurchasetype="6" or osheet.FItemList(i).fpurchasetype="7" then
			%>
				<%'= osheet.FItemList(i).fcurrencyUnit %>
				<input type="button" class="button" value="�ֹ���" onClick="PopOpenorderUTF8('<%= osheet.FItemList(i).FIdx %>','<%= osheet.FItemList(i).Fsitename %>','USD','<%= osheet.FItemList(i).Ftplcompanyid %>')">
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align=center><%= Left(osheet.FItemList(i).Fscheduledate,10) %></td>
	    <td align=right><%= FormatNumber(osheet.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><%= FormatNumber(osheet.FItemList(i).Ftotalbuycash,0) %></td>
	</tr>
	<% next %>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=7 rowspan=2>�Ѱ�</td>
		<td align=right><%= formatNumber(totaljumunsellcash,0) %></td>
		<td align=right><%= formatNumber(totaljumunsuply,0) %></td>
		<td colspan=8 rowspan=2></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align=right><%= formatNumber(totaltotalsellcash,0) %></td>
		<td align=right><%= formatNumber(totalfixsuply,0) %></td>
	</tr>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
        	<% if osheet.HasPreScroll then %>
				<a href="javascript:NextPage('<%= osheet.StartScrollPage-1 %>');">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + osheet.StartScrollPage to osheet.FScrollCount + osheet.StartScrollPage - 1 %>
				<% if i>osheet.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if osheet.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan=20 align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>
</table>

<%
set osheet = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
