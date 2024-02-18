<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� �԰���Ʈ
' History : 2007.01.01 �̻� ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim code, blinkcode, minusjumun, page,designer, statecd, onoffgubun, divcode, rackipgoyn, ipgocheck, yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim fromDate, toDate, vPurchaseType, searchType, searchText, totalsellcash,totalsuply, totalitemno, i
dim alinkcode, linkparam
	page = request("page")
	designer = request("designer")
	statecd = request("statecd")
	code = request("code")				' �԰� �ڵ�
	alinkcode = request("alinkcode")
	blinkcode = request("blinkcode")
	onoffgubun = request("onoffgubun")	' ��/���� ����
	divcode = request("divcode")		' ���� ����
	rackipgoyn = request("rackipgoyn")	'
	vPurchaseType = requestCheckVar(request("purchasetype"),3)
	searchType = request("searchType")
	searchText = request("searchText")
	minusjumun = request("minusjumun")

	'// �԰��� �˻��� �ʿ��� ���� ����
	ipgocheck = request("ipgocheck")
	yyyy1 = request("yyyy1")
	yyyy2 = request("yyyy2")
	mm1	  = request("mm1")
	mm2	  = request("mm2")
	dd1	  = request("dd1")
	dd2	  = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if onoffgubun="" then onoffgubun="all"

if page="" then page=1

code = Trim(code)
blinkcode = Trim(blinkcode)

linkparam="&designer="&designer&"&searchType="&searchType&"&searchText="&searchText&"&onoffgubun="&onoffgubun&"&divcode="&divcode&"&rackipgoyn="&rackipgoyn
linkparam=linkparam & "&purchasetype="&vpurchasetype&"&minusjumun="&minusjumun&"&code="&"&alinkcode="&alinkcode&"&blinkcode="&"&ipgocheck="&ipgocheck

dim oipchul
set oipchul = new CIpChulStorage
	oipchul.FCurrPage = page
	oipchul.Fpagesize=50
	oipchul.FRectCode = code
	oipchul.FRectBLinkCode = blinkcode
	oipchul.FRectALinkCode = alinkcode
	oipchul.FRectDivcode = divcode
	oipchul.FRectRackipgoyn = rackipgoyn
	oipchul.FRectMinusOnly = minusjumun

	if ipgocheck="on" then
		oipchul.FRectExecuteDtStart = fromDate
		oipchul.FRectExecuteDtEnd   = toDate
	end if

	if code="" then
	oipchul.FRectCodeGubun = "ST"  ''�԰�
	oipchul.FRectSocID = designer
	oipchul.FRectOnOffGubun = onoffgubun
	end if

	oipchul.FRectSearchType = searchType
	oipchul.FRectSearchText = searchText

	oipchul.FRectBrandPurchaseType = vPurchaseType
	oipchul.GetIpChulgoList

totalsellcash = 0
totalsuply	  = 0
totalitemno = 0
%>

<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script src="/js/jquery.placeholder.min.js"></script>
<script type="text/javascript">

function PopUpcheBrandInfoEdit(v){
	window.open("/admin/member/popupchebrandinfo.asp?designer=" + v,"PopUpcheBrandInfoEdit","width=640,height=580,scrollbars=yes,resizabled=yes");
}

function IpgoInput(){
	location.href="/admin/newstorage/ipgoinput.asp?menupos=<%= menupos %>";
}

function popipgocheck(iidx){
	var popwin = window.open("poplimitcheckipgoNew.asp?idx=" + iidx ,"popipgoproc","width=900,height=550,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function PopIpgoSheet(v,itype){
	var popwin;
	popwin = window.open('popipgosheet.asp?idx=' + v + '&itype=' + itype,'popipgosheet','width=760,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function ExcelSheet(v,itype){
	window.open('popipgosheet.asp?idx=' + v + '&itype=' + itype + '&xl=on');
}

function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}

function NextPage(page){
	ClearPlaceHolder();
	document.frm.page.value = page;
	document.frm.submit();
}

function popXL(fromDate, toDate) {
	<% if ipgocheck<>"on" then %>
	alert("���� �԰����� �����ϼ���.");
	return;
	<% end if %>

	var popwin = window.open("/admin/newstorage/pop_ipgolist_xl_download.asp?fromDate=" + fromDate + "&toDate=" + toDate + "<%=linkparam%>","popXL","width=100,height=100 scrollbars=yes resizable=yes");
	popwin.focus();
}

function SubmitFrm(frm) {
	ClearPlaceHolder();
	if (frm.code.value.length > 0) {
		if (frm.code.value.substring(0,2).toUpperCase() != "ST") {
			alert("�԰��ڵ尡 �ƴմϴ�.");
			return;
		}
	}

	frm.submit();
}

function ClearPlaceHolder() {
	var frm = document.frm;
	frm.code.value = $('#code').val();
	frm.blinkcode.value = $('#blinkcode').val();
}

function popOpenPPMaster(idx) {
	var popwin;

	popwin = window.open('/admin/newstorage/PurchasedProductModify.asp?menupos=9175&idx=' + idx ,'popOpenPPMaster','width=1400,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
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
		* �԰��ڵ� :
		<textarea class="textarea" id="code" name="code" cols="12" rows="1" placeholder="�ִ�50��"><%= code %></textarea>
		&nbsp;
		* �ֹ��ڵ� :
		<textarea class="textarea" id="blinkcode" name="blinkcode" cols="12" rows="1" placeholder="�ִ�50��"><%= blinkcode %></textarea>
		&nbsp;
		* �ֹ��ڵ�(����) : <input type="text" class="text" name="alinkcode" value="<%= alinkcode %>" size="8" maxlength="8">
		&nbsp;
		<input type="checkbox" name="ipgocheck" <% if ipgocheck="on" then  response.write "checked" %> onclick="EnDisabledDateBox(this)">�԰���
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitFrm(frm)">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �¿�������:
		<input type="radio" name="onoffgubun" value="all" <% if onoffgubun="all" then response.write "checked" %> >��ü
		<input type="radio" name="onoffgubun" value="on" <% if onoffgubun="on" then response.write "checked" %> >�¶���
		<input type="radio" name="onoffgubun" value="off" <% if onoffgubun="off" then response.write "checked" %> >��������
		&nbsp;
		* ���Ա���:
		<input type="radio" name="divcode" value="" <% if divcode="" then response.write "checked" %> >��ü
		<input type="radio" name="divcode" value="001" <% if divcode="001" then response.write "checked" %> >����
		<input type="radio" name="divcode" value="002" <% if divcode="002" then response.write "checked" %> >��Ź
		<input type="radio" name="divcode" value="801" <% if divcode="801" then response.write "checked" %> >Off����
		<input type="radio" name="divcode" value="802" <% if divcode="802" then response.write "checked" %> >Off��Ź
		&nbsp;
		* �������� : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �˻����� :
		<select class="select" name="searchType">
			<option value="" >��ü</option>
			<option value="socname" <% if (searchType = "socname") then %>selected<% end if %> >��ü��</option>
			<option value="socno" <% if (searchType = "socno") then %>selected<% end if %> >����ڹ�ȣ</option>
		</select>
		<input type="text" class="text" name=searchText value="<%= searchText %>" size="15" maxlength="20">
		&nbsp;
		<input type="checkbox" name="minusjumun" <% if minusjumun="on" then response.write "checked" %> >���̳ʽ��ֹ���
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
		<input type="button" class="button" value=" �԰��Է� " onclick="IpgoInput();">
	</td>
	<td align="right">
		<% if oipchul.FTotalCount > 0 then %>
			<input type="button" class="button" value=" �����ޱ� " onclick="popXL('<%= fromDate %>', '<%= toDate %>');">
		<% end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="left">
		�˻���� : <b><%= oipchul.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oipchul.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=60>�԰��ڵ�</td>
	<td width=60>�ֹ��ڵ�</td>
	<td width="60">����IDX</td>
	<td width=60>��������</td>
	<td>����óID</td>
	<td>����ó</td>
	<td width=50>ó����</td>
	<td width=70>������</td>
	<td width=70>�԰���</td>
	<td width=60>�Һ��ڰ�</td>
	<td width=60>���԰�</td>
	<td width=40>����</td>
	<td width=50>����</td>
	<td width=50>����</td>
	<td width=60>�԰�ó��</td>
	<td width=50>������</td>
</tr>
<% if oipchul.FResultCount >0 then %>
<% for i=0 to oipchul.FResultcount-1 %>
<%
totalsellcash = totalsellcash + oipchul.FItemList(i).Ftotalsellcash
totalsuply	  = totalsuply + oipchul.FItemList(i).Ftotalsuplycash
totalitemno	  = totalitemno + oipchul.FItemList(i).ftotalitemno
%>
<tr bgcolor="#FFFFFF" height=24>
	<td align=center><a href="ipgodetail.asp?idx=<%= oipchul.FItemList(i).Fid %>&opage=<%= page %>&menupos=<%=menupos%>"><%= oipchul.FItemList(i).Fcode %></a></td>
	<td align=center>
		<% if Not IsNull(oipchul.FItemList(i).Fblinkcode) then %>
		<a href="/admin/newstorage/orderlist.asp?menupos=537&baljucode=<%= oipchul.FItemList(i).Fblinkcode %>" target="_blank"><%= oipchul.FItemList(i).Fblinkcode %></a>
		<% elseif Not IsNull(oipchul.FItemList(i).Falinkcode) then %>
		<a href="/admin/fran/upchejumunlist.asp?menupos=530&baljucode=<%= oipchul.FItemList(i).Falinkcode %>" target="_blank"><%= oipchul.FItemList(i).Falinkcode %></a>
		<% end if %>
	</td>
	<td align="center">
		<% if (oipchul.FItemList(i).FppMasterIdx <> "" and not(isnull(oipchul.FItemList(i).FppMasterIdx))) then %>
			<a href="#" onclick="popOpenPPMaster(<%= oipchul.FItemList(i).FppMasterIdx %>); return false;"><%= oipchul.FItemList(i).FppMasterIdx %></a>
		<% end if %>
	</td>
	<td align=left><%= oipchul.FItemList(i).fpurchasetypename %></td>
	<td align=left><b><a href="javascript:PopUpcheBrandInfoEdit('<%= oipchul.FItemList(i).Fsocid %>');"><%= oipchul.FItemList(i).Fsocid %></a></b></td>
	<td align=left><%= oipchul.FItemList(i).Fsocname %></td>
	<td align=center><%= oipchul.FItemList(i).Fchargename %></td>
	<td align=center><font color="#777777"><%= Left(oipchul.FItemList(i).Fscheduledt,10) %></font></td>
	<td align=center><%= Left(oipchul.FItemList(i).Fexecutedt,10) %></td>
	<td align=right><font color="<%= oipchul.FItemList(i).GetMinusColor(oipchul.FItemList(i).Ftotalsellcash) %>"><%= FormatNumber(oipchul.FItemList(i).Ftotalsellcash,0) %></font></td>
	<td align=right><font color="<%= oipchul.FItemList(i).GetMinusColor(oipchul.FItemList(i).Ftotalsuplycash) %>"><%= FormatNumber(oipchul.FItemList(i).Ftotalsuplycash,0) %></font></td>
	<td align="right">
		<font color="<%= oipchul.FItemList(i).GetMinusColor(oipchul.FItemList(i).ftotalitemno) %>"><%= FormatNumber(oipchul.FItemList(i).ftotalitemno,0) %></font>
	</td>
	<td align=center><font color="<%= oipchul.FItemList(i).GetDivCodeColor %>"><%= oipchul.FItemList(i).GetDivCodeName %></font></td>
	<td align=right>
	<% if oipchul.FItemList(i).Ftotalsellcash<>0 then %>
	  <%= 100-CLng(oipchul.FItemList(i).Ftotalsuplycash/oipchul.FItemList(i).Ftotalsellcash*100*100)/100 %>%
	<% end if %>
	</td>
	<td align=center>
		<input type="button" class="button" value="�԰�ó��" onClick="popipgocheck('<%= oipchul.FItemList(i).Fid %>')">
	</td>
	<td>
          <a href="javascript:PopIpgoSheet('<%= oipchul.FItemList(i).Fid %>','');"><img src="/images/iexplorer.gif" width=21 border=0></a> <a href="javascript:ExcelSheet('<%= oipchul.FItemList(i).Fid %>','');"><img src="/images/iexcel.gif" width=21 border=0></a>
    </td>
</tr>
<% next %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align=center>�Ѱ�</td>
	<td colspan=8></td>
	<td align=right><%= formatNumber(totalsellcash,0) %></td>
	<td align=right><%= formatNumber(totalsuply,0) %></td>
	<td align="right">
		<%= formatNumber(totalitemno,0) %>
	</td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>

<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=20 align=center>[ �˻������ �����ϴ�. ]</td>
</tr>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if oipchul.HasPreScroll then %>
    		<a href="javascript:NextPage('<%= oipchul.StartScrollPage-1 %>')">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oipchul.StartScrollPage to oipchul.FScrollCount + oipchul.StartScrollPage - 1 %>
    		<% if i>oipchul.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oipchul.HasNextScroll then %>
    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
	</td>
</tr>
</table>


<%
set oipchul = Nothing
%>

<script type="text/javascript">
	EnDisabledDateBox(document.frm.ipgocheck);
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
