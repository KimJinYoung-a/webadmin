<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ҷ�������ǰ����
' History : �̻� ����
'           2021.04.06 �ѿ�� ����(��ǰ���� �Ϻ��ڵ� ���� ���� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim makerid,mode, searchtype, purchasetype, mwdiv, sellyn, onlyisusing, makeruseyn, itemgubun
dim datetype, centermwdiv, monthlymwdiv, yyyy, mm, osummarystock, i, BadOrErrText
	makerid 		= requestcheckvar(request("makerid"),32)
	mode 			= requestcheckvar(request("mode"),32)
	searchtype 		= requestcheckvar(request("searchtype"),3)
	purchasetype 	= requestcheckvar(request("purchasetype"),1)
	mwdiv 			= requestcheckvar(request("mwdiv"),1)
	sellyn 			= requestcheckvar(request("sellyn"),1)
	onlyisusing 	= requestcheckvar(request("onlyisusing"),1)
	makeruseyn	 	= requestcheckvar(request("makeruseyn"),1)
	itemgubun 		= requestcheckvar(request("itemgubun"),3)
	datetype 		= requestcheckvar(request("datetype"),8)
	yyyy 			= requestcheckvar(request("yyyy1"),4)
	mm 				= requestcheckvar(request("mm1"),2)
	centermwdiv		= requestcheckvar(request("centermwdiv"),1)
	monthlymwdiv	= requestcheckvar(request("monthlymwdiv"),1)

if (searchtype = "") then
	searchtype = "bad"
	'datetype = "curr"
	yyyy = Left(now(),4)
	mm   = mid(now(),6,2)
end if

'if (itemgubun = "") then
'	itemgubun = "10"
'end if
datetype = "yyyymm"
' ������ϰ��
if yyyy = Left(now(),4) and mm = mid(now(),6,2) then
	datetype = "curr"
end if

set osummarystock = new CSummaryItemStock
	osummarystock.FRectmakerid = makerid
	osummarystock.FRectSearchType = searchtype
	osummarystock.FRectDatetype   = datetype
	osummarystock.FRectYYYYMM = yyyy+"-"+mm

	'if (datetype = "yyyymm") then
	'	osummarystock.FRectMWDiv = monthlymwdiv
	'else
	'	osummarystock.FRectMWDiv = mwdiv
	'end if
	'osummarystock.FRectlastmwdiv = mwdiv
	osummarystock.FRectMWDiv = mwdiv
	osummarystock.FRectlastmwdiv = monthlymwdiv
	osummarystock.FRectCenterMWDiv = centermwdiv
	osummarystock.FRectSellYN = sellyn
	osummarystock.FRectOnlyIsUsing = onlyisusing
	osummarystock.FRectItemGubun = itemgubun
	osummarystock.FRectPurchaseType = purchasetype
	osummarystock.FRectMakerUseYN = makeruseyn

	if (makerid<>"") then
		osummarystock.FPageSize=500                 ''�߰� 2016/08/04 class�� �����ִµ�.
		osummarystock.GetBadOrErrItemListByBrand
	else
		osummarystock.GetBadOrErrItemListByBrandGroup
	end if

if (searchtype="bad") then
    BadOrErrText = "�ҷ�"
else
    BadOrErrText = "�������"
end if

%>
<script type='text/javascript'>

function PopBadOrErrItemReInput(makerid, acttype) {
	var popwin = window.open('/common/pop_badorerritem_re_input.asp?datetype=<%= datetype %>&yyyy1=<%= yyyy %>&mm1=<%= mm %>&makerid=' + makerid + '&searchtype=<%= searchtype %>&acttype=' + acttype + '&mwdiv=<%= mwdiv %>&itemgubun=<%= itemgubun %>&sellyn=<%= sellyn %>&onlyisusing=<%= onlyisusing %>&makeruseyn=<%= makeruseyn %>&purchasetype=<%=purchasetype%>&centermwdiv=<%= centermwdiv %>&monthlymwdiv=<%= monthlymwdiv %>','PopBadOrErrItemReInput','width=1280,height=800,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function SubmitSearchByBrandNew(makerid, mwdiv, itemgubun) {
	var searchitemgubun = "<%= itemgubun %>";

	if ((itemgubun == "") && (searchitemgubun != "")) {
		itemgubun = searchitemgubun;
	}

	var popwin = window.open('?datetype=<%= datetype %>&yyyy1=<%= yyyy %>&mm1=<%= mm %>&searchtype=<%= searchtype %>&purchasetype=<%= purchasetype %>&onlyisusing=<%= onlyisusing %>&sellyn=<%= sellyn %>&mwdiv=' + mwdiv + '&makerid=' + makerid + '&itemgubun=' + itemgubun + '&centermwdiv=<%= centermwdiv %>' + '&monthlymwdiv=<%= monthlymwdiv %>','SubmitSearchByBrandNew','width=1100,height=600,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'PopItemSellEdit','width=500,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popXL(searchtype, purchasetype, mwdiv, sellyn, onlyisusing, makeruseyn, itemgubun, datetype, yyyy, mm, centermwdiv, monthlymwdiv) {
	var popwin = window.open("/admin/stock/badorerritem_xl_download.asp?searchtype=" + searchtype + "&purchasetype=" + purchasetype + "&mwdiv=" + mwdiv + "&sellyn=" + sellyn + "&onlyisusing=" + onlyisusing + "&makeruseyn=" + makeruseyn + "&itemgubun=" + itemgubun + "&datetype=" + datetype + "&yyyy1=" + yyyy + "&mm1=" + mm + "&centermwdiv=" + centermwdiv + "&monthlymwdiv=" + monthlymwdiv,"popXL","width=300,height=200 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ChangePage(v) {
	var frm = document.frm;

	frm.submit();
}

function jsSetBrandAll() {
    var frm = document.frm;
    frm.makerid.value = 'all';
    document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���ؿ� :
			<% ' �̹��� �̻�� ��û���� ���ܽ�Ŵ. Ư������ �������� ������ ���� ��� ���ذ� ������ �ƴ�?	' 2021.04.15 �ѿ�� %>
			<!--<input type="radio" name="datetype" value="curr" <% if (datetype = "curr") then %>checked<% end if %>> �������
			<input type="radio" name="datetype" value="yyyymm" <% if (datetype = "yyyymm") then %>checked<% end if %>> Ư���������� -->
			<% Call DrawYMBox(yyyy, mm) %>
			&nbsp;
			<input type="radio" name="searchtype" value="bad" <% if (searchtype = "bad") then %>checked<% end if %>> �ҷ���ǰ
			<input type="radio" name="searchtype" value="err" <% if (searchtype = "err") then %>checked<% end if %>> ������ϻ�ǰ
			&nbsp;
			�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
            <input type="button" class="button" value="��ü�귣��" onClick="jsSetBrandAll()">
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<b>�귣�� ����</b>
			&nbsp;
			&nbsp;
			��뿩�� :
			<select class="select" name="makeruseyn">
				<option value="">-����-</option>n
				<option value="Y" <% if (makeruseyn = "Y") then %>selected<% end if %> >�����</option>
				<option value="N" <% if (makeruseyn = "N") then %>selected<% end if %> >������</option>
			</select>
			&nbsp;
			�������� : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<b>��ǰ ����</b>
			&nbsp;
			&nbsp;
			��ǰ���� :
			<select class="select" name="itemgubun">
				<option value="">-����-</option><% '�̹��� �̻�� ��û���� ��ü �߰� %>
				<option value="10" <% if (itemgubun = "10") then %>selected<% end if %> >�»�ǰ(10)</option>
				<option value="OFF" <% if (itemgubun = "OFF") then %>selected<% end if %> >������ü</option>
				<option value="55" <% if (itemgubun = "55") then %>selected<% end if %> >����(55)</option>
				<option value="70" <% if (itemgubun = "70") then %>selected<% end if %> >����(70)</option>
				<option value="75" <% if (itemgubun = "75") then %>selected<% end if %> >����(75)</option>
				<option value="80" <% if (itemgubun = "80") then %>selected<% end if %> >����(80)</option>
				<option value="85" <% if (itemgubun = "85") then %>selected<% end if %> >����(85)</option>
				<option value="90" <% if (itemgubun = "90") then %>selected<% end if %> >����(90)</option>
			</select>
			&nbsp;
            <%'= CHKIIF(datetype<>"yyyymm", "ON���Ա���(����)", "<del>ON���Ա���(����)</del>") %>ON���Ա���(����) :
			<select class="select" name="mwdiv">
				<option value="">-����-</option>
				<option value="M" <% if (mwdiv = "M") then %>selected<% end if %> >����</option>
				<option value="W" <% if (mwdiv = "W") then %>selected<% end if %> >Ư��</option>
				<option value="U" <% if (mwdiv = "U") then %>selected<% end if %> >��ü</option>
				<option value="Z" <% if (mwdiv = "Z") then %>selected<% end if %> >������</option>
			</select>
			&nbsp;
            ���͸��Ա���(����) :
     		<select class="select" name="centermwdiv">
				<option value="">����</option>
				<option value="M" <%= CHKIIF(centermwdiv="M","selected","")%> >����</option>
				<option value="W" <%= CHKIIF(centermwdiv="W","selected","")%> >��Ź</option>
				<option value="X" <%= CHKIIF(centermwdiv="X","selected","")%> >������</option>
			</select>
     		&nbsp;
            <%'= CHKIIF(datetype="yyyymm", "���Ա���(����)", "<del>���Ա���(����)</del>") %>���Ա���(���) :
     		<select class="select" name="monthlymwdiv">
				<option value="">����</option>
				<option value="M" <%= CHKIIF(monthlymwdiv="M","selected","")%> >����</option>
				<option value="W" <%= CHKIIF(monthlymwdiv="W","selected","")%> >��Ź</option>
				<option value="X" <%= CHKIIF(monthlymwdiv="X","selected","")%> >������</option>
			</select>
			&nbsp;
			�Ǹſ���(����) :
			<select class="select" name="sellyn">
				<option value="">-����-</option>
				<option value="Y" <% if (sellyn = "Y") then %>selected<% end if %> >�Ǹ���</option>
				<option value="N" <% if (sellyn = "N") then %>selected<% end if %> >�Ǹž���</option>
			</select>
            &nbsp;
			��뿩��(����) :
			<select class="select" name="onlyisusing">
				<option value="">-����-</option>
				<option value="Y" <% if (onlyisusing = "Y") then %>selected<% end if %> >�����</option>
				<option value="N" <% if (onlyisusing = "N") then %>selected<% end if %> >������</option>
			</select>
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->

<br>

* �귣�� �� ��ǰ������ <font color="red">���������� ����</font>���� �մϴ�.(Ư���� �귣������ �� ��ǰ���� �������)

<% if makerid<>"" then %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if (searchtype = "bad") then %>
				<input type="button" class="button" value="��ǰ" onclick="PopBadOrErrItemReInput('<%= makerid %>', 'actreturn')" border="0">
				&nbsp;
			<% end if %>
        	<input type="button" class="button" value="�ν����" onclick="PopBadOrErrItemReInput('<%= makerid %>', 'actloss')" border="0">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= osummarystock.FTotalCount %></b>
			<% if (osummarystock.FResultCount>=osummarystock.FPageSize) then %>�ִ� <%=osummarystock.FPageSize%> �� ǥ��<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�귣��ID</td>
		<td width="50">�̹���</td>
		<td width="40">���<br>����<br>����</td>
		<td width="40">ON<br>����<br>����</td>
		<td width="40">����<br>����<br>����</td>
		<td width="30">��ǰ<br>����</td>
		<td width="50">��ǰ�ڵ�</td>
		<td width="40">�ɼ�</td>
		<td>��ǰ��<br><font color="blue">[�ɼǸ�]</font></td>

		<td width="50">�Һ��ڰ�</td>
		<td width="50">���԰�</td>
		<td width="30">�Ǹ�<br>����</td>
		<td width="30">���<br>����</td>
		<td width="60"><%= BadOrErrText %><br>����</td>
		<td width="70">���԰���</td>
		<td width="60">�ǻ�<br>��ȿ���</td>
    </tr>
	<% if osummarystock.FResultCount>0 then %>
	<% for i=0 to osummarystock.FResultCount - 1 %>
	<% if (osummarystock.FItemList(i).Fisusing = "Y") then %>
		<tr bgcolor="#FFFFFF" height="30">
	<% else %>
		<tr bgcolor="#BBBBBB" height="30">
	<% end if %>
    	<td><%= osummarystock.FItemList(i).Fmakerid %></td>
    	<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" ></td>
		<td align="center" style="color:<%=GetMwDivColorCd(osummarystock.FItemList(i).flastmwdiv)%>;"><%= osummarystock.FItemList(i).flastmwdiv %></td>
    	<td align="center" style="color:<%=osummarystock.FItemList(i).GetMwDivColor%>;"><%= osummarystock.FItemList(i).Fmwdiv %></td>
		<td align="center" style="color:<%=GetMwDivColorCd(osummarystock.FItemList(i).Fcentermwdiv)%>;"><%= osummarystock.FItemList(i).Fcentermwdiv %></td>
    	<td align="center"><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td align="center"><a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemid %>');"><%= osummarystock.FItemList(i).FItemid %></a></td>
		<td align="center"><%= osummarystock.FItemList(i).FItemoption %></td>
		<td align="left"><a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= osummarystock.FItemList(i).FItemgubun %>&itemid=<%= osummarystock.FItemList(i).FItemid %>&itemoption=<%= osummarystock.FItemList(i).FItemoption %>" target=_blank ><%= osummarystock.FItemList(i).FItemname %></a><br><font color="blue">[<%= osummarystock.FItemList(i).FItemOptionName %>]</font></td>

		<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td align="right"><%= formatnumber(osummarystock.FItemList(i).Fbuycash,0) %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fsellyn %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fisusing %></td>
		<td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fregitemno, 0) %></td>
		<td align="right"><%= formatnumber((osummarystock.FItemList(i).Fbuycash * osummarystock.FItemList(i).Fregitemno),0) %></td>
		<td align="center"><%= FormatNumber(osummarystock.FItemList(i).Frealstock, 0) %></td>
    </tr>
    <% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
</table>

<% else %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="right">
			<% if (searchtype = "bad") then %>
			<input type="button" class="button" value="�����ٿ�ε�(�ҷ�)" onclick="popXL('bad', '<%= purchasetype %>', '<%= mwdiv %>', '<%= sellyn %>', '<%= onlyisusing %>', '<%= makeruseyn %>', '<%= itemgubun %>', '<%= datetype %>', '<%= yyyy %>', '<%= mm %>', '<%= centermwdiv %>', '<%= monthlymwdiv %>')">
			<% end if %>
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><%= osummarystock.FResultCount %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td rowspan="3">�귣��</td>
		<td rowspan="3">�귣���</td>
		<td rowspan="3">��ü��</td>
		<td width="40" rowspan="3">�귣��<br>���<br>����</td>
		<td colspan="11"><%= BadOrErrText %>��ǰ����</td>
		<td rowspan="3">���</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td colspan="4">10</td>
		<td colspan="3">90</td>
		<td colspan="3">��Ÿ</td>
		<td rowspan="2" width="80">�Ұ�</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="55">����</td>
		<td width="55">��Ź</td>
		<td width="55">����</td>
		<td width="55">������</td>
		<td width="55">����</td>
		<td width="55">��Ź</td>
		<td width="55">������</td>
		<td width="55">����</td>
		<td width="55">��Ź</td>
		<td width="55">������</td>
	</tr>
	<% if osummarystock.FResultCount>0 then %>
	<% for i=0 to osummarystock.FResultCount-1 %>
	<% if (osummarystock.FItemList(i).Fuseyn = "Y") then %>
		<tr bgcolor="#FFFFFF" height="30">
	<% else %>
		<tr bgcolor="#BBBBBB" height="30">
	<% end if %>
	    <td><a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', '', '');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td><%= osummarystock.FItemList(i).Fmakername %></td>
		<td align="left"><%= osummarystock.FItemList(i).Fcompany_name %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fuseyn %></td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'M', '10');">
	    	<% if (osummarystock.FItemList(i).Fitem10M <> 0) then %><font color="red"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem10M, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'W', '10');">
	    	<% if (osummarystock.FItemList(i).Fitem10W <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem10W, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'U', '10');">
	    	<% if (osummarystock.FItemList(i).Fitem10U <> 0) then %><font color="green"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem10U, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'Z', '10');">
	    	<% if (osummarystock.FItemList(i).Fitem10U <> 0) then %><font color="black"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem10Z, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'M', '90');">
	    	<% if (osummarystock.FItemList(i).Fitem90M <> 0) then %><font color="red"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem90M, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'W', '90');">
	    	<% if (osummarystock.FItemList(i).Fitem90W <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem90W, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'Z', '90');">
	    	<% if (osummarystock.FItemList(i).Fitem90W <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).Fitem90Z, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'M', 'OFF');">
	    	<% if (osummarystock.FItemList(i).FitemetcM <> 0) then %><font color="red"><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).FitemetcM, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'W', 'OFF');">
	    	<% if (osummarystock.FItemList(i).FitemetcW <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).FitemetcW, 0) %></a>
	    </td>
	    <td align="center">
	    	<a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>', 'Z', 'OFF');">
	    	<% if (osummarystock.FItemList(i).FitemetcW <> 0) then %><b><% end if %>
	    	<%= FormatNumber(osummarystock.FItemList(i).FitemetcZ, 0) %></a>
	    </td>
	    <td align="center">
	    	<% if ((osummarystock.FItemList(i).FOnCnt + osummarystock.FItemList(i).FOffCnt) <> 0) then %><b><% end if %>
	    	<%= FormatNumber((osummarystock.FItemList(i).FOnCnt + osummarystock.FItemList(i).FOffCnt), 0) %>
	    </td>
	    <td align="left">
	    	<% if (searchtype = "bad") then %>
				<input type="button" class="button" value="��ǰ" onclick="PopBadOrErrItemReInput('<%= osummarystock.FItemList(i).FMakerid %>', 'actreturn')" border="0" <% if (osummarystock.FItemList(i).Fcompany_no = "211-87-00620" and Left(Now(), 7) <> "2022-07") then %>disabled<% end if %> >
				&nbsp;
				<input type="button" class="button" value="�������" onclick="PopBadOrErrItemReInput('<%= osummarystock.FItemList(i).FMakerid %>', 'actshopchulgo')" border="0">
				&nbsp;
			<% end if %>
			<% if (searchtype = "bad") then %>
        	<input type="button" class="button" value="���ó��" onclick="PopBadOrErrItemReInput('<%= osummarystock.FItemList(i).FMakerid %>', 'actloss')" border="0">
			<% else %>
			<input type="button" class="button" value="�ν����" onclick="PopBadOrErrItemReInput('<%= osummarystock.FItemList(i).FMakerid %>', 'actloss')" border="0">
			<% end if %>
	    </td>
	</tr>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="16" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
</table>
<% end if %>

<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
