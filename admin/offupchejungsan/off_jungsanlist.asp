<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������������
' History : ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%

dim makerid, yyyy1, mm1, finishflag, page, groupid, vPurchaseType, jgubun, jacctcd, differencekey
dim searchType, searchText, jungsanGubun
makerid 		= requestCheckVar(request("makerid"),32)
yyyy1   		= requestCheckVar(request("yyyy1"),10)
mm1     		= requestCheckVar(request("mm1"),10)
finishflag      = requestCheckVar(request("finishflag"),10)
page            = requestCheckVar(request("page"),10)
vPurchaseType   = requestCheckVar(request("purchasetype"),10)
jgubun          = requestCheckVar(request("jgubun"),10)
jacctcd 		= requestCheckVar(request("jacctcd"),10)
differencekey 	= requestCheckVar(request("differencekey"),10)
searchType 		= requestCheckVar(request("searchType"), 32)
searchText 		= requestCheckVar(request("searchText"), 32)
jungsanGubun    = requestCheckVar(request("jungsanGubun"), 12)

dim comm_cd : comm_cd     = RequestCheckVar(request("comm_cd"),9)

if page="" then page=1


dim taxtype, autojungsan, jungsan_date
taxtype      = requestCheckVar(request("taxtype"),32)
autojungsan  = requestCheckVar(request("autojungsan"),32)
jungsan_date = requestCheckVar(request("jungsan_date"),32)
groupid      = requestCheckVar(request("groupid"),32)

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

''�ӽ�
''yyyy1 = "2006"
''mm1="12"


dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FPageSize   = 50
ooffjungsan.FCurrPage = page
ooffjungsan.FRectYYYYMM = yyyy1 + "-" + mm1
ooffjungsan.FRectfinishflag = finishflag
ooffjungsan.FRectMakerid = makerid
ooffjungsan.FRectTaxtype = taxtype
ooffjungsan.FRectAutojungsan = autojungsan
ooffjungsan.FRectJungsanDate = jungsan_date
ooffjungsan.FRectGroupID = groupid
ooffjungsan.FRectPurchaseType = vPurchaseType
ooffjungsan.FRectJungsanGubunCD = comm_cd
ooffjungsan.FRectJGubun = jgubun
ooffjungsan.FRectjacctcd = jacctcd
ooffjungsan.FRectdifferencekey = differencekey
ooffjungsan.FRectSearchType = searchType
ooffjungsan.FRectSearchText = searchText
ooffjungsan.FRectJungsanGubun = jungsanGubun
ooffjungsan.GetOffJungsanMasterList



dim i
dim orgsellmargin, realsellmargin
orgsellmargin   = 0
realsellmargin  = 0
%>
<script language='javascript'>
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function MakeBrandBatchJungsan(frm){
    if (frm.jgubun.value.length<1){
        alert('���� ��� ������ ���� �ϼ���.');
        frm.jgubun.focus();
        return;
    }

    if (frm.differencekey.value.length<1){
        alert('���� ������ ���� �ϼ���.');
        frm.differencekey.focus();
        return;
    }

    if (frm.itemvatYN.value.length<1){
        alert('��ǰ ���� ������ ���� �ϼ���.');
        frm.itemvatYN.focus();
        return;
    }

    if (confirm('���곻���� �ۼ� �Ͻðڽ��ϱ�?')){
        var queryurl = 'off_jungsan_process.asp?mode=brandbatchprocess&jgubun='+frm.jgubun.value+'&makerid=' + frm.makerid.value + '&yyyy=' + frm.yyyy.value + '&mm=' + frm.mm.value + '&differencekey=' + frm.differencekey.value + '&itemvatYN=' + frm.itemvatYN.value+'&ipchulArr='+frm.ipchulArr.value;

        var popwin = window.open(queryurl ,'off_jungsan_process','width=200, height=200, scrollbars=yes, resizable=yes');
    }
}

function PopDetail(idx){
    var popwin = window.open('off_jungsandetailsum.asp?idx=' + idx ,'off_jungsandetailsum','width=960, height=540, scrollbars=yes, resizable=yes');
    popwin.focus();
}

function PopStateChange(idx){
    var popwin = window.open('off_jungsanstateedit.asp?idx=' + idx ,'off_jungsanstateedit','width=960, height=540, scrollbars=yes, resizable=yes');
    popwin.focus();
}

function DelMaster(idx){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        var popwin = window.open('off_jungsan_process.asp?mode=delmaster&masteridx=' + idx ,'off_jungsan_process','width=200, height=200, scrollbars=yes, resizable=yes');
    }
}

function PopTaxPrintReDirect(itax_no, makerid){
	var popwinsub = window.open("/admin/upchejungsan/red_taxprint.asp?tax_no=" + itax_no + "&makerid=" + makerid,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}

function research(t){

}
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
    <tr align="center" bgcolor="#FFFFFF" >
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
        <td align="left">
            ������ : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;
            ��꼭�������� :
            <select name="taxtype" class="select">
                <option value="" <% if taxtype="" then response.write "selected" %> >����
                <option value="01" <% if taxtype="01" then response.write "selected" %> >����
                <option value="02" <% if taxtype="02" then response.write "selected" %> >�鼼
                <option value="03" <% if taxtype="03" then response.write "selected" %> >����
            </select>&nbsp;&nbsp;
            <!--
            ���ⱸ�� :
            <select name="autojungsan">
            <option value=""  <% if autojungsan="" then response.write "selected" %> >����
            <option value="Y" <% if autojungsan="Y" then response.write "selected" %> >�ڵ�
            <option value="N" <% if autojungsan="N" then response.write "selected" %> >����
            </select>&nbsp;&nbsp;
            -->
            ������ :
            <select name="jungsan_date" class="select">
                <option value="" <% if jungsan_date="" then response.write "selected" %> >����
                <option value="15��" <% if jungsan_date="15��" then response.write "selected" %> >15��
                <option value="����" <% if jungsan_date="����" then response.write "selected" %> >����
                <option value="����" <% if jungsan_date="����" then response.write "selected" %> >����
                <option value="NULL" <% if jungsan_date="NULL" then response.write "selected" %> >������
            </select>
            &nbsp;&nbsp;
            ������� : <% DrawOffJungsanStateCombo "finishflag", finishflag %>
            &nbsp;&nbsp;
            ��ü�������� : 
            <select name="jungsanGubun" class="select">
                <option value="" <% if jungsanGubun="" then response.write "selected" %>>��ü</option>
                <option value="�Ϲݰ���" <% if jungsanGubun="�Ϲݰ���" then response.write "selected" %>>�Ϲݰ���</option>
                <option value="���̰���" <% if jungsanGubun="���̰���" then response.write "selected" %>>���̰���</option>
                <option value="��õ¡��" <% if jungsanGubun="��õ¡��" then response.write "selected" %>>��õ¡��</option>
                <option value="�鼼" <% if jungsanGubun="�鼼" then response.write "selected" %>>�鼼</option>
                <option value="����(�ؿ�)" <% if jungsanGubun="����(�ؿ�)" then response.write "selected" %>>����(�ؿ�)</option>
            </select>
        </td>
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
    		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
    	</td>
    </tr>
    <tr>
        <td bgcolor="#FFFFFF" >
        �������� :
        <% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
        &nbsp;&nbsp;
		�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
		&nbsp;&nbsp;
		��ü(�׷��ڵ�) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >&nbsp;&nbsp;
        ���������ڵ� : <input type="text" class="text" name="jacctcd" value="<%= jacctcd %>" size="7" >

        </td>
    </tr>
    <tr>
        <td bgcolor="#FFFFFF" >
			�����ı��� :
			<% drawSelectBoxJGubun "jgubun",jgubun %>
			���걸�� :
			<% drawSelectBoxOFFJungsanCommCDQuery "comm_cd",comm_cd %>
			&nbsp;&nbsp;
			����
			<input type="text" class="text" name="differencekey" value="<%= differencekey %>" size="2" >
			&nbsp;&nbsp;
			�˻�����:
			<select class="select" name="searchType">
				<option></option>
				<option value="socname" <% if (searchType = "socname") then %>selected<% end if %> >��ü��</option>
				<option value="socno" <% if (searchType = "socno") then %>selected<% end if %> >����ڹ�ȣ</option>
			</select>
			&nbsp;
			<input type="text" class="text" name=searchText value="<%= searchText %>" size="15" maxlength="20">
        </td>
    </tr>
    </form>
</table>
<p>
<!-- ǥ ��ܹ� ��-->
<% if (makerid<>"") and (yyyy1<>"") and (mm1<>"") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="brandbatch" >
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
<tr bgcolor="#FFFFFF">
    <td>
        <select name="jgubun" class="select">
            <option value="">���� ��� ����</option>
            <option value="MM">����</option>
            <option value="CC">������</option>
            <option value="CE">��Ÿ����</option>
        </select>
        <select name="differencekey" class="select">
            <option value="">���� ����
            <option value="0">0��
            <option value="1">1��
            <option value="2">2��
            <option value="3">3��
            <option value="4">4��
            <option value="5">5��
            <option value="6">6��
            <option value="7">7��
            <option value="8">8��
            <option value="9">9��
        </select>
        <select name="itemvatYN" class="select">
            <option value="">��ǰ ���� ���� ����
            <option value="Y">����
            <option value="N">�鼼
        </select>
        <!--
        &nbsp;�����ڵ�<input type="text" name="ipchulArr" value="" size="20">
        -->
        <input type="hidden" name="ipchulArr" value="">
        <input type="button" value=" <%= makerid %> &nbsp;&nbsp;<%= yyyy1 %>�� <%= mm1 %>�� ���� �ۼ� " onClick="MakeBrandBatchJungsan(brandbatch);">
    </td>
</form>
</tr>
</table>
<% end if %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="<%= adminColor("topbar") %>">
        <td colspan="25" align="right" >
            �ѰǼ�: <%= FormatNumber(ooffjungsan.FTotalCount,0) %> &nbsp;&nbsp;
            �ѱݾ�: <%= FormatNumber(ooffjungsan.FTotalSum,0) %> &nbsp;&nbsp;
            Page: <%= page %>/<%= ooffjungsan.FTotalPage %>
        </td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="50">�����</td>
      <td width="50">����<br>���</td>
      <td width="60">����<br>����</td>
      <td width="30">����</td>
      <td width="30">����<br>(��꼭)</td>
      <td width="30">����<br>(��ǰ)</td>
      <td width="90"><a href="javascript:research(frm,'makerid')">�귣��ID</a></td>
      <td width="70">�׷�ID</td>
      <td width="56">��Ź<br>�Ǹ�</td>
      <td width="56">��ü<br>��Ź</td>
      <td width="56">����<br>����</td>
      <td width="56">����<br>����</td>
      <td width="56">���<br>����</td>
      <td width="56">��Ÿ<br>����</td>

      <td width="70">���ǸŰ�</td>
      <td width="70">�Ѹ����</td>
      <td width="70">�������</td>
      <td width="70">�Ѽ�����</td>
      <td width="60">�Һ�<br>����</td>
      <td width="60">����<br>����</td>
      <td width="80"><a href="javascript:research(frm,'state')">����</a></td>
      <td width="70"><a href="javascript:research(frm,'segum')">����<br>������</a></td>
      <td width="70">�Ա���</td>
      <td width="60">��������</td>
      <td width="30">���</td>
    </tr>
    <% if ooffjungsan.FResultCount>0 then %>
    <% for i=0 to ooffjungsan.FResultCount-1 %>
    <%
        if (ooffjungsan.FItemList(i).Ftot_orgsellprice<>0) then
            orgsellmargin = CLng((ooffjungsan.FItemList(i).Ftot_orgsellprice-ooffjungsan.FItemList(i).Ftot_jungsanprice)/ooffjungsan.FItemList(i).Ftot_orgsellprice*100*100)/100
        else
            orgsellmargin = 0
        end if

        if (ooffjungsan.FItemList(i).Ftot_realsellprice<>0) then
            realsellmargin = CLng((ooffjungsan.FItemList(i).Ftot_realsellprice-ooffjungsan.FItemList(i).Ftot_jungsanprice)/ooffjungsan.FItemList(i).Ftot_realsellprice*100*100)/100
        else
            realsellmargin = 0
        end if
    %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td ><a href="javascript:PopDetail('<%= ooffjungsan.FItemList(i).Fidx %>');"><%= ooffjungsan.FItemList(i).FYYYYMM %></a></td>
      	<td ><%= ooffjungsan.FItemList(i).getJGubunName %></td>
      	<td ><%= ooffjungsan.FItemList(i).Fjacc_nm %></td>
      	<td ><%= ooffjungsan.FItemList(i).Fdifferencekey %></td>
      	<td ><font color="<%= ooffjungsan.FItemList(i).GetTaxtypeNameColor %>"><%= ooffjungsan.FItemList(i).GetSimpleTaxtypeName %></font></td>
      	<td ><%= ooffjungsan.FItemList(i).GetItemVatTypeName %></td>
      	<td align="left"><a href="javascript:PopUpcheBrandInfoEdit('<%= ooffjungsan.FItemList(i).Fmakerid %>');"><%= ooffjungsan.FItemList(i).Fmakerid %></a></td>
      	<td align="center"><%= ooffjungsan.FItemList(i).FGroupid %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FTW_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FUW_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FOM_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FSM_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FCM_price,0) %></td>
        <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FET_price,0) %></td>

        <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_orgsellprice,0) %></td>
        <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_realsellprice,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftotalcommission,0) %></td>
      	<td align="center">
      	    <%= orgsellmargin %> %
      	</td>
      	<td align="center">
      	    <% if orgsellmargin<>realsellmargin then %>
      	    <font color="blue"><%= realsellmargin %></font> %
      	    <% else %>
      	    <%= realsellmargin %> %
      	    <% end if %>
      	</td>
      	<td ><a href="javascript:PopStateChange('<%= ooffjungsan.FItemList(i).Fidx %>');"><font color="<%= ooffjungsan.FItemList(i).GetStateColor %>"><%= ooffjungsan.FItemList(i).GetStateName %></font></a></td>
      	<td ><acronym title="<%= ooffjungsan.FItemList(i).Ftaxinputdate %>"><%= ooffjungsan.FItemList(i).Ftaxregdate %></acronym></td>
      	<td ><%= ooffjungsan.FItemList(i).Fipkumdate %></td>
        <td ><%= ooffjungsan.FItemList(i).Fjungsan_gubun %></td>
      	<td >
      	<% if ooffjungsan.FItemList(i).IsEditenable then %>
      	    <a href="javascript:DelMaster('<%= ooffjungsan.FItemList(i).Fidx %>');"><img src="/images/icon_delete2.gif" border="0" width="20"></a>
      	<% else %>
      	    <% if Not IsNULL(ooffjungsan.FItemList(i).Fneotaxno) then %>
      	        <% if (ooffjungsan.FItemList(i).Fbillsitecode="B") then %>
      	        <img src="/images/icon_print02.gif" width="14" height="14" border=0 onclick="window.open('http://www.bill36524.com/popupBillTax.jsp?NO_TAX=<%= ooffjungsan.FItemList(i).Fneotaxno %>&NO_BIZ_NO=2118700620')" style="cursor:hand">
      	        <% else %>
      	        <%= ooffjungsan.FItemList(i).Fbillsitecode %>
      	        <% end if %>
      	    <% end if %>
      	<% end if %>
      	<a href="/admin/upchejungsan/monthjungsanAdm.asp?makerid=<%= ooffjungsan.FItemList(i).Fmakerid %>&yyyy1=<%= LEFT(ooffjungsan.FItemList(i).Fyyyymm,4) %>&mm1=<%= right(ooffjungsan.FItemList(i).Fyyyymm,2) %>" target="_blank">POP</a>
     	</td>
    </tr>
    <% next %>
    <% else %>
    <tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">[ �˻������ �����ϴ�. ]</td>
	</tr>
	<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
			<% if ooffjungsan.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ooffjungsan.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ooffjungsan.StarScrollPage to ooffjungsan.FScrollCount + ooffjungsan.StarScrollPage - 1 %>
				<% if i>ooffjungsan.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ooffjungsan.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set ooffjungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
