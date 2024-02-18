<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epItemManageCls.asp"-->
<%
Dim oFixedItem,page, i
page		= requestCheckvar(request("page"),10)

Dim listtp : listtp = requestCheckvar(request("listtp"),10)
Dim research : research = requestCheckvar(request("research"),10)
Dim showimage : showimage = requestCheckvar(request("showimage"),10)
Dim makerid : makerid = requestCheckvar(request("makerid"),32)
Dim sellyn : sellyn = requestCheckvar(request("sellyn"),10)
Dim itemidarr : itemidarr = requestCheckvar(request("itemidarr"),2000) 
Dim mwdiv : mwdiv = requestCheckvar(request("mwdiv"),10) 
Dim useyn : useyn = requestCheckvar(request("useyn"),10) 
Dim itemcouponyn : itemcouponyn = requestCheckvar(request("itemcouponyn"),10) 
Dim prcCkktype : prcCkktype = requestCheckvar(request("prcCkktype"),10) 
Dim epexcept : epexcept = requestCheckvar(request("epexcept"),10) 

If page = "" Then page = 1
if (research="") and (showimage="") then showimage="on"
if (research="") and (listtp="") then listtp="F"

itemidarr = replace(itemidarr,"'","")
itemidarr = replace(itemidarr,vbCRLF,",")
itemidarr = replace(itemidarr,vbCR,",")
itemidarr = replace(itemidarr,vbLf,",")

SET oFixedItem = new epShopFixedPrice
	oFixedItem.FCurrPage		= page
	oFixedItem.FPageSize		= 50
	oFixedItem.FRectItemIdArr	= itemidarr
	oFixedItem.FRectSellyn	    = sellyn
	oFixedItem.FRectMakerid     = makerid
    oFixedItem.FRectMwDiv       = mwdiv
    oFixedItem.FRectUseYn       = useyn
	oFixedItem.FRectPriceCheckType = prcCkktype
	oFixedItem.FRectEpExceptBrandItem = epexcept

    oFixedItem.FRectItemCouponYN = itemcouponyn

	if (listtp="X") then
		oFixedItem.getNVFixedPriceByNvMapXLLIST
	else
		oFixedItem.getNVFixedPriceLIST
	end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function rePage(sellsite,ErrType1,ErrType2){
	var frm = document.frm;
	frm.sellsite.value=sellsite;
	$("#ErrType1_"+ErrType1).val(ErrType1).prop("checked", true);
	$("#ErrType2_"+ErrType2).val(ErrType2).prop("checked", true);

	frm.submit();
}

function ShowAddNvFixedPrice(){
	if (document.getElementById("addNvFixedPriceTr").style.display==""){
		document.getElementById("addNvFixedPriceTr").style.display="none";	
		document.getElementById("additemid").value="";
		document.getElementById("fixedcash").value="";
	}else{
		document.getElementById("addNvFixedPriceTr").style.display="";	
	}
}

function AddNvFixedPriceAct(){
	var additemid = document.getElementById("additemid").value;
	var fixedcash = document.getElementById("fixedcash").value;

	if (additemid.length<1){
		alert("����� ��ǰ�ڵ带 �Է��ϼ���.");
		document.getElementById("additemid").focus();
		return;
	}

	if (!IsDigit(additemid)){
		alert("��ǰ�ڵ�� ���ڸ� �����մϴ�.");
		document.getElementById("additemid").focus();
		return;
	}

	if (fixedcash.length<1){
		alert("������ ������ �Է��ϼ���.");
		document.getElementById("fixedcash").focus();
		return;
	}

	if (!IsDigit(fixedcash)){
		alert("���������� ���ڸ� �����մϴ�.");
		document.getElementById("fixedcash").focus();
		return;
	}

	var popwin = window.open("epFixdPriceProcess.asp?mode=add&itemid="+additemid+"&fixedcash="+fixedcash,"AddNvFixedPriceAct","width=300, height=300,scrollbars=yes, resizabled=yes");
	popwin.focus();

}

function EditNvFixedPriceUseYn(iitemid,iuseyn){
	var confirmstr = "��ǰ�ڵ�: "+iitemid+"�� �������θ� "+iuseyn+" �� ���� �Ͻðڽ��ϱ�?"
	if (confirm(confirmstr)){
		var popwin = window.open("epFixdPriceProcess.asp?mode=useyn&itemid="+iitemid+"&useyn="+iuseyn,"AddNvFixedPriceAct","width=300, height=300,scrollbars=yes, resizabled=yes");
		popwin.focus();
	}
}

function popRegFileNvMapItem(){
	var popwin = window.open("<%=stsAdmURL%>/admin/etc/naverEp/popRegFileNvMapItem.asp","popRegFileNvMapItem","width=1200, height=1000,scrollbars=yes, resizabled=yes");
	popwin.focus();
}
	

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">

<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
        ��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemidarr" id="itemidarr"><%=replace(itemidarr,",",chr(10))%></textarea>
        &nbsp;&nbsp;
		�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
        &nbsp;&nbsp;
		|
		&nbsp;&nbsp;
		����Ʈ Ÿ�� :
		<input type="radio" name="listtp" value="F" <%=CHKIIF(listtp="F","checked","")%> >
		<%=CHKIIF(listtp="F","<strong>EP���ݰ���LIST ����</strong>","EP���ݰ���LIST ����")%>
		<input type="radio" name="listtp" value="X" <%=CHKIIF(listtp="X","checked","")%> >
		<%=CHKIIF(listtp="X","<strong>EP������XL ����</strong>","EP������XL ����")%>
		
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
	
	�ǸŻ��� : 
        <select name="sellyn" class="select">
            <option value="" <%= CHkIIF(sellyn="","selected","") %> >��ü
            <option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >�Ǹ�
            <option value="N" <%= CHkIIF(sellyn="N","selected","") %> >ǰ��
        </select>&nbsp;
     &nbsp;&nbsp;
    ���Ա��� : 
    <% Call drawSelectBoxMWU("mwdiv",mwdiv) %>
     &nbsp;&nbsp;
    ������뿩�� : 
        <select name="useyn" class="select">
            <option value="" <%= CHkIIF(useyn="","selected","") %> >��ü
            <option value="Y" <%= CHkIIF(useyn="Y","selected","") %> >���
            <option value="N" <%= CHkIIF(useyn="N","selected","") %> >������
        </select>&nbsp;
     &nbsp;&nbsp;
    (����)��ǰ�������� : 
        <select name="itemcouponyn" class="select">
            <option value="" <%= CHkIIF(itemcouponyn="","selected","") %> >��ü
            <option value="Y" <%= CHkIIF(itemcouponyn="Y","selected","") %> >��ǰ��������
            <option value="N" <%= CHkIIF(itemcouponyn="N","selected","") %> >����
            <option value="V" <%= CHkIIF(itemcouponyn="V","selected","") %> >NV��������
            <option value="C" <%= CHkIIF(itemcouponyn="C","selected","") %> >�Ϲ�����
        </select>&nbsp;
    &nbsp;&nbsp;
    ���ݰ��� : 
		<select name="prcCkktype" class="select">
			<option value="" <%= CHkIIF(prcCkktype="","selected","") %> >��ü
            <option value="9" <%= CHkIIF(prcCkktype="9","selected","") %> >�����ǸŰ� <> EP��������
            <option value="2" <%= CHkIIF(prcCkktype="2","selected","") %> >�����ǸŰ� > EP��������
            <option value="3" <%= CHkIIF(prcCkktype="3","selected","") %> >�����ǸŰ� < EP��������
            <option value="1" <%= CHkIIF(prcCkktype="1","selected","") %> >�����ǸŰ� = EP��������

			<option value="99" <%= CHkIIF(prcCkktype="99","selected","") %> >�����ǸŰ� <> EP������
			<option value="22" <%= CHkIIF(prcCkktype="22","selected","") %> >�����ǸŰ� > EP������
			<option value="33" <%= CHkIIF(prcCkktype="33","selected","") %> >�����ǸŰ� < EP������
			<option value="11" <%= CHkIIF(prcCkktype="11","selected","") %> >�����ǸŰ� = EP������
        </select>&nbsp;
    &nbsp;&nbsp;
	EP���� : 
		<select name="epexcept" class="select">
			<option value="" <%= CHkIIF(epexcept="","selected","") %> >��ü
			<option value="N" <%= CHkIIF(epexcept="N","selected","") %> >EP���ܾ���
			<option value="Y" <%= CHkIIF(epexcept="Y","selected","") %> >EP���ܺ귣��/��ǰ
		 </select>&nbsp;	
	&nbsp;&nbsp;
    <input type="checkbox" name="showimage" <%=CHKIIF(showimage="on","checked","")%> >�̹���ǥ��
	</td>
</tr>
</form>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= FormatNumber(oFixedItem.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oFixedItem.FTotalPage,0) %></b>
	</td>
	<!--
	<td colspan="2" align="right">
		<input type="button" id="addNvFixedPriceBtn" value="�űԵ��" onClick="ShowAddNvFixedPrice()">
	</td>
	-->
</tr>
<tr bgcolor="FFFFFF" id="addNvFixedPriceTr" style="display:">
	<td colspan="2" align="left">
		<input type="button" id="popRegFileNvMapItemBtn" value="EP������EXCEL���" onClick="popRegFileNvMapItem()">
	</td>
	<td colspan="17" align="right" >
		��ǰ�ڵ� : <input type="text" name="additemid" id="additemid" value="">
		&nbsp;&nbsp;
		�������� : <input type="text" name="fixedcash" id="fixedcash" value="">
		&nbsp;&nbsp;<input type="button" value="���" onClick="AddNvFixedPriceAct()">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">��ǰ��ȣ</td>
	<% if (showimage="on") then %>
	<td width="50">�̹���</td>
	<% end if %>
    <td width="100">�귣��ID</td>
    <td width="100">��ǰ��</td>
	<td width="60">�Ǹſ���</td>
	<td width="60">��������</td>
    <td width="50">���Ա���</td>
	<td width="80">�����ǸŰ�</td>
	<td width="70">����/����</td>
	<td width="80">EP��������</td>
    <td width="80">EP���������</td>
    <td width="80">EP����������</td>
    <td width="80">EP������뿩��</td>
	<td width="80">EP���������</td>

	<td width="80">EP���ܺ귣��/<br>��ǰ</td>

    <td width="100">EP��Ī�ڵ�</td>
    <td width="80">EP������</td>
    <td width="80">EP���TEN��ǰ��</td>
    <td width="80">EP��Ͼ�����Ʈ��</td>

</tr>
<%
	Dim DiffStat
%>
<% For i=0 to oFixedItem.FResultCount - 1 %>
<%
	DiffStat = ""
%>
<tr align="center" bgcolor="<%=CHKIIF(oFixedItem.FItemList(i).Fuseyn="N","#DDDDDD","#FFFFFF")%>">
	<td><a href="<%=wwwURL%>/<%=oFixedItem.FItemList(i).FItemID%>" target="_blank"><%= oFixedItem.FItemList(i).FItemID %></a></td>
	<% if (showimage="on") then %>
	<td><img src="<%= oFixedItem.FItemList(i).FImageSmall%>" width="50"></td>
	<% end if %>
    <td><%= oFixedItem.FItemList(i).FMakerid %></td>
    <td><%= oFixedItem.FItemList(i).FItemName %></td>
	<td><%= oFixedItem.FItemList(i).FSellyn %></td>
    <td>
		<%= oFixedItem.FItemList(i).getItemLimitStatHtml %>
	</td>
	<td><%= oFixedItem.FItemList(i).FMWdiv %></td>
	<td ><%= oFixedItem.FItemList(i).getSellcashHtml %></td>
	<td ><%= oFixedItem.FItemList(i).getDiscountTypeHtml %></td>
    <td >
	<% if NOT isNULL(oFixedItem.FItemList(i).Ffixedcash) then %>
	<%= Formatnumber(oFixedItem.FItemList(i).Ffixedcash, 0) %>
	<% end if %>
	</td>
    <td><%= oFixedItem.FItemList(i).Fregdt %></td>
    <td><%= oFixedItem.FItemList(i).Fupddt %></td>
    <td>
		<% if (oFixedItem.FItemList(i).Fuseyn="Y") then %>
		<a href="javascript:EditNvFixedPriceUseYn('<%=oFixedItem.FItemList(i).FItemID%>','N')"><%= oFixedItem.FItemList(i).Fuseyn %></a>
		<% else %>
		<a href="javascript:EditNvFixedPriceUseYn('<%=oFixedItem.FItemList(i).FItemID%>','Y')"><%= oFixedItem.FItemList(i).Fuseyn %></a>
		<% end if %>
	</td>
	<td><%= oFixedItem.FItemList(i).Freguserid %></td>
	<td><%= oFixedItem.FItemList(i).getEpExceptStr %></td>
	<td><%= oFixedItem.FItemList(i).FmatchNVMid %></td>
	<td>
        <% if NOT isNULL(oFixedItem.FItemList(i).Fnvminprice) then %>
        <%= Formatnumber(oFixedItem.FItemList(i).Fnvminprice, 0) %>
        <% end if %>
    </td>
	<td>
        <% if NOT isNULL(oFixedItem.FItemList(i).Fnvtensellcash) then %>
        <%= Formatnumber(oFixedItem.FItemList(i).Fnvtensellcash, 0) %>
        <% end if %>
    </td>
	<td><%= oFixedItem.FItemList(i).FNvMaplastupDt %></td>
	
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oFixedItem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oFixedItem.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oFixedItem.StartScrollPage to oFixedItem.FScrollCount + oFixedItem.StartScrollPage - 1 %>
    		<% if i>oFixedItem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oFixedItem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<%
SET oFixedItem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->