<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������긮��Ʈ ���ݰ�꼭 �߱�
' History : 2022.07.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<%
dim i, research, page, ExcDel, productidx, yyyy1, mm1, yyyy2, mm2, dt, makerid, purchasetype, groupid, company_name, ppGubun
dim reportIdx, selectfinishflag, itemid
	page = requestCheckVar(getNumeric(request("page")),8)
	productidx = requestCheckVar(trim(getNumeric(request("productidx"))),8)
	reportIdx = requestCheckVar(trim(getNumeric(request("reportIdx"))),8)
	ExcDel = requestCheckVar(request("ExcDel"),1)
	research = requestCheckVar(request("research"),1)
	yyyy1    = requestCheckVar(request("yyyy1"),4)
	mm1      = requestCheckVar(request("mm1"),2)
	yyyy2    = requestCheckVar(request("yyyy2"),4)
	mm2      = requestCheckVar(request("mm2"),2)
	makerid = requestCheckVar(trim(request("makerid")),32)
	purchasetype = requestCheckVar(request("purchasetype"),2)
	groupid  = requestCheckVar(trim(request("groupid")),6)
	company_name  = requestCheckVar(trim(request("company_name")),64)
	ppGubun = requestCheckVar(trim(request("ppGubun")),32)
	selectfinishflag = requestCheckVar(request("selectfinishflag"),10)
	itemid      = requestCheckvar(request("itemid"),1500)

if page = "" then page = "1"
if ExcDel = "" and research="" then ExcDel = "Y"
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if
if yyyy2="" then
	dt = dateserial(year(Now),month(now),1)
	yyyy2 = Left(CStr(dt),4)
	mm2 = Mid(CStr(dt),6,2)
end if
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

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

dim oPurchasedJungsan
set oPurchasedJungsan = new CPurchasedJungsan
	oPurchasedJungsan.FCurrPage = page
	oPurchasedJungsan.Fpagesize = 100
    oPurchasedJungsan.FRectExcDel = ExcDel
	oPurchasedJungsan.FRectproductidx = productidx
	oPurchasedJungsan.FRectYYYYMM1 = yyyy1 + "-" + mm1
	oPurchasedJungsan.FRectYYYYMM2 = yyyy2 + "-" + mm2
	oPurchasedJungsan.FRectmakerid = makerid
	oPurchasedJungsan.FRectpurchasetype = purchasetype
	oPurchasedJungsan.FRectgroupid = groupid
	oPurchasedJungsan.FRectcompany_name = company_name
	oPurchasedJungsan.FRectppGubun = ppGubun
	oPurchasedJungsan.FRectreportIdx = reportIdx
	oPurchasedJungsan.FRectItemid       = itemid
	oPurchasedJungsan.FRectFinishFlag = selectfinishflag
	oPurchasedJungsan.GetPurchasedJungsanMasterList

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function SubmitFrm(pg) {
	document.frm.page.value=pg;
	document.frm.submit();
}

function finishFlagAllChgProcess(){
    var finishflagvar = '';

    for(var i=0; i<frm.finishflag.length; i++){
        if (frm.finishflag[i].checked){
            finishflagvar=frm.finishflag[i].value;
        }
    }
    if (finishflagvar==''){
        alert('���õ� ���ݰ�꼭 ���°��� �����ϴ�.');
        return;
    }

    frmArr.finishflag.value=finishflagvar;
	frmArr.mode.value='finishflagall';
	frmArr.action="/admin/newstorage/PurchasedProductJungsanProcess.asp";
	var ret = confirm('<%= yyyy1 %>-<%= mm1 %>��ü ������ ���°� ���� �Ͻðڽ��ϱ�?');
	if(ret){
		frmArr.submit();
	}
}

function finishflagarrChgProcess(){
    if ($('input[name="check"]:checked').length == 0) {
        alert('���� �������� �����ϴ�.');
        return;
    }

    var finishflagvar = '';

    for(var i=0; i<frm.finishflag.length; i++){
        if (frm.finishflag[i].checked){
            finishflagvar=frm.finishflag[i].value;
        }
    }
    if (finishflagvar==''){
        alert('���õ� ���ݰ�꼭 ���°��� �����ϴ�.');
        return;
    }

    frmArr.finishflag.value=finishflagvar;
	frmArr.mode.value='finishflagarr';
	frmArr.action="/admin/newstorage/PurchasedProductJungsanProcess.asp";
	var ret = confirm('���� ������ �ۼ��߿��� ��꼭�����û ���·� ���� �Ͻðڽ��ϱ�?');
	if(ret){
		frmArr.submit();
	}
}

function PopPurchasedTaxPrintReDirect(itax_no, groupcode){
	var popPurchasedwinsub = window.open("/admin/newstorage/red_Purchasedtaxprint.asp?tax_no=" + itax_no + "&groupcode="+groupcode ,"Purchasedtaxview","width=1200,height=768,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popPurchasedwinsub.focus();
}

function PopProductidxDetail(productidx){
	var popPurchasedProductidxDetail = window.open("/admin/newstorage/PurchasedProductModify.asp?idx=" + productidx ,"PurchasedProductidxDetail","width=1400,height=768,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popPurchasedProductidxDetail.focus();
}

function PopSheetidxDetail(sheetidx){
	var popPurchasedSheetidxDetail = window.open("/admin/newstorage/PurchasedProductSheetModify.asp?idx=" + sheetidx ,"PurchasedSheetidxDetail","width=1400,height=768,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popPurchasedSheetidxDetail.focus();
}

function toggleChecked(status) {
    $('[name="check"]').each(function () {
        $(this).prop("checked", status);
    });
}

function downloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/newstorage/PurchasedProductJungsanList_excel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}

$(document).ready(function () {
    var checkAllBox = $("#ckall");

    checkAllBox.click(function () {
        var status = checkAllBox.prop('checked');
        toggleChecked(status);
    });
});

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* idx : <input type="text" class="text" name="productidx" value="<%= productidx %>" size="8" maxlength=10>
		&nbsp;
		* �귣��ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
		&nbsp;
		* ��ü(�׷��ڵ�) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="6" maxlength=6>
		&nbsp;
		* ����ڸ� : <input type="text" class="text" name="company_name" value="<%= company_name %>" size="30" maxlength=64>
		<Br><Br>
		* ǰ�ǹ�ȣ : <input type="text" class="text" name="reportIdx" value="<%= reportIdx %>" size="8" maxlength=10>
		&nbsp;
		* ��ǰ�ڵ� : <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<a href="#" onClick="SubmitFrm(1); return false;"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" >
		* �������� : <% DrawYMBoxdynamic "yyyy1",yyyy1,"mm1",mm1,"" %>~<% DrawYMBoxdynamic "yyyy2",yyyy2,"mm2",mm2,"" %>
		&nbsp;
		* �������� : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",purchasetype,"" %>
		&nbsp;
		* ��뱸�� : 
		<% drawCSCommCodeBox true,"Z501","ppGubun",ppGubun,"" %>
		&nbsp;
		* ���� : 
		<select name="selectfinishflag" >
			<option value="">��ü
			<option value="0" <%= CHKIIF(selectfinishflag="0","selected","") %> >�ۼ���
			<option value="1" <%= CHKIIF(selectfinishflag="1","selected","") %> >��꼭�����û
			<option value="3" <%= CHKIIF(selectfinishflag="3","selected","") %> >����Ϸ�
			<!--<option value="0" <%'= CHKIIF(selectfinishflag="0","selected","") %> >������-->
			<!--<option value="1" <%'= CHKIIF(selectfinishflag="1","selected","") %> >��üȮ�δ��-->
			<!--<option value="2" <%'= CHKIIF(selectfinishflag="2","selected","") %> >��üȮ�οϷ�-->
			<!--<option value="3" <%'= CHKIIF(selectfinishflag="3","selected","") %> >����Ȯ��-->
			<!--<option value="7" <%'= CHKIIF(selectfinishflag="7","selected","") %> >�ԱݿϷ�-->
		</select>
	</td>
</tr>
<tr>
    <td bgcolor="#FFFFFF" >
        <label><input type="checkbox" name="ExcDel" value="Y" <%=chkIIF(ExcDel="Y","checked","")%> /> ������ ����</label>
	</td>
</tr>
</table>
<!-- �˻� �� -->

<br />

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
        <input type="radio" name="finishflag" value="0" checked >�ۼ���
        <input type="radio" name="finishflag" value="1" >��꼭�����û
        <input type="radio" name="finishflag" value="3" >����Ϸ�
        <input type="button" value="���û��º���" onClick="finishflagarrChgProcess();" class="button" >

		<% if yyyy1 + "-" + mm1=yyyy2 + "-" + mm2 then %>
			<input type="button" value="<%= yyyy1 %>-<%= mm1 %>��ü���º���" onClick="finishFlagAllChgProcess();" class="button" >
		<% end if %>
	</td>
	<td align="right">
		<input type="button" onclick="downloadexcel();" value="�����ٿ�ε�" class="button">
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<form action="" name="frmArr" method="post" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="finishflag" value="">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oPurchasedJungsan.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oPurchasedJungsan.FTotalPage %></b>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td><input type="checkbox" name="ckall" id="ckall"></td>
	<td width=60>IDX</td>
	<td>������idx</td>
    <td>�����</td>
    <td>�׷��ڵ�</td>
	<td width=100>�귣��ID</td>
    <td>����ڸ�</td>
    <td>��뱸��</td>
    <td>�����Ѿ�</td>
    <td>����ǰ��IDX</td>
    <td>���ݰ�꼭����</td>
    <td>���ݰ�꼭�����</td>
	<td>������</td>
    <td>���</td>
</tr>

<% if oPurchasedJungsan.FResultcount > 0 then %>
<% for i=0 to oPurchasedJungsan.FResultcount-1 %>
<tr bgcolor="<%= CHKIIF(IsNull(oPurchasedJungsan.FItemList(i).Fdeldt), "#FFFFFF", "#EEEEEE") %>" align="center" height="25">
	<td><input type="checkbox" name="check" value="<%= oPurchasedJungsan.FItemList(i).fsheetidx %>" onClick="AnCheckClick(this);"></td>
    <td>
		<a href="#" onclick="PopProductidxDetail('<%= oPurchasedJungsan.FItemList(i).fproductidx %>'); return false;" class="btn3 btnIntb">
		<%= oPurchasedJungsan.FItemList(i).fproductidx %></a>
	</td>
	<td>
		<a href="#" onclick="PopSheetidxDetail('<%= oPurchasedJungsan.FItemList(i).fsheetidx %>'); return false;" class="btn3 btnIntb">
		<%= oPurchasedJungsan.FItemList(i).fsheetidx %></a>
	</td>
	<td><%= oPurchasedJungsan.FItemList(i).fyyyymm %></td>
	<td><%= oPurchasedJungsan.FItemList(i).fgroupCode %></td>
	<td><%= oPurchasedJungsan.FItemList(i).fmakerid %></td>
	<td><%= oPurchasedJungsan.FItemList(i).fcompany_name %></td>
	<td><%= oPurchasedJungsan.FItemList(i).fppGubunname %></td>
	<td>
		<%= FormatNumber(oPurchasedJungsan.FItemList(i).fbuyPrice, 0) %>
	</td>
	<td><%= oPurchasedJungsan.FItemList(i).freportIdx %></td>
    <td><%= GetStateName(oPurchasedJungsan.FItemList(i).ffinishflag) %></td>
	<td><%= oPurchasedJungsan.FItemList(i).ftaxinputdate %></td>
	<td><%= oPurchasedJungsan.FItemList(i).Ftaxregdate %></td>
	<td>
		<% if IsElecTaxExists(oPurchasedJungsan.FItemList(i).fTaxLinkidx,oPurchasedJungsan.FItemList(i).ffinishflag) then %>
			<a href="#" onclick="PopPurchasedTaxPrintReDirect('<%= oPurchasedJungsan.FItemList(i).Fneotaxno %>','<%= oPurchasedJungsan.FItemList(i).fgroupCode %>'); return false;" class="btn3 btnIntb">���</a>
		<% else %>
			<%= oPurchasedJungsan.FItemList(i).Fbillsitecode %>
		<% end if %>
	</td>
</tr>
<% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
        	<% if oPurchasedJungsan.HasPreScroll then %>
				<a href="javascript:SubmitFrm('<%= oPurchasedJungsan.StartScrollPage-1 %>');">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oPurchasedJungsan.StartScrollPage to oPurchasedJungsan.FScrollCount + oPurchasedJungsan.StartScrollPage - 1 %>
				<% if i>oPurchasedJungsan.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:SubmitFrm('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oPurchasedJungsan.HasNextScroll then %>
				<a href="javascript:SubmitFrm('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set oPurchasedJungsan = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
