<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ��ǰ ���� ����
' History : 2007.04.08 ������ ����
'			2022.02.17 �ѿ�� ����(�˻������߰�. ������ �űԹ������� ������)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/newitemcouponcls.asp" -->

<%
dim oitemcoupon, page, research, iSerachType, sSearchTxt, selDate, sSdate, sEdate, onlyvalid, couponGubun, itemcoupontype
dim cpnvalue, i
	research    = requestCheckVar(request("research"),9)
	page        = requestCheckVar(request("page"),9)
	iSerachType = requestCheckVar(request("iSerachType"),9)
	sSearchTxt  = requestCheckVar(request("sSearchTxt"),32)
	onlyvalid   = requestCheckVar(request("onlyvalid"),9)
	selDate     = requestCheckVar(request("selDate"),10)
	sSdate      = requestCheckVar(request("sSdate"),10)
	sEdate      = requestCheckVar(request("sEdate"),10)
	couponGubun = requestCheckVar(request("couponGubun"),10)
	itemcoupontype = requestCheckVar(request("itemcoupontype"),10)
	cpnvalue	= requestCheckVar(request("cpnvalue"),10)

cpnvalue = replace(cpnvalue,"��","")
cpnvalue = replace(cpnvalue,"%","")
cpnvalue = Trim(replace(cpnvalue,",",""))

if Not(IsNumeric(cpnvalue)) then cpnvalue=""
''if (itemcoupontype="") then cpnvalue=""

if page="" then page=1
if research="" then onlyvalid="on"
if research="" and couponGubun="" then couponGubun="C"
    
set oitemcoupon = new CItemCouponMaster
	oitemcoupon.FPageSize=30
	oitemcoupon.FCurrPage = page
	oitemcoupon.FRectOnlyValid = onlyvalid
	oitemcoupon.FRectSearchType = iSerachType
	oitemcoupon.FRectSearchTxt = sSearchTxt
	oitemcoupon.FRectSearchDate = selDate
	oitemcoupon.FRectStartDate = sSdate
	oitemcoupon.FRectEndDate   = sEdate
	oitemcoupon.FRectCouponGubun = couponGubun
	oitemcoupon.FRectitemcoupontype = itemcoupontype
	oitemcoupon.FRectItemCouponValue = cpnvalue
	oitemcoupon.GetItemCouponMasterList

%>
<script type='text/javascript'>

function NextPage(page){
    var frm = document.frmSearch;
    frm.page.value = page;
    frm.submit();
}

function RegItemCoupon(){
	var popwin = window.open('itemcouponmasterreg.asp','RegItemCoupon','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditItemCoupon(itemcouponidx){
	var popwin = window.open('itemcouponmasterreg.asp?itemcouponidx=' + itemcouponidx,'EditItemCoupon','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditItemCouponItemMulti(){
    var popwin = window.open('itemcouponitemlisteidtMulti.asp'  ,'EditItemCouponItemMulti','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function NvItemCouponExcept(){
	var popwin = window.open('/admin/etc/naverEp/exceptNvCpn.asp'  ,'NvItemCouponExcept','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function EditCouponItemList(itemcouponidx){
	var popwin = window.open('itemcouponitemlisteidt.asp?itemcouponidx=' + itemcouponidx,'EditCouponItemList','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=350, height=350');
	winCal.focus();
}

function trim(value) {
	return value.replace(/^\s+|\s+$/g,"");
}

function isUInt(val) {
	var re = /^[0-9]+$/;
	return re.test(val);
}

function SubmitFrm(frm) {
	if ((frm.iSerachType.value == "1") || (frm.iSerachType.value == "2")) {
		if (frm.sSearchTxt.value*0 != 0) {
			alert('�����ڵ�/�̺�Ʈ�ڵ�� ���ڸ� �����մϴ�.');
			return;
		}
	}

	frm.sSearchTxt.value = trim(frm.sSearchTxt.value);

	if (frm.iSerachType.value == "1") {
		if (isUInt(frm.sSearchTxt.value) != true) {
			//?????
			//alert("�����ڵ�� ���ڸ� �����մϴ�.");
			//return;
		}
	}

	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frmSearch" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* <select name="iSerachType">
		<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>�����ڵ�</option>
		<option value="3" <%IF Cstr(iSerachType) = "3" THEN%>selected<%END IF%>>������</option>
		<option value="4" <%IF Cstr(iSerachType) = "4" THEN%>selected<%END IF%>>��������</option>
		<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
		</select>
		<input type="text" name="sSearchTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">
		&nbsp;
		* <select name="selDate">
		<option value="S" <%if Cstr(selDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
		<option value="E" <%if Cstr(selDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
		<option value="R" <%if Cstr(selDate) = "R" THEN %>selected<%END IF%>>������ ����</option>
		</select>
		<input type="text" size="10" name="sSdate" value="<%=sSdate%>" onClick="jsPopCal('sSdate');" style="cursor:hand;">
		~ <input type="text" size="10" name="sEdate" value="<%=sEdate%>" onClick="jsPopCal('sEdate');"  style="cursor:hand;">
		&nbsp;
		<input type="checkbox" name="onlyvalid" <% if onlyvalid="on" then response.write "checked" %> >������������ �� ����
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitFrm(document.frmSearch);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �������� 
		<select name="couponGubun">
			<option value="" <%=CHKIIF(couponGubun="","selected","") %> >��ü
			<option value="C" <%=CHKIIF(couponGubun="C","selected","") %> >�Ϲ�
			<option value="V" <%=CHKIIF(couponGubun="V","selected","") %> >���̹���������
			<option value="P" <%=CHKIIF(couponGubun="P","selected","") %> >�����ι߱�
			<option value="T" <%=CHKIIF(couponGubun="T","selected","") %> >Ÿ��(E-mailƯ��)
		</select>
		&nbsp;
		* ���α��� 
		<select name="itemcoupontype">
			<option value="" <%=CHKIIF(itemcoupontype="","selected","") %> >��ü
			<option value="1" <%=CHKIIF(itemcoupontype="1","selected","") %> >%
			<option value="2" <%=CHKIIF(itemcoupontype="2","selected","") %> >�ݾ�
			<option value="3" <%=CHKIIF(itemcoupontype="3","selected","") %> >��۷�
		</select>
		&nbsp;
		* ���ΰ�(%, ��)
		<input type="text" name="cpnvalue" value="<%=cpnvalue%>" size="7" maxlength="10" style="text-align:right">
	</td>
</tr>
</table>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="�ű� ��ǰ �������" onclick="RegItemCoupon();">
	</td>
	<td align="right">
		<input type="button" class="button" value="���̹������������ܰ���" onclick="NvItemCouponExcept();">
		&nbsp;
		<input type="button" class="button" value="��� ��ǰ����" onclick="EditItemCouponItemMulti();">
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="14" align="left">
		�˻��Ǽ� : <%= oitemcoupon.FTotalCount %> �� Page : <%= page %>/<%= oitemcoupon.FTotalPage %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="50">������ȣ</td>
	<td align="center" width="80">��������</td>
	<td align="center" width="70">�̺�Ʈ�ڵ�<br>(�׷��ڵ�)</td>
	<td >������</td>
	<td >��������</td>
	<td align="center" width="100">���αݾ�</td>
	<td align="center" width="60">����ǰ</td>
	<td align="center" width="100">������</td>
	<td align="center" width="100">������</td>
	<td align="center" width="70">����</td>
	<td align="center" width="120">�⺻<br>��������</td>
	<td align="center" width="80">�����</td>
	<td align="center" width="100">������</td>
	<td align="center" width="40">���</td>
</tr>
<% if oitemcoupon.FResultCount>0 then %>
<% for i=0 to oitemcoupon.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td><%= oitemcoupon.FItemList(i).Fitemcouponidx %></td>
	<td><font color="<%= oitemcoupon.FItemList(i).getCouponGubunColor %>"><%= oitemcoupon.FItemList(i).getCouponGubunName %></font></td>
	<td>
		<%= oitemcoupon.FItemList(i).Fevt_code %>
		<% if Not IsNULL(oitemcoupon.FItemList(i).Fevtgroup_code) then %>
		(<%= oitemcoupon.FItemList(i).Fevtgroup_code %>)
		<% end if %>
	</td>
	<td><a href="javascript:EditItemCoupon('<%= oitemcoupon.FItemList(i).Fitemcouponidx %>')"><%= replace(oitemcoupon.FItemList(i).Fitemcouponname,"4�� ���⼼��","<strong>4�� ���⼼��</strong>") %></a></td>
	<td><%= oitemcoupon.FItemList(i).Fitemcouponexplain %></td>
	<td><%= oitemcoupon.FItemList(i).GetDiscountStr %></td>
	<td><a href="javascript:EditCouponItemList('<%= oitemcoupon.FItemList(i).Fitemcouponidx %>');"><%= oitemcoupon.FItemList(i).Fapplyitemcount %> ��</a></td>
	<td><%= ChkIIF(Right(oitemcoupon.FItemList(i).Fitemcouponstartdate,8)="00:00:00",Left(oitemcoupon.FItemList(i).Fitemcouponstartdate,10),oitemcoupon.FItemList(i).Fitemcouponstartdate) %></td>
	<td><%= ChkIIF(Right(oitemcoupon.FItemList(i).Fitemcouponexpiredate,8)="23:59:59",Left(oitemcoupon.FItemList(i).Fitemcouponexpiredate,10),oitemcoupon.FItemList(i).Fitemcouponexpiredate) %></td>
	<td><font color="<%= oitemcoupon.FItemList(i).GetOpenStateColor %>"><%= oitemcoupon.FItemList(i).GetOpenStateName %></font></td>
	<td><%= oitemcoupon.FItemList(i).GetMargintypeName %></td>
	<td><%= Left(oitemcoupon.FItemList(i).FRegDate,10) %></td>
	<td><%=oitemcoupon.FItemList(i).FlastupDt%></td>
	<td>
		<% if oitemcoupon.FItemList(i).Fopenstate>="7" then %>
		<a href="<%=stsAdmURL%>/admin/dataanalysis/report/simpleQry.asp?menupos=4116&qryidx=221&itmCpnIdx=<%=oitemcoupon.FItemList(i).Fitemcouponidx%>"><img src="/images/documents_icon.png" /></a>
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="14" align="center">
	<% if oitemcoupon.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oitemcoupon.StarScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oitemcoupon.StarScrollPage to oitemcoupon.FScrollCount + oitemcoupon.StarScrollPage - 1 %>
		<% if i>oitemcoupon.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oitemcoupon.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set oitemcoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
