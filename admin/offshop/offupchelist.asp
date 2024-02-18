<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������
' History : 2009.04.07 ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<%
dim shopid, designer, comm_cd, shopusing, partnerusing, page, research, diffCk, offupbea, i
dim hasContOnly, maeipdiv, vPurchaseType, isoffusing, adminopen, diffshopdiv13
	page        = RequestCheckVar(request("page"),9)
	shopid      = RequestCheckVar(request("shopid"),32)
	designer    = RequestCheckVar(request("designer"),32)
	comm_cd     = RequestCheckVar(request("comm_cd"),9)
	shopusing   = RequestCheckVar(request("shopusing"),1)
	partnerusing  = RequestCheckVar(request("partnerusing"),1)
	research    = RequestCheckVar(request("research"),9)
	diffCk      = RequestCheckVar(request("diffCk"),9)
	offupbea    = RequestCheckVar(request("offupbea"),9)
	hasContOnly    = RequestCheckVar(request("hasContOnly"),9)
	maeipdiv = RequestCheckVar(request("maeipdiv"),1)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	isoffusing = requestCheckVar(request("isoffusing"),1)
	adminopen = requestCheckVar(request("adminopen"),1)
	diffshopdiv13      = RequestCheckVar(request("diffshopdiv13"),2)

if page="" then page=1
if (research="") and (hasContOnly="") then hasContOnly="ON"
if (research="") then shopusing="Y"
if diffshopdiv13="on" then
	hasContOnly="OFF"
	comm_cd=""
end if

dim ochargeuser

set ochargeuser = new COffShopChargeUser
	ochargeuser.FCurrPage = page

	if (designer<>"") then
		ochargeuser.FPageSize = 400
	else
		ochargeuser.FPageSize = 100
	end if

	ochargeuser.FRectShopID     = shopid
	ochargeuser.FRectDesigner   = designer
	ochargeuser.FRectComm_cd    = comm_cd
	ochargeuser.FRectShopusing  = shopusing
	ochargeuser.FRectpartnerusing  	= partnerusing
	ochargeuser.FRectOffUpBea   	= offupbea
	ochargeuser.FRectHasContOnly  = hasContOnly
	ochargeuser.FRectmaeipdiv = maeipdiv
	ochargeuser.FRectBrandPurchaseType = vPurchaseType
	ochargeuser.FRectisoffusing = isoffusing
	ochargeuser.FRectadminopen = adminopen

	if (diffCk<>"") then
		ochargeuser.GetOffShopbrandcontractlisterror
	elseif diffshopdiv13<>"" then
		ochargeuser.GetOffShopbrandcontractdiff
	else
	    if (shopid="") and (designer="") then
	        if (offupbea<>"") then
	            ochargeuser.GetOffShopbrandcontractlist
	        end if
	    else
	        if (shopid<>"") then
	    		ochargeuser.GetOffShopDesignerList1
	    	else
	    		ochargeuser.GetOffShopbrandcontractlist
	    	end if
	    end if
	end if
%>
<script type="text/javascript">

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo","width=1280 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popXL() {
	frm.action="/admin/offshop/offupchelist_xl_download.asp";
	frm.target='view';
	frm.submit();
	frm.action='';
	frm.target='';
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

</script>

<iframe id="view" name="view" src="" width="0" height="0" allowtransparency="true" frameborder="0" scrolling="no"></iframe>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" >
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		ShopID : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		&nbsp;
		�귣��ID : <% drawSelectBoxDesignerwithName "designer",designer  %>
     	&nbsp;
     	Shop����� : <% drawSelectBoxUsingYN "shopusing",shopusing %>
		&nbsp;
		��࿩�� :
     	<select name='hasContOnly'>
     		<option value=''>��ü</option>
     		<option value='ON' <% if hasContOnly="ON" then response.write "selected" %>>���Y</option>
     		<option value='OFF' <% if hasContOnly="OFF" then response.write "selected" %>>���N</option>
     	</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�������� : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
		&nbsp;
		OFF��౸�� : <% 'drawSelectBoxOFFChargeDiv "chargediv",chargediv %>
		<% drawSelectBoxOFFJungsanCommCD "comm_cd",comm_cd %>
		&nbsp;
		ON��౸�� :
		<% DrawBrandMWUCombo "maeipdiv",maeipdiv %>
		&nbsp;
		OFF�귣���뿩�� :
		<% drawSelectBoxUsingYN "isoffusing",isoffusing %>
		&nbsp;
		�������ξ��λ�뿩�� :
		<% drawSelectBoxUsingYN "adminopen",adminopen %>
     	&nbsp;
     	SCM���¿��� : <% drawSelectBoxUsingYN "partnerusing",partnerusing %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="diffCk" <%= ChkIIF(diffCk="on","checked","") %> >��ǥ���� ����ġ �˻�
     	&nbsp;
     	<input type="checkbox" name="offupbea" <%= ChkIIF(offupbea="on","checked","") %> >���� ��ü���
     	&nbsp;
     	<input type="checkbox" name="diffshopdiv13" <%= ChkIIF(diffshopdiv13="on","checked","") %> >������ġ(��������ǥY,��������ǥN)
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<% if (shopid="") and (designer="") and (offupbea="") then %>
		<div align="center"><font color="red">���� �Ǵ� �귣�带 �����ϼ���.</font></div>
		<% end if %>
    </td>
    <td align="right">
		<input type="button" class="button" value="�����ٿ�" onclick="popXL();">
    </td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= ochargeuser.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= ochargeuser.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan=2 width="100">ShopID</td>
	<td rowspan=2 width="100">Shop��</td>
	<td rowspan=2>�귣��ID</td>
	<td rowspan=2>�귣���</td>
	<td rowspan=2 width="70">��������</td>
	<td colspan=3>OFF ���</td>
	<td colspan=2>ON ���</td>
	<td rowspan=2 width="50">OFF<br>�귣��<br>��뿩��</td>
	<td rowspan=2 width="50">��������<br>����<br>���¿���</td>
	<td rowspan=2 width="50">SCM<br>���¿���</td>
	<td rowspan=2 width="50">������</td>
	<td rowspan=2 width="50">����</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="90">��౸��</td>
	<td width="50">��ü<br>���Ը���</td>
	<td width="50">SHOP<br>�����</td>
	<td width="50">��౸��</td>
	<td width="50">����</td>
</tr>
<% if ochargeuser.FresultCount >0 then %>
<% for i=0 to ochargeuser.FresultCount-1 %>

<% if ochargeuser.FItemList(i).FShopIsUsing="Y" then %>
	<tr align="center" bgcolor="#FFFFFF">
<% else %>
	<tr align="center" bgcolor="#DDDDDD">
<% end if %>

	<td>
		<%if (ochargeuser.FItemList(i).FShopid="streetshop000") or (ochargeuser.FItemList(i).FShopid="streetshop800") or (ochargeuser.FItemList(i).FShopid="streetshop870") then %>
			<strong><%= ochargeuser.FItemList(i).FShopID %></strong>
		<% else %>
			<%= ochargeuser.FItemList(i).FShopID %>
		<% end if %>
	</td>
	<td>
		<%= ochargeuser.FItemList(i).FShopName %>
	</td>
	<td><font color="<%= CHKIIF(ochargeuser.FItemList(i).IsContractExists,"#000000","#AAAAAA") %>"><%= ochargeuser.FItemList(i).FDesignerId %></font></td>
	<td><font color="<%= CHKIIF(ochargeuser.FItemList(i).IsContractExists,"#000000","#AAAAAA") %>"><%= ochargeuser.FItemList(i).FDesignerName %></font></td>
	<td>
		<%= getBrandPurchaseType(ochargeuser.FItemList(i).fpurchaseType) %>
	</td>

	<% if (ochargeuser.FItemList(i).IsContractExists) then %>
		<td>
			<font color="<%= ochargeuser.FItemList(i).getChargeDivColor %>">
				<%= ochargeuser.FItemList(i).getChargeDivName %>
				<% if (ochargeuser.FItemList(i).Fjungsan_gubun = "���̰���") then %>
				<br />(����)
				<% end if %>
			</font>
		</td>
		<td><%= ochargeuser.FItemList(i).FDefaultMargin %></td>
		<td><%= ochargeuser.FItemList(i).FDefaultSuplyMargin %></td>
	<% else %>
		<td></td>
		<td></td>
		<td></td>
	<% end if %>
	<td>
		<font color="<%= CHKIIF(ochargeuser.FItemList(i).IsContractExists,"#000000","#AAAAAA") %>">
		<%= ochargeuser.FItemList(i).getMwName %></font>
	</td>
	<td>
		<%= ochargeuser.FItemList(i).Fonlinedefaultmargine %>
	</td>
	<td align="center">
		<% if (ochargeuser.FItemList(i).fisoffusing="Y") then  %>
			O
		<% else %>
			X
		<% end if %>
	</td>
	<td align="center">
		<% if (ochargeuser.FItemList(i).FAdminopen="Y") then  %>
			O
		<% else %>
			X
		<% end if %>
	</td>
	<td align="center">
		<% if (ochargeuser.FItemList(i).FPartnerisusing="Y") then  %>
			O
		<% else %>
			X
		<% end if %>
	</td>
	<td><%= ochargeuser.FItemList(i).Fjungsan_date_off %></td>
	<td align="center"><input type="button" class="button" value="����" onclick="editOffDesinger('<%= ochargeuser.FItemList(i).FShopid %>','<%= ochargeuser.FItemList(i).FDesignerId %>');"></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="25" align="center">
	<% if ochargeuser.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ochargeuser.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ochargeuser.StarScrollPage to ochargeuser.FScrollCount + ochargeuser.StarScrollPage - 1 %>
		<% if i>ochargeuser.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ochargeuser.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="20">��ϵ� ������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<%
set ochargeuser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
