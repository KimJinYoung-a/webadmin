<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/deliverypolicycls.asp"-->

<%

dim designer, mduserid , catecode, defaultdeliveryType, isusingbrand, isusingitem, mwdiv
dim i

dim currpage

currpage 	= requestCheckvar(request("currpage"),32)
designer 	= requestCheckvar(request("designer"),32)
mduserid 	= requestCheckvar(request("mduserid"),32)
catecode 	= requestCheckvar(request("catecode"),3)
defaultdeliveryType 	= requestCheckvar(request("defaultdeliveryType"),4)
isusingbrand 	= requestCheckvar(request("isusingbrand"),1)
isusingitem 	= requestCheckvar(request("isusingitem"),1)
mwdiv 	= requestCheckvar(request("mwdiv"),1)

if (currpage = "") then
	currpage = 1
end if



'==============================================================================
dim ODeliveryPolicy

set ODeliveryPolicy = new CDeliveryPolicy

ODeliveryPolicy.FPageSize = 50
ODeliveryPolicy.FCurrPage = currpage
ODeliveryPolicy.FRectUserID = designer
ODeliveryPolicy.FRectMDUserID = mduserid
ODeliveryPolicy.FRectCategoryCode = catecode
ODeliveryPolicy.FRectDefaultDeliveryType = defaultdeliveryType
ODeliveryPolicy.FRectIsUsingBrand = isusingbrand
ODeliveryPolicy.FRectIsUsingItem = isusingitem
ODeliveryPolicy.FRectMWDiv = mwdiv



ODeliveryPolicy.GetList

%>
<script language='javascript'>
function popItemSellEdit(designerid,mwdiv,usingyn){
	var popwin = window.open('/admin/shopmaster/itemviewset.asp?menupos=24&makerid=' + designerid + '&mwdiv=' + mwdiv + '&usingyn=' + usingyn  ,'popItemSellEdit','width=1000,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}
</script>
<!-- �˻� ���� -->



<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
    		�귣�� : <% drawSelectBoxDesignerwithName "designer", designer %>
			&nbsp;
			����� : <% drawSelectBoxCoWorker "mduserid", mduserid %>
			&nbsp;
			ī�װ� : <% SelectBoxBrandCategory "catecode", catecode %>

		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�귣�� �����å : <% drawPartnerCommCodeBox True,"deliveryType","defaultdeliveryType", defaultdeliveryType,"" %>
	     	&nbsp;
			�귣�� ��뿩�� :
			<select class="select" name="isusingbrand">
		     	<option value='' >��ü</option>
		     	<option value='Y' <% if (isusingbrand = "Y") then %>selected<% end if %>>���</option>
		     	<option value='N' <% if (isusingbrand = "N") then %>selected<% end if %>>������</option>
	     	</select>
	     	&nbsp;
			��ǰ ��뿩�� :
			<select class="select" name="isusingitem">
		     	<option value='' selected>��ü</option>
		     	<option value='Y' <% if (isusingitem = "Y") then %>selected<% end if %>>���</option>
		     	<option value='N' <% if (isusingitem = "N") then %>selected<% end if %>>������</option>
	     	</select>
			<!-- ���� ������. 2015-04-08, skyer9
	     	&nbsp;
			�ŷ����� :
			<select class="select" name="mwdiv">
		     	<option value='' selected>��ü</option>
		     	<option value='M' <% if (mwdiv = "M") then %>selected<% end if %>>����</option>
				<option value='W' <% if (mwdiv = "W") then %>selected<% end if %>>��Ź</option>
				<option value='U' <% if (mwdiv = "U") then %>selected<% end if %>>��ü</option>
	     	</select>
			-->
            &nbsp;
            <input type="checkbox" name="exctpl" value="Y" checked disabled> 3PL ����

		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �����ޱ� -->
<%
	Dim exlPsz, exlPg
	exlPsz = 5000
	exlPg = ceil(ODeliveryPolicy.FTotalCount/exlPsz)
%>
<script>
	function fnGetExcel(pg) {
		window.open("brandlist_baesong_excel.asp?currpage="+pg+"&designer=<%=designer%>&mduserid=<%=mduserid%>&defaultdeliveryType=<%=defaultdeliveryType%>&catecode=<%=catecode%>&mwdiv=<%=mwdiv%>&isusingbrand=<%=isusingbrand%>&isusingitem=<%=isusingitem%>");
	}
</script>

<div style="text-align:right; margin:10px 5px;">
	<select id="exlPage" class="select" style="vertical-align: middle;">
	<% for i=1 to exlPg %>
	<option value="<%=i%>"><%=((i-1)*exlPsz)+1%>~<%=chkIIF(i*exlPsz<ODeliveryPolicy.FTotalCount,i*exlPsz,ODeliveryPolicy.FTotalCount)%></option>
	<% next %>
	</select>
	<img src="/images/btn_excel.gif" onClick="fnGetExcel(document.getElementById('exlPage').value)" style="cursor:pointer;vertical-align: middle;" />
</div>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">
			�˻���� : <b><%= ODeliveryPolicy.FTotalCount %></b>
			&nbsp;
			������ : <b><%= currpage %> / <%= ODeliveryPolicy.FTotalPage %></b>
		</td>
		<td colspan="16" align=right>
			*��ǰ�� ��� ��ü����� ��츸 ���������� �����մϴ�.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2">�귣��ID</td>
		<td rowspan="2">��Ʈ��Ʈ��</td>
      	<td rowspan="2">ȸ���</td>
      	<td rowspan="2" width="80">�귣��<br>�����å</td>

      	<td colspan="6">��ǰ����(��ǰ��)</td>
      	<td colspan="3">��ǰ����(�ŷ�����)</td>
      	<td rowspan="2" width="60">��ü��ǰ��</td>

      	<td colspan="2">������ۺ����(��)</td>
      	<td rowspan="2" width="80">���</td>

    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      	<td width="60">1�����̸�</td>
      	<td width="50">1������</td>
      	<td width="50">2������</td>
      	<td width="50">3������</td>
      	<td width="50">5������</td>
      	<td width="60">5�����̻�</td>

      	<td width="40">��ü</td>
      	<td width="40">��Ź</td>
      	<td width="40">����</td>

      	<td width="60">������<br>�ּұݾ�</td>
      	<td width="70">������ۺ�</td>
    </tr>
<% if ODeliveryPolicy.FresultCount < 1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for i = 0 to ODeliveryPolicy.FresultCount - 1 %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= ODeliveryPolicy.FItemList(i).Fuserid %></td>
    	<td><%= ODeliveryPolicy.FItemList(i).Fsocname_kor %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fconame %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).FdefaultdeliveryType %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice0 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice10000 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice20000 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice30000 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice40000 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice50000 %></td>
      	<td><a href="javascript:popItemSellEdit('<%= ODeliveryPolicy.FItemList(i).Fuserid %>','U','<%= isusingitem %>');"><%= ODeliveryPolicy.FItemList(i).Fupchecount %></a></td>
      	<td><a href="javascript:popItemSellEdit('<%= ODeliveryPolicy.FItemList(i).Fuserid %>','W','<%= isusingitem %>');"><% if (ODeliveryPolicy.FItemList(i).FdefaultdeliveryType <> "ETC") and ((ODeliveryPolicy.FItemList(i).Fwitakcount <> 0) or (ODeliveryPolicy.FItemList(i).Fmaeipcount <> 0)) then %><font color=red><b><% end if %><%= ODeliveryPolicy.FItemList(i).Fwitakcount %></a></td>
      	<td><a href="javascript:popItemSellEdit('<%= ODeliveryPolicy.FItemList(i).Fuserid %>','M','<%= isusingitem %>');"><% if (ODeliveryPolicy.FItemList(i).FdefaultdeliveryType <> "ETC") and ((ODeliveryPolicy.FItemList(i).Fwitakcount <> 0) or (ODeliveryPolicy.FItemList(i).Fmaeipcount <> 0)) then %><font color=red><b><% end if %><%= ODeliveryPolicy.FItemList(i).Fmaeipcount %></a></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fitemcount %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).FdefaultFreeBeasongLimit %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).FdefaultDeliverPay %></td>
      	<td>
		<!-- �����̻� + �ٹ����ٻ���� - MD��Ʈ �����̻� ��������(�������� ����:2011.09.01) -->
		<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or ((session("ssAdminLsn") <= "4") and (session("ssAdminPsn") = "11"))) then %>
			<% if (ODeliveryPolicy.FItemList(i).FdefaultdeliveryType = "ETC") then %>
				<% if (ODeliveryPolicy.FItemList(i).Fwitakcount = 0) and (ODeliveryPolicy.FItemList(i).Fmaeipcount = 0) then %>
			<input type="button" class="button" value="�űԼ���" onClick="PopBrandAdminUsingChange('<%= ODeliveryPolicy.FItemList(i).Fuserid %>')">
				<% end if %>
      		<% else %>
      		<input type="button" class="button" value="��������" onClick="PopBrandAdminUsingChange('<%= ODeliveryPolicy.FItemList(i).Fuserid %>')">
      		<% end if %>
      	<% end if %>
      	</td>
    </tr>
	<% next %>
<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
    		<% if ODeliveryPolicy.HasPreScroll then %>
    			<a href="?currpage=<%= ODeliveryPolicy.StartScrollPage-1 %>&menupos=<%= menupos %>&designer=<%= designer %>&mduserid=<%= mduserid %>&catecode=<%= catecode %>&defaultdeliveryType=<%= defaultdeliveryType %>&isusingbrand=<%= isusingbrand %>&isusingitem=<%= isusingitem %>">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i = (0 + ODeliveryPolicy.StartScrollPage) to (ODeliveryPolicy.FScrollCount + ODeliveryPolicy.StartScrollPage - 1) %>
    			<% if i>ODeliveryPolicy.FTotalpage then Exit for %>
    			<% if CStr(currpage)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="?currpage=<%= i %>&menupos=<%= menupos %>&designer=<%= designer %>&mduserid=<%= mduserid %>&catecode=<%= catecode %>&defaultdeliveryType=<%= defaultdeliveryType %>&isusingbrand=<%= isusingbrand %>&isusingitem=<%= isusingitem %>">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if ODeliveryPolicy.HasNextScroll then %>
    			<a href="?currpage=<%= i %>&menupos=<%= menupos %>&designer=<%= designer %>&mduserid=<%= mduserid %>&catecode=<%= catecode %>&defaultdeliveryType=<%= defaultdeliveryType %>&isusingbrand=<%= isusingbrand %>&isusingitem=<%= isusingitem %>">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>


















<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
