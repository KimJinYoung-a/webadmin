<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ֹ� Ŭ����
' Hieditor : 2009.04.17 �̻� ����
'			 2016.07.19 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<%
dim orderserial, searchtype, searchrect, yyyy1,yyyy2,mm1,mm2,dd1,dd2, page, ojumun, ix,iy
dim nowdate,searchnextdate,research, jumundiv, sellchnl, cknodate,ckdelsearch,ckipkumdiv4,ckipkumdiv2, not3pl, ipkumdiv
	searchtype  = requestCheckVar(request("searchtype"),32)
	searchrect  = requestCheckVar(request("searchrect"),32)
	yyyy1       = requestCheckVar(request("yyyy1"),4)
	mm1         = requestCheckVar(request("mm1"),2)
	dd1         = requestCheckVar(request("dd1"),2)
	yyyy2       = requestCheckVar(request("yyyy2"),4)
	mm2         = requestCheckVar(request("mm2"),2)
	dd2         = requestCheckVar(request("dd2"),2)
	jumundiv    = requestCheckVar(request("jumundiv"),10)
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	cknodate    = request("cknodate")
	ckdelsearch = request("ckdelsearch")
	ckipkumdiv4 = request("ckipkumdiv4")
	orderserial = request("orderserial")
	ckipkumdiv2 = request("ckipkumdiv2")
	ipkumdiv	= requestCheckVar(request("ipkumdiv"),1)
	research    = request("research")
	page = request("page")
	not3pl = request("not3pl")

if (page="") then page=1
nowdate = Left(CStr(now()),10)

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

if research="" then ckipkumdiv2="on"
if research="" then not3pl="on"
    
set ojumun = new CJumunMaster
	if (jumundiv="flowers") then
		ojumun.FRectIsFlower = "Y"
	elseif (jumundiv="minus") then
	    ojumun.FRectIsMinus = "Y"
	elseif (jumundiv="foreign") then
	    ojumun.FRectIsForeign = "Y"
	elseif (jumundiv="military") then
	    ojumun.FRectIsMilitary = "Y"
	elseif (jumundiv="pojang") then
	    ojumun.FRectPojangOrder = "Y"
    elseif (jumundiv="sendGift") then
        ojumun.FRectIsSendGift = "Y"
	end if
	
	if cknodate="" then
		ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
		ojumun.FRectRegEnd = searchnextdate
	end if
	
	if ckdelsearch<>"on" then
		ojumun.FRectDelNoSearch="on"
	end if

	if searchtype="01" then
		ojumun.FRectBuyname = searchrect
	elseif searchtype="02" then
		ojumun.FRectReqName = searchrect
	elseif searchtype="03" then
		ojumun.FRectUserID = searchrect
	elseif searchtype="04" then
		ojumun.FRectIpkumName = searchrect
	elseif searchtype="06" then
		ojumun.FRectSubTotalPrice = searchrect
	end if
	
	ojumun.FPageSize = 30
	ojumun.FRectIpkumDiv4 = ckipkumdiv4
	ojumun.FRectIpkumDiv2 = ckipkumdiv2
	ojumun.FRectIpkumDiv = ipkumdiv
	ojumun.FRectOrderSerial = orderserial
	ojumun.FCurrPage = page
	ojumun.FRectSellChannelDiv = sellchnl
	ojumun.FRectExcept3pl = not3pl  ''2017/03/29 �߰�
	ojumun.SearchJumunList

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function ViewOrderDetail(iorderserial){
	var ViewOrderDetail;
	ViewOrderDetail = window.open('/admin/ordermaster/viewordermaster.asp?orderserial=' + iorderserial,'ViewOrderDetail','scrollbars=yes,resizable=yes,width=1024,height=768');
    ViewOrderDetail.focus();
}



function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function SubmitForm(frm) {
	if ((CheckDateValid(frm.yyyy1.value, frm.mm1.value, frm.dd1.value) == true) && (CheckDateValid(frm.yyyy2.value, frm.mm2.value, frm.dd2.value) == true)) {
		frm.submit();
	}
}

// ������ �ٿ�ε�
function fnDownloadExcel() {
	var para = $("#frmSearch").serialize();
	document.location.href = '/admin/ordermaster/orderlist_excel.asp?'+para;
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm" id="frmSearch" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
	�ֹ���ȣ :
	<input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="16">
	&nbsp;
	�˻��Ⱓ :
	<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	(<input type="checkbox" name="cknodate" <% if cknodate="on" then response.write "checked" %> >�Ⱓ�������)
	<input type="checkbox" name="ckipkumdiv2" <% if ckipkumdiv2="on" then response.write "checked" %> >����Ǹ��˻�

    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitForm(document.frm);">
	</td>
</tr>
<tr>
    <td bgcolor="#F4F4F4">
	�˻����� :
	<select name="searchtype">
	<option value="">����</option>
	<option value="01" <% if searchtype="01" then response.write "selected" %> >������</option>
	<option value="02" <% if searchtype="02" then response.write "selected" %> >������</option>
	<option value="03" <% if searchtype="03" then response.write "selected" %> >���̵�</option>
	<option value="04" <% if searchtype="04" then response.write "selected" %> >�Ա���</option>
	<option value="06" <% if searchtype="06" then response.write "selected" %> >�����ݾ�</option>
	</select>
	<input type="text" name="searchrect" value="<%= searchrect %>" size="11" maxlength="16">
	&nbsp;&nbsp;
	�ֹ����� :
	<select name="jumundiv" class="select">
        <option value="">����</option>
        <option value="sendGift"   <% if jumundiv="sendGift"   then response.write "selected" %> >�����ϱ�</option>
		<option value="pojang" <% if jumundiv="pojang" then response.write "selected" %> >�����ֹ�</option>
        <option value="flowers" <% if jumundiv="flowers" then response.write "selected" %> >�ö���ֹ�</option>
        <option value="minus"   <% if jumundiv="minus"   then response.write "selected" %> >���̳ʽ�</option>
        <option value="foreign"   <% if jumundiv="foreign"   then response.write "selected" %> >�ؿܹ��</option>
        <option value="military"   <% if jumundiv="military"   then response.write "selected" %> >���δ�</option>
    </select>
    &nbsp;&nbsp;
    ä�α��� :
    <% drawSellChannelComboBox "sellchnl",sellchnl %>
    &nbsp;
    <input type="checkbox" name="not3pl" <%=CHKIIF(not3pl<>"","checked","")%> > 3PL��������
    &nbsp;
    <input type="checkbox" name="ckipkumdiv4" <% if ckipkumdiv4="on" then response.write "checked" %> >�����Ϸ��̻�˻�
	&nbsp;
	�ŷ����� : <% Call DrawIpkumDivName("ipkumdiv",ipkumdiv,"") %>
    </td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p style="padding:5px 0 4px 0; text-align:right;">
    <% if (C_ManagerUpJob or C_ADMIN_AUTH) then %>
	<img src="http://webadmin.10x10.co.kr/images/btn_excel.gif" alt="download excels" title="���� �ٿ�ε�" onclick="fnDownloadExcel();" style="cursor:pointer;"/>
    <% end if %>
</p>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr bgcolor="#FFFFFF">
	<td colspan="19">
		�� �Ǽ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FTotalCount,0) %></font>
		&nbsp;�� �ݾ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotal,0) %></font>
		&nbsp;��հ��ܰ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotal,0) %></font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="19" align="right">page : <%= ojumun.FCurrPage %>/<%=ojumun.FTotalPage %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="80" align="center">�ֹ���ȣ</td>
	<td width="40" align="center">����</td>
	<td width="60" align="center">ä��</td>
	<td width="60" align="center">Site</td>
	<td width="80" align="center">RdSite</td>
	<td width="70" align="center">UserID</td>

	<% if (C_InspectorUser = False) then %>
		<td width="70" align="center">���</td>
	<% end if %>

	<% if (FALSE) then %>
		<td width="60" align="center">������</td>
		<td width="65" align="center">������</td>
    <% end if %>

	<% if (C_InspectorUser = False) then %>
		<td width="60" align="center">�ֹ��Ѿ�</td>
		<td width="60" align="center">���ʽ�����</td>
		<td width="60" align="center">��ǰ����</td>
		<td width="50" align="center">��Ÿ����</td>
		<td width="60" align="center">���ϸ���</td>
	<% end if %>

	<td width="60" align="center">(��)������</td>
	<td width="74" align="center">�������</td>
	<td width="74" align="center">�ŷ�����</td>
	<td width="40" align="center">����<br>����</td>
	<td width="110" align="center">�ֹ���</td>
</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="19" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" bgcolor="#FFFFFF">
	<% else %>
	<tr class="gray" bgcolor="#FFFFFF">
	<% end if %>

		<td align="center"><a href="#" onclick="ViewOrderDetail('<%= ojumun.FMasterItemList(ix).FOrderSerial %>'); return false;" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td align="center" ><%= CHKIIF(ojumun.FMasterItemList(ix).IsForeignDeliver,ojumun.FMasterItemList(ix).FDlvcountryCode,"") %></td>
		<td align="center"><%= getSellChannelDivName(ojumun.FMasterItemList(ix).Fbeadaldiv) %> </td>
		<td align="center"><font color="<%= ojumun.FMasterItemList(ix).SiteNameColor %>"><%= ojumun.FMasterItemList(ix).FSitename %></font></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FRdSite %></td>

		<% if ojumun.FMasterItemList(ix).UserIDName<>"&nbsp;" then %>
			<td align="center">
				<%'= printUserId(ojumun.FMasterItemList(ix).UserIDName,2,"*") %>
				<%= ojumun.FMasterItemList(ix).UserIDName %>
			</td>
		<% else %>
			<td align="center"></td>
		<% end if %>

		<% if (C_InspectorUser = False) then %>
			<td align="center">
			    <% if ojumun.FMasterItemList(ix).FUserID="" then %>
	
			    <% else %>
					<font color="<%= getUserLevelColor(ojumun.FMasterItemList(ix).fUserLevel) %>"><%= getUserLevelStr(ojumun.FMasterItemList(ix).fUserLevel) %></font>
			    <% end if %>
			</td>
		<% end if %>

		<% if (FALSE) then %>
			<td align="center"><%= ojumun.FMasterItemList(ix).FBuyName %></td>
			<td align="center"><%= ojumun.FMasterItemList(ix).FReqName %></td>
		<% end if %>

		<% if (C_InspectorUser = False) then %>
			<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FTotalSum,0) %></td>
			<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).Fcouponpay,0) %></td>
			<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).getMayItemCouponDiscount,0) %></td>
			<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).Fallatdiscountprice,0) %></td>
			<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).Fmiletotalprice,0) %></td>
		<% end if %>

		<td align="right"><font color="<%= ojumun.FMasterItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FMasterItemList(ix).FSubTotalPrice,0) %></font></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).JumunMethodName %></td>
		<td align="center"><font color="<%= ojumun.FMasterItemList(ix).IpkumDivColor %>"><%= ojumun.FMasterItemList(ix).IpkumDivName %></font></td>
		<td align="center"><font color="<%= ojumun.FMasterItemList(ix).CancelYnColor %>"><%= ojumun.FMasterItemList(ix).CancelYnName %></font></td>
		<td align="center"><%= Left(ojumun.FMasterItemList(ix).GetRegDate,16) %></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="19" height="30" align="center">
			<% if ojumun.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
	
			<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
				<% if ix>ojumun.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>
	
			<% if ojumun.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% end if %>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
