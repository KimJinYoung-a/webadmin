<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� taxRefund ����
' History : 2014.01.17 ������
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/payment/taxRefundMngCls.asp"-->
<%
dim page,shopid,yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate,toDate, Searchtaxrefundkey, schType, scgRealsum, jyyyymm
dim datefg , i, ToTcashsum, intLoop, isedityn, inc3pl
	shopid = requestCheckvar(request("shopid"),32)
	page = requestCheckvar(request("page"),10)
	if page="" then page=1
	yyyy1 = requestCheckvar(request("yyyy1"),4)
	mm1 = requestCheckvar(request("mm1"),2)
	dd1 = requestCheckvar(request("dd1"),2)
	yyyy2 = requestCheckvar(request("yyyy2"),4)
	mm2 = requestCheckvar(request("mm2"),2)
	dd2 = requestCheckvar(request("dd2"),2)

	jyyyymm = requestCheckvar(request("jyyyymm"),7)

	datefg = requestCheckvar(request("datefg"),10)
    inc3pl = requestCheckvar(request("inc3pl"),10)
	Searchtaxrefundkey = requestCheckvar(request("Searchtaxrefundkey"),30)
	schType = requestCheckvar(request("schType"),10)
	scgRealsum = requestCheckvar(request("scgRealsum"),10)
if datefg = "" then datefg = "maechul"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-0)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/����
if (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if


dim oTaxRefund
set oTaxRefund = new CTaxRefund
	oTaxRefund.FRectShopID = shopid
	oTaxRefund.FRectStartDay = fromDate
	oTaxRefund.FRectEndDay = toDate
	oTaxRefund.frectdatefg = datefg
	oTaxRefund.frecttaxrefundkey = Searchtaxrefundkey
	oTaxRefund.frectscgRealsum = scgRealsum
	oTaxRefund.frectschType = schType
	if (Len(jyyyymm)=7) then
	    oTaxRefund.FRectRefundMonth = jyyyymm
    end if
	''oTaxRefund.FRectInc3pl = inc3pl
	oTaxRefund.FPageSize = 200
	oTaxRefund.FCurrPage = page

	if (shopid<>"") then
		oTaxRefund.GetTaxRefundTargetList
	else
		response.write "<script language='javascript'>"
		response.write "alert('������ �����Ͻ� �� �˻��ϼ���.');"
		response.write "</script>"
	end if

dim totcnt, totrealsum, totVatsum

Dim defaultrefundCode
Select Case shopid
	'Case "streetshop011"	defaultrefundCode = "20025720390513"		'���з�
	Case "streetshop011"	defaultrefundCode = "20023120332514"		'���з� '2014/02/05 ������ ����. ������븮�� ��û
	'Case "streetshop014"	defaultrefundCode = "20023120332513"		'��Ÿ
	Case "streetshop014"	defaultrefundCode = "20025720390514"		'��Ÿ '2014/02/05 ������ ����. ������븮�� ��û
	'Case "streetshop018"	defaultrefundCode = "20025710131013"		'�����Ե�
	Case "streetshop018"	defaultrefundCode = "20025710131014"		'�����Ե� '2014/02/05 ������ ����. ������븮�� ��û
End Select
%>
<script language='javascript'>
function addRefundKey(comp, chkid, btnIid, btnSid){
	document.getElementById(chkid).disabled = false;
	document.getElementById(btnIid).style.display = "none";
	document.getElementById(btnSid).style.display = "block";
	document.getElementById(chkid).focus();
	document.getElementById(chkid).value = "<%=defaultrefundCode%>";
}
function updateRefundKey(comp, chkid){
	if(document.getElementById(chkid).value.length < 20){
		alert('20�� �̳��� �Է��ϼ���');
		document.getElementById(chkid).value = "<%=defaultrefundCode%>";
		document.getElementById(chkid).focus();
		return false;
	}
	document.frmSvArr.target = "xLink";
	document.frmSvArr.cmdparam.value = "U";
	document.frmSvArr.midx.value = comp;
	document.frmSvArr.refundkey.value = document.getElementById(chkid).value;
	document.frmSvArr.action = "/admin/offshop/payment/taxRefund_process.asp"
	document.frmSvArr.submit();
}
function delRefundKey(comp){
	if(confirm("���� �Ͻðڽ��ϱ�?")){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "D";
		document.frmSvArr.midx.value = comp;
		document.frmSvArr.action = "/admin/offshop/payment/taxRefund_process.asp"
		document.frmSvArr.submit();
	}
}
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="A">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">�˻�<br>����</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ :
				<% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% drawSelectBoxOffShopAll "shopid",shopid %>
					<% end if %>
				<% else %>
					* ���� : <% drawSelectBoxOffShopAll "shopid",shopid %>
				<% end if %>
				<!--
	            &nbsp;&nbsp;
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	            -->
	            &nbsp;&nbsp;
	            * ����� : <input type="text" name="jyyyymm" value="<%=jyyyymm%>" size="7" maxlength="7"> (YYYY-MM)
			</td>
		</tr>
	    </table>
    </td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
    <td>
    * �˻����� :
    <select name="schType" class="select">
	    <option value="">��ü
	    <option value="0" <%= Chkiif(schType="0","selected","") %> >taxRefund �Է³���
	    <option value="1" <%= Chkiif(schType="1","selected","") %>>taxRefund ���Է³���
	    <option value="2" <%= Chkiif(schType="2","selected","") %>>�ܱ��α��ų���
    </select>

    &nbsp;&nbsp;
    * �����ݾ� :
    <input type="text" name="scgRealsum" size="10" maxlength="10" value="<%= scgRealsum %>">

    &nbsp;&nbsp;
    * TaxRefund�Ϸù�ȣ :
    <input type="text" name="Searchtaxrefundkey" size="25" maxlength="20" value="<%=Searchtaxrefundkey%>"> (20�ڸ�)
    </td>
</tr>

</form>
</table>
<!-- ǥ ��ܹ� ��-->
<Br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oTaxRefund.FTotalCount %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>�ֹ���ȣ</td>
	<td>������</td>
	<td>�ΰ���</td>
	<!--
	<td>ī��</td>
	<td>����</td>
	<td>���ϸ���</td>
	<td>��ǰ��</td>
	<td>����Ʈī��</td>
	-->
	<td>������</td>
	<td>�ܱ��ο���</td>
	<td>�����</td>
	<td>TaxRefund�ڵ�</td>
	<td>���</td>
</tr>
<%
if oTaxRefund.FResultCount > 0 then
for i=0 to oTaxRefund.FResultCount -1
totcnt = totcnt +1
totrealsum=totrealsum+oTaxRefund.FItemList(i).Frealsum
totVatsum=totVatsum+CLNG(FIX(oTaxRefund.FItemList(i).Frealsum/11))
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><%= oTaxRefund.FItemList(i).ForderNo %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).Frealsum,0) %></td>
	<td align="right"><%= FormatNumber(FIX(oTaxRefund.FItemList(i).Frealsum/11),0) %></td>
	<!--
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).Fcardsum,0) %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).Fcashsum,0) %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).Fspendmile,0) %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).FGiftCardPaySum,0) %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).FTenGiftCardPaySum,0) %></td>
	-->
	<td><%= oTaxRefund.FItemList(i).Fshopregdate %></td>
	<td>
	<%
		Select Case oTaxRefund.FItemList(i).Fbuyergubun
			Case "100"	response.write "������"
			Case "200"	response.write "�ܱ���"
			Case Else	response.write "��üũ"
		End Select
	%>
	</td>
	<td><%= oTaxRefund.FItemList(i).FrefundMonth %></td>
	<td>
		<input type="text" id="taxrefundkey<%=i%>" name="taxrefundkey" maxlength="20" size="25" value="<%= oTaxRefund.FItemList(i).Ftaxrefundkey %>" disabled >
	</td>

	<td>
	<% if isNULL(oTaxRefund.FItemList(i).Ftaxrefundkey) then %>
	<input type="button" class="button" id="btnI<%=i%>" value="�Է�" style="display:block;" onClick="addRefundKey('<%= oTaxRefund.FItemList(i).Fidx %>','taxrefundkey<%= i%>','btnI<%=i%>','btnS<%=i%>')">
	<input type="button" class="button" id="btnS<%=i%>" value="����" style="display:none;" onClick="updateRefundKey('<%= oTaxRefund.FItemList(i).Fidx %>', 'taxrefundkey<%=i%>')">
	<% else %>
	<input type="button" class="button" value="����" onClick="delRefundKey('<%= oTaxRefund.FItemList(i).Fidx %>')">
	<% end if %>
	</td>
</tr>
<%
next
%>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td >�հ�</td>
	<td align="right"><%= FormatNumber(totrealsum,0) %></td>
	<td align="right"><%= FormatNumber(totVatsum,0) %></td>
	<!--
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	-->
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
</tr>
<tr height="20">
    <td colspan="16" align="center" bgcolor="#FFFFFF">
        <% if oTaxRefund.HasPreScroll then %>
		<a href="javascript:goPage('<%= oTaxRefund.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oTaxRefund.StartScrollPage to oTaxRefund.FScrollCount + oTaxRefund.StartScrollPage - 1 %>
    		<% if i>oTaxRefund.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oTaxRefund.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="15">�˻� ����� �����ϴ�.</td>
</tr>
<%
end if
%>
</table>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="midx" value="">
<input type="hidden" name="refundkey" value="">
<input type="hidden" name="refundmonth" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->