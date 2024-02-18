<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ����
' History : 2011.12.27 �ѿ�� ����
'						2014.01.03 ������ ���� �˻����� �� �ʵ� �߰�
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/delaytaxcls.asp"-->
<%
dim i, j
dim yyyy1, mm1, yyyy2, mm2,yyyy3,mm3
dim yyyymm, endyyyymm, issueyyymm, makerid ,offgubun, issuegubun
dim designer,groupid ,vPurchaseType,erpCustcd, jgubun, companynoYN
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	yyyy3 = requestCheckVar(request("yyyy3"),4)
	mm3 = requestCheckVar(request("mm3"),2)
	offgubun = requestCheckVar(request("offgubun"),3)
	issuegubun = requestCheckVar(request("issuegubun"),1)
	designer = requestCheckVar(request("designer"),32)
	groupid  = requestCheckVar(request("groupid"),32)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	erpCustcd = requestCheckVar(request("erpCustcd"),16)
	jgubun   = requestCheckVar(request("jgubun"),10)
	companynoYN = requestCheckVar(request("companynoYN"),1)
Dim jacctcdexists : jacctcdexists =requestCheckVar(request("jacctcdexists"),10)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 =  format00(2,cstr(Month(now())))
if (yyyy2="") then yyyy2 = yyyy1
if (mm2="") then mm2 = mm1

yyyymm = yyyy1 + "-" + mm1
endyyyymm = yyyy2 + "-" + mm2
if yyyy3 <> "" then
issueyyymm = yyyy3 + "-" + mm3
end if

If offgubun <> "ON" Then
'	companynoYN = ""
End If

dim ocdelaytax
set ocdelaytax = new CDelayTax
	ocdelaytax.FRectStartYYYYMM = yyyymm
	ocdelaytax.FRectEndYYYYMM = endyyyymm
	ocdelaytax.FRectIssueYYYYMM = issueyyymm
	ocdelaytax.FRectGubun = offgubun
	ocdelaytax.FRectIssueGubun = issuegubun
	ocdelaytax.FRectdesigner = designer
	ocdelaytax.FRectGroupid = groupid
	ocdelaytax.FRectPurchaseType = vPurchaseType
	ocdelaytax.FRecterpCustcd = erpCustcd
    ocdelaytax.FRectJGubun = jgubun
    ocdelaytax.FRectCompanynoYN = companynoYN
	ocdelaytax.FRectJacctcdExists = jacctcdexists
	ocdelaytax.GetDelayTaxDetailList

%>

<script type="text/javascript">

function formSubmit(page) {
	frm.page.value=page;
	frm.submit();
}
function popSearchGroupID(frmname,compname){
    var popwin = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&compname=" + compname,"popSearchGroupID","width=800 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���� :
		<select class="select" name="offgubun">
		<option value="ON" <% if (offgubun = "ON") then %>selected<% end if %> >�¶���</option>
		<option value="OFF" <% if (offgubun = "OFF") then %>selected<% end if %> >��������</option>
		<option value="ETC" <% if (offgubun = "ETC") then %>selected<% end if %> >��Ÿ����</option>
		</select>
		&nbsp;
		����� :  <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
		&nbsp;
		����� : <% Call DrawYMBoxdynamic("yyyy3", yyyy3, "mm3", mm3, "") %>
		&nbsp;
		���౸�� :
		<select class="select" name="issuegubun">
		<option value="1" <% if (issuegubun = "1") then %>selected<% end if %> >�������</option>
		<option value="2" <% if (issuegubun = "2") then %>selected<% end if %> >��������</option>
		<option value="9" <% if (issuegubun = "9") then %>selected<% end if %> >��Ÿ����(������)</option>
		</select>
		&nbsp;
		�����ı��� :
        <% drawSelectBoxJGubun "jgubun",jgubun %>
		&nbsp;&nbsp;
		* �ٹ����� ����� ���� : 
        <select name="companynoYN" class="select">
			<option value="">��ü
			<option value="Y" <%= CHKIIF(companynoYN="Y","selected","") %> >����ڸ�
			<option value="N" <%= CHKIIF(companynoYN="N","selected","") %> >���������
		</select>
		&nbsp;&nbsp;
		<input type="checkbox" name="jacctcdexists" <%= CHKIIF(jacctcdexists="on","checked","") %> >�������� ���� ���길 ����
        

	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="formSubmit('1');">
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" >
		�������� : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
		&nbsp;&nbsp;
		�귣��ID : <% drawSelectBoxDesignerwithName "designer",designer  %>&nbsp;&nbsp;
	  ��ü(�׷��ڵ�) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >
		<input type="button" class="button" value="Code�˻�" onclick="popSearchGroupID(this.form.name,'groupid');" >&nbsp;&nbsp;
		ERP�����ڵ� : <input type="text" class="input" name="erpCustcd" value="<%=erpCustcd%>" size="16">
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
		* �ִ� 3,000�Ǳ��� ǥ�õ˴ϴ�.
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		�˻���� : <b><%= formatnumber(ocdelaytax.FTotalCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">�����</td>
	<td width="150">�귣��</td>
	<td width="150">����ڸ�</td>
	<td>��������<br/>�޴���</td>
	<td width="60">ERP�����ڵ�</td>
	<td width="60">�׷��ڵ�</td>
	<td width="60">������<br/>��������</td>
	<td width="60">����</td>
	<td width="170">�̼���</td>
	<td width="80">�����</td>
	<td width="80">������</td>
	<td width="80">�Ա���</td>
	<td width="80">�����<br>(�����)</td>
	<td width="80">��������</td>
	<td>����ι�</td>
	<td>�������</td>
	<td>����</td>
	<td>���</td>
</tr>
<%
if ocdelaytax.FresultCount > 0 then
%>
	<%
	for i=0 to ocdelaytax.FresultCount-1
	%>
		<tr bgcolor="#FFFFFF" align="center">
			<td nowrap><%= ocdelaytax.FItemList(i).Fyyyymm %></td>
			<td><%= ocdelaytax.FItemList(i).Fmakerid %></td>
			<td><%= ocdelaytax.FItemList(i).Fcompany_name %></td>
			<td><%= ocdelaytax.FItemList(i).fjungsan_hp %></td>
			<td><%= ocdelaytax.FItemList(i).Ferpcust_cd %></td>
			<td><%= ocdelaytax.FItemList(i).Fgroupid %></td>
			<td><%= ocdelaytax.FItemList(i).getTaxTypeName %></td>
			<td><%= ocdelaytax.FItemList(i).Fjungsan_gubun %></td>
			<td><%= ocdelaytax.FItemList(i).Feserotaxkey %></td>
			<td  nowrap><%= ocdelaytax.FItemList(i).Ftaxinputdate %></td>
			<td  nowrap><%= ocdelaytax.FItemList(i).Ftaxregdate %></td>
			<td  nowrap><%= ocdelaytax.FItemList(i).Fipkumdate %></td>
		<td  nowrap align="right"><%= FormatNumber(ocdelaytax.FItemList(i).FjungsanPrice,0)  %></td>
			<td><%= ocdelaytax.FItemList(i).FpurchasetypeName %></td>
			<td><%= ocdelaytax.FItemList(i).FbizsectionName %></td>
			<td><%= ocdelaytax.FItemList(i).FselltypeName %></td>
			<td><%= ocdelaytax.FItemList(i).GetFinishFlagName %></td>
			<td></td>
		</tr>
	<% next %>
		<tr height="25" bgcolor="#ffffff">
			<td colspan="12" align="center">�հ�</td>
			<td align="right"><%=FormatNumber(ocdelaytax.FTot_jungsanPrice,0)%></td>
			<td colspan="5"></td>
		</tr>
<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="18">�˻������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<%
set ocdelaytax = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp"-->