<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��Ÿ�������
' History : 2009.04.07 ������ ����
'			2010.05.13 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<%
dim idx
dim ofranchulgojungsan, shopid

idx = RequestCheckvar(request("idx"),10)

if idx="" then idx="0"


'// ===========================================================================
set ofranchulgojungsan = new CEtcMeachul
ofranchulgojungsan.FRectidx = idx
ofranchulgojungsan.getOneEtcMeachul

dim IsMeaipPriceEditPossible	: IsMeaipPriceEditPossible = True
if (idx <> "0") and (ofranchulgojungsan.FOneItem.Fdivcode <> "GC") and (ofranchulgojungsan.FOneItem.Fdivcode <> "ET") then
	IsMeaipPriceEditPossible = False
end if

'// ===========================================================================
'���ͺμ����
Dim clsBS, arrBizList
Set clsBS = new CBizSection
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	clsBS.FSale = "N"
	arrBizList = clsBS.fnGetBizSectionList
Set clsBS = nothing
'// ===========================================================================
Dim defaultYYYY, defaultMM, defaultShopDiv
Dim i

IF idx="0" THen
    defaultYYYY = Left(DateAdd("m",-1,now()),4)
    defaultMM   = Mid(DateAdd("m",-1,now()),6,2)
    defaultShopDiv = ""
ELSE
    defaultYYYY = ""
    defaultMM   = ""
    defaultShopDiv = ""
END IF
%>
<script type='text/javascript'>

function SaveInfo(frm){
	if (frm.title.value.length<1){
		alert('Title�� �Է��ϼ���');
		frm.title.focus();
		return;
	}

	if (frm.shopdiv.value.length<1){
		alert('������ �Է��ϼ���');
		frm.shopdiv.focus();
		return;
	}

	if (frm.diffKey.value.length<1){
	    alert('���� ������ �Է��ϼ���');
		frm.diffKey.focus();
		return;
	}

	if (frm.shopdiv.value == "7") {
		if ((frm.papertype.value != "200") && (frm.papertype.value != "102")) {
			alert("������ ����(�ؿ�)�ΰ�� \n\n����Ű����� �Ǵ� ������꼭�� ���������� ����� �� �ֽ��ϴ�.");
			frm.papertype.focus();
			return;
		}
        // 20036 => 4010005
		if (frm.selltype.value != "4010005") {
			alert("������ ����(�ؿ�)�ΰ�� \n\n���������� ������ �����մϴ�.");
			frm.selltype.focus();
			return;
		}
	} else if (frm.shopdiv.value == "9") {
	    //if (frm.idx.value!=9861){
    		if (frm.papertype.value != "102") {
    			alert("������ �����ΰ�� \n\n������꼭�� ���������� ����� �� �ֽ��ϴ�.");
    			frm.papertype.focus();
    			return;
    		}
    	//}

		if (frm.selltype.value != "4010005") {
			alert("������ �����ΰ�� \n\n���������� ������ �����մϴ�.");
			frm.selltype.focus();
			return;
		}
	} else {
		if ((frm.papertype.value == "200") || (frm.papertype.value == "102")) {
			alert("���ó������ ���� �Ǵ� �����ΰ�츸 ��� �����մϴ�.");
			frm.papertype.focus();
			return;
		}

		if (frm.selltype.value == "4010005") {
			alert("���ó������ ���� �Ǵ� �����ΰ�츸 ��� �����մϴ�..");
			frm.selltype.focus();
			return;
		}
	}

<% if idx="0" then %>
	if (frm.shopid.value.length<1){
		alert('����ó�� �Է��ϼ���');
		frm.shopid.focus();
		return;
	}

	if (frm.totalbuycash.value.length<1){
		alert('�� ���԰��� �Է��ϼ���');
		frm.totalbuycash.focus();
		return;
	}

	if (frm.totalsuplycash.value.length<1){
		alert('�� ���ް��� �Է��ϼ���');
		frm.totalsuplycash.focus();
		return;
	}
<% elseif (IsMeaipPriceEditPossible) then %>
	if (frm.totalbuycash.value.length<1){
		alert('�� ���԰��� �Է��ϼ���');
		frm.totalbuycash.focus();
		return;
	}
<% end if %>

	if (frm.totalsum.value.length > 0) {
		frm.totalsum.value = replaceAll(frm.totalsum.value, ",", "");
	}

/*
	if (frm.totalsum.value.length<1){
		alert('�� ����ݾ��� �Է��ϼ���');
		frm.totalsum.focus();
		return;
	}


	if ((!frm.statecd[0].checked)&&(!frm.statecd[1].checked)&&(!frm.statecd[2].checked)&&(!frm.statecd[3].checked)){
		alert('���¸� �����ϼ���.');
		frm.statecd[0].focus();
		return;
	}
*/

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function escapeRegExp(string) {
    return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}

// frm.totalsum.value = replaceAll(frm.totalsum.value, ",", "");
function replaceAll(string, find, replace) {
  return string.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}

function changeState(state)
{
	var f = document.frm;

	switch (state)
	{
	case "0":
		var msg = "���������� �����Ͻðڽ��ϱ�?";
		break;
	case "1":
		var msg = "��üȮ�������� �����Ͻðڽ��ϱ�?";
		break;
	case "3":
		var msg = "��üȮ�οϷ�� �����Ͻðڽ��ϱ�?";
		break;
	case "7":
		var msg = "�ԱݿϷ�� �����Ͻðڽ��ϱ�?";
		if (f.ipkumdate.value.length!=10)
		{
			alert("�Ա����� �Է��Ͻʽÿ�.");
			return;
		}

		if (f.taxdate.value.length!=10)
		{
			alert("����������� �����ϴ�.\n\n�������� �ۼ� �� ����������� �Է��Ͻʽÿ�.");
			return;
		}

		break;
	}

	if (confirm(msg))
	{
		f.mode.value = "changeState";
		f.stateCd.value = state;
		f.submit();
	}
}

function changeIssueState(state)
{
	var f = document.frm;
	var msg = "";

	switch (state)
	{
	case "0":
		msg = "�����û���� �����Ͻðڽ��ϱ�?";
		break;
	case "9":
		msg = "����Ϸ�� �����Ͻðڽ��ϱ�?";
		if (f.taxdate.value.length != 10) {
			alert("����������� �����ϴ�.\n\n�������� �ۼ� �� ����������� �Է��Ͻʽÿ�.");
			return;
		}
		break;
	case "NULL":
		msg = "�������� ������ �����Ͻðڽ��ϱ�?";
		break;
	}

	if ((f.paperissuetype.value == "1") && (state == "NULL")) {
		// ����� ��꼭 ���� ����
		msg = "����� ��꼭�� ����û�� ���۵� ���\n�������ݰ�꼭�� �߰��� �����ؾ� �ϰ�\n\n���۵��� ���� ���\nBILL36524 ���� ����� ��꼭�� ����ؾ� �մϴ�.\n\n" + msg;
	}

	if (msg == "") {
		alert("ERROR");
		return;
	}

	if (confirm(msg) == true) {
		f.mode.value = "changeIssueState";
		f.issueStateCd.value = state;
		f.submit();
	}
}

function changeIpkumState(state)
{
	var f = document.frm;
	var msg = "";

	switch (state)
	{
	case "0":
		msg = "�Ա��������� �����Ͻðڽ��ϱ�?";
		break;
	case "5":
		msg = "�Ϻ��Ա����� �����Ͻðڽ��ϱ�?";
		break;
	case "9":
		msg = "�ԱݿϷ�� �����Ͻðڽ��ϱ�?";
		if (f.ipkumdate.value.length != 10) {
			alert("���� �Ա����� �Է��ϼ���");
			return;
		}
		break;
	case "NULL":
		msg = "�Աݻ��� ������ �����Ͻðڽ��ϱ�?";
		break;
	}

	if (msg == "") {
		alert("ERROR");
		return;
	}

	if (confirm(msg) == true) {
		f.mode.value = "changeIpkumState";
		f.ipkumStateCd.value = state;
		f.submit();
	}
}

function jsGetTax(ibizNo, itotSum){
	var sSearchText = ibizNo;
	var itotSum = itotSum;

	if (sSearchText == "2118700620") {
		sSearchText = "";
	}

	var winTax = window.open("/admin/tax/popSetEseroTax.asp?sST="+sSearchText+"&totSum="+itotSum+"&tgType=NRM&iTST=1","popGetTaxInfo","width=1200, height=800, resizable=yes, scrollbars=yes");
	winTax.focus();
}

function fillTaxInfo(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP){
    var frm = document.frm;

	frm.eserotaxkey.value = eTax;
}

</script>

<form name="frm" method=post action="/admin/offshop/etc_meachul_process.asp" style="margin:0px;">
<input type=hidden name="idx" value="<%= ofranchulgojungsan.FOneItem.Fidx %>">
<% if idx="0" then %>
<input type=hidden name="mode" value="addmaster">
<% else %>
<input type=hidden name="mode" value="modimaster">
<input type="hidden" name="stateCd" value="">
<input type="hidden" name="issueStateCd" value="">
<input type="hidden" name="ipkumStateCd" value="">
<% end if %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>IDX</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fidx %></td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">����ó</td>
		<% if idx="0" then %>
		<td bgcolor="#FFFFFF" >
			<% NewdrawSelectBoxShopAll "shopid", shopid %>
		</td>
		<% else %>
		<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fshopid %></td>
		<% end if %>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<% if idx="0" then %>
			<td bgcolor="#FFFFFF" ><% call DrawYMBox(defaultYYYY,defaultMM) %></td>
		<% else %>
			<td bgcolor="#FFFFFF" >
				<% if C_ADMIN_AUTH or C_MngPart or C_PSMngPart then %>
					<% call DrawYMBox(Left(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),4),Right(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),2)) %>
					�� ������,�繫���� ��������
				<% else %>
					<%= Left(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),4) %>-<%= Right(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),2) %>
					<input type="hidden" name="yyyy1" value="<%= Left(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),4) %>">
					<input type="hidden" name="mm1" value="<%= Right(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),2) %>">
				<% end if %>
			</td>
		<% end if %>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF" >
			<% if idx="0" then %>
				<% Call DrawShopDivBox(defaultShopDiv) %>
				/
				<select class="select" name="divcode">
					<option value="GC">���ͺ�
					<option value="ET">��Ÿ����
				</select>
			<% else %>
				<% Call DrawShopDivBox(ofranchulgojungsan.FOneItem.FShopDiv) %>
				/
				<font color="<%= ofranchulgojungsan.FOneItem.GetDivCodeColor %>"><%= ofranchulgojungsan.FOneItem.GetDivCodeName %></font>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	    <td bgcolor="#FFFFFF" >
	    <% if idx="0" then %>
	    <input type="text" name="diffKey" maxlength="2" class="text">
	    <% else %>
	    <input type="text" name="diffKey" value="<%= ofranchulgojungsan.FOneItem.FdiffKey %>" size="2" maxlength="2" class="text">
	    <% end if %>
	    </td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">Title</td>
		<td bgcolor="#FFFFFF" >
			<input type="text" class="text" name=title value="<%= ofranchulgojungsan.FOneItem.Ftitle %>" size="40" maxlength="40" <%If ofranchulgojungsan.FOneItem.Fstatecd>="4" Then %>readOnly<%End If %> >
			(ex) OO�� 4�� 1�� ��ǰ��
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�ѼҺ��ڰ�</td>
		<td bgcolor="#FFFFFF">
			<% if idx="0" then %>

			<% else %>
			<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsellcash,0) %>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>"><b>�����</b></td>
		<td bgcolor="#FFFFFF">
			<% if idx="0" then %>
			<input type=text name=totalsuplycash value="" size=10 maxlength=10 style="border:1px #999999 solid; text-align=right">
			<font color="#AAAAAA">(����ó�� ������ ��ǰ����)</font>
			<% else %>
			<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsuplycash,0) %>
			<font color="#AAAAAA">(����ó�� ������ ��ǰ����)</font>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�Ѹ��԰�</td>
		<td bgcolor="#FFFFFF">
			<% if idx="0" then %>
				<input type=text name=totalbuycash value="" size=10 maxlength=10 style="border:1px #999999 solid; text-align=right">
				<font color="#AAAAAA">(�ҿ� ���:����)</font>
			<% else %>
				<% if IsMeaipPriceEditPossible then %>
					<input type=text name=totalbuycash value="<%= ofranchulgojungsan.FOneItem.Ftotalbuycash %>" size=10 maxlength=10 style="border:1px #999999 solid; text-align=right">
				<% else %>
					<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalbuycash,0) %>
					<input type="hidden" name="totalbuycash" value="<%= ofranchulgojungsan.FOneItem.Ftotalbuycash %>">
				<% end if %>
				<font color="#AAAAAA">(��ü�κ��� ���޹��� ��ǰ����)</font>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<td bgcolor="#FFFFFF" >
		<font color="<%= ofranchulgojungsan.FOneItem.GetStateColor %>"><%= ofranchulgojungsan.FOneItem.GetStateName %></font>

		<% if (ofranchulgojungsan.FOneItem.Fstatecd="0") then %>
		==&gt; <input type="button" class="button" onclick="changeState('1');" value="��üȮ�������� ����">
		<% elseif (ofranchulgojungsan.FOneItem.Fstatecd="1") then %>
		==&gt; <input type="button" class="button" onclick="changeState('3');" value="��üȮ�οϷ�� ����">
		<% elseif (ofranchulgojungsan.FOneItem.Fstatecd="4") or (ofranchulgojungsan.FOneItem.Fstatecd="3") then %>
		==&gt; <input type="button" class="button" onclick="changeState('7');" value="�Ϸ� �� ����">
		<% else %>
		<% end if %>

		<% if (ofranchulgojungsan.FOneItem.Fstatecd="1") or (ofranchulgojungsan.FOneItem.Fstatecd="3") then %>
		<input type="button" class="button" onclick="changeState('0');" value="���������� ����">
		<% elseif (ofranchulgojungsan.FOneItem.Fstatecd="4") then %>
		<input type="button" class="button" onclick="changeState('0');" value="���������� ����">
		<% else %>

	    <% end if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ����</td>
		<td bgcolor="#FFFFFF" >
			<textarea name="etcstr" class="textarea" cols="86" rows="8"><%= ofranchulgojungsan.FOneItem.Fetcstr %></textarea>
		</td>
	</tr>

	<tr>
		<td height="10" bgcolor="#FFFFFF" style="padding: 1px;"></td>
		<td bgcolor="#FFFFFF" style="padding: 1px;"></td>
	</tr>

	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">����μ�</td>
		<td bgcolor="#FFFFFF">
	        <select class="select" name="bizsection_cd">
	        <option value="">--����--</option>
	        <% For i = 0 To UBound(arrBizList,2)	%>
	    		<option value="<%=arrBizList(0,i)%>" <%IF (ofranchulgojungsan.FOneItem.Fbizsection_cd) = Cstr(arrBizList(0,i)) THEN%> selected <%END IF%>><%=arrBizList(1,i)%></option>
	    	<% Next %>
	        </select>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td bgcolor="#FFFFFF">
			<% drawPartnerCommCodeBox true,"sellacccd","selltype",ofranchulgojungsan.FOneItem.Fselltype,"" %>
		</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name="papertype">
				<option value="">����
				<option value="100" <% if ofranchulgojungsan.FOneItem.Fpapertype="100" then response.write "selected" %> > ���� ��꼭
				<option value="101" <% if ofranchulgojungsan.FOneItem.Fpapertype="101" then response.write "selected" %> > �鼼 ��꼭
				<option value="102" <% if ofranchulgojungsan.FOneItem.Fpapertype="102" then response.write "selected" %> > ���� ��꼭
				<option value="200" <% if ofranchulgojungsan.FOneItem.Fpapertype="200" then response.write "selected" %> > ����Ű�����
				<option value="999" <% if ofranchulgojungsan.FOneItem.Fpapertype="999" then response.write "selected" %> > ����
			</select>

	        <select class="select" name="paperissuetype">
	        	<option value="">--����--</option>
				<option value="1" <%IF (ofranchulgojungsan.FOneItem.Fpaperissuetype = "1") THEN%> selected <%END IF%>>������</option>
				<option value="2" <%IF (ofranchulgojungsan.FOneItem.Fpaperissuetype = "2") THEN%> selected <%END IF%>>������</option>
	        </select>
	        *������ = ������ ����
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�̼���</td>
		<td bgcolor="#FFFFFF">
			<% if ofranchulgojungsan.FOneItem.Fpaperissuetype = "2" then %>
				<input type="text" class="text_ro" name="eserotaxkey" value="<%= ofranchulgojungsan.FOneItem.Feserotaxkey %>" size="30" maxlength="32" readonly>
				<input type="button" class="button" value="�˻�" onClick="jsGetTax('<%= ofranchulgojungsan.FOneItem.FbizNo %>','<%= ofranchulgojungsan.FOneItem.Ftotalsum %>');">
		    <% else %>
		        <%= ofranchulgojungsan.FOneItem.Feserotaxkey %>
		        <% if IsNull(ofranchulgojungsan.FOneItem.Feserotaxkey) then %>��Ī����<% end if %>
			<% end if %>
		</td>
	</tr>

	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>"><b>�ѹ���ݾ�</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="totalsum" value="<%= ofranchulgojungsan.FOneItem.Ftotalsum %>" size="10" maxlength="10" style="text-align=right">
			<font color="#AAAAAA">(��꼭 ���� �ݾ�)</font>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">���������</td>
		<td bgcolor="#FFFFFF">
   			<input type="text" id="termTaxDt" name="taxdate" readonly size="11" maxlength="10" value="<%= ofranchulgojungsan.FOneItem.Ftaxdate %>" class="text_ro" style="text-align:center;" />
			<%if (ofranchulgojungsan.FOneItem.Fstatecd > "0") or (ofranchulgojungsan.FOneItem.Fstatecd < "7") then %>
				<% if (Not IsNull(ofranchulgojungsan.FOneItem.Fpapertype)) and (ofranchulgojungsan.FOneItem.Fpapertype <> "100" and Not (ofranchulgojungsan.FOneItem.Fpapertype = "200" and IsNull(ofranchulgojungsan.FOneItem.Finvoiceidx))) then %>
				<img src="/images/calicon.gif" align="absmiddle" border="0" id="btnTaxDt" style="cursor:pointer;" />
				<script type="text/javascript">
					var CAL_TaxDate = new Calendar({
						inputField : "termTaxDt", trigger    : "btnTaxDt",
						bottomBar: true, dateFormat: "%Y-%m-%d",
						onSelect: function() {
							this.hide();
						}
					});
				</script>
				<% end if %>
			<% end if %>
			<font color="#AAAAAA">(��꼭������,����Ű�����)</font>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<td bgcolor="#FFFFFF" >
			<%If idx <> "0" Then %>

				<%= ofranchulgojungsan.FOneItem.GetIssueStateName %>(<%= ofranchulgojungsan.FOneItem.Fpaperissuetype %>)

				<% if (ofranchulgojungsan.FOneItem.Fpaperissuetype = "1") then %>
					<% if (C_ADMIN_AUTH) and (ofranchulgojungsan.FOneItem.FIssueStateCD="9") then %>
					<input type="button" class="button" onclick="changeIssueState('NULL');" value="�������� ����"> [�����ں�]
					<% end if %>
				<% elseif (ofranchulgojungsan.FOneItem.Fpaperissuetype = "2") then %>

					<% if IsNull(ofranchulgojungsan.FOneItem.FIssueStateCD) then %>
						==&gt;
						<input type="button" class="button" onclick="changeIssueState('0');" value="�����û���� ����">
						<input type="button" class="button" onclick="changeIssueState('9');" value="����Ϸ�� ����">
					<% else %>
						==&gt;
						<% if (ofranchulgojungsan.FOneItem.FIssueStateCD="0") then %>
							<input type="button" class="button" onclick="changeIssueState('9');" value="����Ϸ�� ����">
						<% elseif (ofranchulgojungsan.FOneItem.FIssueStateCD="9") then %>

						<% else %>
							ERROR
						<% end if %>
						<input type="button" class="button" onclick="changeIssueState('NULL');" value="�������� ����" <% if (Not C_ADMIN_AUTH) then %>disabled<% end if %> > <% if (C_ADMIN_AUTH) then %>[�����ں�]<% end if %>
					<% end if %>

				<% end if %>

	    	<% end if %>
		</td>
	</tr>
	<tr>
		<td height="10" bgcolor="#FFFFFF" style="padding: 1px;"></td>
		<td bgcolor="#FFFFFF" style="padding: 1px;"></td>
	</tr>

	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�Ա���</td>
		<td bgcolor="#FFFFFF">
   			<input type="text" id="termIpkumDt" name="ipkumdate" readonly size="11" maxlength="10" value="<%= ofranchulgojungsan.FOneItem.Fipkumdate %>" class="text_ro" style="text-align:center;" />
			<%if (ofranchulgojungsan.FOneItem.Fstatecd > "0") or (ofranchulgojungsan.FOneItem.Fstatecd < "7") then %>
			<img src="/images/calicon.gif" align="absmiddle" border="0" id="btnIpkumDt" style="cursor:pointer;" />
			<script type="text/javascript">
				var CAL_IpkumDate = new Calendar({
					inputField : "termIpkumDt", trigger    : "btnIpkumDt",
					bottomBar: true, dateFormat: "%Y-%m-%d",
					onSelect: function() {
						this.hide();
					}
				});
			</script>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<td bgcolor="#FFFFFF" >
			<%If idx <> "0" Then %>

				<%= ofranchulgojungsan.FOneItem.GetIpkumStateName %>

				<% if IsNull(ofranchulgojungsan.FOneItem.FIpkumStateCD) then %>
					==&gt;
					<!--
					<input type="button" class="button" onclick="changeIpkumState('0');" value="�Ա��������� ����">
					-->
					<input type="button" class="button" onclick="changeIpkumState('5');" value="�Ϻ��Ա����� ����">
					<input type="button" class="button" onclick="changeIpkumState('9');" value="�ԱݿϷ�� ����">
				<% else %>
					==&gt;
					<% if (ofranchulgojungsan.FOneItem.FIpkumStateCD="0") then %>
						<input type="button" class="button" onclick="changeIpkumState('5');" value="�Ϻ��Ա����� ����">
						<input type="button" class="button" onclick="changeIpkumState('9');" value="�ԱݿϷ�� ����">
					<% elseif (ofranchulgojungsan.FOneItem.FIpkumStateCD="5") then %>
						<input type="button" class="button" onclick="changeIpkumState('9');" value="�ԱݿϷ�� ����">
					<% elseif (ofranchulgojungsan.FOneItem.FIpkumStateCD="9") then %>

					<% else %>
						ERROR
					<% end if %>
					<input type="button" class="button" onclick="changeIpkumState('NULL');" value="�Աݻ��� ����">
				<% end if %>

	    	<% end if %>
		</td>
	</tr>
	<tr>
		<td height="10" bgcolor="#FFFFFF" style="padding: 1px;"></td>
		<td bgcolor="#FFFFFF" style="padding: 1px;"></td>
	</tr>

	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">���ʵ����</td>
		<td bgcolor="#FFFFFF"><%= ofranchulgojungsan.FOneItem.Fregusername %>(<%= ofranchulgojungsan.FOneItem.Freguserid %>)</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">����ó����</td>
		<td bgcolor="#FFFFFF"><%= ofranchulgojungsan.FOneItem.Ffinishusername %>(<%= ofranchulgojungsan.FOneItem.Ffinishuserid %>)</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td bgcolor="#FFFFFF"><%= ofranchulgojungsan.FOneItem.Fregdate %></td>
	</tr>
	<tr height="30">
		<td colspan=2 align=center bgcolor="#FFFFFF">
		<%If idx="0" Then %>
			<input type="button" class="button" value="��������" onclick="SaveInfo(frm);">
		<% else %>
			<input type="button" class="button" value="��ü����" onclick="SaveInfo(frm);">
		<%End If %>

		</td>
	</tr>
</table>
</form>
<%
set ofranchulgojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
