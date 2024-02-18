<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_refundcheckcls.asp"-->
<%

''http://webadmin.10x10.co.kr/cscenter/refund/refund_check_list.asp?page=4&research=on&menupos=1&divcd=A004&yyyy1=2015&mm1=08&dd1=01&yyyy2=2015&mm2=08&dd2=31&returnmethod=&orderserial=&refundMin=&refundMax=&chkGubun=retbea

dim research, page, i
dim divcd, returnmethod, orderserial, chkGubun, refundMin, refundMax
dim yyyy1, yyyy2, mm1, mm2, dd1, dd2
dim fromDate, toDate
dim exCheckFinish
dim returnmethodIN, retR007, retR910, retR900, dategbn

'===============================================================================
research 		= requestCheckVar(request("research"),32)
page 			= requestCheckVar(request("page"),32)
divcd 			= requestCheckVar(request("divcd"),32)
returnmethod 	= requestCheckVar(request("returnmethod"),32)
orderserial 	= requestCheckVar(request("orderserial"),32)
chkGubun 		= requestCheckVar(request("chkGubun"),32)
refundMin 		= requestCheckVar(request("refundMin"),32)
refundMax 		= requestCheckVar(request("refundMax"),32)
exCheckFinish 	= requestCheckVar(request("exCheckFinish"),32)
retR007 		= requestCheckVar(request("retR007"),32)
retR910 		= requestCheckVar(request("retR910"),32)
retR900 		= requestCheckVar(request("retR900"),32)
dategbn     = requestCheckvar(request("dategbn"),32)
'===============================================================================
yyyy1   = request("yyyy1")
yyyy2   = request("yyyy2")
mm1     = request("mm1")
mm2     = request("mm2")
dd1     = request("dd1")
dd2     = request("dd2")

if (yyyy1="") then
	fromDate = CStr(DateSerial(Year(Now()), (Month(Now()) - 1), 1))
	toDate = CStr(DateSerial(Year(Now()), Month(Now()), 0))

    yyyy1 = CStr(Year(fromDate))
    mm1 = CStr(Month(fromDate))
    dd1 =  CStr(day(fromDate))

    yyyy2 = CStr(Year(toDate))
    mm2 = CStr(Month(toDate))
    dd2 =  CStr(day(toDate))
end if

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if dategbn="" then dategbn="finishdate"

if (retR007 <> "") or (retR910 <> "") or (retR900 <> "") then
	returnmethodIN = "'XXXX'"
	if (retR007 <> "") then
		returnmethodIN = returnmethodIN + ",'R007'"
	end if
	if (retR910 <> "") then
		returnmethodIN = returnmethodIN + ",'R910'"
	end if
	if (retR900 <> "") then
		returnmethodIN = returnmethodIN + ",'R900'"
	end if
end if


'===============================================================================
if (page="") then page = 1
if (research="") then
	''divcd = "A003"
	chkGubun = "err"
	''exCheckFinish = "Y"
end if


'===============================================================================
dim oCCSRefundCheck

set oCCSRefundCheck = new CCSRefundCheck


oCCSRefundCheck.FPageSize = 50
oCCSRefundCheck.FCurrPage = page

oCCSRefundCheck.FRectOrderSerial = orderserial
oCCSRefundCheck.FRectDivCD = divcd
oCCSRefundCheck.FRectReturnMethod = returnmethod
oCCSRefundCheck.FRectStartDate = fromDate
oCCSRefundCheck.FRectEndDate = toDate
oCCSRefundCheck.FRectChkGubun = chkGubun
oCCSRefundCheck.FRectRefundMin = refundMin
oCCSRefundCheck.FRectRefundMax = refundMax

oCCSRefundCheck.FRectExCheckFinish = exCheckFinish
oCCSRefundCheck.FRectReturnMethodIN = returnmethodIN
oCCSRefundCheck.FRectDategbn = dategbn
oCCSRefundCheck.GetRefundCheckList

%>

<script language='javascript'>

function jsSetTitle(divcd) {
	var asidList, asidElements, ele, chkFound;
	var frm = document.frmAct;

	chkFound = false;
	asidList = "-1";
	asidElements = document.getElementsByName("asid");

	for (var i = 0; i < asidElements.length; i++) {
		ele = asidElements[i];
		if (ele.checked == true) {
			chkFound = true;
			asidList = asidList + "," + ele.value;
		}
	}

	if (chkFound != true) {
		alert("���õ� ������ �����ϴ�.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		if (divcd == "J") {
			// ���޸� ����Ȯ�� �� ȯ��
			frm.mode.value = "ipjumRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else if (divcd == "B") {
			// ���Ա� ����ȯ��
			frm.mode.value = "ipjumDiffRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else if (divcd == "P") {
			// ��ǰ��� ����ȯ��
			frm.mode.value = "prdDiffRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else if (divcd == "CB") {
			// CS���� - ������ ȯ��(��ۺ�)
			frm.mode.value = "csDelivRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else if (divcd == "U") {
			// ��ü���� �� ��ȯ��
			frm.mode.value = "upcheJungsanRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else {
			alert("����.");
			return;
		}
	}
}

function popXL() {
	var popwin = window.open("refund_check_xl_download.asp?page=1&research=on&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>&returnmethod=<%= returnmethod %>&chkGubun=<%= chkGubun %>&dategbn=<%= dategbn %>", "reActAccMonthSummary","width=1000,height=1000 scrollbars=yes resizable=yes");
	popwin.focus();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���� :
			<select class="select" name="divcd">
				<option value=""></option>
				<option>------</option>
				<option value="A003" <% if (divcd = "A003") then %>selected<% end if %> >ȯ��</option>
				<option value="A007" <% if (divcd = "A007") then %>selected<% end if %> >ī�����</option>
				<option value="A008" <% if (divcd = "A008") then %>selected<% end if %> >�ֹ����</option>
				<option>------</option>
				<option value="A004" <% if (divcd = "A004") then %>selected<% end if %> >��ǰ(����)</option>
				<option value="A010" <% if (divcd = "A010") then %>selected<% end if %> >��ǰ(�ٹ�)</option>
			</select>
			&nbsp;
			�Ⱓ :
			<select class="select" name="dategbn">
				<option value="regdate" <%=CHKIIF(dategbn="regdate","selected","")%> >������</option>
				<option value="finishdate" <%=CHKIIF(dategbn="finishdate","selected","")%> >�Ϸ���</option>
			</select>
            <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			<select class="select" name="returnmethod">
				<option></option>
				<option>------</option>
				<!--
				<option value="R100" <% if (returnmethod = "R100") then %>selected<% end if %> >�ſ�ī�� ���</option>
				<option value="R120" <% if (returnmethod = "R120") then %>selected<% end if %> >�ſ�ī�� �κ����</option>
				<option value="R400" <% if (returnmethod = "R400") then %>selected<% end if %> >�޴������� ���</option>
				<option value="R020" <% if (returnmethod = "R020") then %>selected<% end if %> >�ǽð���ü ���</option>
				<option>------</option>
				<option value="R050" <% if (returnmethod = "R050") then %>selected<% end if %> >���������� ���</option>
				<option>------</option>
				-->
				<option value="R007" <% if (returnmethod = "R007") then %>selected<% end if %> >������ ȯ��</option>
				<option value="R910" <% if (returnmethod = "R910") then %>selected<% end if %> >��ġ�� ȯ��</option>
				<option value="R900" <% if (returnmethod = "R900") then %>selected<% end if %> >���ϸ��� ȯ��</option>
				<option value="REXC" <% if (returnmethod = "REXC") then %>selected<% end if %> >������/��ġ��/���ϸ��� �̿� ȯ��</option>
				<!--
				<option>------</option>
				<option value="R000" <% if (returnmethod = "R000") then %>selected<% end if %> >ȯ�� ����</option>
				-->
			</select>
			&nbsp;
			�ֹ���ȣ :
			<input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="14">
			&nbsp;
			ȯ�Ҿ� :
			<input type="text" class="text" name="refundMin" value="<%= refundMin %>" size="10">
			~
			<input type="text" class="text" name="refundMax" value="<%= refundMax %>" size="10">
		</td>
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			ȯ�ҹ�� :
			<input type="checkbox" name="retR007" value="Y" <% if (retR007 = "Y") then %>checked<% end if %> > ������ ȯ��
			<input type="checkbox" name="retR910" value="Y" <% if (retR910 = "Y") then %>checked<% end if %> > ��ġ�� ȯ��
			<input type="checkbox" name="retR900" value="Y" <% if (retR900 = "Y") then %>checked<% end if %> > ���ϸ��� ȯ��
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			���� :
			<input type="radio" name="chkGubun" value="" <% if (chkGubun = "") then %>checked<% end if %> > ��ü
			<input type="radio" name="chkGubun" value="addjung" <% if (chkGubun = "addjung") then %>checked<% end if %> > ��ü�߰�����(��ǰ ��)
			<input type="radio" name="chkGubun" value="err" <% if (chkGubun = "err") then %>checked<% end if %> > �ݾ׺���ġ(ȯ��)
			<input type="radio" name="chkGubun" value="ret" <% if (chkGubun = "ret") then %>checked<% end if %> > ��ǰ(��ü�߰�����)
			<input type="radio" name="chkGubun" value="etc" <% if (chkGubun = "etc") then %>checked<% end if %> > ��ü��Ÿ����
			<input type="radio" name="chkGubun" value="retbea" <% if (chkGubun = "retbea") then %>checked<% end if %> disabled> ��ۺ�(���ɹ�ǰ-����)
			<input type="radio" name="chkGubun" value="retbeaTen" <% if (chkGubun = "retbeaTen") then %>checked<% end if %> disabled> ��ۺ�(���ɹ�ǰ-�ٹ�)
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="exCheckFinish" value="Y" <% if (exCheckFinish = "Y") then %>checked<% end if %> > ����Ϸ� ���� ����(��ġ��ȯ��, ����Ȯ��, �ʰ��Ա�, ��ü����ȯ��, CS����)
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

	<div align="right">
		<input type="button" class="button" value="��ۺ�(CS)" onClick="jsSetTitle('CB');">
		&nbsp;
		<input type="button" class="button" value="�ʰ��Ա� ����" onClick="jsSetTitle('B');">
		<!--
		<input type="button" class="button" value="��ǰ��� ����" onClick="jsSetTitle('P');" disabled>
		-->
		&nbsp;
		<input type="button" class="button" value="��ü����ȯ��" onClick="jsSetTitle('U');">
		&nbsp;
		<input type="button" class="button" value="����Ȯ��" onClick="jsSetTitle('J');">
		&nbsp;
		<input type="button" class="button" value="�����ޱ�" onclick="popXL();">
	</div>

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmList" method="post" onSubmit="return false;">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><% = oCCSRefundCheck.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= oCCSRefundCheck.FTotalpage %></b>
			&nbsp;
			<font color="red">ȯ�Ҿ� �հ�</font> : <b><%= FormatNumber(oCCSRefundCheck.FrefundSUM,0) %> ��</b>
			&nbsp;
			<font color="red">��ü�߰����� �հ�</font> : <b><%= FormatNumber(oCCSRefundCheck.FaddjungSUM,0) %> ��</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"></td>
		<td width="70">ASID</td>
		<td width="100">�ֹ���ȣ</td>
		<td width="120">����</td>
		<td width="80">����01</td>
		<td width="80">����02</td>
		<td width="220">����</td>
		<td width="80">ȯ�ҹ��</td>
		<td width="80">���/��ǰ</td>
		<td width="70"><b>ȯ�Ҿ�</b></td>
		<!--
		<td width="70">��ǰ��ۺ�</td>
		-->
		<td width="70">��ü����</td>
		<td width="100">�������</td>
		<td width="70">�����Ա�</td>
		<td width="80">������</td>
		<td width="80">�Ϸ���</td>
		<td>���</td>
	</tr>
<% if oCCSRefundCheck.FresultCount < 1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for i = 0 to oCCSRefundCheck.FResultCount - 1 %>
	<tr class="a" align="center" bgcolor="FFFFFF">
		<td><input type="checkbox" name="asid" value="<%= oCCSRefundCheck.FItemList(i).Fasid %>"></td>
		<td><a href="javascript:Cscenter_Action_List('<%= oCCSRefundCheck.FItemList(i).FOrderserial %>','','')"><%= oCCSRefundCheck.FItemList(i).Fasid %></a></td>
		<td><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= oCCSRefundCheck.FItemList(i).FOrderserial %>')"><%= oCCSRefundCheck.FItemList(i).FOrderserial %></a></td>
		<td><%= DDotFormat(oCCSRefundCheck.FItemList(i).Fdivcdname,7) %></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).Fgubun01name %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).Fgubun01name,5) %></acronym></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).Fgubun02name %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).Fgubun02name,5) %></acronym></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).Ftitle %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).Ftitle,18) %></acronym></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).FreturnmethodName %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).FreturnmethodName,4) %></acronym></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(oCCSRefundCheck.FItemList(i).FOrgRefundRequire, 0) %></td>
		<td align="right" style="padding-right:5px;"><b><%= FormatNumber(oCCSRefundCheck.FItemList(i).Frefundresult, 0) %></b></td>
		<!--
		<td align="right" style="padding-right:5px;"><%= FormatNumber(oCCSRefundCheck.FItemList(i).Freturndeliverpay, 0) %></td>
		-->
		<td align="right" style="padding-right:5px;"><%= FormatNumber(oCCSRefundCheck.FItemList(i).Fadd_upchejungsandeliverypay, 0) %></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).Fadd_upchejungsancause %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).Fadd_upchejungsancause,5) %></acronym></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(oCCSRefundCheck.FItemList(i).FappPrice, 0) %></td>
		<td><acronym title="<%= oCCSRefundCheck.FItemList(i).Fregdate %>"><%= Left(oCCSRefundCheck.FItemList(i).Fregdate,10) %></acronym></td>
		<td><acronym title="<%= oCCSRefundCheck.FItemList(i).Ffinishdate %>"><%= Left(oCCSRefundCheck.FItemList(i).Ffinishdate,10) %></acronym></td>
		<td>
			<% if (oCCSRefundCheck.FItemList(i).Frefundresult <> oCCSRefundCheck.FItemList(i).FOrgRefundRequire) and (oCCSRefundCheck.FItemList(i).FOrgRefundRequire <> 0) then %>
			<font color="red">ȯ�Ҿ� ����ġ</font>
			<% end if %>
		</td>
	</tr>
	<% next %>
<% end if %>
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
    		<% if oCCSRefundCheck.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= oCCSRefundCheck.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>
    		<% for i = 0 + oCCSRefundCheck.StartScrollPage to oCCSRefundCheck.FScrollCount + oCCSRefundCheck.StartScrollPage - 1 %>
    			<% if i > oCCSRefundCheck.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oCCSRefundCheck.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>

<%
set oCCSRefundCheck = Nothing
%>

<form name="frmAct" method="post" onSubmit="return false;" action="refund_check_list_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="asidList" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
