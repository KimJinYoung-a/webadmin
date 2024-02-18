<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_mileagecls.asp" -->

<%

dim i
dim userid, currpage, showdel
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate
dim writeuser, grpby

yyyy1 	= requestCheckvar(request("yyyy1"),32)
mm1 	= requestCheckvar(request("mm1"),32)
dd1 	= requestCheckvar(request("dd1"),32)
yyyy2 	= requestCheckvar(request("yyyy2"),32)
mm2 	= requestCheckvar(request("mm2"),32)
dd2 	= requestCheckvar(request("dd2"),32)

writeuser 	= requestCheckvar(request("writeuser"),32)
grpby 	= requestCheckvar(request("grpby"),32)

if (currpage = "") then currpage = 1

if (grpby = "") then
	grpby = "writeuser"
end if



if (yyyy1="") then
    nowdate = Left(CStr(dateadd("m",-1,now())),10)
	yyyy1   = Left(nowdate,4)
	mm1     = Mid(nowdate,6,2)
	dd1     = 1

	nowdate = DateSerial(Year(Now()), Month(Now()), 1)
	nowdate = Left(DateAdd("d", -1, nowdate),10)
	yyyy2   = Left(nowdate,4)
	mm2     = Mid(nowdate,6,2)
	dd2     = Mid(nowdate,9,2)
end if


'==============================================================================
dim oCCSCenterMileage

set oCCSCenterMileage = New CCSCenterMileage

oCCSCenterMileage.FPageSize = 300
oCCSCenterMileage.FCurrPage= currpage

oCCSCenterMileage.FRectStartDate= Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
oCCSCenterMileage.FRectEndDate= Left(CStr(DateSerial(yyyy2,mm2,dd2)),10)

oCCSCenterMileage.FRectWriteUser = writeuser

oCCSCenterMileage.FRectGrpBy = grpby

oCCSCenterMileage.getCSMileage


''response.write "aaa"
''response.end

%>
<script language='javascript'>

function gotoPage(page)
{
	document.frm.currpage.value = page;
	document.frm.submit();
}

function changeType(showtype)
{
    document.frm.showdetail.value = "on";
	document.frm.showtype.value = showtype;
	document.frm.submit();
}

function popMileageRequest(userid, orderserial, mileage, jukyo) {
	// �ʼ� : ���̵�
	// �ɼ� : �ֹ���ȣ, ���ϸ���, ���䳻��

	if (userid == "") {
		alert("���̵� �����ϴ�.");
		return;
	}

    var popwin = window.open('/cscenter/mileage/pop_mileage_request.asp?userid=' + userid + '&orderserial=' + orderserial + '&mileage=' + mileage + '&jukyo=' + jukyo,'popMileageRequest','width=660,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}



function popYearExpireMileList(yyyymmdd,userid){
    var popwin = window.open('popAdminExpireMileSummary.asp?yyyymmdd=' + yyyymmdd + '&userid=' + userid,'popAdminExpireMileSummary','width=660,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function returnToBankCash(userid)
{
    var popwin = window.open('cs_popReturnToBankCash.asp?userid=' + userid,'cs_popReturnToBankCash','width=400,height=300');
    popwin.focus();
}

function SubmitDelete(idx) {
	var frm = document.frmAction;

	if (confirm("��ġ�� ������ �����Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value = "delete";
	frm.idx.value = idx;
	frm.submit();
}

function jsSubmit() {
	var frm = document.frm;
	if ((frm.grpby[2].checked === true) && (frm.writeuser.value === "")) {
		alert("���̵� �Է��� ��쿡�� �հ豸�� ������ ������ �� �ֽ��ϴ�.");
		return;
	}

	frm.submit();
}

function searchListByReguser(writeuser) {
	var frm = document.frm;
	frm.writeuser.value=writeuser;
	frm.grpby[2].checked=true;
	frm.currpage.value=1;
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="currpage" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			����ھ��̵� : <input type="text" class="text" name="writeuser" value="<%= writeuser %>">
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
          	<input type="button" class="button" value="�˻�" onclick="jsSubmit()">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�հ豸�� :
			<input type="radio" name="grpby" value="writeuser" <%= CHKIIF(grpby="writeuser", "checked", "") %> > �����
			<input type="radio" name="grpby" value="title" <%= CHKIIF(grpby="title", "checked", "") %> > ����
			<input type="radio" name="grpby" value="none" <%= CHKIIF(grpby="none", "checked", "") %> > ����
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td height=25 width="80">�ֹ���ȣ</td>
		<td width="100">�����̵�</td>
      	<td width="280">����</td>
      	<td width="150">�����</td>
      	<td width="100">ó����</td>
      	<td width="100">���ϸ���</td>
      	<td width="150">�����</td>
      	<td width="150">ó����</td>
      	<td>���</td>
    </tr>
<% if (oCCSCenterMileage.FresultCount > 0) then %>
	<% for i=0 to oCCSCenterMileage.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=30><%= oCCSCenterMileage.FItemList(i).Forderserial %></td>
		<td><%= oCCSCenterMileage.FItemList(i).Fuserid %></td>
		<td><%= oCCSCenterMileage.FItemList(i).Ftitle %></td>
		<td align="left" style="padding: 5px;">
			<span onclick="searchListByReguser('<%= oCCSCenterMileage.FItemList(i).Fwriteuser %>')" style="cursor:pointer;" title="��ϳ��� ����">
			<% if (oCCSCenterMileage.FItemList(i).Fusername <> "") then %>
			<%= oCCSCenterMileage.FItemList(i).Fusername %>(<%= oCCSCenterMileage.FItemList(i).Fwriteuser %>)
			<% elseif (oCCSCenterMileage.FItemList(i).Fwriteuser <> "") then %>
			<%= oCCSCenterMileage.FItemList(i).Fwriteuser %>
			<% end if %>
			</span>
		</td>
		<td><%= oCCSCenterMileage.FItemList(i).Ffinishuser %></td>
		<td align="right" style="padding: 5px;"><%= FormatNumber(oCCSCenterMileage.FItemList(i).Frefundresult,0) %></td>
		<td><%= oCCSCenterMileage.FItemList(i).Fregdate %></td>
		<td><%= oCCSCenterMileage.FItemList(i).Ffinishdate %></td>
		<td><%= oCCSCenterMileage.FItemList(i).Fcontents_jupsu %></td>
    </tr>
	<% next %>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="9"> �˻��� ������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
