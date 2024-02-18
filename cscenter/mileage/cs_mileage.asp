<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ���ϸ�������
' History : �̻� ����
'			2023.07.21 �ѿ�� ����(���ϸ����Ҹ� �⺰->���� �� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mileage/sp_pointcls.asp" -->
<!-- #include virtual="/lib/classes/mileage/sp_mileage_logcls.asp" -->
<%
dim i, userid, showdelete, showtype, currpage, showdetail, research, myMileage, myOffMileage, myMileageLog, oExpireMile
dim oBeforeSixMonth, beforeSixMonthSUM, currentdate, expireDate
	userid      = requestCheckVar(trim(request("userid")),32)
	showdelete  = requestCheckVar(trim(request("showdelete")),1)	'�������� ǥ�ÿ���
	showtype    = requestCheckVar(trim(request("showtype")),1)		'���ʽ�(B)����(O)���(S) ���ϸ���
	currpage    = requestCheckVar(getNumeric(trim(request("currpage"))),10)
	showdetail  = requestCheckVar(trim(request("showdetail")),2)
	research  = requestCheckVar(trim(request("research")),2)

if (research = "") then
	''showdelete = "Y"
end if

if (currpage = "") then currpage = 1
if ((showtype <> "S") and (showtype <> "O") and (showtype <> "B") and (showtype <> "X")) then showtype = "A"
if (showdelete = "") then showdelete = "N"
if (showdetail="") then showtype=""

currentdate=date()

' �̹��޸���
expireDate = dateadd("d",-1,dateserial(year(dateadd("m",+1,currentdate)),month(dateadd("m",+1,currentdate)),"01"))

set myMileage = new TenPoint
myMileage.FRectUserID = userid
if (userid <> "") then
    myMileage.getTotalMileage
end if

set myOffMileage = new TenPoint
myOffMileage.FGubun = "my10x10"
myOffMileage.FRectUserID = userid
if (userid <> "") then
    myOffMileage.getOffShopMileagePop
end if

set myMileageLog = New CMileageLog
myMileageLog.FPageSize = 100
myMileageLog.FCurrPage = Cint(currpage)
myMileageLog.FRectUserid = userid
myMileageLog.FRectMileageLogType = showtype
myMileageLog.FRectShowDelete = showdelete

if ((userid <> "") and (showtype <> "") and (showdetail<>"")) then
	if (showtype = "A") then
		myMileageLog.getMileageLogAll
		'myMileageLog.getMileageLog
	else
		myMileageLog.getMileageLog
	end if
end if

' ���Ό��  ���ϸ���
set oExpireMile = new CMileageLog
	oExpireMile.FRectUserid = userid
	oExpireMile.FRectExpireDate = expireDate
	if (userid<>"") then
		'oExpireMile.getNextExpireMileageSum
		oExpireMile.getNextExpireMileageMonthlySum
	end if

set oBeforeSixMonth = new CMileageLog
oBeforeSixMonth.FRectUserid = userid

beforeSixMonthSUM = 0
if (userid<>"") then
    oBeforeSixMonth.GetRealSumBuyMileageBeforeSixMonth()
	beforeSixMonthSUM = oBeforeSixMonth.FOneItem.FbeforesixmonthSUM
end if

%>
<script type='text/javascript'>

function gotoPage(page){
	document.frmpage.currpage.value = page;
	document.frmpage.submit();
}

function changeType(showtype){
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

    var popwin = window.open('/cscenter/mileage/pop_mileage_request.asp?userid=' + userid + '&orderserial=' + orderserial + '&mileage=' + mileage + '&jukyo=' + jukyo,'popMileageRequest','width=1400,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popOffMileageList(userid) {
	if (userid == "") {
		alert("���̵� �����ϴ�.");
		return;
	}

    var popwin = window.open('/admin/offshop/offmileagelist.asp?menupos=651&userid=' + userid,'popOffMileageList','width=1500,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popYearExpireMileList(yyyymmdd,userid){
    var popwin = window.open('/cscenter/mileage/popAdminExpireMileSummary.asp?yyyymmdd=' + yyyymmdd + '&userid=' + userid,'popAdminExpireMileSummary','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}
function popmonthlyExpireMileList(yyyymmdd,userid){
    var popwin = window.open('/cscenter/mileage/popAdminExpireMileMonthlySummary.asp?yyyymmdd=' + yyyymmdd + '&userid=' + userid+'&menupos=<%=menupos%>','popAdminExpireMileSummary','width=1400,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if C_ADMIN_AUTH then %>
	function SubmitFormDelForce(idx) {
		var frm = document.frmAct;

		if (frm.userid.value == "") {
			alert("����.");
			return;
		}

		if (idx == "") {
			alert("����.");
			return;
		}

		if (confirm("[������]����!!!!\n\n�����޿� �ο��� ���ϸ����� �����ϸ� �ȵ˴ϴ�.\n\n���� �Ͻðڽ��ϱ�?") == true) {
			frm.mode.value = "delForce";
			frm.idx.value = idx;
			frm.submit();
		}
	}

	function jsReCalcSum() {
		var frm = document.frmAct;

		if (frm.userid.value == "") {
			alert("���̵� �����ϴ�.");
			return;
		}

		if (confirm("[������]���ϸ����� �����մϴ�.\n\n�����Ͻðڽ��ϱ�?") == true) {
			frm.mode.value = "recalcmile";
			frm.submit();
		}
	}
<% end if %>

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="showtype" value="<%= showtype %>">
<input type="hidden" name="showdetail" value="<%= showdetail %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���̵� : <input type="text" class="text" name="userid" value="<%= userid %>" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;
		<input type="checkbox" name="showdelete" <%= chkIIF(showdelete="Y","checked","") %> value="Y">����(���ų����� ��� ���) ǥ��
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button" value="�˻�" onclick="document.frm.submit()">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		<% if C_ADMIN_AUTH then %>
			<input type="button" class="button" value="���ϸ��� ����[������]" onClick="jsReCalcSum()">
		<% end if %>
	</td>
</tr>
</table>
</form>
<br>

<form name="frmWrite" action="userMileage_Process.asp" onsubmit="return checkForm();" style="margin:0px;">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td>
		<input type="button" class="button" value="������û" onclick="popMileageRequest('<%= userid %>', '', 0, '');">
		������û�� �Ͻø�, CSó������Ʈ�� ��ϵǸ�, ������ ���ΰ� �Բ� �����˴ϴ�.
	</td>
</tr>
<tr>
	<td>
		<span id="divWrite" style="float:left; display:none">
			<input type="hidden" name="mode" value="INS">
			<input type="hidden" name="userID" value="<%=userID%>">
			�ֹ���ȣ :
			<input type="text" name="orderSerial" size="11" maxlength="11" onkeydown="onlyNumber(this,event);" class="text">
			&nbsp;
			������ :
			<input type="text" name="savePoint" size="5" maxlength="5" style="text-align:right;" onkeydown="onlyNumber(this,event);" class="text">
			&nbsp;
			�������� :
			<select class="select" name="etcTitle">
				<option value='' selected>��Ͼ���</option>
				<option value='�Ա�����'>�Ա�����</option>
				<option value='��ǰ����'>��ǰ����</option>
				<option value='�������'>�������</option>
				<option value='CS����'>CS����</option>
				<option value='��ǰ���ȯ��'>��ǰ���ȯ��</option>
				<option value='��Ÿ'>��Ÿ</option>
			</select>
			<input type="submit" class="button" value="���">
		</span>
	</td>
</tr>
</table>
</form>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
		<strong>�������</strong>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="35">
	<td height=25>����</td>
	<td>���縶�ϸ���</td>
	<td>���ʽ� ���ϸ���</td>
	<td>�������� ���ϸ���<br>(�¶���+��ī����)</td>
	<td>������������ ���ϸ���<br>(�¶���+��ī����)</td>
	<td>����� ���ϸ���</td>
	<td>�Ҹ�� ���ϸ���</td>
</tr>
<% if (userid <> "") then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>�¶���</td>
    	<td><strong><%=FormatNumber(myMileage.FTotalMileage,0) %></strong></td>
    	<td><%=FormatNumber(myMileage.FBonusMileage,0) %></td>
    	<td><%=FormatNumber(myMileage.FTotJumunmileage + myMileage.FAcademymileage,0) %></td>
      	<td><%=FormatNumber(myMileage.Fmichulmile + myMileage.FmichulmileACA,0) %></td>
      	<td><%=FormatNumber(myMileage.FSpendMileage*-1,0) %></font></td>
      	<td><%=FormatNumber(myMileage.FrealExpiredMileage*-1,0) %></font></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>��������</td>
    	<td><a href="javascript:popOffMileageList('<%= userID %>')"><strong><%=FormatNumber(myOffMileage.FOffShopMileage,0) %></strong></a></td>
    	<td colspan=5></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>�Ҹ� ��� ���ϸ���</td>
    	<td>
			<!--<a href="javascript:popYearExpireMileList('<%'= oExpireMile.FRectExpireDate %>','<%'= userid %>');">-->
			<a href="#" onclick="popmonthlyExpireMileList('<%= oExpireMile.FRectExpireDate %>','<%= userid %>'); return false;">
			<%= FormatNumber(oExpireMile.FOneItem.getMayExpireTotal,0) %></a>
		</td>
    	<td colspan=5 align=left>
			&nbsp;&nbsp;
			<!--<a href="javascript:popYearExpireMileList('<%'= oExpireMile.FRectExpireDate %>','<%'= userid %>');">-->
			<a href="#" onclick="popmonthlyExpireMileList('<%= oExpireMile.FRectExpireDate %>','<%= userid %>'); return false;">
			* �Ҹ����� : <%= expireDate %></a>
		</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>�ֱ�(6�����̳�) ���������հ�</td>
    	<td><%= FormatNumber(myMileage.FRecentJumunMileage,0) %></td>
    	<td colspan=5 align=left></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>6�������� ���������հ�</td>
    	<td>
			<%= FormatNumber(myMileage.FOldJumunmileage,0) %>
			<% if (beforeSixMonthSUM > 0) and (beforeSixMonthSUM <> myMileage.FOldJumunmileage) then %>
			<br /><font color="red">(<%= FormatNumber(beforeSixMonthSUM,0) %>)</font>
			<% end if %>
		</td>
    	<td colspan=5 align=left> &nbsp;&nbsp;* ���� ���ϸ������� 6������ ���� ������ ǥ�õ��� �ʽ��ϴ�.</td>
    </tr>
    	<% if (myMileage.FAcademyMileage>0) then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>��ī���� �ֹ�����</td>
    	<td><%= FormatNumber(myMileage.FAcademyMileage,0) %></td>
    	<td colspan=5 align=left> &nbsp;&nbsp;* ������������ ���ϸ����� <font color="red">��ǰ����</font> �����˴ϴ�.</td>
    </tr>
		<% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25>��ī���� �ֹ�����</td>
    	<td>����</td>
    	<td colspan=5 align=left> &nbsp;&nbsp;* ������������ ���ϸ����� <font color="red">��ǰ����</font> �����˴ϴ�.</td>
    </tr>
    	<% end if %>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�¶���</td>
    	<td>-</td>
    	<td>-</td>
      	<td>-</td>
		<td>-</td>
      	<td>-</td>
      	<td>-</td>
    </tr>
<% end if %>
</table>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
		<strong>�󼼳��� : </strong>
		<%=chkIIF(showtype="A","<strong>","")%><a href="javascript:changeType('A')">��ü����</a><%=chkIIF(showtype="A","</strong>","")%>
		|
		<%=chkIIF(showtype="B","<strong>","")%><a href="javascript:changeType('B')">���ʽ� ���ϸ���</a><%=chkIIF(showtype="B","</strong>","")%>
		|
		<%=chkIIF(showtype="O","<strong>","")%><a href="javascript:changeType('O')">���� ���ϸ���</a><%=chkIIF(showtype="O","</strong>","")%>
		|
		<%=chkIIF(showtype="S","<strong>","")%><a href="javascript:changeType('S')">��� ���ϸ���</a><%=chkIIF(showtype="S","</strong>","")%>
		|
		<%=chkIIF(showtype="X","<strong>","")%><a href="javascript:changeType('X')">�Ҹ� ���ϸ���</a><%=chkIIF(showtype="X","</strong>","")%>
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		<% if (showdetail="on") then %>
			�˻���� : <b>�� <%= myMileageLog.FTotalCount %> ��</b> ������ : <b><%= currpage %> / <%= myMileageLog.FTotalPage %></b>
		<% end if %>
	</td>
</tr>
<% if (showdetail="on") then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height=25>���̵�</td>
		<td>��������</td>
		<td>����</td>
		<td>���䳻��</td>
		<td>���ϸ���</td>
		<td>�ܾ�</td>
		<td>�����ֹ���ȣ</td>
		<td>��������</td>
		<td>���</td>
	</tr>
	<% if (myMileageLog.FresultCount > 0) then %>
		<% for i=0 to myMileageLog.FResultCount - 1 %>
		<tr align="center" <% if (myMileageLog.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
			<td height=25><%= userid %></td>
			<td><%= Left(myMileageLog.FItemList(i).FRegdate,10) %></td>
			<td><% if myMileageLog.FItemList(i).Fmileage >= 0 then %><font color="blue"><% else %><font color="red"><% end if %><%= myMileageLog.FItemList(i).Fstatusflagstring %></font></td>
			<td><%= myMileageLog.FItemList(i).Fjukyo %></td>
			<td align="right"><% if myMileageLog.FItemList(i).Fmileage >= 0 then %><font color="blue"><% else %><font color="red"><% end if %><%= FormatNumber(myMileageLog.FItemList(i).Fmileage, 0) %></font>&nbsp;&nbsp;</td>
			<td align="right">
				<%
				if (showtype = "A") then
					response.write FormatNumber(myMileageLog.FItemList(i).Fremain, 0)
				else
					response.write "--"
				end if
				%>
				&nbsp;&nbsp;
			</td>
			<td><%= myMileageLog.FItemList(i).Forderserial %></td>
			<td><%= myMileageLog.FItemList(i).Fdeleteyn %></td>
			<td>
				<% if C_ADMIN_AUTH and (myMileageLog.FItemList(i).Fstatusflag = "B") and (myMileageLog.FItemList(i).Fid <> "") and (myMileageLog.FItemList(i).Fid > "0") and (myMileageLog.FItemList(i).Fdeleteyn <> "Y") then %>
					<input type="button" class="button" value="����[������]" onClick="SubmitFormDelForce(<%= myMileageLog.FItemList(i).Fid %>)">
				<% end if %>
			</td>
		</tr>
		<% next %>
		<tr align="center" bgcolor="#FFFFFF">
			<form name="frmpage" method="get" action="" style="margin:0px;">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="userid" value="<%= userid %>">
			<input type="hidden" name="showtype" value="<%= showtype %>">
			<input type="hidden" name="showdelete" value="<%= showdelete %>">
			<input type="hidden" name="currpage" value="<%= currpage %>">
			<input type="hidden" name="showdetail" value="on">
			</form>
			<td colspan="15">
			<% if myMileageLog.HasPreScroll then %>
				<span class="list_link"><a href="javascript:gotoPage(<%= myMileageLog.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + myMileageLog.StartScrollPage to myMileageLog.StartScrollPage + myMileageLog.FScrollCount - 1 %>
				<% if (i > myMileageLog.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(myMileageLog.FCurrPage) then %>
				<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
				<% else %>
				<a href="javascript:gotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
				<% end if %>
			<% next %>
			<% if myMileageLog.HasNextScroll then %>
				<span class="list_link"><a href="javascript:gotoPage(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
			</td>
		</tr>

	<% else %>
		<tr align="center" bgcolor="#FFFFFF">
			<td colspan="15"> �˻��� ������ �����ϴ�.</td>
		</tr>
	<% end if %>
<% end if %>
</table>

<form name="frmAct" method="post" action="domodifymileage.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="idx" value="">
</form>

<%
set myMileageLog = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
