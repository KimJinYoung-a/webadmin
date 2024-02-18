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
<!-- #include virtual="/lib/classes/cscenter/sp_tenCashCls.asp" -->

<%

dim i, userid, showdel, showtype, currpage, showdetail

userid      = request("userid")
showdel  	= request("showdel")		'�������� ǥ�ÿ���
showtype    = request("showtype")		'���ʽ�(B)����(O)���(S) ���ϸ���
currpage    = request("currpage")
showdetail  = request("showdetail")

if (currpage = "") then currpage = 1
if ((showtype <> "S") and (showtype <> "O") and (showtype <> "B")) then showtype = "A"
if (showdel = "") then showdel = "N"
if (showdetail="") then showtype=""



'==============================================================================
dim oTenCash

set oTenCash = new CTenCash

oTenCash.FRectUserID = userid

if (userid<>"") then
    oTenCash.getUserCurrentTenCash
end if



'==============================================================================
dim oTenCashLog

set oTenCashLog = New CTenCash

oTenCashLog.FPageSize=10
oTenCashLog.FCurrPage= currpage
oTenCashLog.FRectUserid = userid
oTenCashLog.FRectShowDelete = showdel

if (userid<>"")  then
	oTenCashLog.gettenCashLog
end if

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


</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="currpage" value="">
	<input type="hidden" name="showtype" value="<%= showtype %>">
	<input type="hidden" name="showdetail" value="<%= showdetail %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���̵� : <input type="text" class="text" name="userid" value="<%= userid %>">
          	&nbsp;
          	<input type="checkbox" name="showdel" <%= chkIIF(showdel="Y","checked","") %> value="Y">����(���ų����� ��� ���) ǥ��
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
          	<input type="button" class="button" value="�˻�" onclick="document.frm.submit()">
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
		    <strong>�������</strong>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td height=25 width="100">����</td>
    	<td width="200">��ġ���ܾ�</td>
    	<td></td>
    </tr>
<% if (userid <> "") then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25></td>
    	<td><strong><%= FormatNumber(oTenCash.Fcurrentdeposit,0) %> ��</strong></td>
    	<td align="left"><% If oTenCash.Fcurrentdeposit <> "0" Then %>&nbsp;<input type="button" class="button" value="��ġ�� ���������� ȯ��" onClick="returnToBankCash('<%=userid%>')"><% End If %></td>
    </tr>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td></td>
    	<td>-</td>
    	<td>-</td>
    </tr>
<% end if %>
</table>

<p><br><p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td height=25>���̵�</td>
      	<td>����</td>
      	<td>����</td>
      	<td>�ݾ�</td>
      	<td>�ܾ�</td>
      	<td>�����ֹ���ȣ</td>
      	<td>��������</td>
    </tr>
<% if (oTenCashLog.FresultCount > 0) then %>
	<% for i=0 to oTenCashLog.FResultCount - 1 %>
    <tr align="center" <% if (oTenCashLog.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
    	<td height=30><%= userid %></td>
    	<td><%= Left(oTenCashLog.FItemList(i).FRegdate,10) %></td>
    	<td><% if oTenCashLog.FItemList(i).Fdeposit >= 0 then %><font color="blue"><% else %><font color="red"><% end if %><%= oTenCashLog.FItemList(i).Fjukyo %></font></td>
    	<td><% if oTenCashLog.FItemList(i).Fdeposit >= 0 then %><font color="blue"><% else %><font color="red"><% end if %><%= oTenCashLog.FItemList(i).Fdeposit %></font></td>
    	<td><%= FormatNumber(oTenCashLog.FItemList(i).FRemain, 0) %></td>
    	<td><%= oTenCashLog.FItemList(i).Forderserial %></td>
    	<td>
    		<%= oTenCashLog.FItemList(i).Fdeleteyn %>
    		<% if oTenCashLog.FItemList(i).Fdeleteyn = "N" then %>
            	<% if C_ADMIN_AUTH then %>
	    		&nbsp;
	    		<input type="button" class="button" value="[������]����" onClick="SubmitDelete(<%= oTenCashLog.FItemList(i).Fidx %>)">
                <% end if %>
    		<% else %>
				&nbsp;
    			<%= oTenCashLog.FItemList(i).Fdeluserid %>
    		<% end if %>
    	</td>
    </tr>
	<% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<form name="frmpage" method="get" action="">
    	<input type="hidden" name="menupos" value="<%= menupos %>">
    	<input type="hidden" name="userid" value="<%= userid %>">
    	<input type="hidden" name="showtype" value="<%= showtype %>">
    	<input type="hidden" name="currpage" value="<%= currpage %>">
    	<input type="hidden" name="showdetail" value="on">
    	</form>
      	<td colspan="7">
	   	<% if oTenCashLog.HasPreScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= oTenCashLog.StartScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oTenCashLog.StartScrollPage to oTenCashLog.StartScrollPage + oTenCashLog.FScrollCount - 1 %>
			<% if (i > oTenCashLog.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oTenCashLog.FCurrPage) then %>
			<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
			<% else %>
			<a href="javascript:gotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
			<% end if %>
		<% next %>
		<% if oTenCashLog.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= i %>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
      	</td>
    </tr>
<% elseif (userid <> "") then %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="7"> �˻��� ������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>

<form name="frmAction" method="post" action="cs_deposit_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="currpage" value="<%= currpage %>">
<input type="hidden" name="idx" value="">
</form>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
