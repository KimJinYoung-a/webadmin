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
<!-- #include virtual="/lib/classes/cscenter/sp_tenGiftCardCls.asp" -->

<%

dim i, userid, showdelete, currpage

userid      = request("userid")
showdelete  = request("showdelete")		'�������� ǥ�ÿ���
currpage    = request("currpage")

if (currpage = "") then currpage = 1
if (showdelete = "") then showdelete = "N"



'==============================================================================
dim oTenGiftCard

set oTenGiftCard = new CTenGiftCard

oTenGiftCard.FRectUserID = userid

if (userid<>"") then
    oTenGiftCard.getUserCurrentTenGiftCard
end if



'==============================================================================
dim oTenGiftCardLog

set oTenGiftCardLog = New CTenGiftCard

oTenGiftCardLog.FPageSize=20
oTenGiftCardLog.FCurrPage= currpage
oTenGiftCardLog.FRectUserid = userid

if (userid<>"")  then
	oTenGiftCardLog.gettenGiftCardLog
end if



dim spendPercentage

if (oTenGiftCard.FgainCash <> 0) then
	spendPercentage = 100*oTenGiftCard.FspendCash/oTenGiftCard.FgainCash
else
	spendPercentage = 0
end if


%>
<script language='javascript'>

function gotoPage(page)
{
	document.frmpage.currpage.value = page;
	document.frmpage.submit();
}

function refundByBank(userid)
{
    var popwin = window.open('cs_popRefundByBank.asp?userid=' + userid,'cs_popRefundByBank','width=400,height=300');
    popwin.focus();
}

/*
function SubmitDelete(idx) {
	var frm = document.frmAction;

	if (confirm("��ġ�� ������ �����Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value = "delete";
	frm.idx.value = idx;
	frm.submit();
}
*/

</script>

<style type "text/css">
<!--
/* @group Blink */
.blink {
	-webkit-animation: blink .75s linear infinite;
	-moz-animation: blink .75s linear infinite;
	-ms-animation: blink .75s linear infinite;
	-o-animation: blink .75s linear infinite;
	 animation: blink .75s linear infinite;
}
@-webkit-keyframes blink {
	0% { opacity: 1; }
	50% { opacity: 1; }
	50.01% { opacity: 0; }
	100% { opacity: 0; }
}
@-moz-keyframes blink {
	0% { opacity: 1; }
	50% { opacity: 1; }
	50.01% { opacity: 0; }
	100% { opacity: 0; }
}
@-ms-keyframes blink {
	0% { opacity: 1; }
	50% { opacity: 1; }
	50.01% { opacity: 0; }
	100% { opacity: 0; }
}
@-o-keyframes blink {
	0% { opacity: 1; }
	50% { opacity: 1; }
	50.01% { opacity: 0; }
	100% { opacity: 0; }
}
@keyframes blink {
	0% { opacity: 1; }
	50% { opacity: 1; }
	50.01% { opacity: 0; }
	100% { opacity: 0; }
}
/* @end */
-->
</style>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���̵� : <input type="text" class="text" name="userid" value="<%= userid %>">
          	&nbsp;
          	<!--
          	<input type="checkbox" name="showdelete" <%= chkIIF(showdelete="Y","checked","") %> value="Y">����(���ų����� ��� ���) ǥ��
          	-->
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
		<td colspan="7">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
		    <strong>�������</strong>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td height=25 width="100">����</td>
    	<td width="150">Giftī�� �ܾ�</td>
    	<td width="150">����Ѿ�</td>
    	<td width="150">��ǰ�����Ѿ�</td>
    	<td width="150">Giftī�������</td>
    	<td width="150">�� ȯ���Ѿ�</td>
    	<td></td>
    </tr>
<% if (userid <> "") then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25></td>
    	<td><strong><%= FormatNumber(oTenGiftCard.FcurrentCash,0) %> ��</strong></td>
    	<td><strong><%= FormatNumber(oTenGiftCard.FgainCash,0) %> ��</strong></td>
    	<td><strong><%= FormatNumber(oTenGiftCard.FspendCash,0) %> ��</strong></td>
    	<td><strong><%= FormatNumber((spendPercentage),0) %> %</strong></td>
    	<td><strong><%= FormatNumber(oTenGiftCard.FrefundCash,0) %> ��</strong></td>
    	<td align="left">
    		<% If oTenGiftCard.FcurrentCash <> "0"  Then %>
    			<% if (CLng(FormatNumber((100*oTenGiftCard.FspendCash/oTenGiftCard.FgainCash),0)) >= 60) or (userid="danbi2612") or (userid="setjddms") or (userid="dadareda") or (userid="eiddr0705") then %>
    				&nbsp;<input type="button" class="button" value="������ ȯ��" onClick="refundByBank('<%=userid%>')">
    			<% end if %>
    		<% End If %>
    	</td>
    </tr>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td></td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    </tr>
<% end if %>
</table>
<br><font color=red>* Giftī�������( = ��ǰ�����Ѿ�/����Ѿ�) �� 60% �̻��� ��� �ܾ��� ȯ���� �����մϴ�.</font>
<% if (userid = "woodpy35") then %>
<br><font color=red class="tab blink">* woodp*** ������ ����Ƽ�� ���� �� ���������� ����Ͻ� ���̽ʴϴ�.(Ȯ�οϷ�, 2018-02-09, skyer9)</font>
<% end if %>

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
<% if (oTenGiftCardLog.FresultCount > 0) then %>
	<% for i=0 to oTenGiftCardLog.FResultCount - 1 %>
    <tr align="center" <% if (oTenGiftCardLog.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
    	<td height=30><%= userid %></td>
    	<td><%= Left(oTenGiftCardLog.FItemList(i).FRegdate,10) %></td>
    	<td><% if oTenGiftCardLog.FItemList(i).FuseCash >= 0 then %><font color="blue"><% else %><font color="red"><% end if %><%= oTenGiftCardLog.FItemList(i).Fjukyo %></font></td>
    	<td><% if oTenGiftCardLog.FItemList(i).FuseCash >= 0 then %><font color="blue"><% else %><font color="red"><% end if %><%= oTenGiftCardLog.FItemList(i).FuseCash %></font></td>
    	<td><%= FormatNumber(oTenGiftCardLog.FItemList(i).FRemain, 0) %></td>
    	<td><%= oTenGiftCardLog.FItemList(i).Forderserial %></td>
    	<td>
    		<%= oTenGiftCardLog.FItemList(i).Fdeleteyn %>
    		<% if oTenGiftCardLog.FItemList(i).Fdeleteyn = "N" then %>
	    		&nbsp;
	    		<!--
	    		<input type="button" class="button" value="����" onClick="SubmitDelete(<%= oTenGiftCardLog.FItemList(i).Fidx %>)">
	    		-->
    		<% else %>
    			<%= oTenGiftCardLog.FItemList(i).Fdeluserid %>
    		<% end if %>
    	</td>
    </tr>
	<% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<form name="frmpage" method="get" action="">
    	<input type="hidden" name="menupos" value="<%= menupos %>">
    	<input type="hidden" name="userid" value="<%= userid %>">
    	<input type="hidden" name="currpage" value="<%= currpage %>">
    	</form>
      	<td colspan="7">
	   	<% if oTenGiftCardLog.HasPreScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= oTenGiftCardLog.StartScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oTenGiftCardLog.StartScrollPage to oTenGiftCardLog.StartScrollPage + oTenGiftCardLog.FScrollCount - 1 %>
			<% if (i > oTenGiftCardLog.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oTenGiftCardLog.FCurrPage) then %>
			<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
			<% else %>
			<a href="javascript:gotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
			<% end if %>
		<% next %>
		<% if oTenGiftCardLog.HasNextScroll then %>
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
