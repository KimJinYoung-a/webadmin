<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : OkCashbag����
' History : ������ ����
'			2023.03.22 �ѿ�� ����(���� ���� ���̵� ���� �ִºκ� ���� ���� ������ �ڵ�ȭ. �ҽ� ǥ���ڵ�� ����.)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/othermall/okcashbagCls.asp"-->
<%
dim sSdate,sEdate, userid, orderserial, vRdSite, OrderType, sPageSize, SearchDateType, CurrPage
dim oCash,intLp
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)
	userid 		= requestCheckVar(Request("uId"),32)
	orderserial	= requestCheckVar(Request("oSn"),12)
	vRdSite		= requestCheckVar(Request("rdsite"),10)
	OrderType = requestCheckVar(Request("otp"),2)
	SearchDateType = requestCheckVar(request("dType"),2)
	CurrPage = requestCheckVar(request("pg"),3)

If vRdSite = "" Then
	vRdSite = "okcashbag"
End If

IF sSdate ="" Then
	sSdate= DateSerial(Year(now()),Month(now()),1)
End IF

IF OrderType="" Then OrderType="N"
sPageSize = 100
IF SearchDateType="" THEN SearchDateType="od"

IF CurrPage="" THEN CurrPage =1

Set oCash = New CashbagCls
	oCash.FCurrPage		= CurrPage
	oCash.FPageSize		= sPageSize
	oCash.FStartDate 	= sSdate
	oCash.FEndDate 		= sEdate
	oCash.Fuserid	 	= userid
	oCash.Forderserial 	= orderserial
	oCash.FOrderType 	= OrderType
	oCash.FSearchType	= SearchDateType
	oCash.FRdSite		= vRdSite

	IF OrderType="N" Then 		'//�����
		oCash.getNormalOrder()
	ELSEIF OrderType ="C" Then	'//��Ұ�
		oCash.getCancelOrder()
	ELSEIF OrderType="UN" or OrderType ="UC" Then '// ��� �� ���� (����,���)
		oCash.getUpdatedOrder()
	END IF

%>

<script type="text/javascript">

function jsChkAll(blnChk){
		var frm, blnChk;
		frm = document.rfrm;

		for (var i=0;i<frm.elements.length;i++){
			//check optioon
			var e = frm.elements[i];

			//check itemEA
			if ((e.type=="checkbox")) {
				e.checked = blnChk ;
				AnCheckClick(e);
		}
	}
}

function downloadexcel(){
	document.sfrm.target = "view";
	document.sfrm.action = "/admin/etc/okcashbag/okcashbagReport_down.asp";
	document.sfrm.submit();
	document.sfrm.target = "";
	document.sfrm.action = "";
}

function NextPage(v){
	document.sfrm.target = "";
	document.sfrm.action = "";
	document.sfrm.pg.value=v;
	document.sfrm.submit();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName+'&FN=sfrm','pCal','width=250, height=200');
	winCal.focus();
}

</script>

<!-- �˻� ���� -->
<form name="sfrm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="pg" value=1>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<select name="dType">
				<option value="od" <% IF SearchDateType="od" Then response.write "selected"%>>�ֹ��� ����</option>
				<option value="ov" <% IF SearchDateType="ov" Then response.write "selected"%>>����� ����</option>
				<!--<option value="ud" <% IF SearchDateType="ud" Then response.write "selected"%>>������ ����</option>-->
			</select>
			<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
			~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">
			&nbsp;
			* ���̵� : <input type="text" size="10" maxlength="32" name="uId" value="<%=userid%>">
			&nbsp;
			* �ֹ���ȣ : <input type="text" size="12" maxlength="12" name="oSn" value="<%=orderserial%>">
			&nbsp;
			* ���޾�ü : 
			<select name="rdsite">
				<option value="okcashbag" <%=ChkIIF(vRdSite="okcashbag","selected","")%>>okcashbag</option>
				<option value="pickle" <%=ChkIIF(vRdSite="pickle","selected","")%>>pickle</option>
			</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="NextPage('');">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<acronym title="�����"><input type="radio" name="otp" value="N" <% IF OrderType="N" Then response.write "checked" %> onClick="document.sfrm.submit();">�����</acronym>
			<acronym title="����� ��³���"><input type="radio" name="otp" value="UN" <% IF OrderType="UN" Then response.write "checked" %> onClick="document.sfrm.submit();">����� ��³���</acronym>
			<acronym title="����� ����� ��ҳ���"><input type="radio" name="otp" value="C" <% IF OrderType="C" Then response.write "checked" %> onClick="document.sfrm.submit();">��Ұ�</acronym>
			<acronym title="��Ұ� ��³���"><input type="radio" name="otp" value="UC" <% IF OrderType="UC" Then response.write "checked" %> onClick="document.sfrm.submit();">��Ұ� ��³���</acronym>
		</td>
	</tr>
</table>
<!-- �˻� �� -->

<Br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">
		<%
		' ������ �̰ų� ���߿�� �̰ų� ������Ʈ �ϰ��
		If C_ADMIN_AUTH or C_SYSTEM_Part or C_partnership_part Then
		%>
			<input type="button" onclick="downloadexcel();" value="�����ٿ�ε�" class="button">
		<% else %>
			�ٿ���Ѿ���
		<% End If %>
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<form name="rfrm" method="post" action="" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�˻���� : <b><%= oCash.FTotalCount %></b>
		&nbsp;
		������ : <b><%= CurrPage %>/ <%= oCash.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll(this.checked);"></td>
	<td align="center" width="100" >�ֹ���ȣ</td>
	<td align="center" width="80">��ٱ��Ϲ�ȣ</td>
	<td align="center">�Ѱ����ݾ�</td>
	<td align="center">�ֹ�����</td>
	<td align="center">�������</td>
	<td align="center">�ֹ���</td>
	<td align="center">ĳ�����ȣ</td>
	<td align="center">��������Ʈ</td>
</tr>
<% if oCash.FresultCount>0 then %>
<% for IntLp=0 to oCash.FresultCount-1 %>

<tr align="center" bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="chkb" onClick="AnCheckClick(this);" value="<%= oCash.FItemList(IntLp).Fidx %>"></td>
	<td align="center"><%= oCash.FItemList(IntLp).FOrderSerial %></td>
	<td align="center"><%= oCash.FItemList(IntLp).FShoppingBagNo %></td>
	<td align="center"><%= FormatNumber(oCash.FItemList(IntLp).FPointCash,0) %></td>
	<td align="center">
		<% if not(isnull(oCash.FItemList(IntLp).FRegdate)) then %>
			<%= DateValue(oCash.FItemList(IntLp).FRegdate) %>
		<% end if %>
	</td>
	<td align="center"><% if DateValue(oCash.FItemList(IntLp).FBeadaldate)="1900-01-01" then Response.Write "�̹��": Else Response.Write DateValue(oCash.FItemList(IntLp).FBeadaldate): End if %></td>
	<td align="center"><%= oCash.FItemList(IntLp).FBuyName %></td>
	<td align="center">0000-****-****-0000</td>
	<td align="center"><%= FormatNumber(oCash.FItemList(IntLp).FPoint,0) %></td>
</tr>   
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="9" align="center">
		<% if oCash.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCash.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for intLp=0 + oCash.StartScrollPage to oCash.FScrollCount + oCash.StartScrollPage - 1 %>
			<% if intLp>oCash.FTotalpage then Exit for %>
			<% if CStr(CurrPage)=CStr(intLp) then %>
			<font color="red">[<%= intLp %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= intLp %>')">[<%= intLp %>]</a>
			<% end if %>
		<% next %>

		<% if oCash.HasNextScroll then %>
			<a href="javascript:NextPage('<%= intLp %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="16" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</form>
<!-- ����Ʈ ���� -->

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set oCash = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
