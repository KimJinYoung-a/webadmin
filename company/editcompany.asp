<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim onepartner,i,page
page = request("page")
if page="" then page=1
set onepartner = new CPartnerUser
onepartner.FCurrpage = page
onepartner.GetOnePartner session("ssBctId")
%>
<script language="javascript">
function useredit(frm){
	for (var i=0;i<frm.elements.length;i++){
	  var e = frm.elements[i];

	  if ((e.name=="txpassword")) {
		if (e.value.length<1){
			alert('�ʼ� �Է� �����Դϴ�.');
			e.focus();
			return;
		}
	  }
	}
	
	if (frm.txnewpassword1.value!=frm.txnewpassword2.value){
		alert('�� ��й�ȣ�� ��ġ���� �ʽ��ϴ�.');
		frm.txnewpassword2.focus();
		return;
	}
	
	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.submit();
	}
}
</script>
<table width="610" border="0" class="a">
	<form name="frmedit" method="post" action="doeditcompany.asp">
	<tr>
		<td width="120">���̵� :</td>
		<td><%= onepartner.FPartnerList(0).FID %></td>
	</tr>
	<tr>
		<td >��ü�� :</td>
		<td><input type="text" name="txcompanyname" value="<%= onepartner.FPartnerList(0).FCompany_name %>"></td>
	</tr>
	<tr>
		<td >��й�ȣ :</td>
		<td><input type="password" name="txpassword" value="" size="12" maxlength="16"></td>
	</tr>
	<tr>
		<td >�ּ� :</td>
		<td>
			<input type="text" name="txaddress1" value="<%= onepartner.FPartnerList(0).FAddress %>">(���� ������)<br>
			<input type="text" name="txaddress2" size="30" value="<%= onepartner.FPartnerList(0).FManager_Address %>">(�Ż絿 123-45)
		</td>
	</tr>
	<tr>
		<td >Ȩ������ :</td>
		<td><input type="text" name="txurl" size="30" value="<%= onepartner.FPartnerList(0).FURL %>" maxlength="128">(http://www.10x10.co.kr)</td>
	</tr>
	<tr>
		<td >����� :</td>
		<td><input type="text" name="txmanagername" size="12" value="<%= onepartner.FPartnerList(0).FManager_Name %>"></td>
	</tr>
	<tr>
		<td >��ȭ :</td>
		<td><input type="text" name="txphone" size="12" value="<%= onepartner.FPartnerList(0).FTel %>">(02-123-4567)</td>
	</tr>
	<tr>
		<td >�ѽ� :</td>
		<td><input type="text" name="txfax" size="12" value="<%= onepartner.FPartnerList(0).FFax %>">(02-123-4568)</td>
	</tr>
	<tr>
		<td >�̸��� :</td>
		<td><input type="text" name="txemail" size="30" value="<%= onepartner.FPartnerList(0).FEmail %>" maxlength="128"></td>
	</tr>
	<tr>
		<td >Ŀ�̼� :</td>
		<td><%= CDbl(onepartner.FPartnerList(0).FCommission)*100 %> %</td>
	</tr>
	<tr>
		<td colspan="2"><br>**��й�ȣ�� �����Ͻ÷��� �Ʒ� ���� ä�� �ֽñ�ٶ��ϴ�.</td>
	</tr>
	<tr>
		<td >�����й�ȣ :</td>
		<td><input type="password" name="txnewpassword1" size="12" value="" maxlength="16"></td>
	</tr>
	<tr>
		<td >�����й�ȣ Ȯ��:</td>
		<td><input type="password" name="txnewpassword2" size="12" value="" maxlength="16"></td>
	</tr>
	<tr>
		<td colspan="2" height="30" align="center"><input type="button" value="����" onClick="useredit(frmedit)"></td>
	</tr>
	</form>
</table>
<%
set onepartner = Nothing
%>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->