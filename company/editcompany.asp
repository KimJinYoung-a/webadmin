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
			alert('필수 입력 사항입니다.');
			e.focus();
			return;
		}
	  }
	}
	
	if (frm.txnewpassword1.value!=frm.txnewpassword2.value){
		alert('새 비밀번호가 일치하지 않습니다.');
		frm.txnewpassword2.focus();
		return;
	}
	
	var ret = confirm('수정 하시겠습니까?');
	if (ret){
		frm.submit();
	}
}
</script>
<table width="610" border="0" class="a">
	<form name="frmedit" method="post" action="doeditcompany.asp">
	<tr>
		<td width="120">아이디 :</td>
		<td><%= onepartner.FPartnerList(0).FID %></td>
	</tr>
	<tr>
		<td >업체명 :</td>
		<td><input type="text" name="txcompanyname" value="<%= onepartner.FPartnerList(0).FCompany_name %>"></td>
	</tr>
	<tr>
		<td >비밀번호 :</td>
		<td><input type="password" name="txpassword" value="" size="12" maxlength="16"></td>
	</tr>
	<tr>
		<td >주소 :</td>
		<td>
			<input type="text" name="txaddress1" value="<%= onepartner.FPartnerList(0).FAddress %>">(서울 강남구)<br>
			<input type="text" name="txaddress2" size="30" value="<%= onepartner.FPartnerList(0).FManager_Address %>">(신사동 123-45)
		</td>
	</tr>
	<tr>
		<td >홈페이지 :</td>
		<td><input type="text" name="txurl" size="30" value="<%= onepartner.FPartnerList(0).FURL %>" maxlength="128">(http://www.10x10.co.kr)</td>
	</tr>
	<tr>
		<td >담당자 :</td>
		<td><input type="text" name="txmanagername" size="12" value="<%= onepartner.FPartnerList(0).FManager_Name %>"></td>
	</tr>
	<tr>
		<td >전화 :</td>
		<td><input type="text" name="txphone" size="12" value="<%= onepartner.FPartnerList(0).FTel %>">(02-123-4567)</td>
	</tr>
	<tr>
		<td >팩스 :</td>
		<td><input type="text" name="txfax" size="12" value="<%= onepartner.FPartnerList(0).FFax %>">(02-123-4568)</td>
	</tr>
	<tr>
		<td >이메일 :</td>
		<td><input type="text" name="txemail" size="30" value="<%= onepartner.FPartnerList(0).FEmail %>" maxlength="128"></td>
	</tr>
	<tr>
		<td >커미션 :</td>
		<td><%= CDbl(onepartner.FPartnerList(0).FCommission)*100 %> %</td>
	</tr>
	<tr>
		<td colspan="2"><br>**비밀번호를 변경하시려면 아래 란을 채워 주시기바랍니다.</td>
	</tr>
	<tr>
		<td >변경비밀번호 :</td>
		<td><input type="password" name="txnewpassword1" size="12" value="" maxlength="16"></td>
	</tr>
	<tr>
		<td >변경비밀번호 확인:</td>
		<td><input type="password" name="txnewpassword2" size="12" value="" maxlength="16"></td>
	</tr>
	<tr>
		<td colspan="2" height="30" align="center"><input type="button" value="저장" onClick="useredit(frmedit)"></td>
	</tr>
	</form>
</table>
<%
set onepartner = Nothing
%>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->