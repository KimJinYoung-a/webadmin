<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim onepartner,i,page
page = request("page")
if page="" then page=1
set onepartner = new CPartnerUser
onepartner.FCurrpage = page
onepartner.GetPartnerList 1
%>
<script language="javascript">
function useradd(frm){

	for (var i=0;i<frm.elements.length;i++){
	  var e = frm.elements[i];

	  if ((e.name=="userid")||(e.name=="userpass")||(e.name=="username")) {
		if (e.value.length<1){
			alert('필수 입력 사항입니다.');
			e.focus();
			return;
		}
	  }

	}



	var ret = confirm('추가 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}

function useredit(frm){

	for (var i=0;i<frm.elements.length;i++){
	  var e = frm.elements[i];

	  if ((e.name=="userid")||(e.name=="userpass")||(e.name=="username")) {
		if (e.value.length<1){
			alert('필수 입력 사항입니다.');
			e.focus();
			return;
		}
	  }

	}



	var ret = confirm('수정 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}
</script>
<table width="610" border="0" class="a">
	<tr>
		<td width="70">ID</td>
		<td width="70">Pass</td>
		<td width="70">Name</td>
		<td width="70">할인율</td>
		<td width="70">커미션</td>
		<td width="70">비고</td>
		<td width="100">E-Mail</td>
		<td width="80">사용여부</td>
		<td width="70">수정</td>
	</tr>
	<% for i=0 to onepartner.FresultCount-1 %>
	<form name="frmedit_<%= i %>" method="post" action="domemberedit.asp">
	<input type="hidden" name="mode" value="partneredit">
	<input type="hidden" name="userid" value="<%= onepartner.FPartnerList(i).FID %>">
	<input type="hidden" name="divcd" value="<%= onepartner.FPartnerList(i).FUserDiv %>">
	<tr>
		<td><%= onepartner.FPartnerList(i).FID %></td>
		<td><input type="text" name="userpass" size="10" value="<%= onepartner.FPartnerList(i).FPassword %>"></td>
		<td><input type="text" name="username" size="10" value="<%= onepartner.FPartnerList(i).FCompany_name %>"></td>
		<td><input type="text" name="discountrate" size="4" value="<%= onepartner.FPartnerList(i).FDiscountrate %>"></td>
		<td><input type="text" name="commission" size="4" value="<%= onepartner.FPartnerList(i).Fcommission %>"></td>
		<td><input type="text" name="bigo" size="4" value="<%= onepartner.FPartnerList(i).Fbigo %>"></td>
		<td><input type="text" name="usermail" size="20" value="<%= onepartner.FPartnerList(i).Femail %>"></td>
		<td>
			<select name="isusing">
			<option value="Y" <% if onepartner.FPartnerList(i).Fisusing="Y" then response.write "selected" %> >Yes</option>
			<option value="N" <% if onepartner.FPartnerList(i).Fisusing="N" then response.write "selected" %> >No</option>
			</select>
		</td>
		<td align="left"><input type="button" value="수정" onclick="useredit(frmedit_<%= i %>)"></td>
	</tr>
	</form>
	<% next %>
	<tr>
		<td colspan="9" align="right">
		<% if onepartner.HasPreScroll then %>
		<a href="?page=<%= onepartner.StartScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + onepartner.StartScrollPage to onepartner.FScrollCount + onepartner.StartScrollPage - 1 %>
			<% if i>onepartner.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if onepartner.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<table width="500" border="0" class="a">
	<form name="frmadd" method="post" action="domemberedit.asp">
	<input type="hidden" name="mode" value="partneradd">
	<input type="hidden" name="divcd" value="999">
	<tr><td colspan="2">파트너추가</td></tr>
	<tr>
		<td width="80">ID</td>
		<td><input type="text" name="userid" size="8" maxlength="32"></td>
	</tr>
	<tr>
		<td>Pass</td>
		<td><input type="text" name="userpass" size="8" maxlength="32"></td>
	</tr>
	<tr>
		<td>Name</td>
		<td><input type="text" name="username" size="16" maxlength="32"></td>
	</tr>
	<tr>
		<td>할인율</td>
		<td><input type="text" name="discountrate" size="6" maxlength="4"></td>
	</tr>
	<tr>
		<td>커미션</td>
		<td><input type="text" name="commission" size="6" maxlength="4"></td>
	</tr>
	<tr>
		<td>비고</td>
		<td><input type="text" name="bigo" size="6" maxlength="4"></td>
	</tr>
	<tr>
		<td>E-mail</td>
		<td><input type="text" name="usermail" size="16" maxlength="32"></td>
	</tr>

	<tr>
		<td></td>
		<td align="left"><input type="button" value="추가" onclick="useradd(frmadd)"></td>
	</tr>
	</form>
</table>
<%
set onepartner = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->