<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체정보
' History : 2009.04.17 최초생성자 모름
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->

<%
dim ogroup,i, groupid

dim opartner
set opartner = new CPartnerUser
	opartner.FCurrpage = 1
	opartner.FRectDesignerID = session("ssBctId")
	opartner.FPageSize = 1
	opartner.GetOnePartnerNUser

groupid = opartner.FOneItem.FGroupid

'groupid = "G00240"

set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = groupid
	ogroup.GetOneGroupInfo

dim OReturnAddr
set OReturnAddr = new CCSReturnAddress
	OReturnAddr.FRectGroupCode = groupid
	OReturnAddr.GetReturnAddressList

%>

<script type="text/javascript">

function CopyZip(flag,post1,post2,add,dong){
	var frm = eval(flag);

	frm.return_zipcode.value= post1 + "-" + post2;
	frm.return_address.value= add;
	frm.return_address2.value= dong;
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function SameReturnAddr(frm, bool){
	if (bool){
		frm.return_zipcode.value = document.frmgroup.return_zipcode.value;
		frm.return_address.value = document.frmgroup.return_address.value;
		frm.return_address2.value = document.frmgroup.return_address2.value;
	}else{
		frm.return_zipcode.value = "";
		frm.return_address.value = "";
		frm.return_address2.value = "";
	}
}

function SameReturnName(frm, bool){
	if (bool){
		frm.deliver_name.value = document.frmgroup.deliver_name.value;
		frm.deliver_phone.value = document.frmgroup.deliver_phone.value;
		frm.deliver_hp.value = document.frmgroup.deliver_hp.value;
		frm.deliver_email.value = document.frmgroup.deliver_email.value;
	}else{
		frm.deliver_name.value = "";
		frm.deliver_phone.value = "";
		frm.deliver_hp.value = "";
		frm.deliver_email.value = "";
	}
}


function ModifyReturnAddress(frm){

	if (frm.return_zipcode.value.length < 1){
		alert('우편번호를 선택하세요.');
		frm.return_zipcode.focus();
		return;
	}

	if (frm.return_address.value.length < 1){
		alert('주소를 정확히 입력하세요.');
		frm.return_address.focus();
		return;
	}

	if (frm.return_address2.value.length < 1){
		alert('주소를 정확히 입력하세요.');
		frm.return_address2.focus();
		return;
	}

	if (frm.deliver_name.value.length < 1){
		alert('배송담당자 이름을 입력하세요.');
		frm.deliver_name.focus();
		return;
	}

	if (frm.deliver_phone.value.length < 1){
		alert('배송담당자 전화번호를 입력하세요.');
		frm.deliver_phone.focus();
		return;
	}

	if (frm.deliver_hp.value.length < 1){
		alert('배송담당자 핸드폰번호를 입력하세요.');
		frm.deliver_hp.focus();
		return;
	}

	if (frm.deliver_email.value.length < 1){
		alert('배송담당자 이메일주소를 입력하세요.');
		frm.deliver_email.focus();
		return;
	}

	var ret = confirm('브랜드 반품정보를 수정 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}

function PopUpcheList(frmname){
	var popwin = window.open("/admin/lib/popupchelist.asp?frmname=" + frmname,"popupchelist","width=700 height=540 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=4><b>* 브랜드별 반품정보 및 택배사 설정</b></td>
</tr>

<!--
<tr>
<td colspan="4" bgcolor="#FFFFFF" height="25">**사무실 주소**</td>
</tr>
-->
<tr>
	<td height="25" width="150" bgcolor="<%= adminColor("tabletop") %>">상호</td>
	<td width="250" bgcolor="#FFFFFF"><b><%= ogroup.FOneItem.FCompany_name %></b></td>
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">그룹코드</td>
	<td bgcolor="#FFFFFF"><b><%= ogroup.FOneItem.FGroupId %></b></td>
</tr>
<!--
<tr>
<td height="25" bgcolor="<%= adminColor("tabletop") %>">배송담당자명</td>
<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_name %></td>
<td bgcolor="<%= adminColor("tabletop") %>">배송담당자 전화번호</td>
<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_phone %></td>
</tr>
<tr>
<td height="25" bgcolor="<%= adminColor("tabletop") %>">배송담당자 이메일</td>
<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_email %></td>
<td bgcolor="<%= adminColor("tabletop") %>">배송담당자 핸드폰번호</td>
<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_hp %></td>
</tr>
<tr>
<form name=frmgroup>
<input type=hidden name=return_zipcode value="<%= ogroup.FOneItem.Freturn_zipcode %>">
<input type=hidden name=return_address value="<%= ogroup.FOneItem.Freturn_address %>">
<input type=hidden name=return_address2 value="<%= ogroup.FOneItem.Freturn_address2 %>">
<input type=hidden name=deliver_name value="<%= ogroup.FOneItem.Fdeliver_name %>">
<input type=hidden name=deliver_phone value="<%= ogroup.FOneItem.Fdeliver_phone %>">
<input type=hidden name=deliver_hp value="<%= ogroup.FOneItem.Fdeliver_hp %>">
<input type=hidden name=deliver_email value="<%= ogroup.FOneItem.Fdeliver_email %>">
</form>
<td height="25" bgcolor="<%= adminColor("tabletop") %>">주소</td>
<td colspan="3" bgcolor="#FFFFFF" >
[<%= ogroup.FOneItem.Freturn_zipcode %>] <%= ogroup.FOneItem.Freturn_address %> <%= ogroup.FOneItem.Freturn_address2 %>
</td>
</tr>
-->
</table>

<br>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
	**브랜드별 반품정보 설정**
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">브랜드ID</td>
	<td>브랜드명</td>
	<td width="200">반품주소</td>
	<td width="80"></td>
	<td width="130">배송 담당자</td>
	<td width="130">CS 담당자</td>
	<td width="130">사용택배사</td>
	<td width="50">변경</td>
</tr>
<% for i=0 to OReturnAddr.FResultCount - 1%>
<tr bgcolor="#FFFFFF">
	<form name=frm<%= i %> method=post action=doupcheedit.asp onsubmit="return false;">
		<input type=hidden name=uid value="<%= OReturnAddr.FItemList(i).Fbrandid %>">
		<input type=hidden name=mode value="modifyreturnaddress">
		<input type=hidden name=makerid value="<%= OReturnAddr.FItemList(i).Fbrandid %>">
		<td height="25" align="center" rowspan="4"><%= OReturnAddr.FItemList(i).Fbrandid %></td>
		<td align="center" rowspan="4"><%= OReturnAddr.FItemList(i).Fstreetname_kor %><br><%= OReturnAddr.FItemList(i).Fstreetname_eng %></td>
		<td rowspan="4">
			<input type="text" class="text" name="return_zipcode" value="<%= OReturnAddr.FItemList(i).FreturnZipcode %>" size="7" maxlength="7">
			<input type="button" class="button" value="검색" onClick="TnFindZipNewdesigner('frm<%= i %>','D')">
			<input type="button" class="button" value="검색(구)" onclick="javascript:popZip('frm<%= i %>');">
			<br>
			<input type="text" class="text" name="return_address" value="<%= OReturnAddr.FItemList(i).FreturnZipaddr %>" size="25" maxlength="64">
			<br>
			<input type="text" class="text" name="return_address2" value="<%= OReturnAddr.FItemList(i).FreturnEtcaddr %>" size="50" maxlength="64">
		</td>
		<td align="left">
	이름
		</td>
		<td align="left">
			<input type="text" class="text" name="deliver_name" value="<%= OReturnAddr.FItemList(i).FreturnName %>" size="8" maxlength="32">
		</td>
		<td align="left">
			<input type="text" class="text" name="cs_name" value="<%= OReturnAddr.FItemList(i).FcsName %>" size="8" maxlength="32">
		</td>
		<td align="center" rowspan="4"><% drawSelectBoxDeliverCompany "defaultsongjangdiv",OReturnAddr.FItemList(i).Fsongjangdiv %></td>
		<td align="center" rowspan="4">
			<input type="button" class="button" value="수정" onclick="ModifyReturnAddress(frm<%= i %>)">
		</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="left">
	전화번호
	</td>
	<td align="left">
		<input type="text" class="text" name="deliver_phone" value="<%= OReturnAddr.FItemList(i).FreturnPhone %>" size="16" maxlength="32">
	</td>
	<td align="left">
		<input type="text" class="text" name="cs_phone" value="<%= OReturnAddr.FItemList(i).FcsPhone %>" size="16" maxlength="32">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="left">
	휴대폰
	</td>
	<td align="left">
		<input type="text" class="text" name="deliver_hp" value="<%= OReturnAddr.FItemList(i).Freturnhp %>" size="16" maxlength="32">
	</td>
	<td align="left">
		<input type="text" class="text" name="cs_hp" value="<%= OReturnAddr.FItemList(i).Fcshp %>" size="16" maxlength="32">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="left">
	이메일
	</td>
	<td align="left">
		<input type="text" class="text" name="deliver_email" value="<%= OReturnAddr.FItemList(i).FreturnEmail %>" size="16" maxlength="64">
	</td>
	<td align="left">
		<input type="text" class="text" name="cs_email" value="<%= OReturnAddr.FItemList(i).FcsEmail %>" size="16" maxlength="64">
	</td>
</tr>
</form>
<% next %>
</table>

<%
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->