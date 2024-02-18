<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Dim oCoupang, i, mode, maeipdiv
Dim makerid
makerid	= request("id")

If makerid = "" AND maeipdiv = "" Then
	Call Alert_Close("브랜드ID or 업체구분값이 없습니다.")
	dbget.Close: Response.End
End IF

Set oCoupang = new CCoupang
	oCoupang.FRectMakerid = makerid
	oCoupang.getTenCoupangOneBrandDeliver
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
	function fnSaveForm() {
		var frm = document.frm;
	<% If maeipdiv = "U" Then %>
		if($("#phoneNumber2").val() == ""){
			alert('추가전화번호를 입력하세요');
			$("#phoneNumber2").focus();
			return false;
		}
		if($("#deliveryCode").val() == ""){
			alert('택배사를 선택하세요');
			$("#deliveryCode").focus();
			return false;
		}
		if($("#returnZipCode").val() == ""){
			alert('우편번호를 입력하세요');
			$("#returnZipCode").focus();
			return false;
		}
		if($("#returnAddress").val() == ""){
			alert('주소1을 입력하세요');
			$("#returnAddress").focus();
			return false;
		}
		if($("#returnAddressDetail").val() == ""){
			alert('주소2를 입력하세요');
			$("#returnAddressDetail").focus();
			return false;
		}
		if($("#jeju").val() == ""){
			alert('도서산간_제주를 입력하세요');
			$("#jeju").focus();
			return false;
		}
		if($("#notJeju").val() == ""){
			alert('도서산간_제주외를 입력하세요');
			$("#notJeju").focus();
			return false;
		}
	<% End If %>
		if(confirm("저장하시겠습니까?")) {
			frm.action="procBrandMapping.asp";
			frm.submit();
		}
	}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="maeipdiv" value="<%= maeipdiv %>">
<input type="hidden" name="gubun" value="popup">
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oCoupang.FItemList(0).FId %></td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">브랜드명(한글)</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oCoupang.FItemList(0).FSocname_kor %></td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">브랜드명(영문)</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oCoupang.FItemList(0).FSocname %></td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">추가전화번호</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" id="phoneNumber2" name="phoneNumber2" value="<%= oCoupang.FItemList(0).FDeliverPhone %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">택배사</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<select name="deliveryCode" id="deliveryCode" class="select">
			<option value=""></option>
			<option value="HYUNDAI" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "2", "selected", "") %> >롯데택배</option>
			<option value="KGB" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "18", "selected", "") %> >로젠택배</option>
			<option value="EPOST" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "8", "selected", "") %> >우체국</option>
			<option value="HANJIN" <%= Chkiif((oCoupang.FItemList(0).FDefaultSongjangDiv = "1") OR (oCoupang.FItemList(0).FDefaultSongjangDiv = "36"), "selected", "") %> >한진택배</option>
			<option value="CJGLS" <%= Chkiif( (oCoupang.FItemList(0).FDefaultSongjangDiv = "3") OR (oCoupang.FItemList(0).FDefaultSongjangDiv = "4"), "selected", "") %> >CJ대한통운</option>
			<option value="KDEXP" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "21", "selected", "") %> >경동택배</option>
			<option value="DONGBU" <%= Chkiif((oCoupang.FItemList(0).FDefaultSongjangDiv = "39") OR (oCoupang.FItemList(0).FDefaultSongjangDiv = "41"), "selected", "") %> >드림택배(구 KG로지스)</option>
			<option value="ILYANG" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "26", "selected", "") %> >일양택배</option>
			<option value="CHUNIL" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "31", "selected", "") %> >천일택배</option>
			<option value="AJOU" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "10", "selected", "") %> >아주택배</option>
			<option value="CSLOGIS" <%= Chkiif((oCoupang.FItemList(0).FDefaultSongjangDiv = "5") OR (oCoupang.FItemList(0).FDefaultSongjangDiv = "24"), "selected", "") %> >SC로지스</option>
			<option value="DAESIN" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "34", "selected", "") %> >대신택배</option>
			<option value="CVS" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "35", "selected", "") %> >CVS택배</option>
			<option value="HDEXP" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "37", "selected", "") %> >합동택배</option>
			<option value="DADREAM">다드림</option>
			<option value="DHL" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "91", "selected", "") %> >DHL</option>
			<option value="UPS">UPS</option>
			<option value="FEDEX">FEDEX</option>
			<option value="REGISTPOST">우편등기</option>
			<option value="DIRECT">업체직송</option>
			<option value="COUPANG">쿠팡자체배송</option>
			<option value="EMS">우체국 EMS</option>
			<option value="TNT">TNT</option>
			<option value="USPS">USPS</option>
			<option value="IPARCEL">i-parcel</option>
			<option value="GSMNTON">GSM NtoN</option>
			<option value="SWGEXP">성원글로벌</option>
			<option value="PANTOS">범한판토스</option>
			<option value="ACIEXPRESS">ACI Express</option>
			<option value="DAEWOON">대운글로벌</option>
			<option value="AIRBOY">에어보이익스프레스</option>
			<option value="KGLNET">KGL네트웍스</option>
			<option value="KUNYOUNG" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "29", "selected", "") %> >건영택배</option>
			<option value="SLX">SLX택배</option>
			<option value="HONAM" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "33", "selected", "") %> >호남택배</option>
			<option value="LINEEXPRESS">LineExpress</option>
			<option value="SFEXPRESS">순풍택배</option>
			<option value="TWOFASTEXP">2FastsExpress</option>
			<option value="ECMS">ECMS익스프레스
		</select>
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">우편번호</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" name="returnZipCode" id="returnZipCode" value="<%= oCoupang.FItemList(0).FReturn_zipcode %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">주소1</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" name="returnAddress" id="returnAddress" value="<%= oCoupang.FItemList(0).FReturn_address %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">주소2</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" size="50" name="returnAddressDetail" id="returnAddressDetail" value="<%= oCoupang.FItemList(0).FReturn_address2 %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">도서산간_제주</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" name="jeju" id="jeju" value="<%= oCoupang.FItemList(0).FJeju %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">도서산간_제주외</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" name="notJeju" id="notJeju" value="<%= oCoupang.FItemList(0).FNotJeju %>">
	</td>
</tr>
<tr align="center">
	<td colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="저장" onclick="fnSaveForm();">
		<input type="button" class="button" value="취소" onclick="self.close();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
