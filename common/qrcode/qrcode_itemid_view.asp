<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.charset = "utf-8" %>
<%
'#######################################################
'	Description : 상품 QRCODE
'	History	:  2015.06.02 한용민 생성
'#######################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim itemgubun, itemoption, itemid, wd, ht, qt, imgPath, msg
	itemgubun = getNumeric(requestcheckvar(request("itemgubun"),2))
	itemoption = getNumeric(requestcheckvar(request("itemoption"),4))
	itemid = getNumeric(requestcheckvar(request("itemid"),10))
	wd = getNumeric(requestcheckvar(request("wd"),3))
	ht = getNumeric(requestcheckvar(request("ht"),3))
	qt = requestcheckvar(request("qt"),1)

if wd = "" then wd = 500
if ht = "" then ht = 500
if qt = "" then qt = "M"

if itemid = "" then
	response.write "<script type='text/javascript'>alert('상품번호가 없습니다.');</script>"
	response.end
end if

msg = "http://m.10x10.co.kr/offshop/view/iteminfo.asp?itemid=" & itemid

'// 구글 Chart API - QRCode URL (반드시 UTF-8로 전송)
imgPath = "http://chart.apis.google.com/chart?cht=qr&chl=" & URLEncodeUTF8(msg) & "&choe=UTF-8&chs=" & wd & "x" & ht & "&chld=" & qt & "|1"
%>
<script type="text/javascript">

function pageclose(){
	self.close();
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<img src="<%= imgPath %>" width="<%= wd %>" height="<%= ht %>" onclick="pageclose();" />
	</td>
</tr>
<tr>
	<td align="left">
		<br><br><br>
		QR코드를 다운로드 하실려면 마우스 오른쪽 버튼을 누르신후 저장해 주세요.
	</td>
</tr>
</table>
<!-- 액션 끝 -->
