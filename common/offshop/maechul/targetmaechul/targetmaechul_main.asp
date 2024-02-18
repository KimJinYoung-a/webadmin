<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  목표매출 관리
' History : 2013.03.15 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
Dim yyyy1 ,mm1 ,i , shopid ,research
	research = requestcheckvar(request("research"),2)
	shopid = requestcheckvar(request("shopid"),32)
	yyyy1 = requestcheckvar(request("yyyy1"),4)
	mm1 = requestcheckvar(request("mm1"),2)

	if yyyy1 = "" then yyyy1 = year(now())
	if mm1 = "" then mm1 = month(now())

if research <> "ON" and shopid = "" then
	'/매장
	if (C_IS_SHOP) then

		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	end if
end if

if shopid = "" then
	response.write "<script language='javascript'>"
	response.write "	alert('매장을 선택해 주세요');"
	response.write "</script>"
end if
%>

<script language="javascript">

var iTimer;
var tempTerm = 0;

//1초 있다가 한번만 실행
function reRetry() {
	var tagetmaechul_zone = document.getElementById("tagetmaechul_zone");

    if(tagetmaechul_zone.document.body.readyState == "complete") {
        clearInterval(iTimer);
    }else {
        reFrame();
    }

    if(tempTerm>0) {
        clearInterval(iTimer);
    } else {
        reFrame();
        tempTerm++;
    }
}

function reFrame(){
	document.all.tagetmaechul_zone.src='/common/offshop/maechul/targetmaechul/targetmaechul_sub.asp?shopid=<%=shopid%>&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&gubuntype=2&menupos=<%=menupos%>';

}

function frmsubmit(){
	frm.submit();
}

</script>

<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="research" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		&nbsp;&nbsp;
		* 기간 : <% DrawYMBox yyyy1,mm1 %>
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>
</table>
<!---- /검색 ---->

<Br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr>
	<td align="left">
		<% if shopid <> "" then %>
			<!-- 목표매출 -->
			<iframe id="tagetmaechul" name="tagetmaechul" src="/common/offshop/maechul/targetmaechul/targetmaechul_sub.asp?shopid=<%=shopid%>&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&gubuntype=1&menupos=<%=menupos%>" width="100%" onload="this.style.height=this.contentWindow.document.body.scrollHeight;"  frameborder="0" scrolling="no" allowtransparency="true"></iframe>
		<% end if %>
	</td>
<tr>
</tr>
	<td align="right">
		<% if shopid <> "" then %>
			<!-- 조닝별 매출 1초후 로딩-->
			<iframe id="tagetmaechul_zone" name="tagetmaechul_zone" src='/common/offshop/maechul/targetmaechul/targetmaechul_sub.asp?shopid=<%=shopid%>&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&gubuntype=2&menupos=<%=menupos%>' width="100%" onload="this.style.height=this.contentWindow.document.body.scrollHeight;"  frameborder="0" scrolling="no" allowtransparency="true"></iframe>
			<!--<iframe id="tagetmaechul_zone" onload="iTimer = setInterval('reRetry()',1000);this.style.height=this.contentWindow.document.body.scrollHeight;" name="tagetmaechul_zone" src="/common/offshop/maechul/targetmaechul/targetmaechul_loading.asp" width="100%" frameborder="0" scrolling="no" allowtransparency="true"></iframe>-->
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- 표 중간바 끝-->

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
