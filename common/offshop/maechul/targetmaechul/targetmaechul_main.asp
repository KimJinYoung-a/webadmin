<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��ǥ���� ����
' History : 2013.03.15 �ѿ�� ����
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
	'/����
	if (C_IS_SHOP) then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	end if
end if

if shopid = "" then
	response.write "<script language='javascript'>"
	response.write "	alert('������ ������ �ּ���');"
	response.write "</script>"
end if
%>

<script language="javascript">

var iTimer;
var tempTerm = 0;

//1�� �ִٰ� �ѹ��� ����
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

<!---- �˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="research" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���� : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		&nbsp;&nbsp;
		* �Ⱓ : <% DrawYMBox yyyy1,mm1 %>
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit();">
	</td>
</tr>
</table>
<!---- /�˻� ---->

<Br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr>
	<td align="left">
		<% if shopid <> "" then %>
			<!-- ��ǥ���� -->
			<iframe id="tagetmaechul" name="tagetmaechul" src="/common/offshop/maechul/targetmaechul/targetmaechul_sub.asp?shopid=<%=shopid%>&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&gubuntype=1&menupos=<%=menupos%>" width="100%" onload="this.style.height=this.contentWindow.document.body.scrollHeight;"  frameborder="0" scrolling="no" allowtransparency="true"></iframe>
		<% end if %>
	</td>
<tr>
</tr>
	<td align="right">
		<% if shopid <> "" then %>
			<!-- ���׺� ���� 1���� �ε�-->
			<iframe id="tagetmaechul_zone" name="tagetmaechul_zone" src='/common/offshop/maechul/targetmaechul/targetmaechul_sub.asp?shopid=<%=shopid%>&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&gubuntype=2&menupos=<%=menupos%>' width="100%" onload="this.style.height=this.contentWindow.document.body.scrollHeight;"  frameborder="0" scrolling="no" allowtransparency="true"></iframe>
			<!--<iframe id="tagetmaechul_zone" onload="iTimer = setInterval('reRetry()',1000);this.style.height=this.contentWindow.document.body.scrollHeight;" name="tagetmaechul_zone" src="/common/offshop/maechul/targetmaechul/targetmaechul_loading.asp" width="100%" frameborder="0" scrolling="no" allowtransparency="true"></iframe>-->
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- ǥ �߰��� ��-->

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
