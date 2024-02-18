<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/cardPartialCancelCls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->

<%

''한글 한글

Dim orderserial : orderserial=requestCheckvar(request("orderserial"),20)
Dim tid         : tid=requestCheckvar(request("tid"),100)
Dim i, rURi
dim paddparam

''rw orderserial
''rw tid

Dim sqlStr, ipkumdate, acctamount, pggubun
sqlStr = " select ipkumdate, P.acctamount, IsNull(m.pggubun,'') as pggubun, IsNull(p.paddparam,'') as paddparam from db_order.dbo.tbl_order_Master M"
sqlStr = sqlStr & " 	left join db_order.dbo.tbl_order_Paymentetc P"
sqlStr = sqlStr & " 	on M.orderserial=P.orderserial"
sqlStr = sqlStr & " 	and P.acctdiv='100'"
sqlStr = sqlStr & " where M.orderserial='"&orderserial&"'"

rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
    ipkumdate  = rsget("ipkumdate")
    acctamount = rsget("acctamount")
	pggubun = rsget("pggubun")
	paddparam = rsget("paddparam")
end if
rsget.close


rURi = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?" + "noTid=" + tid + "&noMethod=1"

if (pggubun = "KA") then
	rURi = "https://mms.cnspay.co.kr/trans/retrieveIssueLoader.do?TID=" + tid + "&type=0"
end if

if (pggubun = "NP") then
    rURi = "/cscenter/taxsheet/popNpayRecirect.asp?orderserial=" + orderserial + "&tid="&tid
end if

if (pggubun = "PY") then
    rURi = "/cscenter/taxsheet/popPaycoRecirect.asp?orderserial=" + orderserial + "&tid="&tid + "&paddparam=" + paddparam
end if

if (pggubun = "TS") then
	if (application("Svr_Info")="Dev") then
		rURi = "https://pay.toss.im/payfront/web/external/sales-check?payToken=" & tid & "&transactionId=12637496-8a46-488c-bc30-febded96656f"
	else
		rURi = "https://pay.toss.im/payfront/web/external/sales-check?payToken=" & tid & "&transactionId=12637496-8a46-488c-bc30-febded96656f"
	end if
end if

Dim ocardCancel
set ocardCancel = new CCardPartialCancel
ocardCancel.FRectOrderserial = orderserial
ocardCancel.getCardCancelList

if (ocardCancel.FResultCount<1) then
    set ocardCancel = Nothing
    dbget.Close()

    response.redirect rURi
    response.end
else

%>
<style>
body, tr, td {font-size:9pt; font-family:굴림,verdana; color:#433F37; line-height:19px;}
table, img {border:none}

/* Padding ******/
.pl_01 {padding:1 10 0 10; line-height:19px;}
.pl_03 {font-size:20pt; font-family:굴림,verdana; color:#FFFFFF; line-height:29px;}

/* Link ******/
.a:link  {font-size:9pt; color:#333333; text-decoration:none}
.a:visited { font-size:9pt; color:#333333; text-decoration:none}
.a:hover  {font-size:9pt; color:#0174CD; text-decoration:underline}

.txt_03a:link  {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:visited {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:hover  {font-size: 8pt;line-height:18px;color:#EC5900; text-decoration:underline}

.buttoncss {
	font-family: "Verdana", "돋움";
	font-size: 9pt;
	background-color: #E6E6E6;
	border: 1px outset #BABABA;
	color: #000000;
	height: 20px;
	cursor:hand;
}

</style>
<script language='javascript'>
window.resizeTo(680,400);

function receiptinicis(tid){
	var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?" + "noTid=" + tid + "&noMethod=1";
	var popwin = window.open(receiptUrl,"INIreceipt","width=415,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function receiptKakao(tid){
	var status = "toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes,resizable=yes,width=420,height=540";
    var url = "https://mms.cnspay.co.kr/trans/retrieveIssueLoader.do?TID="+tid+"&type=0";
    var popwin = window.open(url,"popupIssue",status);
	popwin.focus();
}

function receiptNaverPay(tid){
	var status = "toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes,resizable=yes,width=420,height=540";
    var url = "/cscenter/taxsheet/popNpayRecirect.asp?orderserial=<%=orderserial%>&tid="+tid;
    var popwin = window.open(url,"popupIssue",status);
	popwin.focus();
}

function receiptPaycoPay(tid){
	var status = "toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes,resizable=yes,width=420,height=540";
    var url = "/cscenter/taxsheet/popPaycoRecirect.asp?orderserial=<%=orderserial%>&tid="+tid;
    var popwin = window.open(url,"popupIssue",status);
	popwin.focus();
}

// 토스 전표 팝업
function receiptTossPay(tid){
	<% if (application("Svr_Info")="Dev") then %>
	var receiptUrl = "https://pay.toss.im/payfront/web/external/sales-check?payToken="+tid+"&transactionId=12637496-8a46-488c-bc30-febded96656f";
	<% else %>
	var receiptUrl = "https://pay.toss.im/payfront/web/external/sales-check?payToken="+tid+"&transactionId=12637496-8a46-488c-bc30-febded96656f";
	<% end if %>
	var popwin = window.open(receiptUrl,"Tossreceipt","width=415,height=600");
	popwin.focus();
}

</script>
<body bgcolor="#FFFFFF" text="#242424" leftmargin=0 topmargin=0 marginwidth=0 marginheight=0 bottommargin=0 rightmargin=0><center>

<table width="650" border="0" cellspacing="0" cellpadding="0">
    <tr>
	    <!---- 팝업제목 시작 ---->
	    <td valign="top" bgcolor="#af1414"><img src="http://fiximage.10x10.co.kr/web2011/mytenbyten/pop_reciept_sel_tit.gif" width="650" height="60"></td>
	    <!---- 팝업제목 끝 ---->
  	</tr>
  	<tr>
    	<td style="padding:0px 15px">

    		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		        <tr>
          			<td style="padding:25px 0 7px 0;">
		        		<b>신용카드 부분취소 증빙서류 내역입니다.</b>
          			</td>
        		</tr>
		       	<tr>
          			<td>

        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top:solid 3px #be0808; border-bottom:solid 1px #eaeaea; padding-top:3px;">
        <tr align="center" height="30" bgcolor="#fcf6f6">
            <td width="60" style="border-bottom:solid 1px #eaeaea;" >구분</td>
            <td width="60" style="border-bottom:solid 1px #eaeaea;" >취소차수</td>
            <td style="border-bottom:solid 1px #eaeaea;" >결제일</td>
            <td style="border-bottom:solid 1px #eaeaea;" >취소요청일</td>
            <td width="100" style="border-bottom:solid 1px #eaeaea;" >취소액</td>
            <td width="100" style="border-bottom:solid 1px #eaeaea;" >승인잔액</td>
            <td width="100" style="border-bottom:solid 1px #eaeaea;" >전표</td>
        </tr>
        <tr align="center" height="30" bgcolor="#FFFFFF" >
            <td align="center" >최초<br>승인 </td>
            <td >&nbsp;</td>
            <td ><%= ipkumdate %></td>
            <td >&nbsp;</td>
            <td >&nbsp;</td>
            <td ><%= FormatNumber(acctamount,0) %></td>
            <%
			if (pggubun = "KA") then
				rURi = "javascript:receiptKakao('" + tid + "');"
			elseif (pggubun = "NP") then
			    rURi = "javascript:receiptNaverPay('" + tid + "');"
			elseif (pggubun = "PY") then
			    rURi = "javascript:receiptPaycoPay('" + tid + "');"
			elseif (pggubun = "TS") then
			    rURi = "javascript:receiptTossPay('" + tid + "');"
			else
				rURi = "javascript:receiptinicis('" + tid + "');"
			end if


            %>
            <td ><a href="<%= rURi %>"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/cs_icon01.gif" border="0"></a></td>
        </tr>

        <% for i=0 to ocardCancel.FResultCount-1 %>
        <tr align="center" height="30" bgcolor="#FFFFFF" style="border-top:solid 1px #eaeaea;">
            <td style="border-top:solid 1px #eaeaea;">취소 </td>
            <td style="border-top:solid 1px #eaeaea;"><%= ocardCancel.FItemList(i).Fcancelrequestcount%></td>
            <td style="border-top:solid 1px #eaeaea;">&nbsp;</td>
            <td style="border-top:solid 1px #eaeaea;"><%= ocardCancel.FItemList(i).Fregdate%></td>
            <td style="border-top:solid 1px #eaeaea;"><b><%= FormatNumber(ocardCancel.FItemList(i).Fcancelprice,0) %></b></td>
            <td style="border-top:solid 1px #eaeaea;"><%= FormatNumber(ocardCancel.FItemList(i).Frepayprice,0)%></td>
            <%
			if (pggubun = "KA") then
				rURi = "javascript:receiptKakao('" + ocardCancel.FItemList(i).Fnewtid + "');"
			elseif (pggubun = "NP") then
			    rURi = "javascript:receiptNaverPay('" + tid + "');"  '' 최초결제키만 되는듯.
			elseif (pggubun = "PY") then
			    rURi = "javascript:receiptPaycoPay('" + tid + "');"
			elseif (pggubun = "TS") then
			    rURi = "javascript:receiptTossPay('" + tid + "');"
			else
				rURi = "javascript:receiptinicis('" + ocardCancel.FItemList(i).Fnewtid + "');"
			end if
            %>
            <td style="border-top:solid 1px #eaeaea;"><a href="<%= rURi %>"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/cs_icon01.gif" border="0"></a></td>
        </tr>
        <% next %>
        </table>
            </td>
            </tr>
        </table>
</td></tr>
</table>
<%
end if
set ocardCancel = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp" -->
