<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<% Response.Addheader "P3P","policyref='/w3c/p3p.xml', CP='NOI DSP LAW NID PSA ADM OUR IND NAV COM'"%>

<!-- #include virtual="/login/checkUserGuestlogin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/classes/ordercls/cashreceiptcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/taxsheet_cls.asp"-->

<%

dim i, j
dim userid, sitename
sitename = "10x10"

dim orderserial
dim IsBiSearch

userid = getLoginUserID
orderserial = requestCheckVar(request("orderserial"),11)

if orderserial="" then
    orderserial = GetGuestLoginOrderserial
end if


''세금계산서 발행 요청건이 있으면 리다이렉트
Dim chkReg
chkReg = chkRegTax(orderserial)  ''Y : 발행, N : 발행전, none : 신청전
if (chkReg<>"none") then
    response.write "<script language='javascript'>alert('세금계산서 발행 요청건이 있습니다. \n\n세금계산서와 현금영수증은 동시에 발행하실 수 없습니다.');</script>"
    response.write "<script language='javascript'>window.close();</script>"
	Response.End
end if


if (orderserial="") then
	response.write "<script>alert('올바른 접속이 아닙니다..');</script>"
	response.write "<script>window.close();</script>"
	response.end
end if

''데이콤 가상계좌 현금영수증 발행내역 있는지 체크. ======= 201004==
dim DCcashreceiptNo
if (chkDacomCyberPayCashReciptExists(orderserial, DCcashreceiptNo)) then
    IF (application("Svr_Info")="Dev") then
        response.redirect "http://pg.dacom.net:7080/transfer/cashreceipt.jsp?orderid="&orderserial&"&mid=ttenbyten01&servicetype=SC0040&seqno=001"
    ELSE
        response.redirect "http://pg.dacom.net/transfer/cashreceipt.jsp?orderid="&orderserial&"&mid=tenbyten01&servicetype=SC0040&seqno=001"
    END IF
    dbget.Close() : response.end
end if
''=================================================================

dim ocashreceipt, receiptAlreadyExists
set ocashreceipt = new CCashReceipt
ocashreceipt.FRectIsSucces = "00"
ocashreceipt.FRectOrderserial = orderserial
ocashreceipt.FRectCancelyn = "N"
ocashreceipt.GetReceiptByOrderSerial

receiptAlreadyExists = (ocashreceipt.FResultcount>0)
%>

<% if (receiptAlreadyExists) then %>
<html>
<head>
<title>무통장 입금 현금영수증 발행</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="css/group.css" type="text/css">
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
</style>
<script>
	function showreceipt(tid){
 		var showreceiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/Cash_mCmReceipt.jsp?noTid=" + tid + "&clpaymethod=22";
 		window.open(showreceiptUrl,"showreceipt","width=380,height=540, scrollbars=no,resizable=no");
	}

</script>

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#242424" leftmargin=0 topmargin=15 marginwidth=0 marginheight=0 bottommargin=0 rightmargin=0><center>
<table width="632" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="83" background="http://fiximage.10x10.co.kr/web2007/receipt/cash_top.gif" style="padding:0 0 0 64">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="3%" valign="top"><img src="http://fiximage.10x10.co.kr/web2007/receipt/title_01.gif" width="8" height="27" vspace="5"></td>
          <td width="97%" height="40" class="pl_03"><font color="#FFFFFF"><b>현금결제 영수증 발급내용</b></font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td align="center" bgcolor="6095BC"><table width="620" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td bgcolor="#FFFFFF" style="padding:0 0 0 56">
		  <table width="510" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="7"><img src="http://fiximage.10x10.co.kr/web2007/receipt/life.gif" width="7" height="30"></td>
                <td background="http://fiximage.10x10.co.kr/web2007/receipt/center.gif"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon03.gif" width="12" height="10">
                  <b>고객님께서 요청하신 현금영수증 발급 내용입니다. </b></td>
                <td width="8"><img src="http://fiximage.10x10.co.kr/web2007/receipt/right.gif" width="8" height="30"></td>
              </tr>
            </table>
            <br>
            <table width="510" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="407"  style="padding:0 0 0 9"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon.gif" width="10" height="11">
                  <strong><font color="433F37">발급내역</font></strong></td>
                <td width="103">&nbsp;</td>
              </tr>
              <tr>
                <td colspan="2"  style="padding:0 0 0 23">
		  		<table width="470" border="0" cellspacing="0" cellpadding="0">
					<tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="109" height="26">발 급 결 과</td>
                      <td width="343"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td>성공</td>
                            <td width='142' align='right'>
                            <img src='http://fiximage.10x10.co.kr/web2007/receipt/button_02.gif' width='94' height='24' border='0' onclick="showreceipt('<%= ocashreceipt.FoneItem.FTid %>')" onMouseOver="this.style.cursor='hand';"></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="109" height="25">승 인 번 호</td>
                      <td width="343"><%= ocashreceipt.FoneItem.FResultCashNoAppl %></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>

                    <tr>
                      <td width='18' align='center'><img src='http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif' width='7' height='7'></td>
                      <td width='109' height='25'>총 현금결제금액</td>
                      <td width='343'><%= ocashreceipt.FoneItem.Fcr_price %></td>
                    </tr>

                    <tr>
                      <td height='1' colspan='3' align='center'  background='http://fiximage.10x10.co.kr/web2007/receipt/line.gif'></td>
                    </tr>
                    <tr>
                      <td width='18' align='center'><img src='http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif' width='7' height='7'></td>
                      <td width='109' height='25'>발 행 구 분</td>
                      <td width='343'>
                      <%
    					IF ocashreceipt.FoneItem.Fuseopt = "0" THEN
							response.write "소비자 소득공제용"
						ELSE
    						response.write "사업자 지출증빙용"
    					END IF
					  %>
						</td>
                    </tr>
                    <tr>
                      <td height='1' colspan='3' align='center'  background='http://fiximage.10x10.co.kr/web2007/receipt/line.gif'></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <br>
           </td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><img src="http://fiximage.10x10.co.kr/web2007/receipt/bottom01.gif" width="632" height="13"></td>
  </tr>
</table>
</center></body>
</html>
<% end if %>
<%
set ocashreceipt = Nothing
%>

<% if Not receiptAlreadyExists then %>
<form name=frm method=post action="INIreceiptReq.asp">
<input type=hidden name=sitename value="<%= sitename %>">
<input type=hidden name=orderserial value="<%= orderserial %>">
</form>
<script language='javascript'>
frm.submit();
</script>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->