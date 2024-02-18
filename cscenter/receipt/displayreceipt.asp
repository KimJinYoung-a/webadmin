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
<!-- #include virtual="/lib/classes/cscenter/cs_cashreceiptcls.asp"-->

<%
dim idx
idx = session("lastreceiptidx")

dim ocashreceipt
set ocashreceipt = new CCashReceipt
ocashreceipt.FRectIdx = idx
ocashreceipt.GetOneCashReceipt


%>
<html>
<head>
<title>무통장 입금 현금영수증 발행</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
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

	var openwin=window.open("childwin.html","childwin","width=299,height=149");
	openwin.close();
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
    <td height="83" background="<%
    			IF ocashreceipt.FoneItem.FResultCode = "00" THEN 		'성공인 경우
					response.write "http://fiximage.10x10.co.kr/web2007/receipt/cash_top.gif"
				ELSE
    				response.write "http://fiximage.10x10.co.kr/web2007/receipt/spool_top.gif"
    			END IF
				%>"style="padding:0 0 0 64">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="3%" valign="top"><img src="http://fiximage.10x10.co.kr/web2007/receipt/title_01.gif" width="8" height="27" vspace="5"></td>
          <td width="97%" height="40" class="pl_03"><font color="#FFFFFF"><b>현금결제 영수증 발급결과</b></font></td>
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
                <% IF ocashreceipt.FoneItem.FResultCode = "00" THEN %>
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
                      <td width="109" height="25">결 과 내 용</td>
                      <td width="343"><%= ocashreceipt.FoneItem.Fresultmsg %></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
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
                      <td height='1' colspan='3' align='center'  background='http://fiximage.10x10.co.kr/web2007/receipt/line.gif'></td>
                    </tr>
                    <tr>
                      <td width='18' align='center'><img src='img/icon02.gif' width='7' height='7'></td>
                      <td width='109' height='25'>총 현금결제금액</td>
                      <td width='343'><%= ocashreceipt.FoneItem.Fcr_price %></td>
                    </tr>

                    <tr>
                      <td height='1' colspan='3' align='center'  background='http://fiximage.10x10.co.kr/web2007/receipt/line.gif'></td>
                    </tr>
                    <tr>
                      <td width='18' align='center'><img src='img/icon02.gif' width='7' height='7'></td>
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
              <% else %>
              <table width="470" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="109" height="26">발 급 결 과</td>
                      <td width="343"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><font color="red">실패</font></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="109" height="25">결 과 내 용</td>
                      <td width="343"><%= ocashreceipt.FoneItem.Fresultmsg %></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td colspan=3 align=center><a href="javascript:history.back();"><font color="blue">[재시도]</font></a></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
              </table>
              <% end if %>
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
  <tr>
    <td align="center"><input type="button" value="닫기" onClick="self.close();"></td>
  </tr>
</table>
</center></body>
</html>
<%
set ocashreceipt = Nothing
%>
<script >
opener.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
