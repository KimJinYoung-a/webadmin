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
dim i, j
dim orderserial, hand
orderserial = request("orderserial")
hand        = request("hand")

orderserial = orderserial

dim sitename : sitename = "10x10"

dim myorder
set myorder = new CMyOrder
myorder.FRectSiteName = sitename
myorder.FRectOrderserial = orderserial
myorder.GetOneReceiptOrder

if (not myorder.FOrderExist) then
	response.write "<script>alert('결제 완료(확인) 된 건 및 \r\n2005년 1월 이후 주문건에 대해서만 영수증 발급가능합니다.');</script>"
	response.write "<script>window.close();</script>"
	response.end
end if

''발급가능날짜 - 최대 2달로 설정
dim availdate
availdate = dateAdd("m",-2,now())

if (myorder.FMasterItem.FRegDate<availdate) then
	response.write "<script>alert('최근 두달 이내 주문건에 대해서만 현금 영수증 발급가능합니다.');</script>"
	response.write "<script>window.close();</script>"
	response.end
end if

%>

<html>
<head>
<title>무통장 입금 현금영수증 발행</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style>
body, tr, td {font-size:9pt; font-family:굴림,verdana; color:#433F37; line-height:18px;}
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

<script language="javascript">

// 영수증 선택에 따른 분류
function RCP1(){
	document.ini.useopt.value="0" // 소비자 소득공제용
}

function RCP2(){
	document.ini.useopt.value="1" // 사업자 지출증빙용
}

var openwin;

function pay(frm)
{
	  // 필수항목 체크 (상품명, 상품가격, 구매자명, 구매자 이메일주소, 구매자 전화번호, 영수증 발행 용도)
<% if (hand="") then %>
	if(frm.useopt.value == "")
	{
		alert("현금영수증 발행용도를 선택하세요. 필수항목입니다.");
		return false;
	}
	else if(frm.useopt.value == "0")
	{
	    // 휴대폰도 가능.
		//if(frm.reg_num.value.length !=13){
		//	alert("소비자 소득공제용 영수증을 선택하셨습니다. 올바른 주민등록번호 13자리를 입력하세요.");
		//	return false;
		//}

		if(frm.reg_num.value.length !=13 && frm.reg_num.value.length !=10 && frm.reg_num.value.length !=11)
		{
			alert("주민등록번호 13자리 또는 올바른 휴대폰 번호 10자리(11자리)를 입력하세요.");
			frm.reg_num.focus();
			return false;
		}
		else if(frm.reg_num.value.length == 13)
		{
			var obj = frm.reg_num.value;
                	var sum=0;

                	for(i=0;i<8;i++) { sum+=obj.substring(i,i+1)*(i+2); }

                	for(i=8;i<12;i++) { sum+=obj.substring(i,i+1)*(i-6); }

                	sum=11-(sum%11);

                	if (sum>=10) { sum-=10; }

                	if (obj.substring(12,13) != sum || (obj.substring(6,7) !=1 && obj.substring(6,7) != 2))
                	{

                	    alert("주민등록번호에 오류가 있습니다. 다시 확인하십시오.");
                	    frm.reg_num.focus();
                	    return false;

	        	}

	        }
	        else if(frm.reg_num.value.length == 11 ||frm.reg_num.value.length == 10 )
	        {
	        	var obj = frm.reg_num.value;
	        	if (obj.substring(0,3)!= "011" && obj.substring(0,3)!= "017" && obj.substring(0,3)!= "016" && obj.substring(0,3)!= "018" && obj.substring(0,3)!= "019" && obj.substring(0,3)!= "010")
	        	{
	        		alert("주민등록번호 13자리 또는 올바른 휴대폰 번호 10자리(11자리)를 입력하세요. ");
	        		frm.reg_num.focus();
	        		return false;
	        	}

	        	var chr;
			for(var i=0; i<obj.length; i++){

	        		chr = obj.substr(i, 1);
	        		if( chr < '0' || chr > '9') {
   					alert("숫자가 아닌 문자가 휴대폰 번호에 추가되어 오류가 있습니다, 다시 확인 하십시오. ");
   					frm.reg_num.focus();
   					return false;
  				}
			}
	       }
	}else if(frm.useopt.value == "1"){
	    // 휴대폰도 가능.
		//if(frm.reg_num.value.length !=10){
		//	alert("사업자 지출증빙용 영수증을 선택하셨습니다. 올바른 사업자번호 10자리를 입력하세요.");
		//	return false;
		//}

		if(frm.reg_num.value.length !=10  && frm.reg_num.value.length !=11 && frm.reg_num.value.length !=13)
		{
			alert("올바른 주민등록번호 13자리, 사업자등록번호 10자리 또는 휴대폰 번호 10자리(11자리)를 입력하세요.");
			frm.reg_num.focus();
			return false;
		}
		else if(frm.reg_num.value.length == 13)
		{
			var obj = frm.reg_num.value;
            	var sum=0;

            	for(i=0;i<8;i++) { sum+=obj.substring(i,i+1)*(i+2); }

            	for(i=8;i<12;i++) { sum+=obj.substring(i,i+1)*(i-6); }

            	sum=11-(sum%11);

            	if (sum>=10) { sum-=10; }

            	if (obj.substring(12,13) != sum || (obj.substring(6,7) !=1 && obj.substring(6,7) != 2))
            	{

            	    alert("주민등록번호 13자리 또는 올바른 휴대폰 번호 10자리(11자리)를 입력하세요. ");
            	    frm.reg_num.focus();
            	    return false;

        	    }

        }
		else if(frm.reg_num.value.length == 10 && frm.reg_num.value.substring(0,1)!= "0"){
   			var vencod = frm.reg_num.value;
   			var sum = 0;
   			var getlist =new Array(10);
   			var chkvalue =new Array("1","3","7","1","3","7","1","3","5");
   			for(var i=0; i<10; i++) { getlist[i] = vencod.substring(i, i+1); }
   			for(var i=0; i<9; i++) { sum += getlist[i]*chkvalue[i]; }
   			sum = sum + parseInt((getlist[8]*5)/10);
   			sidliy = sum % 10;
   			sidchk = 0;
   			if(sidliy != 0) { sidchk = 10 - sidliy; }
   			else { sidchk = 0; }
   			if(sidchk != getlist[9]) {
   				alert("올바른 사업자 번호를 입력하시기 바랍니다. ");
   				frm.reg_num.focus();
   			    return false;
   			}
   			else
			{
			    //alert("number ok");
			    //return;
			}

		}
		else if(frm.reg_num.value.length == 11 ||frm.reg_num.value.length == 10 )
        {
        	var obj = frm.reg_num.value;
        	if (obj.substring(0,3)!= "011" && obj.substring(0,3)!= "017" && obj.substring(0,3)!= "016" && obj.substring(0,3)!= "018" && obj.substring(0,3)!= "019" && obj.substring(0,3)!= "010")
        	{
        		alert("실제 번호를 입력하시지 않아 실행에 실패하였습니다. 다시 입력하시기 바랍니다. ");
        		frm.reg_num.focus();
        		return false;
        	}

        	var chr;
		for(var i=0; i<obj.length; i++){

        		chr = obj.substr(i, 1);
        		if( chr < '0' || chr > '9') {
				alert("실제 번호를 입력하시지 않아 실행에 실패하였습니다. 다시 입력하시기 바랍니다. ");
				frm.reg_num.focus();
				return false;
			}
		}
       }
	}
<% end if %>
	var sum_price = eval(frm.sup_price.value) + eval(frm.tax.value) + eval(frm.srvc_price.value);
	if(frm.cr_price.value != sum_price){
		alert("총액은 공급가+부가세+봉사료입니다.더한 금액이 틀립니다");
		return false;
	}
	if(frm.cr_price.value < 100){
		alert("현금영수증 발행시 최소금액은 100 원입니다.");
		return false;
	}

	if(frm.goodname.value == "")
	{
		alert("상품명이 빠졌습니다. 필수항목입니다.");
		return false;
	}
	else if(frm.cr_price.value == "")
	{
		alert("현금결제금액이 빠졌습니다. 필수항목입니다.");
		return false;
	}
	else if(frm.sup_price.value == "")
	{
		alert("공급가액이 빠졌습니다. 필수항목입니다.");
		return false;
	}
	else if(frm.tax.value == "")
	{
		alert("부가세가 빠졌습니다. 필수항목입니다.");
		return false;
	}
	else if(frm.srvc_price.value == "")
	{
		alert("봉사료가 빠졌습니다. 필수항목입니다.");
		return false;
	}
	else if(frm.buyername.value == "")
	{
		alert("구매자명이 빠졌습니다. 필수항목입니다.");
		return false;
	}
	else if(frm.reg_num.value == "")
	{
		alert("주민등록번호(또는 사업자번호)가 빠졌습니다. 필수항목입니다.");
		return false;

	}
	else if(frm.buyeremail.value == "")
	{
		alert("구매자 이메일주소가 빠졌습니다. 필수항목입니다.");
		return false;
	}
	else if(frm.buyertel.value == "")
	{
		alert("구매자 전화번호가 빠졌습니다. 필수항목입니다.");
		return false;
	}

	// 더블클릭으로 인한 중복요청을 방지하려면 반드시 confirm()을
	// 사용하십시오.

	if(confirm("현금영수증을 발행하시겠습니까?"))
	{
		disable_click();
		openwin = window.open("childwin.html","childwin","width=299,height=149");
		return true;
	}
	else
	{
		return false;
	}
}


// 영수증 발행용도 리스트 보이기

var main_cnt = 1

function showhide(num){

    for (i=1; i<=main_cnt; i++){

      menu=eval("document.all.block"+i+".style");

      if (num == i){

      	if (menu.display == "block") {

        	menu.display="none";
        }
        else{
        	menu.display="block";
        }

     }
     else{

     	menu.display="none";
     }
   }
}



function enable_click(){
	document.ini.clickcontrol.value = "enable"
}

function disable_click(){
	document.ini.clickcontrol.value = "disable"
}

function focus_control(){
	if(document.ini.clickcontrol.value == "disable")
		openwin.focus();
}

function reCalcuTax(){
    var frm = document.ini;
    var cr_price = frm.cr_price.value;
    var sup_price = parseInt(cr_price*10/11);
    var tax       = cr_price*1-sup_price*1;

    frm.sup_price.value = sup_price;
     frm.tax.value = tax;
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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
</head>

<!-----------------------------------------------------------------------------------------------------
※ 주의 ※
 아래의 body TAG의 내용중에
 onload="javascript:enable_click()" onFocus="javascript:focus_control()" 이 부분은 수정없이 그대로 사용.
 아래의 form TAG내용도 수정없이 그대로 사용.
------------------------------------------------------------------------------------------------------->

<body bgcolor="#FFFFFF" text="#242424" leftmargin=0 topmargin=0 marginwidth=0 marginheight=0 bottommargin=0 rightmargin=0 onload="javascript:enable_click()" onFocus="javascript:focus_control()"><center>
<form name=ini method=post action="INIreceipt.asp" onSubmit="return pay(this)">
<input type=hidden name=goodname value="<%= myorder.GetGoodsName %>">
<% if (hand="") then %>
<input type=hidden name=cr_price value="<%= myorder.FMasterItem.FsubtotalPrice %>">
<% end if %>
<input type=hidden name=sup_price value="<%= myorder.FMasterItem.GetSuppPrice %>">
<input type=hidden name=tax value="<%= myorder.FMasterItem.GetTaxPrice %>">
<input type=hidden name=srvc_price value="0">
<input type=hidden name=buyername value="<%= myorder.FMasterItem.FBuyName %>">
<input type=hidden name=orderserial value="<%= orderserial %>">
<input type=hidden name=userid value="<%= myorder.FMasterItem.FUserID %>">
<input type=hidden name=sitename value="<%= sitename %>">
<input type=hidden name=paymethod value="<%= myorder.FMasterItem.FAccountDiv %>">

<input type=hidden name=buyertel value="000-000-0000">

<table width="632" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="85" background="http://fiximage.10x10.co.kr/web2007/receipt/cash_top.gif" style="padding:0 0 0 64">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="3%" valign="top"><img src="http://fiximage.10x10.co.kr/web2007/receipt/title_01.gif" width="8" height="27" vspace="5"></td>
          <td width="97%" height="40" class="pl_03"><font color="#FFFFFF"><b>현금영수증 발행요청</b></font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td align="center" bgcolor="6095BC"><table width="620" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td bgcolor="#FFFFFF" style="padding:2 0 0 56">
            <table width="510" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="7"><img src="http://fiximage.10x10.co.kr/web2007/receipt/life.gif" width="7" height="30"></td>
                <td background="http://fiximage.10x10.co.kr/web2007/receipt/center.gif"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon03.gif" width="12" height="10">
                  <b>정보를 기입하신 후 발행버튼을 눌러주십시오.</b></td>
                <td width="8"><img src="http://fiximage.10x10.co.kr/web2007/receipt/right.gif" width="8" height="30"></td>
              </tr>
            </table>
            <table width="510" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="10" colspan="2"></td>
              </tr>
              <tr>
                <td width="510" colspan="2"  style="padding:0 0 0 23">
                <table width="470" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="26">상 품 명</td>
                      <td width="320"><%= myorder.GetGoodsName %></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <% if (hand<>"") then %>
                      <td width="180" height="26">발행금액</td>
                      <td width="320"><input type="text" name="cr_price" value="<%= myorder.FMasterItem.FsubtotalPrice %>" size="10" maxlength=9 onKeyUp="reCalcuTax()">
                      <% else %>
                      <td width="180" height="26">결제한금액</td>
                      <td width="320"><%= Formatnumber(myorder.FMasterItem.FsubtotalPrice,0) %>
                      <% end if %>
<!--                      &nbsp;&nbsp;<font color=red>(현금으로 결제한 총금액:공급가+부가세)</font>  -->
                      </td>
                    </tr>
<!--
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="26">공 급 가 액</td>
                      <td width="320"><%= FormatNumber(myorder.FMasterItem.GetSuppPrice,0) %></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="26">부 가 세</td>
                      <td width="320"><%= FormatNumber(myorder.FMasterItem.GetTaxPrice,0) %></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
-->
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="25">구 매 자 명</td>
                      <td width="343"><%= myorder.FMasterItem.FBuyName %></td>
                    </tr>
 <!--
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
-->
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="25">전 자 우 편</td>
                      <td width="343"><input type=text name=buyeremail size=20 value="<%= myorder.FMasterItem.FBuyEmail %>"></td>
                    </tr>
<!--
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="109" height="25">이 동 전 화</td>
                      <td width="343"><input type=text name=buyertel size=20 value="<%= myorder.FMasterItem.FBuyHp %>"></td>
                    </tr>
 -->
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="25">주민등록번호<br>(사업자번호,휴대폰번호)</td>
                      <td width="343">
                      <% if (hand="") then %>
                      <input type=text name=reg_num size=13 maxlength=13 value="">
                      <% else %>
                      <input type=text name=reg_num size=18 maxlength=18 value="">
                      <% end if %>
                      &nbsp;&nbsp;&nbsp;<font color=red>"-"를 뺀 숫자만 입력하세요</font></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td colspan=2>현금영수증 발행용도를 선택하세요</td>
                    </tr>
                    <tr>
                      <td colspan=3>
                    	 <table width=100% cellspacing=0 cellpadding=0 border=0>
                    	   <tr>
                    	     <td align=center>
                    		     <input type=radio checked name=choose value=1 Onclick= "javascript:RCP1()">소비자 소득공제용&nbsp;&nbsp;&nbsp;&nbsp;
				     			 <input type=radio name=choose value=1 Onclick= "javascript:RCP2()">사업자 지출증빙용
                    	     </td>
                    	   </tr>
                    	 </table>
                      </td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td colspan="3" >
                        ● 소득 공제용 - 주민등록 번호와 휴대폰 번호로 발급 가능<br>
                        ● 지출 증빙용 - 주민등록 번호, 휴대폰 번호, 사업자 번호로 발급 가능<br>
                      </td>
                    </tr>
                    <tr valign="bottom">
                      <td height="40" colspan="3" align="center"><input type=image src="http://fiximage.10x10.co.kr/web2007/receipt/button_08.gif" width="63" height="25"></td>
                    </tr>
<!--
                    <tr valign="bottom">
                      <td height="45" colspan="3">전자우편과 이동전화번호를 입력받는 것은 영수증 발행 성공 내역을 E-MAIL 또는 SMS 로
                   알려드리기 위함이오니 반드시 기입하시기 바랍니다.</td>
                    </tr>
-->
                  </table></td>
              </tr>
            </table>
            <br>
          </td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><img src="http://fiximage.10x10.co.kr/web2007/receipt/BOTTOM01.GIF" width="632" height="13"></td>
  </tr>
</table>
</center>

<!--
상점아이디.
테스트를 마친 후, 발급받은 아이디로 바꾸어 주십시오.
-->
<% if (application("Svr_Info")	= "Dev") then %>
<input type=hidden name=mid value="INIpayTest">
<% else %>
<input type=hidden name=mid value="teenxteen4">
<% end if %>

<!--
UID.
테스트를 마친후, 발급받은 상점아이디로 바꾸어 주십시오.
(반드시 mid와 동일한 값을 입력)
-->
<input type=hidden name=uid value="">

<!--
화폐단위
WON 또는 CENT
주의 : 미화승인은 별도 계약이 필요합니다.
-->
<input type=hidden name=currency value="WON">

<!-- 삭제/수정 불가 -->
<input type=hidden name=clickcontrol value="">
<input type=hidden name=useopt value="0">

</form>
</body>
</html>
<%
set myorder = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->