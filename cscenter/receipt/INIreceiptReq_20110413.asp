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
	response.write "<script>alert('���� �Ϸ�(Ȯ��) �� �� �� \r\n2005�� 1�� ���� �ֹ��ǿ� ���ؼ��� ������ �߱ް����մϴ�.');</script>"
	response.write "<script>window.close();</script>"
	response.end
end if

''�߱ް��ɳ�¥ - �ִ� 2�޷� ����
dim availdate
availdate = dateAdd("m",-2,now())

if (myorder.FMasterItem.FRegDate<availdate) then
	response.write "<script>alert('�ֱ� �δ� �̳� �ֹ��ǿ� ���ؼ��� ���� ������ �߱ް����մϴ�.');</script>"
	response.write "<script>window.close();</script>"
	response.end
end if

%>

<html>
<head>
<title>������ �Ա� ���ݿ����� ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style>
body, tr, td {font-size:9pt; font-family:����,verdana; color:#433F37; line-height:18px;}
table, img {border:none}

/* Padding ******/
.pl_01 {padding:1 10 0 10; line-height:19px;}
.pl_03 {font-size:20pt; font-family:����,verdana; color:#FFFFFF; line-height:29px;}

/* Link ******/
.a:link  {font-size:9pt; color:#333333; text-decoration:none}
.a:visited { font-size:9pt; color:#333333; text-decoration:none}
.a:hover  {font-size:9pt; color:#0174CD; text-decoration:underline}

.txt_03a:link  {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:visited {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:hover  {font-size: 8pt;line-height:18px;color:#EC5900; text-decoration:underline}
</style>

<script language="javascript">

// ������ ���ÿ� ���� �з�
function RCP1(){
	document.ini.useopt.value="0" // �Һ��� �ҵ������
}

function RCP2(){
	document.ini.useopt.value="1" // ����� ����������
}

var openwin;

function pay(frm)
{
	  // �ʼ��׸� üũ (��ǰ��, ��ǰ����, �����ڸ�, ������ �̸����ּ�, ������ ��ȭ��ȣ, ������ ���� �뵵)
<% if (hand="") then %>
	if(frm.useopt.value == "")
	{
		alert("���ݿ����� ����뵵�� �����ϼ���. �ʼ��׸��Դϴ�.");
		return false;
	}
	else if(frm.useopt.value == "0")
	{
	    // �޴����� ����.
		//if(frm.reg_num.value.length !=13){
		//	alert("�Һ��� �ҵ������ �������� �����ϼ̽��ϴ�. �ùٸ� �ֹε�Ϲ�ȣ 13�ڸ��� �Է��ϼ���.");
		//	return false;
		//}

		if(frm.reg_num.value.length !=13 && frm.reg_num.value.length !=10 && frm.reg_num.value.length !=11)
		{
			alert("�ֹε�Ϲ�ȣ 13�ڸ� �Ǵ� �ùٸ� �޴��� ��ȣ 10�ڸ�(11�ڸ�)�� �Է��ϼ���.");
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

                	    alert("�ֹε�Ϲ�ȣ�� ������ �ֽ��ϴ�. �ٽ� Ȯ���Ͻʽÿ�.");
                	    frm.reg_num.focus();
                	    return false;

	        	}

	        }
	        else if(frm.reg_num.value.length == 11 ||frm.reg_num.value.length == 10 )
	        {
	        	var obj = frm.reg_num.value;
	        	if (obj.substring(0,3)!= "011" && obj.substring(0,3)!= "017" && obj.substring(0,3)!= "016" && obj.substring(0,3)!= "018" && obj.substring(0,3)!= "019" && obj.substring(0,3)!= "010")
	        	{
	        		alert("�ֹε�Ϲ�ȣ 13�ڸ� �Ǵ� �ùٸ� �޴��� ��ȣ 10�ڸ�(11�ڸ�)�� �Է��ϼ���. ");
	        		frm.reg_num.focus();
	        		return false;
	        	}

	        	var chr;
			for(var i=0; i<obj.length; i++){

	        		chr = obj.substr(i, 1);
	        		if( chr < '0' || chr > '9') {
   					alert("���ڰ� �ƴ� ���ڰ� �޴��� ��ȣ�� �߰��Ǿ� ������ �ֽ��ϴ�, �ٽ� Ȯ�� �Ͻʽÿ�. ");
   					frm.reg_num.focus();
   					return false;
  				}
			}
	       }
	}else if(frm.useopt.value == "1"){
	    // �޴����� ����.
		//if(frm.reg_num.value.length !=10){
		//	alert("����� ���������� �������� �����ϼ̽��ϴ�. �ùٸ� ����ڹ�ȣ 10�ڸ��� �Է��ϼ���.");
		//	return false;
		//}

		if(frm.reg_num.value.length !=10  && frm.reg_num.value.length !=11 && frm.reg_num.value.length !=13)
		{
			alert("�ùٸ� �ֹε�Ϲ�ȣ 13�ڸ�, ����ڵ�Ϲ�ȣ 10�ڸ� �Ǵ� �޴��� ��ȣ 10�ڸ�(11�ڸ�)�� �Է��ϼ���.");
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

            	    alert("�ֹε�Ϲ�ȣ 13�ڸ� �Ǵ� �ùٸ� �޴��� ��ȣ 10�ڸ�(11�ڸ�)�� �Է��ϼ���. ");
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
   				alert("�ùٸ� ����� ��ȣ�� �Է��Ͻñ� �ٶ��ϴ�. ");
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
        		alert("���� ��ȣ�� �Է��Ͻ��� �ʾ� ���࿡ �����Ͽ����ϴ�. �ٽ� �Է��Ͻñ� �ٶ��ϴ�. ");
        		frm.reg_num.focus();
        		return false;
        	}

        	var chr;
		for(var i=0; i<obj.length; i++){

        		chr = obj.substr(i, 1);
        		if( chr < '0' || chr > '9') {
				alert("���� ��ȣ�� �Է��Ͻ��� �ʾ� ���࿡ �����Ͽ����ϴ�. �ٽ� �Է��Ͻñ� �ٶ��ϴ�. ");
				frm.reg_num.focus();
				return false;
			}
		}
       }
	}
<% end if %>
	var sum_price = eval(frm.sup_price.value) + eval(frm.tax.value) + eval(frm.srvc_price.value);
	if(frm.cr_price.value != sum_price){
		alert("�Ѿ��� ���ް�+�ΰ���+������Դϴ�.���� �ݾ��� Ʋ���ϴ�");
		return false;
	}
	if(frm.cr_price.value < 100){
		alert("���ݿ����� ����� �ּұݾ��� 100 ���Դϴ�.");
		return false;
	}

	if(frm.goodname.value == "")
	{
		alert("��ǰ���� �������ϴ�. �ʼ��׸��Դϴ�.");
		return false;
	}
	else if(frm.cr_price.value == "")
	{
		alert("���ݰ����ݾ��� �������ϴ�. �ʼ��׸��Դϴ�.");
		return false;
	}
	else if(frm.sup_price.value == "")
	{
		alert("���ް����� �������ϴ�. �ʼ��׸��Դϴ�.");
		return false;
	}
	else if(frm.tax.value == "")
	{
		alert("�ΰ����� �������ϴ�. �ʼ��׸��Դϴ�.");
		return false;
	}
	else if(frm.srvc_price.value == "")
	{
		alert("����ᰡ �������ϴ�. �ʼ��׸��Դϴ�.");
		return false;
	}
	else if(frm.buyername.value == "")
	{
		alert("�����ڸ��� �������ϴ�. �ʼ��׸��Դϴ�.");
		return false;
	}
	else if(frm.reg_num.value == "")
	{
		alert("�ֹε�Ϲ�ȣ(�Ǵ� ����ڹ�ȣ)�� �������ϴ�. �ʼ��׸��Դϴ�.");
		return false;

	}
	else if(frm.buyeremail.value == "")
	{
		alert("������ �̸����ּҰ� �������ϴ�. �ʼ��׸��Դϴ�.");
		return false;
	}
	else if(frm.buyertel.value == "")
	{
		alert("������ ��ȭ��ȣ�� �������ϴ�. �ʼ��׸��Դϴ�.");
		return false;
	}

	// ����Ŭ������ ���� �ߺ���û�� �����Ϸ��� �ݵ�� confirm()��
	// ����Ͻʽÿ�.

	if(confirm("���ݿ������� �����Ͻðڽ��ϱ�?"))
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


// ������ ����뵵 ����Ʈ ���̱�

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
�� ���� ��
 �Ʒ��� body TAG�� �����߿�
 onload="javascript:enable_click()" onFocus="javascript:focus_control()" �� �κ��� �������� �״�� ���.
 �Ʒ��� form TAG���뵵 �������� �״�� ���.
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
          <td width="97%" height="40" class="pl_03"><font color="#FFFFFF"><b>���ݿ����� �����û</b></font></td>
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
                  <b>������ �����Ͻ� �� �����ư�� �����ֽʽÿ�.</b></td>
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
                      <td width="180" height="26">�� ǰ ��</td>
                      <td width="320"><%= myorder.GetGoodsName %></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <% if (hand<>"") then %>
                      <td width="180" height="26">����ݾ�</td>
                      <td width="320"><input type="text" name="cr_price" value="<%= myorder.FMasterItem.FsubtotalPrice %>" size="10" maxlength=9 onKeyUp="reCalcuTax()">
                      <% else %>
                      <td width="180" height="26">�����ѱݾ�</td>
                      <td width="320"><%= Formatnumber(myorder.FMasterItem.FsubtotalPrice,0) %>
                      <% end if %>
<!--                      &nbsp;&nbsp;<font color=red>(�������� ������ �ѱݾ�:���ް�+�ΰ���)</font>  -->
                      </td>
                    </tr>
<!--
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="26">�� �� �� ��</td>
                      <td width="320"><%= FormatNumber(myorder.FMasterItem.GetSuppPrice,0) %></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="26">�� �� ��</td>
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
                      <td width="180" height="25">�� �� �� ��</td>
                      <td width="343"><%= myorder.FMasterItem.FBuyName %></td>
                    </tr>
 <!--
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
-->
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="25">�� �� �� ��</td>
                      <td width="343"><input type=text name=buyeremail size=20 value="<%= myorder.FMasterItem.FBuyEmail %>"></td>
                    </tr>
<!--
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="109" height="25">�� �� �� ȭ</td>
                      <td width="343"><input type=text name=buyertel size=20 value="<%= myorder.FMasterItem.FBuyHp %>"></td>
                    </tr>
 -->
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td width="180" height="25">�ֹε�Ϲ�ȣ<br>(����ڹ�ȣ,�޴�����ȣ)</td>
                      <td width="343">
                      <% if (hand="") then %>
                      <input type=text name=reg_num size=13 maxlength=13 value="">
                      <% else %>
                      <input type=text name=reg_num size=18 maxlength=18 value="">
                      <% end if %>
                      &nbsp;&nbsp;&nbsp;<font color=red>"-"�� �� ���ڸ� �Է��ϼ���</font></td>
                    </tr>
                    <tr>
                      <td height="1" colspan="3" align="center"  background="http://fiximage.10x10.co.kr/web2007/receipt/line.gif"></td>
                    </tr>
                    <tr>
                      <td width="18" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/icon02.gif" width="7" height="7"></td>
                      <td colspan=2>���ݿ����� ����뵵�� �����ϼ���</td>
                    </tr>
                    <tr>
                      <td colspan=3>
                    	 <table width=100% cellspacing=0 cellpadding=0 border=0>
                    	   <tr>
                    	     <td align=center>
                    		     <input type=radio checked name=choose value=1 Onclick= "javascript:RCP1()">�Һ��� �ҵ������&nbsp;&nbsp;&nbsp;&nbsp;
				     			 <input type=radio name=choose value=1 Onclick= "javascript:RCP2()">����� ����������
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
                        �� �ҵ� ������ - �ֹε�� ��ȣ�� �޴��� ��ȣ�� �߱� ����<br>
                        �� ���� ������ - �ֹε�� ��ȣ, �޴��� ��ȣ, ����� ��ȣ�� �߱� ����<br>
                      </td>
                    </tr>
                    <tr valign="bottom">
                      <td height="40" colspan="3" align="center"><input type=image src="http://fiximage.10x10.co.kr/web2007/receipt/button_08.gif" width="63" height="25"></td>
                    </tr>
<!--
                    <tr valign="bottom">
                      <td height="45" colspan="3">���ڿ���� �̵���ȭ��ȣ�� �Է¹޴� ���� ������ ���� ���� ������ E-MAIL �Ǵ� SMS ��
                   �˷��帮�� �����̿��� �ݵ�� �����Ͻñ� �ٶ��ϴ�.</td>
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
�������̵�.
�׽�Ʈ�� ��ģ ��, �߱޹��� ���̵�� �ٲپ� �ֽʽÿ�.
-->
<% if (application("Svr_Info")	= "Dev") then %>
<input type=hidden name=mid value="INIpayTest">
<% else %>
<input type=hidden name=mid value="teenxteen4">
<% end if %>

<!--
UID.
�׽�Ʈ�� ��ģ��, �߱޹��� �������̵�� �ٲپ� �ֽʽÿ�.
(�ݵ�� mid�� ������ ���� �Է�)
-->
<input type=hidden name=uid value="">

<!--
ȭ�����
WON �Ǵ� CENT
���� : ��ȭ������ ���� ����� �ʿ��մϴ�.
-->
<input type=hidden name=currency value="WON">

<!-- ����/���� �Ұ� -->
<input type=hidden name=clickcontrol value="">
<input type=hidden name=useopt value="0">

</form>
</body>
</html>
<%
set myorder = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->