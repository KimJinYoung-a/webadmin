<%
dim protocol : protocol = "http"
If request.serverVariables("SERVER_PORT") = "443" Then protocol = "https"
%>     
<html>
<head>
    <title>LG������ ���ڰ��� �ſ�ī����� ���� ������</title>
<!-- ���� �Ұ�  ����-->
	<script language="javascript" src="<%= protocol %>://www.vpay.co.kr/KVPplugin.js" type="text/javascript"></script>
	<script language="javascript" src="<%= protocol %>://mpi.dacom.net/XPayMPI/XPayMPIOCX.js" type="text/javascript"></script>
	<script language="javascript">
		function XPayMPIProcessResult(Step1ControlResult, Step1ServerResult, Step1Msg, Step2ControlResult, Step2ServerResult, Step2Msg,  Cavv, MD, ECI, CAVVALG ) { }
		
	</script>
<!-- //���� �Ұ�  �� -->
	<script language="javascript">
		/*
		* ������û ó�� 
		*/
		function doPay_ActiveX(){
			ret = xpay_card(document.getElementById('LGD_PAYINFO'));
			if( "0000" == ret ) { //��������
				/*
				* �������� ȭ�� ó��
				*/
				alert("������û �մϴ�...")
				document.getElementById("LGD_PAYINFO").submit();        
			} else { //��������
				/*
				* �������� ȭ�� ó��
				*/        
			}
		}
	</script>
</head>
<body>
	<div id="LGD_ACTIVEX_DIV"></div><!-- ActiveX��ġ ���н� �ȳ� Layer -->
	<script language="javascript" src="<%= protocol %>://xpay.lgdacom.net/xpay/js/xpay_card.js" type="text/javascript"></script><!-- ActiveX��ġ �� �����Լ� -->
	
    <form method="post" id="LGD_PAYINFO" action="CardApp.asp">
        <table>	
<!-- ������������ ���� -->		
            <tr>
                <td>�������̵�(t�� ������ ���̵�) </td>
                <td><input type="text" name="CST_MID" id="CST_MID" value="youareagirl"/></td>
            </tr>
            <tr>
                <td>����,�׽�Ʈ </td>
                <td><input type="text" name="CST_PLATFORM" id = "CST_PLATFORM" value="test"/></td>
            </tr>
            <tr>
                <td>�ֹ���ȣ(����ũ) </td>
                <td><input type="text" name="LGD_OID" id="LGD_OID" value=""/></td>
            </tr>
			<tr>
                <td>�����ݾ� </td>
                <td><input type="text" name="LGD_AMOUNT" id="LGD_AMOUNT" value="1000"/></td>
            </tr>
			<tr>
                <td>�����ڸ�</td>
                <td><input type="text" name="LGD_BUYER" id="LGD_BUYER" value="�׽�Ʈ������"/></td>
            </tr>
            <tr>
                <td>��ǰ��</td>
                <td><input type="text" name="LGD_PRODUCTINFO" id="LGD_PRODUCTINFO" value="�׽�Ʈ��ǰ"/></td>
            </tr>
			<tr>
                <td>������</td>
                <td><input type="text" name="LGD_MERTNAME" id="LGD_MERTNAME" value="�׽�Ʈ����"/></td>
            </tr>
            <tr>
                <td>�̸����ּ�</td>
                <td><input type="text" name="LGD_BUYEREMAIL" id="LGD_BUYEREMAIL" value=""/></td>
            </tr>
			<tr>
                <td>ī�弱��</td>
                <td>
                	<select name="LGD_CARDTYPE" id="LGD_CARDTYPE" onchange="setLGD_NOINT()">
                		<option value="XX">����</option>
                		<option value="11">����(KB)</option>
						<option value="21">��ȯ</option>
						<option value="31">��(BC)</option>
						<option value="41">����</option>
						<option value="51">�Ｚ</option>
						<option value="61">����</option>
						<option value="71">�Ե�</option>
						<option value="36">��Ƽ</option>
						<option value="32">�ϳ�</option>
						<option value="33">�츮</option>
						<option value="42">����</option>
						<option value="34">����</option>
						<option value="35">����</option>
						<option value="46">����</option>
						<option value="29">����ĳ��Ż</option>
						<option value="4V">�ؿ�VISA</option>
						<option value="4M">�ؿ�MASTER</option>
						<option value="4J">�ؿ�JCB</option>
						<option value="6D">�ؿ�DINERS CLUB</option>
						<option value="91">NH</option>
                	</select>
                </td>
            </tr>
			<tr>
                <td>ISP����������(KVP_NOINT_INF)</td>
                <td><input type="text" name="KVP_NOINT_INF" id="KVP_NOINT_INF" value="0100-2:3,0204-2:3"/>��)����/��/�츮 2,3���� ������ �� --> 0100-2:3,0204-2:3,0700-2:3</td>
            </tr>
			<tr>
                <td>ISPǥ���Һΰ�����(KVP_QUOTA_INF)</td>
                <td><input type="text" name="KVP_QUOTA_INF" id="KVP_QUOTA_INF" value="0:2:3:4:5:6:7:8:9:10:11:12"/></td>
            </tr>
			<tr>
                <td>�Ƚ�Ŭ�� ����������(LGD_NOINTINF)</td>
                <td><input type="text" name="LGD_NOINTINF" id="LGD_NOINTINF" value=""/>��)����/�Ｚ 2,3���� ������ �� --> 41-2:3,51-2:3</td>
            </tr>
			<tr>
                <td>�Ƚ�Ŭ�� �Һΰ���(LGD_INSTALL)</td>
                <td>
                	<select name="LGD_INSTALL" onchange="setLGD_NOINT()">
                		<option value="0">�Ͻú�</option>
						<option value="2">2����</option>
						<option value="3">3����</option>
						<option value="4">4����</option>
						<option value="5">5����</option>
						<option value="6">6����</option>
						<option value="7">7����</option>
						<option value="8">8����</option>
						<option value="9">9����</option>
						<option value="10">10����</option>
						<option value="11">11����</option>
						<option value="12">12����</option>
                	</select>
				</td>
            </tr>

<!--//������������ ��  -->
<!-- ���� �Ұ�  ����-->		
            <tr>
                <td>�������� </td>
                <td><input type="text" name="LGD_AUTHTYPE" id="LGD_AUTHTYPE"/></td>
            </tr>
			<tr>
                <td>KVP_CURRENCY</td>
                <td><input type="text" name="KVP_CURRENCY" id="KVP_CURRENCY"/></td>
            </tr>
			<tr>
                <td>KVP_OACERT_INF</td>
                <td><input type="text" name="KVP_OACERT_INF" id="KVP_OACERT_INF"/></td>
            </tr>	
			<tr>
                <td>KVP_RESERVED1</td>
                <td><input type="text" name="KVP_RESERVED1" id="KVP_RESERVED1"/></td>
            </tr>						
			<tr>
                <td>KVP_RESERVED2</td>
                <td><input type="text" name="KVP_RESERVED2" id="KVP_RESERVED2"/></td>
            </tr>						
			<tr>
                <td>KVP_RESERVED3</td>
                <td><input type="text" name="KVP_RESERVED3" id="KVP_RESERVED3"/></td>
            </tr>	
			<tr>
                <td>KVP_GOODNAME</td>
                <td><input type="text" name="KVP_GOODNAME" id="KVP_GOODNAME"/></td>
            </tr>				
			<tr>
                <td>KVP_CARDCOMPANY</td>
                <td><input type="text" name="KVP_CARDCOMPANY" id="KVP_CARDCOMPANY"/></td>
            </tr>			
			<tr>
                <td>KVP_PRICE</td>
                <td><input type="text" name="KVP_PRICE" id="KVP_PRICE"/></td>
            </tr>			
			<tr>
                <td>KVP_PGID</td>
                <td><input type="text" name="KVP_PGID" id="KVP_PGID"/></td>
            </tr>			    
			<tr>
                <td>KVP_QUOTA</td>
                <td><input type="text" name="KVP_QUOTA" id="KVP_QUOTA" /></td>
            </tr>
            <tr>
                <td>KVP_NOINT </td>
                <td><input type="text" name="KVP_NOINT" id="KVP_NOINT"/></td> 
			<tr>
            <tr>
                <td>KVP_SESSIONKEY </td>
                <td><input type="hidden" name="KVP_SESSIONKEY" id="KVP_SESSIONKEY" /></td> 
			<tr>
			<tr>
                <td>KVP_ENCDATA</td>
                <td><input type="hidden" name="KVP_ENCDATA" id="KVP_ENCDATA"/></td> 
			<tr>
			<tr>
                <td>KVP_CARDCODE</td>
                <td><input type="text" name="KVP_CARDCODE" id="KVP_CARDCODE"/></td> 
			<tr>
			<tr>
                <td>KVP_CONAME</td>
                <td><input type="text" name="KVP_CONAME" id="KVP_CONAME"/></td> 
			<tr>
            <tr>
                <td>LGD_PAN </td>
                <td><input type="text" name="LGD_PAN" /></td>
            </tr>
            <tr>
                <td>LGD_NOINT</td>
                <td><input type="text" name="LGD_NOINT" id="LGD_NOINT" /></td>
            </tr>
            <tr>
                <td>VBV_ECI</td>
                <td><input type="text" name="VBV_ECI" id="VBV_ECI"/></td>
            </tr>
            <tr>
                <td>VBV_CAVV</td>
                <td><input type="text" name="VBV_CAVV" id="VBV_CAVV" /></td>
            </tr>
            <tr>
                <td>VBV_XID</td>
                <td><input type="text" name="VBV_XID" id="VBV_XID" /></td>
            </tr>
            <tr>
                <td>LGD_EXPYEAR</td>
                <td><input type="text" name="LGD_EXPYEAR" id="LGD_EXPYEAR" /></td>
            </tr>
            <tr>
                <td>LGD_EXPMON</td>
                <td><input type="text" name="LGD_EXPMON" id="LGD_EXPMON" /></td>
            </tr>
<!-- //���� �Ұ�  �� -->
      </table>
		<input type="button" value="�����ϱ�" onclick="doPay_ActiveX()"/><br/>
    </form>

</body>
</html>
