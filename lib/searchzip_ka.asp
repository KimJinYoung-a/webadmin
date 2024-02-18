<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"

'###########################################################
' Description :  �����ȣ ã��(īī�� API)
' History : 2019.06.13 ������ ����Ʈ ����
'           2022.09.28 �ѿ�� ����Ʈ ���� ���� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<%
	dim fiximgPath
	'�̹��� ��� ����(SSL ó��)
	if request.ServerVariables("SERVER_PORT_SECURE")<>1 then
		fiximgPath = "http://fiximage.10x10.co.kr"
	else
		fiximgPath = "/fiximage"
	end If
	
	Dim strTarget
	Dim strMode
	strTarget	= requestCheckVar(Request("target"),32)
	strMode     = requestCheckVar(Request("strMode"),32)

%>
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<title>�ٹ����� 10X10 : �����ȣã��</title>
<link rel="stylesheet" type="text/css" href="/css/frontCopy/default.css" />
<link rel="stylesheet" type="text/css" href="/css/frontCopy/preVst/common_ssl.css" />
<link rel="stylesheet" type="text/css" href="/css/frontCopy/preVst/mytenten_ssl.css" />
<link rel="stylesheet" type="text/css" href="/css/frontCopy/preVst/popup.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	$(function(){
		searchZipKakaoLocalPc();
	});

	function detailInputAddress() {
		$("#basicAddrInputArea").empty().html($("#taddr1").val()+$("#taddr2").val()+$("#extraAddr").val());
		$("#searchZipWrap").hide();
		$("#content").show();
		$(".popWrap").css('background-color', '#FFFFFF'); 		
		$("#extraAddr2").focus();
	}

	function returnAddressSearch() {
		$("#content").hide();
		$("#searchZipWrap").show();
		searchZipKakaoLocalPc();
	}	

	<%'// ��â�� �� ������ %>
	function CopyZipAPI()	{
		var frm = eval("opener.document.<%=strTarget%>");
		var basicAddr;
		var basicAddr2;
		basicAddr = "";
		basicAddr2 = "";
		basicAddr = $("#taddr1").val()+$("#taddr2").val()+$("#extraAddr").val();
		basicAddr2 = $("#extraAddr2").val();
		//basicAddr  = basicAddr.replace(/?/g,"/");		 
		//basicAddr2 = basicAddr2.replace(/?/g,"/");		

		<% if strMode="A" then %>
			// copy
			frm.reqzipcode.value		= $("#tzip").val();
			frm.reqzipaddr.value		= basicAddr
			frm.reqaddress.value		= basicAddr2
			// focus
			frm.reqaddress.focus();

		<% elseif strMode="B" then %>
			// copy
			frm.zipcode.value		= $("#tzip").val();
			frm.zipaddr.value		= basicAddr
			frm.useraddr.value		= basicAddr2
			// focus
			frm.useraddr.focus();

		<% elseif strMode="C" then %>
			// copy
			frm.company_zipcode.value		= $("#tzip").val();
			frm.company_address.value		= basicAddr
			frm.company_address2.value		= basicAddr2
			// focus
			frm.company_address2.focus();

		<% elseif strMode="D" then %>
			// copy
			frm.return_zipcode.value		= $("#tzip").val();
			frm.return_address.value		= basicAddr
			frm.return_address2.value		= basicAddr2
			// focus
			frm.return_address2.focus();

		<% elseif strMode="E" then %>
			// copy
			frm.zipcode.value		= $("#tzip").val();
			frm.addr1.value		= basicAddr
			frm.addr2.value		= basicAddr2
			// focus
			frm.addr2.focus();

		<% elseif strMode="F" then %>
			// copy
			frm.shopzipcode.value		= $("#tzip").val();
			frm.shopaddr1.value		= basicAddr
			frm.shopaddr2.value		= basicAddr2
			// focus
			frm.shopaddr2.focus();

		<% elseif strMode="G" then %>
			// copy
			frm.sPCd.value		= $("#tzip").val();
			frm.sAddr.value		= basicAddr + " " + basicAddr2;
			// focus
			frm.sAddr.focus();

		<% elseif strMode="I" then %>
			// copy
			frm.p_return_zipcode.value		= $("#tzip").val();
			frm.p_return_address.value		= basicAddr
			frm.p_return_address2.value		= basicAddr2
			// focus
			frm.p_return_address2.focus();

		<% elseif strMode="J" then %>
			// copy
			frm.returnZipcode.value		= $("#tzip").val();
			frm.returnZipaddr.value		= basicAddr
			frm.returnEtcaddr.value		= basicAddr2
			// focus
			frm.returnEtcaddr.focus();

		<% end if %>

		// close this window
		window.close();
	}

    function searchZipKakaoLocalPc() {
        // ���� scroll ��ġ�� �����س��´�.
		var currentScroll = Math.max(document.body.scrollTop, document.documentElement.scrollTop);
		// �����ȣ ã�� ã�� ȭ���� ���� element
		var element_wrap = document.getElementById('searchZipWrap');
		daum.postcode.load(function(){
			new daum.Postcode({
				oncomplete: function(data) {
					var addr = ''; // �ּ� ����
					var extraAddr = ''; // �����׸� ����

					<%'//����ڰ� ������ �ּ� Ÿ�Կ� ���� �ش� �ּ� ���� �����´�.%>
					if (data.userSelectedType === 'R') { // ����ڰ� ���θ� �ּҸ� �������� ���
						addr = data.roadAddress;
					} else { // ����ڰ� ���� �ּҸ� �������� ���(J)
						addr = data.jibunAddress;
					}

					<%'// ����ڰ� ������ �ּҰ� ���θ� Ÿ���϶� �����׸��� �����Ѵ�.%>
					if(data.userSelectedType === 'R'){
						<%'// ���������� ���� ��� �߰��Ѵ�. (�������� ����)%>
						<%'// �������� ��� ������ ���ڰ� "��/��/��"�� ������.%>
						if(data.bname !== '' && /[��|��|��]$/g.test(data.bname)){
							extraAddr += data.bname;
						}
						<%'// �ǹ����� �ְ�, ���������� ��� �߰��Ѵ�.%>
						if(data.buildingName !== '' && data.apartment === 'Y'){
							extraAddr += (extraAddr !== '' ? ', ' + data.buildingName : data.buildingName);
						}
						<%'// ǥ���� �����׸��� ���� ���, ��ȣ���� �߰��� ���� ���ڿ��� �����.%>
						if(extraAddr !== ''){
							extraAddr = ' (' + extraAddr + ')';
						}
						<%'// ���յ� �����׸��� �ش� �ʵ忡 �ִ´�.%>
						$("#extraAddr").val(extraAddr);
					} else {
						$("#extraAddr").val("");
					}

					<%'// �����ȣ�� �ּ� ������ �ش� �ʵ忡 �ִ´�.%>
					$("#tzip").val(data.zonecode);
					$("#taddr1").val(addr);

					<%'// iframe�� ���� element�� �Ⱥ��̰� �Ѵ�.%>
					<%'// (autoClose:false ����� �̿��Ѵٸ�, �Ʒ� �ڵ带 �����ؾ� ȭ�鿡�� ������� �ʴ´�.)%>
					<%'//element_wrap.style.display = 'none';%>

					<%'// �����ȣ ã�� ȭ���� ���̱� �������� scroll ��ġ�� �ǵ�����.%>
					document.body.scrollTop = currentScroll;
				},
				<%'// ����ڰ� �ּҸ� Ŭ��������%>
				onclose : function(state) {
					if(state === 'COMPLETE_CLOSE'){
						detailInputAddress();
					}
				},
				onresize : function(size) {
					//for (var key in this) {
					//	console.log("attributes : " + key + ", value : " + this[key]);
                    //}
                    //document.getElementById("__daum__layer_"+this.viewerNo).style.height = size.height+"px";
                    //parent.self.scrollTo(0, 0);
                    element_wrap.style.height = size.height + 'px';
                    parent.self.scrollTo(0, 0);
				},				
				width : '100%',
				height : '100%',
				hideMapBtn : true,
				hideEngBtn : true,
				shorthand : false
			}).embed(element_wrap);
	    });
		<%'// iframe�� ���� element�� ���̰� �Ѵ�.%>
        element_wrap.style.display = 'block';
    }
</script>
</head>
<body>
<div class="heightgird popV18">
	<div class="popWrap" style="background-color:#ececec;">
		<div class="popHeader">
			<h1>�ּ� �Է�</h1>
		</div>
		<div id="searchZipWrap" style="display:none;border:0px solid;width:100%;height:700px;margin:5px 0;position:relative"></div>							
		<div class="popContent tPad30">
			<%' content %>
			<div class="mySection">
				<div id="content" style="display:none;">
					<p class="rt" style="margin-bottom:-20px;"><a href="" onclick="returnAddressSearch();" class="btn btnS2 btnGry2"><span class="fn">�ּ� �ٽ� �˻�</span></a></p>
					<fieldset>
						<legend>�ּ� �Է� ��</legend>
						<table class="baseTable rowTable docForm">
						<caption class="visible">������ �ּҸ� �Է����ּ���</caption>
						<colgroup>
							<col width="120" /> <col width="*" />
						</colgroup>
						<tbody>
						<tr>
							<th scope="row">�ּ�</th>
							<td>
								<div class="rPad15">
									<span id="basicAddrInputArea"></span>
								</div>
								<div class="tPad07">
									<input type="text" class="txtInp box5" style="width:90%;" name="extraAddr2" id="extraAddr2" placeholder="���ּ� �Է�" />
								</div>
							</td>
						</tr>
						</tbody>
						</table>

						<div class="btnArea ct tPad20">
							<input type="submit" class="btn btnS1 btnRed btnW100" onclick="CopyZipAPI();" value="���" />
							<button type="button" class="btn btnS1 btnGry btnW100" onclick="window.close();">���</button>
						</div>
					</fieldset>
				</div>
			</div>
			<%' //content %>
		</div>
	</div>
	<div class="popFooter">
		<div class="btnArea">
			<button type="button" class="btn btnS1 btnGry2" onclick="window.close();">�ݱ�</button>
		</div>
	</div>
</div>
<form name="tranFrmApi" id="tranFrmApi" method="post">
	<input type="hidden" name="tzip" id="tzip">
	<input type="hidden" name="taddr1" id="taddr1">
	<input type="hidden" name="taddr2" id="taddr2">
    <input type="hidden" name="extraAddr" id="extraAddr">
	<input type="hidden" name="target" id="target" value="<%=strTarget%>">
	<input type="hidden" name="strMode" id="strMode" value="<%=strMode%>">	
</form>
<script src="https://ssl.daumcdn.net/dmaps/map_js_init/postcode.v2.js"></script>
</body>
</html>