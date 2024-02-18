<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"

'###########################################################
' Description :  �����ȣ ã��(īī�� API)
' History : 2019.06.13 ������ ����
'           2019.07.30 �ѿ�� ����Ʈ ���� ����
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
<style type="text/css">
html, body, blockquote, caption, dd, div, dl, dt, fieldset, form, frame, h1, h2, h3, h4, h5, h6, hr, iframe, input, legend, li, object, ol, p, pre, q, select, table, textarea, tr, td, ul, button {margin:0; padding:0;}
ol, ul {list-style:none;}
fieldset, img {border:0;}
h1, h2, h3, h4, h5, h6 {font-style:normal; font-size:12px;}
hr {display:none;}
table {border-collapse:collapse; border:0; empty-cells:show;}
textarea {resize:none;}
input, button {border:0;}
button {overflow:visible;}

body, h1, h2, h3 ,h4 {font-size:12px; font-family:dotum, dotumche, '����', '����ü', verdana, tahoma, sans-serif; line-height:1.6; color:#555;}
a {color:inherit; text-decoration:none;}
a:link, a:active, a:visited {color:#555;}
a:hover {text-decoration:none;}
a:hover {text-decoration:none;}
legend {visibility:hidden; width:0; height:0;}
caption {overflow:hidden; width:0; height:0; font-size:0; line-height:0; text-indent:-9999px;}
button {border:0; cursor:pointer;}
input[type=number]::-webkit-inner-spin-button {-webkit-appearance:none;}

html, body {height:100%;}

/* Popup layout */
body > .heightgird {min-height:100%; height:auto;}
.heightgird {position:relative;}
.popWrap {padding-bottom:45px;}
.popWrap .popHeader {padding:27px 15px 15px; background:#d50c0c; color:#fff;}
.popContent {padding:30px; font-size:11px;}
.popFooter {position:absolute; bottom:0; width:100%; padding:0; border-top:1px solid #ddd; background:#f5f5f5;}
.popFooter .btnArea {float:right; padding:8px 30px 11px 0;}
.popFooter .btnArea .btn {padding:5px 11px 3px 24px; border:0; border-bottom:1px solid #efefef; background:#999 url(http://fiximage.10x10.co.kr/web2013/common/btn_close_popup.gif) 11px center no-repeat;}
.popFooter .btnArea .btn:hover {border:0; border-bottom:1px solid #efefef; background:#8a8a8a url(http://fiximage.10x10.co.kr/web2013/common/btn_close_popup.gif) 11px center no-repeat;}
.popFooter button {font-family:Dotum; font-weight:normal;}
.popWrap .popHeader h1 img {vertical-align:top;}

</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
    $(function(){
        searchZipKakao();
	});

	<%'// ��â�� �� ������ %>
	function CopyZipAPI()	{
		var frm = eval("opener.document.<%=strTarget%>");
		var basicAddr;
		var basicAddr2;
		var chkAddr;
		var tmpaddr;
		basicAddr = "";
		basicAddr2 = "";

		<%'// �⺻�ּ� �Է°��� �����.%>
		tmpaddr = $("#taddr1").val().split(" ");

		if (tmpaddr.length >= 3)
		{
			if (tmpaddr[2].substring(tmpaddr[2].length-1, tmpaddr[2].length)=="��")
			{
				basicAddr = tmpaddr[0]+" "+tmpaddr[1]+" "+tmpaddr[2];
				chkAddr = "2";
			}
			else
			{
				basicAddr = tmpaddr[0]+" "+tmpaddr[1];
				chkAddr = "1";
			}
		}
		else
		{
			basicAddr = tmpaddr[0]+" "+tmpaddr[1];
			chkAddr = "1";
		}

		<%'// ���ּ� �Է°��� �����.%>
		for (var iadd=parseInt(chkAddr)+1;iadd < parseInt(tmpaddr.length);iadd++)
		{
			basicAddr2 += tmpaddr[iadd]+" ";
		}
		if ($("#extraAddr").val()!="")
		{
			basicAddr2 = basicAddr2 + $("#extraAddr").val();
		}

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

    // �����ȣ ã�� ã�� ȭ���� ���� element
    var element_wrap = $("#searchZipWrap");

    function searchZipKakao() {
        // ���� scroll ��ġ�� �����س��´�.
        var currentScroll = Math.max(document.body.scrollTop, document.documentElement.scrollTop);
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
						CopyZipAPI();
					}
				},
				width : '100%',
				height : '89%',
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
<img src="//fiximage.10x10.co.kr/web2019/common/tit_post.jpg" style="width:100%">
<div id="searchZipWrap" style="display:none;border:1px solid;width:500px;height:700px;margin:5px 0;position:relative">
</div>
<form name="tranFrmApi" id="tranFrmApi" method="post">
	<input type="hidden" name="tzip" id="tzip">
	<input type="hidden" name="taddr1" id="taddr1">
	<input type="hidden" name="taddr2" id="taddr2">
    <input type="hidden" name="extraAddr" id="extraAddr">
</form>
<script src="https://ssl.daumcdn.net/dmaps/map_js_init/postcode.v2.js"></script>
</body>
</html>