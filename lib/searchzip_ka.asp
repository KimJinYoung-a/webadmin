<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"

'###########################################################
' Description :  우편번호 찾기(카카오 API)
' History : 2019.06.13 원승현 프론트 생성
'           2022.09.28 한용민 프론트 로직 이전 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<%
	dim fiximgPath
	'이미지 경로 지정(SSL 처리)
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
<title>텐바이텐 10X10 : 우편번호찾기</title>
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

	<%'// 모창에 값 던져줌 %>
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
        // 현재 scroll 위치를 저장해놓는다.
		var currentScroll = Math.max(document.body.scrollTop, document.documentElement.scrollTop);
		// 우편번호 찾기 찾기 화면을 넣을 element
		var element_wrap = document.getElementById('searchZipWrap');
		daum.postcode.load(function(){
			new daum.Postcode({
				oncomplete: function(data) {
					var addr = ''; // 주소 변수
					var extraAddr = ''; // 참고항목 변수

					<%'//사용자가 선택한 주소 타입에 따라 해당 주소 값을 가져온다.%>
					if (data.userSelectedType === 'R') { // 사용자가 도로명 주소를 선택했을 경우
						addr = data.roadAddress;
					} else { // 사용자가 지번 주소를 선택했을 경우(J)
						addr = data.jibunAddress;
					}

					<%'// 사용자가 선택한 주소가 도로명 타입일때 참고항목을 조합한다.%>
					if(data.userSelectedType === 'R'){
						<%'// 법정동명이 있을 경우 추가한다. (법정리는 제외)%>
						<%'// 법정동의 경우 마지막 문자가 "동/로/가"로 끝난다.%>
						if(data.bname !== '' && /[동|로|가]$/g.test(data.bname)){
							extraAddr += data.bname;
						}
						<%'// 건물명이 있고, 공동주택일 경우 추가한다.%>
						if(data.buildingName !== '' && data.apartment === 'Y'){
							extraAddr += (extraAddr !== '' ? ', ' + data.buildingName : data.buildingName);
						}
						<%'// 표시할 참고항목이 있을 경우, 괄호까지 추가한 최종 문자열을 만든다.%>
						if(extraAddr !== ''){
							extraAddr = ' (' + extraAddr + ')';
						}
						<%'// 조합된 참고항목을 해당 필드에 넣는다.%>
						$("#extraAddr").val(extraAddr);
					} else {
						$("#extraAddr").val("");
					}

					<%'// 우편번호와 주소 정보를 해당 필드에 넣는다.%>
					$("#tzip").val(data.zonecode);
					$("#taddr1").val(addr);

					<%'// iframe을 넣은 element를 안보이게 한다.%>
					<%'// (autoClose:false 기능을 이용한다면, 아래 코드를 제거해야 화면에서 사라지지 않는다.)%>
					<%'//element_wrap.style.display = 'none';%>

					<%'// 우편번호 찾기 화면이 보이기 이전으로 scroll 위치를 되돌린다.%>
					document.body.scrollTop = currentScroll;
				},
				<%'// 사용자가 주소를 클릭했을때%>
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
		<%'// iframe을 넣은 element를 보이게 한다.%>
        element_wrap.style.display = 'block';
    }
</script>
</head>
<body>
<div class="heightgird popV18">
	<div class="popWrap" style="background-color:#ececec;">
		<div class="popHeader">
			<h1>주소 입력</h1>
		</div>
		<div id="searchZipWrap" style="display:none;border:0px solid;width:100%;height:700px;margin:5px 0;position:relative"></div>							
		<div class="popContent tPad30">
			<%' content %>
			<div class="mySection">
				<div id="content" style="display:none;">
					<p class="rt" style="margin-bottom:-20px;"><a href="" onclick="returnAddressSearch();" class="btn btnS2 btnGry2"><span class="fn">주소 다시 검색</span></a></p>
					<fieldset>
						<legend>주소 입력 폼</legend>
						<table class="baseTable rowTable docForm">
						<caption class="visible">나머지 주소를 입력해주세요</caption>
						<colgroup>
							<col width="120" /> <col width="*" />
						</colgroup>
						<tbody>
						<tr>
							<th scope="row">주소</th>
							<td>
								<div class="rPad15">
									<span id="basicAddrInputArea"></span>
								</div>
								<div class="tPad07">
									<input type="text" class="txtInp box5" style="width:90%;" name="extraAddr2" id="extraAddr2" placeholder="상세주소 입력" />
								</div>
							</td>
						</tr>
						</tbody>
						</table>

						<div class="btnArea ct tPad20">
							<input type="submit" class="btn btnS1 btnRed btnW100" onclick="CopyZipAPI();" value="등록" />
							<button type="button" class="btn btnS1 btnGry btnW100" onclick="window.close();">취소</button>
						</div>
					</fieldset>
				</div>
			</div>
			<%' //content %>
		</div>
	</div>
	<div class="popFooter">
		<div class="btnArea">
			<button type="button" class="btn btnS1 btnGry2" onclick="window.close();">닫기</button>
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