<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<%
'###########################################################
' Description :  mobile 우편번호 찾기
' History : 2016.06.16 원승현 생성
'###########################################################
%>
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
	Dim strTarget
	Dim strMode, protocolAddr
	strTarget	= requestCheckVar(Request("target"),32)
	strMode     = requestCheckVar(Request("strMode"),32)


'response.write strTarget

	dim PageSize	: PageSize = getNumeric(requestCheckVar(request("psz"),5))
	dim CurrPage : CurrPage = getNumeric(requestCheckVar(request("cpg"),8))
	if CurrPage="" then CurrPage=1
	if PageSize="" then PageSize=5
%>
<script>
	var zipcodeScroll;
	$(function(){
		// 우편번호 찾기 탭,레이어
		$(".zipcode .tabContainer .tabCont").css("display", "none");
		$(".zipcode .tabContainer .tabCont:first-child").css("display", "block");
		$(".zipcode").delegate(".tabNav li", "click", function() {
			var index = $(this).parent().children().index(this);
			$(this).siblings().removeClass();
			$(this).addClass("current");
			$(this).parent().next(".tabContainer").children().hide().eq(index).show();
			zipcodeScroll.onResize();
			return false;
		});
		zipcodeScroll = new Swiper(".zipcodeLayer .scrollArea .swiper-container", {
			scrollbar:'.zipcodeLayer .swiper-scrollbar',
			direction:'vertical',
			slidesPerView:'auto',
			mousewheelControl: true,
			freeMode: true
		});
		$(window).resize(function() {
			var lyrH = $(".layerPopup").outerHeight();
			$(".layerPopup").css('margin-top', -lyrH/2);
		});

		if($(".layerPopup").is(":visible")){
			$(".layerMask").show();
			var lyrH = $(".layerPopup").outerHeight();
			$(".layerPopup").css('margin-top', -lyrH/2);
		} else {
			$(".layerMask").hide();
		}
	});

	<%'// 검색 %>
	function SubmitForm(stype) {

		<%'// 지번 일 경우 %>
		if (stype=="jibun")
		{
			if ($("#tJibundong").val().length < 2) { alert("검색어를 두 글자 이상 입력하세요."); return; }
			$("#sGubun").val(stype);
			$("#sJibundong").val($("#tJibundong").val());
			$("#cpg").val(1);
			$("#keyword").val("");
		}

		<%'// 도로명+건물번호 일 경우 %>
		if (stype=="RoadBnumber")
		{
			if ($("#city11").val()=="")
			{
				alert('시/도를 선택해 주세요.');
				return;
			}

			<%'// 세종특별자치시는 시군구가 없어서 체크안함 %>
			if ($("#city11").val()!="세종특별자치시")
			{
				if ($("#city12").val()=="")
				{
					alert('시/군/구를 선택해 주세요.');
					return;
				}
			}
			if ($("#NameRoadBnumber").val()=="")
			{
				alert('도로명을 입력해 주세요.');
				$("#NameRoadBnumber").focus();
				return;	
			}
			if ($("#NumberRoadBnumber").val()=="")
			{
				alert('건물번호를 입력해 주세요.');
				$("#NumberRoadBnumber").focus();
				return;	
			}

			$("#sGubun").val(stype);
			$("#sSido").val($("#city11").val());
			$("#sGungu").val($("#city12").val());
			$("#sRoadName").val($("#NameRoadBnumber").val());
			$("#sRoadBno").val($("#NumberRoadBnumber").val());
		}

		<%'// 도로명에 동(읍/면)+지번 일 경우 %>
		if (stype=="RoadBjibun")
		{
			if ($("#city21").val()=="")
			{
				alert('시/도를 선택해 주세요.');
				return;
			}

			<%'// 세종특별자치시는 시군구가 없어서 체크안함 %>
			if ($("#city21").val()!="세종특별자치시")
			{
				if ($("#city22").val()=="")
				{
					alert('시/군/구를 선택해 주세요.');
					return;
				}
			}
			if ($("#DongRoadBjibun").val()=="")
			{
				alert('동(읍/면)을 입력해 주세요.');
				$("#DongRoadBjibun").focus();
				return;	
			}
			if ($("#JibunRoadBjibun").val()=="")
			{
				alert('지번을 입력해 주세요.');
				$("#JibunRoadBjibun").focus();
				return;	
			}
			$("#sGubun").val(stype);
			$("#sSido").val($("#city21").val());
			$("#sGungu").val($("#city22").val());
			$("#sRoaddong").val($("#DongRoadBjibun").val());
			$("#sRoadjibun").val($("#JibunRoadBjibun").val());
		}

		<%'// 도로명에 건물명 일 경우 %>
		if (stype=="RoadBname")
		{
			if ($("#city31").val()=="")
			{
				alert('시/도를 선택해 주세요.');
				return;
			}

			<%'// 세종특별자치시는 시군구가 없어서 체크안함 %>
			if ($("#city31").val()!="세종특별자치시")
			{
				if ($("#city32").val()=="")
				{
					alert('시/군/구를 선택해 주세요.');
					return;
				}
			}
			if ($("#NameRoadBname").val()=="")
			{
				alert('건물명을 입력해 주세요.');
				$("#NameRoadBname").focus();
				return;	
			}
			$("#sGubun").val(stype);
			$("#sSido").val($("#city31").val());
			$("#sGungu").val($("#city32").val());
			$("#sRoadBname").val($("#NameRoadBname").val());
		}

		$.ajax({
			type:"get",
			url:"/apps/academy/lib/searchzipNewProc.asp",
		    data: $("#searchProcFrm").serialize(),
		    dataType: "text",
			async:false,
			cache:true,
			success : function(Data, textStatus, jqXHR){
				if (jqXHR.readyState == 4) {
					if (jqXHR.status == 200) {
						if(Data!="") {
							res = Data.split("|");
							if (res[0]=="OK")
							{
								if (stype=="jibun")
								{
									if (res[1]=="<li class='nodata'>검색된 주소가 없습니다.</li>")
									{
										SubmitFormAPI();
									}
									else
									{
										$("#resultJibun").show();
										setTimeout(function () {
											window.$('html,body').animate({scrollTop:$("#resultJibun").offset().top}, 0);
										}, 10);
										if (res[1]=="<li class='nodata'>검색된 주소가 없습니다.</li>")
										{
											$("#JibunHelp").hide();
										}
										else
										{
											$("#JibunHelp").show();
										}
										$("#jibunaddrList").empty().html(res[1]);
										if (res[3]!="")
										{
											$("#addrpaging").empty().html(res[3]);
										}
										if (res[2] > 100)
										{
											$("#cautionTxtJibun").empty().html("<p></p><p>검색 결과가 많을 경우 지번 또는 건물명과 함께 검색해주세요</p><p class='ex'>예) 동숭동 1-45, 동숭동 동숭아트센타</p>");
											$("#cautionTxtJibun").show();
										}
										else
										{
											$("#cautionTxtJibun").empty();
										}
										$("#jibuntotalcntView").empty().html("총 <b>"+numberWithCommas(res[2])+"</b> 건");
									}
								}

								if (stype=="RoadBnumber")
								{
									$("#resultRoadBnumber").show();
									setTimeout(function () {
										window.$('html,body').animate({scrollTop:$("#resultRoadBnumber").offset().top}, 0);
									}, 10);
									if (res[1]=="<li class='nodata'>검색된 주소가 없습니다.</li>")
									{
										$("#RoadBnumberHelp").hide();
									}
									else
									{
										$("#RoadBnumberHelp").show();
									}
									$("#RoadBnumberaddrList").empty().html(res[1]);
								}

								if (stype=="RoadBjibun")
								{
									$("#resultRoadBjibun").show();
									setTimeout(function () {
										window.$('html,body').animate({scrollTop:$("#resultRoadBjibun").offset().top}, 0);
									}, 10);
									if (res[1]=="<li class='nodata'>검색된 주소가 없습니다.</li>")
									{
										$("#RoadBjibunHelp").hide();
									}
									else
									{
										$("#RoadBjibunHelp").show();
									}
									$("#RoadBjibunaddrList").empty().html(res[1]);
								}

								if (stype=="RoadBname")
								{
									$("#resultRoadBname").show();
									setTimeout(function () {
										window.$('html,body').animate({scrollTop:$("#resultRoadBname").offset().top}, 0);
									}, 10);
									if (res[1]=="<li class='nodata'>검색된 주소가 없습니다.</li>")
									{
										$("#RoadBnameHelp").hide();
									}
									else
									{
										$("#RoadBnameHelp").show();
									}
									$("#RoadBnameaddrList").empty().html(res[1]);
								}

								setTimeout(function () {
									zipcodeScroll.onResize();
								}, 50);

							}
							else
							{
								errorMsg = res[1].replace(">?n", "\n");
								alert(errorMsg );
								return false;
							}
						} else {
							alert("잘못된 접근 입니다1.");
							return false;
						}
					}
				}
			},
			error:function(jqXHR, textStatus, errorThrown){
				alert("잘못된 접근 입니다2!");
				return false;
			}
		});
	}


	<%'// 시군구 리스트 가져옴 %>
	function getgunguList(v, stype) {

		$("#sGubun").val("gungureturn");
		$("#sSidoGubun").val(v);

		if (v=="")
		{
			alert("시/도를 선택해 주세요.");
			return false;
		}

		<%'// 세종특별자치시는 시군구가 없으므로 안타도됨 %>
		if (v=="세종특별자치시")
		{
			$("#"+stype).empty().html("<option value=''>시/군/구 없음</option>");
		}
		else
		{
			$.ajax({
				type:"POST",
				url:"/apps/academy/lib/searchzipNewProc.asp",
			   data: $("#searchProcFrm").serialize(),
			   dataType: "text",
				async:false,
				cache:true,
				success : function(Data, textStatus, jqXHR){
					if (jqXHR.readyState == 4) {
						if (jqXHR.status == 200) {
							if(Data!="") {
								res = Data.split("|");
								if (res[0]=="OK")
								{
									$("#"+stype).empty().html(res[1]);
								}
								else
								{
									errorMsg = res[1].replace(">?n", "\n");
									alert(errorMsg );
									return false;
								}
							} else {
								alert("잘못된 접근 입니다.3");
								return false;
							}
						}
					}
				},
				error:function(jqXHR, textStatus, errorThrown){
					alert("잘못된 접근 입니다!4");
					return false;
				}
			});
		}
	}

	function numberWithCommas(x) {
		return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
	}

	function setBackAction(x, y, z) {
		$("#"+x).hide();
		$("#"+y).show();
		$("#"+z).show();
		setTimeout(function () {
			zipcodeScroll.onResize();
		}, 50);
	}

	<%'// form에 각 값들 넣고 기본, 상세 주소 입력값 만듦 %>
	function setAddr(zip, sido, gungu, dong, eupmyun, ri, official_bld, jibun, road, building_no, type, wp, uwp) {

		var basicAddr; // 기본주소
		var basicAddr2; // 상세주소
		var roadbasicAddr; // 도로명으로 검색할시 표시할 지번주소

		$("#zip").val(zip);
		$("#sido").val(sido);
		$("#gungu").val(gungu);
		$("#dong").val(dong);
		$("#eupmyun").val(eupmyun);
		$("#ri").val(ri);
		$("#official_bld").val(official_bld);
		$("#jibun").val(jibun);
		$("#road").val(road);
		$("#building_no").val(building_no);

		if (type=="jibun")
		{
			<%'// 기본주소 입력값을 만든다.%>
			basicAddr = "["+zip+"] "+sido+" "+gungu;
			if (dong=="")
			{
				basicAddr = basicAddr + " "+eupmyun;
			}
			else
			{
				basicAddr = basicAddr + " "+dong;
			}
			if (ri!="")
			{
				basicAddr = basicAddr + " "+ri;
			}
			<%'// 상세주소 입력값을 만든다.%>
			if (official_bld!="")
			{
				basicAddr = basicAddr + " "+official_bld+" "+jibun;
			}
			else
			{
				basicAddr = basicAddr + " "+jibun;
			}
			$("#Jibunfinder").hide();
			$("#resultJibun").hide();
			$("#jibunDetail").show();
		}

		if (type=="RoadBnumber")
		{
			<%'// 기본주소 입력값을 만든다.%>
			basicAddr = "["+zip+"] "+sido+" "+gungu;
			if (eupmyun!="")
			{
				basicAddr = basicAddr + " "+eupmyun+" "+road;
			}
			else
			{
				basicAddr = basicAddr + " "+road;
			}
			if (building_no!="")
			{
				basicAddr = basicAddr + " "+building_no;
			}
			<%'// 상세주소 입력값을 만든다.%>
			if (official_bld!="")
			{
				basicAddr2 = ""+official_bld+"";
			}

			<%' // 지번주소 입력값을 만든다.%>
			roadbasicAddr = sido+" "+gungu;
			if (dong=="")
			{
				roadbasicAddr = roadbasicAddr + " "+eupmyun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+dong;
			}
			if (ri!="")
			{
				roadbasicAddr = roadbasicAddr + " "+ri;
			}
			if (official_bld!="")
			{
				roadbasicAddr = roadbasicAddr + " "+official_bld+" "+jibun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+jibun;
			}
			$("#RoadBnumberJibunDetail").empty().html("지번 주소 : "+roadbasicAddr);
			$("#RoadBnumberfinder").hide();
			$("#resultRoadBnumber").hide();
			$("#RoadBnumberDetail").show();
		}

		if (type=="RoadBjibun")
		{
			<%'// 기본주소 입력값을 만든다.%>
			basicAddr = "["+zip+"] "+sido+" "+gungu;
			if (eupmyun!="")
			{
				basicAddr = basicAddr + " "+eupmyun+" "+road;
			}
			else
			{
				basicAddr = basicAddr + " "+road;
			}
			if (building_no!="")
			{
				basicAddr = basicAddr + " "+building_no;
			}
			<%'// 상세주소 입력값을 만든다.%>
			if (official_bld!="")
			{
				basicAddr2 = ""+official_bld+"";
			}

			<%' // 지번주소 입력값을 만든다.%>
			roadbasicAddr = sido+" "+gungu;
			if (dong=="")
			{
				roadbasicAddr = roadbasicAddr + " "+eupmyun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+dong;
			}
			if (ri!="")
			{
				roadbasicAddr = roadbasicAddr + " "+ri;
			}
			if (official_bld!="")
			{
				roadbasicAddr = roadbasicAddr + " "+official_bld+" "+jibun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+jibun;
			}
			$("#RoadBjibunJibunDetail").empty().html("지번 주소 : "+roadbasicAddr);
			$("#RoadBjibunfinder").hide();
			$("#resultRoadBjibun").hide();
			$("#RoadBjibunDetail").show();
		}

		if (type=="RoadBname")
		{
			<%'// 기본주소 입력값을 만든다.%>
			basicAddr = "["+zip+"] "+sido+" "+gungu;
			if (eupmyun!="")
			{
				basicAddr = basicAddr + " "+eupmyun+" "+road;
			}
			else
			{
				basicAddr = basicAddr + " "+road;
			}
			if (building_no!="")
			{
				basicAddr = basicAddr + " "+building_no;
			}
			<%'// 상세주소 입력값을 만든다.%>
			if (official_bld!="")
			{
				basicAddr2 = ""+official_bld+"";
			}

			<%' // 지번주소 입력값을 만든다.%>
			roadbasicAddr = sido+" "+gungu;
			if (dong=="")
			{
				roadbasicAddr = roadbasicAddr + " "+eupmyun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+dong;
			}
			if (ri!="")
			{
				roadbasicAddr = roadbasicAddr + " "+ri;
			}
			if (official_bld!="")
			{
				roadbasicAddr = roadbasicAddr + " "+official_bld+" "+jibun;
			}
			else
			{
				roadbasicAddr = roadbasicAddr + " "+jibun;
			}
			$("#RoadBnameJibunDetail").empty().html("지번 주소 : "+roadbasicAddr);
			$("#RoadBnamefinder").hide();
			$("#resultRoadBname").hide();
			$("#RoadBnameDetail").show();
		}

		$("#"+wp).empty().html(basicAddr);
		if (basicAddr2!="")
		{
			$("#"+uwp).val(basicAddr2);
		}
		$("#"+uwp).focus();

		setTimeout(function () {
			zipcodeScroll.onResize();
		}, 50);
	}


	<%'// 모창에 값 던져줌 %>
	function CopyZip(x, y)	{

		<%'// api로 검색시에는 CopyZipAPI로 던져줌 %>
		if ($("#keyword").val()!="")
		{
			CopyZipAPI(x, y);
			return false;
		}

		var frm = document.<%=strTarget%>;
		var basicAddr;
		var basicAddr2;

		<%'// 기본주소 입력값을 만든다.%>
		basicAddr = $("#sido").val()+" "+$("#gungu").val();

		if (y=="jibun")
		{
			<%'// 상세주소 입력값을 만든다.%>
			if ($("#dong").val()=="")
			{
				basicAddr2 = $("#eupmyun").val();;
			}
			else
			{
				basicAddr2 = $("#dong").val();
			}
			if ($("#ri").val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#ri").val();
			}
			if ($("#official_bld").val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#official_bld").val()+" "+$("#jibun").val();
			}
			else
			{
				basicAddr2 = basicAddr2 + " "+$("#jibun").val();
			}
			if ($("#"+x).val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#"+x).val();
			}
		}
		if (y=="RoadBnumber")
		{
			if ($("#eupmyun").val()!="")
			{
				basicAddr2 = $("#eupmyun").val()+" "+$("#road").val();
			}
			else
			{
				basicAddr2 = $("#road").val();
			}
			if ($("#building_no").val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#building_no").val();
			}
			if ($("#"+x).val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#"+x).val();
			}
		}
		if (y=="RoadBjibun")
		{
			if ($("#eupmyun").val()!="")
			{
				basicAddr2 = $("#eupmyun").val()+" "+$("#road").val();
			}
			else
			{
				basicAddr2 = $("#road").val();
			}
			if ($("#building_no").val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#building_no").val();
			}
			if ($("#"+x).val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#"+x).val();
			}

		}
		if (y=="RoadBname")
		{
			if ($("#eupmyun").val()!="")
			{
				basicAddr2 = $("#eupmyun").val()+" "+$("#road").val();
			}
			else
			{
				basicAddr2 = $("#road").val();
			}
			if ($("#building_no").val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#building_no").val();
			}
			if ($("#"+x).val()!="")
			{
				basicAddr2 = basicAddr2 + " "+$("#"+x).val();
			}
		}


		// copy
		<%
			Select Case strTarget
				Case "frmWrite"		'나의 주소록 Form
		%>
			//frm.zip1.value			= post1;
			//frm.zip2.value			= post2;
			frm.zip.value		= $("#zip").val();
			frm.reqZipaddr.value	= basicAddr;
			frm.reqAddress.value	= basicAddr2;
		<%		Case "buyer" %>
			frm.buyZip1.value		= post1;
			frm.buyZip2.value		= post2;
			frm.buyAddr1.value		= add;
			frm.buyAddr2.value		= dong;
		<%		Case "userinfo" %>
			frm.txZip1.value		= post1;
			frm.txZip2.value		= post2;
			frm.txAddr1.value		= add;
			frm.txAddr2.value		= dong;
		<%		Case Else %>
			frm.txZip.value		= $("#zip").val();
			frm.txAddr1.value		= basicAddr;
			frm.txAddr2.value		= basicAddr2;
		<%	End Select %>
		closeLayer();

	}

	function jsPageGo(icpg){
		var frm = document.searchProcFrm;
		frm.cpg.value=icpg;

		$.ajax({
			type:"get",
			url:"/apps/academy/lib/searchzipNewProc.asp",
			data: $("#searchProcFrm").serialize(),
			dataType: "text",
			async:false,
			cache:true,
			success : function(Data, textStatus, jqXHR){
				if (jqXHR.readyState == 4) {
					if (jqXHR.status == 200) {
						if(Data!="") {
							var str;
							for(var i in Data)
							{
								 if(Data.hasOwnProperty(i))
								{
									str += Data[i];
								}
							}
							str = str.replace("undefined","");
							res = str.split("|");
							if (res[0]=="OK")
							{
								$("#resultJibun").show();
								$("#jibunaddrList").empty().html(res[1]);
								if (res[3]!="")
								{
									$("#addrpaging").empty().html(res[3]);
								}
								if (res[2] > 100)
								{
									$("#cautionTxtJibun").empty().html("<p></p><p>검색 결과가 많을 경우 지번 또는 건물명과 함께 검색해주세요</p><p class='ex'>예) 동숭동 1-45, 동숭동 동숭아트센타</p>");
									$("#cautionTxtJibun").show();
								}
								else
								{
									$("#cautionTxtJibun").empty();
								}
								setTimeout(function () {
									zipcodeScroll.onResize();
								}, 50);
							}
							else
							{
								errorMsg = res[1].replace(">?n", "\n");
								alert(errorMsg );
								return false;
							}
						} else {
							alert("잘못된 접근 입니다.5");
							return false;
						}
					}
				}
			},
			error:function(jqXHR, textStatus, errorThrown){
				alert("잘못된 접근 입니다6!");
				return false;
			}
		});

	}

	<%' 검색 juso.go.kr api 사용영역 %>
	function SubmitFormAPI()
	{
		if ($("#tJibundong").val().length < 2) { alert("검색어를 두 글자 이상 입력하세요."); return; }
		$("#keyword").val($("#tJibundong").val());
		$("#currentPage").val(1);
		$.ajax({
/*
		     url :"http://www.juso.go.kr/addrlink/addrLinkApiJsonp.do"
			,type:"post"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,cache:true
			,crossDomain:true
*/
			 url : "/lib/sz_gate.asp" 
			,type:"get"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,cache:true
			,success:function(xmlStr){
				if(navigator.appName.indexOf("Microsoft") > -1){
					var xmlData = new ActiveXObject("Microsoft.XMLDOM");
					xmlData.loadXML(xmlStr.returnXml)
				}else{
					var xmlData = xmlStr.returnXml;
				}
				$("#jibunaddrList").html("");
				var errCode = $(xmlData).find("errorCode").text();
				var errDesc = $(xmlData).find("errorMessage").text();
				if(errCode != "0"){
					alert(errCode+"="+errDesc);
				}else{
					if ($(xmlData).find("totalCount").text()=="0")
					{
						$("#Jibunfinder").show();
						$("#JibunHelp").hide();
						$("#resultJibun").show();
						$("#addrpaging").empty();
						$("#jibunaddrList").empty().html("<li class='nodata'>검색된 주소가 없습니다.</li>");
					}
					else
					{
						if(xmlStr != null){

							$("#Jibunfinder").show();
							$("#resultJibun").show();
							$("#JibunHelp").show();
							$("#jibuntotalcntView").empty().html("총 <b>"+$(xmlData).find("totalCount").text()+"</b> 건");
							if (parseInt($(xmlData).find("totalCount").text())>=100)
							{
								$("#cautionTxtJibun").empty().html("<p></p><p>검색 결과가 많을 경우 지번 또는 건물명과 함께 검색해주세요</p><p class='ex'>예) 동숭동 1-45, 동숭동 동숭아트센타</p>");
								$("#cautionTxtJibun").show();
							}
							fnDisplayPaging_New_nottextboxdirectJS($("#currentPage").val(),$(xmlData).find("totalCount").text(),$("#countPerPage").val(),4,'jsPageGoAPI');
							makeList(xmlData);
							setTimeout(function () {
								zipcodeScroll.onResize();
							}, 50);
						}
					}
				}
			}
		});
	}

	<%'// 페이징 자바스크립트 버전 %>
	function fnDisplayPaging_New_nottextboxdirectJS(strCurrentPage, intTotalRecord, intRecordPerPage, intBlockPerPage, strJsFuncName) {
		var intCurrentPage;
		var strCurrentPath;
		var vPageBody;
		var intStartBlock;
		var intEndBlock;
		var intTotalPage;
		var strParamName;
		var intLoop;

		<%'// 현재 페이지 설정 %>
		intCurrentPage = strCurrentPage;

		<%'// 해당 페이지에 표시되는 시작페이지와 마지막페이지 설정 %>
		intStartBlock = parseInt((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1;
		intEndBlock = parseInt((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage;

		<%'// 총 페이지 수 설정 %>
		intTotalPage = parseInt((intTotalRecord - 1)/intRecordPerPage) + 1

		if (intTotalPage < 1)
		{
			intTotalPage = 1;
		}

		vPageBody = "";
		vPageBody = vPageBody + "<div class='pagination'>";

		if (intCurrentPage = 1)
		{
			vPageBody = vPageBody + "<a href='javascript:'  class='btnPrev'><span>이전 페이지</span></a>";
		}
		else
		{
			vPageBody = vPageBody + "<a href='javascript:" + strJsFuncName + "(" + (intCurrentPage - 1) + ")'  class='btnPrev'><span>이전 페이지</span></a>";
		}

		vPageBody = vPageBody + " <span> ";
		vPageBody = vPageBody + "<input type='number' class='pageNum' value='"+intCurrentPage+"' min='1' max='"+intTotalPage+"' onkeypress='if(event.keyCode==13){fnDirPg" + strJsFuncName + "(this.value); return false;}' />";
		vPageBody = vPageBody + "/ "+intTotalPage+"</span>";

		if (intTotalPage >= intCurrentPage + 1)
		{
			vPageBody = vPageBody + "<a href='javascript:" + strJsFuncName + "(" + (intCurrentPage + 1) + ")'  class='btnNext'><span>다음 페이지</span></a>";
		}
		else
		{
			vPageBody = vPageBody + "<a href='javascript:'  class='btnNext'><span>다음 페이지</span></a>";
		}
		vPageBody = vPageBody + "</div>"
		$("#addrpaging").empty().html(vPageBody);

	}

	function jsPageGoAPI(icomp)
	{
		$("#currentPage").val(icomp);
		$.ajax({
/*
		     url :"http://www.juso.go.kr/addrlink/addrLinkApiJsonp.do"
			,type:"post"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,crossDomain:true
			,cache:true
*/
			 url : "/lib/sz_gate.asp" 
			,type:"get"
			,data:$("#searchProcApi").serialize()
			,dataType:"jsonp"
			,cache:true
			,success:function(xmlStr){
				if(navigator.appName.indexOf("Microsoft") > -1){
					var xmlData = new ActiveXObject("Microsoft.XMLDOM");
					xmlData.loadXML(xmlStr.returnXml)
				}else{
					var xmlData = xmlStr.returnXml;
				}
				$("#jibunaddrList").html("");
				var errCode = $(xmlData).find("errorCode").text();
				var errDesc = $(xmlData).find("errorMessage").text();
				if(errCode != "0"){
					alert(errCode+"="+errDesc);
				}else{
					if ($(xmlData).find("totalCount").text()=="0")
					{
						
						$("#Jibunfinder").show();
						$("#resultJibun").show();
						$("#JibunHelp").hide();
						$("#jibunaddrList").empty().html("<li class='nodata'>검색된 주소가 없습니다.</li>");
					}
					else
					{
						if(xmlStr != null){
							$("#Jibunfinder").show();
							$("#resultJibun").show();
							$("#JibunHelp").show();
							$("#jibuntotalcntView").empty().html("총 <b>"+$(xmlData).find("totalCount").text()+"</b> 건");
							fnDisplayPaging_New_nottextboxdirectJS($("#currentPage").val(),$(xmlData).find("totalCount").text(),$("#countPerPage").val(),4,'jsPageGoAPI');
							makeList(xmlData);
							setTimeout(function () {
								zipcodeScroll.onResize();
							}, 50);
						}
					}
				}
			}
		});

	}

	function makeList(xmlStr){
		var htmlStr = "";
		$(xmlStr).find("juso").each(function(){
			var s = "'"+$(this).find('zipNo').text()+"','"+$(this).find('jibunAddr').text()+"','jibunDetailtxt','jibunDetailAddr2'";
			htmlStr += '<li><a href="" onclick="setAddrAPI('+s+');return false;">'+$(this).find('jibunAddr').text() +'<br>';
			htmlStr += "<span>도로명주소 : "+ $(this).find('roadAddr').text() +"</span></a></li>";
		});
		$("#jibunaddrList").empty().html(htmlStr);

	}

	function setAddrAPI(zip, addr, wp, uwp)
	{
		var basicAddr; // 기본주소

		basicAddr = "["+zip+"] "+addr;

		basicAddr = basicAddr.replace("  "," ");
		addr = addr.replace("  "," ");

		$("#tzip").val(zip);
		$("#taddr1").val(addr);

		$("#"+wp).empty().html(basicAddr);
		$("#"+uwp).focus();

		$("#Jibunfinder").hide();
		$("#resultJibun").hide();
		$("#jibunDetail").show();

		setTimeout(function () {
			zipcodeScroll.onResize();
		}, 50);
	}

	<%'// 모창에 값 던져줌 %>
	function CopyZipAPI(x, y)	{
		var frm = eval("document.<%=strTarget%>");
		var basicAddr;
		var basicAddr2;
		var chkAddr;
		var tmpaddr;
		basicAddr = "";
		basicAddr2 = "";

		<%'// 기본주소 입력값을 만든다.%>
		tmpaddr = $("#taddr1").val().split(" ");

		if (tmpaddr.length >= 3)
		{
			if (tmpaddr[2].substring(tmpaddr[2].length-1, tmpaddr[2].length)=="구")
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

		<%'// 상세주소 입력값을 만든다.%>
		for (var iadd=parseInt(chkAddr)+1;iadd < parseInt(tmpaddr.length);iadd++)
		{
			basicAddr2 += tmpaddr[iadd]+" ";
		}
		if ($("#"+x).val()!="")
		{
			basicAddr2 = basicAddr2 + $("#"+x).val();
		}

		// copy
		<%
			Select Case strTarget
				Case "frmWrite"		'나의 주소록 Form
		%>
			//frm.zip1.value			= post1;
			//frm.zip2.value			= post2;
			frm.zip.value		= $("#tzip").val();
			frm.reqZipaddr.value	= basicAddr;
			frm.reqAddress.value	= basicAddr2;
		<%		Case "buyer" %>
			frm.buyZip.value		= $("#tzip").val();
			frm.buyAddr1.value		= basicAddr;
			frm.buyAddr2.value		= basicAddr2;
		<%		Case "userinfo" %>
			frm.txZip.value		= $("#tzip").val();
			frm.txAddr1.value		= basicAddr;
			frm.txAddr2.value		= basicAddr2;
		<%		Case Else %>
			frm.txZip.value		= $("#tzip").val();
			frm.txAddr1.value		= basicAddr;
			frm.txAddr2.value		= basicAddr2;
		<%	End Select %>
		closeLayer();

	}

</script>
<div class="wrap">
	<div class="container headC">
		<%' 우편번호 찾기 레이어팝업 %>
		<div class="layerPopup zipcodeLayer">
			<div class="layerCont">
				<h2>우편번호 찾기</h2>
				<button type="button" class="layerClose" onclick="closeLayer();return false;"><span>닫기</span></button>
				<div class="scrollArea">
					<div class="swiper-container">
						<div class="swiper-wrapper">
							<div class="swiper-slide">
								<div class="zipcode">
									<ul class="tabNav tab1">
										<li class="current"><a href="#tabcont1">도로명 주소</a></li>
										<li><a href="#tabcont2">지번 주소</a></li>
									</ul>
									<div class="findZipcode tabContainer">
										<%' 도로명 주소 %>
										<div id="#tabcont1" class="tabCont">
											<div class="zipcodeTab">
												<ul class="tabNav tab2">
													<li class="current"><a href="#tabcont1-1">도로명 주소<br />+ 건물번호</a></li>
													<li><a href="#tabcont1-2">동(읍/면)<br />+ 지번</a></li>
													<li><a href="#tabcont1-3">건물명</a></li>
												</ul>
												<div class="tabContainer">
													<%' 도로명 주소+ 건물번호 %>
													<div id="tabcont1-1" class="tabCont">
														<%' 검색 %>
														<div class="finder" id="RoadBnumberfinder">
															<p class="help">도로명, 건물번호 를 입력 후 검색해주세요.<br />예) 대학로12길(도로명) 31(건물번호)</p>
															<ul>
																<li>
																	<label for="city11">시/도</label>
																	<select id="city11" onchange="getgunguList(this.value, 'city12')">
																		<option value="">시/도 선택</option>
																		<option value="서울특별시">서울특별시</option>
																		<option value="경기도">경기도</option>
																		<option value="강원도">강원도</option>
																		<option value="인천광역시">인천광역시</option>
																		<option value="충청북도">충청북도</option>
																		<option value="충청남도">충청남도</option>
																		<option value="대전광역시">대전광역시</option>
																		<option value="경상북도">경상북도</option>
																		<option value="경상남도">경상남도</option>
																		<option value="세종특별자치시">세종특별자치시</option>
																		<option value="대구광역시">대구광역시</option>
																		<option value="부산광역시">부산광역시</option>
																		<option value="울산광역시">울산광역시</option>
																		<option value="전라북도">전라북도</option>
																		<option value="전라남도">전라남도</option>
																		<option value="광주광역시">광주광역시</option>
																		<option value="제주특별자치도">제주특별자치도</option>
																	</select>
																</li>
																<li>
																	<label for="city12">시/군/구</label>
																	<select id="city12">
																		<option>시/군/구 선택</option>
																	</select>
																</li>
																<li>
																	<label for="road">도로명</label>
																	<input type="text" id="NameRoadBnumber" class="frmInputV16" />
																</li>
																<li>
																	<label for="buildingno">건물번호</label>
																	<input type="text" id="NumberRoadBnumber" onkeydown="javascript: if (event.keyCode == 13) {SubmitForm('RoadBnumber');}" class="frmInputV16" />
																</li>
															</ul>
															<div class="btnGroup">
																<div><button type="button" class="btn btnB1 btnYgn" onclick="SubmitForm('RoadBnumber');" >확인</button></div>
															</div>
															<div class="reference">
																<p>도로명 주소 검색 결과가 없을 경우,<br />도로명 주소 안내시스템을 참고해주시길 바랍니다</p>
																<a href="http://www.juso.go.kr" target="_blank" class="cBlu1">http://www.juso.go.kr</a>
															</div>
														</div>

														<%' 검색 결과 %>
														<div class="result" id="resultRoadBnumber" style="display:none;">
															<p class="help" id="RoadBnumberHelp">아래 주소 중 해당하는 주소를 선택해주세요.</p>
															<ul id="RoadBnumberaddrList"></ul>
														</div>

														<!-- 상세주소 입력 -->
														<div class="form" id="RoadBnumberDetail" style="display:none;">
															<p class="help">상세 주소를 입력해주세요.</p>
															<div class="address">
																<p><span id="RoadBnumberDetailTxt"></span><span id="RoadBnumberJibunDetail"></span></p>
																<input type="text" title="상세주소 입력" id="RoadBnumberDetailAddr2" placeholder="상세 주소를 입력해주세요" onkeydown="javascript: if (event.keyCode == 13) {CopyZip('RoadBnumberDetailAddr2', 'RoadBnumber');}" />
															</div>
															<div class="btnGroup">
																<div><button type="button" class="btn btnB1 btnWht2" onclick="setBackAction('RoadBnumberDetail','resultRoadBnumber','RoadBnumberfinder');return false;">이전</button></div>
																<div><button type="button" class="btn btnB1 btnYgn" onclick="CopyZip('RoadBnumberDetailAddr2', 'RoadBnumber');">주소입력</button></div>
															</div>
														</div>
													</div>

													<%' 동(읍/면)+ 지번 %>
													<div id="tabcont1-2" class="tabCont">
														<!-- 검색 -->
														<div class="finder" id="RoadBjibunfinder">
															<p class="help">동(읍/면), 지번 입력 후 검색해주세요.<br />예) 동숭동(동) 1-45 (지번)</p>
															<ul>
																<li>
																	<label for="city21">시/도</label>
																	<select id="city21" onchange="getgunguList(this.value, 'city22')" >
																		<option value="">시/도 선택</option>
																		<option value="서울특별시">서울특별시</option>
																		<option value="경기도">경기도</option>
																		<option value="강원도">강원도</option>
																		<option value="인천광역시">인천광역시</option>
																		<option value="충청북도">충청북도</option>
																		<option value="충청남도">충청남도</option>
																		<option value="대전광역시">대전광역시</option>
																		<option value="경상북도">경상북도</option>
																		<option value="경상남도">경상남도</option>
																		<option value="세종특별자치시">세종특별자치시</option>
																		<option value="대구광역시">대구광역시</option>
																		<option value="부산광역시">부산광역시</option>
																		<option value="울산광역시">울산광역시</option>
																		<option value="전라북도">전라북도</option>
																		<option value="전라남도">전라남도</option>
																		<option value="광주광역시">광주광역시</option>
																		<option value="제주특별자치도">제주특별자치도</option>
																	</select>
																</li>
																<li>
																	<label for="city22">시/군/구</label>
																	<select id="city22">
																		<option>시/군/구 선택</option>
																	</select>
																</li>
																<li>
																	<label for="town">동(읍/면)</label>
																	<input type="text" id="DongRoadBjibun" />
																</li>
																<li>
																	<label for="addressno">지번</label>
																	<input type="text" id="JibunRoadBjibun" onkeydown="javascript: if (event.keyCode == 13) {SubmitForm('RoadBjibun');}" />
																</li>
															</ul>
															<div class="btnGroup">
																<div><button type="button" class="btn btnB1 btnYgn" onclick="SubmitForm('RoadBjibun');">확인</button></div>
															</div>
															<div class="reference">
																<p>도로명 주소 검색 결과가 없을 경우,<br />도로명 주소 안내시스템을 참고해주시길 바랍니다</p>
																<a href="http://www.juso.go.kr" target="_blank" class="cBlu1">http://www.juso.go.kr</a>
															</div>
														</div>

														<%' 검색 결과 %>
														<div class="result" id="resultRoadBjibun" style="display:none;">
															<p class="help" id="RoadBjibunHelp">아래 주소 중 해당하는 주소를 선택해주세요.</p>
															<ul id="RoadBjibunaddrList"></ul>
														</div>

														<%' 상세주소 입력 %>
														<div class="form" id="RoadBjibunDetail" style="display:none;">
															<p class="help">상세 주소를 입력해주세요.</p>
															<div class="address">
																<p><span id="RoadBjibunDetailTxt"></span><span id="RoadBjibunJibunDetail"></span></p>
																<input type="text" title="상세주소 입력" placeholder="상세 주소를 입력해주세요" id="RoadBjibunDetailAddr2" onkeydown="javascript: if (event.keyCode == 13) {CopyZip('RoadBjibunDetailAddr2', 'RoadBjibun');}" />
															</div>
															<div class="btnGroup">
																<div><button type="button" class="btn btnB1 btnWht2" onclick="setBackAction('RoadBjibunDetail','resultRoadBjibun','RoadBjibunfinder');return false;">이전</button></div>
																<div><button type="button" class="btn btnB1 btnYgn" onclick="CopyZip('RoadBjibunDetailAddr2', 'RoadBjibun');">주소입력</button></div>
															</div>
														</div>
													</div>

													<%' 건물명 %>
													<div id="tabcont1-3" class="tabCont">
														<!-- 검색 -->
														<div class="finder">
															<p class="help">건물명을 입력 후 검색해주세요.<br />예) 자유빌딩(건물번호)</p>
															<ul>
																<li>
																	<label for="city31">시/도</label>
																	<select id="city31" onchange="getgunguList(this.value, 'city32')">
																		<option value="">시/도 선택</option>
																		<option value="서울특별시">서울특별시</option>
																		<option value="경기도">경기도</option>
																		<option value="강원도">강원도</option>
																		<option value="인천광역시">인천광역시</option>
																		<option value="충청북도">충청북도</option>
																		<option value="충청남도">충청남도</option>
																		<option value="대전광역시">대전광역시</option>
																		<option value="경상북도">경상북도</option>
																		<option value="경상남도">경상남도</option>
																		<option value="세종특별자치시">세종특별자치시</option>
																		<option value="대구광역시">대구광역시</option>
																		<option value="부산광역시">부산광역시</option>
																		<option value="울산광역시">울산광역시</option>
																		<option value="전라북도">전라북도</option>
																		<option value="전라남도">전라남도</option>
																		<option value="광주광역시">광주광역시</option>
																		<option value="제주특별자치도">제주특별자치도</option>
																	</select>
																</li>
																<li>
																	<label for="city32">시/군/구</label>
																	<select id="city32">
																		<option>시/군/구 선택</option>
																	</select>
																</li>
																<li>
																	<label for="building">건물명</label>
																	<input type="text" id="NameRoadBname" onkeydown="javascript: if (event.keyCode == 13) {SubmitForm('RoadBname');}" />
																</li>
															</ul>
															<div class="btnGroup">
																<div><button type="button" class="btn btnB1 btnYgn" onclick="SubmitForm('RoadBname');">확인</button></div>
															</div>
															<div class="reference">
																<p>도로명 주소 검색 결과가 없을 경우,<br />도로명 주소 안내시스템을 참고해주시길 바랍니다</p>
																<a href="http://www.juso.go.kr" target="_blank" class="cBlu1">http://www.juso.go.kr</a>
															</div>
														</div>

														<%' 검색 결과 %>
														<div class="result" id="resultRoadBname" style="display:none;">
															<p class="help">아래 주소 중 해당하는 주소를 선택해주세요.</p>
															<ul id="RoadBnameaddrList"></ul>
														</div>

														<%' 상세주소 입력 %>
														<div class="form" id="RoadBnameDetail" style="display:none;">
															<p class="help">상세 주소를 입력해주세요.</p>
															<div class="address">
																<p><span id="RoadBnameDetailTxt"></span><span id="RoadBnameJibunDetail"></span></p>
																<input type="text" title="상세주소 입력" placeholder="상세 주소를 입력해주세요" id="RoadBnameDetailAddr2" onkeydown="javascript: if (event.keyCode == 13) {CopyZip('RoadBnameDetailAddr2', 'RoadBname');}" />
															</div>
															<div class="btnGroup">
																<div><button type="button" class="btn btnB1 btnWht2" onclick="setBackAction('RoadBnameDetail','resultRoadBname','RoadBnamefinder');return false;">이전</button></div>
																<div><button type="button" class="btn btnB1 btnYgn" onclick="CopyZip('RoadBnameDetailAddr2', 'RoadBname');">주소입력</button></div>
															</div>
														</div>
													</div>
												</div>
											</div>
										</div>
										<%'// 도로명 주소 %>

										<%' 지번 주소 %>
										<div id="tabcont2" class="tabCont">
											<%' 검색 %>
											<div class="finder" id="Jibunfinder">
												<p class="help">찾고 싶으신 주소의 동(읍/면)을 입력해주세요.<br />예) 동숭동, 역삼1동</p>
												<ul>
													<li>
														<label for="dong">통합검색</label>
														<input type="text" id="tJibundong" placeholder="종로구 동숭동 1-45" onkeydown="javascript: if (event.keyCode == 13) {SubmitForm('jibun');}" />
													</li>
												</ul>
												<div class="btnGroup">
													<button type="button" class="btn btnB1 btnYgn" onclick="SubmitForm('jibun');">검색</button>
												</div>
											</div>

											<%' 검색 결과 %>
											<div class="result" id="resultJibun" style="display:none;">
												<p class="help" id="JibunHelp">아래 주소 중 해당하는 주소를 선택해주세요.<span id="cautionTxtJibun"></span></p>
												<p class="total" id="jibuntotalcntView"></p>
												<ul id="jibunaddrList"></ul>
												<div id="addrpaging" class="paging"></div>
											</div>

											<%' 상세주소 입력 %>
											<div class="form" id="jibunDetail" style="display:none;">
												<p class="help">상세 주소를 입력해주세요.</p>
												<div class="address">
													<p><div id="jibunDetailtxt"></div></p>
													<input type="text" title="상세주소 입력" id="jibunDetailAddr2" value="" placeholder="상세 주소를 입력해주세요" onkeydown="javascript: if (event.keyCode == 13) {CopyZip('jibunDetailAddr2', 'jibun');}" />
												</div>
												<div class="btnGroup">
													<div><button type="button" class="btn btnB1 btnWht2" onclick="setBackAction('jibunDetail','resultJibun','Jibunfinder');return false;">이전</button></div>
													<div><button type="button" class="btn btnB1 btnYgn" onclick="CopyZip('jibunDetailAddr2', 'jibun');">주소입력</button></div>
												</div>
											</div>
										</div>
										<%'// 지번 주소 %>
									</div>
								</div>
							</div>
						</div>
						<div class="swiper-scrollbar"></div>
					</div>
				</div>
			</div>
		</div>
		<!--// 우편번호 찾기 레이어팝업 -->
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
<form name="searchProcFrm" id="searchProcFrm" method="post">
	<input type="hidden" name="sGubun" id="sGubun">
	<input type="hidden" name="sJibundong" id="sJibundong">
	<input type="hidden" name="sSidoGubun" id="sSidoGubun">
	<input type="hidden" name="sSido" id="sSido">
	<input type="hidden" name="sGungu" id="sGungu">
	<input type="hidden" name="sRoadName" id="sRoadName">
	<input type="hidden" name="sRoadBno" id="sRoadBno">
	<input type="hidden" name="sRoaddong" id="sRoaddong">
	<input type="hidden" name="sRoadjibun" id="sRoadjibun">
	<input type="hidden" name="sRoadBname" id="sRoadBname">
	<input type="hidden" name="cpg" id="cpg" value="<%=currpage%>">
	<input type="hidden" name="psz" id="psz" value="<%=pagesize%>">
</form>

<form name="tranFrm" id="tranFrm" method="post">
	<input type="hidden" name="zip" id="zip">
	<input type="hidden" name="sido" id="sido">
	<input type="hidden" name="gungu" id="gungu">
	<input type="hidden" name="dong" id="dong">
	<input type="hidden" name="eupmyun" id="eupmyun">
	<input type="hidden" name="ri" id="ri">
	<input type="hidden" name="official_bld" id="official_bld">
	<input type="hidden" name="jibun" id="jibun">
	<input type="hidden" name="road" id="road">
	<input type="hidden" name="building_no" id="building_no">
</form>

<form name="searchProcApi" id="searchProcApi" method="post">
	<input type="hidden" name="currentPage" id="currentPage" value="1"/>
	<input type="hidden" name="countPerPage" id="countPerPage" value="5"/> 
	<input type="hidden" name="confmKey" id="confmKey" value="U01TX0FVVEgyMDE2MDcwNDIwMjE0NDEzNTk5"/>
	<input type="hidden" name="keyword" id="keyword" value=""/>
</form>

<form name="tranFrmApi" id="tranFrmApi" method="post">
	<input type="hidden" name="tzip" id="tzip">
	<input type="hidden" name="taddr1" id="taddr1">
	<input type="hidden" name="taddr2" id="taddr2">
</form>