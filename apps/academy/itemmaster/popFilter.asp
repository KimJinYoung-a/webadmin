<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/CategoryMaster/displaycate/classes/displaycateCls.asp"-->
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 필터"

Dim cDisp, i, pdiv
SET cDisp = New cDispCate
cDisp.FCurrPage = 1
cDisp.FPageSize = 2000
cDisp.FRectDepth = 1
cDisp.FRectSiteGubun="upche"
cDisp.GetDispCateList()

pdiv = requestCheckVar(request("div"),1)

'전달값 처리
dim sellyn, limityn, sdiv, sortupdown, cate1, cate2
sellyn  = RequestCheckVar(request("sellyn"),2)
sortupdown = RequestCheckVar(request("sortupdown"),1)
limityn = RequestCheckVar(request("limityn"),1)
sdiv = RequestCheckVar(request("sdiv"),10)
cate1 = RequestCheckVar(request("cate1"),10)
cate2 = RequestCheckVar(request("cate2"),10)

if (sellyn="") then sellyn="YS"
If (limityn="") Then limityn="A"
If (sortupdown="") Then sortupdown="u"
If (sdiv="") Then sdiv="Reg"
%>
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<script>
$(function() {
	// select box control
	$('#ctgy1').on('change', function () {
		$('#ctgy2').removeAttr('disabled');
	});

	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});

});

function fnSearchFilterSelect(formname,formdata){
	if(formname!=''){
		eval("$('#"+formname+"')").val(formdata);
	}
}

function fnSearchSortFilterSelect(formdata,btnnum){
	if($("#sdiv").val()==formdata){
		if($("#ssort").val()=="u"){
			$("#ssort").val("d");
			eval("$('#btn"+btnnum +"')").removeClass("srtUp");
			eval("$('#btn"+btnnum +"')").addClass("srtDown");
		}else{
			$("#ssort").val("u");
			eval("$('#btn"+btnnum +"')").removeClass("srtDown");
			eval("$('#btn"+btnnum +"')").addClass("srtUp");
		}
	}else{
		$("#sdiv").val(formdata);
		$("#ssort").val("u");
		eval("$('#btn"+btnnum +"')").removeClass("srtDown");
		eval("$('#btn"+btnnum +"')").addClass("srtUp");
	}
}

function fnAppCallWinConfirm(){
	var arrinfochk='', arrinfocont='', arrinfocd='';
	var jsontxt;
	var iteminfo = document.iteminfo;

	//if($("#depth1 option:selected").val()==""){
	//	alert("카테고리를 선택해주세요.");
	//}else{
		jsontxt = $("#depth1").val() + "," + $("#depth2").val() + "," + $("#sellyn").val() + "," + $("#limityn").val() + "," + $("#sdiv").val() + "," + $("#ssort").val() + ",Y";
		//alert(jsontxt);
		fnAPPopenerJsCallClose("fnSearchFilterSet(\"" + jsontxt + "\")");
	//}
}

jQuery(document).ready(function(){
<% if cate1<>"" then %>
	//카테고리 지정
	startCategoryRequest('2depth','<%=cate1%>','<%=cate2%>');
	$('#depth2').removeAttr('disabled');
<% end if %>

	//$("#depth1").change(function(){
		fnAPPShowRightConfirmBtns();
	//});
	$('#depth1').on('change', function () {
		$('#depth2').removeAttr('disabled');
		$('.addBtn button').removeAttr('disabled');
		$('.addBtn button').addClass('active');
	});
});

// AJAX 프로그램
var parentFrmName = "searchForm";
var xmlHttp;
var xmlDoc;
var xmlHttpMode, xmlHttpParam1, xmlHttpParam2;
var xmlHttpDefaultSet;
var xmlProcessId = 0;

function createXMLHttpRequest() {
	if (window.ActiveXObject) {
			xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
	} else if (window.XMLHttpRequest) {
			xmlHttp = new XMLHttpRequest();
	}
}

function startCategoryRequest(mode,param,param2) {
	var frm = eval("document." + parentFrmName);
	xmlHttpMode = mode;
	xmlHttpParam1 = param;
	xmlHttpParam2 = param2;
	//frm.catecode_a.value=param;
	//$("input[name='depthname']").val($("select[name='depth1'] option:selected").text());

	createXMLHttpRequest();
	xmlHttp.onreadystatechange = callback;
	xmlHttp.open("GET", "/apps/academy/itemmaster/reqCategoryResponse.asp?param1=" + param, true);
	xmlHttp.send(null);
}

function callback() {
	if(xmlHttp.readyState == 4) {
            if(xmlHttp.status == 200) {
                    // 정상적인 데이타 반환
                    // 전체(TXT) : xmlHttp.responseText
                    if (window.ActiveXObject) {
                            // XML 로 변환한다.
                            // 텍스트 앞부분에서 "<" 이전 문자들을 제거한다.(공백문자 제거용,  이렇게 안하면 변환이 안된다 --)
                            xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
                            var rawXML = xmlHttp.responseText;
                            var filteredML;

                            var index = 0;
                            for (var i = 0; i < rawXML.length; i++) {
                                    if (rawXML.charAt(i) == "<") {
                                            index = i;
                                            break;
                                    }
                            }
                            filteredML = rawXML.substring(index);
                            xmlDoc.loadXML(filteredML);
                    } else if (window.XMLHttpRequest) {
                            xmlDoc = xmlHttp.responseXML;
                    }
					//alert(xmlHttp.responseText);
                    process();
            } else if (xmlHttp.status == 204){
                    // 데이터가 존재하지 않을 경우
                    alert("데이타가 존재하지 않습니다.(CODE : 200)");
            } else if (xmlHttp.status == 500){
                    // 에러발생시
                    alert("데이타 수신중 에러가 발생하였습니다.(CODE : 500)");
            }
    }
}

function process() {
	var frm = eval("document." + parentFrmName);
	var buf;
	var length = xmlDoc.getElementsByTagName("value1").length;
//	alert(length);
	if (xmlHttpMode=="1depth"){
		frm.depth1.length = (length*1+1);

			frm.depth1.options[0].value= "";
			frm.depth1.options[0].text= " 대분류 ";

		for (i=0;i<length;i++){
			frm.depth1.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.depth1.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam1){
				frm.depth1.options[i + 1].selected = true;
			}
		}

		//디폴트값
		if (xmlHttpParam2!="") { startCategoryRequest('2depth',xmlHttpParam1,xmlHttpParam2); }
	}else if (xmlHttpMode=="2depth"){
		frm.depth2.length = (length*1 + 1);

			frm.depth2.options[0].value= "";
			frm.depth2.options[0].text= " 중분류 ";

		for (i=0;i<length;i++){
			frm.depth2.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.depth2.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam2){
				frm.depth2.options[i + 1].selected = true;
			}
		}
		if ((xmlHttpParam2=="")&&(frm.depth2.length>0)) frm.depth2.options[0].selected = true;
	}
}
</script>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<h1 class="hidden">필터</h1>
			<!-- for dev msg : 레이어 작업시 이 부분만 가져가면 됩니다.-->
			<form name="searchForm" id="searchForm" method="get" style="margin:0px;">
			<input type="hidden" name="sellyn" id="sellyn" value="<%=sellyn%>">
			<input type="hidden" name="limityn" id="limityn" value="<%=limityn%>">
			<input type="hidden" name="sdiv" id="sdiv" value="Reg">
			<input type="hidden" name="ssort" id="ssort" value="<%=sortupdown%>">
			<div class="filterWrap">
				<dl class="dfCompos">
					<dt>카테고리</dt>
					<dd class="selectBtn">
						<div class="grid2">
							<select name="depth1" id="depth1" onchange="startCategoryRequest('2depth',this.value)">
								<option value="">대분류</option>
								<% If cDisp.FResultCount > 0 Then %>
								<% For i=0 To cDisp.FResultCount-1 %>
								<option value="<%=cDisp.FItemList(i).FCateCode%>" <%=chkIIF(cStr(cate1)=cStr(cDisp.FItemList(i).FCateCode),"selected","")%>><%=cDisp.FItemList(i).FCateName%></option>
								<% Next %>
								<% End If %>
							</select>
						</div>
						<div class="grid2">
							<select name="depth2" id="depth2" disabled="disabled">
								<option>중분류</option>
							</select>
						</div>
					</dd>
				</dl>
				<% If pdiv = "7" Then %>
				<% If sellyn<>"N" Then %>
				<dl class="dfCompos">
					<dt>판매상태</dt>
					<dd class="selectBtn">
						<div class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sellyn="YS","selected","")%>" onClick="fnSearchFilterSelect('sellyn','YS');">전체</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sellyn="Y","selected","")%>" onClick="fnSearchFilterSelect('sellyn','Y');">판매중</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sellyn="S","selected","")%>" onClick="fnSearchFilterSelect('sellyn','S');">일시품절</button></div>
					</dd>
				</dl>
				<% End If %>
				<dl class="dfCompos">
					<dt>한정구분</dt>
					<dd class="selectBtn">
						<div class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(limityn="A","selected","")%>" onClick="fnSearchFilterSelect('limityn','A');">전체</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(limityn="Y","selected","")%>" onClick="fnSearchFilterSelect('limityn','Y');">한정</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(limityn="N","selected","")%>" onClick="fnSearchFilterSelect('limityn','N');">비한정</button></div>
					</dd>
				</dl>
				<dl class="dfCompos">
					<dt>정렬기준</dt>
					<dd class="selectBtn">
						<ul>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Reg","selected","")%>" onClick="fnSearchSortFilterSelect('Reg',1);"><span class="sort <%=chkIIF(sdiv="Reg" and sortupdown="d","srtDown","srtUp")%>" id="btn1">등록순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Sales","selected","")%>" onClick="fnSearchSortFilterSelect('Sales',2);"><span class="sort <%=chkIIF(sdiv="Sales" and sortupdown="d","srtDown","srtUp")%>" id="btn2">매출순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="SaleCount","selected","")%>" onClick="fnSearchSortFilterSelect('SaleCount',3);"><span class="sort <%=chkIIF(sdiv="SaleCount" and sortupdown="d","srtDown","srtUp")%>" id="btn3">판매량순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Favo","selected","")%>" onClick="fnSearchSortFilterSelect('Favo',4);"><span class="sort <%=chkIIF(sdiv="Favo" and sortupdown="d","srtDown","srtUp")%>" id="btn4">관심등록순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Price","selected","")%>" onClick="fnSearchSortFilterSelect('Price',5);"><span class="sort <%=chkIIF(sdiv="Price" and sortupdown="d","srtDown","srtUp")%>" id="btn5">가격순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Disc","selected","")%>" onClick="fnSearchSortFilterSelect('Disc',6);"><span class="sort <%=chkIIF(sdiv="Disc" and sortupdown="d","srtDown","srtUp")%>" id="btn6">할인순</span></button></li>
						</ul>
					</dd>
				</dl>
				<% Else %>
				<dl class="dfCompos">
					<dt>대기상태</dt>
					<dd class="selectBtn">
						<ul>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sellyn="YS","selected","")%>" onClick="fnSearchFilterSelect('sellyn','YS');">전체</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sellyn="8","selected","")%>" onClick="fnSearchFilterSelect('sellyn','8');">임시저장</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sellyn="1","selected","")%>" onClick="fnSearchFilterSelect('sellyn','1');">대기</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sellyn="2","selected","")%>" onClick="fnSearchFilterSelect('sellyn','2');">보류</button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sellyn="0","selected","")%>" onClick="fnSearchFilterSelect('sellyn','0');">반려</button></li>
						</ul>
					</dd>
				</dl>
				<dl class="dfCompos">
					<dt>한정구분</dt>
					<dd class="selectBtn">
						<div class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(limityn="A","selected","")%>" onClick="fnSearchFilterSelect('limityn','A');">전체</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(limityn="Y","selected","")%>" onClick="fnSearchFilterSelect('limityn','Y');">한정</button></div>
						<div class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(limityn="N","selected","")%>" onClick="fnSearchFilterSelect('limityn','N');">비한정</button></div>
					</dd>
				</dl>
				<dl class="dfCompos">
					<dt>정렬기준</dt>
					<dd class="selectBtn">
						<ul>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Reg","selected","")%>" onClick="fnSearchSortFilterSelect('Reg',1);"><span class="sort <%=chkIIF(sdiv="Reg" and sortupdown="d","srtDown","srtUp")%>" id="btn1">등록순</span></button></li>
							<li class="grid3"><button type="button" class="btnM1 btnGry <%=chkIIF(sdiv="Price","selected","")%>" onClick="fnSearchSortFilterSelect('Price',2);"><span class="sort <%=chkIIF(sdiv="Price" and sortupdown="d","srtDown","srtUp")%>" id="btn2">가격순</span></button></li>
						</ul>
					</dd>
				</dl>
				<% End If %>
			</div>
			</form>
			<!-- //for dev msg : 레이어 작업시 이 부분만 가져가면 됩니다.-->
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->