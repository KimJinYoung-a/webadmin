<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/CategoryMaster/displaycate/classes/displaycateCls.asp"-->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 카테고리 선택"

Dim cDisp, i
SET cDisp = New cDispCate
cDisp.FCurrPage = 1
cDisp.FPageSize = 2000
cDisp.FRectDepth = 1
cDisp.FRectSiteGubun="upche"
cDisp.GetDispCateList()

Dim waititemid
waititemid = requestCheckVar(request("waititemid"),10)
%>
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<script type="text/javascript">
<!--
$(function() {
	// select box control
	$('#depth1').on('change', function () {
		// 카테고리가 2개면 추가버튼 비활성화
		if($("#tbl_DispCate li").length>1) {
			$('.addBtn button').attr('disabled','disabled').removeClass('active');
		} else {
			$('#depth2').removeAttr('disabled');
			$('.addBtn button').removeAttr('disabled').addClass('active');
		}
	});
});

// AJAX 프로그램
var parentFrmName = "cate";
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

function startCategoryRequest(mode,param) {
	var frm = eval("document." + parentFrmName);
	xmlHttpMode = mode;
	xmlHttpParam1 = param;
	xmlHttpParam2='';
	frm.catecode_a.value=param;
	$("input[name='depthname']").val($("select[name='depth1'] option:selected").text());

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

function addDispCateItem(dcd,cnm,div,dpt) {
	// 기존에 값에 중복 카테고리 여부 검사
	var CheckOverlap=false;
	$("input[name='catecode']").each(function(i){
		if($("input[name='catecode']:eq(" + i + ")").val() == dcd){
			CheckOverlap=true;
			return false;
		}
		else{
			CheckOverlap=false;
			return;
		}
	});
	if(CheckOverlap){
		alert("이미 지정된 같은 카테고리가 있습니다.");
	}else{
		// 행추가
		var oRow;
		oRow = "<li><div><span>" + $("input[name='depthname']").val() + "</span>";
		if(div=="y") {
			oRow += "<input type='hidden' name='isDefault' value='y'>";
		} else {
			oRow += "<input type='hidden' name='isDefault' value='n'>";
		}
		oRow += "<input type='hidden' name='catecode' value='" + dcd + "'>";
		oRow += "<input type='hidden' name='catedepth' value='" + dpt + "'>";
		oRow += "<input type='hidden' name='arrdepthname' value='" + $("input[name='depthname']").val() + "'>";
		oRow += "<button type='button' class='btnListDel' onClick='delDispCateItem(" + dcd + ")'>삭제</button></div></li>";
		$("#tbl_DispCate").append(oRow);
		$("input[name='selDispCateDiv']").val("n");
		$("input[name='checkcate']").val(Number(Number($("input[name='checkcate']").val())+1));
		chgodr('HelpInfo',1,'','');
	}
}

// 선택 전시카테고리 삭제
function delDispCateItem(catecode){
	if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
		$("input[name='catecode']").each(function(i){
			//alert($("input[name='isDefault']:eq(" + i + ")[value='y']").length);
			if($("input[name='catecode']:eq(" + i + ")").val() == catecode){
				if($("input[name='isDefault']:eq(" + i + ")[value='y']").length==1){
					$("input[name='selDispCateDiv']").val("y");
				}
				$(this).closest('li').remove();
				$("input[name='checkcate']").val(Number($("input[name='checkcate']").val())-1);
				return false;
			}
			else{
				return;
			}
		});

		// 카테고리가 하나도 없으면 안내 문구 표시
		if($("#tbl_DispCate li").length==0) {
			chgodr('HelpInfo',2,'','');
		}
	}
}

function sendDispCateItem(){
//alert("ok");
	if($("#tbl_DispCate li").length>1){
		alert("카테고리는 2개까지 선택 가능합니다.");
		return;
	}

	var dcd,cnm,div,dpt;

	dcd = $("input[name='catecode_a']").val();
	cnm = $("select[name='depth1'] option:selected");

	div = $("input[name='selDispCateDiv']").val();
	dpt = dcd.length/3
	if(dpt==0) {
		alert('카테고리를 선택해주세요.');
		return;
	}
	addDispCateItem(dcd,cnm,div,dpt);
	fnAPPShowRightConfirmBtns();//확인버튼 활성화

	//선택 박스 초기화
	$("#depth1").val("");
	$("#depth2").val("").attr('disabled','disabled');

	// 카테고리가 2개면 추가버튼 비활성화
	if($("#tbl_DispCate li").length>1) {
		$('.addBtn button').attr('disabled','disabled').removeClass('active');
	}
}

function js2DepthSelectBox(depthval) {
	$("input[name='catecode_a']").val(depthval);
	$("input[name='depthname']").val($("select[name='depth1'] option:selected").text() + " > " + $("select[name='depth2'] option:selected").text());
}

function fnAppCallWinConfirm(){
	var catename;
	var arrcode;
	var arrdepth;
	var arrdefault;
	var jsontxt;
	if($("input[name='checkcate']").val()<1){
		alert('카테고리를 추가해주세요.');
	}
	else{
		$("input[name=isDefault]").each(function(idx){
			// 기본 카테고리 정보 넘기기
			if(idx<1){
				arrcode = $("input[name=catecode]:eq(" + idx + ")").val();
				arrdepth = $("input[name=catedepth]:eq(" + idx + ")").val();
				arrdefault = $("input[name=isDefault]:eq(" + idx + ")").val();
			}else{
				arrcode += "," + $("input[name=catecode]:eq(" + idx + ")").val();
				arrdepth += "," + $("input[name=catedepth]:eq(" + idx + ")").val();
				arrdefault += "," + $("input[name=isDefault]:eq(" + idx + ")").val();
			}
		});
		$("#arrcatecode").val(arrcode);
		$("#arrcatedepth").val(arrdepth);
		$("#arrisdefault").val(arrdefault);
		document.cate.action="/apps/academy/itemmaster/popup/WaitDIYItemPopupDetailinfoEdit_Process.asp";
		document.cate.target="FrameCKP";
		document.cate.submit();
	}
}

function fnDetailInfoEnd(){
	$("input[name=isDefault]").each(function(idx){
		// 기본 카테고리 정보 넘기기
		if($("input[name=isDefault]:eq(" + idx + ")").val()=="y"){
			catename = $("input[name=arrdepthname]:eq(" + idx + ")").val();
			if($("input[name='isDefault']").length>1){
				catename += " 외 " + ($("input[name='isDefault']").length - 1) + "건"
			}else{
				catename += ""
			}
		}
		if(idx<1){
			arrcode = $("input[name=catecode]:eq(" + idx + ")").val();
			arrdepth = $("input[name=catedepth]:eq(" + idx + ")").val();
			arrdefault = $("input[name=isDefault]:eq(" + idx + ")").val();
		}else{
			arrcode += "," + $("input[name=catecode]:eq(" + idx + ")").val();
			arrdepth += "," + $("input[name=catedepth]:eq(" + idx + ")").val();
			arrdefault += "," + $("input[name=isDefault]:eq(" + idx + ")").val();
		}
	});
	jsontxt = catename + "!" + arrcode + "!" + arrdepth + "!" + arrdefault;
	fnAPPopenerJsCallClose("fnCategorySet(\"" + jsontxt + "\")");
}

function chgodr(hidediv,v,formname,formdata){
	if(hidediv!=''){
		if (v == 1){
			eval("$('#"+hidediv+"')").css("display","none");
		}else{
			eval("$('#"+hidediv+"')").css("display","");
		}
	}
	if(formname!=''){
		eval("$('#"+formname+"')").val(formdata);
	}
}
//-->
</script>

</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form name="cate" method="post" onsubmit="return false;" autocomplete="off">
		<input type="hidden" name="waititemid" value="<%=waititemid%>">
		<input type="hidden" name="mode" value="cate">
		<input type='hidden' name='catecode_a' value='' />
		<input type="hidden" name="selDispCateDiv" value="y">
		<input type="hidden" name="depthname">
		<input type="hidden" name="checkcate" value="0">
		<input type="hidden" name="arrcatecode" id="arrcatecode">
		<input type="hidden" name="arrcatedepth" id="arrcatedepth">
		<input type="hidden" name="arrisdefault" id="arrisdefault">
		<div class="content bgGry">
			<h1 class="hidden">카테고리 설정</h1>
			<div class="ctgySetting">
				<div class="selectBtn">
					<div class="grid2">
						<select name="depth1" id="depth1" onchange="startCategoryRequest('2depth',this.value)">
							<option value="">대분류</option>
							<% If cDisp.FResultCount > 0 Then %>
							<% For i=0 To cDisp.FResultCount-1 %>
							<option value="<%=cDisp.FItemList(i).FCateCode%>"><%=cDisp.FItemList(i).FCateName%></option>
							<% Next %>
							<% End If %>
						</select>
					</div>
					<div class="grid2">
						<select name="depth2" id="depth2" disabled="disabled" onChange="js2DepthSelectBox(this.value)">
							<option>중분류</option>
						</select>
					</div>
				</div>
				<div class="addBtn">
					<button type="button" class="btnB1 btnDkGry" disabled="disabled" onclick="sendDispCateItem();"><span class="itemAdd">추가</span></button>
				</div>
				<div class="addCtgyList">
					<ul id="tbl_DispCate"><% If waititemid<>"" Then %><% =getDispCategoryWait(waititemid) %><% End If %></ul>
				</div>
				<div ></div>
				<div class="linkNotice" id="HelpInfo" style="display:<% If waititemid<>"" Then %><% If getDispCategoryWaitCount(waititemid) > 0 Then %>none<% End If %><% End If %>">
					<p class="fs1-5r">카테고리 추가 후, 확인 버튼을 눌러주세요</p>
					<ul class="tMar1-5r">
						<li>추가된 카테고리는 더핑거스 웹사이트의 <br />해당 카테고리 리스트에 노출됩니다.</li>
						<li>(카테고리는 2개까지 선택할 수 있습니다.)</li>
					</ul>
				</div>
			</div>
		</div>
		</form>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
set cDisp = Nothing
%>
<script>
<!--
jQuery(document).ready(function(){
	fnAPPShowRightConfirmBtns();
});
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->