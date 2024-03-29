<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 
' History : 최초생성자모름
'			2017.04.10 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim cdl, cdm, cds
cdl = requestCheckVar(request("cdl"),3)
cdm = requestCheckVar(request("cdm"),3)
cds = requestCheckVar(request("cds"),3)

%>
<script type="text/javascript">
// AJAX 프로그램
var parentFrmName = "frmcate";
var xmlHttp;
var xmlDoc;
var xmlHttpMode, xmlHttpParam1, xmlHttpParam2, xmlHttpParam3;
var xmlHttpDefaultSet;
var xmlProcessId = 0;

function Trim(str){
 return str.replace(/\s/g,""); // \ -> 역슬래쉬 입니다.
}


function createXMLHttpRequest() {
        if (window.ActiveXObject) {
                xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
        } else if (window.XMLHttpRequest) {
                xmlHttp = new XMLHttpRequest();
        }
}

function startRequest( mode,cdl,cdm,cds) {

		xmlHttpMode = mode;
		xmlHttpParam1 = cdl;
		xmlHttpParam2 = cdm;
		xmlHttpParam3 = cds;


		//alert('mode=' + mode + ',cdl=' + cdl + ',cdm=' + cdm + ',cds=' + cds);
        createXMLHttpRequest();
        xmlHttp.onreadystatechange = callback;
        xmlHttp.open("GET", "/common/module/normal_action_response.asp?mode=" + mode + "&param1=" + cdl + "&param2=" + cdm + "&param3=" + cds, true);
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

// 여기만 변경한다. 해당 페이지에서 ajax 를 이용해 받은 데이타를 페이지에 표시한다.
function process() {
	var frm = eval("document." + parentFrmName);
	var buf;
	var length = xmlDoc.getElementsByTagName("value1").length;

	if (xmlHttpMode=="cdl"){
		frm.cdl.length = (length*1+1);

		for (i=0;i<length;i++){
			frm.cdl.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.cdl.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam1){
				frm.cdl.options[i + 1].selected = true;
			}
		}

		//디폴트값
		if (xmlHttpParam2!="") { startRequest('cdm',xmlHttpParam1,xmlHttpParam2,xmlHttpParam3); }
	}else if (xmlHttpMode=="cdm"){
		frm.cdm.length = (length*1 + 1);
		frm.cds.length = 1;
		for (i=0;i<length;i++){
			frm.cdm.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.cdm.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam2){
				frm.cdm.options[i + 1].selected = true;
			}
		}
		if ((xmlHttpParam2=="")&&(frm.cdm.length>0)) frm.cdm.options[0].selected = true;
		if ((xmlHttpParam2=="")&&(frm.cds.length>0)) frm.cds.options[0].selected = true;

		//디폴트값
		if (xmlHttpParam3!="") { startRequest('cds',xmlHttpParam1,xmlHttpParam2,xmlHttpParam3); }
	}else if (xmlHttpMode=="cds"){
		frm.cds.length = (length*1 + 1);

		for (i=0;i<length;i++){
			frm.cds.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.cds.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam3){
				frm.cds.options[i + 1].selected = true;
			}
		}
		if ((xmlHttpParam3=="")&&(frm.cds.length>0)) frm.cds.options[0].selected = true;
	}
}

function selectCategory(frm){
	if ((frm.cdl.selectedIndex<0)||(frm.cdm.selectedIndex<0)||(frm.cds.selectedIndex<0)){
		alert('카테고리를 세단계 모두 선택해주세요.');
		return;
	}

	var cd1 = frm.cdl[frm.cdl.selectedIndex].value;
	var cd2 = frm.cdm[frm.cdm.selectedIndex].value;
	var cd3 = frm.cds[frm.cds.selectedIndex].value;

	var cd1name = frm.cdl[frm.cdl.selectedIndex].text;
	var cd2name = frm.cdm[frm.cdm.selectedIndex].text;
	var cd3name = frm.cds[frm.cds.selectedIndex].text;

	if ((cd1=="")||(cd2=="")||(cd3=="")){
		alert('카테고리를 세단계 모두 선택해주세요.');
		return;
	}

	opener.setCategory(cd1,cd2,cd3,cd1name,cd2name,cd3name);
	window.close();
}
</script>
<body >
<table width="640" border="0" cellspacing="3" cellpadding="0" align="center">
<form name="frmcate">
<tr>
	<td>
	<select name="cdl" onchange="startRequest('cdm',this.value,'','')"  style='width:200;' size="15">

	</select>
	<select name="cdm" onchange="startRequest('cds',eval(parentFrmName).cdl.value,this.value,'')"  style='width:200;' size="15">

	</select>
	<select name="cds"  style='width:200;'  size="15">

	</select>

	</td>
</tr>
<tr>
	<td align="center"><input type="button" value="카테고리선택" onclick="selectCategory(frmcate);"></td>
</tr>
</form>
</table>
</body>
<script language='javascript'>
document.onload = getOnload();

function getOnload(){
    startRequest('cdl','','','');
    
	//startRequest('cdl','<%= cdl %>','<%= cdm %>','<%= cds %>');
	//startRequest('cdl','<%= cdl %>','','');
}
</script>
<!-- #include virtual="/common/lib/poptail.asp"-->