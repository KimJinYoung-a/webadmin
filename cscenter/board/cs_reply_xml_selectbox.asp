<%
'###########################################################
' Description : 1:1 상담
' History : 이상구 생성
'			2021.09.10 한용민 수정(이문재이사님요청 자사몰 필드추가, 소스표준화)
'###########################################################
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

// AJAX 프로그램
var xmlHttp;
var xmlDoc;
var xmlHttpMode;
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

function startRequest(mode, gubuncode, masteridx, sitename) {
	xmlHttpMode = mode;

	createXMLHttpRequest();
	xmlHttp.onreadystatechange = callback;
	xmlHttp.open("GET", "/cscenter/board/cs_reply_xml_response.asp?mode=" + mode + "&gubuncode=" + gubuncode + "&masteridx=" + masteridx + "&sitename=" + sitename, true);
	xmlHttp.send(null);
}

function callback() {
	if(xmlHttp.readyState == 4) {
        if(xmlHttp.status == 200) {
			// 정상적인 데이타 반환
            // 전체(TXT) : xmlHttp.responseText

			// 텍스트 앞부분에서 "<" 이전 문자들을 제거한다.(공백문자 제거용,  이렇게 안하면 변환이 안된다 --)
			var rawXML = xmlHttp.responseText;
			var filteredXML;
			var index = 0;
            for (var i = 0; i < rawXML.length; i++) {
                if (rawXML.charAt(i) == "<") {
                    index = i;
                    break;
                }
            }
            filteredXML = rawXML.substring(index);

            if (window.ActiveXObject) {
                // XML 로 변환한다.
                xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
                xmlDoc.loadXML(filteredXML);
            } else if (window.XMLHttpRequest) {
				if (xmlHttp.responseXML) {
					xmlDoc = xmlHttp.responseXML;
				} else {
					var parser = new DOMParser();
					xmlDoc = parser.parseFromString(filteredXML, 'text/xml');
				}
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
	var frm = eval("document.frm");
	var buf;
	var xmlItemCount = xmlDoc.getElementsByTagName("value1").length;

	if (xmlItemCount < 1) {
		return;
	}

	var selectObj, value1, value2, value3;
	if (xmlHttpMode == "replymaster") {
		frm.masteridx.length = xmlItemCount;

		for (i = 0; i < xmlItemCount; i++) {
			selectObj = frm.masteridx.options[i];
			value1 = xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			value2 = xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;
			// value3 = xmlDoc.getElementsByTagName("value3")[i].firstChild.nodeValue;

			selectObj.value = value1;
			selectObj.text = value2;
		}
	} else if (xmlHttpMode == "replydetail") {
		frm.detailidx.length = xmlItemCount;

		for (i = 0; i < xmlItemCount; i++) {
			selectObj = frm.detailidx.options[i];
			value1 = xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			value2 = xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;
			value3 = xmlDoc.getElementsByTagName("value3")[i].firstChild.nodeValue;

			selectObj.value = value3;
			selectObj.text = value2;
		}
	} else if (xmlHttpMode == "replysitename") {
		frm.sitenameidx.length = xmlItemCount;

		for (i = 0; i < xmlItemCount; i++) {
			selectObj = frm.sitenameidx.options[i];
			value1 = xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			value2 = xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;
			// value3 = xmlDoc.getElementsByTagName("value3")[i].firstChild.nodeValue;

			selectObj.value = value1;
			selectObj.text = value2;
		}
	}
}

function setSiteName(sitename) {
    var frm = eval("document.frm");
    if (sitename == '10x10') {
        frm.sitenameidx.selectedIndex = 1;
    } else {
        frm.sitenameidx.selectedIndex = 2;
    }
    requestSelectBoxmaster(frm.sitenameidx.value);
}

function requestSelectBoxsitename() {
	startRequest("replysitename", "0001", "", "");
}

//function requestSelectBoxMaster() {
//	startRequest("replymaster", "0001", "", "");
//}

function requestSelectBoxmaster(sitename) {
	var frm = eval("document.frm");
	if (sitename == "") {
		sitename = -1;
	}

	if (sitename == "XX") {
		return;
	}

    TnChangePrefaceNew(sitename == '10x10' ? '00' : '55');

	startRequest("replymaster", "0001", "", sitename);
	$("#detailidx option").remove();
}

function requestSelectBoxDetail(masteridx) {
	if (masteridx == "") {
		masteridx = -1;
	}

	if (masteridx == "XX") {
		return;
	}

	startRequest("replydetail", "0001", masteridx, "");
}

</script>

<select class="select" id="sitenameidx" name="sitenameidx" onchange="requestSelectBoxmaster(this.value)" ></select>
<select class="select" id="masteridx" name="masteridx" onchange="requestSelectBoxDetail(this.value)" ></select>
<select class="select" id="detailidx" name="detailidx" onchange="fnSelectBoxDetailSelected(this.value)" ></select>
