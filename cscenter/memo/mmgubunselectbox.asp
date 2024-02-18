<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  cs 메모
' History : 2007.10.26 이상구 생성
'           2016.12.07 한용민 수정
'###########################################################
%>
<%
dim tmp_mmgubun, tmp_qadiv
if tmp_mmgubun="" then	tmp_mmgubun = request("mmgubun")
if tmp_qadiv="" then	tmp_qadiv = request("qadiv")
%>

<script type="text/javascript">

// AJAX 프로그램
var parentFrmName = "frm";
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

function startRequest( mode,mmgubun,qadiv) {
		xmlHttpMode = mode;
		xmlHttpParam1 = mmgubun;
		xmlHttpParam2 = qadiv;

		//alert('mode=' + mode + ',mmgubun=' + mmgubun + ',qadiv=' + qadiv);
        createXMLHttpRequest();
        xmlHttp.onreadystatechange = callback;
        xmlHttp.open("GET", "/cscenter/memo/mmgubunselectbox_response.asp?mode=" + mode + "&param1=" + mmgubun + "&param2=" + qadiv, true);
        xmlHttp.send(null);
}

function callback() {
	if(xmlHttp.readyState == 4) {
            if(xmlHttp.status == 200) {
                    // 정상적인 데이타 반환
                    // 전체(TXT) : xmlHttp.responseText
                    //if (window.ActiveXObject) {
                    if ("ActiveXObject" in window) {
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
                        var parser = new DOMParser();
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
                        xmlDoc = parser.parseFromString(filteredML, "text/xml");
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

	if (xmlHttpMode=="mmgubun"){
		frm.mmgubun.length = (length*1+1);

		for (i=0;i<length;i++){
			frm.mmgubun.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.mmgubun.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam1){
				frm.mmgubun.options[i + 1].selected = true;
			}
		}

		//디폴트값
		if (xmlHttpParam1!="") { startRequest('qadiv',xmlHttpParam1,xmlHttpParam2,xmlHttpParam3); }
	}else if (xmlHttpMode=="qadiv"){
		frm.qadiv.length = (length*1 + 1);

		if (length == 0) {
			frm.qadiv.options[0].text = "구분을 선택하세요";
		} else {
			frm.qadiv.options[0].text = "";
		}

		for (i=0;i<length;i++){
			frm.qadiv.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue.replace("XX", "");
			frm.qadiv.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam2){
				frm.qadiv.options[i + 1].selected = true;
			}
		}
		if ((xmlHttpParam2=="")&&(frm.qadiv.length>0)) frm.qadiv.options[0].selected = true;
	}
}

</script>

구분:<select class="select" name="mmgubun" onchange="startRequest('qadiv',this.value,'','')"  ></select>
구분상세:<select class="select" name="qadiv" ><option value="">구분을 선택하세요</option></select>

<script type="text/javascript">
/*
document.onload = getOnload();

function getOnload(){
	startRequest('mmgubun','<%= tmp_mmgubun %>','<%= tmp_qadiv %>');
}
*/
</script>
