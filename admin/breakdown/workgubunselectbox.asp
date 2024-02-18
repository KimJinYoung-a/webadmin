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

function startRequest( mode, work_part_sn, work_type, work_target ) {
	xmlHttpMode = mode;
	xmlHttpParam1 = work_part_sn;
	xmlHttpParam2 = work_type;
	xmlHttpParam3 = work_target;

	if (mode === "work_target" && work_type === "") {
		return;
	}

	// alert("mode = " + mode + ", work_part_sn = " + work_part_sn + ", work_type = " + work_type + ", work_target = " + work_target);

    createXMLHttpRequest();
    xmlHttp.onreadystatechange = callback;
    xmlHttp.open("GET", "/admin/breakdown/workgubunselectbox_response.asp?mode=" + mode + "&param1=" + work_part_sn + "&param2=" + work_type + "&param3=" + work_target, true);
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
	var frm = eval("document." + parentFrmName);
	var buf;
	var length = xmlDoc.getElementsByTagName("value1").length;

	if (xmlHttpMode === "work_type") {
		frm.work_type.length = (length*1+1);

		frm.work_type.options[0].text = "";
		frm.work_type.options[0].value = "";
		frm.work_type.options[0].selected = true;

		for (i=0;i<length;i++){
			frm.work_type.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.work_type.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam2){
				frm.work_type.options[i + 1].selected = true;
			}
		}

		//디폴트값
		if (xmlHttpParam1!="") { startRequest('work_target',xmlHttpParam1,xmlHttpParam2,xmlHttpParam3); }
	}else if (xmlHttpMode=="work_target"){
		frm.work_target.length = (length*1 + 1);

		frm.work_target.options[0].text = "";
		frm.work_target.options[0].value = "";
		frm.work_target.options[0].selected = true;

		for (i=0;i<length;i++){
			frm.work_target.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue.replace("XX", "");
			frm.work_target.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam3){
				frm.work_target.options[i + 1].selected = true;
			}
		}
		if ((xmlHttpParam2=="")&&(frm.work_target.length>0)) frm.work_target.options[0].selected = true;
	}
}

</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td width="200">
			구분:<select class="select" name="work_type" onchange="startRequest('work_target',xmlHttpParam1,this.value,'',''); hideFrame();" style="width:120px;" ></select>
		</td>
		<td>
			구분상세:<select class="select" name="work_target" onchange="changeWorkType();" style="width:120px;"><option value=""></option></select>
		</td>
	</tr>
</table>
