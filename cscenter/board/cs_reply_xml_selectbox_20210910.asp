<script type="text/javascript">
// AJAX ���α׷�
var xmlHttp;
var xmlDoc;
var xmlHttpMode;
var xmlHttpDefaultSet;
var xmlProcessId = 0;

function Trim(str){
	return str.replace(/\s/g,""); // \ -> �������� �Դϴ�.
}


function createXMLHttpRequest() {
	if (window.ActiveXObject) {
		xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
	} else if (window.XMLHttpRequest) {
		xmlHttp = new XMLHttpRequest();
	}
}

function startRequest(mode, gubuncode, masteridx) {
	xmlHttpMode = mode;

	createXMLHttpRequest();
	xmlHttp.onreadystatechange = callback;
	xmlHttp.open("GET", "/cscenter/board/cs_reply_xml_response.asp?mode=" + mode + "&gubuncode=" + gubuncode + "&masteridx=" + masteridx, true);
	xmlHttp.send(null);
}

function callback() {
	if(xmlHttp.readyState == 4) {
        if(xmlHttp.status == 200) {
			// �������� ����Ÿ ��ȯ
            // ��ü(TXT) : xmlHttp.responseText

			// �ؽ�Ʈ �պκп��� "<" ���� ���ڵ��� �����Ѵ�.(���鹮�� ���ſ�,  �̷��� ���ϸ� ��ȯ�� �ȵȴ� --)
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
                // XML �� ��ȯ�Ѵ�.
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
            // �����Ͱ� �������� ���� ���
            alert("����Ÿ�� �������� �ʽ��ϴ�.(CODE : 200)");
        } else if (xmlHttp.status == 500){
            // �����߻���
            alert("����Ÿ ������ ������ �߻��Ͽ����ϴ�.(CODE : 500)");
        }
    }
}

// ���⸸ �����Ѵ�. �ش� ���������� ajax �� �̿��� ���� ����Ÿ�� �������� ǥ���Ѵ�.
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
	}
}

function requestSelectBoxMaster() {
	startRequest("replymaster", "0001", "");
}

function requestSelectBoxDetail(masteridx) {
	if (masteridx == "") {
		masteridx = -1;
	}

	if (masteridx == "XX") {
		return;
	}

	startRequest("replydetail", "0001", masteridx);
}

</script>

&nbsp;
<select class="select" name="masteridx" onchange="requestSelectBoxDetail(this.value)"  ></select>
&nbsp;
<select class="select" name="detailidx" onchange="fnSelectBoxDetailSelected(this.value)" ></select>
&nbsp;

<script language='javascript'>
/*
// ���� �������� �Ʒ� ������ �־�� �Ѵ�.
document.onload = getOnload();

function getOnload(){
	requestSelectBoxMaster();
}

function fnSelectBoxDetailSelected(v) {
	// do something;
}

*/
</script>
