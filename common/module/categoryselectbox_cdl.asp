<%
'//�¶��� �������� ����
dim tmp_cdl, tmp_cdm
if tmp_cdl="" then	tmp_cdl = request("selC")
if tmp_cdm="" then	tmp_cdm = request("selCM")
%>
<script type="text/javascript">
// AJAX ���α׷�
var parentFrmName = "frm";
var xmlHttp;
var xmlDoc;
var xmlHttpMode, xmlHttpParam1, xmlHttpParam2;
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

function startRequest( mode,cdl,cdm) {

	if(cdl!="110") {
		document.frm.selCM.value="";
		document.frm.selCM.style.display="none";
	} else {
		document.frm.selCM.style.display="";
	}
	if(mode=="cdl"||mode=="cdm"&&cdl=="110") {
		xmlHttpMode = mode;
		xmlHttpParam1 = cdl;
		xmlHttpParam2 = cdm;

		//alert('mode=' + mode + ',cdl=' + cdl + ',cdm=' + cdm);
        createXMLHttpRequest();
        xmlHttp.onreadystatechange = callback;
        xmlHttp.open("GET", "/common/module/normal_action_response.asp?mode=" + mode + "&param1=" + cdl + "&param2=" + cdm, true);
        xmlHttp.send(null);
	}
}

function callback() {
	if(xmlHttp.readyState == 4) {
            if(xmlHttp.status == 200) {
                    // �������� ����Ÿ ��ȯ
                    // ��ü(TXT) : xmlHttp.responseText
                    if (window.ActiveXObject) {
                            // XML �� ��ȯ�Ѵ�.
                            // �ؽ�Ʈ �պκп��� "<" ���� ���ڵ��� �����Ѵ�.(���鹮�� ���ſ�,  �̷��� ���ϸ� ��ȯ�� �ȵȴ� --)
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
	var frm = eval("document." + parentFrmName);
	frm = document.frm;
	var buf;
	var length = xmlDoc.getElementsByTagName("value1").length;

	if (xmlHttpMode=="cdl"){
		frm.selC.length = (length*1+1);

		for (i=0;i<length;i++){
			frm.selC.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.selC.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam1){
				frm.selC.options[i + 1].selected = true;
			}
		}

		//����Ʈ��
		if (xmlHttpParam1!="") { startRequest('cdm',xmlHttpParam1,xmlHttpParam2); }
	}else if (xmlHttpMode=="cdm"){
		frm.selCM.length = (length*1 + 1);
		for (i=0;i<length;i++){
			frm.selCM.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.selCM.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam2){
				frm.selCM.options[i + 1].selected = true;
			}
		}
		if ((xmlHttpParam2=="")&&(frm.selCM.length>0)) frm.selCM.options[0].selected = true;
	}
}

function selectCategory(frm){
	if ((frm.selC.selectedIndex<0)||(frm.selCM.selectedIndex<0)){
		alert('ī�װ��� ���ܰ� ��� �������ּ���.');
		return;
	}

	var cd1 = frm.selC[frm.selC.selectedIndex].value;
	var cd2 = frm.selCM[frm.selCM.selectedIndex].value;

	var cd1name = frm.selC[frm.selC.selectedIndex].text;
	var cd2name = frm.selCM[frm.selCM.selectedIndex].text;

	if ((cd1=="")||(cd2=="")){
		alert('ī�װ��� ���ܰ� ��� �������ּ���.');
		return;
	}

	opener.setCategory(cd1,cd2,cd1name,cd2name);
	window.close();
}
</script>

	ī�װ� :
	<select class="select" name="selC" onchange="startRequest('cdm',this.value,'')"  ></select>
	<select class="select" name="selCM" style="display:none"></select>

<script language='javascript'>
document.onload = getOnload();

function getOnload(){
	startRequest('cdl','<%= tmp_cdl %>','<%= tmp_cdm %>');
}
</script>