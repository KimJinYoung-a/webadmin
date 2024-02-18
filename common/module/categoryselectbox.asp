<%
dim tmp_cdl, tmp_cdm, tmp_cds
if tmp_cdl="" then	tmp_cdl = request("cdl")
if tmp_cdm="" then	tmp_cdm = request("cdm")
if tmp_cds="" then	tmp_cds = request("cds")
%>
<script type="text/javascript">
// AJAX ���α׷�
var parentFrmName = "frm";
var xmlHttp;
var xmlDoc;
var xmlHttpMode, xmlHttpParam1, xmlHttpParam2, xmlHttpParam3;
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
		//����Ʈ��
		if (xmlHttpParam1!="") { startRequest('cdm',xmlHttpParam1,xmlHttpParam2,xmlHttpParam3); }
		frm.cdl.options[0].text = "-��ü-"
		frm.cdl.options[0].value = ""
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
		if ((xmlHttpParam3=="")&&(frm.cds.length>0)) frm.cds.options[0].selected = true;

		//����Ʈ��
		if (xmlHttpParam2!="") { startRequest('cds',xmlHttpParam1,xmlHttpParam2,xmlHttpParam3); }
		frm.cdm.options[0].text = "-��ü-"
		frm.cdm.options[0].value = ""
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
		frm.cds.options[0].text = "-��ü-"
		frm.cds.options[0].value = ""
	}
}

function selectCategory(frm){
	if ((frm.cdl.selectedIndex<0)||(frm.cdm.selectedIndex<0)||(frm.cds.selectedIndex<0)){
		alert('ī�װ��� ���ܰ� ��� �������ּ���.');
		return;
	}

	var cd1 = frm.cdl[frm.cdl.selectedIndex].value;
	var cd2 = frm.cdm[frm.cdm.selectedIndex].value;
	var cd3 = frm.cds[frm.cds.selectedIndex].value;

	var cd1name = frm.cdl[frm.cdl.selectedIndex].text;
	var cd2name = frm.cdm[frm.cdm.selectedIndex].text;
	var cd3name = frm.cds[frm.cds.selectedIndex].text;

	if ((cd1=="")||(cd2=="")||(cd3=="")){
		alert('ī�װ��� ���ܰ� ��� �������ּ���.');
		return;
	}

	opener.setCategory(cd1,cd2,cd3,cd1name,cd2name,cd3name);
	window.close();
}
</script>

	ī�װ� :
	<select class="select" name="cdl" onchange="startRequest('cdm',this.value,'','')"  ></select>
	<select class="select" name="cdm" onchange="startRequest('cds',eval(parentFrmName).cdl.value,this.value,'')"  ></select>
	<select class="select" name="cds"  ></select>

<script language='javascript'>
document.onload = getOnload();

function getOnload(){
	startRequest('cdl','<%= tmp_cdl %>','<%= tmp_cdm %>','<%= tmp_cds %>');
}
</script>