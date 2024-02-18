<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  cs �޸�
' History : 2007.10.26 �̻� ����
'           2016.12.07 �ѿ�� ����
'###########################################################
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

		//����Ʈ��
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
			����:<select class="select" name="work_type" onchange="startRequest('work_target',xmlHttpParam1,this.value,'',''); hideFrame();" style="width:120px;" ></select>
		</td>
		<td>
			���л�:<select class="select" name="work_target" onchange="changeWorkType();" style="width:120px;"><option value=""></option></select>
		</td>
	</tr>
</table>
