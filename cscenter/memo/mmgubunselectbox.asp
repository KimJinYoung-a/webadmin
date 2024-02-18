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
<%
dim tmp_mmgubun, tmp_qadiv
if tmp_mmgubun="" then	tmp_mmgubun = request("mmgubun")
if tmp_qadiv="" then	tmp_qadiv = request("qadiv")
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
                    // �������� ����Ÿ ��ȯ
                    // ��ü(TXT) : xmlHttp.responseText
                    //if (window.ActiveXObject) {
                    if ("ActiveXObject" in window) {
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

	if (xmlHttpMode=="mmgubun"){
		frm.mmgubun.length = (length*1+1);

		for (i=0;i<length;i++){
			frm.mmgubun.options[i + 1].value= xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue;
			frm.mmgubun.options[i + 1].text= xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue;

			if (xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue==xmlHttpParam1){
				frm.mmgubun.options[i + 1].selected = true;
			}
		}

		//����Ʈ��
		if (xmlHttpParam1!="") { startRequest('qadiv',xmlHttpParam1,xmlHttpParam2,xmlHttpParam3); }
	}else if (xmlHttpMode=="qadiv"){
		frm.qadiv.length = (length*1 + 1);

		if (length == 0) {
			frm.qadiv.options[0].text = "������ �����ϼ���";
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

����:<select class="select" name="mmgubun" onchange="startRequest('qadiv',this.value,'','')"  ></select>
���л�:<select class="select" name="qadiv" ><option value="">������ �����ϼ���</option></select>

<script type="text/javascript">
/*
document.onload = getOnload();

function getOnload(){
	startRequest('mmgubun','<%= tmp_mmgubun %>','<%= tmp_qadiv %>');
}
*/
</script>
