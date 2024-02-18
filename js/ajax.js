// AJAX ���α׷�
//
// ����
// ������������ �񵿱������� ����� �������� ������ ǥ���Ҷ� ����Ѵ�.
// �Ʒ��� ������ ǥ���Ѵ�.
// a.asp �� ������������̸�, b.asp �� ������������� ǥ���� ����Ÿ�� xml �� �����ִ� �����������̴�.
//
// startAjaxSample(param1) �� startAjaxSample(param1,param2, param3) ���� ���·� ������ �����ϴ�.
// ����� �ҷ��� URL �� �����ϰ�( initializeURL() ), ���Ͻ� ������ ��ũ��Ʈ�Լ����� �����ϸ�( initializeReturnFunction() ), AJAX �� �����Ѵ�.( startRequest() )
//
// ���Ͻ� ������ ��ũ��Ʈ�Լ����� ����Ʈ�� processAjax() ��, initializeReturnFunction() �� �̿��Ͽ� ���� �����ϴ�.
// ����� xmlDoc �� ���õ� ����Ÿ�� �����ȭ�鿡 ǥ���Ѵ�.
//
// �߰� : 200ms �� ������ �߰�(select �ױ� �Ǵ� ����Ŭ��� ��ó)
// �߰� : Ư������ ó��(����Ÿ�� �޾ƿ��� ������(�Ʒ� ���ÿ����� b.asp)���� <,>,& �� ����, &lt;,&gt;,&amp; ������ ��ȯ�Ͽ� �����Ѵ�.)
// �߰� : ���Ͻ� ������ ��ũ��Ʈ�Լ����� �����Ҽ� �ִ�.(���������� �������� ajax �� ���ɼ� �ִ�)
// �߰� : �����ڵ带 Ȯ���Ҽ� �ִ� ����� �����Ѵ�. ����Ʈ�� alert() �� �����Ѵ�.
// �߰� : �����̸� 100ms �� ���δ�.
//
// ����
/*

a.asp
// ============================================================================
<script language="JavaScript" src="/js/ajax.js"></script>
<script>
function startAjaxSample(param1) {
        initializeURL("b.asp?id=" + param1);
        initializeReturnFunction("processAjaxSample()");
        initializeErrorFunction("onErrorAjaxSample()");
        startRequest();
}

function processAjaxSample() {
        // ���� �ؽ�Ʈ
        // xmlHttp.responseText

        var length = xmlDoc.getElementsByTagName("value1").length;
        if (length < 1) {
                alert("����Ÿ�� �������� �ʽ��ϴ�.");
                return;
        }

        for (var i = 0; i < length; i++) {
                alert(xmlDoc.getElementsByTagName("value1")[i].firstChild.nodeValue);
                alert(xmlDoc.getElementsByTagName("value2")[i].firstChild.nodeValue);
        }
}

function onErrorAjaxSample() {
        alert("ERROR : " + xmlHttp.status);
}
</script>

b.asp
// ============================================================================
<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<?xml version="1.0"  encoding="euc-kr"?>
<response>
  <item>
    <value1>111</value1>
    <value2>Ư��111+_)(*</value2>
  </item>
  <item>
    <value1>444</value1>
    <value2>Ư��,./":;'}{][</value2>
  </item>
  <item>
    <value1>222&gt;</value1>
    <value2>Ư��$#@!</value2>
  </item>
  <item>
    <value1>222&gt;</value1>
    <value2>Ư��^%&amp;</value2>
  </item>
  <item>
    <value1>222&gt;</value1>
    <value2>�ѱ�222&lt;</value2>
  </item>
  <item>
    <value1>333</value1>
    <value2>Ư��333~`-=\|?</value2>
  </item>
</response>
*/

// ============================================================================
var xmlActionURL = "";
var xmlReturnFunction = "processAjax()";
var xmlErrorFunction = "";
var xmlProcessId = 0;
var xmlHttp;
var xmlDoc;

function initializeURL(url) {
        xmlActionURL = url;
}

function initializeReturnFunction(func) {
        xmlReturnFunction = func;
}

function initializeErrorFunction(errfunc) {
        xmlErrorFunction = errfunc;
}

function createXMLHttpRequest() {
        if (window.ActiveXObject) {
                xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
        } else if (window.XMLHttpRequest) {
                xmlHttp = new XMLHttpRequest();
        }
}

function doStartRequest(processid) {
        if (xmlProcessId != processid) {
                return;
        }
        createXMLHttpRequest();
        xmlHttp.onreadystatechange = callback;
        xmlHttp.open("GET", xmlActionURL, true);
        xmlHttp.send(null);
        // alert("" + processid + " ��° ��ûó��");
}

function startRequest() {
        // ����Ŭ�� or <select> ������ �ٷ��� ��û�� ���� ó��
        // 100ms �� Ÿ�̸Ӹ� �ΰ� �ش�ð����� ��û�� ������ ���� ��û ����
        if (xmlProcessId >= 10000) {
                xmlProcessId = 1;
        } else {
                xmlProcessId = xmlProcessId + 1;
        }
        setTimeout("doStartRequest(" + xmlProcessId + ")", 100);
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

                        eval(xmlReturnFunction);
                } else if (xmlHttp.status == 204){
                        // �����Ͱ� �������� ���� ���
                        if (xmlErrorFunction == "") {
                                alert("����Ÿ�� �������� �ʽ��ϴ�.(CODE : " + xmlHttp.status + ")");
                        } else {
                                eval(xmlErrorFunction);
                        }
                } else if (xmlHttp.status == 500){
                        // �����߻���
                        if (xmlErrorFunction == "") {
                                alert("����Ÿ ������ ������ �߻��Ͽ����ϴ�.(CODE : " + xmlHttp.status + ")");
                        } else {
                                eval(xmlErrorFunction);
                        }
                }
        }
}