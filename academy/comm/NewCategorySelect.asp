<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim iid, dftDiv
iid     = RequestCheckvar(request("iid"),10)
dftDiv  = RequestCheckvar(request("dftDiv"),1)

%>
<script type="text/javascript">
// AJAX ���α׷�
var parentFrmName = "frmcate";
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
        xmlHttp.open("GET", "/academy/comm/NewCategory_Action_response.asp?mode=" + mode + "&param1=" + cdl + "&param2=" + cdm + "&param3=" + cds, true);
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

	opener.addCateItem(cd1,cd1name,cd2,cd2name,cd3,cd3name,frm.div.value);
	self.close();
}
</script>
<body bgcolor="#F4F4F4">
<!-- �ش� -->
<table width="630" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
		<tr>
			<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
				<font color="#333333"><b>ī�װ� ����/�߰�</b></font>
			</td>
			<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:5; border-bottom:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999" bgcolor="#FFFFFF">
		<b>��ǰ�� ���� ī�װ��� �����մϴ�.</b><br>
		���ʺ��� ��з�, �ߺз�, �Һз��̸� ���ʴ�� ���ð����մϴ�.<br>
		<font color="#E08050">�� ī�װ� ������ [�⺻] �� [�߰�]�� ������, <b>�⺻ ī�װ��� �Ѱ����� ����<b>�����մϴ�.</font>

	</td>
</tr>
</table>
<!-- ���� -->
<table width="630" border="0" cellspacing="3" cellpadding="0" align="center">
<form name="frmcate">
<tr><td colspan="2" height="5"></td></tr>
<tr>
	<td colspan="2" align="center">
		<select name="cdl" onchange="startRequest('cdm',this.value,'','')" class="textarea" style='width:200;' size="15">
	
		</select>
		<select name="cdm" onchange="startRequest('cds',eval(parentFrmName).cdl.value,this.value,'')" class="textarea"  style='width:200;' size="15">
	
		</select>
		<select name="cds"  style='width:200;' class="textarea"  size="15">
	
		</select>
	</td>
</tr>
<tr><td colspan="2" height="5"></td></tr>
<tr>
	<td align="left">
		<input type="button" class="button" value="â�ݱ�" onclick="self.close()">
	</td>
	<td align="right" style="font-size:10pt;">
		ī�װ� ����
		<select name="div" class="select">
			<option value="D" <%= ChkIIF(dftDiv="D","selected","") %> >�⺻ ī�װ�</optioN>
			<option value="A" <%= ChkIIF(dftDiv<>"D","selected","") %> >�߰� ī�װ�</optioN>
		</select>
		<input type="button" class="button" value="ī�װ�����" onclick="selectCategory(frmcate);">
	</td>
</tr>
</form>
</table>
</body>
<script language='javascript'>

function getOnload(){
	startRequest('cdl','','','');
}

function getOnUnload(){
    xmlHttp = null;
    xmlDoc  = null;
}

window.onload = getOnload;
window.onunload = getOnUnload;

</script>
<!-- #include virtual="/common/lib/poptail.asp"-->