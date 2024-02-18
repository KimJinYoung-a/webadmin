<%@ codepage="65001" language="VBScript" %>
<!-- #include virtual="/lib/util/commlib.asp" -->
<%
response.charset = "utf-8"

const C_DIMX = 10
redim protocolNM(C_DIMX)
redim protocolJson(C_DIMX)
redim protocolURL(C_DIMX)
redim protocolComment(C_DIMX)
redim protocolFix(C_DIMX)

dim protocolVersion : protocolVersion="1"   ''프로토콜 버전 1

dim DFTProto : DFTProto = "'OS':'ios','versioncode':'70','versionname':'1.69','version':'"&protocolVersion&"'"
protocolNM(0) = "firstconnection"
protocolJson(0) = protocolJson(0) + "'type':'firstconnection'"
protocolJson(0) = protocolJson(0) + ",'pushid':'12345'"
protocolJson(0) = protocolJson(0) + ",'uuid':'06081FFF'"
protocolJson(0) = protocolJson(0) + ",'idfa':'ac601e3178b8'"
protocolURL(0) = "/apps/protoV1/startupProc.asp"
protocolComment(0) = ""
protocolFix(0) = 1

protocolNM(1) = "1차 login protocol"
protocolJson(1) = protocolJson(1) + "'type':'login'"
protocolJson(1) = protocolJson(1) + ",'id':'fingertest01'"
protocolJson(1) = protocolJson(1) + ",'pass':'cube1010'"
protocolJson(1) = protocolJson(1) + ",'pushid':'12345'"
protocolJson(1) = protocolJson(1) + ",'uuid':'06081FFF'"
protocolJson(1) = protocolJson(1) + ",'idfa':'ac601e3178b8'"
protocolJson(1) = protocolJson(1) + ",'pass2':''"
protocolJson(1) = protocolJson(1) + ",'saved_id':'Y'"
protocolURL(1) = "/apps/protoV1/loginProc.asp"
protocolComment(1) = ""
protocolFix(1) = 1

protocolNM(2) = "2차 login protocol"
protocolJson(2) = protocolJson(2) + "'type':'login'"
protocolJson(2) = protocolJson(2) + ",'id':'fingertest01'"
protocolJson(2) = protocolJson(2) + ",'pass':'cube1010'"
protocolJson(2) = protocolJson(2) + ",'pushid':'12345'"
protocolJson(2) = protocolJson(2) + ",'uuid':'06081FFF'"
protocolJson(2) = protocolJson(2) + ",'idfa':'ac601e3178b8'"
protocolJson(2) = protocolJson(2) + ",'pass2':'vldrjtm01'"
protocolJson(2) = protocolJson(2) + ",'saved_id':'Y'"
protocolURL(2) = "/apps/protoV1/loginProc.asp"
protocolComment(2) = ""
protocolFix(2) = 1

protocolNM(3) = "logout protocol"
protocolJson(3) = protocolJson(3) + "'type':'logout'"
protocolJson(3) = protocolJson(3) + ",'pushid':'12345'"
protocolJson(3) = protocolJson(3) + ",'uuid':'06081FFF'"
protocolJson(3) = protocolJson(3) + ",'idfa':'ac601e3178b8'"
protocolURL(3) = "/apps/protoV1/logoutProc.asp"
protocolComment(3) = ""
protocolFix(3) = 1

protocolNM(4) = "2차 비밀번호 삭제"
protocolJson(4) = protocolJson(4) + "'type':'set2ndpassdel'"
protocolJson(4) = protocolJson(4) + ",'id':'fingertest01'"
protocolURL(4) = "/apps/protoV1/set2ndpassdel.asp"
protocolComment(4) = "테스트용 임시 페이지"
protocolFix(4) = 1

protocolNM(5) = "2차 비밀번호 설정 protocol"
protocolJson(5) = protocolJson(5) + "'type':'set2ndpass'"
protocolJson(5) = protocolJson(5) + ",'id':'fingertest01'"
protocolJson(5) = protocolJson(5) + ",'pass2':'vldrjtm01'"
protocolJson(5) = protocolJson(5) + ",'saved_id':'Y'"
protocolURL(5) = "/apps/protoV1/set2ndpass.asp"
protocolComment(5) = ""
protocolFix(5) = 1

protocolNM(6) = "push id protocol"
protocolJson(6) = protocolJson(6) + "'type':'reg'"
protocolJson(6) = protocolJson(6) + ",'pushid':'12345'"
protocolJson(6) = protocolJson(6) + ",'uuid':'06081FFF'"
protocolJson(6) = protocolJson(6) + ",'idfa':'ac601e3178b8'"
protocolURL(6) = "/apps/protoV1/deviceProc.asp"
protocolComment(6) = ""
protocolFix(6) = 1

protocolNM(7) = "push yn protocol"
protocolJson(7) = protocolJson(7) + "'type':'adpush'"
protocolJson(7) = protocolJson(7) + ",'pushid':'12345'"
protocolJson(7) = protocolJson(7) + ",'uuid':'06081FFF'"
protocolJson(7) = protocolJson(7) + ",'idfa':'ac601e3178b8'"
protocolJson(7) = protocolJson(7) + ",'notiyn':'A'"
protocolURL(7) = "/apps/protoV1/pushYnProc.asp"
protocolComment(7) = ""
protocolFix(7) = 1

protocolNM(8) = "dashboard protocol"
protocolJson(8) = protocolJson(8) + "'type':'dashboard'"
protocolURL(8) = "/apps/protoV1/dashboardProc.asp"
protocolComment(8) = "로그인 후 테스트 가능합니다.(1,2차 로그인 후 테스트)"
protocolFix(8) = 1

protocolNM(9) = "notice list protocol"
protocolJson(9) = protocolJson(9) + "'type':'notilist'"
protocolJson(9) = protocolJson(9) + ",'movepage':1"
protocolURL(9) = "/apps/protoV1/noticeListProc.asp"
protocolComment(9) = "로그인 후 테스트 가능합니다."
protocolFix(9) = 1

dim i
%>
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<style type="text/css">
body {font-size: 9pt; font-family: "굴림";color:#000000}
table { border:1px solid gray }
</style>
<script language='javascript'>
function jREQ(i){
    var frm = document.frmjson;
    var iURL = frm.jurl[i].value
    var jsonVal = frm.json[i].value
    
   
    document.smFrm.json.value=jsonVal;
    document.smFrm.action=iURL;
    document.smFrm.target="ifrmtarget";
    
    document.smFrm.submit();
}

function jscustom(src){
    document.location.href=src;
}

function jsCallbackFunc(retval){
    alert(retval);
}

function FnGotoBrand(v){
	$.ajax({
		url: '/apps/appCom/wish/webview/lib/act_getBrandUrl.asp?makerid='+v,
		cache: false,
		success: function(message) {
			if(message!="") {
			    
				window.location.href = "custom://brandproduct.custom?" + message;
			}
		}
	});
}
</script>
<script type="text/javascript" src="/lib/js/jquery-1.7.1.min.js"></script>

</head>
<body>
    <table width="100%"  cellspacing="0" cellpadding="4">
    <tr>
        <td colspan="2"><strong>더핑거스 아티스트 API Test</strong>
        </td>
    </tr>
    
    <tr>
        <td colspan="2">
            ==================================================================================
            
        </td>
    </tr>
	    <tr >
        <td colspan="2" >
    <form name="smFrm" method="post" action="">
    <input type="hidden" name="json" value="">
    </form>
    <iframe id="ifrmtarget" name="ifrmtarget" frameborder="1" style="top:0;left:0;z-index:-1;width:100%;height:100%;border:1"></iframe>
        </td>
    </tr>
    <form name="frmjson">
    <% for i=0 to C_DIMX-1 %>
    <tr>
        <td>
            <%=CHKIIF(protocolFix(i)=1,"<strong>","")%>
            <%=i+1%>. <%=protocolNM(i)%>  <font color="<%=CHKIIF(protocolFix(i)=1,"#000000","gray")%>">| <%=protocolComment(i)%> </font>
            <%=CHKIIF(protocolFix(i)=1,"</strong>","")%>
        <br>URL : <%=protocolURL(i)%>
        </td>
        <td width="50" align="center"> 
        <input type="button" value="REQ" onClick="jREQ('<%=i%>')"> 
        </td>
    </tr>
    <tr>
        <td colspan="2">
        <input type="hidden" name="jurl" value="<%=protocolURL(i)%>">
        <input type="text" name="json" value='<%="{" + replace(DFTProto&","&protocolJson(i),"'","""")+ "}"%>' size="180">
        </td>
        
    </tr>
    <tr>
    <td colspan="2"></td>
    </tr>
    <% next %>
    </form>

    </table>
</body>
</html>