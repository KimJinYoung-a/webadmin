<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<!-- #include virtual="/lib/util/htmllib_UTF8.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<%
dim wehagoRedirectUri
IF application("Svr_Info")="Dev" THEN
    wehagoRedirectUri = "http://localhost:11117/cscenter/taxsheet/popCommonsampleWehagoTaxApiCallback.asp"
else
    wehagoRedirectUri = "https://webadmin.10x10.co.kr/cscenter/taxsheet/popCommonsampleWehagoTaxApiCallback.asp"
end if

' 위하고측에서 쿠키유지를 8시간까지 하고 있다고 함
' 처음 위하고 접속시 토큰값을 저장하고 그 토큰으로 8시간 통신을 한다. 토큰이 없을때만 로그인해서 토큰 받아옴.
dim WEHAGO_access_token, WEHAGO_state, WEHAGO_wehago_id, WEHAGO_time, isWehagoLogin
WEHAGO_access_token=session("WEHAGO_access_token")
WEHAGO_state=session("WEHAGO_state")
WEHAGO_wehago_id=session("WEHAGO_wehago_id")
WEHAGO_time=session("WEHAGO_time")
isWehagoLogin=false

if WEHAGO_time<>"" then
    isWehagoLogin=true

    if datediff("h",session("WEHAGO_time"),now())>=7 then
        isWehagoLogin=false
        WEHAGO_access_token=""
        WEHAGO_state=""
        WEHAGO_wehago_id=""
        WEHAGO_time=""
        session("WEHAGO_access_token") = ""
        session("WEHAGO_state") = ""
        session("WEHAGO_wehago_id") = ""
        session("WEHAGO_time") = ""
        Call fn_RDS_SSN_SET()
    end if
end if

%>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>위하고 아이디로 로그인</title>
    <script type="text/javascript" src="https://static.wehago.com/support/jquery/jquery-3.2.1.min.js"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/jquery/jquery.cookie.js"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/wehago.0.1.2.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/wehagoLogin-1.1.6.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/service/common/wehagoCommon-0.2.8.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/service/invoice/wehagoInvoice-0.0.5.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="/cscenter/taxsheet/wehago_globals.js" charset="utf-8"></script>
    <script type="text/javascript" src="/js/gibberish-aesUTF8.js"></script>

<script type="text/javascript">

</script>
</head>
<body>

<div style="width: 600px; margin: 200px auto 0; padding: 50px; border: 1px solid #9297a4;">
    <h1>위하고 자동 로그인</h1>

    <!-- 위하고 아이디로 로그인 버튼 노출 영역 -->
    <div id="wehago_id_login"></div>
    <!-- // 위하고 아이디로 로그인 버튼 노출 영역 -->
    <script type="text/javascript">

        function AES_Encode(key,plain_text){
            GibberishAES.size(256);	
            return GibberishAES.aesEncrypt(plain_text, key);
        }
        function AES_Decode(key,base64_text){
            GibberishAES.size(256);	
            return GibberishAES.aesDecrypt(base64_text, key);
        }
        var AES256Key = "E86916E2CF3846B9BB6880CBC0447C35"
        var ID_AES256 = AES_Encode(AES256Key,"tozzinet")
        var pw_AES256 = AES_Encode(AES256Key,"tenbyten!!")
        var ID_AES256_URLEncode = encodeURIComponent(ID_AES256);
        var pw_AES256_URLEncode = encodeURIComponent(pw_AES256);

        var wehago_id_login = new wehago_id_login({
            app_key: "98730f8cfdef4f77af17ce8ee08282fb",  // AppKey
            service_code: "10x10",  // ServiceCode
            redirect_uri: "http://localhost:11117/cscenter/taxsheet/popCommonsampleWehagoTaxApiCallback.asp",  // Callback URL
            mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
            is_auto_login:"T", // T-자동로그인 설정, F-기존 로그인(기존 로그인은 is_auto_login,id,pw 생략 후 사용 가능합니다.)
            id:ID_AES256_URLEncode,
            pw:pw_AES256_URLEncode
        });

        var state = wehago_id_login.getUniqState();
        wehago_id_login.setButton("white", 1, 40);
        wehago_id_login.setDomain(".10x10.co.kr");
        wehago_id_login.setState(state);
        wehago_id_login.setPopup();  // 위하고 로그인페이지를 팝업으로 띄울경우
        wehago_id_login.init_wehago_id_login();

        $(function() {
            <%
            ' 위하고 토큰이 있는경우
            if isWehagoLogin then
            %>
                location.replace("<%= wehagoRedirectUri %>?access_token=<%=WEHAGO_access_token%>&wehago_id=<%=WEHAGO_wehago_id%>&state=<%=WEHAGO_state%>");
            <%
            ' 위하고 토큰이 없는경우 자동로그인을 시켜서 토큰을 받아옴.
            else
            %>
                // 이렇게하면 자동로그인 됨.
                //setTimeout(function(){
                //    $("#wehago_id_login_anchor").trigger("click");   		
                //},1000);
            <% end if %>
        });

    </script>
</div>
</body>
</html>
