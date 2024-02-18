<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 공용세금계산서 발행 위하고 api 연동
' History : 2022.10.31 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderutf8.asp"-->
<!-- #include virtual="/lib/util/htmllib_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/common/jungsan/wehagoApiFunction.asp" -->
<%
dim taxIdx  : taxIdx  = requestCheckVar(getNumeric(request("taxIdx")),10)

if taxIdx="" then
    response.write "세금계산서번호가 없습니다."
    session.codePage = 949
    dbget.close() : response.end
end if

dim oTax, repEmail
set oTax = new CTax
    oTax.FRecttaxIdx = taxIdx
    oTax.GetTaxRead

dim sell_hp, sell_hp1, sell_hp2, sell_hp3
dim buy_hp, buy_hp1, buy_hp2, buy_hp3

sell_hp = Split(oTax.FOneItem.FsupplyRepTel, "-")
buy_hp = Split(oTax.FOneItem.FrepTel, "-")

if (UBound(sell_hp) >= 0) then
	sell_hp1 = sell_hp(0)
end if

if (UBound(sell_hp) >= 1) then
	sell_hp2 = sell_hp(1)
end if

if (UBound(sell_hp) >= 2) then
	sell_hp3 = sell_hp(2)
end if

if (UBound(buy_hp) >= 0) then
	buy_hp1 = buy_hp(0)
end if

if (UBound(buy_hp) >= 1) then
	buy_hp2 = buy_hp(1)
end if

if (UBound(buy_hp) >= 2) then
	buy_hp3 = buy_hp(2)
end if
repEmail = db2html(oTax.FOneItem.FrepEmail)

IF application("Svr_Info")="Dev" THEN
    sell_hp1 = "010"
    sell_hp2 = "9177"
    sell_hp3 = "8708"
    buy_hp1 = "010"
    buy_hp2 = "9177"
    buy_hp3 = "8708"
    repEmail = "tozzinet@10x10.co.kr"
end if

if (oTax.FOneItem.Fbilldiv = "52") or (oTax.FOneItem.Fbilldiv = "55") then
	response.write "텐바이텐 이외 사업자 발행불가"
    session.codePage = 949
    dbget.close() : response.end
end if

dim reg_socno
dim reg_subsocno
dim reg_socname
dim reg_ceoname
dim reg_socaddr
dim reg_socstatus
dim reg_socevent
dim reg_managername
dim reg_managerphone
dim reg_managermail

dim tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID, tplbillUserPass
dim busiNo
reg_socno			= oTax.FOneItem.FsupplyBusiNo
reg_subsocno		= oTax.FOneItem.FsupplyBusiSubNo
reg_socname			= oTax.FOneItem.FsupplyBusiName
reg_ceoname			= oTax.FOneItem.FsupplyBusiCEOName
reg_socaddr			= oTax.FOneItem.FsupplyBusiAddr
reg_socstatus		= oTax.FOneItem.FsupplyBusiType
reg_socevent		= oTax.FOneItem.FsupplyBusiItem
reg_managername		= oTax.FOneItem.FsupplyRepName
reg_managerphone	= oTax.FOneItem.FsupplyRepTel
reg_managermail		= oTax.FOneItem.FsupplyRepEmail
busiNo = oTax.FOneItem.FbusiNo
IF application("Svr_Info")="Dev" THEN
    reg_socno = "2222222227"
    busiNo = "1111111119"
    reg_managerphone	= "01091778708"
    reg_managermail		= "tozzinet@10x10.co.kr"
end if

dim FG_VAT : FG_VAT = "1"			'// 1과세, 3면세, 2영세(잘못된것 아님 : 빌365)

if IsNull(oTax.FOneItem.Ftaxtype) then
	oTax.FOneItem.Ftaxtype = ""
end if

'// Y : 과세 / N : 면세 / 0 : 영세
Select Case oTax.FOneItem.Ftaxtype
	Case "Y"
		FG_VAT = "1"
	Case "N"
		FG_VAT = "3"
	Case "0"
		FG_VAT = "2"
	Case Else
		response.write "과세구분 설정 에러"
        session.codePage = 949
        dbget.close() : response.end
End Select

dim isueDate

if IsNull(oTax.FOneItem.FisueDate) then
	oTax.FOneItem.FisueDate = ""
end if

if (oTax.FOneItem.FisueDate = "") then
	response.write "발행일자 설정 에러"
    session.codePage = 949
    dbget.close() : response.end
else
	isueDate = oTax.FOneItem.FisueDate
end if

dim ipkumdate : ipkumdate = ""

if IsNull(oTax.FOneItem.Fipkumdate) then
	oTax.FOneItem.Fipkumdate = ""
end if

'// 고객 주문의 경우 입금일자
ipkumdate = oTax.FOneItem.Fipkumdate

dim consignYN

if IsNull(oTax.FOneItem.FconsignYN) then
	oTax.FOneItem.FconsignYN = ""
end if

if (oTax.FOneItem.FconsignYN = "") then
	response.write "위수탁구분 설정 에러"
    session.codePage = 949
    dbget.close() : response.end
else
	consignYN = oTax.FOneItem.FconsignYN
end if

if (oTax.FOneItem.Fbilldiv = "99") then
	Call Get3PLUpcheInfoByTPLCompanyid(oTax.FOneItem.Ftplcompanyid, tplcompanyname, tplgroupid, tplbillUserID, tplbillUserPass)
end if

dim wehagoRedirectUri
IF application("Svr_Info")="Dev" THEN
    wehagoRedirectUri = "http://localhost:11117/cscenter/taxsheet/popCommonWehagoTaxApiCallback.asp?taxIdx="& taxIdx &""
else
    wehagoRedirectUri = "https://webadmin.10x10.co.kr/cscenter/taxsheet/popCommonWehagoTaxApiCallback.asp?taxIdx="& taxIdx &""
end if
%>
<script type="text/javascript" src="https://static.wehago.com/support/jquery/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="https://static.wehago.com/support/jquery/jquery.cookie.js"></script>
<script type="text/javascript" src="https://static.wehago.com/support/wehago.0.1.2.js" charset="utf-8"></script>
<script type="text/javascript" src="https://static.wehago.com/support/wehagoLogin-1.1.6.min.js" charset="utf-8"></script>
<script type="text/javascript" src="https://static.wehago.com/support/service/common/wehagoCommon-0.2.8.min.js" charset="utf-8"></script>
<script type="text/javascript" src="https://static.wehago.com/support/service/invoice/wehagoInvoice-0.0.5.min.js" charset="utf-8"></script>
<script type="text/javascript" src="/js/gibberish-aesUTF8.js"></script>
<script type="text/javascript">
    <!-- #include virtual="/common/jungsan/wehago_globals_js.asp"-->
</script>

<div style="width: 300px; margin: 20px auto 0; padding: 10px; border: 1px solid #9297a4;">
    <h2>발행일 : <%= isueDate %></h2>
    <h3>오래 기다려도 자동발행이 안될경우 눌러주세요.</h3>

    <!-- 위하고 아이디로 로그인 버튼 노출 영역 -->
    <div id="wehago_id_login"></div>
    <!-- // 위하고 아이디로 로그인 버튼 노출 영역 -->
</div>
<script type="text/javascript">
    var AES256Key = "<%= AES256Key %>"
    <% IF application("Svr_Info")="Dev" THEN %>
        var ID_AES256 = AES_Encode(AES256Key,"tozzinet")
        var pw_AES256 = AES_Encode(AES256Key,"tenbyten!!")
        <%
        ' 빌36524용 테스트 계정
        'var ID_AES256 = AES_Encode(AES256Key,"BILLTEST02")
        'var pw_AES256 = AES_Encode(AES256Key,"bizon#720")
        %>
    <% else %>
        <%
        ' 빌36524용 계정. 빌에서는 이계정으로 발행 했으며, 위하고 들어오면서 tenbyten 으로 통합되서 발행되는듯함. 재무팀(최현희) 에서 tenbyten 아이디로 전부 발행하면 된다고 말해줌
        Select Case oTax.FOneItem.Fbilldiv
        '    '// 고객 - 공급자 텐바이텐
        '    Case "01"
        '        var ID_AES256 = AES_Encode(AES256Key,"customer")
        '        var pw_AES256 = AES_Encode(AES256Key,"20011010")
        '    '// 고객 - 공급자(업체별)
        '    Case "11"
        '        var ID_AES256 = AES_Encode(AES256Key,"customer")
        '        var pw_AES256 = AES_Encode(AES256Key,"20011010")
        '    '// 가맹점 - 공급자 텐바이텐
        '    Case "02"
        '        var ID_AES256 = AES_Encode(AES256Key,"accounts")
        '        var pw_AES256 = AES_Encode(AES256Key,"20011010")
        '    '// 프로모션 - 공급자 텐바이텐
        '    Case "03"
        '        var ID_AES256 = AES_Encode(AES256Key,"promotion")
        '        var pw_AES256 = AES_Encode(AES256Key,"20011010")
        '    '// 기타 - 공급자 텐바이텐
        '    Case "51"
        '        var ID_AES256 = AES_Encode(AES256Key,"accounts")
        '        var pw_AES256 = AES_Encode(AES256Key,"20011010")
            '// 3PL업체
            Case "99"
        %>
                var ID_AES256 = AES_Encode(AES256Key,"<%= tplbillUserID %>")
                var pw_AES256 = AES_Encode(AES256Key,"<%= tplbillUserPass %>")
        <%
            Case Else
        %>
                var ID_AES256 = AES_Encode(AES256Key,"tenbyten")
                var pw_AES256 = AES_Encode(AES256Key,"tenbyten10!")
        <%
        End Select
    end if
    %>
    var ID_AES256_URLEncode = encodeURIComponent(ID_AES256);
    var pw_AES256_URLEncode = encodeURIComponent(pw_AES256);

    var wehago_id_login = new wehago_id_login({
        app_key: "<%= wehagoAppKey %>",  // AppKey
        service_code: "<%= wehagoServiceCode %>",  // ServiceCode
        redirect_uri: "<%= wehagoRedirectUri %>",  // Callback URL
        <% IF application("Svr_Info")="Dev" THEN %>
            //mode: "dev",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
            mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
        <% else %>
            mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
        <% end if %>
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
        setTimeout(function(){
            if (confirm('발행일 : <%= isueDate %>\n발행 하시겠습니까?')){
                $("#wehago_id_login_anchor").trigger("click");   		
            }
        },1000);
    });

</script>
<%
function Get3PLUpcheInfoByTPLCompanyid(tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID, byRef tplbillUserPass)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID, billUserPass as tplbillUserPass "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner_tpl t with (nolock)"
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.tplcompanyid = '" + CStr(tplcompanyid) + "' "
	'' response.write sqlStr
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
		tplcompanyname = db2html(rsget("tplcompanyname"))
		tplgroupid = rsget("tplgroupid")
		tplbillUserID = rsget("tplbillUserID")
		tplbillUserPass = rsget("tplbillUserPass")
	end if
	rsget.close
end function

session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
