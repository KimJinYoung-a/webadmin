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
' Description : 수수료 세금계산서 발행 위하고 api 연동
' History : 2022.11.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyheadutf8.asp"-->
<!-- #include virtual="/lib/util/htmllib_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls_utf8.asp"-->
<!-- #include virtual="/common/jungsan/wehagoApiFunction.asp" -->
<%
dim i, repEmail, jungsan_name, isueDate, autotype
dim makerid, yyyy1,mm1, onoffGubun, jidx, isauto, nextjidx, FG_VAT
makerid 		= requestCheckvar(request("makerid"),32)
yyyy1   		= requestCheckvar(request("yyyy1"),10)
mm1     		= requestCheckvar(request("mm1"),10)
onoffGubun     	= requestCheckvar(request("onoffGubun"),10)
jidx            = requestCheckvar(request("jidx"),10)
isauto          = requestCheckvar(request("isauto"),10)
nextjidx        = requestCheckvar(request("nextjidx"),10)
autotype 		= requestCheckvar(request("autotype"),32)
dim groupid
groupid = getPartnerId2GroupID(makerid)

dim ojungsanTaxCC
set ojungsanTaxCC = new CUpcheJungsanTax
ojungsanTaxCC.FRectMakerid = makerid
ojungsanTaxCC.FRectTargetGbn = onoffGubun
ojungsanTaxCC.FRectJjungsanIdx = jidx
ojungsanTaxCC.getOneUpcheJungsanTax

dim PrdCommissionSum : PrdCommissionSum = 0

if (ojungsanTaxCC.FresultCount>0) then
	if (ojungsanTaxCC.FOneItem.IsCommissionTax) then
	    PrdCommissionSum = ojungsanTaxCC.FOneItem.Ftotalcommission
	end if

    FG_VAT = ojungsanTaxCC.FOneItem.getBill_FG_VAT
end if

IF application("Svr_Info")<>"Dev" THEN
    if PrdCommissionSum = 0 then
        if (request("autotype")="V2") then
        response.write "<script type='text/javascript'>"&vbCRLF
        response.write "opener.addResultLog('"&request("jidx")&"','수수료0');"&vbCRLF
        response.write "opener.fnNextEvalProc();"&vbCRLF
        response.write "</script>"
        else
        response.write "<script type='text/javascript'>alert('수수료 매출정보가 없습니다.');</script>"
        response.write "수수료 매출정보가 없습니다"
        end if
        session.codePage = 949
        dbget.close() : response.end
    end if
end if
if ojungsanTaxCC.FOneItem.IsEvaledTax then
    if (request("autotype")="V2") then
    response.write "<script type='text/javascript'>"&vbCRLF
    response.write "opener.addResultLog('"&request("jidx")&"','기정산확정');"&vbCRLF
    response.write "opener.fnNextEvalProc();"&vbCRLF
    response.write "</script>"
    else
    response.write "<script type='text/javascript'>alert('이미 정산 확정된 내역입니다.');</script>"
    response.write "이미 정산 확정된 내역입니다."
    end if
    session.codePage = 949
    dbget.close()	:	response.End
end if

dim opartner, ogroup
dim stypename

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid
opartner.FPageSize = 1
opartner.GetOnePartnerNUser

set ogroup = new CPartnerGroup
ogroup.FRectGroupid = ojungsanTaxCC.FOneItem.Fgroupid
ogroup.GetOneGroupInfo

if ogroup.FResultCount<1 then
    if (request("autotype")="V2") then
        response.write "<script type='text/javascript'>"&vbCRLF
        response.write "opener.addResultLog('"&request("jidx")&"','그룹미지정/정산정보없음');"&vbCRLF
        response.write "opener.fnNextEvalProc();"&vbCRLF
        response.write "</script>"
    else
        response.write "<script type='text/javascript'>alert('그룹 코드가 지정되지 않았거나, 정산정보가 없습니다.');</script>"
        response.write "그룹 코드가 지정되지 않았거나, 정산정보가 없습니다"
    end if
    session.codePage = 949
	dbget.close()	:	response.End
end if

dim MaySocialNo : MaySocialNo=FALSE ''주민번호로 발급
if IsMaySocialNo(ogroup.FOneItem.Fcompany_no) then
    MaySocialNo = true
    ogroup.FOneItem.Fcompany_no = ogroup.FOneItem.FdecCompNo
end if

jungsan_name=ogroup.FOneItem.Fjungsan_name

if (NOT MaySocialNo) then
    if LEN(replace(ogroup.FOneItem.Fcompany_no,"-",""))<>10 then
        if (request("autotype")="V2") then
            response.write "<script type='text/javascript'>"&vbCRLF
            response.write "opener.addResultLog('"&request("jidx")&"','사업자번호');"&vbCRLF
            response.write "opener.fnNextEvalProc();"&vbCRLF
            response.write "</script>"
        else
            response.write "<script type='text/javascript'>alert('사업자 번호가 올바르지 않습니다.');</script>"
            response.write "사업자 번호가 올바르지 않습니다."& replace(ogroup.FOneItem.Fcompany_no,"-","") & "::" & LEN(replace(ogroup.FOneItem.Fcompany_no,"-",""))
        end if
        session.codePage = 949
    	dbget.close()	:	response.End
    end if
end if

stypename = "세금계산서"
dim jungsan_hpall, jungsan_hp1,jungsan_hp2,jungsan_hp3, reg_socno, busiNo, buyceoname, buycompany_address1, buycompany_address2
jungsan_hpall = Trim(ogroup.FOneItem.Fjungsan_hp)
jungsan_hpall = split(jungsan_hpall,"-")

if UBound(jungsan_hpall)>=0 then
	jungsan_hp1 = jungsan_hpall(0)
end if
if UBound(jungsan_hpall)>=1 then
	jungsan_hp2 = jungsan_hpall(1)
end if
if UBound(jungsan_hpall)>=2 then
	jungsan_hp3 = jungsan_hpall(2)
end if

if (jungsan_hp2="") and (jungsan_hp3="") and (Len(jungsan_hp1)=11) then
    jungsan_hp3 = MID(jungsan_hp1,8,4)
    jungsan_hp2 = MID(jungsan_hp1,4,4)
    jungsan_hp1 = LEFT(jungsan_hp1,3)
end if
repEmail = db2html(ogroup.FOneItem.Fjungsan_email)
reg_socno = "211-87-00620"
busiNo = ogroup.FOneItem.Fcompany_no
buyceoname=ogroup.FOneItem.Fceoname
buycompany_address1 = ogroup.FOneItem.Fcompany_address
buycompany_address2 = ogroup.FOneItem.Fcompany_address2

IF application("Svr_Info")="Dev" THEN
    reg_socno = "2222222227"
    busiNo = "1111111119"
    buyceoname = "한용민"
    buycompany_address1 = "서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층"
    buycompany_address2 = "서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 15층"
    jungsan_hp1 = "010"
    jungsan_hp2 = "9177"
    jungsan_hp3 = "8708"
    repEmail = "tozzinet@10x10.co.kr"
    jungsan_name = "한용민"

    isueDate = date()
end if

Dim EVAL_CompanyNo  : EVAL_CompanyNo = "2118700620"

IF application("Svr_Info")<>"Dev" THEN
    if (replace(ogroup.FOneItem.Fcompany_no,"-","")=EVAL_CompanyNo) then
        if (request("autotype")="V2") then
            response.write "<script type='text/javascript'>"&vbCRLF
            response.write "opener.addResultLog('"&request("jidx")&"','텐바이텐사업자발행불가');"&vbCRLF
            response.write "opener.fnNextEvalProc();"&vbCRLF
            response.write "</script>"
        else
            response.write "<script type='text/javascript'>alert('텐바이텐 사업자 발행 불가.');</script>"
            response.write "텐바이텐 사업자 발행 불가."
        end if
        session.codePage = 949
        dbget.close()	:	response.End
    end if
end if

dim wehagoRedirectUri
IF application("Svr_Info")="Dev" THEN
    wehagoRedirectUri = "http://localhost:11117/admin/upchejungsan/popUWehagoTaxApiCallback.asp?makerid="& makerid &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&onoffGubun="& onoffGubun &"&jidx="& jidx &"&isauto="& isauto &"&autotype="& autotype &"&nextjidx="& nextjidx &""
else
    wehagoRedirectUri = "https://webadmin.10x10.co.kr/admin/upchejungsan/popUWehagoTaxApiCallback.asp?makerid="& makerid &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&onoffGubun="& onoffGubun &"&jidx="& jidx &"&isauto="& isauto &"&autotype="& autotype &"&nextjidx="& nextjidx &""
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
        var ID_AES256 = AES_Encode(AES256Key,"tenbyten")
        var pw_AES256 = AES_Encode(AES256Key,"tenbyten10!")
    <% end if %>
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
        <%
        ' 위하고 토큰이 있는경우
        if isWehagoLogin then
        %>
            location.replace("<%= wehagoRedirectUri %>&access_token=<%=WEHAGO_access_token%>&wehago_id=<%=WEHAGO_wehago_id%>&state=<%=WEHAGO_state%>");
        <%
        ' 위하고 토큰이 없는경우 자동로그인을 시켜서 토큰을 받아옴.
        else
        %>
            // 자동로그인
            setTimeout(function(){
                $("#wehago_id_login_anchor").trigger("click");
            },1000);
        <% end if %>
    });

</script>
<%
function IsMaySocialNo(icompanyno)
    IsMaySocialNo = false
    if isNULL(icompanyno) then Exit function
    IsMaySocialNo = LEN(trim(replace(icompanyno,"-","")))=13
end function

set ojungsanTaxCC = Nothing
set opartner = Nothing
set ogroup = Nothing

session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
