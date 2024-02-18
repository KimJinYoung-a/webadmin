<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 개인정보 문서 파기 관리
' History : 2019.08.13 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/isms/personaldata_cls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim i, idxarr, userid, menupos
    menupos = requestcheckvar(request("menupos"),10)

userid = session("ssBctId")

if request.form("idx").count<1 then
    response.write "<script type='text/javascript'>"
    response.write "    alert('저장할값이 없습니다.');"
    response.write "    self.close();"
    response.write "</script>"
    dbget.close() : response.end
end if

for i=1 to request.form("idx").count
    idxarr = idxarr & request.form("idx")(i) & ","
next

%>
<script type='text/javascript'>

function downFileconfirm(){
    if ( !frmupdate.ck1.checked ){
        alert('문서분쇄에 동의 체크를 해주세요.');
        frmupdate.ck1.focus();
        return;
    }
    if ( !frmupdate.ck2.checked ){
        alert('CD,USB폐기에 동의 체크를 해주세요.');
        frmupdate.ck2.focus();
        return;
    }
    if ( !frmupdate.ck3.checked ){
        alert('컴퓨터 파일 삭제에 동의 체크를 해주세요.');
        frmupdate.ck3.focus();
        return;
    }
    if ( !frmupdate.ck4.checked ){
        alert('기타저장(웹하드,이메일 등) 매체 파기/삭제에 동의 체크를 해주세요.');
        frmupdate.ck4.focus();
        return;
    }

    frmupdate.mode.value = "downFileconfirmArr";
    frmupdate.target="_self"
    frmupdate.action="/admin/isms/personaldata_process.asp";
    frmupdate.submit();
}

// 전체선택
function totalCheck(tmpval){
    if (tmpval.checked){
        frmupdate.ck1.checked = true
        frmupdate.ck2.checked = true
        frmupdate.ck3.checked = true
        frmupdate.ck4.checked = true
    }else{
        frmupdate.ck1.checked = false
        frmupdate.ck2.checked = false
        frmupdate.ck3.checked = false
        frmupdate.ck4.checked = false
    }
}

</script>
</head>
<body>
<form name="frmupdate" method="post" action="/admin/isms/personaldata_process.asp" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idxarr" value="<%= idxarr %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td><h2>고객정보 파기 확인서</h2></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>
        1. <%= session("ssBctCname") %>가 텐바이텐을 통하여 상품판매, 배송, 반품 등의 제반 업무를 진행하는데 있어 취득한 제반 정보의 파기와 관련하여
        <br>&nbsp;&nbsp;&nbsp;다음의 내용을 확인하여 주시기 바랍니다.
        <Br>2. <%= session("ssBctCname") %>가 2개월 이상 개인정보 파기와 확인서를 미작성시, 어드민 사용에 제약이 있음을 알려 드립니다.
        <br>3. 고객 정보보호 및 공정한 거래질서 확립을 위하여 노력해 주시기 바랍니다.
    </td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>
        <br><%= session("ssBctCname") %>가 텐바이텐을 통하여 판매한 상품과 관련한 계약 이행을 위한 제반 고객정보 등에 대하여 다음과 같이 파기 업무를 이행하였음을 확인 합니다.
    </td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>
        1. 파기 정보 : 계약 이행을 위한 제반 고객 정보(성명, 전화번호, 주소 등)
        <br>2. 파기 방법
        <br>&nbsp;&nbsp;&nbsp;
        &nbsp;&nbsp;<input type="checkbox" name="ck1"> 문서분쇄
        &nbsp;&nbsp;<input type="checkbox" name="ck2"> CD,USB폐기
        &nbsp;&nbsp;<input type="checkbox" name="ck3"> 컴퓨터 파일 삭제
        &nbsp;&nbsp;<input type="checkbox" name="ck4"> 기타저장(웹하드,이메일 등) 매체 파기/삭제
    </td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>
        <br><%= session("ssBctCname") %>은 개인정보 관련 법률을 준수할 것이며, 상기 고객정보 파기 업무의 불이행 또는 불완전 이행으로 발행하는 행정적,
        민형사상 책임은 이 부담할 것을 확약 합니다.
        <br><br><p align="center"><%= year(date()) %>년 <%= month(date()) %>월 <%= day(date()) %>일 <%= session("ssBctCname") %></p>
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td>
        <input type="checkbox" name="ckall" onClick="totalCheck(this);">모두확인
        &nbsp;
        <input type="button" value="확인"  onclick="downFileconfirm();" class="button" />
    </td>
</tr>
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->