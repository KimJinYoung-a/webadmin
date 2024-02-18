<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<%
'###########################################################
' Description : 운송장전송주소오류관리
' Hieditor : 2022.06.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/logistics/songjang/SongJangSendClass.asp"-->

<%
Dim i, page, osongjang, reload, SongJangGubun, siteseq, songjangdiv, gubuncd
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
    reload = requestcheckvar(request("reload"),2)
    SongJangGubun = requestcheckvar(trim(request("SongJangGubun")),10)
    siteseq = requestcheckvar(getNumeric(request("siteseq")),10)
    songjangdiv = requestcheckvar(trim(request("songjangdiv")),32)
    gubuncd = requestcheckvar(trim(request("gubuncd")),3)

if page = "" then page = 1
if SongJangGubun="" then
    SongJangGubun="GENERAL"
end if

set osongjang = new cSongJangSendError
	osongjang.FPageSize = 50
	osongjang.FCurrPage = 1
    osongjang.FrectSongJangGubun = SongJangGubun
    osongjang.Frectsiteseq = siteseq
    osongjang.FRectSongjangDiv = songjangdiv
    osongjang.Frectgubuncd = gubuncd
    osongjang.GetSongJangSendErrorList()
%>

<script type="text/javascript">

function SongJangSendErrorEdit(idx, SongJangGubun){
	var popwin = window.open('/admin/logics/songjang/SongJangSendErrorEdit.asp?idx='+idx+'&SongJangGubun='+SongJangGubun+'&menupos=<%=menupos%>','SongJangSendErroraddreg','width=1200,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function NextPage(page){
	document.frm.page.value= page;
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method=get action="" style="margin:0px;">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reload" value="on">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
    <td align="left">
        * 운송장구분 : <% drawSelectBoxSongJangGubun "SongJangGubun",SongJangGubun,"" %>
        &nbsp;
        * 업체 : <% drawSelectBoxSiteSeq "siteseq",siteseq,"" %>
        &nbsp;
        * 출고구분 : <% drawSelectBoxgubuncd "gubuncd",gubuncd,"" %>
        &nbsp;
        * 택배사 :
        <% Call drawSelectBoxDeliverCompany ("songjangdiv", songjangdiv) %>
    </td>
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="검색" onClick="NextPage('');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
        운송장 구분은 일반송장과 반품송장 두가지 전송이 있으며
        <br>고객 배송 주소 오입력으로 인한 택배사 운송장 전송 오류건 입니다.
        <br>당일 오류건은 반드시 당일 택배사 배송차가 출발하기 이전에 수정 되어야 합니다.
        <br>당일 미수정시 허브나 터미널 택배사 전산에서 조회되지 않습니다.
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
    <td colspan="15">
        검색결과 : <b><%= osongjang.FTotalCount %></b>
        &nbsp;
        페이지 : <b><%= page %>/ <%= osongjang.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>번호</td>
    <td>업체</td>
    <td>출고구분</td>
    <td>택배사</td>
    <td>송장번호</td>
    <td>주문번호</td>
    <td>이름</td>
    <td>전화번호</td>
    <td>휴대폰번호</td>
    <td>출고일</td>
    <td>우편번호</td>
    <td>주소</td>
    <td>비고</td>
</tr>
<% if osongjang.FresultCount>0 then %>
<% for i=0 to osongjang.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" align="center">
    <td>
        <%= osongjang.FItemList(i).fidx %>
    </td>
    <td>
        <%= getSiteSeqnamestr(osongjang.FItemList(i).fsiteseq) %>
    </td>
    <td>
        <%= getgubuncdname(osongjang.FItemList(i).fgubuncd) %>
    </td>
    <td>
        <%= osongjang.FItemList(i).Fdivname %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fSONGJANGNO %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fORDERSERIAL %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fnm %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fTEL_NO %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fHP_NO %>
    </td>
    <td>
        <%= left(osongjang.FItemList(i).fregdate,10) %>
        <br><%= mid(osongjang.FItemList(i).fregdate,12,22) %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fZIP_NO %>
    </td>
    <td><%= osongjang.FItemList(i).fADDR %>&nbsp;<%= osongjang.FItemList(i).fADDR_ETC %></td>
    <td>
        <input type="button" class="button" value="수정" onclick="SongJangSendErrorEdit('<%= osongjang.FItemList(i).fidx %>','<%= SongJangGubun %>');">
    </td>
</tr>
<% next %>

<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>

</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<%
session.codePage = 949
 %>