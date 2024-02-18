<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 송장대역관리
' Hieditor : 2021.04.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/logistics/invoice_band_cls.asp"-->
<%
Dim i, page, osongjang, isusing, reload, siteseq, gubuncd, osongjanglog, songjangdiv
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
    isusing = requestcheckvar(request("isusing"),10)
    siteseq = requestcheckvar(getNumeric(request("siteseq")),10)
    reload = requestcheckvar(request("reload"),2)
    gubuncd = requestcheckvar(trim(request("gubuncd")),3)
    songjangdiv = requestcheckvar(trim(request("songjangdiv")),32)

if page = "" then page = 1
if reload="" and isusing="" then isusing="Y"
if siteseq="" and siteseq="" then siteseq="10"
''if gubuncd="" and gubuncd="" then gubuncd="00"
if reload="" and songjangdiv="" then
    'songjangdiv = "1"   ' 한진택배
    songjangdiv = "2"   ' 롯데택배
end if

set osongjang = new cinvoice_band_list
	osongjang.FPageSize = 50
	osongjang.FCurrPage = page
    osongjang.Frectisusing = isusing
    osongjang.Frectsiteseq = siteseq
    osongjang.Frectgubuncd = gubuncd
    osongjang.FRectSongjangDiv = songjangdiv

    osongjang.finvoice_band()

set osongjanglog = new cinvoice_band_list
	osongjanglog.FPageSize = 5
	osongjanglog.FCurrPage = 1
    osongjanglog.Frectsiteseq = siteseq
    osongjanglog.Frectgubuncd = gubuncd
    osongjanglog.FRectSongjangDiv = songjangdiv

	osongjanglog.finvoice_band_log()
%>

<script type="text/javascript">

function invoice_band_reg(iidx){
	var popwin = window.open('/admin/logics/invoice_band_reg.asp?iidx='+iidx+'&menupos=<%=menupos%>','addreg','width=1200,height=700,scrollbars=yes,resizable=yes');
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
        * 업체 : <% drawSelectBoxSiteSeq "siteseq",siteseq,"" %>
        &nbsp;
        * 택배사 :
        <% Call drawSelectBoxDeliverCompany ("songjangdiv", songjangdiv) %>
        &nbsp;
        * 출고구분 : <% drawSelectBoxgubuncd "gubuncd",gubuncd,"" %>
        &nbsp;
        * 사용여부 : <% drawSelectBoxisusingYN "isusing",isusing,"" %>
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
<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left"></td>
    <td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
    <td colspan="3">
        ※ 최근 마지막 출고된 송장내역 로그 5개
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>송장번호(검증키포함)</td>
    <td>실제송장번호</td>
    <td>주문번호</td>
</tr>
<% if osongjanglog.FresultCount>0 then %>
<% for i=0 to osongjanglog.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" align="center">
    <td>
        <%= osongjanglog.FItemList(i).fSONGJANGNO %>
    </td>
    <td>
        <%= osongjanglog.FItemList(i).fREALSONGJANGNO %>
    </td>
    <td>
        <%= osongjanglog.FItemList(i).fORDERSERIAL %>
    </td>
</tr>
<% next %>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
</table>
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
    </td>
    <td align="right">
        <input type="button" class="button" value="신규등록" onclick="invoice_band_reg('');">
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
    <td>택배사</td>
    <td>출고구분</td>
    <td>송장번호(검증키포함)</td>
    <td>실제송장번호</td>
    <td>
        남은송장수
        <br>(8시간주기업데이트)
    </td>
    <td>
        기본송장여부
        <br>(현재로직스실제사용대역)
    </td>
    <td>사용여부</td>
    <td>최초등록</td>
    <td>최종수정</td>
    <td>비고</td>
</tr>
<% if osongjang.FresultCount>0 then %>
<% for i=0 to osongjang.FresultCount-1 %>
<% if osongjang.FItemList(i).fbasicsongjangyn = "Y" then %>
<tr align="center" bgcolor="#FFFFaa" align="center">
<% else %>
<tr align="center" bgcolor="#FFFFFF" align="center">
<% end if %>
    <td>
        <%= osongjang.FItemList(i).fiidx %>
    </td>
    <td>
        <%= getSiteSeqnamestr(osongjang.FItemList(i).fsiteseq) %>
    </td>
    <td>
        <%= osongjang.FItemList(i).Fdivname %>
    </td>
    <td>
        <%= getgubuncdname(osongjang.FItemList(i).fgubuncd) %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fstartsongjangno %> - <%= osongjang.FItemList(i).fendsongjangno %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fstartrealsongjangno %> - <%= osongjang.FItemList(i).fendrealsongjangno %>
    </td>
    <td>
        <% if osongjang.FItemList(i).fbasicsongjangyn="Y" then %>
            <% if osongjang.FItemList(i).fremainsongjangcount="0" then %>
                <%= osongjang.FItemList(i).fendrealsongjangno-osongjang.FItemList(i).fstartrealsongjangno %>
            <% else %>
                <%= osongjang.FItemList(i).fremainsongjangcount %>
            <% end if %>
        <% else %>
            <% 'if osongjang.fcurrentbasicsongjangidx > osongjang.FItemList(i).fiidx then %>
                <%= osongjang.FItemList(i).fremainsongjangcount %>
            <% 'else %>
                <%'= osongjang.FItemList(i).fendrealsongjangno-osongjang.FItemList(i).fstartrealsongjangno %>
            <% 'end if %>
        <% end if %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fbasicsongjangyn %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fisusing %>
    </td>
    <td>
        <%= osongjang.FItemList(i).freguserid %>
        <br><%= left(osongjang.FItemList(i).fregdate,10) %>
        <br><%= mid(osongjang.FItemList(i).fregdate,12,22) %>
    </td>
    <td>
        <%= osongjang.FItemList(i).flastuserid %>
        <br><%= left(osongjang.FItemList(i).flastupdate,10) %>
        <br><%= mid(osongjang.FItemList(i).flastupdate,12,22) %>
    </td>
    <td>
        <input type="button" class="button" value="수정" onclick="invoice_band_reg('<%= osongjang.FItemList(i).fiidx %>');">
    </td>
</tr>
<% next %>
<tr bgcolor="FFFFFF">
    <td colspan="15" align="center">
        <% if osongjang.HasPreScroll then %>
        <a href="javascript:NextPage('<%= osongjang.StartScrollPage-1 %>')">[pre]</a>
        <% else %>
            [pre]
        <% end if %>

        <% for i=0 + osongjang.StartScrollPage to osongjang.FScrollCount + osongjang.StartScrollPage - 1 %>
            <% if i>osongjang.FTotalpage then Exit for %>
            <% if CStr(page)=CStr(i) then %>
            <font color="red">[<%= i %>]</font>
            <% else %>
            <a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
            <% end if %>
        <% next %>

        <% if osongjang.HasNextScroll then %>
            <a href="javascript:NextPage('<%= i %>')">[next]</a>
        <% else %>
            [next]
        <% end if %>
    </td>
</tr>

<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>

</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
