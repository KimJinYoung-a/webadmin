<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 배송추적 예외처리 브랜드-택배사 목록 기타/퀵 익일로 처리
' Hieditor : 2019.06.27 eastone 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim i
Dim songjangdiv : songjangdiv	  = requestCheckVar(request("songjangdiv"),10)
Dim makerid     : makerid         = requestCheckVar(request("makerid"),32)
Dim page        : page            = requestCheckVar(request("page"),10)


if (page="") then page=1

dim oDeliveryTrackExcept
SET oDeliveryTrackExcept = New CDeliveryTrack
oDeliveryTrackExcept.FCurrPage = page
oDeliveryTrackExcept.FPageSize = 50
oDeliveryTrackExcept.FRectsongjangDiv = songjangdiv
oDeliveryTrackExcept.FRectMakerid     = makerid

oDeliveryTrackExcept.getDeliveryTrackExceptFinBrandList()


%>
<script language="javascript">
function jsSubmit(frm) {
	frm.submit();
}

function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function addExceptBrand(comp){
    var frm = comp.form;
    if (frm.exceptmakerid.value.length<1){
        alert("브랜드ID를 입력해주세요.");
        frm.exceptmakerid.focus();
        return;
    }

    if (frm.exceptsongjangdiv.value.length<1){
        alert("택배사를 선택해 주세요.");
        frm.addexceptdlv.focus();
        return;
    }

    if (confirm("추가 하시겠습니까?")){
        frm.mode.value="addexceptbrand";
        frm.submit();
    }
    
}

function delThis(comp,imakerid,isongjangdiv){
    if (confirm('삭제하시겠습니까?')){
        var iurl = "DeliveryTrackingSummary_Process.asp?exceptmakerid="+imakerid+"&mode=delexceptbrand&exceptsongjangdiv="+isongjangdiv;
        var popwin=window.open(iurl,'dlExceptBrand','width=200 height=200 scrollbars=yes resizable=yes');
        popwin.focus();
    }
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td  width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        &nbsp; 택배사 : <% Call drawTrackDeliverBox("songjangdiv",songjangdiv, "") %>
        브랜드ID : <input type="text" class="text" name="makerid" value="<%= makerid %>" size="16" > 
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(document.frm);">
	</td>
</tr>
</table>
</form>

<p />
<form name="frmexcept" method="post" action="DeliveryTrackingSummary_Process.asp">
<input type="hidden" name="mode" value="addexceptbrand">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
        [기타택배사 자동처리 브랜드 목록]<br>
        매일 새벽 전전일 내역을 배치로 배송완료 처리 합니다.(출고일+1일)
	</td>
    <td colspan="2" align="right">
    브랜드ID : 
    <input type="text" name="exceptmakerid" value="" size="16" maxlength="32">
    <% Call drawTrackDeliverBox("exceptsongjangdiv","99", "") %>
    
    <input type="button" value="추가" onClick="addExceptBrand(this);">
    </td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="130">브랜드ID</td>
    <td width="120">택배사</td>
    <td width="120">등록일</td>
    <td width="120">등록자</td>
    <td >비고</td>

</tr>
<% for i = 0 to (oDeliveryTrackExcept.FResultCount - 1) %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%=oDeliveryTrackExcept.FItemList(i).Fmakerid %></td>
    <td><%=oDeliveryTrackExcept.FItemList(i).Fdivname %></td>
    <td><%=oDeliveryTrackExcept.FItemList(i).Fregdt %></td>
    <td><%=oDeliveryTrackExcept.FItemList(i).Freguserid %></td>
    <td align="center">
    <input type="button" value="삭제" onClick="delThis(this,'<%=oDeliveryTrackExcept.FItemList(i).Fmakerid %>','<%=oDeliveryTrackExcept.FItemList(i).Fsongjangdiv %>');">
    </td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="5" align="center">
        <% if oDeliveryTrackExcept.HasPreScroll then %>
        <a href="javascript:goPage('<%= oDeliveryTrackExcept.StartScrollPage-1 %>');">[pre]</a>
        <% else %>
            [pre]
        <% end if %>

        <% for i=0 + oDeliveryTrackExcept.StartScrollPage to oDeliveryTrackExcept.FScrollCount + oDeliveryTrackExcept.StartScrollPage - 1 %>
            <% if i>oDeliveryTrackExcept.FTotalpage then Exit for %>
            <% if CStr(page)=CStr(i) then %>
            <font color="red">[<%= i %>]</font>
            <% else %>
            <a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
            <% end if %>
        <% next %>

        <% if oDeliveryTrackExcept.HasNextScroll then %>
            <a href="javascript:goPage('<%= i %>');">[next]</a>
        <% else %>
            [next]
        <% end if %>
    </td>
</tr>
</table>
</form>

<p />


<%
SET oDeliveryTrackExcept = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
