<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : 상품일괄변경[관리자]
' History : 2021.11.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
if not(C_ADMIN_AUTH or C_MD_AUTH or C_SYSTEM_Part) then
    response.write "<script type='text/javascript'>alert('권한이 없습니다. MD팀,개발팀 파트장 이상 접근 가능 합니다.');</script>"
    dbget.close() : response.end
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function chmakermwdiv(){
    if ( $('#makerid').val() == ''){
        alert('변경하실 브랜드ID를 선택해 주세요.');
        return;
    }
    $('#frmmakeritemch').attr('target', 'view');
    $('#frmmakeritemch').attr('action', '/admin/itemmaster/item_change_all_process.asp');
    $('#mode').val('makerchmwdiv');
    $('#frmmakeritemch').submit();
    $('#mode').val('');
}

function chmakermargin(){
    if ( $('#makerid').val() == ''){
        alert('변경하실 브랜드ID를 선택해 주세요.');
        return;
    }
    if ( $('#margin').val() == '' || $('#margin').val() == 0){
        alert('변경하실 마진 % 을 입력해 주세요.');
        return;
    }
    $('#frmmakeritemch').attr('target', 'view');
    $('#frmmakeritemch').attr('action', '/admin/itemmaster/item_change_all_process.asp');
    $('#mode').val('makerchmargin');
    $('#frmmakeritemch').submit();
    $('#mode').val('');
}

function chmakersellyn_n(){
    if ( $('#makerid').val() == ''){
        alert('변경하실 브랜드ID를 선택해 주세요.');
        return;
    }
    $('#frmmakeritemch').attr('target', 'view');
    $('#frmmakeritemch').attr('action', '/admin/itemmaster/item_change_all_process.asp');
    $('#mode').val('makerchsellyn_n');
    $('#frmmakeritemch').submit();
    $('#mode').val('');
}

function chMoveMaker(){
    if ( $('#makerid').val() == ''){
        alert('변경하실 브랜드ID를 선택해 주세요.');
        return;
    }
    if ( $('#toMakerid').val() == ''){
        alert('이동될 브랜드ID를 선택해 주세요.');
        return;
    }
    $('#frmmakeritemch').attr('target', 'view');
    $('#frmmakeritemch').attr('action', '/admin/itemmaster/item_change_all_process.asp');
    $('#mode').val('MoveMaker');
    $('#frmmakeritemch').submit();
    $('#mode').val('');
}
</script>
<form name="frmmakeritemch" id="frmmakeritemch" method="post" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" id="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td>
		변경할 브랜드ID : 
		<input type="text" class="text" name="makerid" id="makerid" value="" size="15" maxlength=32 >
		<input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'makerid');" >
    </td>
</tr>
<tr class="a" height="25" bgcolor="#FFFFFF">
    <td>
        해당 브랜드 상품 모두 계약구분(
		<label><input type="radio" name="mwdiv" value="M" checked>매입</label>
		<label><input type="radio" name="mwdiv" value="W" >위탁</label>
		<label><input type="radio" name="mwdiv" value="U" >업체</label>
        )으로 변경. 브랜드 대표 마진은 직접 변경 하셔야 합니다.
        <input class="button" type="button" value="계약구분변경하기" onClick="chmakermwdiv();">
    </td>
</tr>
<tr class="a" height="25" bgcolor="#FFFFFF">
    <td>
        해당 브랜드 상품 모두 마진(<input type="text" class="text" name="margin" id="margin" value="0" size="5" maxlength=5 >
        )% 으로 변경.
        <input class="button" type="button" value="마진변경하기" onClick="chmakermargin();">
    </td>
</tr>
<tr class="a" height="25" bgcolor="#FFFFFF">
    <td>
        해당 브랜드 상품 모두 <input class="button" type="button" value="판매안함으로변경하기" onClick="chmakersellyn_n();">
    </td>
</tr>
<tr class="a" height="25" bgcolor="#FFFFFF">
    <td>
        해당 브랜드 상품 모두
		<input type="text" class="text" name="toMakerid" id="toMakerid" value="" size="15" maxlength="32" />
		<input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'toMakerid');" />
        <input class="button" type="button" value="브랜드로 변경하기" onClick="chMoveMaker();" />
        ※ 마진 및 계약구분은 별도로 변경해야합니다.
    </td>
</tr>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
    <iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" ></iframe>
<% else %>
    <iframe id="view" name="view" src="" width="100%" height=0 frameborder="0" ></iframe>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->