<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/shintvshopping/shintvshoppingCls.asp"-->
<%
Dim midx, page, i, shipCostCode, itemid
page		= request("page")
midx 		= request("midx")
shipCostCode	= request("shipCostCode")
itemid  	= request("itemid")

If page = "" Then page = 1

If NOT isNumeric(midx) Then
	Response.Write "<script language=javascript>alert('잘못된 접근입니다.');window.close();</script>"
	dbget.close()	:	response.End
End If

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

Dim oShintvshoppingMaster
Set oShintvshoppingMaster = new CShintvshopping
	oShintvshoppingMaster.FCurrPage			= page
	oShintvshoppingMaster.FPageSize			= 50
	oShintvshoppingMaster.FRectMasterIdx	= midx
	oShintvshoppingMaster.FRectShipCostCode	= shipCostCode
	oShintvshoppingMaster.FRectItemID		= itemid
	oShintvshoppingMaster.getssgMarginItemDetailList
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function popCateSelect(){
	$.ajax({
		url: "/admin/etc/ssg/act_CategorySelect.asp",

		cache: false,
		success: function(message) {
			$("#lyrCateAdd").empty().append(message).fadeIn();
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function jsAddItemID() {
	var frm = document.frm;

	if (frm.itemid.value == '') {
		alert('상품코드를 입력하세요.');
		return;
	}

	if (confirm('저장하시겠습니까?')) {
		frm.delIdx.value = '';
		frm.submit();
	}
}

function delItem(v)
{
	$("#delIdx").val(v);
	document.frm.submit();
}

function goPage(pg){
    //frm.page.value = pg;
    //frm.submit();
	location.href = '?page='+pg+'&midx=<%= midx %>';
}
</script>

<form name="frm" action="procShintvshoppingBeasongCode.asp" methd="post" style="margin:0px;">
<input type="hidden" name="mode" value="itemDetail">
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" id="delIdx" name="delIdx" value="">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
				상품ID :
				<textarea class="textarea" name="itemid" cols="16" rows="2"></textarea>
				<input type="button" value="저 장" onClick="jsAddItemID()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		검색결과 : <b><%= FormatNumber(oShintvshoppingMaster.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oShintvshoppingMaster.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품코드</td>
	<td width="100">관리</td>
</tr>
<% For i=0 to oShintvshoppingMaster.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= oShintvshoppingMaster.FItemList(i).Fitemid %></td>
	<td><input type="button" class="button" value="삭제" onclick="delItem(<%= oShintvshoppingMaster.FItemList(i).FIdx %>);"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oShintvshoppingMaster.HasPreScroll then %>
		<a href="javascript:goPage('<%= oShintvshoppingMaster.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oShintvshoppingMaster.StartScrollPage to oShintvshoppingMaster.FScrollCount + oShintvshoppingMaster.StartScrollPage - 1 %>
    		<% if i>oShintvshoppingMaster.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oShintvshoppingMaster.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% Set oShintvshoppingMaster = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
