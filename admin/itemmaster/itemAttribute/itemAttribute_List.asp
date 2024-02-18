<%@ language=vbscript %>
<% option explicit %>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%
'###############################################
' Discription : 상품속성 관리
' History : 2013.08.02 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim attribDiv, attribUsing, dispCate
Dim oAttrib, lp
Dim page

'// 파라메터 접수
attribDiv = request("attribDiv")
attribUsing = request("attribUsing")
dispCate = request("dispCate")
page = request("page")
if attribUsing="" then attribUsing="Y"
if page="" then page="1"


'// 페이지정보 목록
	set oAttrib = new CAttrib
	oAttrib.FPageSize = 20
	oAttrib.FCurrPage = page
	oAttrib.FRectattribDiv = attribDiv
	oAttrib.FRectattribUsing = attribUsing
    oAttrib.GetAttribList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function() {
	//속성구분 로드
	chgDispCate("<%=dispCate%>","<%=attribDiv%>");

	$("input[type=submit]").button();

  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// 행 정렬
	$( "#attrList" ).sortable({
		placeholder: "ui-state-highlight",
		handle: ".rowHaddle",
		start: function(event, ui) {
			ui.placeholder.html('<td height="24" colspan="8" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).parent().find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});

function popAttribute(attrCd){
    var popwin = window.open('popItemAttribEdit.asp?attribCd='+attrCd,'popAttribManage','width=450,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popLinkItem(attrCd) {
    var popwin = window.open('popItemAttribLinkItem.asp?attribCd='+attrCd,'popAttribManage','width=1200,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function chkAllItem() {
	if($("input[name='chkCd']:first").attr("checked")=="checked") {
		$("input[name='chkCd']").attr("checked",false);
	} else {
		$("input[name='chkCd']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkCd']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 상품속성을 선택해주세요.");
		return;
	}
	if(confirm("선택하신 속성들을 지정하신 정보로 저장하시겠습니까?")) {
		document.frmList.target="_self";
		document.frmList.action="doItemAttrModify.asp";
		document.frmList.submit();
	}
}

function chgDispCate(dc,ad) {
	// 전시카테고리 선택에 따른 상품속성 선택상자 변경
	$.ajax({
		url: "act_itemAttrSelectBox.asp?dispcate="+dc+"&attribDiv="+ad,
		cache: false,
		success: function(message)
		{
			$("#attrSelBox").empty().append(message);
		}
	});

}
</script>
<!-- 상단 검색폼 시작 -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="page" value="" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
	    전시카테고리:
	    <%=getDispCateSelectbox("dispCate",dispCate,"onchange='chgDispCate(this.value)'")%>
	    &nbsp;/&nbsp;
	    속성구분:
		<span id="attrSelBox"></span>
		&nbsp;/&nbsp;
	    사용구분:
		<select name="attribUsing" class="select">
			<option value="A">전체</option>
			<option value="Y" <%=chkIIF(attribUsing="Y","selected","")%> >사용함</option>
			<option value="N" <%=chkIIF(attribUsing="N","selected","")%> >사용안함</option>
		</select>
	</td>
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="검색" />
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 5px 0;">
<tr>
    <td align="left">
    	<input type="button" value="선택저장" class="button" onClick="saveList()" title="우선순위 및 사용여부를 일괄저장합니다.">
    </td>
    <td align="right">
    	<input type="button" value="신규속성 등록" class="button" onClick="popAttribute('');">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 목록 시작 -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="attrArr">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%=oAttrib.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oAttrib.FtotalPage%></b>
	</td>
</tr>
<colgroup>
	<col width="40" />
    <col width="50" />
    <col width="80" />
    <col width="*" />
    <col width="*" />
    <col width="70" />
    <col width="140" />
	<col width="160" />
</colgroup>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><span class="ui-icon ui-icon-arrowthick-2-n-s"></span></td>
    <td><input type="checkbox" name="allChk" onclick="chkAllItem()"></td>
    <td>속성코드</td>
    <td>속성구분</td>
    <td>속성명</td>
    <td>우선<br>순위</td>
    <td>사용여부</td>
	<td><span class="ui-icon ui-icon-wrench"></span></td>
</tr>
<tbody id="attrList">
<%	for lp=0 to oAttrib.FResultCount - 1 %>
<tr align="center" bgcolor="<%=chkIIF(oAttrib.FItemList(lp).FattribUsing="N","#DDDDDD","#FFFFFF")%>">
	<td><span class="rowHaddle ui-icon ui-icon-grip-solid-horizontal" style="cursor:grab;" title="정렬순서를 변경합니다."></span></td>
    <td><input type="checkbox" name="chkCd" value="<%=oAttrib.FItemList(lp).FattribCd%>" /></td>
    <td><%=oAttrib.FItemList(lp).FattribCd%></td>
    <td><%="[" & oAttrib.FItemList(lp).FattribDiv & "] " & oAttrib.FItemList(lp).FattribDivName %></td>
    <td align="left"><%=oAttrib.FItemList(lp).FattribName & chkIIF(oAttrib.FItemList(lp).FattribNameAdd<>""," / " & oAttrib.FItemList(lp).FattribNameAdd,"") %></td>
    <td><input type="text" name="sort<%=oAttrib.FItemList(lp).FattribCd%>" size="3" class="text" value="<%=oAttrib.FItemList(lp).FattribSortNo%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oAttrib.FItemList(lp).FattribCd%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oAttrib.FItemList(lp).FattribUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">사용함</label><input type="radio" name="use<%=oAttrib.FItemList(lp).FattribCd%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oAttrib.FItemList(lp).FattribUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">사용안함</label>
		</span>
    </td>
	<td>
		<input type="button" value="속성수정" onclick="popAttribute('<%=oAttrib.FItemList(lp).FattribCd%>')" class="ui-button ui-corner-all" style="font-size:11px;" />
		<input type="button" value="상품연결" onclick="popLinkItem('<%=oAttrib.FItemList(lp).FattribCd%>')" class="ui-button ui-corner-all" style="font-size:11px;" />
	</td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="8" align="center">
    <% if oAttrib.HasPreScroll then %>
		<a href="javascript:goPage('<%= oAttrib.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oAttrib.StartScrollPage to oAttrib.FScrollCount + oAttrib.StartScrollPage - 1 %>
		<% if lp>oAttrib.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oAttrib.HasNextScroll then %>
		<a href="javascript:goPage('<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<!-- 목록 끝 -->
<%
	set oAttrib = Nothing
%>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->