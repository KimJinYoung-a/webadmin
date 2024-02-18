<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/pop_dealitem_group.asp
' Description :  딜 상품 그룹등록
' History : 2022.10.17 정태훈 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<%
Dim idx : idx = Request("idx")
Dim groupCode : groupCode = Request("groupCode")
Dim sTarget : sTarget = request("sTarget")
dim cdealGroup, cdealGroupDetail, arrList, title, sort, maxSortNum

set cdealGroup = new CDealSelect
cdealGroup.FRectDealCode = idx
arrList = cdealGroup.fnGetDealItemGroup
maxSortNum=cdealGroup.fnGetDealItemGroupSortInfo
set cdealGroup = nothing

if maxSortNum="" or isnull(maxSortNum) then maxSortNum=0

if groupCode <> "" then
set cdealGroupDetail = new CDealSelect
cdealGroupDetail.FRectDealCode = idx
cdealGroupDetail.FRectGroupCode = groupCode
cdealGroupDetail.fnGetDealItemGroupDetail
title = cdealGroupDetail.Ftitle
sort = cdealGroupDetail.Fsort
set cdealGroupDetail = nothing
end if

if sort="" or isnull(sort) then sort=maxSortNum+1

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
$(function(){
	$("#accordion").accordion();
	//드래그
	$("#subList").sortable({
		placeholder: "ui-state-highlight",
		cancel : ".sortablearrow",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).find("input[name^='sSort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sSort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
    var update = function() {
        var now = $('#title').val().length;
		var strValue = $('#title').val();
		var str2="";
		var len = 0;
        if(now>20){
			str2 = strValue.substr(0, 20);
			$('#title').val(str2);
		}
		$('#count').text(now);
		
    };
    // input, keyup, paste 이벤트와 update 함수를 바인드한다
    $('#title').bind('input keyup paste', function() {
        setTimeout(update, 0)
    });
    update();
});

function jsGroupSubmit(){
	if(!document.frmG.title.value){
		alert("그룹명을 입력해주세요");
		document.frmG.title.focus();
		return false;
	}else{
		document.frmG.submit();
	}
}
function jsDelGroup(groupcode){
	if(confirm("정말 삭제하시겠어요?")){
		document.frmGM.groupCode.value=groupcode;
		document.frmGM.submit();
	}
}
function fnRegistGroup(groupcode){
	location.href="/admin/itemmaster/deal/pop_dealitem_group.asp?idx=<%=idx%>&groupcode="+groupcode
}
// 그룹 순서 일괄 저장
function jsGroupSortSave() {
	var frm;
	var sValue, sSort, sDisp ;
	frm = document.frmL;
	sValue = "";
	sSort = "";
	sDisp = ""; 

	var itemid;
	if (frm.groupcode.length > 0){
		for (var i=0;i<frm.groupcode.length;i++){ 
			if(!IsDigit(frm.sSort[i].value)){
				alert("순서지정은 숫자만 가능합니다.");
				frm.sSort[i].focus();
				return;
			}
			itemid = frm.groupcode[i].value;
			if (sValue==""){
				sValue = frm.groupcode[i].value;
			}else{
				sValue =sValue+","+frm.groupcode[i].value;
			}
			// 정렬순서
			if (sSort==""){
				sSort = frm.sSort[i].value;
			}else{
				sSort =sSort+","+frm.sSort[i].value;
			}
    	}
	}
	frm.itemidarr.value = sValue;
	frm.sortarr.value = sSort;
	frm.submit();
}
</script>
<form name="frmGM" method="post" action="dodealitemgroup.asp">
    <input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="mode" value="del">
	<input type="hidden" name="groupCode">
</form>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
	<td align="left">딜 상품 그룹 등록</td>
</tr>
</table>
<hr>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
 <tr>
 	<td>
 		<form name="frmG" method="post" action="dodealitemgroup.asp">
		<input type="hidden" name="idx" value="<%=idx%>">
		<% if groupCode <> "" then %>
		<input type="hidden" name="mode" value="update">
		<% else %>
		<input type="hidden" name="mode" value="add">
		<% end if %>
		<input type="hidden" name="groupCode" value="<%=groupCode%>">
		<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
						<tr>
							<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">그룹명</td>
							<td bgcolor="#FFFFFF"><input type="text" name="title" id="title" size="40" maxlength="32" class="text" value="<%=title%>"><span id="count">0</span>/20자</td>
						</tr>
						<tr>
							<td align="center" bgcolor="<%= adminColor("tabletop") %>">정렬순서</td>
							<td bgcolor="#FFFFFF"><input type="text" size="2" name="sort" class="text" value="<%=sort%>"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td align="center">
		<p>
			<% if groupCode <> "" then %>
	    	<input type="button" class="button" style="height:30px; width:100px;" value="수정" onClick="jsGroupSubmit();">
			<% else %>
			<input type="button" class="button" style="height:30px; width:100px;" value="저장" onClick="jsGroupSubmit();">
			<% end if %>
	    </p>
	</td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
	<td align="left">딜 상품 그룹 관리</td>
</tr>
</table>
<% IF isArray(arrList) THEN %>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>
	    <form name="frmL" method="post" action="dodealitemgroup.asp">
	    <input type="hidden" name="idx" value="<%=idx%>">
		<input type="hidden" name="mode" value="sort">
		<input type="hidden" name="itemidarr">
		<input type="hidden" name="sortarr">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center">
				<td bgcolor="<%= adminColor("tabletop") %>" width="15%">그룹코드</td>
				<td bgcolor="<%= adminColor("tabletop") %>">그룹명</td>
				<td bgcolor="<%= adminColor("tabletop") %>">정렬</td>
				<td width="80" bgcolor="<%= adminColor("tabletop") %>" width="15%">관리</td>
			</tr>
			<tbody id="subList">
			<%dim sumi,i ,eGCMoArr, intg %>
			<% FOR intg = 0 To UBound(arrList,2)
				sumi = 0
				eGCMoArr = arrList(0,intg)
			%>
			<tr bgcolor="#ffffff">
				<td align="center"><%=arrList(0,intg)%><input type="hidden" name="groupcode" value="<%=arrList(0,intg)%>"></td>
				<td align="center"><%=db2html(arrList(1,intg))%></td>
				<td align="center" class="sortablearrow"><input type="text" size="2" name="sSort" class="text" value="<%=db2html(arrList(2,intg))%>"></td>
				<td align="center" class="sortablearrow">
					<input type="button" name="btnD" value="수정" onclick="fnRegistGroup(<%=arrList(0,intg)%>)"  class="button">
					<input type="button" name="btnD" value="삭제" onclick="jsDelGroup(<%=arrList(0,intg)%>)"  class="button">
				</td>
			</tr>
			<%   intg = intg+sumi
			NEXT%>
			</tbody>
			</table>
	</form>
	</td>
</tr>
<tr>
	<td align="center">
		<p>
			<input type="button" class="button" style="height:30px; width:100px;" value="저장" onClick="jsGroupSortSave();">
	    </p>
	</td>
</tr>
</table>
<% END IF %>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
