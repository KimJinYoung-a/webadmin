<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 다이 상품 등록 대기 상품 
' Hieditor : 2010.10.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemRegCls.asp"-->
<%

Dim owaititem,ix,page,itemname, i

page = requestCheckvar(request("page"),10)
if (page="") then page=1
itemname = requestCheckvar(request("itemname"),64)

  	if itemname <> "" then
		if checkNotValidHTML(itemname) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if

set owaititem = new CWaitItemlist
owaititem.FPageSize = 20
owaititem.FCurrPage = page
owaititem.FRectDesignerID = session("ssBctID")
owaititem.FRectitemname = itemname
owaititem.FRectCurrState = "junstnotreged"  ''등록대기 혹은 등록보류, 등록거부인 상품만 보여줌
owaititem.WaitProductList

%>

<script language='javascript'>

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function ViewItemDetail(itemno){
	var popwin = window.open('/academy/itemmaster/viewDIYitem/viewDIYitem.asp?itemid='+itemno ,'popwin','width=1024,height=960,scrollbars=yes,status=yes');
	popwin.focus();
}

function TnSearchItem(){
	document.frm.page.value = "";
	document.frm.submit();
}
function ChangeOrderMakerFrame(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}

	var ret = confirm('선택 상품을 삭제하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;
				}
			}
		}
		frm.submit();
	}
}

</script>
<script>
// ============================================================================
// 옵션수정
function PopDIYItemOptionEdit(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/academy/comm/pop_diywaititemoptionedit.asp?' + param ,'PopDIYItemOptionEdit','width=700,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// 이미지수정
function PopDIYItemImageEdit(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/academy/comm/pop_itemimage.asp?' + param ,'PopDIYItemImageEdit','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="1">
    <tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="30" >
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
    		상품명 검색 : <input type="text" name="itemname" size="20" value="<%= itemname %>">&nbsp;<a href="javascript:TnSearchItem()"><img src="/admin/images/search2.gif" width="74" height="22" align="absmiddle" border="0"></a>
        </td>
        <td valign="top" align="right">
        	검색결과 : 총 <font color="red"><% = owaititem.FTotalCount %></font>개&nbsp;&nbsp;&nbsp;
        	<input type="button" value="선택상품삭제" onClick="ChangeOrderMakerFrame()">
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>

	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    	<td width="30">선택</td>
			<td width="60"><b>임시</b>코드</td>
			<td width="80">업체코드</td>
			<td>상품명</td>
			<td width="60">판매가</td>
			<td width="100">제조사</td>
			<td width="80">등록요청일</td>
			<td width="60">상태</td>
			<td width="50">옵션</td>
	    </tr>
<% if owaititem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to owaititem.FresultCount-1 %>
	   <form name="frmBuyPrc_<%= ix %>" method="post">
	   <input type="hidden" name="itemid" value="<%= owaititem.FItemList(ix).Fitemid %>">
		<tr class="a" height="25" bgcolor="#FFFFFF">
			<td align="center">
			<% If (owaititem.FItemList(ix).FCurrState <> 7) then %>
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
			<% else %>
			<input type="checkbox" name="cksel" disabled >
			<% End if %>
			</td>
			<td align="center"><%= owaititem.FItemList(ix).Fitemid %></td>
			<td align="center"><%= owaititem.FItemList(ix).Fupchemanagecode %></td>
			<% if owaititem.FItemList(ix).FCurrState="7" then %>
			<td align="left">&nbsp;<% =owaititem.FItemList(ix).Fitemname %>&nbsp;&nbsp;<a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<% =owaititem.FItemList(ix).Flinkitemid %>" target="_blank"><font color="blue">(보기)</font></a></td>
			<% else %>
			<td align="left"><a href="diy_wait_item_modify.asp?itemid=<% =owaititem.FItemList(ix).Fitemid %>&menupos=<%= menupos %>&fingerson=on"><% =owaititem.FItemList(ix).Fitemname %></a>&nbsp;&nbsp;<a href="javascript:ViewItemDetail('<% =owaititem.FItemList(ix).Fitemid %>')"><font color="blue">(미리보기)</font></a></td>
			<% end if %>
			<td align="center"><%= FormatNumber(owaititem.FItemList(ix).Fsellcash,0) %></td>
			<td align="center"><% if owaititem.FItemList(ix).Fmakername="" then %>&nbsp;<% else %><% =owaititem.FItemList(ix).Fmakername %><% end if %></td>
			<td align="center"><% =FormatDateTime(owaititem.FItemList(ix).Fregdate,2) %></td>
			<td align="center">
				<% if owaititem.FItemList(ix).FCurrState="0" or owaititem.FItemList(ix).FCurrState="2" then %>
				<font color="<%= owaititem.FItemList(ix).GetCurrStateColor %>" onmouseover="OnOffMessegeBox('on','<%= owaititem.FItemList(ix).Frejectmsg %>','<%= Left(owaititem.FItemList(ix).FrejectDate,10) %>')"><%= owaititem.FItemList(ix).GetCurrStateName %></font>
				<% else %>
				<font color="<%= owaititem.FItemList(ix).GetCurrStateColor %>"><%= owaititem.FItemList(ix).GetCurrStateName %></font>
				<% end if %>
			</td>
			<td align="center">
            <% if (owaititem.FItemList(ix).FCurrState <> "7") then %>
				<a href="javascript:PopDIYItemOptionEdit('<%= owaititem.FItemList(ix).Fitemid %>')">
				<img src="/images/icon_modify.gif" border="0" align="absbottom">
				</a>
            <% end if %>
			</td>
		</tr>
		</form>
    <% next %>
<% end if %>



<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
		<% if owaititem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= owaititem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + owaititem.StartScrollPage to owaititem.StartScrollPage + owaititem.FScrollCount - 1 %>
			<% if (ix > owaititem.FTotalpage) then Exit for %>
			<% if CStr(ix) = CStr(owaititem.FCurrPage) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if owaititem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>

<form name="frmArrupdate" method="post" action="delwaititemarr.asp">
<input type="hidden" name="mode" value="del">
<input type="hidden" name="itemid" value="">
</form>
<script language="javascript">
<!-- // 내용 미리보기 레이어 컨트롤 및 레이어 생성 //
	function OnOffMessegeBox(sw,msg,dt)
	{
		var mx, my, strMsg;

		//마우스 좌표
		mx = event.clientX;
		my = event.clientY+document.body.scrollTop;
		
		//내용 정리
		strMsg = "<table cellpadding=0 cellspacing=0 border=0 width=230 onmouseout=\"OnOffMessegeBox('off','', '')\" class='a' style='border:#606090 1px solid;'>"
				+ "<tr><td bgcolor=#E8E8EF style='padding:3 3 3 3;'><b>등록보류사유</b></td></tr>"
				+ "<tr><td bgcolor=#FFFFFF style='padding:3 3 3 3;'>" + msg + "</td></tr>"
				+ "<tr><td bgcolor=#F8F8FF style='padding:3 3 3 3;' align=right>" + dt + "</td></tr>"
				+ "</table>";
		
		if(sw=="on")
		{
			document.all.popMessege.style.top = my - 10;
			document.all.popMessege.style.left = mx - 180;
			document.all.popMessege.innerHTML = strMsg;
			document.all.popMessege.style.visibility = 'visible';
		} else	{
			document.all.popMessege.style.visibility = 'hidden';
		}
	}
//-->
</script>
<div name="popMessege" id="popMessege" style="z-index:20; position:absolute; top:10; left:10; visibility:hidden;"></div>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->