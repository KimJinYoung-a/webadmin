<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 한용민 2008프론트에서이동 2009용으로 변경
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/category_main_EventBannerCls.asp"-->

<%
dim mode,i,page ,cdl , cdm , idx
	mode = request("mode")
	page = request("page")
	idx = request("idx")	
	cdl = request("cdl")
	cdm = request("cdm")	
%>

<script language="javascript">

function subcheck(){
	var frm=document.inputfrm;

	if (frm.cdl.value.length<1) {
		alert('카테고리를 선택해주세요..');
		frm.cdl.focus();
		return;
	}
	
	if (frm.evt_code.value.length< 1 ){
		 alert('이벤트 번호를 입력해주세요');
	frm.evt_code.focus();
	return;
	}

	if (frm.viewidx.value.length< 1 ){
		 alert('표시순서를 숫자로 입력해주세요.');
	frm.viewidx.focus();
	return;
	}

	if (frm.cdl.value == '110'){
		if (frm.cdm.value==''){
			alert('감성채널은 중카테고리를 선택해야만 합니다');			
			return;
		}
	}

	frm.submit();
}

function chimg(im,v){

	frm=eval("document." + v);
	frm.src=im;
}

function popEventList(){
	var frm=document.inputfrm;

	if (frm.cdl.value.length<1) {
		alert('카테고리를 선택해주세요..');
		frm.cdl.focus();
		return;
	}
	
	window.open('ViewEventList_Main_EventBanner.asp?selC=010','popasd','width=800,height=600,scrollbars=yes');
}

function changecontent()
{
	document.inputfrm.action='?';
	document.inputfrm.submit();

}

</script>
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="20">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top"><b>카테고리 메인 이벤트 베너 등록/수정</b></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="inputfrm" method="post" action="doMainEventBanner.asp">
<input type="hidden" name="mode" value="<% =mode %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<% if mode="add" then %>
<tr>
	<td width="100" bgcolor="#F0F0FD" align="center">카테고리선택</td>
	<td bgcolor="#FFFFFF">
<select class='select' name="cdl">
<option value='010' selected>디자인문구</option>
</select>
<select class='select' name="cdm">
<option value='' <% If cdm = "" Then Response.Write "selected" End If %>>전체다이어리</option>
<option value='10' <% If cdm = "10" Then Response.Write "selected" End If %>>심플다이어리</option>
<option value='20' <% If cdm = "20" Then Response.Write "selected" End If %>>일러스트다이어리</option>
<option value='30' <% If cdm = "30" Then Response.Write "selected" End If %>>캐릭터다이어리</option>
<option value='40' <% If cdm = "40" Then Response.Write "selected" End If %>>포토다이어리</option>
</select>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">이벤트 번호</td>
	<td bgcolor="#FFFFFF"><input type="text" name="evt_code" size="8"><input type="button" name="evtbtn" class="button" value="검색" onclick="popEventList();"></td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">표시순서</td>
	<td bgcolor="#FFFFFF"><input type="text" name="viewidx" size="4"></td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">사용유무</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" checked>Y
		<input type="radio" name="isusing" value="N">N
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td colspan="2" align="center">
			<input type="button" value=" 저장 " onclick="subcheck();"> &nbsp;&nbsp;
			<input type="button" value=" 취소 " onclick="history.back();">
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New CateEventBanner
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.frectidx = idx
	fmainitem.GetEventBannerList

if cdl = "" then cdl = fmainitem.FItemList(0).fcdl
if cdm = "" then cdm = fmainitem.FItemList(0).Fcdm
%>
<tr>
	<td width="100" align="center" bgcolor="#F0F0FD">카테고리</td>
	<td bgcolor="#FFFFFF">
<select class='select' name="cdl">
<option value='010' selected>디자인문구</option>
</select>
<select class='select' name="cdm">
<option value='' <% If cdm = "" Then Response.Write "selected" End If %>>전체다이어리</option>
<option value='10' <% If cdm = "10" Then Response.Write "selected" End If %>>심플다이어리</option>
<option value='20' <% If cdm = "20" Then Response.Write "selected" End If %>>일러스트다이어리</option>
<option value='30' <% If cdm = "30" Then Response.Write "selected" End If %>>캐릭터다이어리</option>
<option value='40' <% If cdm = "40" Then Response.Write "selected" End If %>>포토다이어리</option>
</select>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">이벤트명</td>
	<td bgcolor="#FFFFFF">
		<%="[" & fmainitem.FItemList(0).Fevt_code & "] " & fmainitem.FItemList(0).Fevt_name %>
		<input type="hidden" name="evt_code" value="<%=fmainitem.FItemList(0).Fevt_code%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">표시순서</td>
	<td bgcolor="#FFFFFF"><input type="text" name="viewidx" size="4" value="<%=fmainitem.FItemList(0).FviewIdx%>"></td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">사용유무</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%if fmainitem.FItemList(0).FIsusing="Y" then response.write "checked" %> checked>Y
		<input type="radio" name="isusing" value="N" <%if fmainitem.FItemList(0).FIsusing="N" then response.write "checked" %>>N
		<input type="hidden" name="orgUsing" value="<%=fmainitem.FItemList(0).FIsusing%>">
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td colspan="2" align="center">
		<input type="button" value=" 저장 " onclick="subcheck();"> &nbsp;&nbsp;
		<input type="button" value=" 취소 " onclick="history.back();">
	</td>
</tr>
<% end if %>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
