<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/vipcornerCls.asp"-->
<%
'###############################################
' PageName : vip_write.asp
' Discription : 우수회원 전용코너 관리
' History : 2015.04.15 원승현 생성
'###############################################

dim justDate,mode,i, idx, vIdx
mode=request("mode")
vIdx = request("idx")

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--

document.domain = "10x10.co.kr";

function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb,orgImgName){

	window.open("vip_PopImgInput.asp?divName="+divnm+"&inputname="+iptNm+"&ImagePath="+vPath+"&maxFileSize="+Fsize+"&maxFileWidth="+Fwidth+"&makeThumbYn="+thumb+"&orgImgName="+orgImgName,'imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
}

function subcheck(){
	var frm=document.inputfrm;


	if(!frm.evt_code.value) {
		alert("이벤트코드를 넣어주세요.");
		return;
	}
	if(!frm.image1.value) {
		alert("pc용 이미지를 넣어주세요.");
		return;
	}

	frm.submit();
}

function delitems()
{
	if (confirm("삭제 하시겠습니까?")){
		document.inputfrm.mode.value="delete";
		document.inputfrm.submit();
	}
}

//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="dovip_Process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<input type="hidden" name="idx" value="<% =vIdx %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>우수회원전용코너 등록/수정</b></font>
	</td>
</tr>
<% if mode="add" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="evt_code" value="" size="10">&nbsp;
		<input type="button" value="이벤트 코드 찾기" onclick="window.open('/admin/eventmanage/event/?menupos=870', '', 'width=1200, height=600, toolbar=yes, scrollbars=yes');" class="button">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>pc용 이미지 (390*189)</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image1div','image1','i1','390','189','false','');"/>
		<input type="hidden" name="image1" value="">
		<div align="right" id="image1div"></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>모바일/앱용 이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image2div','image2','i2','390','189','true','');"/>		
		<input type="hidden" name="image2" value="">
		<div align="right" id="image2div"></div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">순서</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="orderby" value="" size="5">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
	<td bgcolor="#FFFFFF">
		<select name="isusing">
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New CVip
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.FRectidx=vIdx
	fmainitem.GetVipCornerModify
%>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="evt_code" value="<%=fmainitem.FItemList(0).FevtCode%>" size="10">&nbsp;
		<input type="button" value="이벤트 코드 찾기" onclick="window.open('/admin/eventmanage/event/?menupos=870', '', 'width=1200, height=600, toolbar=yes, scrollbars=yes');" class="button">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>pc용 이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image1div','image1','i1','390','189','false','');"/>
		<input type="hidden" name="image1" value="<%= fmainitem.FItemList(0).Fpcimg %>">
		<div align="right" id="image1div"><% IF fmainitem.FItemList(0).Fpcimg<>"" THEN %><img src="<%=webImgUrl%>/vipcorner/<%= fmainitem.FItemList(0).Fpcimg %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>모바일/앱용 이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image2div','image2','i2','390','189','true','');"/>		
		<input type="hidden" name="image2" value="<%= fmainitem.FItemList(0).Fmaing %>">
		<div align="right" id="image2div"><% IF fmainitem.FItemList(0).Fmaing<>"" THEN %><img src="<%=webImgUrl%>/vipcorner/<%= fmainitem.FItemList(0).Fmaing %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">순서</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="orderby" value="<%= fmainitem.FItemList(0).Forderby %>" size="5">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
	<td bgcolor="#FFFFFF">
		<select name="isusing">
			<option value="Y" <% If fmainitem.FItemList(0).Fisusing="Y" Then %> selected <% End If %>>Y</option>
			<option value="N" <% If fmainitem.FItemList(0).Fisusing="N" Then %> selected <% End If %>>N</option>
		</select>
	</td>
</tr>

<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" 저장 " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<% if mode="edit" then %><input type="button" value=" 삭제 " class="button" onclick="delitems();"> &nbsp;&nbsp;<% end if %>
		<input type="button" value=" 취소 " class="button" onclick="window.close();">
	</td>
</tr>
</form>
</table>

<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
