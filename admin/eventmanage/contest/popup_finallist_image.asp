<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/contestCls.asp"-->

<%
Dim cPollList, i, page , idx , image_url , link_url, isusing ,regdate, vContest, vPollIdx, vIdx, vImgCode, vImgName, vImgName2, vSortNo
vContest = Request("contest")
vPollIdx = Request("usernum")
vIdx = Request("idx")

set cPollList = new ClsContest
cPollList.FContest = vContest
cPollList.FUserNum = vPollIdx

If vIdx <> "" Then
	cPollList.FIdx = vIdx
	cPollList.FPollImageDetail
	vImgCode = cPollList.FOneItem.fimg_code
	vImgName = cPollList.FOneItem.fimg_name
	vImgName2 = cPollList.FOneItem.fimg_name2
	vSortNo  = cPollList.FOneItem.fsortno
Else
	vSortNo = "0"
End IF

cPollList.FPollImageList
%>

<script language="javascript">

document.domain = "10x10.co.kr";

function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

	window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

//저장
function reg(){
	if (frm.code.value==''){
		alert('이미지구분을 입력해주세요');
		frm.code.focus();
		return;
	}
	
	if (frm.image_url.value==''){
		alert('이미지를 올려주세요');
		return;
	}
	
	if (frm.code.value=='2' && frm.image_url2.value==''){
		alert('작은 이미지도 올려주세요');
		return;
	}
	
	if (frm.sortno.value==''){
		alert('정렬번호를 입력해주세요');
		frm.sortno.focus();
		return;
	}
	
	if (isNaN(frm.sortno.value)){
		alert('정렬번호를 숫자로 입력해주세요');
		frm.sortno.value = "0";
		frm.sortno.focus();
		return;
	}
	
	frm.action='popup_finallist_image_proc.asp';
	frm.submit();
}

function del_image(idx)
{
	if(confirm("선택하신 이미지를 삭제하시겠습니까?\n삭제 후 복구 불가능합니다.") == true) {
		location.href = "popup_finallist_image_proc.asp?gubun=del&contest=<%=vContest%>&poll_idx=<%=vPollIdx%>&idx="+idx+"";
	}
}

function viewtr(cd)
{
	if(cd == 2)
	{
		document.getElementById("img2").style.display = "block";
	}
	else
	{
		document.getElementById("img2").style.display = "none";
	}
}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="idx" value="<%=vIdx%>">
<input type="hidden" name="poll_idx" value="<%=vPollIdx%>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="80">공모전 No.</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;<%= vContest %><input type="hidden" name="contest" value="<%= vContest %>">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>이미지구분</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;<%=FImageCodeList(vImgCode,"onchange='viewtr(this.value)'")%>			
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image_urldiv','image_url','<%=vContest%>','2000','235','true');"/>		
		<input type="hidden" name="image_url" value="<%= vImgName %>">
		<div align="right" id="image_urldiv"><% IF vImgName<>"" THEN %><img src="<%=webImgUrl%>/contest/<%=vContest%>/<%= vImgName %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30" id="img2" style="display:none;">
	<td>작은 이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image_urldiv2','image_url2','<%=vContest%>','2000','235','true');"/>		
		<input type="hidden" name="image_url2" value="<%= vImgName2 %>">
		<div align="right" id="image_urldiv2"><% IF vImgName2<>"" THEN %><img src="<%=webImgUrl%>/contest/<%=vContest%>/<%= vImgName2 %>" width=50 height=50 ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>정렬번호</td>
	<td bgcolor="#FFFFFF" align="left">
		&nbsp;<input type="text" name="sortno" value="<%= vSortNo %>" size="3"> (숫자가 클수록 앞에 나옵니다.)
	</td>
</tr>
<%
	If vIdx <> "" AND vImgCode = 2 Then
		Response.Write "<script language='javascript'>document.getElementById(""img2"").style.display='';</script>"
	End If
%>
<tr align="center" bgcolor="FFFFFF" height="30">
	<td colspan="2">
		<table width="100%">
		<tr>
			<td>&nbsp;<input type="button" onclick="location.href='popup_finallist_image.asp?contest=<%= vContest %>&usernum=<%=vPollIdx%>';" value="새글" class="button"></td>
			<td align="right"><input type="button" onclick="reg();" value="저장" class="button">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>

<br>
※ poll_list_image 인 경우<br>Front단 리스트 순서가 아래 리스트 순서대로 그대로 나타납니다.
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">	
	<td align="center">번호</td>
	<td align="center">이미지구분</td>
	<td align="center">이미지</td>
	<td align="center"></td>
</tr>
<% for i=0 to cPollList.FresultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= i+1 %></td>
	<td align="center"><%= cPollList.FItemList(i).fcodename %></td>
	<td align="center"><img src="<%=webImgUrl%>/contest/<%=vContest%>/<%= cPollList.FItemList(i).fimg_name %>" width=30 height=30 ></td>
	<td align="center">
		<input type="button" value="수정" onClick="location.href='popup_finallist_image.asp?contest=<%=vContest%>&usernum=<%=vPollIdx%>&idx=<%= cPollList.FItemList(i).fidx %>'">
		<input type="button" value="삭제" onClick="del_image('<%= cPollList.FItemList(i).fidx %>')">
	</td>
</tr>
<% next %>
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

<% set cPollList = nothing %>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->