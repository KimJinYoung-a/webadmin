<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/popManageColorCode.asp
' Description :  상품 컬러 코드등록
' History : 2009.03.24 허진원 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Dim oitem, sMode, iColorCD, lp
Dim sColorName, sColorIcon, iSortNo, sIsUsing
iColorCD = Request.Querystring("iCD")

'// 기본값
sMode = "I"	'등록

'// 색상코드가 있으면 수정모드
if iColorCD<>"" then
	sMode = "U"	'수정
	set oitem = new CItemColor
	oitem.FRectColorCD = iColorCD
	oitem.GetColorList

	if oitem.FResultCount>0 then
		sColorName	= oitem.FItemList(0).FcolorName
		sColorIcon	= oitem.FItemList(0).FcolorIcon
		iSortNo		= oitem.FItemList(0).FsortNo
		sIsUsing	= oitem.FItemList(0).FisUsing
	else
		Alert_return("잘못된 번호입니다.")
		dbget.close()	:	response.End
	end if

	set oitem = Nothing
end if
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.scName.value){
			alert("컬러명을 입력해주세요.");			
			return false;
		}

		if((!document.frmImg.scIcon.value)&&document.frmImg.mode.value=="I"){
			alert("찾아보기 버튼을 눌러 업로드할 컬러칩 이미지를 선택해 주세요.");			
			return false;
		}

		if(!document.frmImg.icSort.value){
			alert("정렬번호를 숫자로 입력해주세요.");			
			return false;
		}
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 상품 색상코드 관리</div>
<table width="350" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/items/itemColorCodeProcess.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="<%=sMode%>">
<% if sMode="U" then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">컬러코드</td>
	<td bgcolor="#FFFFFF"><input type="text" name="icCode" size="4" readonly value="<%=iColorCD%>"></td>
</tr>
<% end if %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">컬러명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scName" size="10" maxlength="12" value="<%=sColorName%>"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">컬러칩아이콘</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="scIcon">
		<% IF sColorIcon <> "" THEN %>
			<br>현재 파일명 : <%=right(sColorIcon,len(sColorIcon)-instrRev(sColorIcon,"/"))%>
		<% END IF %>
	</td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">정렬번호</td>
	<td bgcolor="#FFFFFF"><input type="text" name="icSort" size="4" maxlength="4" style="text-align:right" value="<%=iSortNo%>"></td>
</tr>
<% if sMode="U" then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="scUse" value="Y" <% if sIsUsing="Y" then Response.Write "checked" %>>사용
		<input type="radio" name="scUse" value="N" <% if sIsUsing="N" then Response.Write "checked" %>>삭제
	</td>
</tr>
<% end if %>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<% if sMode="I" then %>
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		<% Else %>
		<a href="popManageColorCode.asp"><img src="/images/icon_cancel.gif" border="0"></a>
		<% end if %>
	</td>
</tr>
</form>
</table>
<br>
<%
	'####### 컬러칩 목록 #######
	set oitem = new CItemColor
	oitem.FPageSize = 50
	oitem.FRectUsing = "Y"
	oitem.GetColorList
%>
<table width="350" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td bgcolor="#DDDDFF">코드</td>
	<td bgcolor="#DDDDFF">Icon</td>
	<td bgcolor="#DDDDFF">코드명</td>
	<td bgcolor="#DDDDFF">정렬번호</td>
	<td bgcolor="#DDDDFF">사용</td>
</tr>
<%
	if oitem.FResultCount>0 then
		for lp=0 to oitem.FResultCount-1
%>
<tr align="center">
	<td bgcolor="#FFFFFF"><%=oitem.FItemList(lp).FcolorCode%></td>
	<td bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="1" bgcolor="#dddddd"><tr><td bgcolor="#FFFFFF"><img src="<%=oitem.FItemList(lp).FcolorIcon%>" width="12" height="12" hspace="2" vspace="2"></td></tr></table></td>
	<td bgcolor="#FFFFFF"><a href="popManageColorCode.asp?iCD=<%=oitem.FItemList(lp).FcolorCode%>"><%=oitem.FItemList(lp).FcolorName%></a></td>
	<td bgcolor="#FFFFFF"><%=oitem.FItemList(lp).FsortNo%></td>
	<td bgcolor="#FFFFFF"><%=oitem.FItemList(lp).FisUsing%></td>
</tr>
<%
		next
	else
		Response.Write "<tr><td colspan=5 height=50 align=center bgcolor=#F8F8F8>등록된 색상이 없습니다.</td></tr>"
	end if
%>
</table>
<%
	set oitem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->