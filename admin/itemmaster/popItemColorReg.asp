<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/popItemColorReg.asp
' Description :  상품 컬러 등록
' History : 2009.03.28 허진원 생성
'           2011.04.22 허진원 : 수정시 해당 상품의 관련 색상 목록 출력
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Dim oitem, sMode, iColorCD, itemId, lp
Dim sColorName, sColorIcon, sitemname, slistImage
iColorCD = Request.Querystring("iCD")
itemId = Request.Querystring("iid")

'// 기본값
sMode = "I"	'등록

'// 색상코드가 있으면 수정모드
if iColorCD<>"" then
	sMode = "U"	'수정
	set oitem = new CItemColor
	oitem.FRectColorCD = iColorCD
	oitem.FRectitemid = itemid
	oitem.GetColorItemList

	if oitem.FResultCount>0 then
		sColorName	= oitem.FItemList(0).FcolorName
		sColorIcon	= oitem.FItemList(0).FcolorIcon
		sitemname	= oitem.FItemList(0).FitemName
		slistImage	= oitem.FItemList(0).FlistImage
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
		if(!document.frmItemColor.iid.value){
			alert("상품코드를 입력해주세요.");
			document.frmItemColor.iid.focus();
			return false;
		}

		if(!document.frmItemColor.iCD.value){
			alert("컬러칩을 선택해주세요.");
			return false;
		}

		if((!document.frmItemColor.sBasicImage.value)&&document.frmItemColor.mode.value=="I"){
			alert("찾아보기 버튼을 눌러 업로드할 상품 이미지를 선택해 주세요.");			
			return false;
		}
	}

	//색상코드 선택
	function selColorChip(cd) {
		var i;
		document.frmItemColor.iCD.value= cd;
		for(i=0;i<=30;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}

	// 색상삭제
	function delItemColor() {
		if(confirm("지금 보시는 상품색상 정보를 삭제하시겠습니까?\n\n※삭제가 완료되면 다시 복구할 수 없습니다.")) {
			document.frmItemColor.mode.value="D";
			document.frmItemColor.submit();
		}
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 상품/색상 등록</div>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmItemColor" method="post" action="<%= uploadImgUrl %>/linkweb/items/itemColorProcess.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="<%=sMode%>">
<input type="hidden" name="iCD" value="<%=iColorCD%>">
<input type="hidden" name="oCD" value="<%=iColorCD%>">
<% if sMode="I" then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
	<td bgcolor="#FFFFFF"><input type="text" name="iid" size="10" maxlength="8"></td>
</tr>
<% else %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
	<td bgcolor="#FFFFFF"><input type="text" name="iid" size="10" readonly value="<%=itemid%>"></td>
</tr>
<% end if %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">선택색상</td>
	<td bgcolor="#FFFFFF"><%=FnSelectColorBar(iColorCd,13)%></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">색상별 상품이미지</td>
	<td bgcolor="#FFFFFF">
		<table border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="file" name="sBasicImage"></td>
		</tr>
		<tr>
			<td colspan="2"><font color="#808080">※이미지는 1000px×1000px, 560kb이하의 JPG형식</font></td>
		</tr>
		<% IF slistImage <> "" THEN %>
		<tr>
			<td valign="top">현재 이미지 :</td>
			<td><img src="<%=slistImage%>" width="100" border="0" align="absmiddle"></td>
		</tr>
		<% END IF %>
		</table>
	</td>
</tr>	
<tr>
	<td colspan="2" bgcolor="#FFFFFF">
		<table width="100%" border="0" class="a">
		<tr>
			<td>
				<% if sMode="U" then %>
				<a href="javascript:delItemColor();"><img src="/images/icon_delete.gif" border="0"></a>
				<% end if %>
			</td>
			<td align="right">
				<input type="image" src="/images/icon_confirm.gif">
				<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<%
'// 수정모드이면 상품관련 색상목록 출력
if sMode="U" then

	set oitem = new CItemColor
	oitem.FRectItemId	= itemid
	oitem.FPageSize		= 30
	oitem.FCurrPage		= 1
	oitem.FRectUsing	= "Y"
	oitem.GetColorItemList
%>
<!-- 리스트 시작 -->
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">검색결과 : <b><%= oitem.FTotalCount%></b></td>
	</tr>
	</form>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>상품이미지</td>
		<td>컬러칩</td>
		<td>상품명</td>
		<td>등록일시</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="4" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for lp=0 to oitem.FresultCount-1 %>
	<tr align="center">
		<td bgcolor="#FFFFFF"><a href="popItemColorReg.asp?iCD=<%=oitem.FItemList(lp).FcolorCode%>&iid=<%=oitem.FItemList(lp).FitemId%>"><img src="<%=oitem.FItemList(lp).FsmallImage%>" border="0" width="50"></a></td>
		<td bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="1" bgcolor="#dddddd"><tr><td bgcolor="#FFFFFF"><img src="<%=oitem.FItemList(lp).FcolorIcon%>" width="12" height="12" hspace="2" vspace="2"></td></tr></table></td>
		<td bgcolor="#FFFFFF"><a href="popItemColorReg.asp?iCD=<%=oitem.FItemList(lp).FcolorCode%>&iid=<%=oitem.FItemList(lp).FitemId%>"><%=oitem.FItemList(lp).Fitemname%></a></td>
		<td bgcolor="#FFFFFF"><%=left(oitem.FItemList(lp).Fregdate,10)%></td>
    </tr>
	<% next %>
</table>
<%
	end if
	set oitem = Nothing
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->