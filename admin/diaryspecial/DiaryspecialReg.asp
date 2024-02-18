<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 스페셜 다이어리 등록페이지
'	History		: 2015.10.05 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/diaryspecial/diaryspecialCls.asp"-->
<%

dim ospecial, mode, idx
dim pcmainimage, pcoverimage, pctext
dim mobileimage, mobiletext
dim itemid1, itemid2, itemid3, itemid4, itemid5
dim linkgubun, linkcode, sortnum, isusing, regdate
dim detailitemimage1, detailitemimage2, detailitemimage3, detailitemimage4, detailitemimage5
idx = request("idx")
Mode = request("mode")

IF Mode = "" THEN Mode = "add"
if idx 	<> "" then mode	= "edit"

IF Mode= "edit" THEN
	set ospecial = new CDiaryspecial
	ospecial.FRecidx = idx
	ospecial.fnGetDiaryspecial_oneitem()
		idx			=	ospecial.Foneitem.Fidx
		pcmainimage	=	ospecial.Foneitem.Fpcmainimage
		pcoverimage	=	ospecial.Foneitem.Fpcoverimage
		pctext		=	ospecial.Foneitem.Fpctext
		mobileimage	=	ospecial.Foneitem.Fmobileimage
		mobiletext	=	ospecial.Foneitem.Fmobiletext
		itemid1		=	ospecial.Foneitem.Fitemid1
		itemid2		=	ospecial.Foneitem.Fitemid2
		itemid3		=	ospecial.Foneitem.Fitemid3
		itemid4		=	ospecial.Foneitem.Fitemid4
		itemid5		=	ospecial.Foneitem.Fitemid5
		linkgubun	=	ospecial.Foneitem.Flinkgubun
		linkcode	=	ospecial.Foneitem.Flinkcode
		sortnum		=	ospecial.Foneitem.Fsortnum
		isusing		=	ospecial.Foneitem.Fisusing
		regdate		=	ospecial.Foneitem.Fregdate

		detailitemimage1	=	ospecial.Foneitem.Fdetailitemimage1
		detailitemimage2	=	ospecial.Foneitem.Fdetailitemimage2
		detailitemimage3	=	ospecial.Foneitem.Fdetailitemimage3
		detailitemimage4	=	ospecial.Foneitem.Fdetailitemimage4
		detailitemimage5	=	ospecial.Foneitem.Fdetailitemimage5
	set ospecial = nothing
End IF

if isusing="" then isusing="N"
if linkgubun="" then linkgubun="i"
if sortnum="" then sortnum="99"

%>
<style type="text/css">
.line1-1{border-top:2px solid gray;}
.line1-2{border-top:2px solid gray;}
.line2-1{border-top:2px solid gray;}
.line2-2{border-top:2px solid gray;}
.line3-1{border-top:2px solid gray;}
.line3-2{border-top:2px solid gray;}
</style>

<script anguage="javascript">
// 새상품 추가 팝업
//function findProd(){
//		var popwin;
//		popwin = window.open("/admin/Diary2009/pop_additemlist.asp", "popup_item", "width=900,height=600,scrollbars=yes,resizable=yes");
//		popwin.focus();
//}

function showimage(img){
	var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
}

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

function delimage(gubun)
{
	var aa = eval("document.frmreg."+gubun+"");
	aa.value = "";
	frmreg.submit();
}

function frmsubmit(){

	if (frmreg.mobiletext.value == ''){
		alert('모바일 문구를 입력해 주세요.');
		frmreg.mobiletext.focus();
		return;
	}
	
	if (frmreg.iid1.value != '' && !IsDouble(frmreg.iid1.value)){
		alert('상품코드는 숫자만 가능합니다.');
		frmreg.iid1.focus();
		return;
	}

	if (frmreg.iid2.value != '' && !IsDouble(frmreg.iid2.value)){
		alert('상품코드는 숫자만 가능합니다.');
		frmreg.iid2.focus();
		return;
	}
	
	if (frmreg.iid3.value != '' && !IsDouble(frmreg.iid3.value)){
		alert('상품코드는 숫자만 가능합니다.');
		frmreg.iid3.focus();
		return;
	}
	
	if (frmreg.iid4.value != '' && !IsDouble(frmreg.iid4.value)){
		alert('상품코드는 숫자만 가능합니다.');
		frmreg.iid4.focus();
		return;
	}
	
	if (frmreg.iid5.value != '' && !IsDouble(frmreg.iid5.value)){
		alert('상품코드는 숫자만 가능합니다.');
		frmreg.iid5.focus();
		return;
	}
	
	if (!IsDouble(frmreg.linkcode.value)){
		alert('링크코드가 없거나 숫자가 아닙니다.');
		frmreg.linkcode.focus();
		return;
	}

	if (!IsDouble(frmreg.sortnum.value)){
		alert('정렬순서가 없거나 숫자가 아닙니다.');
		frmreg.sortnum.focus();
		return;
	}
	frmreg.submit();
}
document.domain = "10x10.co.kr";
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmreg" method="post" action="/admin/Diaryspecial/DiaryspecialRegProc.asp">
<input type="hidden" name="mode" value="<%= Mode %>">
<input type="hidden" name="did" value="<%= idx %>">
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" height="25">
			<td colspan="2" align="center"><b>다이어리 신규등록</b></td>
		</tr>

		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> IDX</td>
			<td bgcolor="#FFFFFF" align="left"><%= idx %></td>
		</tr>

		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td class="line1-1" nowrap width="150"> PC메인이미지</td>
			<td class="line1-2" bgcolor="#FFFFFF" align="left">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('pcmainimages','pcmainimage','pcmain','2000','750','false');"/>
				(<b><font color="red">JPG,GIF</font></b>만가능)
					<input type="hidden" name="pcmainimage" value="<%= pcmainimage %>">
					<div align="right" id="pcmainimages"><% IF pcmainimage<>"" THEN %><img src="<%= pcmainimage %>" width="100" height="100" style="cursor:pointer" onclick="showimage('<%= pcmainimage %>');"><a href="javascript:delimage('pcmainimage');">[삭제]</a><% End IF %></div>
			</td>
		</tr>

		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> PC오버이미지</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('pcoverimages','pcoverimage','pcover','2000','750','false');"/>
				(<b><font color="red">JPG,GIF</font></b>만가능)
					<input type="hidden" name="pcoverimage" value="<%= pcoverimage %>">
					<div align="right" id="pcoverimages"><% IF pcoverimage<>"" THEN %><img src="<%= pcoverimage %>" width="100" height="100" style="cursor:pointer" onclick="showimage('<%= pcoverimage %>');"><a href="javascript:delimage('pcoverimage');">[삭제]</a><% End IF %></div>
			</td>
		</tr>
		
<!--
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> PC문구</td>
			<td bgcolor="#FFFFFF" align="left">
-->
				<input type="hidden" class="text" name="pctext" id="pctext" size="80" value="<%= pctext %>">
<!--
			</td>
		</tr>
-->
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td class="line2-1" nowrap width="150"> 모바일 메인이미지</td>
			<td class="line2-2" bgcolor="#FFFFFF" align="left">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('mobileimages','mobileimage','moimage','2000','750','false');"/>
				(<b><font color="red">JPG,GIF</font></b>만가능)
					<input type="hidden" name="mobileimage" value="<%= mobileimage %>">
					<div align="right" id="mobileimages"><% IF mobileimage<>"" THEN %><img src="<%= mobileimage %>" width="100" height="100" style="cursor:pointer" onclick="showimage('<%= mobileimage %>');"><a href="javascript:delimage('mobileimage');">[삭제]</a><% End IF %></div>
			</td>
		</tr>

		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> 모바일 문구</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" class="text" name="mobiletext" id="mobiletext" size="80" value="<%=mobiletext%>">
			</td>
		</tr>

		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td rowspan="5" class="line3-1" nowrap width="150"> 연관 상품코드</td>

			<td bgcolor="#FFFFFF"  class="line3-1" align="left">

				상품코드 : <input type="text" class="text" maxlength="10" name="iid1" id="iid1" value="<%= ItemID1 %>" >
				이 미 지 : <input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('detailitemimages1','detailitemimage1','dtimage1','2000','750','false');"/>
				(<b><font color="red">JPG,GIF</font></b>만가능)
					<input type="hidden" name="detailitemimage1" value="<%= detailitemimage1 %>">
					<div align="right" id="detailitemimages1"><% IF detailitemimage1<>"" THEN %><img src="<%= detailitemimage1 %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= detailitemimage1 %>');"><a href="javascript:delimage('detailitemimage1');">[삭제]</a><% End IF %></div>
				<br>
				<!--<input type="button" class="button" value="상품찾기" onClick="findProd();">-->
			</td>
		</tr>

		<tr>
			<td bgcolor="#FFFFFF" align="left">

				상품코드 : <input type="text" class="text" maxlength="10" name="iid2" id="iid2" value="<%= ItemID2 %>" >
				이 미 지 : <input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('detailitemimages2','detailitemimage2','dtimage2','2000','750','false');"/>
				(<b><font color="red">JPG,GIF</font></b>만가능)
					<input type="hidden" name="detailitemimage2" value="<%= detailitemimage2 %>">
					<div align="right" id="detailitemimages2"><% IF detailitemimage2<>"" THEN %><img src="<%= detailitemimage2 %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= detailitemimage2 %>');"><a href="javascript:delimage('detailitemimage2');">[삭제]</a><% End IF %></div>
				<br>
				<!--<input type="button" class="button" value="상품찾기" onClick="findProd();">-->
			</td>
		</tr>

		<tr>
			<td bgcolor="#FFFFFF" align="left">

				상품코드 : <input type="text" class="text" maxlength="10" name="iid3" id="iid3" value="<%= ItemID3 %>" >
				이 미 지 : <input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('detailitemimages3','detailitemimage3','dtimage3','2000','750','false');"/>
				(<b><font color="red">JPG,GIF</font></b>만가능)
					<input type="hidden" name="detailitemimage3" value="<%= detailitemimage3 %>">
					<div align="right" id="detailitemimages3"><% IF detailitemimage3<>"" THEN %><img src="<%= detailitemimage3 %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= detailitemimage3 %>');"><a href="javascript:delimage('detailitemimage3');">[삭제]</a><% End IF %></div>
				<br>
				<!--<input type="button" class="button" value="상품찾기" onClick="findProd();">-->
			</td>
		</tr>

		<tr>
			<td bgcolor="#FFFFFF" align="left">

				상품코드 : <input type="text" class="text" maxlength="10" name="iid4" id="iid4" value="<%= ItemID4 %>" >
				이 미 지 : <input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('detailitemimages4','detailitemimage4','dtimage4','2000','750','false');"/>
				(<b><font color="red">JPG,GIF</font></b>만가능)
					<input type="hidden" name="detailitemimage4" value="<%= detailitemimage4 %>">
					<div align="right" id="detailitemimages4"><% IF detailitemimage4<>"" THEN %><img src="<%= detailitemimage4 %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= detailitemimage4 %>');"><a href="javascript:delimage('detailitemimage4');">[삭제]</a><% End IF %></div>
				<br>
				<!--<input type="button" class="button" value="상품찾기" onClick="findProd();">-->
			</td>
		</tr>

		<tr>
			<td bgcolor="#FFFFFF" align="left">

				상품코드 : <input type="text" class="text" maxlength="10" name="iid5" id="iid5" value="<%= ItemID5 %>" >
				이 미 지 : <input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('detailitemimages5','detailitemimage5','dtimage5','2000','750','false');"/>
				(<b><font color="red">JPG,GIF</font></b>만가능)
					<input type="hidden" name="detailitemimage5" value="<%= detailitemimage5 %>">
					<div align="right" id="detailitemimages5"><% IF detailitemimage5<>"" THEN %><img src="<%= detailitemimage5 %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= detailitemimage5 %>');"><a href="javascript:delimage('detailitemimage5');">[삭제]</a><% End IF %></div>
				<br>
				<!--<input type="button" class="button" value="상품찾기" onClick="findProd();">-->
			</td>
		</tr>


		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td class="line3-1" nowrap> 메인이미지 링크구분</td>
			<td class="line3-1" bgcolor="#FFFFFF" align="left">
				<input type="radio" maxlength="10" name="linkgubun" value="i" <% IF linkgubun="i" THEN %>checked<% END IF %>>상품
				<input type="radio" maxlength="10" name="linkgubun" value="e" <% IF linkgubun="e" THEN %>checked<% END IF %>>이벤트
			</td>
		</tr>

		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> 메인이미지 링크코드</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" class="text" maxlength="10" name="linkcode" id="linkcode" value="<%= linkcode %>" >
			</td>
		</tr>

		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> 정렬순서</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" class="text" maxlength="3" name="sortnum" id="sortnum" value="<%= sortnum %>" >
			</td>
		</tr>

		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> 사용여부</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="radio" name="ius" value="Y" <% IF isUsing="Y" THEN %>checked<% END IF %>>사용
				<input type="radio" name="ius" value="N" <% IF isUsing="N" THEN %>checked<% END IF %> >사용안함
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center" bgcolor="<%=adminColor("green")%>"><br>
		<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="frmsubmit();" style="cursor:pointer">
		<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="frmreg.reset();" style="cursor:pointer">
		<img src="http://webadmin.10x10.co.kr/images/icon_new_registration.gif" border="0" onClick="location.href='/admin/diaryspecial/DiaryspecialReg.asp';" style="cursor:pointer">
	</td>
</tr>
</form>
</table>

<form name="imginputfrm" method="post" action="">
<input type="hidden" name="YearUse" value="2015">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- 리스트 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->