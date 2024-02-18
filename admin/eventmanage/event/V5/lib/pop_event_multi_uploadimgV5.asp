<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/pop_event_uploadimg.asp
' Description :  이벤트 이미지 등록
' History : 2007.02.22 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
Dim sFolder, sImg1, sImg2, sImg3, sImg4, sImg5
dim sImg6, sImg7, sImg8, sImg9, sImg10
dim sName, sSpan, slen, arrImg, vYear, menuidx
Dim sOpt, sImgName1, sImgName2, sImgName3, sImgName4, sImgName5
dim sImgName6, sImgName7, sImgName8, sImgName9, sImgName10
dim idx1, idx2, idx3, idx4, idx5, idx6, idx7, idx8, idx9, idx10
dim ArrcMultiContentsSwife, ArrcMultiContentsSwife2, ix
dim cEvtCont, cEvtCont2, uploadok, eCode

sFolder = Request.Querystring("sF") 
menuidx = Request.Querystring("menuidx")
sName = Request.Querystring("sName")
sSpan = Request.Querystring("sSpan")
sOpt = Request.Querystring("sOpt")
vYear = Request("yr")
uploadok = Request("uploadok")
eCode = requestCheckvar(request("eC"),16)
if uploadok="Y" then
%>
<script>
	opener.document.location.replace("/admin/eventmanage/event/v5/template/slide/pop_culture_themeslide.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>");
	self.close();
</script>
<%
end if

IF menuidx <> "" THEN
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    cEvtCont.FRectIDX = menuidx	'멀티 컨텐츠 메뉴 idx
    cEvtCont.FRectDevice = "M"	'멀티 컨텐츠 디바이스 구분
    ArrcMultiContentsSwife=cEvtCont.fnGetMultiContentsSwifeList
    set cEvtCont = nothing
end if
If isArray(ArrcMultiContentsSwife) Then
	For ix = 0 To UBound(ArrcMultiContentsSwife,2)
		if ix=0 then
			idx1 = ArrcMultiContentsSwife(0,ix)
			sImg1 = ArrcMultiContentsSwife(1,ix)
		end if
		if ix=1 then
			idx2 = ArrcMultiContentsSwife(0,ix)
			sImg2 = ArrcMultiContentsSwife(1,ix)
		end if
		if ix=2 then
			idx3 = ArrcMultiContentsSwife(0,ix)
			sImg3 = ArrcMultiContentsSwife(1,ix)
		end if
		if ix=3 then
			idx4 = ArrcMultiContentsSwife(0,ix)
			sImg4 = ArrcMultiContentsSwife(1,ix)
		end if
		if ix=4 then
			idx5 = ArrcMultiContentsSwife(0,ix)
			sImg5 = ArrcMultiContentsSwife(1,ix)
		end if
		if ix=5 then
			idx6 = ArrcMultiContentsSwife(0,ix)
			sImg6 = ArrcMultiContentsSwife(1,ix)
		end if
		if ix=6 then
			idx7 = ArrcMultiContentsSwife(0,ix)
			sImg7 = ArrcMultiContentsSwife(1,ix)
		end if
		if ix=7 then
			idx8 = ArrcMultiContentsSwife(0,ix)
			sImg8 = ArrcMultiContentsSwife(1,ix)
		end if
		if ix=8 then
			idx9 = ArrcMultiContentsSwife(0,ix)
			sImg9 = ArrcMultiContentsSwife(1,ix)
		end if
		if ix=9 then
			idx10 = ArrcMultiContentsSwife(0,ix)
			sImg10 = ArrcMultiContentsSwife(1,ix)
		end if
	Next
end if

IF sImg1 <> "" THEN
	arrImg = split(sImg1,"/")
	slen = ubound(arrImg)
	sImgName1 = arrImg(slen)
END IF
IF sImg2 <> "" THEN
	arrImg = split(sImg2,"/")
	slen = ubound(arrImg)
	sImgName2 = arrImg(slen)
END IF
IF sImg3 <> "" THEN
	arrImg = split(sImg3,"/")
	slen = ubound(arrImg)
	sImgName3 = arrImg(slen)
END IF
IF sImg4 <> "" THEN
	arrImg = split(sImg4,"/")
	slen = ubound(arrImg)
	sImgName4 = arrImg(slen)
END IF
IF sImg5 <> "" THEN
	arrImg = split(sImg5,"/")
	slen = ubound(arrImg)
	sImgName5 = arrImg(slen)
END IF
IF sImg6 <> "" THEN
	arrImg = split(sImg6,"/")
	slen = ubound(arrImg)
	sImgName6 = arrImg(slen)
END IF
IF sImg7 <> "" THEN
	arrImg = split(sImg7,"/")
	slen = ubound(arrImg)
	sImgName7 = arrImg(slen)
END IF
IF sImg8 <> "" THEN
	arrImg = split(sImg8,"/")
	slen = ubound(arrImg)
	sImgName8 = arrImg(slen)
END IF
IF sImg9 <> "" THEN
	arrImg = split(sImg9,"/")
	slen = ubound(arrImg)
	sImgName9 = arrImg(slen)
END IF
IF sImg10 <> "" THEN
	arrImg = split(sImg10,"/")
	slen = ubound(arrImg)
	sImgName10 = arrImg(slen)
END IF

vYear = Request("yr")
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");			
			return false;
		}
		document.all.dvLoad.style.display = "";
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이미지 업로드 처리</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/event_admin/V5/event_upload_multi.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="idx1" value="<%=idx1%>">
<input type="hidden" name="sImg1" value="<%=sImg1%>">
<input type="hidden" name="idx2" value="<%=idx2%>">
<input type="hidden" name="sImg2" value="<%=sImg2%>">
<input type="hidden" name="idx3" value="<%=idx3%>">
<input type="hidden" name="sImg3" value="<%=sImg3%>">
<input type="hidden" name="idx4" value="<%=idx4%>">
<input type="hidden" name="sImg4" value="<%=sImg4%>">
<input type="hidden" name="idx5" value="<%=idx5%>">
<input type="hidden" name="sImg5" value="<%=sImg5%>">
<input type="hidden" name="idx6" value="<%=idx6%>">
<input type="hidden" name="sImg6" value="<%=sImg6%>">
<input type="hidden" name="idx7" value="<%=idx7%>">
<input type="hidden" name="sImg7" value="<%=sImg7%>">
<input type="hidden" name="idx8" value="<%=idx8%>">
<input type="hidden" name="sImg8" value="<%=sImg8%>">
<input type="hidden" name="idx9" value="<%=idx9%>">
<input type="hidden" name="sImg9" value="<%=sImg9%>">
<input type="hidden" name="idx10" value="<%=idx10%>">
<input type="hidden" name="sImg10" value="<%=sImg10%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<input type="hidden" name="yr" value="<%=vYear%>">
<input type="hidden" name="sOpt" value="<%=sOpt%>">
<input type="hidden" name="menuidx" value="<%=menuidx%>">
<input type="hidden" name="eCode" value="<%=eCode%>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지1</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg1"></td>
	</tr>
	<%IF sImg1 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName1%></td>
	</tr>
	<%END IF%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지2</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg2"></td>
	</tr>
	<%IF sImg2 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName2%></td>
	</tr>
	<%END IF%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지3</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg3"></td>
	</tr>
	<%IF sImg3 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName3%></td>
	</tr>
	<%END IF%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지4</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg4"></td>
	</tr>
	<%IF sImg4 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName4%></td>
	</tr>
	<%END IF%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지5</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg5"></td>
	</tr>
	<%IF sImg5 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName5%></td>
	</tr>
	<%END IF%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지6</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg6"></td>
	</tr>
	<%IF sImg6 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName6%></td>
	</tr>
	<%END IF%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지7</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg7"></td>
	</tr>
	<%IF sImg7 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName7%></td>
	</tr>
	<%END IF%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지8</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg8"></td>
	</tr>
	<%IF sImg8 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName8%></td>
	</tr>
	<%END IF%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지9</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg9"></td>
	</tr>
	<%IF sImg9 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName9%></td>
	</tr>
	<%END IF%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지10</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg10"></td>
	</tr>
	<%IF sImg10 <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName10%></td>
	</tr>
	<%END IF%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			+ 최대 파일사이즈 1MB(1,024KB) 이하만,<br>
			+ gif,jpg,png 타입의 파일만 등록가능
		</td>
	</tr>
</form>	
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<div id="dvLoad" style="display:none;top:50px;left:20;position:absolute;background-color:gray;">
	<table border="0" class="a" cellpadding="5" cellspacing="5">
		<tr>
			<td> <font color="#FFFFFF">업로드 처리중입니다. 잠시만 기다려주세요~~</font></td>
		</tr>
	</table>
</div>