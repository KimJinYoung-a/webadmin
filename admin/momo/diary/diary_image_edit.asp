<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 다이어리
' Hieditor : 2009.12.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim idx,diary_date,title,contents,mainimage1 , i , diarytype
dim mainimage2,mainimage3,isusing,regdate,diary_order
	idx = request("idx")
	
	if idx = "" then 
		response.write "<script>alert('idx 값이 없습니다.'); self.close();</script>"
		dbget.close() : response.end
	end if
	
dim oMainContents
	set oMainContents = new cdiary_list
	oMainContents.FRectIdx = idx
	oMainContents.fdiarycontents_oneitem

if oMainContents.ftotalcount > 0 then	
	diary_date = oMainContents.FOneItem.fdiary_date
	title = oMainContents.FOneItem.ftitle
	contents = oMainContents.FOneItem.fcontents
	mainimage1 = oMainContents.FOneItem.fmainimage1
	mainimage2 = oMainContents.FOneItem.fmainimage2
	mainimage3 = oMainContents.FOneItem.fmainimage3
	isusing = oMainContents.FOneItem.fisusing	
	diary_order = oMainContents.FOneItem.fdiary_order
	diarytype = oMainContents.FOneItem.fdiarytype
end if
%>
	
<script language='javascript'>

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
	function SaveMainContents(){
		frmcontents.submit();				
	}

</script>

			
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="center">
			
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="gray">
<form name="frmcontents" method="post" action="/admin/momo/diary/diary_process.asp">
<input type="hidden" name="mode" value="image">			
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">Idx :</td>
	    <td>
	        <%= idx %><input type="hidden" name="idx" value="<%= idx %>">
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>상세이미지1</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('mainimage1div','mainimage1','mainimage1','2000','800','true');"/>		
			<input type="hidden" name="mainimage1" value="<%= mainimage1 %>">
			<div align="right" id="mainimage1div"><% IF mainimage1<>"" THEN %><img src="<%=webImgUrl%>/momo/diary/mainimage1/<%= mainimage1 %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>상세이미지2</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('mainimage2div','mainimage2','mainimage2','2000','800','true');"/>		
			<input type="hidden" name="mainimage2" value="<%= mainimage2 %>">
			<div align="right" id="mainimage2div"><% IF mainimage2<>"" THEN %><img src="<%=webImgUrl%>/momo/diary/mainimage2/<%= mainimage2 %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
		</td>
	</tr>		
	<tr align="center" bgcolor="#FFFFFF">
		<td>상세이미지3</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('mainimage3div','mainimage3','mainimage3','2000','800','true');"/>		
			<input type="hidden" name="mainimage3" value="<%= mainimage3 %>">
			<div align="right" id="mainimage3div"><% IF mainimage3<>"" THEN %><img src="<%=webImgUrl%>/momo/diary/mainimage3/<%= mainimage3 %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
	    <td  align="center" colspan=2>
	    	<input type="button" value=" 저 장 " onClick="SaveMainContents();" class="button">
	    </td>
	</tr>	
</form>
</table>

<%
	set oMainContents = nothing
%>
<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->