<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station 에디터 수정  
' Hieditor : 2008.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incsessionadmin.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<%
dim YearUse ,editor_no ,editor_name , isusing ,  comment_isusing
dim listimgName, list2imgName , barnerimgName, barner2imgName, main1imgName , main2imgName , image_main_link
dim main3imgName, main4imgName, main5imgName, list2015imgName
	YearUse = 2009
	editor_no = requestCheckVar(getNumeric(request("editor_no")),10)
	
dim oip, i
set oip = new ceditor_list
	oip.frecteditor_no = editor_no
	if editor_no <> "" then
		oip.feditor_list()
		editor_no = ReplaceBracket(oip.FItemList(0).feditor_no)
		editor_name = oip.FItemList(0).feditor_name
		isusing = oip.FItemList(0).fisusing
		comment_isusing = oip.FItemList(0).fcomment_isusing
		listimgName = oip.FItemList(0).fimage_list
		list2imgName = oip.FItemList(0).fimage_list2
		barnerimgName = oip.FItemList(0).fimage_barner
		barner2imgName = oip.FItemList(0).fimage_barner2
		main1imgName = oip.FItemList(0).fimage_main
		main2imgName = oip.FItemList(0).fimage_main2
		main3imgName = oip.FItemList(0).fimage_main3
		main4imgName = oip.FItemList(0).fimage_main4
		main5imgName = oip.FItemList(0).fimage_main5				
		image_main_link = oip.FItemList(0).fimage_main_link
		list2015imgName = oip.FItemList(0).fimage_list2015
	end if
%>
<script type='text/javascript'>

	//''이미지 저장
	function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,fheight,thumb){
	
		window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
		document.imginputfrm.divName.value=divnm;
		document.imginputfrm.inputname.value=iptNm;
		document.imginputfrm.ImagePath.value = vPath;
		document.imginputfrm.maxFileSize.value = Fsize;
		document.imginputfrm.maxFileWidth.value = Fwidth;
		document.imginputfrm.maxFileheight.value = fheight;	
		document.imginputfrm.makeThumbYn.value = thumb;
		document.imginputfrm.orgImgName.value = document.getElementById(iptNm).value;
		document.imginputfrm.target='imginput';
		document.imginputfrm.action='editor_img_input.asp';
		document.imginputfrm.submit();
	}

	document.domain = "10x10.co.kr";

	//''이벤트저장
	function editor_reg(mode){
				
		if (mode == 'add'){
			frm.mode.value='add';
		}else if(mode == 'edit'){
			frm.mode.value='edit';
		}
		
		frm.submit();		
	}
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" action="/admin/culturestation/editor_edit_process.asp" style="margin:0px;">
<input type="hidden" name="mode" >
<input type="hidden" name="editor_no" value="<%=editor_no%>">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>editor명</b><br></td>
		<td><input type="text" name="editor_name" size="50" value="<%= editor_name %>">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>사용여부</b><br></td>
		<td><select name="isusing">
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>기능</b><br></td>
		<td><input type="checkbox" name="comment_isusing" value="ON" <% if comment_isusing = "ON" then response.write " checked" %>>코맨트사용
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>기본 이미지</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('listdiv','listimgName','list','200','50','50','false');"  class="button"> 
			(<b><font color="red">50x50</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="listimgName" id="listimgName" value="<%= listimgName %>">
			<div align="right" id="listdiv">
				<% if listimgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/list/<%= oip.FItemList(i).fimage_list %>" width="50" height="50">
				<% end if %>	
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>기본 #2013</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('list2div','list2imgName','list2','200','246','221','false');"  class="button"> 
			(<b><font color="red">246x221</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="list2imgName" id="list2imgName" value="<%= list2imgName %>">
			<div align="right" id="list2div">
				<% if list2imgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/list2/<%= oip.FItemList(i).fimage_list2 %>" width="50" height="50">
				<% end if %>	
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>기본 #2015</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('list2015div','list2015imgName','list2015','200','296','100','false');"  class="button"> 
			(<b><font color="red">296x100</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="list2015imgName" id="list2015imgName" value="<%= list2015imgName %>">
			<div align="right" id="list2015div">
				<% if list2015imgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/list2015/<%= oip.FItemList(i).fimage_list2015 %>" width="50" height="50">
				<% end if %>	
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>배너 이미지</b><br></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('barnerdiv','barnerimgName','barner','200','190','78','false');"  class="button"> 
			(<b><font color="red">190x78</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="barnerimgName" id="barnerimgName" value="<%= barnerimgName %>">
			<div align="right" id="barnerdiv">
				<% if barnerimgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/barner/<%= oip.FItemList(i).fimage_barner %>" width="50" height="50">
				<% end if %>				
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>배너 #2013</b><br></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('barner2div','barner2imgName','barner2','200','192','80','false');"  class="button"> 
			(<b><font color="red">192x80</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="barner2imgName" id="barner2imgName" value="<%= barner2imgName %>">
			<div align="right" id="barner2div">
				<% if barner2imgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/barner2/<%= oip.FItemList(i).fimage_barner2 %>" width="50" height="50">
				<% end if %>
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>메인 이미지1</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main1div','main1imgName','main1','600','745','5000','false');"  class="button"> 
			(<b><font color="red">가로745</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="main1imgName" id="main1imgName" value="<%= main1imgName %>">
			<div align="right" id="main1div">
				<% if main1imgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/main1/<%= main1imgName %>" width="50" height="50">
				<% end if %>				
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>메인 이미지2</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main2div','main2imgName','main2','600','745','5000','false');"  class="button"> 
			(<b><font color="red">가로745</font></b><b><font color="red"></font></b>) 없으면 등록하지 마세요.
			<input type="hidden" name="main2imgName" id="main2imgName" value="<%= main2imgName %>">
			<div align="right" id="main2div">
				<% if main2imgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/main2/<%= main2imgName %>" width="50" height="50">
				<% end if %>				
			</div>
		</td>
	</tr>
		
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>메인 이미지3</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main3div','main3imgName','main3','600','745','5000','false');"  class="button"> 
			(<b><font color="red">가로745</font></b><b><font color="red"></font></b>) 없으면 등록하지 마세요.
			<input type="hidden" name="main3imgName" id="main3imgName" value="<%= main3imgName %>">
			<div align="right" id="main3div">
				<% if main3imgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/main3/<%= main3imgName %>" width="50" height="50">
				<% end if %>				
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>메인 이미지4</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main4div','main4imgName','main4','600','745','5000','false');"  class="button"> 
			(<b><font color="red">가로745</font></b><b><font color="red"></font></b>) 없으면 등록하지 마세요.
			<input type="hidden" name="main4imgName" id="main4imgName" value="<%= main4imgName %>">
			<div align="right" id="main4div">
				<% if main4imgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/main4/<%= main4imgName %>" width="50" height="50">
				<% end if %>				
			</div>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>메인 이미지5</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main5div','main5imgName','main5','600','745','5000','false');"  class="button"> 
			(<b><font color="red">가로745</font></b><b><font color="red"></font></b>) 없으면 등록하지 마세요.
			<input type="hidden" name="main5imgName" id="main5imgName" value="<%= main5imgName %>">
			<div align="right" id="main5div">
				<% if main5imgName <> "" then %>
					<img src="<%=webImgUrl%>/culturestation/editor/<%= yearUse %>/main5/<%= main5imgName %>" width="50" height="50">
				<% end if %>				
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2"><b>메인이미지 이미지맵 & 링크 코드</b> <font color="red"> map이름 절대 수정금지 !!</font></td>
	</tr>
	
	<% 
	'//수정
	if editor_no <> "" then 
	%>
	
		
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
			<textarea rows="15" cols="100" name="image_main_link"><%= image_main_link %></textarea>
		</td>
	</tr>	
	<% 
	'//저장
	else 
	%>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
			<textarea rows="15" cols="100" name="image_main_link"><map name="ImgMap1"></map><map name="ImgMap2"></map><map name="ImgMap3"></map><map name="ImgMap4"></map><map name="ImgMap5"></map></textarea>
		</td>
	</tr>
	<% end if %>
	
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
			<% 
			'//수정
			if editor_no <> "" then 
			%>
				<input type="button" value="editor수정" onclick="editor_reg('edit');" class="button">
			<% 
			'//신규
			else 
			%>
				<input type="button" value="editor신규저장" onclick="editor_reg('add');" class="button">
			<% end if %>
		</td>
	</tr>
</table>
</form>
<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="YearUse" value="<%= YearUse %>">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="maxFileheight" value="">	
	<input type="hidden" name="makeThumbYn" value="">
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->

<%
	set oip = nothing
%>