<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecture_stuffcls.asp"-->
<%
dim idx,mode
dim olec

idx = request("idx")
mode = request("mode")

if idx="" then idx=0
set olec = new CLectureStuffDetail
olec.GetLectureStuffDetail idx

%>
<script language="JavaScript">
<!--
function CheckForm(){
	if (document.lecform.itemid.value.length < 1){
		alert("제품번호를 등록해주세요");
		document.lecform.itemid.focus();
	}
	else if (document.lecform.lecturer.value.length < 1){
		alert("강사명을 등록해주세요");
		document.lecform.lecturer.focus();
	}
	else if (document.lecform.stuff.value.length < 1){
		alert("재료를 등록해주세요");
		document.lecform.stuff.focus();
	}
	else if (document.lecform.needtime.value.length < 1){
		alert("소요시간을 등록해주세요");
		document.lecform.needtime.focus();
	}
	else{
		document.lecform.submit();
	}
}

function calender_open(objectname) {
//       document.all.cal.style.display="";
//	   document.all.cal.style.left = event.offsetX;
//	   document.all.cal.style.top = event.offsetY + 200;
//	   document.lecform.objname.value = objectname;

//	   alert("X-좌표 : " + event.offsetX + "\n" + "Y-좌표 : " + event.offsetY);
}

//-->

function popLectureItemList(frm){
	var popwin = window.open('lecregitems.asp','lecitem','width=600,height=500,status=no,resizable=yes,scrollbars=yes');
	popwin.focus();
}
</script>
<form method="post" name="lecform" action="http://partner.10x10.co.kr/admin/lecture/dostuff.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="idx" value="<% = idx %>">
<input type="hidden" name="mode" value="<% = mode %>">
<table width="800" border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td >Idx</td>
	<td bgcolor="#FFFFFF"> <% = olec.Fidx %></td>
</tr>
<% if mode = "add" then %>
<tr bgcolor="#DDDDFF">
	<td >상품ID</td>
	<td bgcolor="#FFFFFF"><input type="text" name="itemid" value="0" size="9" maxlength="9">
	<!--
	<input type=button value="목록에서선택" onClick="popLectureItemList();">
	-->
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강사명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturer" value="" size="30" maxlength="32"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >재료구성</td>
	<td bgcolor="#FFFFFF"><input type="text" name="stuff" size="50" maxlength="128"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >소요시간</td>
	<td bgcolor="#FFFFFF"><input type="text" name="needtime" value="0" size="20" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >메인이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="mainimg" value="0" size="50" maxlength="12">(300*250)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >리스트이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="icon1" value="0" size="50" maxlength="12">(150*110)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상세이미지1</td>
	<td bgcolor="#FFFFFF"><input type="file" name="file1" value="0" size="50" maxlength="12">(600*400)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>이미지설명1</td>
	<td bgcolor="#FFFFFF"><textarea name="contents1" rows="10" cols="80"></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상세이미지2</td>
	<td bgcolor="#FFFFFF"><input type="file" name="file2" value="0" size="50" maxlength="12">(600*400)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>이미지설명2</td>
	<td bgcolor="#FFFFFF"><textarea name="contents2" rows="10" cols="80"></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상세이미지3</td>
	<td bgcolor="#FFFFFF"><input type="file" name="file3" value="0" size="50" maxlength="12">(600*400)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>이미지설명3</td>
	<td bgcolor="#FFFFFF"><textarea name="contents3" rows="10" cols="80"></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상세이미지4</td>
	<td bgcolor="#FFFFFF"><input type="file" name="file4" value="0" size="50" maxlength="12">(600*400)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>이미지설명4</td>
	<td bgcolor="#FFFFFF"><textarea name="contents4" rows="10" cols="80"></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF">
	&nbsp;&nbsp;&nbsp;
	<input type=radio name=isusing value=Y checked > 사용중(전시함)
	<input type=radio name=isusing value=N  > 사용안함(전시안함)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right" height="30"><input type="button" value="내용저장" onclick="CheckForm();return false;">&nbsp;&nbsp;&nbsp;</td>
</tr>

<% else %>
<tr bgcolor="#DDDDFF">
	<td >상품ID</td>
	<td bgcolor="#FFFFFF"><input type="text" name="itemid" size="9" maxlength="9" value="<% = olec.Fitemid %>">
	<input type=button value="목록에서선택" onClick="popLectureItemList();">
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강사명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturer" size="30" maxlength="32" value="<% = olec.Flecturer %>"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >재료구성</td>
	<td bgcolor="#FFFFFF"><input type="text" name="stuff" size="50" maxlength="128" value="<% = olec.Fstuff %>"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >소요시간</td>
	<td bgcolor="#FFFFFF"><input type="text" name="needtime" size="20" maxlength="12" value="<% = olec.Fneedtime %>"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >메인이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="mainimg" size="50" maxlength="12">(300*250)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >리스트이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="icon1" size="50" maxlength="12">(150*110)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상세이미지1</td>
	<td bgcolor="#FFFFFF"><input type="file" name="file1" size="50" maxlength="12">(600*400)<br>
		<input type="checkbox" name="dl_file1">삭제 (<% = olec.Ffile1 %>)
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>이미지설명1</td>
	<td bgcolor="#FFFFFF"><textarea name="contents1" rows="10" cols="80"><% = olec.Fcontents1 %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상세이미지2</td>
	<td bgcolor="#FFFFFF"><input type="file" name="file2" size="50" maxlength="12">(600*400)<br>
			<input type="checkbox" name="dl_file2">삭제 (<% = olec.Ffile2 %>)
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>이미지설명2</td>
	<td bgcolor="#FFFFFF"><textarea name="contents2" rows="10" cols="80"><% = olec.Fcontents2 %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상세이미지3</td>
	<td bgcolor="#FFFFFF"><input type="file" name="file3" size="50" maxlength="12">(600*400)<br>
			<input type="checkbox" name="dl_file3">삭제 (<% = olec.Ffile3 %>)
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>이미지설명3</td>
	<td bgcolor="#FFFFFF"><textarea name="contents3" rows="10" cols="80"><% = olec.Fcontents3 %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상세이미지4</td>
	<td bgcolor="#FFFFFF"><input type="file" name="file4" size="50" maxlength="12">(600*400)<br>
			<input type="checkbox" name="dl_file4">삭제 (<% = olec.Ffile4 %>)
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>이미지설명4</td>
	<td bgcolor="#FFFFFF"><textarea name="contents4" rows="10" cols="80"><% = olec.Fcontents4 %></textarea></td>
</tr>

<tr bgcolor="#DDDDFF">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF">
	&nbsp;&nbsp;&nbsp;
	<% if olec.FIsUsing ="Y" then %>
	<input type=radio name=isusing value=Y checked > 사용중(전시함)
	<input type=radio name=isusing value=N  > 사용안함(전시안함)
	<% else %>
	<input type=radio name=isusing value=Y  > 사용중(전시함)
	<input type=radio name=isusing value=N checked > 사용안함(전시안함)
	<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right" height="30"><input type="button" value="내용저장" onclick="CheckForm();return false;">&nbsp;&nbsp;&nbsp;</td>
</tr>
<% end if %>
</table>
</form>
<%
set olec = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->