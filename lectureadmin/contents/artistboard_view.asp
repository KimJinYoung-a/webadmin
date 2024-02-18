<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/artistboard_cls.asp"-->
<%
Dim CBCont
Dim lecuserid, idx, userid, title, content, regdate, imgurl1, imgurl2
idx = requestCheckVar(request("idx"),10)
lecuserid = Session("ssBctId")
Set CBCont = new CArtistRoomBoard
	CBCont.Fidx = idx
	CBCont.FLecuserid = lecuserid
	CBCont.fnGetContent
	userid = CBCont.FUserid
	title = CBCont.FTitle
	content = CBCont.FContent
	imgurl1 = CBCont.FImgUrl1
	imgurl2 = CBCont.FImgUrl2
	regdate = CBCont.FRegdate
	
Set CBCont = nothing

%>
<script language="javascript">
<!--
	function jsDel(){
		if(confirm("삭제하시겠습니까?")){			
			document.frmDel.submit();
		}
	}
	
	function window.onload() {	 
		<% if (imgurl1 <> "") then %>
		  if (fimgurl1.width > 400) { fimgurl1.width=400; }
		<% end if %>
		<% if (imgurl2 <> "") then %>
		  if (fimgurl2.width> 400) { fimgurl2.width=400; }
		<% end if %>
	}
//-->
</script>
<!-- 보기 화면 시작 -->
<form name="frmDel" method="post" action="http://image.thefingers.co.kr/linkweb/artist/procboard.asp" enctype="multipart/form-data">
<input type="hidden" name="lecuserid" value="<%=lecuserid%>">
<input type="hidden" name="sUID" value="<%=userid%>">
<input type="hidden" name="sM" value="D">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="retUrl" value="http://webadmin.10x10.co.kr/lectureadmin/contents/artistboard_list.asp?menupos=979">
</form>
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2" >
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b> 상세 정보</b></td>
			<td height="26" align="right"><%=regdate%>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="15%" bgcolor="#DDDDFF">작성자 아이디</td>
	<td bgcolor="#FFFFFF"><%=userid%></td>
</tr>
<tr>
	<td align="center"  bgcolor="#DDDDFF">제목</td>
	<td bgcolor="#F8F8FF"><%=db2html(title)%></td>
</tr>
<tr>
	<td align="center"  bgcolor="#DDDDFF">내용</td>
	<td bgcolor="#FFFFFF">
		<%=nl2br(db2html(content))%><br><br>
		<%IF imgurl1 <> "" THEN %>	<img src="<%=imgurl1%>" id="fimgurl1"><%END IF%><Br><br>
		<%IF imgurl2 <> "" THEN %><img src="<%=imgurl2%>" id="fimgurl2"><%END IF%>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
	<%IF Cstr(userid) = Cstr(session("ssBctId")) THEN%>
		<img src="/images/icon_modify.jpg" onClick="self.location='artistboard_modi.asp?menupos=<%=menupos%>&idx=<%=idx%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="jsDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
	<%END IF%>	
	
		<img src="/images/icon_reply.gif" onClick="self.location='artistboard_write.asp?menupos=<%=menupos %>&idx=<%=idx%>'" style="cursor:pointer" align="absmiddle">
		<img src="/images/icon_list.gif" onClick="self.location='artistboard_list.asp?menupos=<%=menupos %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>

</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
