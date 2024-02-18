<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<%

Function getMainPosNoArr()
	dim strSQL,tmpArr
	strSQL =" SELECT PosNo " &_
			" FROM db_diary2010.dbo.tbl_diaryMain " &_
			" GROUP BY PosNo ORDER BY Posno "

	rsget.open strSQL,dbget,2
	IF not rsget.Eof Then
		tmpArr = rsget.getRows()
	End IF
	rsget.Close
	getMainPosNoArr = tmpArr
End Function

Function getMainArr(Pn)
	dim strSQL,tmpArr

	StrSQL =" SELECT TOP 1 id,PosNo,Img,Url  "&_
			" FROM db_diary2010.dbo.tbl_diaryMain "&_
			" WHERE PosNo="& Pn &" ORDER BY id desc "

	rsget.open strSQL,dbget,2
	IF not rsget.Eof Then
		tmpArr = rsget.getRows()
	End IF
	rsget.Close
	getMainArr = tmpArr
End Function

Function getImgUrl(Img)
	IF img<>"" Then
		IF application("Svr_Info")="Dev" THEN
			getImgUrl = "http://testimgstatic.10x10.co.kr/diary/2009/main/"& Img
		ELSE
			getImgUrl = "http://webimage.10x10.co.kr/diary_collection/2009/main/"& Img
		End IF
	End IF
End Function


dim ArrPosNo ,ArrList ,tmpArr
dim intLp , intLp2

ArrPosNo = getMainPosNoArr


Redim ArrList(3,23) '//23행 4열 배열 생성후 0행 사용안함

IF isArray(ArrPosNo) Then
	FOR intLp = 0 To Ubound(ArrPosNo,2)
		tmpArr = getMainArr(ArrPosNo(0,intLp))
		IF isArray(tmpArr) Then
			FOR intLp2 = 0 To Ubound(tmpArr,1)
				ArrList(intLp2,ArrPosNo(0,intLp)) = tmpArr(intLp2,0)
			NEXT
		End IF
	NEXT
End IF


%>
<script language="javascript">


function jsInsert(pn){

	var popmitemreg = window.open('/admin/diary2009/lib/pop_DiaryMainItem_reg.asp?pn='+ pn,'popreg','width=620,height=800,resizable=yes,scrollbars=yes')
	popmitemreg.focus();
}

function jsEdit(pn){
	if (pn==13){
		document.location.href="/admin/diary2009/lib/pop_DiaryMainFlash_reg.asp?pn="+pn;
		//var popmitemreg = window.open('/admin/diary2009/lib/pop_DiaryMainFlash_reg.asp?pn='+ pn,'popreg','width=620,height=800,resizable=yes,scrollbars=yes')
	}else{
		var popmitemreg = window.open('/admin/diary2009/lib/pop_DiaryMainItem_reg.asp?md=edit&pn='+ pn,'popreg','width=620,height=800,resizable=yes,scrollbars=yes')
		popmitemreg.focus();
	}

}
function preview(){
	var poppre = window.open('<%= wwwUrl %>/chtml/diary/pre_make_diarymain.asp','poppre','');
	poppre.focus();
}

function make(){
	if(confirm('실서버에 적용합니다.')){
		var popreal = window.open('<%= wwwUrl %>/chtml/diary/make_diarymain.asp','popreal','width=620,height=800,resizable=yes,scrollbars=yes');
		popreal.focus();
	}
}

function showimage(img){
	var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
}

window.onload = function(){
	window.resizeTo(550,700);
}

document.domain = "10x10.co.kr";

</script>
<table width="450" height="400" border="0" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="2" rowspan="2">
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv1" style="cursor:pointer"><% IF ArrList(2,1)<>"" Then %><img src="<%= getImgUrl(ArrList(2,1)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,1)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('1');" value="등록">
				<input type="button" class="button" onClick="jsEdit('1');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv2" style="cursor:pointer"><% IF ArrList(2,2)<>"" Then %><img src="<%= getImgUrl(ArrList(2,2)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,2)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('2');" value="등록"> <input type="button" class="button" onClick="jsEdit('2');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv3" style="cursor:pointer;"><% IF ArrList(2,3)<>"" Then %><img src="<%= getImgUrl(ArrList(2,3)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,3)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('3');" value="등록"><input type="button" class="button" onClick="jsEdit('3');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv4" style="cursor:pointer"><% IF ArrList(2,4)<>"" Then %><img src="<%= getImgUrl(ArrList(2,4)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,4)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('4');" value="등록"><input type="button" class="button" onClick="jsEdit('4');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv5" style="cursor:pointer"><% IF ArrList(2,5)<>"" Then %><img src="<%= getImgUrl(ArrList(2,5)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,5)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('5');" value="등록"><input type="button" class="button" onClick="jsEdit('5');" value="수정">
		</td></tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv6" style="cursor:pointer"><% IF ArrList(2,6)<>"" Then %><img src="<%= getImgUrl(ArrList(2,6)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,6)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('6');" value="등록"><input type="button" class="button" onClick="jsEdit('6');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv7" style="cursor:pointer"><% IF ArrList(2,7)<>"" Then %><img src="<%= getImgUrl(ArrList(2,7)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,7)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('7');" value="등록"><input type="button" class="button" onClick="jsEdit('7');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv8" style="cursor:pointer"><% IF ArrList(2,8)<>"" Then %><img src="<%= getImgUrl(ArrList(2,8)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,8)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('8');" value="등록"><input type="button" class="button" onClick="jsEdit('8');" value="수정">
		</td></tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv9" style="cursor:pointer"><% IF ArrList(2,9)<>"" Then %><img src="<%= getImgUrl(ArrList(2,9)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,9)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('9');" value="등록"><input type="button" class="button" onClick="jsEdit('9');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv10" style="cursor:pointer"><% IF ArrList(2,10)<>"" Then %><img src="<%= getImgUrl(ArrList(2,10)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,10)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('10');" value="등록"><input type="button" class="button" onClick="jsEdit('10');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv11" style="cursor:pointer"><% IF ArrList(2,11)<>"" Then %><img src="<%= getImgUrl(ArrList(2,11)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,11)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('11');" value="등록"><input type="button" class="button" onClick="jsEdit('11');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv12" style="cursor:pointer"><% IF ArrList(2,12)<>"" Then %><img src="<%= getImgUrl(ArrList(2,12)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,12)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('12');" value="등록"><input type="button" class="button" onClick="jsEdit('12');" value="수정">
		</td></tr>
		</table>
	</td>
	<td colspan="2" rowspan="2">
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv13" style="cursor:pointer"><% IF ArrList(2,13)<>"" Then %><img src="<%= getImgUrl(ArrList(2,13)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,13)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('13');" value="등록"><input type="button" class="button" onClick="jsEdit('13');" value="관리">
		</td></tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv14" style="cursor:pointer"><% IF ArrList(2,14)<>"" Then %><img src="<%= getImgUrl(ArrList(2,14)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,14)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('14');" value="등록"><input type="button" class="button" onClick="jsEdit('14');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv15" style="cursor:pointer"><% IF ArrList(2,15)<>"" Then %><img src="<%= getImgUrl(ArrList(2,15)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,15)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('15');" value="등록"><input type="button" class="button" onClick="jsEdit('15');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv16" style="cursor:pointer"><% IF ArrList(2,16)<>"" Then %><img src="<%= getImgUrl(ArrList(2,16)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,16)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('16');" value="등록"><input type="button" class="button" onClick="jsEdit('16');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv17" style="cursor:pointer"><% IF ArrList(2,17)<>"" Then %><img src="<%= getImgUrl(ArrList(2,17)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,17)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('17');" value="등록"><input type="button" class="button" onClick="jsEdit('17');" value="수정">
		</td></tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv18" style="cursor:pointer"><% IF ArrList(2,18)<>"" Then %><img src="<%= getImgUrl(ArrList(2,18)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,18)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('18');" value="등록"><input type="button" class="button" onClick="jsEdit('18');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv19" style="cursor:pointer"><% IF ArrList(2,19)<>"" Then %><img src="<%= getImgUrl(ArrList(2,19)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,19)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('19');" value="등록"><input type="button" class="button" onClick="jsEdit('19');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv20" style="cursor:pointer"><% IF ArrList(2,20)<>"" Then %><img src="<%= getImgUrl(ArrList(2,20)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,20)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('20');" value="등록"><input type="button" class="button" onClick="jsEdit('20');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv21" style="cursor:pointer"><% IF ArrList(2,21)<>"" Then %><img src="<%= getImgUrl(ArrList(2,21)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,21)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('21');" value="등록"><input type="button" class="button" onClick="jsEdit('21');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv22" style="cursor:pointer"><% IF ArrList(2,22)<>"" Then %><img src="<%= getImgUrl(ArrList(2,22)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,22)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('22');" value="등록"><input type="button" class="button" onClick="jsEdit('22');" value="수정">
		</td></tr>
		</table>
	</td>
	<td>
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		<tr><td><div id="dv23" style="cursor:pointer"><% IF ArrList(2,23)<>"" Then %><img src="<%= getImgUrl(ArrList(2,23)) %>" border="0" width="50" height="50" onClick="showimage('<%= getImgUrl(ArrList(2,23)) %>')"><% End IF %></div></td></tr>
		<tr><td height="30" align="center">
				<input type="button" class="button" onClick="jsInsert('23');" value="등록"><input type="button" class="button" onClick="jsEdit('23');" value="수정">
		</td></tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
		<input type="button" class="button" value="미리보기" onclick="preview();">
		<input type="button" class="button" value="적용" onclick="make();">
	</td>
</tr>
</table>
<form name="regfrm" target="regframe" action="">
<input type="hidden" name="" value="">
<input type="hidden" name="" value=""
</form>
<iframe name="regframe" src="" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->