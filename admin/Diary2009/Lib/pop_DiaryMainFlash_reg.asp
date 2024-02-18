<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->

<%

dim intPosNo , intID , strIMG , strURL , strMode

dim ArrList,intLp
dim StrSQL



intPosNo = request("pn")
strMode = request("md")
IF intPosNo="" Then
	Alert_close("오류")
End IF


Function getImgUrl(Img)
	IF img<>"" Then
		IF application("Svr_Info")="Dev" THEN
			getImgUrl = "http://testimgstatic.10x10.co.kr/diary/2009/main/"& Img
		ELSE
			getImgUrl = "http://webimage.10x10.co.kr/diary_collection/2009/main/"& Img
		End IF
	End IF
End Function

IF intPosNo<>""  Then
	StrSQL =" SELECT id,PosNo,Img,Url,isUsing  "&_
			" FROM db_diary2010.dbo.tbl_diaryMain "&_
			" WHERE isUsing='Y' and PosNo="& intPosNo &" ORDER BY id desc "
	rsget.open StrSQL,dbget,2

	IF not rsget.Eof then
		ArrList = rsget.getRows()
	End IF

	rsget.close
End IF
%>
<script language="javascript">

//전체선택

function jsChkAll(blnChk){
	    var frm, blnChk;
		frm = document.rfrm;

 		for (var i=0;i<frm.elements.length;i++){
			//check optioon
			var e = frm.elements[i];

			//check itemEA
			if ((e.type=="checkbox")) {
				e.checked = blnChk ;
				AnCheckClick(e);
		}
	}
}

//삭제
function jsDel(sType, iValue){
		var frm;
		var sValue;
		frm = document.rfrm;
		sValue = "";

		if (sType ==0) {
			//if(!frm.chkb) return;

			if (frm.chkb.length > 1){
				for (var i=0;i<frm.chkb.length;i++){
					if(frm.chkb[i].checked){
					   	if (sValue==""){
							sValue = frm.chkb[i].value;
					   	}else{
							sValue =sValue+","+frm.chkb[i].value;
					   	}
					}
				}
			}else{
				if(frm.chkb.checked){
					sValue = frm.chkb.value;
				}
			}

			if (sValue == "") {
				alert('선택 상품이 없습니다.');
				return;
			}
			document.frmmn.id.value = sValue;
		}else{
			document.frmmn.id.value = iValue;
		}
		alert(sValue);
		if(confirm("선택하신 상품을 삭제하시겠습니까?")){
			document.frmmn.submit();
		}
}
function make(){
	if(confirm('실서버에 적용합니다.')){
		var popreal = window.open('<%= wwwUrl %>/chtml/diary/make_diarymain_flash.asp','popins','width=620,height=800,resizable=yes,scrollbars=yes');
		popreal.focus();
	}
}
//window.onload = function(){
//	window.resizeTo(600,500);
//}

</script>

<table  width="500" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmmn" method="post" action="<%= uploadImgUrl %>/linkWeb/diary/DiaryMainReg_Proc.asp" enctype="multipart/form-data">
<input type="hidden" name="md" value="del">
<input type="hidden" name="id" value="">
<input type="hidden" name="pn" value="<%= intPosNo %>">
<tr bgcolor="#FFFFFF">
	<td align="right">
		<input type="button" class="button" value="적용" onClick="make();">
		<input type="button" class="button" value="삭제" onClick="jsDel(0,'')">
	</td>
</tr>
</form>
</table>
<table  width="500" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="rfrm" method="post" action="<%= uploadImgUrl %>/linkWeb/diary/DiaryMainReg_Proc.asp" target="regframe" enctype="multipart/form-data" >
<input type="hidden" name="md" value="<%= strMode %>">
<input type="hidden" name="id" value="<%= intID%>">
<input type="hidden" name="pn" value="<%= intPosNo %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width="20" align="center"><input type="checkbox" name="chkAll" onClick="jsChkAll(this.checked);"></td>
	<td width="50" align="center">이미지</td>
	<td width="200">URL</td>
	<td width="50" align="center">사용여부</td>
</tr>
<% IF isArray(ArrList) Then %>
<% FOR intLp=0 To Ubound(ArrList,2) %>
<tr bgcolor="#FFFFFF">
	<td><input type="checkbox" name="chkb" onClick="AnCheckClick(this);" value="<%= ArrList(0,intLp) %>"></td>
	<td align="center"><img src="<%= getImgUrl(ArrList(2,intLp)) %>" width="50" border="0"></td>
	<td><%= ArrList(3,intLp) %></td>
	<td align="center"><%= ArrList(4,intLp) %></td>
</tr>
<% NEXT %>
<% End IF %>

</form>
</table>

<iframe name="regframe" src="" frameborder="0" width="0" height="0"></iframe>

<!-- #include virtual="/lib/db/dbclose.asp" -->