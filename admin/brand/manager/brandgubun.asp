<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.30 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/managerCls.asp"-->
<%
Dim list, page, i, makerid
	page	= request("page")
	makerid	= request("makerid")
	menupos	= request("menupos")
	
If page = ""	Then page = 1
	
SET list = new cmanager
	list.FCurrPage		= page
	list.FPageSize		= 100
	list.frectisusing = "Y"
	list.sbbrandgubunlist
%>

<script language="javascript">

var ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

function jsedit() {
	var brandgubun=""; var hello_yn=""; var interview_yn=""; var tenbytenand_yn=""; var artistwork_yn=""; var shop_collection_yn=""; var shop_event_yn=""; var lookbook_yn="";
	var chkSel = 0;
	
	var chkI = document.getElementsByName("chkI")
	var tmphello_yn = document.getElementsByName("hello_yn")
	var tmpinterview_yn = document.getElementsByName("interview_yn")
	var tmptenbytenand_yn = document.getElementsByName("tenbytenand_yn")
	var tmpartistwork_yn = document.getElementsByName("artistwork_yn")
	var tmpshop_collection_yn = document.getElementsByName("shop_collection_yn")
	var tmpshop_event_yn = document.getElementsByName("shop_event_yn")	
	var tmplookbook_yn = document.getElementsByName("lookbook_yn")

	for (var i=0;i<chkI.length;i++){
		if (chkI[i].checked){
			if(tmphello_yn[i].options[tmphello_yn[i].selectedIndex].value==""){			
				alert('hello 권한을 선택해주세요');
				tmphello_yn[i].focus();
				return;
			}
			if(tmpinterview_yn[i].options[tmpinterview_yn[i].selectedIndex].value==""){			
				alert('interview 권한을 선택해주세요');
				tmpinterview_yn[i].focus();
				return;
			}
			if(tmptenbytenand_yn[i].options[tmptenbytenand_yn[i].selectedIndex].value==""){			
				alert('tenbytenand 권한을 선택해주세요');
				tmptenbytenand_yn[i].focus();
				return;
			}
			if(tmpartistwork_yn[i].options[tmpartistwork_yn[i].selectedIndex].value==""){			
				alert('artistwork 권한을 선택해주세요');
				tmpartistwork_yn[i].focus();
				return;
			}
			if(tmpshop_collection_yn[i].options[tmpshop_collection_yn[i].selectedIndex].value==""){			
				alert('shop_collection 권한을 선택해주세요');
				tmpshop_collection_yn[i].focus();
				return;
			}
			if(tmpshop_event_yn[i].options[tmpshop_event_yn[i].selectedIndex].value==""){			
				alert('shop_event 권한을 선택해주세요');
				tmpshop_event_yn[i].focus();
				return;
			}
			if(tmplookbook_yn[i].options[tmplookbook_yn[i].selectedIndex].value==""){			
				alert('lookbook 권한을 선택해주세요');
				tmplookbook_yn[i].focus();
				return;
			}	
													
			chkSel++;
			brandgubun = brandgubun + chkI[i].value + ",";
			hello_yn = hello_yn + tmphello_yn[i].value + ",";
			interview_yn = interview_yn + tmpinterview_yn[i].value + ",";
			tenbytenand_yn = tenbytenand_yn + tmptenbytenand_yn[i].value + ",";
			artistwork_yn = artistwork_yn + tmpartistwork_yn[i].value + ",";
			shop_collection_yn = shop_collection_yn + tmpshop_collection_yn[i].value + ",";
			shop_event_yn = shop_event_yn + tmpshop_event_yn[i].value + ",";
			lookbook_yn = lookbook_yn + tmplookbook_yn[i].value + ",";
		}
	}
	if(chkSel<=0) {
		alert("선택상품이 없습니다.");
		return;
	}

	document.frmedit.mode.value = "EDIT";
	document.frmedit.brandgubunarr.value = brandgubun;
	document.frmedit.hello_ynarr.value = hello_yn;
	document.frmedit.interview_ynarr.value = interview_yn;
	document.frmedit.tenbytenand_ynarr.value = tenbytenand_yn;
	document.frmedit.artistwork_ynarr.value = artistwork_yn;
	document.frmedit.shop_collection_ynarr.value = shop_collection_yn;
	document.frmedit.shop_event_ynarr.value = shop_event_yn;
	document.frmedit.lookbook_ynarr.value = lookbook_yn;
	document.frmedit.submit();
}
	
</script>

<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>브랜드매뉴권한지정</b>

<form name="frmedit" method="post" action="brandgubunProcess.asp" style="margin:0px;">
	<input type="hidden" name="brandgubunarr" value="">
	<input type="hidden" name="hello_ynarr" value="">
	<input type="hidden" name="interview_ynarr" value="">
	<input type="hidden" name="tenbytenand_ynarr" value="">
	<input type="hidden" name="artistwork_ynarr" value="">
	<input type="hidden" name="shop_collection_ynarr" value="">
	<input type="hidden" name="shop_event_ynarr" value="">
	<input type="hidden" name="lookbook_ynarr" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mode" value="">
</form>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% if list.fresultcount >0 then %>	
			<input class="button" type="button" id="btnEditSel" value="선택 수정" onClick="jsedit();">
			&nbsp;&nbsp;
	    <% end if %>	
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=list.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= list.FTotalPage %></b>		
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
	<td>브랜드구분</td>
	<td>브랜드구분명</td>
	<td>
		<b><font color="orange">업체등록</font></b>			
		<Br>hello
	</td>
	<td>
		<b><font color="blue">MD등록</font></b>
		<Br>interview
	</td>
	<td>
		<b><font color="blue">MD등록</font></b>
		<Br>tenbytenand
	</td>
	<td>
		<b><font color="orange">업체등록</font></b>		
		<Br>artistwork
	</td>
	<td>
		<b><font color="red">업체등록->MD컨펌</font></b>
		<Br>shop_collection
	</td>
	<td>
		<b><font color="blue">MD등록</font></b>		
		<Br>shop_event
	</td>
	<td>
		<b><font color="red">업체등록->MD컨펌</font></b>		
		<Br>lookbook
	</td>
	<td>최근수정</td>
	<td>비고</td>
</tr>
<% if list.fresultcount >0 then %>
<% For i = 0 to list.fresultcount -1 %>
<tr height="25" bgcolor="FFFFFF" align="center">
	<td align="center"><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= list.FItemlist(i).fbrandgubun %>"></td>
	<td><%= list.FItemlist(i).fbrandgubun %></td>
	<td><%= list.FItemlist(i).fbrandgubunname %></td>
	<td><% drawSelectBoxUsingYN "hello_yn", list.FItemlist(i).fhello_yn %></td>
	<td><% drawSelectBoxUsingYN "interview_yn", list.FItemlist(i).finterview_yn %></td>
	<td><% drawSelectBoxUsingYN "tenbytenand_yn", list.FItemlist(i).ftenbytenand_yn %></td>
	<td><% drawSelectBoxUsingYN "artistwork_yn", list.FItemlist(i).fartistwork_yn %></td>
	<td><% drawSelectBoxUsingYN "shop_collection_yn", list.FItemlist(i).fshop_collection_yn %></td>
	<td><% drawSelectBoxUsingYN "shop_event_yn", list.FItemlist(i).fshop_event_yn %></td>
	<td><% drawSelectBoxUsingYN "lookbook_yn", list.FItemlist(i).flookbook_yn %></td>
	<td>
		<%= list.FItemlist(i).flastupdate %>
		<br>(<%= list.FItemlist(i).flastadminid %>)
	</td>
	<td>
	</td>
</tr>
<% Next %>

<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If list.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= list.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + list.StartScrollPage to list.StartScrollPage + list.FScrollCount - 1 %>
			<% If (i > list.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(list.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If list.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</form>
</table>

<%
SET list = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->