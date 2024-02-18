<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Thanks 10x10 리스트 페이지  
' History : 2008.04.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<%
Dim oip,i,page ,isusing_search 
	isusing_search = request("isusing_searchbox")
	page = request("page")
	if page = "" then page = 1

'// 고객이 쓴글 리스트
set oip = new cthanks10x10_list
	oip.FPageSize = 20
	oip.FCurrPage = page
	oip.frectisusing = isusing_search
	oip.fthanks10x10_list()
%>

<script language="javascript">

function regpop(idx,gubun){
	var regpop = window.open('/admin/culturestation/thanks10x10_reg.asp?idx='+idx+'&gubun='+gubun,'regpop','width=900,height=600,scrollbars=yes,resizable=yes');
	regpop.focus();
}

function del(idx){
	var ret;
		ret = confirm('고객님이 작성한 글을 정말 삭제 하시겠습니까?');
	
	if (ret){
	var del = window.open('/admin/culturestation/thanks10x10_process.asp?idx='+idx+'&mode=del','del','width=800,height=600,scrollbars=yes,resizable=yes');
	del.focus();
	}
}

function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}	

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}	

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function thanks10x10_reg(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
					
				}
			}
		}
			var tot;
			tot = upfrm.fidx.value;
			upfrm.fidx.value = ""
		var addreg;
		addreg = window.open("/admin/culturestation/thanks10x10_process.asp?idx=" +tot, "addreg","width=400,height=300,scrollbars=yes,resizable=yes");
		addreg.focus();
	}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="mode">
	<input type="hidden" name="fidx">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<select name="isusing_searchbox" value="<%=isusing_search %>">
				<option value="" <% if isusing_search = "" then response.write " selected" %>>실서버등록여부</option>
				<option value="Y" <% if isusing_search = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing_search = "N" then response.write " selected" %>>N</option>
			</select> 
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left"><input type="button" class="button" value="실서버등록" onclick="thanks10x10_reg(frm);">
		</td>
		<td align="right">	
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td align="center">번호</td>
		<td align="center">IMAGE</td>
		<td align="center">고객ID</td>		
		<td align="center">내용</td>		
		<td align="center">등록일</td>	
		<td align="center">답글여부</td>		
		<td align="center">실서버<br>적용</td>
		<td align="center">비고</td>
    </tr>
	<% for i=0 to oip.FresultCount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			<!--for문 안에서 i 값을 가지고 루프-->	
    <tr align="center" bgcolor="#FFFFFF">
 			<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			<td align="center"><%= oip.fitemlist(i).fidx %><input type="hidden" name="idx" value="<%= oip.fitemlist(i).fidx %>"></td>		
			<td align="center"><%= drawgubun(oip.FItemList(i).fgubun) %></td>
			<td align="center"><%= oip.fitemlist(i).fuserid %></td>		
			<td align="center"><%= left(oip.fitemlist(i).fcontents,20)&"..." %></td>
			<td align="center"><%= left(oip.fitemlist(i).freg_date,10) %></td>
			<td align="center">
				<% if oip.fitemlist(i).fcomment = "" then %>
					<a href="javascript:regpop('<%= oip.fitemlist(i).fidx %>','add');">N [작성]</a>
				<% else %>
					<a href="javascript:regpop('<%= oip.fitemlist(i).fidx %>','edit');">Y [수정]</a>
				<% end if%>
			</td>
			<td align="center"><%= oip.fitemlist(i).fisusing_display %></td>	
			<td align="center"><input type="button" class="button" value="삭제" onclick="javascript:del(<%= oip.fitemlist(i).fidx %>);"></td>
    </tr>   
	</form>
	<% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
