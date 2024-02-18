<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 yesno
' Hieditor : 2009.10.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i,page , yesnoid , yesnoword , isusing
	menupos = request("menupos")
	page = request("page")
	yesnoid = request("yesnoidsearch")
	yesnoword = request("yesnoword")
	isusing = request("isusing")			
	if page = "" then page = 1

'// 리스트
set ocontents = new cyesno_list
	ocontents.FPageSize = 20
	ocontents.FCurrPage = page
	ocontents.frectyesnoid = yesnoid
	ocontents.frectyesnoword = yesnoword
	ocontents.frectisusing = isusing			
	ocontents.fyesno_list()
%>

<script language="javascript">

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

//신규등록 & 수정
function reg(yesnoid){
	var reg = window.open('/admin/momo/yesno/yesno_reg.asp?yesnoid='+yesnoid,'reg','width=600,height=600,scrollbars=yes,resizable=yes');
	reg.focus();
}


//코맨트보기
function regcomment(yesnoid){
	var regcomment = window.open('/admin/momo/yesno/yesno_comment_list.asp?yesnoid='+yesnoid,'regcomment','width=1024,height=768,scrollbars=yes,resizable=yes');
	regcomment.focus();
}

// 진행중으로 변경 
function changestats(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.yesnoid.value = upfrm.yesnoid.value + frm.keyid.value + "," ;
						
				}
			}
		}

			var tot;
			tot = upfrm.yesnoid.value;
			upfrm.yesnoid.value = ""
		var changestats;

		changestats = window.open("/admin/momo/yesno/yesno_process.asp?yesnoid=" +tot + "&mode=ing" , "changestats","width=400,height=300,scrollbars=yes,resizable=yes");
		changestats.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="yesnoid">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>		
		<td align="left">
			yesnoid:<input type="text" name="yesnoidsearch" value="<%=yesnoid%>" size=10>
			&nbsp; yesnoword:<input type="text" name="yesnoword" value="<%=yesnoword%>" size=20>
			&nbsp; 사용여부:
			<select name="isusing" value="<%=isusing%>">
				<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>		
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm.submit();">
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
		<td align="left">
			<input type="button" onclick="changestats(frm);" value="진행중으로변경" class="button">
		</td>
		<td align="right">				
			<input type="button" onclick="reg('');" value="YESNO신규등록" class="button">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ocontents.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ocontents.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= ocontents.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td align="center">yesnoID</td>
		<td align="center">yesnoword</td>
		<td align="center">메인이미지</td>
		<td align="center">등록일</td>
		<td align="center">사용여부</td>
		<td align="center">코맨트수</td>
		
		<td align="center">비고</td>
    </tr>
	<% for i=0 to ocontents.FresultCount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			
	
    <% if ocontents.FItemList(i).fisusing = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
    <% else %>    
    <tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<% end if %>
		<td align="center">
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		</td>
		<td align="center">
			<%= ocontents.FItemList(i).fyesnoid %><input type="hidden" name="keyid" value="<%= ocontents.FItemList(i).fyesnoid %>">
		</td>
		<td align="center">
			<%= ocontents.FItemList(i).fyesnoword %>
		</td>		
		<td align="center">
			<img src="<%=webImgUrl%>/momo/yesno/main/<%= ocontents.FItemList(i).fmainimage %>" width=50 height=50>
		</td>
		<td align="center">
			<%= FormatDate(ocontents.FItemList(i).fregdate,"0000.00.00") %>
		</td>
		<td align="center">
			<%= ocontents.FItemList(i).fisusing %>
		</td>
		<td align="center">
			<% if ocontents.FItemList(i).fcommentcount > 0 then %>
			<a href="javascript:regcomment(<%= ocontents.FItemList(i).fyesnoid %>)" onfocus="this.blur();"><%= ocontents.FItemList(i).fcommentcount %></a>
			<% else %>
			<%= ocontents.FItemList(i).fcommentcount %>
			<% end if %>
		</td>			
		<td align="center"><input type="button" onclick="reg(<%= ocontents.FItemList(i).fyesnoid %>);" class="button" value="수정"></td>
    </tr>   
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if ocontents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ocontents.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocontents.StartScrollPage to ocontents.StartScrollPage + ocontents.FScrollCount - 1 %>
				<% if (i > ocontents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocontents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocontents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set ocontents = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->