<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이미지 & 링크 파일 생성 리스트 페이지   
' History : 2009.03.24 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/brandstreet/brandstreet_cls.asp"-->

<%
dim research,isusing, fixtype, linktype, poscode, validdate
dim page

	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")

if ((research="") and (isusing="")) then 
    isusing = "Y"
    validdate = "on"
end if

if page="" then page=1

dim oposcode
set oposcode = new cbrandstreet_list
	oposcode.FRectPosCode = poscode
	if (poscode<>"") then
	    oposcode.fposcode_oneitem
	end if

dim oMainContents
set oMainContents = new cbrandstreet_list
	oMainContents.FPageSize = 10
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = isusing
	oMainContents.FRectPosCode = poscode
	oMainContents.FRectvaliddate = validdate
	oMainContents.fcontents_list

dim i
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

// 플래쉬 실서버 적용
function AssignFlashReal(upfrm,poscode,imagecount){
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
		var AssignFlashReal;
		AssignFlashReal = window.open("<%=wwwUrl%>/chtml/brandstreet_event_flashmake.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignFlashReal","width=800,height=600,scrollbars=yes,resizable=yes");
		AssignFlashReal.focus();
}

// 플래쉬 실서버 적용
function AssignlinkReal(upfrm,poscode,imagecount){
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
		var AssignFlashReal;
		AssignFlashReal = window.open("<%=wwwUrl%>/chtml/brandstreet_event_linkmake.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignFlashReal","width=800,height=600,scrollbars=yes,resizable=yes");
		AssignFlashReal.focus();
}

//포스 코드 등록 & 수정
function popPosCodeManage(){
    var popPosCodeManage = window.open('/admin/brandstreet/imagemake_poscode.asp','popPosCodeManage','width=800,height=600,scrollbars=yes,resizable=yes');
    popPosCodeManage.focus();
}

//이미지신규등록 & 수정
function AddNewMainContents(idx){
    var AddNewMainContents = window.open('/admin/brandstreet/imagemake_contents.asp?idx='+ idx,'AddNewMainContents','width=800,height=600,scrollbars=yes,resizable=yes');
    AddNewMainContents.focus();
}


</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="fidx">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		    <!--<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전-->
		    사용구분
			<select name="isusing">
			<option value="">전체
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
			<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
			</select>
			적용구분
			<% call DrawMainPosCodeCombo("poscode", poscode,"") %>
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
		<td align="left">
		    <% if (poscode<>"") then %>
			    <% if oposcode.FOneItem.fimagetype="flash" then %>
			    	<a href="javascript:AssignFlashReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> Flash Real 적용</a>
			    <% elseif oposcode.FOneItem.fimagetype="link" then %>
			    	<% if poscode = "200" then %>
					<a href="javascript:AssignlinkReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> 배너 Real 적용</a>
			    	<% end if %>
			    <% elseif oposcode.FOneItem.fimagetype="multi" then %>
			    	
			    	<a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>
			    <% end if %>			   
		    <% end if %>
		</td>
		<td align="right">
			<% if C_ADMIN_AUTH then %>
			<input type="button" value="코드관리" class="button" onClick="popPosCodeManage();">
			<% end if %>
		
			<a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0"></a>			
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oMainContents.FResultCount > 0 then %> 
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oMainContents.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
 		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	    <td align="center">Idx</td>
	    <td align="center">Image</td>
	    <td align="center">구분명</td>
	    <td align="center">LinkType</td>
	    <td align="center">우선순위</td>
	    <td align="center">사용여부</td>
	    <td align="center">등록일</td>
		    <% if Not(oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Fimagetype="flash") then %>
				<td></td>
		    <% end if %> 
    </tr>
	<% for i=0 to oMainContents.FResultCount - 1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			<!--for문 안에서 i 값을 가지고 루프-->	 		
		<% if oMainContents.FItemList(i).FIsusing="N" then %>
			<tr align="center" bgcolor="#DDDDDD" onmouseover=this.style.background="orange"; onmouseout=this.style.background="#FFFFFF";>
		<% else %>
			<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background="#FFFFFF";>
		<% end if %>	
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		
	    <td align="center"><%= oMainContents.FItemList(i).Fidx %><input type="hidden" name="idx" value="<%= oMainContents.FItemList(i).Fidx %>"></td>
	    <td align="center">
	    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');">
	    	<img width=40 height=40 src="<%=uploadUrl%>/brandstreet/main/<%= oMainContents.FItemList(i).fimagepath %>" border="0">
	    	</a>
	    </td>
	    <td align="center">
	    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><%= oMainContents.FItemList(i).Fposname %></a>
	    	(<%= oMainContents.FItemList(i).flinkpath %>)
	    </td>
	    <td align="center"><%= oMainContents.FItemList(i).fimagetype %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fimage_order %></td>
	    <td align="center"><%= oMainContents.FItemList(i).FIsusing %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fregdate %></td> 

	    <% if Not(oMainContents.FItemList(i).Fposcode="200" or oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Fimagetype="flash" or oMainContents.FItemList(i).Fimagetype="multi") then %>
			<td>
				<a href="javascript:AssignDailyReal('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>
			</td>
	    <% end if %> 

	</tr>
	</form>	
	<% next %>
    </tr>   
    
<% else %>

	<tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oMainContents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oMainContents.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oMainContents.StartScrollPage to oMainContents.StartScrollPage + oMainContents.FScrollCount - 1 %>
				<% if (i > oMainContents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oMainContents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oMainContents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

		


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

