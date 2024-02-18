<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 포스페이지 관리
' History : 2009.09.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/sitemaster_cls.asp"-->

<%
dim research,isusing, fixtype, linktype, poscode, validdate, gubun
dim page, loginuserid
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2,nowdate, datesearch


'==============================================================================
yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)
datesearch = RequestCheckvar(request("datesearch"),1)

	validdate= RequestCheckvar(request("validdate"),2)
	isusing = RequestCheckvar(request("isusing"),1)
	research= RequestCheckvar(request("research"),2)
	poscode = RequestCheckvar(request("poscode"),10)
	fixtype = request("fixtype")
	page    = RequestCheckvar(request("page"),10)
	validdate= RequestCheckvar(request("validdate"),2)
	gubun	= RequestCheckvar(request("gubun"),24)
	loginuserid = session("ssBctId")

if ((research="") and (isusing="") ) then 
    isusing = "Y"
    validdate = "on"
end if

if page="" then page=1
if gubun="" then gubun="index"
If poscode="" Then poscode=999

dim oposcode
set oposcode = new cposcode_list
	oposcode.FRectPosCode = poscode
	if (poscode<>"") then
	    oposcode.fposcode_oneitem
	end if
dim oMainContents
set oMainContents = new cposcode_list
	oMainContents.FPageSize = 100
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = isusing
	oMainContents.FRectPosCode = poscode
	oMainContents.FRectvaliddate = validdate
	If yyyy1 <> "" And datesearch="Y" Then
	oMainContents.FRectSearchSDate = yyyy1 + "-" + mm1 + "-" + dd1
	End If
	If yyyy2 <> "" And datesearch="Y" Then
	oMainContents.FRectSearchEDate = yyyy2 + "-" + mm2 + "-" + dd2
	End if
	oMainContents.FRectGubun = gubun
	oMainContents.fcontents_list


if yyyy1="" Or yyyy2="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), CLng(Mid(nowdate,6,2)),Mid(nowdate,9,2))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
	yyyy2 = Left(nowdate,4)
	mm2 = Mid(nowdate,6,2)
	dd2 = Mid(nowdate,9,2)
end If

dim i
%>

<script type="text/javascript" src="http://www.10x10.co.kr/lib/js/jquery-1.7.1.min.js"></script>
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
		AssignFlashReal = window.open("<%=wwwFingers%>/chtml/main_make_flash.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignFlashReal","width=800,height=600,scrollbars=yes,resizable=yes");
		AssignFlashReal.focus();
}

// 이미지 실서버 적용
function AssignimageReal(upfrm,poscode,imagecount,is2016){
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
		var AssignimageReal;
		if (is2016==1){
		    AssignimageReal = window.open("<%=www1Fingers%>/chtml/main_make_image.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignimageReal","width=800,height=600,scrollbars=yes,resizable=yes");
		}else{
    		AssignimageReal = window.open("<%=wwwFingers%>/chtml/main_make_image.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignimageReal","width=800,height=600,scrollbars=yes,resizable=yes");
    	}
		AssignimageReal.focus();
}

// XML 실서버 적용
function AssignXMLReal(upfrm,poscode,imagecount,is2016,uid){
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
			var newimagecount
			newimagecount = $("input:checkbox[name=cksel]:checked").length
		var AssignXMLReal;
		if (is2016==1){
		    AssignXMLReal = window.open("","AssignXMLReal","width=800,height=600,scrollbars=yes,resizable=yes");
			AssignXMLReal.location.href = "<%=www1Fingers%>/chtml/2016main_make_XML.asp?idx="+tot + '&poscode='+poscode+ '&loginuserid='+uid+'&imagecount='+newimagecount;
		}else{
    		AssignXMLReal = window.open("", "AssignXMLReal","width=800,height=600,scrollbars=yes,resizable=yes");
			AssignXMLReal.location.href = "<%=wwwFingers%>/chtml/main_make_XML.asp?idx=" +tot + '&poscode='+poscode+ '&loginuserid='+uid+'&imagecount='+newimagecount;
    	}
		AssignXMLReal.focus();
}


//포스 코드 등록 & 수정
function popPosCodeManage(){
    var popPosCodeManage = window.open('/academy/sitemaster/imagemake_poscode.asp','popPosCodeManage','width=800,height=600,scrollbars=yes,resizable=yes');
    popPosCodeManage.focus();
}

//이미지신규등록 & 수정
function AddNewMainContents(idx){
    var AddNewMainContents = window.open('/academy/sitemaster/imagemake_contents.asp?gubun=<%=gubun%>&poscode=<%=poscode%>&idx='+ idx,'AddNewMainContents','width=800,height=600,scrollbars=yes,resizable=yes');
    AddNewMainContents.focus();
}

//페이지 이동
function goPage(pg) {
	var frm = document.frm;
	frm.page.value=pg;
	frm.submit();
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
			기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %> <input type="checkbox" name="datesearch" value="Y"<% If datesearch="Y" Then Response.write " checked"%>>기간적용&nbsp;
		    사용구분
			<select name="isusing">
			<option value="">전체
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
			<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
			</select>
			&nbsp;&nbsp;
			그룹구분
			<% call DrawGroupGubunCombo ("gubun", gubun, "") %>
			&nbsp;&nbsp;
			적용구분
			<% call DrawMainPosCodeCombo("poscode", poscode, "", gubun) %>

			<% if poscode = "999" then %>
				<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전
			<% end if %>
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
		    <% 
		    '//적용구분 선택시에만 뿌림
		    if (poscode<>"") then 
		    %>
			    <% if oposcode.FOneItem.fimagetype="flash" then %>
			    	<a href="javascript:AssignFlashReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> Flash Real 적용</a>
			    <% elseif oposcode.FOneItem.fimagetype="xml" then %>
			    	<!--<a href="javascript:AssignXMLReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>,0, '<%= loginuserid %>');"><img src="/images/refreshcpage.gif" border="0"> XML Real 적용</a>-->
			    	&nbsp;
			    	<a href="javascript:AssignXMLReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>,1, '<%= loginuserid %>');"><img src="/images/refreshcpage.gif" border="0"> XML Real 적용(2016 개편)</a>
			    <% elseif oposcode.FOneItem.fimagetype="multi" then %>
			    	<a href="javascript:AssignTest('<%= poscode %>');"><img src="/images/icon_search.jpg" border="0"> 미리보기</a> 
			    	&nbsp;&nbsp;
			    	<a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>
				<% else %>
					<a href="javascript:AssignimageReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>,0);"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>
					&nbsp;
					<a href="javascript:AssignimageReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>,1);"><img src="/images/refreshcpage.gif" border="0"> Real 적용(2016 개편)</a>			    			    
			    <% end if %>
		    <% end if %>
		</td>
		<td align="right">
			<% if C_ADMIN_AUTH then %>
			<input type="button" value="코드관리" class="button" onClick="popPosCodeManage();">
			<% end if %>
		
			<input type="button" value="신규등록" class="button" onClick="javascript:AddNewMainContents('0');">						
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
		<td align="center">imagecolor</td>
	    <td align="center">구분명</td>
		<td align="center">itemcontents</td>
	    <td align="center">LinkType</td>
	    <td align="center">우선순위</td>
	    <td align="center">사용여부</td>
		<td align="center">시작일</td>
		<td align="center">종료일</td>
		<% if poscode="999" then %>
		<td align="center">상태</td>
		<% end if %>
	    <% if poscode="999" then %>
				<td align="center">등록자 ID</td>
				<td align="center">마지막 리얼적용 일시</td>
			    <td align="center">마지막 리얼적용 이미지</td>
		<% end if %>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to oMainContents.FResultCount - 1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			 		
		<% if oMainContents.FItemList(i).FIsusing="N" then %>
			<tr bgcolor="#DDDDDD">
		<% else %>
			<tr  <% if poscode="999" then %> <% if (date() >=  cdate(oMainContents.FItemList(i).FSdate) ) AND (date() <= cdate(oMainContents.FItemList(i).FEdate) ) and oMainContents.FItemList(i).FIsusing = "Y" then %> bgcolor="<%= adminColor("pink") %>" <% else %> bgcolor="#FFFFFF"<% end if %> <% else %> bgcolor="#FFFFFF" <% end if %>>
		<% end if %>	
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		
	    <td align="center"><%= oMainContents.FItemList(i).Fidx %><input type="hidden" name="idx" value="<%= oMainContents.FItemList(i).Fidx %>"></td>
	    <td align="center">
	    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');">
	    	<img width=40 height=40 src="<%=imgFingers%>/main/<%= oMainContents.FItemList(i).fimagepath %>" border="0">
	    	</a>
	    </td>
		<td align="center" bgcolor="<%= oMainContents.FItemList(i).fleftimagecolor %>"><%= oMainContents.FItemList(i).fleftimagecolor %></td>
	    <td align="center">
	    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');">
	    	<%= oMainContents.FItemList(i).Fposname %>
	    	<%
	    	if oMainContents.FItemList(i).fitemid <> 0 and oMainContents.FItemList(i).fitemid <> "" then 
	    		response.write "(" & oMainContents.FItemList(i).fitemid & ")"
	    	elseif oMainContents.FItemList(i).fevt_code <> 0 and oMainContents.FItemList(i).fevt_code <> "" then 
	    		response.write "(" & oMainContents.FItemList(i).fevt_code & ")"
	    	end if
	    	%></a>
	    </td>
		<td align="center"><%= oMainContents.FItemList(i).frelation_itemcontents %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fimagetype %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fimage_order %></td>
		<td align="center"><%= oMainContents.FItemList(i).FIsusing %></td>
		<td align="center"> <% '시작일,종료일 %>
			<% 
				Response.Write left(oMainContents.FItemList(i).FSdate,10)
			%>
		</td>
		<td align="center"> <% '시작일,종료일 %>
			<% 
				Response.Write left(oMainContents.FItemList(i).FEdate,10)
		'					Response.Write "<br />"
			%>
		</td>
		<% if poscode="999" then %>
		<td align="center">
			<%
			if (date() >=  cdate(oMainContents.FItemList(i).FSdate) ) AND (date() <= cdate(oMainContents.FItemList(i).FEdate) )  then
				if oMainContents.FItemList(i).FIsusing = "Y" then
					Response.write " <span style=""color:blue"">진행중</span>"
				else
					Response.write " <span style=""color:green"">대기중</span>"
				end if
			elseif date() < cdate(oMainContents.FItemList(i).FSdate) then
				Response.write " <span style=""color:green"">대기중</span>"
			else
				Response.write " <span style=""color:red"">종료</span>"
			end if

		'				Response.Write "<br />"
			%>
		</td> <% '상태 %>
		<% end if %>
	    <% if poscode="999" then %>
				<td align="center"><%= oMainContents.FItemList(i).fxmluserid %></td>
				<td align="center"><%= oMainContents.FItemList(i).fxmlregdate %></td>
			    <td align="center">
	    			<img width=40 height=40 src="<%= oMainContents.FItemList(i).fxmlimage %>" border="0">
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
				<span class="list_link"><a href="javascript:goPage(<%= oMainContents.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oMainContents.StartScrollPage to oMainContents.StartScrollPage + oMainContents.FScrollCount - 1 %>
				<% if (i > oMainContents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oMainContents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:goPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oMainContents.HasNextScroll then %>
				<span class="list_link"><a href="javascript:goPage(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
