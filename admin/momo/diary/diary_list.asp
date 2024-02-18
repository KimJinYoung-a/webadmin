<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 다이어리
' Hieditor : 2009.12.01 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i,page , isusing , diarytype , diary_date
	diary_date = request("diary_date")
	menupos = request("menupos")
	page = request("page")
	diarytype = request("diarytype")
	isusing = request("isusing")
	if page = "" then page = 1

'// 리스트
set ocontents = new Cdiary_list
	ocontents.FPageSize = 20
	ocontents.FCurrPage = page
	ocontents.frectisusing = isusing
	ocontents.frectdiarytype = diarytype	
	ocontents.frectdiary_date = diary_date	
	ocontents.fdiary_contents_list()
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
	function AddNewMainContents(idx){
	    var AddNewMainContents = window.open('/admin/momo/diary/diary_edit.asp?idx='+ idx,'AddNewMainContents','width=800,height=768,scrollbars=yes,resizable=yes');
	    AddNewMainContents.focus();
	}
	
	//이미지 신규등록 & 수정
	function Addimage(idx){
	    var Addimage = window.open('/admin/momo/diary/diary_image_edit.asp?idx='+ idx,'Addimage','width=600,height=400,scrollbars=yes,resizable=yes');
	    Addimage.focus();
	}

	//코맨트보기
	function regcomment(idx){
		var regcomment = window.open('/admin/momo/diary/diary_comment_list.asp?idx='+idx,'regcomment','width=1024,height=768,scrollbars=yes,resizable=yes');
		regcomment.focus();
	}
	
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="fidx">	
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>		
		<td align="left">
			날짜:<input type="text" name="diary_date" size=10 value="<%= diary_date %>">			
			<a href="javascript:calendarOpen3(frm.diary_date,'시작일',frm.diary_date.value)">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>		
		    사용구분
			<select name="isusing">
			<option value="">전체
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
			<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
			</select>
			고객참여여부: <select name="diarytype">
				<option value="" <% if diarytype = "" then response.write " selected" %>>선택</option>
				<option value="withyou" <% if diarytype="withyou" then response.write " selected" %>>withyou</option>
				<option value="with10x10" <% if diarytype="with10x10" then response.write " selected" %>>with10x10</option>
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
		</td>
		<td align="right">		
			<input type="button" value="신규등록" class="button" onClick="javascript:AddNewMainContents('');">					
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
	    <td align="center">Idx</td>
	    <td align="center">날짜</td>
	    <td align="center">제목</td>
	    <td align="center">고객<br>참여여부</td>
	    <td align="center">사용여부</td>
	    <td align="center">코맨트수</td>
	    <td align="center">비고</td>
    </tr>
    
	<% for i=0 to ocontents.fresultcount -1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			 		
		<% if ocontents.FItemList(i).fisusing="N" then %>
			<tr bgcolor="#DDDDDD" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
		<% else %>
			<tr bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
		<% end if %>	
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		
	    <td align="center"><%= ocontents.FItemList(i).fidx %></td>
	    <td align="center"><%= FormatDate(ocontents.FItemList(i).fdiary_date,"0000.00.00") %></td>
	    <td align="center" ><%= chrbyte(ocontents.FItemList(i).ftitle,20,"Y") %></td> 				    
	    <td align="center"><%= ocontents.FItemList(i).fdiarytype %></td>
	    <td align="center"><%= ocontents.FItemList(i).fisusing %></td> 
	    <td align="center">
			<% if ocontents.FItemList(i).fcommentcount > 0 then %>
			<a href="javascript:regcomment(<%= ocontents.FItemList(i).fidx %>)" onfocus="this.blur();">보기[<%= ocontents.FItemList(i).fcommentcount %>]</a>
			<% else %>
			<%= ocontents.FItemList(i).fcommentcount %>
			<% end if %>	    
	    </td>
	    <td align="center">
	    	<input type="button" onclick="AddNewMainContents(<%= ocontents.FItemList(i).fidx %>);" class="button" value="내용수정하기">
	    	<input type="button" onclick="Addimage(<%= ocontents.FItemList(i).fidx %>);" class="button" value="이미지등록">	    	
	    </td>
	</tr>
	</form>	
	<% next %>			
    </tr>   

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if ocontents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ocontents.StartScrollPage-1 %>&isusing=<%=isusing%>&diarytype=<%=diarytype%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocontents.StartScrollPage to ocontents.StartScrollPage + ocontents.FScrollCount - 1 %>
				<% if (i > ocontents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocontents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>&diarytype=<%=diarytype%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocontents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>&diarytype=<%=diarytype%>">[next]</a></span>
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