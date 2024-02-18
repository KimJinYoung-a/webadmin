<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  마일리지 구분 
' History : 2007.10.23 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mileage/mileage_class.asp"-->

<%
dim page , isusing , seachjukyocd
	Page = Request("Page") 						'가지고 넘어온 Page 번호를 저장
		if Page = "" then 							'가지고 넘어온 Page 번호가 없다면
		Page = 1 
		end if
	isusing = request("isusingbox")
	seachjukyocd = request("seachjukyocd")
		
dim omileage , i
set omileage = new Cmileagelist
	omileage.FPageSize = 25							'한페이지에 들어갈 페이지수
	omileage.Fcurrpage = Page
	omileage.frectisusing = isusing
	omileage.frectseachjukyocd = seachjukyocd
	omileage.fmileagelist()
	
'########################################################### 구분셀렉트박스	
Sub Drawisusing(gubunbox,gubunid)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str
	
	response.write "<select class='select' name='" & gubunbox & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if gubunid ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">선택</option>"								'선택이란 단어가 나오도록.

	response.write "<option value='Y'"							'옵션의 값이 없으면
		if gubunid ="Y" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">사용</option>"
	
	response.write "<option value='N'"							'옵션의 값이 없으면
		if gubunid ="N" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">미사용</option>"		
	response.write "</select>"	
End Sub	
'########################################################### 년도 셀렉트박스	
%>	

<script language="javascript">

	function NextPage(page){
	frm.page.value= page;
	frm.submit();
	}
	
	function add(menupos){
	var popup
	popup = window.open('mileage_add.asp?menupos='+menupos,'add' , 'width=400,height=180,scrollbars=yes,resizable=yes');
	popup.focus();
	}
	
	function del(jukyocd){
	var popup
	popup = window.open('mileage_del_process.asp?jukyocd='+jukyocd,'del' , 'width=1,height=1,scrollbars=yes,resizable=yes');
	popup.focus();
	}

	function edit(jukyocd,menupos){
	var popup
	popup = window.open('mileage_edit.asp?jukyocd='+jukyocd+'&menupos='+menupos,'edit' , 'width=400,height=180,scrollbars=yes,resizable=yes');
	popup.focus();
	}	
	
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			상태:
			<% Drawisusing "isusingbox",isusing %>
			&nbsp;
			코드번호:
			<input type="text" class="text" name="seachjukyocd" size="10" value="<%= seachjukyocd %>">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm.submit()">
		</td>
	</tr>
	</form>
</table>	
	
<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="신규등록" onClick="javascript:add('<%= menupos %>');">
		</td>
	</tr>
</table>

<p>


<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#DDDDFF>
		<td align="center">
			마일리지코드번호
		</td>
		<td align="center">
			코드명
		</td>
		<td align="center">
			상태
		</td>
		<td align="center">
			비고
		</td>
	</tr>
<% if omileage.FResultCount > 0 then %>
	<% for i = 0 to omileage.FResultCount - 1 %>
	<tr bgcolor=#FFFFFF>
		<td align="center">
			<%= omileage.flist(i).fjukyocd %>
		</td>
		<td align="center">
			<%= omileage.flist(i).fjukyoname %>
		</td>
		<td align="center">
			<% if ucase(omileage.flist(i).fisusing) = "Y" then %>
			사용
			<% else %>
			미사용
			<% end if %>
		</td>
		<td align="center">
			<input type="button" class="button" value="수정" onclick="edit('<%= omileage.flist(i).fjukyocd %>','<%= menupos %>');">
			&nbsp;
			<input type="button" class="button" value="삭제" onclick="del('<%= omileage.flist(i).fjukyocd %>');">				
		</td>				
	</tr>
	<% next %>
<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	  	<td colspan="15"> 검색 결과가 없습니다.</td>
	</tr>
<% end if %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if omileage.HasPreScroll then %>
				<a href="javascript:NextPage('<%= omileage.StartScrollPage-1 %>')">[pre]</a>
	   		<% else %>
	    		[pre]
	   		<% end if %>
	
	    	<% for i=0 + omileage.StartScrollPage to omileage.FScrollCount + omileage.StartScrollPage - 1 %>
	    		<% if i>omileage.FTotalpage then Exit for %>
		    		<% if CStr(page)=CStr(i) then %>
		    		<font color="red">[<%= i %>]</font>
		    		<% else %>
		    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		    		<% end if %>
	    	<% next %>
	
	    	<% if omileage.HasNextScroll then %>
	    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    	<% else %>
	    		[next]
    		<% end if %>
		</td>
	</tr>
</table>	


<% 
set omileage = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
