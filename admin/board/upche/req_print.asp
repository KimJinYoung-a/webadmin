<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<%
'###########################################################
' Description : 업체 입점문의
' History : 2008.09.01 한용민 수정/추가
'###########################################################
%>
<%
dim i, j
dim commmode
	commmode=request("commmode")
dim page,gubun, onlymifinish
dim research, searchkey,catevalue
dim ipjumYN
	page = request("pg")
	gubun = request("gubun")
	onlymifinish = request("onlymifinish")
	research = request("research")
	searchkey = request("searchkey")
	catevalue=request("catevalue")
	ipjumYN=request("ipjumYN")
	if research="" and onlymifinish="" then onlymifinish="on"

	'// 기본값으로 입점의뢰서
	if gubun="" then gubun="01"
	if (page = "") then page = "1"

dim companyrequest
	set companyrequest = New CCompanyRequest
	companyrequest.read(request("id"))

%>
<STYLE TYPE="text/css">
<!--
	A:link, A:visited, A:active { text-decoration: none; }
	A:hover { text-decoration:underline; }
	BODY, TD, UL, OL, PRE { font-size: 10pt; }
	INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #ffffff; color: #000000; }
-->
</STYLE>

<!-- 업체정보 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="black">
<form method="post" name="f" action="/admin/board/upche/req_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="finish">
<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
	<tr bgcolor="FFFFFF" align="center">
		<td colspan=5><b><font size=3 color="blue">협력업체 관리 선정기준</font></b></td>	
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">회사명</td>	
		<td><%= db2html(companyrequest.results(0).companyname) %></td>	
		<td bgcolor="<%= adminColor("gray") %>">대표자명</td>			
		<td><%= db2html(companyrequest.results(0).chargename) %></td>				
	</tr>	
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">주소</td>	
		<td><%= db2html(companyrequest.results(0).address) %></td>
		<td bgcolor="<%= adminColor("gray") %>">구매고객</td>	
		<td>
			<%= db2html(companyrequest.results(0).cur_target) %>
		</td>							
	</tr>		
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">성명</td>			
		<td><%= db2html(companyrequest.results(0).chargename) %></td>	
		<td bgcolor="<%= adminColor("gray") %>">직책(부서명)</td>			
		<td><%= db2html(companyrequest.results(0).chargeposition) %></td>			
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">Tel</td>			
		<td><%= db2html(companyrequest.results(0).phone) %></td>	
		<td bgcolor="<%= adminColor("gray") %>">H.P</td>			
		<td><%= db2html(companyrequest.results(0).hp) %></td>	
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">사업자<br>등록번호</td>	
		<td><%= db2html(companyrequest.results(0).license_no) %></td>
		<td bgcolor="<%= adminColor("gray") %>">이메일</td>			
		<td><a href="mailto:<%= db2html(companyrequest.results(0).email) %>"><%= db2html(companyrequest.results(0).email) %></a></td>	
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">업태</td>	
		<td>
			<% 
			if companyrequest.results(0).Service <> "" then
				if left(companyrequest.results(0).Service,1) <> 0 then response.write "제조. "
				if mid(companyrequest.results(0).Service,3,1) <> 0 then response.write "도매. "
				if mid(companyrequest.results(0).Service,5,1) <> 0 then response.write "소매. "	 
				if mid(companyrequest.results(0).Service,7,1) <> 0 then response.write "수출. "
				if mid(companyrequest.results(0).Service,9,1) <> 0 then response.write "서비스. "
				if mid(companyrequest.results(0).Service,11,1) <> 0 then response.write "수입. "	
				if right(companyrequest.results(0).Service,1) <> 0 then response.write "기타. "
			end if
			%>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">주품목</td>	
		<td>
			<% Drawcatelarge "catelargebox",companyrequest.results(0).cd1 %>(<% Drawcatemid companyrequest.results(0).cd1,"catemidbox",companyrequest.results(0).cd2 %>)
		</td>				
	</tr>	
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">물류</td>	
		<td>
			<% 
			if companyrequest.results(0).physical = 0 then 
				response.write "물류시설 자체보유"
				response.write "("& companyrequest.results(0).physical_name & ")"
			else 
				response.write "물류전문업체 특정"
				response.write "("& companyrequest.results(0).physical_name & ")"
			end if
			%>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">제조</td>	
		<td>
			<% 
			if companyrequest.results(0).manufacturing = 0 then 
				response.write "생산공장 자체보유"
				response.write "("& companyrequest.results(0).manufacturing_name & ")"
			else 
				response.write "외부업체 특정"
				response.write "("& companyrequest.results(0).manufacturing_name & ")"
			end if
			%>
		</td>				
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">산업재산권 취득</td>	
		<td>
			<%= companyrequest.results(0).industrial %>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">라이센스 취득</td>	
		<td>
			<%= companyrequest.results(0).license %>
		</td>				
	</tr>	
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">유통방법</td>	
		<td>
			<% 
			if left(companyrequest.results(0).utong,1) <> 0 then response.write "아직미판매 "
			if mid(companyrequest.results(0).utong,3,1) <> 0 then response.write "백화점 "
			if mid(companyrequest.results(0).utong,5,1) <> 0 then response.write "할인점 "	 
			if mid(companyrequest.results(0).utong,7,1) <> 0 then response.write "대리점 "
			if mid(companyrequest.results(0).utong,9,1) <> 0 then response.write "종합쇼핑몰 "
			if mid(companyrequest.results(0).utong,11,1) <> 0 then response.write "홈쇼핑 "	
			if mid(companyrequest.results(0).utong,13,1) <> 0 then response.write "타사거래처 "
			if mid(companyrequest.results(0).utong,15,1) <> 0 then response.write "자사몰 "	
			if right(companyrequest.results(0).utong,1) <> 0 then response.write "자사숍 "				
			%>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">사업자형태</td>	
		<td>
			<% 
			if companyrequest.results(0).tax = 0 then 
				response.write "간이 "
			elseif  companyrequest.results(0).tax = 1 then 
				response.write "면세 "
			elseif  companyrequest.results(0).tax = 2 then 
				response.write "일반 "			
			else
				response.write "법인 "
			end if
			%>
		</td>				
	</tr>	
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">회사URL</td>	
		<td>
			<%
				dim arrUrl
				arrUrl = split(companyrequest.results(0).companyurl,",")
				if ubound(arrUrl)>0 then
					Response.Write "<a href='"
					if Left(arrUrl(0),7) <> "http://" then Response.Write "http://"
					Response.Write arrUrl(0) & "' target='_blank'>" & arrUrl(0) & "</a>"
					Response.Write "<br><br><b>입점쇼핑몰</b> : " & arrUrl(1)
				else
					Response.Write "<a href='"
					if Left(companyrequest.results(0).companyurl,7) <> "http://" then Response.Write "http://"
					Response.Write companyrequest.results(0).companyurl & "' target='_blank'>" & companyrequest.results(0).companyurl & "</a>"
				end if
			%>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">구분</td>	
		<td>
			<%= companyrequest.code2name(companyrequest.results(0).reqcd) %>
		</td>				
	</tr>		
	<tr bgcolor="FFFFFF" align="center">			
		<td bgcolor="<%= adminColor("gray") %>">상품명(브랜드명)</td>	
		<td colspan=3>
			<%= db2html(companyrequest.results(0).reqcomment) %>
		</td>				
	</tr>
	<tr bgcolor="FFFFFF" align="center">			
		<td bgcolor="<%= adminColor("gray") %>">첨부파일</td>	
		<td>
			<% if (companyrequest.results(0).attachfile <> "") then %>
				<a href="http://imgstatic.10x10.co.kr<%= companyrequest.results(0).attachfile %>" target="_blank">다운받기</a>
			<% else %>
				없음
			<% end if %>
		</td>						
		<td bgcolor="<%= adminColor("gray") %>">처리상태</td>	
		<td>
			<% if (IsNull(companyrequest.results(0).finishdate) = true) then %>
				미완료
			<% else %>
				<%= FormatDate(companyrequest.results(0).finishdate, "0000-00-00") %>
			<% end if %>
		</td>
	</tr>		
	<tr bgcolor="FFFFFF" align="center">			
		<td bgcolor="<%= adminColor("gray") %>">회사설명</td>	
		<td colspan=3>
			<%= nl2br(db2html(companyrequest.results(0).companycomments)) %>
		</td>				
	</tr>
	<tr bgcolor="FFFFFF" align="center">			
		<td colspan=4 align="left">
		<input type="button" value="프린트" class="button" onclick="javascript:window.print();">
		</td>				
	</tr>	
</table><br>
<!-- 업체정보 끝 -->

<!-- #include virtual="/lib/db/dbclose.asp" -->