<%@ CODEPAGE = 0 %>
<% option explicit %>
<%
'###########################################################
' Description :  옥션 상품 관리 페이지
' History : 2007.09.11 한용민 생성
'###########################################################

'0 : ANSI (기본값) 
'949 : 한국어 (EUC-KR) 
'65001 : 유니코드 (UTF-8) 
'65535 : 유니코드 (UTF-16)
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->
	
<%
dim oip, i,page, auction,ten,makerid,magin,auction_category
	makerid = request("makeridbox")
	ten = request("tenbox")
	auction = request("auctionbox")					'상태값 검색을 위한 변수
	magin = request("maginbox")
	auction_category = request("auction_categorybox")
	page = Request("Page") 						'가지고 넘어온 Page 번호를 저장
		if Page = "" then 							'가지고 넘어온 Page 번호가 없다면
		Page = 1 
		end if
	
set oip = new Cauctionlist        			'클래스 지정
oip.FPageSize = 100							'한페이지에 들어갈 페이지수
oip.Fcurrpage = Page
oip.frectauction = auction
oip.frectten = ten
oip.frectmakerid = makerid
oip.frectmagin = magin
oip.fauction_category = auction_category
oip.fauctionlist()								'클래스를 실행


Sub Drawauction(selectboxname, stats)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str ,a

	response.write "<select name='" & selectboxname & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if stats ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">선택</option>"								'선택이란 단어가 나오도록.

	response.write "<option value='y'"							'옵션의 값이 없으면
		if stats ="y" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">Y</option>"
	
	response.write "<option value='n'"							'옵션의 값이 없으면
		if stats ="n" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">N</option>"
			
	response.write "</select>"
End Sub
'##################################################################
Sub Drawten(selectboxname, stats)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str ,a

	response.write "<select name='" & selectboxname & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if stats ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">선택</option>"								'선택이란 단어가 나오도록.

	response.write "<option value='y'"							'옵션의 값이 없으면
		if stats ="y" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">Y</option>"
	
	response.write "<option value='n'"							'옵션의 값이 없으면
		if stats ="n" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">N</option>"
			
	response.write "</select>"
End Sub
'##################################################################
Sub Drawmakerid(boxname, stats)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색.boxname은 sub구문에서만 쓰임
	dim userquery, tem_str

	response.write "<select name='" & boxname & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if stats ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">선택</option>"								'선택이란 단어가 나오도록.

	'사용자 검색 옵션 내용 DB에서 가져오기
		userquery = "select makerid"
		userquery = userquery + " from [db_item].dbo.tbl_auction a"
		userquery = userquery + " left join [db_item].[dbo].tbl_item b"
		userquery = userquery + " on a.ten_itemid = b.itemid"
		userquery = userquery + " group by makerid"
	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(stats) = Lcase(rsget("makerid")) then 	
				tem_str = " selected"								
			end if
			response.write "<option value='" & rsget("makerid") & "' " & tem_str & ">" & db2html(rsget("makerid")) & "</option>"
			tem_str = ""				
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
'##################################################################
Sub Drawmagin(boxname, stats)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색.boxname은 sub구문에서만 쓰임
		dim userquery, tem_str ,a

	response.write "<select name='" & boxname & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if stats ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">선택</option>"								'선택이란 단어가 나오도록.

	response.write "<option value='20'"							'옵션의 값이 없으면
		if stats ="20" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">20%이상</option>"
	
	response.write "<option value='10000'"							'옵션의 값이 없으면
		if stats ="10000" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">20%미만</option>"
			
	response.write "</select>"
End Sub
'##################################################################
Sub Draw_auction_category(boxname, stats)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색.boxname은 sub구문에서만 쓰임
	dim userquery, tem_str

	response.write "<select name='" & boxname & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if stats ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">선택</option>"								'선택이란 단어가 나오도록.

	'사용자 검색 옵션 내용 DB에서 가져오기
		userquery = "select auction_cate_code"
		userquery = userquery + " from [db_item].dbo.tbl_auction"
		userquery = userquery + " group by auction_cate_code"
	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(stats) = Lcase(rsget("auction_cate_code")) then 	
				tem_str = " selected"								
			end if
			response.write "<option value='" & rsget("auction_cate_code") & "' " & tem_str & ">"
			if rsget("auction_cate_code") = "10010100" then
				response.write "노트/연습장"
			elseif rsget("auction_cate_code") = "10010200" then
				response.write "클리어파일"
			elseif rsget("auction_cate_code") = "10010300" then
				response.write "포스트잇/메모지"
			elseif rsget("auction_cate_code") = "10010400" then
				response.write "화이트/수정용품"
			elseif rsget("auction_cate_code") = "10010500" then
				response.write "클립/집게/홀더"
			elseif rsget("auction_cate_code") = "10010600" then
				response.write "칼/가위/자"
			elseif rsget("auction_cate_code") = "10010700" then
				response.write "스템플러/리무버"
			elseif rsget("auction_cate_code") = "10010800" then
				response.write "풀/테이프"		
			elseif rsget("auction_cate_code") = "10010900" then
				response.write "펀치"
			elseif rsget("auction_cate_code") = "10010900" then
				response.write "문구세트"
			elseif rsget("auction_cate_code") = "10011000" then
				response.write "화방/제도용품"
			elseif rsget("auction_cate_code") = "10011200" then
				response.write "문구용품기타"
				
			elseif rsget("auction_cate_code") = "10030100" then
				response.write "접착식앨범"
			elseif rsget("auction_cate_code") = "10030200" then
				response.write "포켓식앨범"
			elseif rsget("auction_cate_code") = "10030300" then
				response.write "앨범기타"
			elseif rsget("auction_cate_code") = "10040100" then
				response.write "만년필"
			elseif rsget("auction_cate_code") = "10040200" then
				response.write "매직/네임펜/마카"
			elseif rsget("auction_cate_code") = "10040301" then
				response.write "유성펜"
			elseif rsget("auction_cate_code") = "10040302" then
				response.write "수성펜"
			elseif rsget("auction_cate_code") = "10040400" then
				response.write "샤프/연필/색연필"
			elseif rsget("auction_cate_code") = "10040500" then
				response.write "형광/사인펜"
			elseif rsget("auction_cate_code") = "10040600" then
				response.write "제도용펜/특수펜"
			elseif rsget("auction_cate_code") = "10040700" then
				response.write "필기구기타"
			
			elseif rsget("auction_cate_code") = "10050100" then
				response.write "도장"
			elseif rsget("auction_cate_code") = "10050200" then
				response.write "스탬프"
			elseif rsget("auction_cate_code") = "10050300" then
				response.write "파일/바인더"
			elseif rsget("auction_cate_code") = "10050400" then
				response.write "명함집/케이스"	
			elseif rsget("auction_cate_code") = "10050500" then
				response.write "잉크/토너"
			elseif rsget("auction_cate_code") = "10050600" then
				response.write "서류보관함"
			elseif rsget("auction_cate_code") = "10050700" then
				response.write "칠판/보드"
			elseif rsget("auction_cate_code") = "10050900" then
				response.write "사무용가구"																														
			elseif rsget("auction_cate_code") = "10051000" then
				response.write "사무용품기타"
				
			elseif rsget("auction_cate_code") = "10060101" then
				response.write "케릭터다이어리"
			elseif rsget("auction_cate_code") = "10060102" then
				response.write "일러스트다이어리"
			elseif rsget("auction_cate_code") = "10060103" then
				response.write "만년다이어리"	
			elseif rsget("auction_cate_code") = "10060104" then
				response.write "핸드매이드다이어리"
			elseif rsget("auction_cate_code") = "10060201" then
				response.write "스터디다이어리"
			elseif rsget("auction_cate_code") = "10060202" then
				response.write "포토다이어리"
			elseif rsget("auction_cate_code") = "10060301" then
				response.write "프랭클린다이어리"
			elseif rsget("auction_cate_code") = "10060302" then
				response.write "시스템다이어리"
			elseif rsget("auction_cate_code") = "99140700" then
				response.write "다이어리속지"
			elseif rsget("auction_cate_code") = "10060500" then
				response.write "다이어리기타"
					
			elseif rsget("auction_cate_code") = "10070100" then
				response.write "계산기"	
			elseif rsget("auction_cate_code") = "10071000" then
				response.write "사무기기기타"
			elseif rsget("auction_cate_code") = "10090200" then
				response.write "편지/서류봉투"
			elseif rsget("auction_cate_code") = "10090300" then
				response.write "견출지/라벨류"
			elseif rsget("auction_cate_code") = "10090700" then
				response.write "편지지/엽서"
			elseif rsget("auction_cate_code") = "10090800" then
				response.write "장부/양식/서식지"
			elseif rsget("auction_cate_code") = "10090900" then
				response.write "종이류기타"
				
			elseif rsget("auction_cate_code") = "99140100" then
				response.write "이색상품"
			elseif rsget("auction_cate_code") = "99140200" then
				response.write "캐릭터용품"	
			elseif rsget("auction_cate_code") = "99140300" then
				response.write "주문제작/맞춤선물"
			elseif rsget("auction_cate_code") = "99140400" then
				response.write "포토앨범/박스/홀더"
			elseif rsget("auction_cate_code") = "99140500" then
				response.write "키덜트용품"
			elseif rsget("auction_cate_code") = "99140600" then
				response.write "디자인소품"		
			elseif rsget("auction_cate_code") = "99140700" then
				response.write "아이디어소품"
			elseif rsget("auction_cate_code") = "99140800" then
				response.write "장식소품"	
			elseif rsget("auction_cate_code") = "99140900" then
				response.write "기타상품"										
			end if 
			response.write "</option>"
			tem_str = ""				
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
'##################################################################
%>

<!-- #include virtual="/admin/auction/auction.js"-->

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="fidx">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드: <% drawmakerid "makeridbox" , makerid %>
			텐재고: <% Drawten "tenbox", ten %> 
			옥션등록: <% Drawauction "auctionbox", auction %> 
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			마진: <% Drawmagin "maginbox", magin %>
			카테고리: <% Draw_auction_category "auction_categorybox", auction_category %>
		</td>
	</tr>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<select name="auctionup_select" onchange="javascript:auctionup(this.value,frm);">
				<option value="">AUCTION upload 여부</option>
				<option value="y">Y</option>		  
				<option value="n">N</option>	
			</select><br>					
			카테고리변경 : <div id="cd1_display" style="display:inline">	
				<select name="cd1" onchange="javascript:search1();">
					<option value="">대카테고리선택</option>
					<option value="1">문구/사무/용지</option>		  
					<option value="2">꽃/팬시/서비스</option>	
				</select>	
			</div>		  
			<div id="cd2_display_1" style="display:none">	
				<select name="cd2_1" onchange="javascript:search2('cd2_1');">
					<option value="">중카테고리선택</option>
					<option value="1">문구용품</option>		  
					<option value="2">앨범</option>	
					<option value="3">필기구</option>
					<option value="4">사무용품</option>
					<option value="5">다이어리</option>
					<option value="6">사무기기</option>
					<option value="7">종이류</option>								
				</select>				
			</div>
			<div id="cd2_display_2" style="display:none">	
				<select name="cd2_2" onchange="javascript:search2('cd2_2');">
					<option value="">중카테고리선택</option>
					<option value="1">디자인/아이디어소품</option>		  							
				</select>				
			</div>		  
			<div id="cd3_display_1" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">소카테고리선택</option>
					<option value="10010100">노트/연습장</option>		  
					<option value="10010200">클리어파일</option>	
					<option value="10010300">포스트잇/메모지</option>
					<option value="10010400">화이트/수정용품</option>
					<option value="10010500">클립/집게/홀더</option>
					<option value="10010600">칼/가위/자</option>
					<option value="10010700">스템플러/리무버</option>
					<option value="10010800">풀/테이프</option>
					<option value="10010900">펀치</option>				
					<option value="10011000">문구세트</option>												
					<option value="10011100">화방/제도용품</option>
					<option value="10011200">문구용품기타</option>							
				</select>				
			</div>	
			<div id="cd3_display_2" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">소카테고리선택</option>
					<option value="10030100">접착식앨범</option>		  
					<option value="10030200">포켓식앨범</option>	
					<option value="10030300">앨범기타</option>						
				</select>				
			</div>
			<div id="cd3_display_3" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">소카테고리선택</option>
					<option value="10040100">만년필</option>		  
					<option value="10040200">매직/네임펜/마카</option>	
					<option value="10040301">유성펜</option>
					<option value="10040302">수성펜</option>
					<option value="10040400">샤프/연필/색연필</option>
					<option value="10040500">형광/사인펜</option>
					<option value="10040600">제도용펜/특수펜</option>
					<option value="10040700">필기구기타</option>						
				</select>				
			</div>		
			<div id="cd3_display_4" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">소카테고리선택</option>
					<option value="10050100">도장</option>		  
					<option value="10050200">스템프</option>	
					<option value="10050300">파일/바인더</option>
					<option value="10050400">명함집/케이스</option>
					<option value="10050500">잉크/토너</option>
					<option value="10050600">서류보관함</option>
					<option value="10050700">칠판/보드</option>
					<option value="10050900">사무용가구</option>						
					<option value="10051000">사무용품기타</option>	
				</select>	
			</div>		
			<div id="cd3_display_5" style="display:none">							
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">소카테고리선택</option>
					<option value="10060101">캐릭터다이어리</option>		  
					<option value="10060102">일러스트다이어리</option>	
					<option value="10060103">만년다이어리</option>
					<option value="10060104">핸드메이드다이어리</option>
					<option value="10060201">스터디다이어리</option>
					<option value="10060202">포토다이어리</option>
					<option value="10060301">프랭클린플래너</option>
					<option value="10060302">시스템다이어리기타</option>						
					<option value="10060400">다이이리속지</option>
					<option value="10060500">다이어리기타</option>
					<option value="10060100">팬시다이어리</option>
					<option value="10060200">기능성다이어리</option>
					<option value="10060300">시스템다이어리</option>													
				</select>
			</div>		
			<div id="cd3_display_6" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">소카테고리선택</option>
					<option value="10070100">계산기</option>		  
					<option value="10071000">사무기기기타</option>							
				</select>
			</div>	
			<div id="cd3_display_7" style="display:none">							
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">소카테고리선택</option>
					<option value="10090200">편지/서류봉투</option>		  
					<option value="10090300">견출지/라벨류</option>	
					<option value="10090700">편지지/엽서</option>
					<option value="10090800">장부/양식/서식지</option>
					<option value="10090900">종이류기타</option>					
				</select>
			</div>	
			<div id="cd3_display_8" style="display:none">							
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">소카테고리선택</option>
					<option value="99140100">이색상품</option>		  
					<option value="99140200">케릭터용품</option>	
					<option value="99140300">주문제작/맞춤선물</option>
					<option value="99140400">포토앨범/박스/홀더</option>
					<option value="99140500">키덜트용품</option>
					<option value="99140600">디자인용품</option>
					<option value="99140700">아이디어소품</option>
					<option value="99140800">장식소품</option>						
					<option value="99140900">기타소품</option>		
				</select>
			</div>
		</td>
		<td align="right">
			<input type="button" value="등록(상품)" onclick="reg('item');" class="button">
			<input type="button" value="등록(이벤트)" onclick="reg('event');" class="button">
			<input type="button" value="excel출력" onclick="xmlprint(frm)" class="button">
		</td>
	</tr>
</form>	
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oip.FResultCount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="17">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
   		<td align="center">이미지</td>
		<td align="center">idx(수정)</td>
		<td align="center">상품코드</td>
		<td align="center">상품옵션</td>
		<td align="center">브랜드</td>
		<td align="center">상품명</td>
		<td align="center">가격</td>
		<td align="center">마진</td>
		<td align="center">텐재고수량</td>
		<td align="center">텐재고</td>
		<td align="center">품절</td>
		<td align="center">단종</td>
		<td align="center">옥션등록</td>
		<td align="center">옥션카테고리</td>
		<td align="center">비고</td>
    </tr>
	<% for i=0 to oip.FresultCount-1 %>
		<form action="/admin/auction/auction_process.asp" name="frmBuyPrc<%=i%>" method="get">			<!--for문 안에서 i 값을 가지고 루프-->
		<input type="hidden" name="mode">	
    	<tr align="center" bgcolor="#FFFFFF">
			<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			<td align="center">
			<img src="<%= oip.flist(i).FImageSmall %>" width="50" height="50">
			</td>
			<td align="center"><a href="javascript:edit('<%= oip.flist(i).idx %>','<%= oip.flist(i).ten_itemid %>')"><%= oip.flist(i).idx %></a><input type="hidden" name="idx" value="<%= oip.flist(i).idx %>"></td>
			<td align="center"><%= oip.flist(i).ten_itemid %><input type="hidden" name="itemid" value="<%= oip.flist(i).ten_itemid %>"></td>
			<td align="center"><%= oip.flist(i).ten_option %></td>
			<td align="center"><%= oip.flist(i).ten_makerid %></td>
			<td align="center"><%= oip.flist(i).ten_itemname %></td>
			<td align="center"><%= oip.flist(i).fsellcash %>원</td>
			<td align="center"><%= oip.flist(i).GetCalcuMarginRate %>%</td>
			<td align="center"><%= oip.flist(i).ten_jaego %></td>
			<td align="center"><% if oip.flist(i).ten_jaego >= 10 then
					response.write "Y"
				else 
					response.write "N"
				end if %></td>
			<td align="center">
				<% if oip.flist(i).IsSoldOut then %>
					<font color=red>품절</font>
    			<% end if %>
    		</td>
    		<td align="center">	
    		<% if oip.flist(i).Fdanjongyn="Y" then %>
			<font color="#33CC33">단종</font>
			<% elseif oip.flist(i).Fdanjongyn="S" then %>
			<font color="#33CC33">일시<br>품절</font>
			<% end if %>
			</td>
			<td align="center"><%= oip.flist(i).auction_isusing %></td>
			<td align="center">
			
			<% if oip.flist(i).auction_cate_code = "10010100" then
				response.write "노트/연습장"
			elseif oip.flist(i).auction_cate_code = "10010200" then
				response.write "클리어파일"
			elseif oip.flist(i).auction_cate_code = "10010300" then
				response.write "포스트잇/메모지"
			elseif oip.flist(i).auction_cate_code = "10010400" then
				response.write "화이트/수정용품"
			elseif oip.flist(i).auction_cate_code = "10010500" then
				response.write "클립/집게/홀더"
			elseif oip.flist(i).auction_cate_code = "10010600" then
				response.write "칼/가위/자"
			elseif oip.flist(i).auction_cate_code = "10010700" then
				response.write "스템플러/리무버"
			elseif oip.flist(i).auction_cate_code = "10010800" then
				response.write "풀/테이프"		
			elseif oip.flist(i).auction_cate_code = "10010900" then
				response.write "펀치"	
			elseif oip.flist(i).auction_cate_code = "10011000" then
				response.write "문구세트"
			elseif oip.flist(i).auction_cate_code = "10011100" then
				response.write "화방/제도용품"
			elseif oip.flist(i).auction_cate_code = "10011200" then
				response.write "문구용품기타"
				
			elseif oip.flist(i).auction_cate_code = "10030100" then
				response.write "접착식앨범"
			elseif oip.flist(i).auction_cate_code = "10030200" then
				response.write "포켓식앨범"
			elseif oip.flist(i).auction_cate_code = "10030300" then
				response.write "앨범기타"
			elseif oip.flist(i).auction_cate_code = "10040100" then
				response.write "만년필"
			elseif oip.flist(i).auction_cate_code = "10040200" then
				response.write "매직/네임펜/마카"
			elseif oip.flist(i).auction_cate_code = "10040301" then
				response.write "유성펜"
			elseif oip.flist(i).auction_cate_code = "10040302" then
				response.write "수성펜"
			elseif oip.flist(i).auction_cate_code = "10040400" then
				response.write "샤프/연필/색연필"
			elseif oip.flist(i).auction_cate_code = "10040500" then
				response.write "형광/사인펜"
			elseif oip.flist(i).auction_cate_code = "10040600" then
				response.write "제도용펜/특수펜"
			elseif oip.flist(i).auction_cate_code = "10040700" then
				response.write "필기구기타"
			
			elseif oip.flist(i).auction_cate_code = "10050100" then
				response.write "도장"
			elseif oip.flist(i).auction_cate_code = "10050200" then
				response.write "스탬프"
			elseif oip.flist(i).auction_cate_code = "10050300" then
				response.write "파일/바인더"
			elseif oip.flist(i).auction_cate_code = "10050400" then
				response.write "명함집/케이스"	
			elseif oip.flist(i).auction_cate_code = "10050500" then
				response.write "잉크/토너"
			elseif oip.flist(i).auction_cate_code = "10050600" then
				response.write "서류보관함"
			elseif oip.flist(i).auction_cate_code = "10050700" then
				response.write "칠판/보드"
			elseif oip.flist(i).auction_cate_code = "10050900" then
				response.write "사무용가구"																														
			elseif oip.flist(i).auction_cate_code = "10051000" then
				response.write "사무용품기타"
				
			elseif oip.flist(i).auction_cate_code = "10060101" then
				response.write "케릭터다이어리"
			elseif oip.flist(i).auction_cate_code = "10060102" then
				response.write "일러스트다이어리"
			elseif oip.flist(i).auction_cate_code = "10060103" then
				response.write "만년다이어리"	
			elseif oip.flist(i).auction_cate_code = "10060104" then
				response.write "핸드매이드다이어리"
			elseif oip.flist(i).auction_cate_code = "10060201" then
				response.write "스터디다이어리"
			elseif oip.flist(i).auction_cate_code = "10060202" then
				response.write "포토다이어리"
			elseif oip.flist(i).auction_cate_code = "10060301" then
				response.write "프랭클린다이어리"
			elseif oip.flist(i).auction_cate_code = "10060302" then
				response.write "시스템다이어리기타"
			elseif oip.flist(i).auction_cate_code = "99140700" then
				response.write "다이어리속지"
			elseif oip.flist(i).auction_cate_code = "10060500" then
				response.write "다이어리기타"
			elseif oip.flist(i).auction_cate_code = "10060100" then
				response.write "팬시다이어리"
			elseif oip.flist(i).auction_cate_code = "10060200" then
				response.write "기능성다이어리"
			elseif oip.flist(i).auction_cate_code = "10060300" then
				response.write "시스템다이어리"
													
			elseif oip.flist(i).auction_cate_code = "10070100" then
				response.write "계산기"	
			elseif oip.flist(i).auction_cate_code = "10071000" then
				response.write "사무기기기타"
			elseif oip.flist(i).auction_cate_code = "10090200" then
				response.write "편지/서류봉투"
			elseif oip.flist(i).auction_cate_code = "10090300" then
				response.write "견출지/라벨류"
			elseif oip.flist(i).auction_cate_code = "10090700" then
				response.write "편지지/엽서"
			elseif oip.flist(i).auction_cate_code = "10090800" then
				response.write "장부/양식/서식지"
			elseif oip.flist(i).auction_cate_code = "10090900" then
				response.write "종이류기타"
				
			elseif oip.flist(i).auction_cate_code = "99140100" then
				response.write "이색상품"
			elseif oip.flist(i).auction_cate_code = "99140200" then
				response.write "캐릭터용품"	
			elseif oip.flist(i).auction_cate_code = "99140300" then
				response.write "주문제작/맞춤선물"
			elseif oip.flist(i).auction_cate_code = "99140400" then
				response.write "포토앨범/박스/홀더"
			elseif oip.flist(i).auction_cate_code = "99140500" then
				response.write "키덜트용품"
			elseif oip.flist(i).auction_cate_code = "99140600" then
				response.write "디자인소품"		
			elseif oip.flist(i).auction_cate_code = "99140700" then
				response.write "아이디어소품"
			elseif oip.flist(i).auction_cate_code = "99140800" then
				response.write "장식소품"	
			elseif oip.flist(i).auction_cate_code = "99140900" then
				response.write "기타상품"										
			end if %><br>(<%= oip.flist(i).auction_cate_code %>)
			</td>
			<td align="center">
		<!--비고란구분시작 -->	
				<input type="button" value="삭제" onclick="DelMe(frmBuyPrc<%=i%>,'<%= oip.flist(i).idx %>');">
		<!--비고란구분끝 -->
			</td>
    	</tr>   
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="7" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
        	<% if oip.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oip.StartScrollPage-1 %>')">[pre]</a>
	   		<% else %>
	    		[pre]
	   		<% end if %>
	
	    	<% for i=0 + oip.StartScrollPage to oip.FScrollCount + oip.StartScrollPage - 1 %>
	    		<% if i>oip.FTotalpage then Exit for %>
		    		<% if CStr(page)=CStr(i) then %>
		    		<font color="red">[<%= i %>]</font>
		    		<% else %>
		    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		    		<% end if %>
	    	<% next %>
	
	    	<% if oip.HasNextScroll then %>
	    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    	<% else %>
	    		[next]
    		<% end if %>
		</td>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<iframe frameboarder=0 height=0 width=0 name="view" id="view"></iframe>