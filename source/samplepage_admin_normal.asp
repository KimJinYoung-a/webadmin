<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%

dim sellyn,usingyn

%>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 :
			&nbsp;
			카테고리 :
			<br>
			상품코드 :
			<input type="text" class="text" name="" value="" size="32"> (쉼표로 복수입력가능)
			&nbsp;
			상품명 :
			<input type="text" class="text" name="" value="" size="32" maxlength="32">
			<br>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	&nbsp;
	     	사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;
	     	단종
			<select class="select" name="">
	     	<option value='' selected>전체</option>
	     	<option value=''>단종</option>
	     	<option value=''>MD품절</option>
	     	<option value=''>일시품절</option>
	     	<option value=''>단종아님</option>
	     	</select>
	     	&nbsp;
	     	한정
			<select class="select" name="">
	     	<option value='' selected>전체</option>
	     	<option value=''>비한정</option>
	     	<option value=''>한정</option>
	     	<option value=''>한정(0)</option>
	     	</select>
	     	&nbsp;
	     	거래구분:<% drawSelectBoxMWU "usingyn", usingyn %>
	     	&nbsp;
	     	과세
			<select class="select" name="">
	     	<option value='' selected>전체</option>
	     	<option value=''>과세</option>
	     	<option value=''>면세</option>
	     	</select>
	     	&nbsp;
	     	할인
			<select class="select" name="">
	     	<option value='' selected>전체</option>
	     	<option value=''>할인</option>
	     	<option value=''>할인안함</option>
	     	</select>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="전체선택" onClick="">
			&nbsp;
			마진율 : <input type="text" class="text" name="" size="3" maxlength="5">
			<input type="button" class="button" value="선택상품적용" onClick="">
			&nbsp;
			<input type="button" class="button" value="저장" onClick="">
			각종 처리 및 액션이 들어가는 곳입니다.
		</td>
		<td align="right">
			<img src="/images/icon_star.gif" border="0">
			<img src="/images/icon_plus.gif" border="0">
			<img src="/images/icon_minus.gif" border="0">
			<img src="/images/icon_arrow_up.gif" border="0">
			<img src="/images/icon_arrow_down.gif" border="0">
			<img src="/images/icon_arrow_left.gif" border="0">
			<img src="/images/icon_arrow_right.gif" border="0">

			<img src="/images/question.gif" border="0">

			<img src="/images/btn_word.gif" border="0">
			<img src="/images/btn_excel.gif" border="0">
			<img src="/images/icon_word.gif" border="0">
			<img src="/images/icon_excel.gif" border="0">
			<img src="/images/icon_reload.gif" border="0">
			<img src="/images/icon_go.gif" border="0">


		</td>
	</tr>
</table>
<!-- 액션 끝 -->



<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b>350</b>
			&nbsp;
			페이지 : <b>1 / 20</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="150">adminColor("tablebar")</td>
    	<td width="100"><%= adminColor("tabletop") %></td>
      	<td width="100">항목3</td>
      	<td width="100">항목4</td>
      	<td width="100">항목5</td>
      	<td>비고</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tablebg") %>">
    	<td>adminColor("tablebg")</td>
    	<td><%= adminColor("tablebg") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("pink") %>">
    	<td>adminColor("pink")</td>
    	<td><%= adminColor("pink") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("green") %>">
    	<td>adminColor("green")</td>
    	<td><%= adminColor("green") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("sky") %>">
    	<td>adminColor("sky")</td>
    	<td><%= adminColor("sky") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("gray") %>">
    	<td>adminColor("gray")</td>
    	<td><%= adminColor("gray") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
    	<td>adminColor("dgray")</td>
    	<td><%= adminColor("dgray") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>&nbsp;</td>
    	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			[pre]
			<font color="red">1</font>
			2
			3
			4
			[next]
		</td>
	</tr>
</table>

<p>

<!-- 각종 포맷 시작 -->
>>각종 CSS
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="150">구분</td>
    	<td width="300">디스플레이</td>
    	<td width="100">class</td>
      	<td>비고</td>

    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>아이콘</td>
    	<td>
    		<input type="button" class="icon" value="#" onClick="">
    		<input type="button" class="icon" value="*" onClick="">
    		<input type="button" class="icon" value="1" onClick="">
    		<input type="button" class="icon" value="2" onClick="">
    		<input type="button" class="icon" value="@" onClick="">
    		<input type="button" class="icon" value=">" onClick="">
    		<input type="button" class="icon" value="711" onClick="">
    	</td>
      	<td>icon</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>일반버튼</td>
    	<td><input type="button" class="button" value="버튼" onClick=""></td>
      	<td>button</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>검색버튼</td>
    	<td><input type="button" class="button_s" value="검색" onClick=""></td>
      	<td>button_s</td>
      	<td>차후에 칼라 변경 예정(일반버튼과 구별)</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>관리자버튼</td>
    	<td><input type="button" class="button_auth" value="액션" onClick=""  ></td>
      	<td>button_s</td>
      	<td>관리자만 보이는 버튼입니다.</td>
    </tr>

    <tr align="center" bgcolor="#FFFFFF">
    	<td>input_text</td>
    	<td><input type="text" class="text" name="" ></td>
      	<td>text</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>input_text (readonly)</td>
    	<td><input type="text" class="text_ro" name="" readonly></td>
      	<td>text_ro</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>input_textarea</td>
    	<td><textarea class="textarea" name="" rows="2"></textarea></td>
      	<td>textarea</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>input_textarea (readonly)</td>
    	<td><textarea class="textarea_ro" name="" rows="2" readonly></textarea></td>
      	<td>textarea_ro</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>select</td>
    	<td>
    		<select class="select" name="">
    			<option value='' selected>전체</option>
	     		<option value=''>옵션01</option>
	     		<option value=''>옵션02</option>
	     		<option value=''>옵션03</option>
	     		<option value=''>옵션04</option>
	     	</select>
    	</td>
      	<td>select</td>
      	<td></td>
    </tr>

</table>




<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->