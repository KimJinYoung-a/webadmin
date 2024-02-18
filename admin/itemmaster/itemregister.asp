<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : 온라인상품등록
' History : 서동석 생성
'			2017.11.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

dim i,j, designer, rentalItemFlag
'==============================================================================
Sub SelectBoxDesignerItem()
   dim query1 
   %><select name="designer" class="select" onchange="TnDesignerNMargineAppl(this.value);">
     <option value=''>-- 업체선택 --</option><%
   query1 = " select userid,socname_kor,defaultmargine, maeipdiv, IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType from [db_user].[dbo].tbl_user_c"
   query1 = query1 + " where isusing='Y'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           response.write("<option value='"&rsget("userid")& "," & rsget("defaultmargine") & "," & rsget("maeipdiv") & "," & rsget("defaultFreeBeasongLimit") & "," & rsget("defaultDeliverPay") & "," & rsget("defaultDeliveryType") & "'>" & rsget("userid") & "  [" & replace(db2html(rsget("socname_kor")),"'","") & "]" & "</option>")
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

'// 렌탈 상품은 일단 테스트로 특정 유저만 노출함
If C_ADMIN_AUTH Then
	rentalItemFlag = true
Else
	rentalItemFlag = true
End If
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript' src="/js/ckeditor/ckeditor.js"></script>
<script type="text/javascript">
<!-- #include file="./itemregister_javascript.asp"-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>상품등록</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>신상품을 등록합니다.</b>
			<!--
            <br>- 매주 화요일 까지 등록하셔야 수요일에 승인 후 업데이트 됩니다.
            <br>- 설명이나 내용이 부족한 경우 승인 거부될 수 있습니다.
            -->
			<br>- 기본틀생성을 이용하여 빠르게 상품을 등록할수 있습니다.
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>
<input type="button" class="button" value="기본틀생성" onClick="UseTemplate();"><br><br>

<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/itemregisterWithImage_process.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="designerid">
<input type="hidden" name="defaultmargin">
<input type="hidden" name="defaultmaeipdiv">
<input type="hidden" name="defaultFreeBeasongLimit">
<input type="hidden" name="defaultDeliverPay">
<input type="hidden" name="defaultDeliveryType">
<input type="hidden" name="DFcolorCD" value="">
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="itemoptioncode3">

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>

</table>
<!-- 표 상단바 끝-->


<!-- 1.일반정보 --> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>1.일반정보</strong></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">브랜드ID :</td>
	<td bgcolor="#FFFFFF" colspan="3"><% NewDrawSelectBoxDesignerChangeMargin "makerid", designer, "marginData", "TnDesignerNMargineAppl2" %></td>
	<% 'SelectBoxDesignerItem %> <!--(사용업체만 표시됩니다)-->
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="text" name="itemname" maxlength="64" size="50" class="text" id="[on,off,off,off][상품명]">&nbsp;
	</td>
</tr>
<!-- 업체등록시에는 사용안함(MD만 등록가능) -->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품카피 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="designercomment" size="60" maxlength="128" class="text" id="[off,off,off,off][상품카피]"><br>
	</td>
</tr>
</table>

<!-- 2.구분 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>2.구분</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
<input type="hidden" name="cd1" value="">
<input type="hidden" name="cd2" value="">
<input type="hidden" name="cd3" value="">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="재고/매출 등의 관리 카테고리" style="cursor:help;">관리 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="text" name="cd1_name" value="" id="[on,off,off,off][카테고리]" size="20" readonly class="text_ro">
		<input type="text" name="cd2_name" value="" id="[on,off,off,off][카테고리]" size="20" readonly class="text_ro">
		<input type="text" name="cd3_name" value="" id="[on,off,off,off][카테고리]" size="20" readonly class="text_ro">
		
		<input type="button" value="카테고리 선택" class="button" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="프론트에 진열될 카테고리" style="cursor:help;">전시 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td id="lyrDispList"><table class="a" id="tbl_DispCate"></table></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table> 
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	 </td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품구분 :</td>
	<td bgcolor="#FFFFFF" >
		<label><input type="radio" name="itemdiv" value="01" checked onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">일반상품</label>
		<br>
		<label><input type="radio" name="itemdiv" value="06" onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">주문 제작상품</label>
		<input type="checkbox" name="reqMsg" value="10" onClick="checkItemDiv(this);">주문제작 문구 필요<font color=red>(주문시 이니셜등 제작문구가 필요한경우 체크)</font>
		<br>
		<!--업체등록시에는 사용안함(MD만 등록가능) -->
		<label><input type="radio" name="itemdiv" value="08" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">티켓상품</label>
		<label><input type="radio" name="itemdiv" value="09" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">Present상품</label>
		<label><input type="radio" name="itemdiv" value="11" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">상품권상품</label>
		<% If rentalItemFlag Then %>
			<label><input type="radio" name="itemdiv" value="30" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">렌탈상품</label>
		<% End If %>
		<label><input type="radio" name="itemdiv" value="23" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">B2B상품</label>
		<label><input type="radio" name="itemdiv" value="17" onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">마케팅전용상품</label>
	</td>
    <td bgcolor="#FFFFFF">
        <div id="lyRequre" style="display:none;padding-left:22px;">
		예상제작소요일 <input type="text" name="requireMakeDay" value="0" size="2" class="text" id="[off,on,off,off][예상제작소요일]">일
		<font color="red">(상품발송전 상품제작 기간)</font>
		</div>
	</td>
</tr>
<!-- 업체등록시에는 사용안함(MD만 등록가능) -->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">텐바이텐 독점구분 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="tenOnlyYn" value="Y" >독점상품</label>
		<label><input type="radio" name="tenOnlyYn" value="N" checked>일반상품</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">구매 가능 연령 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="adultType" value="0" checked>전체연령</label>
		<label><input type="radio" name="adultType" value="1" >구매시성인인증</label>
		<label><input type="radio" name="adultType" value="2" >미성년 조회 불가</label>
	</td>
</tr>
</table>

<!-- 3.가격정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>3.가격정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">마진 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][마진]">%
		<input type="button" value="공급가 자동계산" class="button" onclick="CalcuAuto(itemreg);">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">판매가(소비자가) :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="sellcash" maxlength="16" size="12" class="text" id="[on,on,off,off][소비자가]" onKeyup="CalcuAuto(itemreg);">원
		<input type="hidden" name="sellvat">
	</td>
	<td width="15%" bgcolor="#DDDDFF">공급가 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="buycash" maxlength="16" size="12" class="text" id="[on,on,off,off][공급가]" >원
		(<b>부가세 포함가</b>)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">마일리지 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="mileage" maxlength="32" size="10" id="[on,on,off,off][마일리지]" value="0" ReadOnly > (판매가의 1%)
	</td>
	<td width="15%" bgcolor="#DDDDFF">과세, 면세 여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="vatinclude" value="Y" checked onclick="TnGoClear(this.form);CalcuAuto(itemreg);">과세</label>
		<label><input type="radio" name="vatinclude" value="N" onclick="TnGoClear(this.form);CalcuAuto(itemreg);">면세</label>
	</td>
</tr>
</table>

<!-- 4.관리정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>4.관리정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="itemid" value="" size="20" class="text_ro" readonly id="[off,off,off,off][상품코드]">
	    (상품등록 완료시 부여됩니다.)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="상품 상세 속성" style="cursor:help;">상품속성 :</td>
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3"></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">업체상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="upchemanagecode" value="" size="20" maxlength="32" class="text" id="[off,off,off,off][업체상품코드]">
	    (업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">ISBN :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		ISBN 13 <input type="text" name="isbn13" class="text" value="" size="13" maxlength="13">
		/ 부가기호 <input type="text" name="isbn_sub" class="text" value="" size="5" maxlength="5"><br />
		ISBN 10 <input type="text" name="isbn10" class="text" value="" size="10" maxlength="10"> (Optional)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">연관상품등록 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="relateItems" value="" size="40" class="text" id="[off,off,off,off][연관상품]">
	    (연관상품은 최대 6개까지 등록가능, 상품번호를 콤마(,)로 구분하여 입력)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">판매여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="sellyn" value="Y">판매함</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="N" checked>판매안함</label>
	</td>
	<td width="15%" bgcolor="#DDDDFF">사용여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="isusing" value="Y" onclick="TnChkIsUsing(this.form)">사용함</label>&nbsp;&nbsp;
		<label><input type="radio" name="isusing" value="N" onclick="TnChkIsUsing(this.form)">사용안함</label>
	</td>
</tr>
</table>

<!-- 5.기본정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>5.기본정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">제조사 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="makername" maxlength="32" size="25" class="text" id="[on,off,off,off][제조사]">&nbsp;(제조업체명)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">원산지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	 <p> 
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="0" checked onClick="jsSetArea(this.value);"> 식품 외</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="1"  onClick="jsSetArea(this.value);"> 농축산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="2"  onClick="jsSetArea(this.value);"> 수산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="3"  onClick="jsSetArea(this.value);"> 축산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="4"  onClick="jsSetArea(this.value);"> 농수산가공품</span>
	  </p>
	  <p><input type="text" name="sourcearea" maxlength="64" size="64" class="text" id="[on,off,off,off][원산지]" /></p>
	  <div id="dvArea0" style="display:;">
	  <p><strong>ex: 한국, 중국, 중국OEM, 일본 등 </strong></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea1" style="display:none;">
	  <p><strong>국내산 :</strong> 국산, 국내산 또는 시·도명, 시·군명(대한민국, 한국X)  <span style="margin-right:10px;">ex. 쌀(국산)</span></BR>
	   <strong>수입산 :</strong> 통관시의 수입국가명 <span style="margin-right:10px;">ex. 곶감(중국산)</span></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea2" style="display:none;">
	  <p><strong>국내산 :</strong> 국산,국내산 또는 연근해산(양식 수산물은 시·군명 가능)   <span style="margin-right:10px;">ex. 갈치(국산), 오징어(연근해산)</span> </BR>
	  	<strong>원양산 :</strong> 원양산 또는 원양산(해역명)   <span style="margin-right:10px;">ex. 참치[원양산(대서양)]</span> </BR>
	    <strong>수입산 :</strong> 통관시의 수입국가명 <span style="margin-right:10px;">ex. 농어(중국산)</span></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea3" style="display:none;">
	  <p>소고기의 경우 식육의 종류(한우/육우/젖소구분) 및 원산지   <span style="margin-right:10px;">ex. 쇠고기(횡성산 한우), 쇠고기(호주산)</span></BR>
	  - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea4" style="display:none;">
	  <p><strong>98%이상 원료가 있는 경우:</strong>  한가지 원료만 표시 가능    <span style="margin-right:10px;">ex. 쇠고기(미국산)</span> </BR>
	  	<strong>복합 원료를 사용한 경우:</strong> 혼합비율이 높은 순으로 2개 국가   <span style="margin-right:10px;">ex. 고추장[밀가루(미국산),고춧가루(국내산)]</span></BR>
	  - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div> 
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품무게 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" id="[on,off,off,off][상품무게]" style="text-align:right" value="0">g &nbsp;(그램단위로 입력, ex:1.5kg→ 1500) / 해외배송시 배송비 산출을 위한 것이므로 정확히 입력.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">검색키워드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="keywords" maxlength="250" size="60" class="text" id="[on,off,off,off][검색키워드]">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)
	</td>
</tr>
</table>

<!-- 5-1.품목상세정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- 품목상세정보 </strong> &nbsp;<font color=gray>상품정보제공고시 관련 법안 추진에 따라 아래 내용을 정확히 입력해주시기 바랍니다.</font></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목선택 :</td>
	<td bgcolor="#FFFFFF">
		<% DrawInfoDiv "infoDiv", "", " onchange='chgInfoDiv(this.value);'" %>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:none">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목내용 :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList"></td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">상품상세페이지에 내용이 포함 되어있더라도 정확히 입력바랍니다. 부정확하거나 잘못된 정보 입력시, 그에 대한 책임을 물을 수도 있습니다.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsource" maxlength="64" size="50" class="text">&nbsp;(ex:플라스틱,비즈,금,...)
	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsize" maxlength="64" size="50" class="text">
		<select name="unit" class="select">
		<option value="">직접입력</option>
		<option value="mm">mm</option>
		<option value="cm" selected>cm</option>
		<option value="m²">m²</option>
		<option value="km">km</option>
		<option value="m²">m²</option>
		<option value="km²">km²</option>
		<option value="ha">ha</option>
		<option value="m³">m³</option>
		<option value="cm³">cm³</option>
		<option value="L">L</option>
		<option value="g">g</option>
		<option value="Kg">Kg</option>
		<option value="t">t</option>
		</select>
		&nbsp;(ex:7.5x15(cm))
		</td>
</tr>
</table>
<!-- 5-2.안전인증정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- 안전인증정보</strong></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">
		안전인증대상 :
		<input type="button" value="안전인증 필수 품목 확인" onclick="jsSafetyPopup();" class="button" />
	</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
		<tr align="left" height="30">
			<td bgcolor="#FFFFFF">
				<label><input type="radio" name="safetyYn" value="Y" checked onclick="chgSafetyYn(document.itemreg)" /> 대상</label>
				<label><input type="radio" name="safetyYn" value="N" onclick="chgSafetyYn(document.itemreg)" /> 대상아님</label>
				<label><input type="radio" name="safetyYn" value="I" onclick="chgSafetyYn(document.itemreg)" /> 상품설명에 표기</label>
				<label><input type="radio" name="safetyYn" value="S" onclick="chgSafetyYn(document.itemreg)" /> 안전기준준수</label>
				<input type="hidden" name="auth_go_catecode" id="auth_go_catecode" value="">
				<input type="hidden" name="real_safetydiv" id="real_safetydiv" value="">
				<input type="hidden" name="real_safetynum" id="real_safetynum" value="">
				<input type="hidden" name="real_safetyidx" id="real_safetyidx" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxSafetyDivCode "safetyDiv", "", "Y", "" %>
				인증번호 <input type="text" name="safetyNum" id="[off,off,off,off][안전인증 인증번호]" size="35" maxlength="25" value="" />
				<input type="button" id="safetybtn" value="추   가" onclick="jsSafetyAuth();" class="button">
				<input type="hidden" name="issafetyauth" id="issafetyauth" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<div id="safetyDivList"></div>
				<div id="safetyYnI" style="display:none;">
					<font color="blue">상품 설명에 표기(표기대상 상품인경우 상품 상세 페이지에 인증번호와 모델명, KC 마크를 꼭 표기해주세요.)</font>
				</div>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td bgcolor="#FFFFFF" colspan=2>
		* 인증정보를 입력 안 하거나, 잘못된 인증정보를 입력한 경우 발견 <strong><font color='red'>즉시 판매정지 또는 삭제</font></strong> 됩니다.<br>
		* <strong><font color='red'>안전기준준수</font></strong> 대상일경우 인증번호가 없으며, KC마크를 표시하지 않아야 됩니다.<br>
		* 입력한 인증정보는 제품안전정보센터에서 제공된 정보를 기준으로 조회되며, <strong><font color='red'>검증되지 않은 정보는 등록이 불가</font></strong>능합니다.<br>
		* 정상적인 인증정보를 입력했음에도 불구하고 등록이 안될경우에 "상품설명에 표기"로 설정이 가능하며, 상품 상세 페이지에 모델명과 표기대상 상품인경우 인증번호,KC마크를 표기해야 합니다.<br>
		* 안전인증정보 관련 문의는 홈페이지(<u><a href="http://safetykorea.kr" target="_blank">http://safetykorea.kr</a></u>)로 확인해 주시기 바랍니다.
	</td>
</tr>
</table>

<!-- 6.배송정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>6.배송정보</strong>
        </td>
        <td align="right">
        	<input type="button" class="button" value="계약조건으로 세팅" onclick="TnAutoChkDeliver()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">매입특정구분 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="mwdiv" value="M" checked onclick="TnCheckUpcheYN(this.form);">매입</label>
		<label><input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);">특정</label>
		<label><input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);">업체배송</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송구분 :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverytype" value="1" checked  onclick="TnCheckUpcheDeliverYN(this.form);">텐바이텐배송</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);">업체(무료)배송</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);">텐바이텐무료배송</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);">업체조건배송(개별 배송비부과)</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);">업체착불배송</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송방법 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverfixday" value="" checked onclick="TnCheckFixday(this.form)">택배(일반)</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="X" onclick="TnCheckFixday(this.form)">화물</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="C" disabled onclick="TnCheckFixday(this.form)">플라워지정일</label>
		<label><input type="radio" name="deliverfixday" value="G" onclick="TnCheckFixday(this.form)">해외직구</label>
		<label><input type="radio" name="deliverfixday" value="L" disabled onclick="TnCheckFixday(this.form)">클래스</label>
		<span id="lyrFreightRng" style="display:none;">
			<br />&nbsp;
			반품/교환 시 화물배송 비용(편도) :
			최소 <input type="text" name="freight_min" class="text" size="6" value="0" style="text-align:right;">원 ~
			최대 <input type="text" name="freight_max" class="text" size="6" value="0" style="text-align:right;">원
		</span>
		<br>&nbsp;<font color="red">(플라워 상품인 경우만 수도권배송, 서울배송, 플라워지정일 옵션이 사용가능합니다.)</font>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송지역 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverarea" value="" checked>전국배송</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="C" disabled >수도권배송</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="S" disabled >서울배송</label>
		<label><input type="checkbox" name="deliverOverseas" value="Y" checked title="해외배송은 상품무게가 입력이 돼야 완료됩니다.">해외배송</label>
	</td>
</tr>
<input type="hidden" name="pojangok" value="Y">
</table>

<!-- 7.옵션정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>7.옵션정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">옵션구분 :</td>
	<td width="85%" bgcolor="#FFFFFF">
		<label><input type="radio" name="useoptionyn" value="Y" onClick="TnCheckOptionYN(this.form);">옵션사용함</label>&nbsp;&nbsp;
		<label><input type="radio" name="useoptionyn" value="N" onClick="TnCheckOptionYN(this.form);" checked>옵션사용안함</label>
	</td>
</tr>
<!----- 옵션구분 DIV ----->
<tr id="opttype" style="display:none" height="40">
    <td width="15%" bgcolor="#DDDDFF">옵션 구분  :</td>
    <td width="85%" bgcolor="#FFFFFF">
        <label><input type="radio" name="optlevel" value="1" onClick="TnCheckOptionYN(this.form);" checked >단일 옵션 (옵션 구분 1개)</label>
        <label><input type="radio" name="optlevel" value="2" onClick="TnCheckOptionYN(this.form);" >이중 옵션 (옵션 구분 최대 3개)</label><!--<font color="blue">※ 매입특정구분이 업체배송인 경우만 선택가능합니다.</font> //2016.05.19 정윤정 삭제--> 
    </td>
</tr>
<!----- 단일 옵션 DIV ----->
<tr id="optlist" style="display:none" height="30">
    <td width="15%" bgcolor="#DDDDFF">옵션 설정 :</td>
  	<td width="85%" bgcolor="#FFFFFF">
      	<table width="500" border="0" cellspacing="0" cellpadding="0" class="a" >
      	<tr>
      	    <td width="100">옵션 구분명 :</td>
      	    <td width="400"><input type="text" name="optTypeNm" value="" size="20" maxlength="20" class="text" id="[off,off,off,off][옵션 구분명]"></td>
      	</tr>
      	<tr>
      	    <td colspan="2">
              <select multiple name="realopt" class="select" style="width:400px;height:120px;"></select>
            </td>
        </tr>
        <tr>
            <td colspan="2">
              <input type="button" value="기본옵션추가" name="btnoptadd" class="button" onclick="popNormalOptionAdd();" >
              <input type="button" value="전용옵션추가" name="btnetcoptadd" class="button" onclick="popEtcOptionAdd();">
              <input type="button" value="선택옵션삭제" name="btnoptdel" class="button" onclick="delItemOptionAdd()" >
              <br><br>
              - 기본옵션추가 : 색상, 사이즈등 기본적으로 정의된 옵션을 추가 하실 수 있습니다.<br>
              - 전용옵션추가 : 기본옵션에 정의되지 않은 상품전용옵션을 지정하실 수 있습니다.<br>
              - 선택옵션삭제 : 선택된 옵션을 삭제합니다.<br>
              - 주의사항 : 한번 저장된 옵션은 <font color=red>삭제가 불가능</font>합니다.<br>
              <br>
            </td>
        </tr>
        </table>
  	</td>
</tr>
<%
dim iMaxCols : iMaxCols = 3
dim iMaxRows : iMaxRows = 9
%>
<!----- 멀티 옵션 DIV ----->
<tr id="optlist2" style="display:none" height="30">
    <td width="15%" bgcolor="#DDDDFF">옵션설정 :</td>
    <td width="85%" bgcolor="#FFFFFF">
        <table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
        <tr align="center"  bgcolor="#DDDDFF">
            <td width="100">옵션구분명</td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="text" name="optionTypename<%= j+1 %>" value="" size="18" maxlength="20" class="text" id="[off,off,off,off][옵션 구분명<%= j %>]">
            </td>
            <% Next %>
            <td width="80">(등록예시)<br>색상</td>
            <td width="80">(등록예시)<br>사이즈</td>
        </tr>
        <tr height="2" bgcolor="#FFFFFF">
            <td colspan="6"></td>
        </tr>
        <% for i=0 to iMaxRows-1 %>
        <tr align="center"  bgcolor="#FFFFFF">
            <td>옵션명 <%= i+1 %></td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="hidden" name="itemoption<%= j+1 %>" value="">
                <input type="text" name="optionName<%= j+1 %>" size="18" maxlength="20" class="text" id="[off,off,off,off][옵션명<%= i %><%= j %>]">
            </td>
            <% next %>
            <td>
                <% if i=0 then %>
                빨강
                <% elseif i=1 then %>
                파랑
                <% elseif i=2 then %>
                노랑
                <% elseif i=3 then %>
                베이지
                <% end if %>
            </td>
            <td>
                <% if i=0 then %>
                XL
                <% elseif i=1 then %>
                L
                <% elseif i=2 then %>
                S
                <% end if %>
            </td>
        </tr>
        <% next %>
        </table>
     </td>
</tr>

<!----- 기본 색상 DIV ----->
<tr id="lyDFColor" height="30" style="display:;">
	<td colspan="2" bgcolor="#FFFFFF" style="padding:0px;">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="0">
		<tr>
			<td width="15%" bgcolor="#DDDDFF">기본 색상선택 :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-left:1px solid <%= adminColor("tablebg") %>;"><%=FnSelectColorBar("",25)%></td>
		</tr>
		<tr>
			<td width="15%" rowspan="2" bgcolor="#DDDDFF" style="border-top:1px solid <%= adminColor("tablebg") %>;">색상별 상품이미지 :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
				<input type="file" size="40" name="imgDFColor" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text">
				<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgDFColor, 40, 1000, 1000)"> (선택,1000X1000,<b><font color="red">jpg</font></b>)
			</td>
		</tr>
		<tr>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
		      - 색상별 이미지는 별도로 등록을 하지않으면 상품 기본이미지가 사용됩니다.
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<!-- 8.한정정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>8.한정정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td width="15%" bgcolor="#DDDDFF" rowspan="2">한정판매구분 :</td>
	<td width="35%" bgcolor="#FFFFFF">   
		<label><input type="radio" name="limityn" value="N" onClick="this.form.limitno.readOnly=true; this.form.limitno.value=''; this.form.limitno.className='text_ro';document.all.dvDisp.style.display = 'none'; this.form.limitdispyn[0].checked = false; this.form.limitdispyn[1].checked = true;" checked>비한정판매</label>&nbsp;&nbsp;
		<label><input type="radio" name="limityn" value="Y" onClick="this.form.limitno.readOnly=false; this.form.limitno.className='text';document.all.dvDisp.style.display = '';">한정판매</label>
		<div id="dvDisp" style="display:none;" >
			&nbsp;-> 한정노출여부: 
			<input type="radio" name="limitdispyn" value="Y">노출 
			<input type="radio" name="limitdispyn" value="N" checked>비노출
		</div>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">한정수량 :</td>
	<td width="35%" bgcolor="#FFFFFF" >
		<input type="text" name="limitno" maxlength="32" size="8" readonly class="text_ro" id="[off,on,off,off][한정수량]">(개)
	</td>
</tr>
<tr>
	<td colspan="3" bgcolor="#FFFFFF"><font color="red">** 옵션이 있는경우 옵션별로 한정수량이 일괄 설정됩니다.(개별설정은 등록후 수정가능)</font></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">최소/최대 판매수 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		최소
		<input type="text" name="orderMinNum" maxlength="5" size="5" class="text" id="[off,on,off,off][최소판매수]" value="1">
		/ 최대
		<input type="text" name="orderMaxNum" maxlength="5" size="5" class="text" id="[off,on,off,off][최대판매수]" value="100">
		(한 주문에 판매 제한 수)
	</td>
</tr>
</table>

<!-- 9.상품설명 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>9.상품설명</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 설명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<!--<label><input type="radio" name="usinghtml" value="N"  >일반TEXT</label>
		<label><input type="radio" name="usinghtml" value="H" checked>TEXT+HTML</label>
		<label><input type="radio" name="usinghtml" value="Y">HTML사용</label>
		<br>
		-->
		<input type="hidden" name="usinghtml" value="Y" />
		<textarea name="itemcontent" rows="18" class="textarea" style="width:100%" id="[on,off,off,off][상품설명]"></textarea>
		<script>
		//
		window.onload = new function(){
			var itemContEditor = CKEDITOR.replace('itemcontent',{
				height : 450,
				// 업로드된 파일 목록
				//filebrowserBrowseUrl : '/browser/browse.asp',
				// 파일 업로드 처리 페이지
				filebrowserImageUploadUrl : '<%= ItemUploadUrl %>/linkweb/items/itemEditorContentUpload.asp'
			});
			itemContEditor.on( 'change', function( evt ) {
			    // 입력할 때 textarea 정보 갱신
			    document.itemreg.itemcontent.value = evt.editor.getData();
			});
		}
		</script>
		<div class="lpad10">
			※ 상품상세 영역의 최대 넓이(폭)는 1,000px입니다.
		</div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">아이템 동영상 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="90" id="[off,off,off,off][아이템동영상]"></textarea>
	    <br>※ Youtube, Vimeo 동영상만 가능(Youtube : 소스코드값 입력, Vimeo : 임베딩값 입력)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문시 유의사항 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="ordercomment" rows="5" cols="90" class="textarea" id="[off,off,off,off][유의사항]"></textarea><br>
		<font color="red">특별한 배송기간이나 주문시 확인해야만 하는 사항</font>을 입력하시면 고객불만이나 환불을 줄일수 있습니다.
	</td>
</tr>
</table>

<!-- 10.이미지정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>10.이미지정보</strong>
		<br>- 텐바이텐에서 이미지를 등록할 경우에는 필수항목인 기본이미지만 입력하시기 바랍니다.
		<br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
		<br>&nbsp;&nbsp;(이미지사이즈나 <font color=red>가로세로폭의 사이즈</font>를 규격에 넘지 않게 등록해주세요. 규격초과시 등록이 되지 않습니다.)
		<br>- <font color=red>포토샾에서 Save For Web으로, Optimize체크, 압축율 80%이하</font>로 만드신 후 올려주시기 바랍니다.
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">기본이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgbasic" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
      <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgbasic,40, 1000, 1000)"> (<font color=red>필수</font>,1000X1000,<b><font color="red">jpg</font></b>)
      <!-- // 사용암함 // <br><input type="checkbox" name="regimg"> 가등록이미지사용 - 이미지를 <font color=red>나중에 등록</font>할경우에는 가등록이미지사용을 체크하세요.-->
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">흰배경(누끼)이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgmask" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
      <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgmask,40, 1000, 1000)"> (선택,1000X1000,<b><font color="red">jpg</font></b>)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF" title="텐바이텐에서만 업로드 가능한 기본이미지 입니다.">텐바이텐기본이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgtenten" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
      <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgtenten,40, 1000, 1000)"> (선택,1000X1000,<b><font color="red">jpg</font></b>)
  	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd2, 40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd3, 40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지4 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd4, 40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지5 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd5, 40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
   	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>※ 기존의 제품설명이미지는 사용하지 않고 상품설명이미지를 사용합니다. 기존에 등록된 제품설명이미지는 사용은 하되 추가 수정은 되지않고 삭제만 됩니다.</strong></font>
 	</td>
 </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">PC상품설명이미지 #1 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#1 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 800, 1600)"> (선택,800X1600, Max 800KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">PC상품설명이미지 #2 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#2 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 800, 1600)"> (선택,800X1600, Max 800KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">PC상품설명이미지 #3 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#3 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 800, 1600)"> (선택,800X1600, Max 800KB,jpg,gif)
  	</td>
  </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td bgcolor="#FFFFFF">
      <input type="button" value="PC상품설명이미지추가" class="button" onClick="InsertImageUp()">
  	</td>
  </tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="MobileimgIn">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">모바일상품상세이미지 #1 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addmoblieimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#1 이미지지우기" class="button" onClick="ClearImage2(this.form.addmoblieimgname[0],40, 640, 1200)"> (선택,640X1200, Max 400KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">모바일상품상세이미지 #2 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addmoblieimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#2 이미지지우기" class="button" onClick="ClearImage2(this.form.addmoblieimgname[1],40, 640, 1200)"> (선택,640X1200, Max 400KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">모바일상품상세이미지 #3 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addmoblieimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="#3 이미지지우기" class="button" onClick="ClearImage2(this.form.addmoblieimgname[2],40, 640, 1200)"> (선택,640X1200, Max 400KB,jpg,gif)
  	</td>
  </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>※ 모바일 상품상세 이미지는 앞으로 이 영역으로 대체 됩니다. html은 사용하지 않을 예정이오니 이쪽으로 업로드 해주시기 바랍니다.<br>※ 모바일 상품상세에는 이미지를 잘라서 올려주시기 바랍니다.</strong></font>
 	</td>
 </tr>
  <tr align="left">
  	<td bgcolor="#FFFFFF">
      <input type="button" value="모바일상품상세이미지추가" class="button" onClick="InsertMobileImageUp()">
  	</td>
  </tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="저장하기" class="button" onClick="SubmitSave()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
</form>
<% if (application("Svr_Info")	= "Dev") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<script type="text/javascript">
	// 안전인증체크. 전안법
	jsSafetyCheck('','');
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->