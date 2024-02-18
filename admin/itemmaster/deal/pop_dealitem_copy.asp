<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/pop_dealitem_copy.asp
' Description :  딜 상품 복사 리스트
' History : 2023.01.10 정태훈
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	'변수선언
	Dim iCurrpage, iPageSize, iPerCnt, isResearch, sSdate, sEdate, intLoop, stext, dispCate, idx
	Dim oDeal, arrList, iTotCnt, iTotalPage, strTxt, sdiv, datediv, viewdiv, isusing, arrCate, maxDepth

	idx = requestCheckVar(Request("idx"),10)	'현재 페이지 번호
	dispCate	= requestCheckVar(Request("disp"),16) 		'전시 카테고리
	maxDepth = 2
	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격

	isusing 		= requestCheckVar(Request("isusing"),1)
	viewdiv 		= requestCheckVar(Request("viewdiv"),1)
	datediv 		= requestCheckVar(Request("datediv"),1)
	sdiv 		= requestCheckVar(Request("sdiv"),10)
	strTxt 		= requestCheckVar(Request("stext"),32)
	
	isResearch = requestCheckVar(Request("isResearch"),1)
	if isResearch ="" then isResearch ="0"
	'## 검색 #############################
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	'데이터 가져오기
	set oDeal = new ClsDeal
		oDeal.FCPage = iCurrpage		'현재페이지
		oDeal.FPSize = iPageSize		'한페이지에 보이는 레코드갯수
		oDeal.FSearchDateDiv 	= datediv	'검색일 구분
		oDeal.FSsDate 	= sSdate	'검색 시작일
		oDeal.FSeDate 	= sEdate	'검색 종료일
		oDeal.FSearchDiv 	= sdiv	'검색구분
		oDeal.FSeTxt 	= strTxt	'검색어
		oDeal.FSViewDiv 	= viewdiv	'유형 구분
		oDeal.FSIsUsing 	= isusing	'사용 구분
		oDeal.FSdispCate 	= dispCate	'전시카테고리 검색
 		arrList = oDeal.fnGetCopyItemDealList	'데이터목록 가져오기
 		iTotCnt = oDeal.FTotCnt	'전체 데이터  수
 	set oDeal = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
	function jsSearch(sType){
		var frm = document.frmEvt
		if (sType == "A"){
				frm.iSD.value = "";
				frm.iED.value = "";
				frm.eventstate.value = "";
				frm.sEtxt.value = "";
				frm.selC.value = "";
		}
		if(frm.sdiv.value=="itemid" && frm.stext.value!=""){
			if(isNaN(frm.stext.value)){
				alert("상품번호 검색은 숫자만 입력해주세요!");
				return false;
			}
		}

		frm.submit();
	}

	function jsCopyItem(dealcode){
        if(confirm("상품을 복사 하시겠습니까?")){
            $.ajax({
                type: "POST",
                url: "dodealitemcopy.asp",
                data: "mode=copy&idx=<%=idx%>&dealcode="+dealcode,
                cache: false,
                async: false,
                success: function(data) {
                    if(data.response=="ok") {
                        alert(data.message);
						opener.jsItemCopyAfter();
						self.close();
                    } else {
                        alert(data.message);
                    }
                }
            });
        }
	}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmEvt" method="get" onSubmit="return jsSearch('E');">
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">검색 조건</td>
	<td>
		<table>
		<tr>
			<td>
				검색어 : 
				<select name="sdiv" class="select">
					<option value="itemid"<% If sdiv="itemid" Then Response.write " selected" %>>딜상품코드</option>
					<option value="itemname"<% If sdiv="itemname" Then Response.write " selected" %>>상품명</option>
					<option value="register"<% If sdiv="register" Then Response.write " selected" %>>작성자</option>
					<option value="makerid"<% If sdiv="makerid" Then Response.write " selected" %>>브랜드아이디</option>
				</select>
				<input type="text" name="stext" size="50" value="<%=strTxt%>" onkeydown="if(event.keyCode==13) jsSearch('E');">
			</td>
		</tr>
		</table>
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>" align="center"><input type="button" class="button_s" value="검색" onClick="javascript:jsSearch('E');"></td>
</tr>
</form>
</table><br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="13">
			<table width="100%">
			<tr>
				<td>검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>딜상품코드</td>
		<td>카테고리</td>
		<td>유형</td>
		<td>딜상품명</td>
		<td>상품수</td>
	 </tr>
	 <% If isArray(arrList) Then %>
	 <% For intLoop = 0 To UBound(arrList,2) %>
	 <tr bgcolor="#FFFFFF" onclick="jsCopyItem(<%=arrList(0,intLoop)%>)">
		<td align="center"><%=arrList(1,intLoop)%></td>
		<td align="center">
		<%
			If arrList(12,intLoop) <> "" Then
			arrCate = Split(arrList(12,intLoop),"^^")
			If ubound(arrCate)>0 Then
			Response.write arrCate(0) & " > " & arrCate(1)
			Else
			Response.write arrCate(0)
			End If
			End If
		%>
		</td>
		<td align="center"><% If arrList(2,intLoop)="1" Then %>상시딜<% Else %>기간딜<% End If %></td>
		<td><%=arrList(5,intLoop)%></td>
		<td align="center">
			<%=arrList(15,intLoop)%>
		</td>
	 </tr>
	 <% Next %>
	 <% Else %>
	 <tr bgcolor="#FFFFFF">
		<td colspan="13" align="center" height="25">
			등록된 내용이 없습니다.
		</td>
	 </tr>
	 <% End If %>
	 <tr bgcolor="#FFFFFF">
		<td colspan="13" bgcolor="#FFFFFF" align="center">
			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,"" %>
		</td>
	 </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->