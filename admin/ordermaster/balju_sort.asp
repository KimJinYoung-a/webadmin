<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 출고지시서 정렬순서관리
' History : 2020.12.17 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/order/tenbalju.asp"-->

<%
dim midx, title , isusing, i, page
dim sqlshopinfo , c_shopdiv ,oshop ,j
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
	page = requestCheckVar(getNumeric(request("page")),10)
    midx = requestCheckVar(getNumeric(request("midx")),10)
	title = requestCheckVar(request("title"),128)
    isusing = requestCheckVar(request("isusing"),1)

if page="" then page=1

dim osort		
set osort = new CTenBalju
	osort.FPageSize = 50
	osort.FCurrPage = page
	osort.frectmidx = midx
	osort.frecttitle = title
	osort.frectisusing = isusing	
	osort.GetBaljusortList()

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function gobaljureg(midx){
	var popwin = window.open('/admin/ordermaster/balju_sort_reg.asp?midx='+midx,'addreg','width=1280,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//폼 전송
function gosubmit(page){
    frm.page.value=page;
	frm.submit();
}

function totalCheck(){
	var f = document.frmArr;
	var objStr = "check";
	var chk_flag = true;
	for(var i=0; i<f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(!f.elements[i].checked) {
				chk_flag = f.elements[i].checked;
				break;
			}
		}
	}

	for(var i=0; i < f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(chk_flag) {
				f.elements[i].checked = false;
			} else {
				f.elements[i].checked = true;
			}
		}
	}
}

function gobaljumasterdel() {
    if ($('input[name="check"]:checked').length == 0) {
        alert('선택 아이템이 없습니다.');
        return;
    }
    var ret = confirm('선택내용을 삭제 하시겠습니까?');
    if (ret) {
        frmArr.action="/admin/ordermaster/balju_sort_process.asp";
        frmArr.mode.value="baljumasterdel";
        frmArr.target="view";
        frmArr.submit();
    }
}

function gobaljumasterreg() {
    if ($('input[name="check"]:checked').length == 0) {
        alert('선택 아이템이 없습니다.');
        return;
    }
    var ret = confirm('선택내용을 저장 하시겠습니까?');
    if (ret) {
        frmArr.action="/admin/ordermaster/balju_sort_process.asp";
        frmArr.mode.value="baljumasterreg";
        frmArr.target="view";
        frmArr.submit();
    }
}

function CheckClick(identikey){
	var f = document.frmArr;

	for(var i=0; i<f.check.length; i++){
		if(f.check[i].value==identikey){
			f.check[i].checked=true;
			break;
		}
	}
}

</script>

<!-- 표 상단바 시작-->
<form name="frm" method="get" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<input type="hidden" name="mode">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a">
		<tr>
			<td>				
				* 번호 : <input type="text" name="midx" value="<%= midx %>" size=8 maxlength=10 >
		     	&nbsp;
                * 제목 : <input type="text" name="title" value="<%= title %>" size=40 maxlength=128 >
			</td>
		</tr>
		<tr>
			<td>
		     	* 사용여부 : <% drawSelectBoxisusingYN "isusing",isusing,"" %>
			</td>
		</tr>
			
		</table>	
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('1');">
	</td>
</tr>
</table>
</form>
<!-- 표 상단바 끝-->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="선택삭제" onClick="gobaljumasterdel();">
        <!--<input type="button" class="button" value="선택수정" onClick="gobaljumasterreg();">-->
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onClick="gobaljureg('');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<form name="frmArr" id="frmArr" method="post" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= osort.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= osort.FTotalPage %></b>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><input type="checkbox" name="ckall" id="ckall" onclick="totalCheck()"></td>
    <td>번호</td>
	<td>제목</td>
    <td>레코드등록수</td>
	<td>사용여부</td>
    <!--<td>정렬</td>-->
	<td>등록일</td>
    <td>최종수정일</td>
	<td>비고</td>
</tr>
<% if osort.FresultCount>0 then %>
<%
For i =0 To osort.fresultcount -1
%>
<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
	<td>
		<input type="checkbox" name="check" value="<%= osort.FItemList(i).fmidx %>" />
	</td>   
	<td>
		<%= osort.FItemList(i).fmidx %>
	</td>
	<td>
		<%= ReplaceBracket(osort.FItemList(i).ftitle) %>
	</td>
	<td>
		<%= osort.FItemList(i).frackcodecount %>	
	</td>	
	<td>
		<%= osort.FItemList(i).fisusing %>	
	</td>
	<!--<td>-->
        <!--<input type="text" name="sortno_<%'= osort.FItemList(i).fmidx %>" value="<%'= osort.FItemList(i).fsortno %>" onKeyup="CheckClick('<%'= osort.FItemList(i).fmidx %>')" size=8 maxlength=10 >-->
	<!--</td>-->
	<td>
		<%= osort.FItemList(i).fregdate %>
        <Br><%= osort.FItemList(i).fregadminid %>
	</td>
	<td>
        <% if osort.FItemList(i).flastupdate<>"" then %>
		    <%= osort.FItemList(i).flastupdate %>
        <% end if %>
        <% if osort.FItemList(i).flastadminid<>"" then %>
            <Br><%= osort.FItemList(i).flastadminid %>
        <% end if %>
	</td>
	<td>
		<input type="button" onclick="gobaljureg('<%=osort.FItemList(i).fmidx%>');" value="레코드등록/수정" class="button">
	</td>
</tr>
<%
Next
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if osort.HasPreScroll then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= osort.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + osort.StartScrollPage to osort.StartScrollPage + osort.FScrollCount - 1 %>
			<% if (i > osort.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(osort.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if osort.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<%
else
%>
<tr bgcolor="#FFFFFF" height="30">
	<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
</tr>
<%
End If
%>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
    <iframe id="view" name="view" src="" width="100%" height="300" frameborder="0" scrolling="no"></iframe>
<% else %>
    <iframe id="view" name="view" src="" width="100%" height="0" frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set osort = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
