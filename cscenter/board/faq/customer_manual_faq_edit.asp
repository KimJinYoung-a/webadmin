<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객안내FAQ
' Hieditor : 2019.10.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/board/faq/customer_manual_faq_cls.asp"-->

<%
Dim ofaq,menupos, i,fidx,gubun,contents,solution,isusing,regdate,lastupdate,lastadminid
dim omanager ,managershopyn, manualtype
	fidx = requestCheckVar(getNumeric(request("fidx")),10)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)

manualtype="customer_faq"

set ofaq = new cfaq_list
	ofaq.frectfidx = fidx
	ofaq.frectmanualtype = manualtype

	'//수정시에만 쿼리
	if fidx <> "" then		
		ofaq.Getcustomer_manual_faq_one()

		if ofaq.ftotalcount >0 then			
            fidx = ofaq.FOneItem.ffidx
            gubun = ofaq.FOneItem.fgubun
            contents = ReplaceBracket(ofaq.FOneItem.fcontents)
            solution = ReplaceBracket(ofaq.FOneItem.fsolution)
            isusing = ofaq.FOneItem.fisusing
            regdate = ofaq.FOneItem.fregdate
            lastupdate = ofaq.FOneItem.flastupdate
            lastadminid = ofaq.FOneItem.flastadminid
		end if
	end if

if isusing="" then isusing="Y"
%>

<script type="text/javascript">

	function fnfaq_write(){
		if (frm.gubun.value=='') {
			alert('구분 선택해 주세요');
			frm.gubun.focus();
			return;
		}
		if (frm.contents.value=='') {
			alert('문의내용을 입력해 주세요');
			frm.contents.focus();
			return;
		}
		if (frm.solution.value=='') {
			alert('처리방법을 입력해 주세요');
			frm.solution.focus();
			return;
		}
		if (frm.isusing.value=='') {
			alert('사용여부를 선택해 주세요');
			frm.isusing.focus();			
			return;
		}
		
		frm.action='/cscenter/board/faq/customer_manual_faq_process.asp';
		frm.mode.value = "faqreg";
		frm.submit();
	}
	
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		고객안내용FAQ 등록/수정
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">

<% if fidx<>"" then %>
    <tr bgcolor="#FFFFFF">
        <td align="center">번호<br></td>
        <td>
            <%= fidx %><input type="hidden" name="fidx" value="<%= fidx %>">
        </td>
    </tr>
<% else %>
    <input type="hidden" name="fidx" value="<%= fidx %>">
<% end if %>

<tr bgcolor="#FFFFFF">
	<td align="center">구분</td>
	<td>
		<% Drawcustomerfaqgubun "gubun",gubun,"" %>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td align="center">문의내용</td>
	<td>
		<textarea name="contents" rows="7" cols="100"><%= contents %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">처리방법</td>
	<td>
		<textarea name="solution" rows="35" cols="100"><%= solution %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">사용여부<br></td>
	<td>
        <% drawSelectBoxisusingYN "isusing",isusing,"" %>
	</td>
</tr>

<% if fidx<>"" then %>
    <tr bgcolor="#FFFFFF">
        <td align="center">최종수정<br></td>
        <td>
            <%= left(lastupdate,10) %>
            <br><%= mid(lastupdate,11,22) %>
            <% if lastadminid<>"" then %>
                <br><%= lastadminid %>
            <% end if %>
        </td>
    </tr>
<% end if %>

<tr bgcolor="#FFFFFF">
	<td align="center" colspan=2>
		<input type="button" value="저장" class="button" onclick="fnfaq_write();">
	</td>
</tr>
</table>	
</form>

<%
set ofaq = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
