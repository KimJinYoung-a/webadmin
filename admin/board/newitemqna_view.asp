<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 게시판관리>>[SITE]고객상품문의
' Hieditor : 최초 생성자 모름
'			 2017.05.19 한용민 수정(이메일발송수정. 고객이 선택한것과 상관없이 막 쏘게 되어 있었음)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/item_qnacls.asp" -->
<%
dim id,page,makerid,notupbea,mifinish,research
id= request("id")
page=request("page")
makerid=request("makerid")
notupbea=request("notupbea")
mifinish=request("mifinish")
research= request("research")
dim itemqna
set itemqna = new CItemQna
itemqna.FRectID = id
itemqna.getOneItemQna

%>
<script type='text/javascript'>
function ActReply(frm){
	var userid, username;
	userid = "<%= Replace(itemqna.FOneItem.Fuserid, Chr(34), "") %>";
	username = "<%= Replace(itemqna.FOneItem.Fusername, Chr(34), "") %>";

	if(frm.replycontents.value.length < 1) {
		alert("답변 내용을 적어주셔야 합니다.");
		frm.replycontents.focus();
		return;
	}

	if (userid.length>1){
		if (frm.replycontents.value.indexOf(userid) >= 0) {
			alert("입력불가!!\n\n고객 아이디를 답변내용에 입력할 수 없습니다.");
			return;
		}
	}
	if (username.length>1){
		if (frm.replycontents.value.indexOf(username) >= 0) {
			alert("입력불가!!\n\n고객 이름을 답변내용에 입력할 수 없습니다.");
			return;
		}
	}

	if(confirm("상품에 대해 답변 하시겠습니까?")) {
		frm.submit();
	}
}

</script>
<table width="800" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frm" method=post action="doitemqnarelpy.asp">
<input type="hidden" name="id" value="<%= itemqna.FOneItem.Fid %>">
<input type="hidden" name="imsitxt">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="notupbea" value="<%= notupbea %>">
<input type="hidden" name="mifinish" value="<%= mifinish %>">
<input type="hidden" name="research" value="<%= research %>">

<% if itemqna.FOneItem.IsReplyOk then %>
<input type="hidden" name="mode" value="reply">
<% else %>
<input type="hidden" name="mode" value="firstreply">
<% end if %>
<tr bgcolor="#FFFFFF">
	<td width=100><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FOneItem.Fitemid %>" target="_aitemlink"><img src="<%= itemqna.FOneItem.Flistimage %>" width=100 border="0"></a></td>
	<td >
	    상품코드 : <a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FOneItem.Fitemid %>" target="_aitemlink"><%= itemqna.FOneItem.FItemID %></a> <br>
		상품명 : <%= itemqna.FOneItem.FItemName %> <br>
		브랜드 : <%= itemqna.FOneItem.FMakerid %>(<%= itemqna.FOneItem.FbrandName %>) <br>
		가격 : <%= FormatNumber(itemqna.FOneItem.FSellcash,0) %> <br>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td valign=top align=center>질문자</td>
	<td>
		<%= itemqna.FOneItem.Fusername %>(<%= itemqna.FOneItem.Fuserid %>)  
		날짜 : <%= itemqna.FOneItem.Fregdate %> 
		등급 : <%= getUserLevelStrByDate(itemqna.FOneItem.fUserLevel, left(itemqna.FOneItem.Fregdate,10)) %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td valign=top align=center>질문</td>
	<td><%=chkiif(itemqna.FOneItem.FSecretYN="Y","<font color='red'>&lt;비밀글&gt;</font>","")%>
		<%= nl2br(itemqna.FOneItem.FContents) %>
	</td>
</tr>
<!--<tr bgcolor="#FFFFFF">
	<td valign=top align=center>이메일수신</td>
	<td>
		<input type="radio" name="emailok" value="Y" <% 'if itemqna.FOneItem.Femailok="Y" then response.write " checked" %>>발송 
		<input type="radio" name="emailok" value="N" <% 'if itemqna.FOneItem.Femailok="N" then response.write " checked" %>>미발송
		<br>
		<input type="text" name="usermail" value="<% '= itemqna.FOneItem.FUsermail %>" class="input_01" size="30" maxlength=80>
	</td>
</tr>-->
<tr bgcolor="#FFFFFF">
	<td valign=top align=center>답변</td>
	<td>
		  <% SelectBoxQnaPreface "02" %>&nbsp;
		  <% SelectBoxQnaCompliment "" %><br>
		  <textarea name="replycontents" cols="80" rows="15" class="input_01"><%= (itemqna.FOneItem.FReplyContents) %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align=center><input type="button" value="답변등록" onclick="ActReply(frm)">
	<a href="/admin/board/newitemqna_list.asp?page=<%= page %>&makerid=<%= makerid %>&notupbea=<%= notupbea %>&mifinish=<%=  mifinish%>&research=<%= research %>">[이전목록으로]</a>
	<a href="/admin/board/newitemqna_list.asp">[전체목록으로]</a>
	</td>
</tr>
</form>
</table>
<iframe name="PrefaceFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

 function TnChangePreface(SelectGubun){
	PrefaceFrame.location.href="/admin/board/preface_select.asp?gubun=" + SelectGubun + "&userid=<%= itemqna.FOneItem.Fuserid %>&masterid=02";
 }

 function TnChangeText(str){
var basictext;
basictext = "안녕하세요, 텐바이텐 고객센터입니다.\n"
basictext = basictext + "(내용)\n"
basictext = basictext + "만족스러운 답변이 되셨는지요?\n"
basictext = basictext + "감사합니다.\n"

	if(str == ''){
		document.frm.replycontents.value = basictext;
	}
	else{
		document.frm.replycontents.value = str;
	}
 }
<% if itemqna.FOneItem.IsReplyOk then %>
<% else %>
TnChangeText('');
<% end if %>
//-->
</script>
<iframe name="ComplimentFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

 function TnChangeCompliment(SelectGubun){
	ComplimentFrame.location.href="/admin/board/compliment_select.asp?masterid=02&gubun=" + SelectGubun;
 }

 function TnChangeText2(str){

	if(str == ''){
	}
	else{
		document.frm.replycontents.value = document.frm.imsitxt.value + "\n" + str;
	}
 }
//-->
</script>
<%
set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp" -->
