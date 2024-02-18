<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs센터 간단 메모 입력
' History : 2012.05.21 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->

<%
dim ocsmemo , id , qadiv ,orderserial,  isEditMode ,userid
	id              = RequestCheckVar(request("id"),9)
	orderserial = requestCheckVar(request("orderserial"),11)
	userid = requestCheckVar(request("userid"),32)
	
'/cs메모
set ocsmemo = New CCSMemo

if (id <> "") then
	ocsmemo.FRectId = id
	ocsmemo.FRectUserID = userid
	ocsmemo.FRectOrderserial = orderserial
	ocsmemo.GetCSMemoDetail
	
	userid = ocsmemo.FOneItem.FUserID
	orderserial = ocsmemo.FOneItem.Forderserial
	qadiv = ocsmemo.FOneItem.Fqadiv
	
	isEditMode = true
else
	ocsmemo.GetCSMemoBlankDetail
	''mayBe Inbound
	ocsmemo.FOneItem.FmmGubun = "1"
	isEditMode = false
end if

if qadiv = "" then qadiv = "20"
%>

<script type='text/javascript'>

var NowDoing = false;
<% if (orderserial<>"") or (userid<>"") then %>
    NowDoing = true;
<% end if %>

function checkDoing(){
    if (!NowDoing){
        NowDoing=true;
        setDoingState();
    }
}

//cs메모 등록
function SubmitSave()
{
    if ((document.frm.orderserial.value.length<1)&&(document.frm.userid.value.length<1)) {
	    alert("주문번호, 아이디 중 하나는 입력 되어야 합니다.");
		return;
	}
	
	if (document.frm.contents_jupsu.value == "") {
		alert("메모내용을 입력하세요.");
		document.frm.contents_jupsu.focus();
		return;
	}
	
	if (document.frm.qadiv.value.length<1){
	    alert("문의 유형을 선택 하세요.");
		document.frm.qadiv.focus();
		return;
	}
	
	if(document.frm.id.value == "") {
    	document.frm.mode.value = "write";
    	document.frm.submit();
	}else{
    	document.frm.mode.value = "modify";
    	document.frm.submit();
	}
}

//cs메모 완료처리
function SubmitFinish(){
	if (document.frm.contents_jupsu.value == "") {
			alert("메모내용을 입력하세요.");
			return;
	}		
	
    if (confirm("완료처리하겠습니까?") == true) {
            document.frm.mode.value = "finish";
            document.frm.submit();
    }
}

//cs메모 삭제
function SubmitDelete()
{
    if (confirm("삭제하겠습니까?") == true) {
            document.frm.mode.value = "delete";
            document.frm.submit();
    }
}


</script>

<!-- CS메모-CALL 시작-->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
		        <td>
		        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS메모</b>
		        </td>
		        <td align="right">
		            
		            <input type="button" class="button" value="<%= chkIIF(isEditMode,"수정","저장") %>" onclick="javascript:SubmitSave();">
			       	<input type="button" class="button" value="완료" <%= chkIIF((Not isEditMode) or (ocsmemo.FOneItem.Fdivcd<>"2"),"disabled","") %> onclick="javascript:SubmitFinish();">
			        <input type="button" class="button" value="삭제" <%= chkIIF(isEditMode,"","disabled") %> onclick="javascript:SubmitDelete();">
			        <!-- 닫기 버튼 필요없슴
			        <input type="button" class="button" value="닫기" onclick="javascript:window.close();">
			         -->
			    </td>
			</tr>	    
		</table>
	</td>
</tr>
<form name="frm" onsubmit="return false;" method="post" action="/common/cscenter/popCallRing_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="id" value="<%= ocsmemo.FOneItem.Fid %>">
<tr>
	<td width="50" bgcolor="<%= adminColor("tabletop") %>">구분</td>
	<td bgcolor="#FFFFFF">
        <select name="mmGubun" onChange="setGubunState(this);" class="select">
            <option value="0" <% if ocsmemo.FOneItem.FmmGubun = "0" then %>selected<% end if %>>일반메모</option>
        </select>
    </td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
	<td bgcolor="#FFFFFF">
	    <table width="370" cellpadding="0" cellspacing="0" border="0" >
    	<tr>
    	    <input type="text" name="orderserial" class="text" value="<%= orderserial %>" size="20" onKeyDown="checkDoing();" >
    	</tr>
    	</table>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">고객ID</td>
	<td bgcolor="#FFFFFF">
	    <table width="370" cellpadding="0" cellspacing="0" border="0" >
    	<tr>
    	    <input type="text" name="userid" class="text" value="<%= userid %>" size="20" onKeyDown="checkDoing();">
    	</tr>
    	</table>
	</td>
</tr>
<% if id = "" then %>
<% else %>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">접수일</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.fregdate %>" size="26" readonly>&nbsp;
    		등록자ID : <%= ocsmemo.FOneItem.Fwriteuser %>
    	</td>
    </tr>
<% end if %>
<% if ucase(ocsmemo.FOneItem.Ffinishyn) <> "Y" then %>
<% else %>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">완료일</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.Ffinishdate %>" size="26" readonly>&nbsp;
    		처리자ID : <%= ocsmemo.FOneItem.Ffinishuser %>
    	</td>
    </tr>
<% end if %>	 
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">문의유형</td>
	<td bgcolor="#FFFFFF">
	    
	    <% if ocsmemo.FOneItem.Fdivcd="2" then %>
	    <input type=hidden name="divcd" value="2">
	    <input type="checkbox" name="dummi" checked disabled >처리요청
	    <% else %>
	    <input type="checkbox" name="divcd" value="2" >처리요청
	    <% end if %>
	    
        <!-- 유형 : -->
        &nbsp;&nbsp;
        <% Call DrawMemoDivCombo("qadiv",qadiv) %>
		
    </td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">메모<br>내용</td>
	<td bgcolor="#FFFFFF"><textarea name="contents_jupsu" class="textarea" cols="60" rows="10" onKeyDown="checkDoing();"><%= replace(db2html(ocsmemo.FOneItem.Fcontents_jupsu),"<br>",vbCrlf) %></textarea></td>
</tr>
</table>
<!-- CS메모-CALL 끝-->

<%
	set ocsmemo = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->