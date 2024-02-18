<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/openGift.asp
' Description :  전체증정이벤트 관리 369등.
' History : 2010.04 서동석 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/openGiftCls.asp"-->

<%


dim evtCode, frontOpen
dim i, page
page = requestCheckVar(request("page"),10)
evtCode = requestCheckVar(request("evtCode"),10)
frontOpen = requestCheckVar(request("frontOpen"),10)

if page="" then page=1

Dim oOpenGift
set oOpenGift=new CopenGift
oOpenGift.FCurrPage = page
oOpenGift.FPageSize = 30
oOpenGift.getOpenGiftList

Dim urlPara : urlPara = Server.UrlEnCode("&evtCode="&evtCode&"&page="&page&"&frontOpen="&frontOpen)
%>

<script language="javascript">
function jsLastEvent(){
	  var winLast,eKind;
	  eKind = 1;
	  var pTarget = '<%= Server.URLEncode("openGift_Reg.asp?menupos=1184") %>';
	  winLast = window.open('pop_event_lastlist.asp?menupos=<%=menupos%>&eventkind='+eKind+'&pTarget='+pTarget,'pLast','width=550,height=600, scrollbars=yes')
	  winLast.focus();
	}
	
function changeGiftScope(eCode,ogiftType){
    var frm = document.frmSm;
    var confirmStr ='사은품 증정 대상을 전체증정으로 변경하시겠습니까?';
    if (ogiftType==9){
        confirmStr ='사은품 증정 대상을 다이어리증정타입으로 변경하시겠습니까?';
    }
    
    if (confirm(confirmStr)){
        frm.eCode.value=eCode;
        frm.opengiftType.value=ogiftType;
        frm.submit();
    }
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmEvt" method="get"  >
	<input type="hidden" name="menupos" value="<%=menupos%>">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">  
			이벤트 코드 : <input type="text" name="evtCode" value="<%=evtCode%>" maxlength="10" size="9">
			
			프런트오픈 : 
			<select name="frontOpen">
			<option value="">전체
			<option value="Y" <%= CHKIIF(frontOpen="Y","selected","") %> >오픈
			</select>
		</td>
        </td>
    		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:frmEvt.submit();">
		</td>
	</tr>	
	</form>
</table>
<!-- 표 상단바 끝-->
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
    <tr height="40" valign="bottom">       
        <td align="left">
        	<input type="button" value="이벤트 지정" onclick="jsLastEvent();" class="button">
	    </td>
	    <td align="right">
	       <!-- input type="button" value="스케쥴" onclick="jsSchedule();"  class="button" -->
	       <!-- <% if C_ADMIN_AUTH then %><input type="button" value="코드관리" onclick="jsCodeManage();"  class="button"><%END IF%> -->
        </td>        
	</tr>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="12">검색결과 : <b><%=oOpenGift.FTotalCount%></b>&nbsp;&nbsp;페이지 : <b><%=page%> / <%=oOpenGift.FtotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td nowrap>이벤트코드</td>
    	<td nowrap>구분</td>
    	<td nowrap>범위</td>
    	<td nowrap>이벤트명</td>
    	<td nowrap>이벤트기간</td>
    	<td nowrap>이벤트상태</td>
    	<td nowrap>프런트오픈</td>
    	<td nowrap>사은품</td>
    	<td nowrap>등록자</td>
    	<td nowrap>비고</td>
    </tr>
    <% for i= 0 to oOpenGift.FREsultCount -1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= oOpenGift.FItemList(i).FEvent_Code %></td>
        <td><%= oOpenGift.FItemList(i).getOpengiftTypeName %></td>
        <td><%= oOpenGift.FItemList(i).getOpengiftScopeName %></td>
        <td align="left"><a href="openGift_Reg.asp?eC=<%= oOpenGift.FItemList(i).FEvent_Code %>&menupos=<%=menupos%>"><%= oOpenGift.FItemList(i).FEvent_Name %></a></td>
        <td><%= oOpenGift.FItemList(i).Fevt_startdate%>~<%= oOpenGift.FItemList(i).Fevt_enddate %></td>
        <td><%= oOpenGift.FItemList(i).getEventStateName %></td>
        <td><%= oOpenGift.FItemList(i).FfrontOpen %></td>
        <td><a href="/admin/shopmaster/gift/giftlist.asp?eC=<%= oOpenGift.FItemList(i).FEvent_Code %>&menupos=1045&fcSc=1"><%= oOpenGift.FItemList(i).FGiftCNT %></a>
        <% if oOpenGift.FItemList(i).FGiftCNT<>oOpenGift.FItemList(i).FALLGiftCNT then %>
        <strong><font color="red">(<%= oOpenGift.FItemList(i).FALLGiftCNT %>)</font></strong>
        <% end if %>
        </td>
        <td><%= oOpenGift.FItemList(i).Freguser %></td>
        <td>
        <% if oOpenGift.FItemList(i).FGiftCNT<>oOpenGift.FItemList(i).FALLGiftCNT then %>
        <img src="/images/icon_arrow_link.gif" onClick="changeGiftScope('<%= oOpenGift.FItemList(i).FEvent_Code %>,'<%=oOpenGift.FItemList(i).FopengiftType%>'')" style="cursor:pointer">
        <% end if %>
        </td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
		<td colspan="12" height="30" align="center">
		<% if oOpenGift.HasPreScroll then %>
			<a href="?page=<%= oOpenGift.StarScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oOpenGift.StarScrollPage to oOpenGift.FScrollCount + oOpenGift.StarScrollPage - 1 %>
			<% if i>oOpenGift.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oOpenGift.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>

<form name="frmSm" method="post" action="openGift_Process.asp" >
<input type="hidden" name="imod" value="chgScope"> 
<input type="hidden" name="eCode" value=""> 
<input type="hidden" name="opengiftType" value=""> 
<input type="hidden" name="opengiftScope" value=""> 
</form>
<%
set oOpenGift=Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->