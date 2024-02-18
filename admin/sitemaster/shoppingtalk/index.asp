<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : GIFT TALK 관리
' Hieditor : 강준구 생성
'			 2022.07.08 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/sitemaster/shoppingtalk/classes/shoppingtalkCls.asp" -->

<%
	Dim vCurrPage, i, j, vTalkIdx, vUserID, vItemID, vTheme, vKeyword, vUseYN, vItem, vItemTmp
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	vCurrPage = requestCheckVar(getNumeric(request("cpg")),10)
	If vCurrPage = "" Then vCurrPage = 1
	iPerCnt 		= 10
	
	If isNumeric(vCurrPage) = False Then
		Response.Write "<script>alert('잘못된 경로입니다.');location.href='/';</script>"
		dbget.close()
		Response.End
	End If
	Dim vKey1, vKey2
	vKey1 = requestCheckVar(request("key1"),3)
	vKey2 = requestCheckVar(request("key2"),3)
	vKeyword = vKey2
	IF vKeyword = "" Then
		vKeyword = vKey1
	End IF
	
	Dim cTalk
	SET cTalk = New CShoppingTalk
	cTalk.FPageSize = 10
	cTalk.FCurrpage = vCurrPage
	cTalk.FRectTalkIdx = vTalkIdx
	cTalk.FRectUserId = vUserID
	cTalk.FRectItemId = vItemID
	cTalk.FRectTheme = vTheme
	cTalk.FRectKeyword = vKeyword
	'cTalk.FRectUseYN = "y"
	cTalk.fnShoppingTalkList
%>

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

function code_manage()
{
	window.open('PopManageCode.asp','coopcode','width=600,height=768');
}

function jsGoPage(iP){
	document.frmpage.cpg.value = iP;
	document.frmpage.submit();
}

function talkuseyn(tidx, v)
{
	var m = "";
	var a = "";
	if(v == "n"){
		m = "삭제";
		a = "y";
	}else{
		m = "OPEN";
		a = "n";
	}
	if(confirm("선택하신 쇼핑톡을 "+m+"하시겠습니까?") == true) {
		$('input[name="action"]').val("update");
		$('input[name="talkidx"]').val(tidx);
		$('input[name="useyn"]').val(v);
		frm1.submit();
     }else{
		$('input[name="action"]').val("");
		$('input[name="talkidx"]').val("");
		$('input[name="useyn"]').val("");
		$("select[name=useyn"+tidx+"]").val(a);
		return;
     }
}

function talk_commentlist(tidx)
{
	window.open('commentlist.asp?talkidx='+tidx+'','commentlist','width=500,height=570');
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		<table width="100%" class="a">
		<tr>
			<td>검색결과 : <b><%= cTalk.FTotalCount %></b></td>
			<td align="right"><input type="button" onClick="code_manage()" value="상황,대상코드관리" class="button"></td>
		</tr>
		</table>
	</td>
</tr>
<% If (cTalk.FResultCount < 1) Then %>
<%
	Else
		For i = 0 To cTalk.FResultCount-1
%>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30" style="cursor:pointer">
			<td><%=cTalk.FItemList(i).FTalkIdx%></td>
			<td><%=CHKIIF(cTalk.FItemList(i).FTheme="2","비교","추천")%></td>
			<td><%=cTalk.FItemList(i).FUserID%></td>
			<td><%=cTalk.FItemList(i).FRegdate%></td>
			<td>
				<select name="useyn<%=cTalk.FItemList(i).FTalkIdx%>" class="select" onChange="talkuseyn('<%=cTalk.FItemList(i).FTalkIdx%>',this.value);">
				<option value="y" <%=CHKIIF(cTalk.FItemList(i).FUseYN="y","selected","")%>>OPEN</option>
				<option value="n" <%=CHKIIF(cTalk.FItemList(i).FUseYN="n","selected","")%>>삭제처리</option>
				</select>
			</td>
			<td><%=cTalk.FItemList(i).FCommCnt%>
				<% If cTalk.FItemList(i).FCommCnt > 0 Then %>
					&nbsp;[<a href="javascript:talk_commentlist('<%=cTalk.FItemList(i).FTalkIdx%>');">댓글리스트</a>]
				<% End If %>
			</td>
		</tr>
		<tr style="cursor:pointer">
			<td colspan="7" bgcolor="#FFFFFF">
			<%
			'### 0:good, 1:bad, 2:itemid, 3:itemname, 4:makerid, 5:brandname, 6:listimage, 7:icon1image, 8:icon2image, 9:basicimage, 10:idx
			vItem = cTalk.FItemList(i).FItem
			vItem = Right(vItem,Len(vItem)-5)
			If cTalk.FItemList(i).FTheme = "2" Then
			%>
				<table class="a" border="0">
				<tr>
					<%
					For j = LBound(Split(vItem,",item,")) To UBound(Split(vItem,",item,"))
						vItemTmp = Split(vItem,",item,")(j)
					%>
						<td <%=CHKIIF(j=0,""," style='padding-left:50px;'")%>><%=CHKIIF(j=0,"A","B")%>선택 : <%=db2html(Split(vItemTmp,"|blank|")(0))%><br>
							<a href="http://m.10x10.co.kr/gift/gifttalk/talk_view.asp?talkidx=<%=cTalk.FItemList(i).FTalkIdx%>" target="_blank"><img src="http://webimage.10x10.co.kr/image/icon1/<%=GetImageSubFolderByItemid(Split(vItemTmp,"|blank|")(2)) & "/" & Split(vItemTmp,"|blank|")(7)%>" style="width:130px;" border="0" /></a>
						</td>
					<%
					Next
					%>
					<td style="padding-left:30px;" valign="top"><%=cTalk.FItemList(i).FContents%></td>
				</tr>
				</table>
			<%
			Else
			%>
				<table class="a" border="0">
				<tr>
					<td>찬성 : <%=Split(vItem,"|blank|")(0)%></td>
					<td align="right">반대 : <%=Split(vItem,"|blank|")(1)%></td>
					<td rowspan="1" style="padding-left:30px;"><%=cTalk.FItemList(i).FContents%></td>
				</tr>
				<tr>
					<td align="center" colspan="2"><a href="http://m.10x10.co.kr/gift/gifttalk/talk_view.asp?talkidx=<%=cTalk.FItemList(i).FTalkIdx%>" target="_blank"><img src="http://webimage.10x10.co.kr/image/basic/<%=GetImageSubFolderByItemid(Split(vItem,"|blank|")(2)) & "/" & Split(vItem,"|blank|")(9)%>" style="width:150px;" border="0" /></a></td>
				</tr>
				</table>
			<% End If %>
			</td>
		</tr>
		<% Next %>
	<%
	iTotalPage 	=  int((cTalk.FTotalCount-1)/10) +1
	iStartPage = (Int((vCurrPage-1)/iPerCnt)*iPerCnt) + 1
	
	If (vCurrPage mod iPerCnt) = 0 Then
		iEndPage = vCurrPage
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
	%>
	
	<form name="frmpage" method="get" action="<%=CurrURL()%>" style="margin:0px;">
	<input type="hidden" name="cpg" value="">
	<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(vCurrPage) then
			%>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
			<%		else %>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
			<%
					end if
				next
			%>
	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
		</td>
	</tr>
    </form>
<% End If %>
</table>
<form name="frm1" action="proc.asp" target="prociframe" style="margin:0px;">
<input type="hidden" name="gubun" value="talk">
<input type="hidden" name="action" value="">
<input type="hidden" name="talkidx" value="">
<input type="hidden" name="useyn" value="">
</form>
<iframe name="prociframe" id="prociframe" src="" width="0" height="0"></iframe>
<% SET cTalk = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->