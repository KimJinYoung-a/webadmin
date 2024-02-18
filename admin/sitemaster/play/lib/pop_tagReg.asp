<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 플레이 상세페이지 태그관리
' Hieditor : 2013-09-03 이종화 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<%
Dim idx , subidx , playcate , oPlayTag , i
	idx  = requestCheckVar(request("idx"),10)
	subidx  = requestCheckVar(request("subidx"),10)
	playcate = requestCheckVar(request("playcate"),10)

IF idx = "" THEN
	Response.Write "<script>alert('잘못된 경로입니다.\nNo. 번호가 있어야 합니다.');</script>"
	dbget.close()
	Response.End
END IF
IF IsNumeric(idx) = False THEN
	Response.Write "<script>alert('잘못된 경로입니다.\nNo. 번호가 있어야 합니다.');</script>"
	dbget.close()
	Response.End
END If

set oPlayTag = new CPlayContents
	oPlayTag.FRectIdx = idx
	oPlayTag.FRectsubIdx = subidx
	oPlayTag.FRectPlaycate = playcate
	oPlayTag.GetRowTagContent()

%>
<script src="/js/jquery-1.7.1.min.js" type="text/javascript"></script>
<script type="text/javascript">
	$(document).ready(function(){
		// 옵션추가 버튼 클릭시
		$("#addItemBtn").click(function(){
			// item 의 최대번호 구하기
			var lastItemNo = $("#imgIn tr:last").attr("class").replace("item", "");

			var newitem = $("#imgIn tr:eq(1)").clone();
			newitem.removeClass();
			newitem.find("td:eq(0)").attr("rowspan", "1");
			newitem.find("#tagname").attr("value", "");
			newitem.find("#tagurl").attr("value", "");
			newitem.find("#tagurl2").attr("value", "");
			newitem.find("#tagurl3").attr("value", "");
			newitem.find("#tagurl4").attr("value", "");
			newitem.addClass("item"+(parseInt(lastItemNo)+1));

			$("#imgIn").append(newitem);
		});
	});

	function chgextxt(v) {
	var urllink = document.getElementById("extxt");
		switch(v) {
			case "1":
				urllink.value='검색어자동입력(미구현)';
				break;
			case "2":
				urllink.value='55073 <--이벤트번호 입력';
				break;
			case "3":
				urllink.value='392832 <--상품코드 입력';
				break;
			case "4":
				urllink.value='102102104 <--카테고리번호 입력';
				break;
			case "5":
				urllink.value='ithinkso <--브랜드아이디 입력';
				break;
		}
	}
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 태그관리<br/>※태그 미입력시 자동 삭제 됩니다. (URL 입력 유무 상관 없음)※<br/>※URL 미입력시 검색페이지로 이동합니다.※<br/>※현재 보이는 순서대로 페이지에 뿌려집니다.※<br/>※<span style="color:blue;font-weight:800;">하단 예제) 참조 문의사항은 시스템팀으로 연락바랍니다.</span>※
<br/>※<span style="color:red;font-weight:300;">PC-URL : 기존 URL 입력 예)/event/eventmain.asp?eventid=이벤트코드</span>※
<br/>※<span style="color:red;font-weight:300;">MO-URL : 모바일 URL입력 예)/category/category_itemprd.asp?itemid=상품코드</span>※
<br/>※<span style="color:red;font-weight:300;">APP-URL : SELECT선택후 해당 코드 입력 (이벤트코드,상품코드,브랜드명,카테고리번호 중 1개)</span>※
</div>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="red">
<tr>
	<td colspan="3">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50">예제)</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50"><input type="text" value="가방" size="15" readonly/></td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="250"><input type="text" value="/event/eventmain.asp?eventid=이벤트코드" size="35" readonly/></td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="250"><input type="text" value="/category/category_itemprd.asp?itemid=상품코드" size="35" readonly/></td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="300">
				<select onchange="chgextxt(this.value);">
					<option value="">=선택=</option>
					<option value="1">상품상세</option>
					<option value="2">이벤트</option>
					<option value="3">브랜드</option>
					<option value="4">카테고리</option>
				</select>
				<input type="text" id="extxt" value="&lt;-- 선택후 해당 번호 입력" size="30" readonly/>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<form name="frmtag" method="post" action="/admin/sitemaster/play/lib/tagProc.asp" >
<input type="hidden" name="mode" value="tag"/>
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="subidx" value="<%=subidx%>"/>
<input type="hidden" name="playcate" value="<%=playcate%>"/>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a">
<tr>
	<td colspan="3">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="a" id="imgIn">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50">&nbsp;</td>
			<td bgcolor="<%= adminColor("tabletop") %>">태그입력</td>
			<td bgcolor="<%= adminColor("tabletop") %>">PC-URL입력</td>
			<td bgcolor="<%= adminColor("tabletop") %>">MO-URL입력</td>
			<td bgcolor="<%= adminColor("tabletop") %>">APP-URL입력</td>
		</tr>
		<% If oPlayTag.FTotalCount > 0  Then %>
		<% for i=0 to oPlayTag.FTotalCount - 1 %>
		<tr class="item<%=i+1%>">
			<td bgcolor="<%= adminColor("tabletop") %>" width="50">태그등록</td>
			<td bgcolor="#FFFFFF" width="50"><input type="text" name="tagname" value="<%=oPlayTag.FItemList(i).Ftagname%>" size="15" id="tagname" /></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl" value="<%=oPlayTag.FItemList(i).Ftagurl%>" size="35" id="tagurl"/></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl2" value="<%=oPlayTag.FItemList(i).Ftagurl2%>" size="35" id="tagurl2"/></td>
			<td bgcolor="#FFFFFF" width="300">
				<select name="tagurl3" id="tagurl3">
					<option value="" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="","selected","")%>>==선택==</option>
					<option value="1" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="1","selected","")%>>상품상세</option>
					<option value="2" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="2","selected","")%>>이벤트</option>
					<option value="3" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="3","selected","")%>>브랜드</option>
					<option value="4" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="4","selected","")%>>카테고리</option>
				</select>
				<input type="text" name="tagurl4" value="<%=oPlayTag.FItemList(i).Ftagurl4%>" size="30" id="tagurl4"/>
			</td>
		</tr>
		<% next%>
		<% Else %>
		<tr class="item1">
			<td bgcolor="<%= adminColor("tabletop") %>">태그등록</td>
			<td bgcolor="#FFFFFF" width="50"><input type="text" name="tagname" value="" size="15" id="tagname" /></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl" value="" size="35" id="tagurl"/></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl2" value="" size="35" id="tagurl2"/></td>
			<td bgcolor="#FFFFFF" width="300">
				<select name="tagurl3" id="tagurl3">
					<option value="">=선택=</option>
					<option value="1">상품상세</option>
					<option value="2">이벤트</option>
					<option value="3">브랜드</option>
					<option value="4">카테고리</option>
				</select>
				<input type="text" name="tagurl4" value="" size="30" id="tagurl4"/>
			</td>
		</tr>
		<% End If %>
		</table>
	</td>
</tr>
<tr>
	<td align="left" colspan="1">
		<INPUT TYPE="button" id="addItemBtn" value="태그 추가"/>
	</td>
	<td align="right">
		<input type="image" src="/images/icon_confirm.gif"/>
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</table>
</form>
<%
	set oPlayTag = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->