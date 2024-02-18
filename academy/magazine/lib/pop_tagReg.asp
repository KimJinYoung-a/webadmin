<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 매거진 상세페이지 태그관리
' Hieditor : 2016-03-04 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/academy/magazineCls.asp" -->
<%
Dim idx, oMagaZineTag , i
	idx  = requestCheckVar(request("idx"),10)

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

set oMagaZineTag = new CMagazineContents
	oMagaZineTag.FRectIdx = idx
	oMagaZineTag.GetRowTagContent()

%>
<script src="/js/jquery-1.7.1.min.js" type="text/javascript"></script>
<script type="text/javascript">

    $(document).ready(function(){

        // 옵션추가 버튼 클릭시
        $("#addItemBtn").click(function(){
            // item 의 최대번호 구하기
            var lastItemNo = $("#tagadd tr:last").attr("class").replace("item", "");

            var newitem = $("#tagadd tr:eq(2)").clone();
            newitem.removeClass();
            newitem.find("td:eq(0)").attr("rowspan", "1");
            newitem.find("#tagname").attr("value", "");
            newitem.addClass("item"+(parseInt(lastItemNo)+1));

            $("#tagadd").append(newitem);

        });

    });
</script>
<div style="padding: 0 5 5 5">
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 태그관리<br/>※태그 미입력시 자동 삭제 됩니다.※<br/>※현재 보이는 순서대로 페이지에 뿌려집니다.※
</div>

<form name="frmtag" method="post" action="/academy/magazine/lib/tagProc.asp" >
<input type="hidden" name="mode" value="tag"/>
<input type="hidden" name="idx" value="<%=idx%>"/>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a">
<tr>
	<td colspan="2">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="a" id="tagadd">

		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50">&nbsp;</td>
			<td bgcolor="<%= adminColor("tabletop") %>">태그입력</td>
		</tr>
		<tr class="hidden">
			<td bgcolor="<%= adminColor("tabletop") %>"></td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50"></td>
		</tr>
		<% If oMagaZineTag.FTotalCount > 0  Then %>
			<% for i=0 to oMagaZineTag.FTotalCount - 1 %>
				<tr class="item<%= i+1 %>">
					<td bgcolor="<%= adminColor("tabletop") %>" width="50">태그등록</td>
					<td bgcolor="#FFFFFF" width="50"><input type="text" name="tagname" value="<%= oMagaZineTag.FItemList(i).Fsearchkw %>" size="15" id="tagname" /></td>
				</tr>
			<% next %>
		<% Else %>
			<tr class="item1">
				<td bgcolor="<%= adminColor("tabletop") %>">태그등록</td>
				<td bgcolor="#FFFFFF" width="50"><input type="text" name="tagname" value="" size="15" id="tagname" /></td>
			</tr>
		<% End If %>
		<tr class="hidden2">
			<td bgcolor="<%= adminColor("tabletop") %>"></td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50"></td>
		</tr>


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
<% set oMagaZineTag = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->