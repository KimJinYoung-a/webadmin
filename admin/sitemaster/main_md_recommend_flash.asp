<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : PC메인관리 MD픽
' History : 서동석 생성
'			2022.07.01 한용민 수정(isms취약점조치)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/pcmain/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_event_rotationcls.asp"-->
<%

dim i
dim page, malltype
dim isusing, research
dim itemid, sdate, edate, realdate, lowestPrice
dim realdatereset, getdate

page = request("page")
isusing = request("isusing")
research = request("research")
itemid = request("itemid")
sdate = request("iSD")
edate = request("iED")
realdate = request("realdate")
realdatereset = request("realdatereset")
getdate = request("getdate")

lowestPrice = request("lowestPrice")
if (page = "") then
        page = "1"
end if

if research = "" and isusing="" then isusing="Y"

if realdatereset = "1" then 
	realdate = ""
else
	if realdate = "" then realdate = date()
end if 
if getdate="" then getdate=realdate
'==============================================================================
dim mdchoicerotate
set mdchoicerotate = new CMainMdChoiceRotate

mdchoicerotate.FCurrPage = CInt(page)
mdchoicerotate.FPageSize = 100
mdchoicerotate.FRectIsUsing = isusing
mdchoicerotate.FRectItemID = itemid
mdchoicerotate.FRectSDate = sdate
mdchoicerotate.FRectEDate = edate
mdchoicerotate.FRectrealdate = realdate
mdchoicerotate.FRectIsLowestPrice = lowestPrice
mdchoicerotate.list

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<script type="text/javascript">
function NextPage(page){
	document.frm.page.value=page;
	document.frm.submit();
}

function frmChange()
{
	var vfm = document.vfrm;
	if(confirm("<%=realdate%>일의 전시순서가 현재 목록에 보이는 상태 그대로 모두 적용됩니다.\n전체 적용 하시겠습니까?"))
	{
		vfm.action="doMainMdChoiceChange.asp";
		vfm.submit()
	}
	else
		return;
}

var chkUsing="<%=isusing%>";
function usingAllChange()
{
	if(chkUsing=="Y") { chkUsing = "N"; }
	else { chkUsing = "Y"; }

	for (var i=0;i<document.vfrm.isusing.length;i++){
		document.vfrm.isusing[i].value=chkUsing;
	}
}

function writeItem(idx) {
	if(idx==0) {
		var mode = "write";
	} else {
		var mode = "modify";
	}
	var mdcWrPop = window.open("main_md_recommend_flash_write.asp?mode="+mode+"&idx="+idx,"popMdcWr","width=1200,height=700,scrollbars=yes");
	mdcWrPop.focus();
}

function editItemImage(itemid) {
	var param = "itemid=" + itemid;

	//if(makerid =="ithinkso"){
		//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=1200,height=700,scrollbars=yes,resizable=yes');
	//}else{
		popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=1000,height=900,scrollbars=yes,resizable=yes');
	//}
	popwin.focus();
}

function popupMainPreview(realdate){			
	var popwin; 		
	popwin = window.open("/admin/sitemaster/main_preview.asp?realdate="+realdate, "popup_main_preview", "width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function popRegArrayItem() {
	var popwin;
    var popwin = window.open('main_md_recommend_flash_writes.asp?realdate=<%=realdate%>','popRegArray','width=1200,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popCurrentStock(itemid) {
	var popwin;
    var popwin = window.open('/admin/stock/itemcurrentstock.asp?menupos=<%= request("menupos") %>&itemid='+itemid,'popRegArray','width=1200,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function checkdate() {
	var frm = document.frm;
		frm.realdate.value = '';
}

function delItem(idx) {
	var vFrm = document.delfrm;
	if(confirm("해당 상품이 삭제 됩니다.\n적용 하시겠습니까?"))
	{
		vFrm.idx.value = idx;
		vFrm.realdatereset.value = (document.frm.realdatereset.checked) ? 1 : 0;
		vFrm.action="doMainMdChoiceChange.asp";
		vFrm.submit();
	} else {
		return;
	}
}

$(function() {
	<% if realdate <> "" then %>
	$("#subList").sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="45" colspan="13" style="border:1px solid #F9BD01;">&nbsp;</td>');
			$(".etcInfo").hide();
		},
		stop: function(){
			var i=0;
			$(this).find("input[name^='disporder']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='disporder']").each(function(){
				$(this).val(i);
				i++;
			});
			$(".etcInfo").show();
		}
	});
	<% end if %>
})

function fnMobileMDPickCopy(){
	if(confirm("모바일 MD Pick의 내용을 가져와 PC메인에 적용합니다.\n가져오시겠습니까?\n\n※ (주의) PC메인에 바로 반영됩니다.")) {
		if($("#getdate").val()!=""){
			$.ajax({
				type: "POST",
				url: "ajaxMDRecommendCopy.asp",
				data: "getdate="+$("#getdate").val(),
				cache: false,
				success: function(message) {
					if(message=="OK") {
						alert("복사 완료");
						window.location.reload();
					} else {
						alert("복사에 실패했습니다.");
					}
				},
				error: function(err) {
					alert(err.responseText);
				}
			});
		}
		else{
			alert("카피일을 선택해주세요.");
		}
	}
}
</script>

<form name="refreshFrm" method="post" style="margin:0;">
</form>

<form name="delfrm" method="post" action="">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<input type="hidden" name="realdate" value="<%=realdate%>">
	<input type="hidden" name="mode" value="del" />
	<input type="hidden" name="realdatereset" value="" />
	<input type="hidden" name="idx" />
</form>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상품코드 :
		<input type="text" name="itemid" value="<%= itemid %>" size=9 maxlength=9 class="text"> &nbsp;/
		등록일 : 
		<input id="iSD" name="iSD" value="<%=sdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="iED" name="iED" value="<%=edate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		 &nbsp; &nbsp; &nbsp;최저가 여부 : 
		<select name="lowestPrice" class="select">
		<option value="" >전체
		<option value="Y" <%=chkIIF(lowestPrice="Y","selected","")%> >사용
		<option value="N" <%=chkIIF(lowestPrice="N","selected","")%> >사용안함
		</select> &nbsp;
		<div style="float:right;vertical-align:middle;padding-right:20px;">
			지정일자 : 
			<input id="realdate" name="realdate" value="<%=realdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="realdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> &nbsp;
			[ 전체검색 : <input type="checkbox" name="realdatereset" value="1" <%=chkiif(realdatereset = "1" ,"checked","")%> onclick="checkdate()" style="vertical-align:middle"/> ]
		</div>
		<br>
	</td>
	<script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "iSD", trigger    : "iSD_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_End = new Calendar({
			inputField : "iED", trigger    : "iED_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_End = new Calendar({
			inputField : "realdate", trigger    : "realdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
				$("input:checkbox[name='realdatereset']").prop("checked", false);
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
	</script>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0;">
	<tr>
		<td align="right">
			<input id="getdate" name="getdate" value="<%=getdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="getdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_End = new Calendar({
					inputField : "getdate", trigger    : "getdate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			<input type="button" class="button" value="모바일MDPICK가져오기" onClick="fnMobileMDPickCopy();">
			<% if realdate <> "" then %>
			<input type="button" class="button" value="전시순서변경" onClick="frmChange()">
			&nbsp;
			<% end if %>
			<input type="button" class="button" value="복수상품등록" onClick="javascript:popRegArrayItem();">
			&nbsp;
			<input type="button" class="button" value="신규등록" onClick="javascript:writeItem(0);">
			&nbsp;
			<input type="button" class="button" value="미리보기" onClick="popupMainPreview('<%=realdate%>')">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<form name="vfrm" method="POST" action="">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="sUsing" value="<%= isusing %>">
<input type="hidden" name="realdate" value="<%=realdate%>">
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		검색결과 : <b><%=mdchoicerotate.Ftotalcount%></b>
		&nbsp;
		페이지 : <b><%=page%> / <%=mdchoicerotate.Ftotalpage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center">상품코드</td>
	<td align="center" class="etcInfo">이미지</td>
	<td align="center" class="etcInfo">상품구분</td>
	<td align="center">상품정보</td>
	<td align="center">전시순서</td>
    <td align="center" class="etcInfo">시작일</td>
    <td align="center" class="etcInfo">종료일</td>
	<td align="center" class="etcInfo">등록일</td>
	<td align="center" class="etcInfo">최저가 여부</td>	
	<td align="center" class="etcInfo">판매여부</td>
	<td align="center" class="etcInfo">등록자</td>
	<td align="center" class="etcInfo">최종작업자</td>
	<td align="center" class="etcInfo">상품 삭제</td>
</tr>
<tbody id="subList">
<% for i=0 to mdchoicerotate.FResultcount -1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<input type="hidden" name="idx" value="<%= mdchoicerotate.FItemList(i).Fidx %>">
			<a href="javascript:writeItem(<%= mdchoicerotate.FItemList(i).Fidx %>);"><%= mdchoicerotate.FItemList(i).Flinkitemid %></a>
		</td>
		<td align="center" class="etcInfo">
			<% if mdchoicerotate.FItemList(i).FTentenImg <> "" then %>
				<img src="<%= mdchoicerotate.FItemList(i).FTentenImg %>" border=0 width="56">
			<% else %>				
				<img src="<%= mdchoicerotate.FItemList(i).Fphotoimg %>" border=0 width="56">					
			<% end if %>
			<br/><button type="button" onClick="editItemImage('<%= mdchoicerotate.FItemList(i).Flinkitemid %>')">수정</button>
		</td>
		<td align="center" class="etcInfo">
			<% if mdchoicerotate.FItemList(i).FItemDiv = "21" then %>
				<span style="color:blue">딜 상품</span>			
			<% else %>
				<span style="color:red">일반 상품</span>							
			<% end if %>			
		</td>				
		<td align="left">
			
			<a href="javascript:writeItem(<%= mdchoicerotate.FItemList(i).Fidx %>);">
				<%=chkIIF(mdchoicerotate.FItemList(i).Ftextinfo="" or isnull(mdchoicerotate.FItemList(i).Ftextinfo),"","TEXT : " & ReplaceBracket(mdchoicerotate.FItemList(i).Ftextinfo) & "<br>") %>
				LINK : <%= ReplaceBracket(mdchoicerotate.FItemList(i).Flinkinfo) %>
			</a>
			<% if mdchoicerotate.FItemList(i).Flinkitemid > 0 then %>
			<table cellpadding="1" cellspacing="1" class="a etcInfo" style="padding-top:15px;">
				<tr>
					<td>판매가 :</td>
					<td><%=FormatNumber(mdchoicerotate.FItemList(i).Forgprice,0)%> <%=mdchoicerotate.FItemList(i).saleCouponPriceCheck(mdchoicerotate.FItemList(i).Fsailyn , mdchoicerotate.FItemList(i).FitemCouponYn , mdchoicerotate.FItemList(i).Forgprice , mdchoicerotate.FItemList(i).Fsailprice , mdchoicerotate.FItemList(i).FitemCouponType)%></td>
				</tr>
				<tr>
					<td>마진 :</td>
					<td><%=fnPercent(mdchoicerotate.FItemList(i).Forgsuplycash,mdchoicerotate.FItemList(i).Forgprice,1)%> <%=mdchoicerotate.FItemList(i).priceMarginCheck( mdchoicerotate.FItemList(i).Fsailyn , mdchoicerotate.FItemList(i).FitemCouponYn , mdchoicerotate.FItemList(i).FitemCouponType , mdchoicerotate.FItemList(i).Fsailsuplycash , mdchoicerotate.FItemList(i).Fsailprice , mdchoicerotate.FItemList(i).Fcouponbuyprice , mdchoicerotate.FItemList(i).Fbuycash)%></td>
				</tr>
				<tr>
					<td>계약구분 :</td>
					<td><%=fnColor(mdchoicerotate.FItemList(i).Fmwdiv,"mw")%>-<%=mdchoicerotate.FItemList(i).deliveryTypeName(mdchoicerotate.FItemList(i).Fdeliverytype)%></td>
				</tr>
				<tr>
					<td>재고현황 :</td>
					<td><a href="javascript:popCurrentStock('<%= mdchoicerotate.FItemList(i).Flinkitemid %>');">[보기]</a></td>
				</tr>
			</table>
			<% end if %>
		</td>
		<td align="center">
			<input type="text" name="disporder" value="<%= mdchoicerotate.FItemList(i).FDisporder %>" size="3" style="text-align:center" class="text">
		</td>
		<td align="center" class="etcInfo">
			<%= formatdate(mdchoicerotate.FItemList(i).Fstartdate,"0000.00.00") %>
			<br>
			<%
				If cdate(mdchoicerotate.FItemList(i).Fstartdate) <= date() and  cdate(mdchoicerotate.FItemList(i).Fenddate) >= date()  Then
					Response.write " <span style=""color:red"">(적용중)</span>"					
				end If
			%>
		</td>
		<td align="center" class="etcInfo">
			<%= formatdate(mdchoicerotate.FItemList(i).Fenddate,"0000.00.00") %><br>
			<%
				If clng(datediff("d", now() , mdchoicerotate.FItemList(i).Fenddate)) < 0 Or clng(datediff("h", now() , mdchoicerotate.FItemList(i).Fenddate )) < 0  Then 
					Response.write " <span style=""color:red"">(종료)</span>"
				ElseIf cInt(datediff("d", mdchoicerotate.FItemList(i).Fenddate , now())) < 1  Then '오늘날짜
					If cInt(datediff("h", now() , mdchoicerotate.FItemList(i).Fenddate )) >= 0 And cInt(datediff("h", now() , mdchoicerotate.FItemList(i).Fenddate )) < 24 Then ' 오늘
						Response.write " <span style=""color:red"">(약 "& cInt(datediff("h", now() , mdchoicerotate.FItemList(i).Fenddate )) &" 시간후 종료)</span>"
					End If 
				End If 
			%>
		</td>		
		<td align="center" class="etcInfo">
			<%= FormatDateTime(mdchoicerotate.FItemList(i).Fregdate,2) %>
		</td>
		<td align="center" class="etcInfo">
			<%
				If mdchoicerotate.FItemList(i).FLowestPrice = "Y" Then
					Response.write "사용"
				Else
					Response.write "사용안함"
				End If
			%>
		</td>
		<td align="center" class="etcInfo">
			<% if mdchoicerotate.FItemList(i).IsSoldOut then %>
			<font color="red"><%=mdchoicerotate.FItemList(i).FSellyn%></font>
			<% else %>
			<font color="blue"><%=mdchoicerotate.FItemList(i).FSellyn%></font>
			<% end if %>
		</td>
		<td align="center" class="etcInfo"><%= mdchoicerotate.FItemList(i).Fregname %></td>
		<td align="center" class="etcInfo"><%= mdchoicerotate.FItemList(i).Fworkername %></td>
		<td align="center" class="etcInfo"><button class="button" onclick="delItem('<%=mdchoicerotate.FItemList(i).Fidx %>');return false;">상품삭제</button></td>
	</tr>
<% next %>
</tbody>
	<tr>
		<td colspan="13" align="center" bgcolor="white">
			<% if mdchoicerotate.HasPreScroll then %>
				<a href="javascript:NextPage('<%= mdchoicerotate.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + mdchoicerotate.StarScrollPage to mdchoicerotate.FScrollCount + mdchoicerotate.StarScrollPage - 1 %>
				<% if i>mdchoicerotate.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if mdchoicerotate.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>
</form>
<%
set mdchoicerotate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->