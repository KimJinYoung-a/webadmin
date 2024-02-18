<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : mdpick_insert.asp
' Discription : 모바일 mdpick
' History : 2014.01.28 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/new_mdpickCls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate , lp , ii
Dim sDt, sTm, eDt, eTm , gubun , mdpicktitle , prevDate , topview , userlevelgubun
	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then
	mode = "add"
Else
	mode = "modify"
End If

If idx <> "" then
	dim mdpickList
	set mdpickList = new Cmdpick
	mdpickList.FRectIdx = idx
	mdpickList.GetOneContents()

	gubun			=	mdpickList.FOneItem.Fgubun
	mdpicktitle		=	mdpickList.FOneItem.Fmdpicktitle
	mainStartDate	=	mdpickList.FOneItem.Fstartdate
	mainEndDate		=	mdpickList.FOneItem.Fenddate
	isusing			=	mdpickList.FOneItem.Fisusing
	topview			=	mdpickList.FOneItem.Ftopview
	userlevelgubun	=	mdpickList.FOneItem.Fuserlevelgubun

	set mdpickList = Nothing
End If

If Isnull(topview) Or topview = "0" Or  topview = "" Then topview = 4

Dim oSubItemList
set oSubItemList = new Cmdpick
	oSubItemList.FPageSize = 100
	oSubItemList.FRectlistIdx = idx
	If idx <> "" then
		oSubItemList.GetContentsItemList()
	End If


if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = date
	end if
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = date
	end if
	eTm = "23:59:59"
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (frm.sTm.value.length != 8) {
			alert("시간을 정확히 입력하세요");
			frm.sTm.focus();
			return;
		}

		if (frm.eTm.value.length != 8) {
			alert("시간을 정확히 입력하세요");
			frm.eTm.focus();
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/today_mdpick/";
	}
	$(function(){
	//달력대화창 설정
	var arrDayMin = ["일","월","화","수","목","금","토"];
	var arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
    	showCurrentAtPos: 1,
      	showOn: "button",
      	<% if Idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
    	}
    });
    $("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
      	showOn: "button",
      	<% if Idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });

	//라디오버튼
	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
	$(".rdoIsLowestPrice").buttonset().children().next().attr("style","font-size:11px;");	

	for ( ii = 1 ; ii<=4 ; ii++ )
	{
		$( "#subList"+ii ).sortable({
			placeholder: "ui-state-highlight",
			start: function(event, ui) {
				ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
				$(this).find('.etcInfo').hide();
			},
			stop: function(){
				var i=99999;
				$(this).find("input[name^='sort']").each(function(){
					if(i>$(this).val()) i=$(this).val()
				});
				if(i<=0) i=1;
				$(this).find("input[name^='sort']").each(function(){
					$(this).val(i);
					i++;
				});
				$(this).find('.etcInfo').show();
			}
		});
	}
});

//소재
function popSubEdit(subidx,v,t) {
<% if idx <>"" then %>
    var popwin = window.open('popSubItemEdit.asp?listIdx=<%=idx%>&subIdx='+subidx+'&gubun='+v+'&topview='+t,'popTemplateManage','width=800,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품검색 일괄 등록 (구버전)
function popRegSearchItem(v,t) {
<% if idx <> "" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/mobile/today_mdpick/doSubRegItemCdArray.asp?listidx=<%=idx%>||gubun="+v+'||topview='+t, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품검색 일괄 등록 (신버전)
function popRegMdPickRcmdItem(v,t) {
<% if idx <> "" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAutoPick.asp?sellyn=Y&usingyn=Y&defaultmargin=0&gubun="+v+"&acURL=/admin/mobile/today_mdpick/doSubRegItemCdArray.asp?listidx=<%=idx%>||gubun="+v+'||topview='+t, "popup_item", "width=1150,height=600,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품코드 일괄 등록
function popRegArrayItem(v,t) {
<% if idx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?listIdx=<%=idx%>&gubun='+v+'&topview='+t,'popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품코드 일괄 등록
function popRegArrayDealItem(v,t) {
<% if idx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?dealyn=Y&listIdx=<%=idx%>&gubun='+v+'&topview='+t,'popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 소재를 선택해주세요.");
		return;
	}
	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}

// 상품이미지 수정
function editItemImage(itemid) {
	var param = "itemid=" + itemid;
	popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=1000,height=900,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//'아이템 삭제
function itemdel(v){
	if (confirm("상품이 삭제됩니다 삭제 하시겠습니까?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.mode.value = "itemdel";
		document.frmdel.action="doListModify.asp";
		document.frmdel.submit();
	}
}

function popCurrentStock(itemid) {
	var popwin;
    var popwin = window.open('/admin/stock/itemcurrentstock.asp?menupos=<%= request("menupos") %>&itemid='+itemid,'popRegArray','width=1024,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>
<form name="frmdel" method="POST" action="">
<input type="hidden" name="mode" />
<input type="hidden" name="chkIdx" />
</form>
<table width="1140" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="domdpick.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<tr bgcolor="#FFFFFF">
	<% If idx = ""  Then %>
	<td colspan="2" align="center" height="35">등록 진행 중 입니다.</td>
	<% Else %>
	<td bgcolor="#FFF999" colspan="2" align="center" height="35">수정 진행 중 입니다.</td>
	<% End If %>
</tr>
<% If idx <> ""  Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">idx</td>
	<td><%=idx%></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">노출기간</td>
    <td>
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">제목</td>
    <td>
		<input type="text" name="mdpicktitle" size="100" value="<%If idx="" then%><%=sDt%>&nbsp;시작 mdpick 입니다<% Else %><%=mdpicktitle%><% End if%>" />
    </td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">가격구분</td>
	<td><div style="float:left;"><input type="radio" name="userlevelgubun" value="1" <%=chkiif(userlevelgubun = "1","checked","")%> checked />기존 &nbsp;&nbsp;&nbsp; <input type="radio" name="userlevelgubun" value="2" <%=chkiif(userlevelgubun = "2","checked","")%>/>VIP이상</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>

<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2">
		<input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/>
	</td>
</tr>
</form>
</table>

<%
	If idx <> "" then
%>
<p><b>▶ 소재 정보</b></p>
<!-- // 등록된 소재 목록 -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="1140" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="9">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	총 <%=oSubItemList.FTotalCount%> 건 /
		    	<input type="button" value="전체선택" class="button" onClick="chkAllItem()">
		    	<input type="button" value="상태저장" class="button" onClick="saveList()" title="표시순서 및 사용여부를 일괄저장합니다.">
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="30" />
<col width="30" />
<col width="30" />
<col width="30" />
<col width="400" />
<col width="40" />
<col width="90" />
<col width="105" />
<col width="30" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>소재번호</td>
    <td>이미지</td>
    <td>상품코드</td>
    <td>상품정보</td>
    <td>표시순서</td>
    <td>사용여부</td>
    <td>최저가 여부</td>	
    <td>상품삭제</td>
</tr>
<% For ii = 1 To 4%>
<% If  ii = 1 Then %><tr><td align="center" bgcolor="#F3F3F3" colspan="9">MD`S PICK</td></tr><% End If %>
<% If  ii = 2 Then %><tr><td align="center" bgcolor="#F3F3F3" colspan="9">NEW</td></tr><% End If %>
<% If  ii = 3 Then %><tr><td align="center" bgcolor="#F3F3F3" colspan="9">BEST (2017 리뉴얼후 사용안함)</td></tr><% End If %>
<% If  ii = 4 Then %><tr><td align="center" bgcolor="#F3F3F3" colspan="9">SALE</td></tr><% End If %>
<tr bgcolor="#FFFFFF">
	<td colspan="9">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="right">
		    	<input type="button" value="상품코드로 등록" class="button" onClick="popRegArrayItem('<%=ii%>','<%=topview%>')" />
				<input type="button" value="딜상품 등록" class="button" onClick="popRegArrayDealItem('<%=ii%>','<%=topview%>')" />
		    	<input type="button" value="상품 추가<%=chkiif(ii>1 ,"(구)","")%>" class="button" onClick="popRegSearchItem('<%=ii%>','<%=topview%>')" />
				<% If ii > 1 Then %>
				<input type="button" value="추천 상품 추가" class="button" onClick="popRegMdPickRcmdItem('<%=ii%>','<%=topview%>')" />
				<% End If %>
		    	<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('','<%=ii%>','<%=topview%>')" style="cursor:pointer;" align="absmiddle">
		    </td>
		</tr>
		</table>
	</td>
</tr>
<tbody id="subList<%=ii%>">
<% For lp=0 to oSubItemList.FResultCount-1 %>
<% If oSubItemList.FItemList(lp).Fgubun = ii Then %>
<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","#FFFFFF","#F3F3F3")%>#FFFFFF">
    <td><input type="checkbox" name="chkIdx" value="<%=oSubItemList.FItemList(lp).FsubIdx%>" /><input type="hidden" name="chkgubun<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fgubun%>"/></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>,'<%=ii%>','<%=topview%>')" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FsubIdx%></td>
    <td style="cursor:pointer;">	
    <%		
		if oSubItemList.FItemList(lp).FFrontImage <> "" then
			Response.Write "<img src='" & oSubItemList.FItemList(lp).FFrontImage & "' height='50' />"
		else
			response.write "<img src='" & chkiif(oSubItemList.FItemList(lp).FTentenImg = "" or isnull(oSubItemList.FItemList(lp).FTentenImg), oSubItemList.FItemList(lp).FsmallImage , oSubItemList.FItemList(lp).FTentenImg ) & "' height='50' />"
		end if
    %>
	<br/><input type="button" onclick="editItemImage(<%=oSubItemList.FItemList(lp).FItemid%>,'<%=ii%>','<%=topview%>')" value="수정"/>
    </td>
    <td>
    <%
    	if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
    		Response.Write "<input type='text' value='" & oSubItemList.FItemList(lp).FItemid & "' readonly size='5'/>"
    	end if
    %>
    </td>
	<td style="text-align:left;">
		<input type="text" name="itemname<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fitemname%>" size="60">
		<table cellpadding="1" cellspacing="1" class="a etcInfo" style="padding-top:15px;">
			<tr>
				<td>판매가 :</td>
				<td><%=FormatNumber(oSubItemList.FItemList(lp).Forgprice,0)%> <%=oSubItemList.FItemList(lp).saleCouponPriceCheck(oSubItemList.FItemList(lp).Fsailyn , oSubItemList.FItemList(lp).FitemCouponYn , oSubItemList.FItemList(lp).Forgprice , oSubItemList.FItemList(lp).Fsailprice , oSubItemList.FItemList(lp).FitemCouponType)%></td>
			</tr>
			<tr>
				<td>마진 :</td>
				<td><%=fnPercent(oSubItemList.FItemList(lp).Forgsuplycash,oSubItemList.FItemList(lp).Forgprice,1)%> <%=oSubItemList.FItemList(lp).priceMarginCheck( oSubItemList.FItemList(lp).Fsailyn , oSubItemList.FItemList(lp).FitemCouponYn , oSubItemList.FItemList(lp).FitemCouponType , oSubItemList.FItemList(lp).Fsailsuplycash , oSubItemList.FItemList(lp).Fsailprice , oSubItemList.FItemList(lp).Fcouponbuyprice , oSubItemList.FItemList(lp).Fbuycash)%></td>
			</tr>
			<tr>
				<td>계약구분 :</td>
				<td><%=fnColor(oSubItemList.FItemList(lp).Fmwdiv,"mw")%>-<%=oSubItemList.FItemList(lp).deliveryTypeName(oSubItemList.FItemList(lp).Fdeliverytype)%></td>
			</tr>
			<tr>
				<td>재고현황 :</td>
				<td><a href="javascript:popCurrentStock('<%= oSubItemList.FItemList(lp).FItemid %>');">[보기]</a></td>
			</tr>
		</table>
	</td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">사용</label><input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">삭제</label>
		</span>
    </td>
    <td>
		<span class="rdoIsLowestPrice">
		<input type="radio" name="lowestPrice<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoIsLowestPrice<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FLowestPrice="Y","checked","")%> /><label for="rdoIsLowestPrice<%=lp%>_1">사용</label><input type="radio" name="lowestPrice<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoIsLowestPrice<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FLowestPrice="N" or oSubItemList.FItemList(lp).FLowestPrice="","checked","")%> /><label for="rdoIsLowestPrice<%=lp%>_2">사용안함</label>
		</span>
    </td>	
	<td><input type="button" value="상품삭제" onclick="itemdel('<%=oSubItemList.FItemList(lp).FsubIdx%>');"/></td>
</tr>
<% End If %>
<% Next %>
</tbody>
<% Next %>
</table>
</form>
<%
	End If
	set oSubItemList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
