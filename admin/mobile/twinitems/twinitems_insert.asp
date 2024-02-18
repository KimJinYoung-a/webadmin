<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : itemtwins_insert.asp
' Discription : 모바일 단품배너
' History : 2017-08-02 이종화 생성
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/today_twinitemsCls.asp" -->
<%
Dim mode
Dim idx
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim ordertext , isusing
Dim stdt , eddt
Dim L_img , L_maincopy , L_itemname	, L_itemid , L_newbest , R_img , R_maincopy	, R_itemname , R_itemid	, R_newbest , iteminfo

idx = requestCheckvar(request("idx"),16)

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 


'// 수정시
If idx <> "" then
	dim twinitemsOne
	set twinitemsOne = new CMainbanner
	twinitemsOne.FRectIdx = idx
	twinitemsOne.GetOneContents()

	mainStartDate		=	twinitemsOne.FOneItem.Fstartdate
	mainEndDate			=	twinitemsOne.FOneItem.Fenddate 
	isusing				=	twinitemsOne.FOneItem.Fisusing
	ordertext			=	twinitemsOne.FOneItem.Fordertext
	L_img				=	twinitemsOne.FOneItem.FL_img		
	L_maincopy			=	twinitemsOne.FOneItem.FL_maincopy
	L_itemname			=	twinitemsOne.FOneItem.FL_itemname
	L_itemid			=	twinitemsOne.FOneItem.FL_itemid	
	L_newbest			=	twinitemsOne.FOneItem.FL_newbest	
	R_img				=	twinitemsOne.FOneItem.FR_img		
	R_maincopy			=	twinitemsOne.FOneItem.FR_maincopy
	R_itemname			=	twinitemsOne.FOneItem.FR_itemname
	R_itemid			=	twinitemsOne.FOneItem.FR_itemid	
	R_newbest			=	twinitemsOne.FOneItem.FR_newbest	
	iteminfo			=	twinitemsOne.FOneItem.Fiteminfo	
	set twinitemsOne = Nothing

	Dim ii
	if not isnull(iteminfo) then 
		If ubound(Split(iteminfo,"^^")) > 0 Then ' 이미지 2개 정보
			For ii = 0 To ubound(Split(iteminfo,"^^"))
				If CStr(L_itemid) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) And L_img = (staticImgUrl & "/mobile/twinitems") Then
					L_img =  webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(L_itemid) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If

				If CStr(R_itemid) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) And R_img = (staticImgUrl & "/mobile/twinitems") Then
					R_img =  webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(R_itemid) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If
			Next 
		End If 
	end if
End If 

dim dateOption
dateOption = request("dateoption")

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
elseif dateOption <> "" then
	sDt = dateOption
	sTm = "00:00:00"
else
	sDt = date
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
elseif dateOption <> "" then	
	eDt = dateOption
	eTm = "23:59:59"
else
	eDt = date
	eTm = "23:59:59"
end If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (frm.L_itemid.value == ""){
			alert("좌측 상품코드를 넣어주세요.");
			return;
		}

		if (frm.L_itemname.value == ""){
			alert("좌측 상품이름을 넣어주세요.");
			frm.L_itemname.focus();
			return;
		}

		if (frm.L_maincopy.value == ""){
			alert("좌측 서브카피를 넣어주세요.");
			frm.L_maincopy.focus();
			return;
		}

		if (frm.R_itemid.value == ""){
			alert("우측 상품코드를 넣어주세요.");
			return;
		}

		if (frm.R_itemname.value == ""){
			alert("우측 상품이름을 넣어주세요.");
			frm.R_itemname.focus();
			return;
		}

		if (frm.R_maincopy.value == ""){
			alert("우측 서브카피를 넣어주세요.");
			frm.R_maincopy.focus();
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/twinitems/";
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

});


function chgtype(v){
	if (v == "1"){
		$("#additem1").css("display","none");
		$("#additem2").css("display","none");
		$("#additem3").css("display","none");
	}else{
		$("#additem1").css("display","");
		$("#additem2").css("display","");
	}
}

// 상품정보 접수
function fnGetItemInfo(iid,v) {
	if (iid != "")
	{
		$.ajax({
			type: "GET",
			url: "/admin/sitemaster/wcms/act_iteminfo.asp?itemid="+iid,
			dataType: "xml",
			cache: false,
			async: false,
			timeout: 5000,
			beforeSend: function(x) {
				if(x && x.overrideMimeType) {
					x.overrideMimeType("text/xml;charset=euc-kr");
				}
			},
			success: function(xml) {
				if($(xml).find("itemInfo").find("item").length>0) {
					var rst = $(xml).find("itemInfo").find("item").find("itemname").text();
					//$("#lyItemInfo"+v).fadeIn();
					$("#lyItemInfo"+v).text(rst);
				} else {
					//$("#lyItemInfo"+v).fadeOut();
				}
			},
			error: function(xhr, status, error) {
				alert("상품번호를 다시 넣어 주세요");
				return; 
				// $("#lyItemInfo"+v).fadeOut();
				/*alert(xhr + '\n' + status + '\n' + error);*/
			}
		});
	}
}
</script>
<table width="90%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/twinitems_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
	<td colspan="4" bgcolor="#FFF999" align="center"><%=chkiif(mode="add","입력페이지 입니다.","수정페이지 입니다.")%></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="5%">노출기간</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="5%">좌측형<br/><br/>배너정보</td>
	<td width="40%">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align="left">
					상품코드 : <input type="text" name="L_itemid" value="<%=L_itemid%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'1')" title="상품코드" />
					<br/>
					<% If L_img <> "" Then %>
					<img src="<%=L_img%>" width="120" height="120"/>
					<% Else %>
					<img src="/images/admin_login_logo2.png" width="120" height="120" /></br>이미지를 등록 해주세요.
					<% End If %>
				</td>
				<td align="right">
					메인카피 : <input type="text" name="L_maincopy" value="<%=L_maincopy%>" maxlength="10" size="20"/>
					<br><font color="red"><strong>※ 최대 10자 제한 ※</strong></font>
					<br/><br/>
					상품명&nbsp; : &nbsp;&nbsp;<input type="text" name="L_itemname" value="<%=L_itemname%>" size="20" maxlength="8"/>
					<br><font color="red"><strong>※ 최대 8자 제한 ※</strong></font>
					<br/><br/><span id="lyItemInfo1"></span>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<%=L_img%><br/>
					이미지 등록 : <input type="file" name="L_img" class="file" title="이벤트 #1" require="N" style="width:80%;" />
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="L_newbest" value="0" checked/> 사용안함&nbsp;&nbsp;&nbsp;<input type="radio" name="L_newbest" value="1" <%=chkiif(L_newbest="1","checked","")%>/> NEW&nbsp;&nbsp;&nbsp; <input type="radio" name="L_newbest" value="2" <%=chkiif(L_newbest="2","checked","")%>/> BEST
				</td>
				<td align="right" width="50%" style="padding-right:30px;">
					<input type="checkbox" name="L_delimg" value="Y" id="L_delimg"/> <label for="L_delimg">이미지 삭제</label>
				</td>
			</tr>
		</table>
	</td>
	<td bgcolor="#FFF999" align="center" width="5%">우측형<br/><br/>배너정보</td>
	<td width="40%">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align="left">
					메인카피 : <input type="text" name="R_maincopy" value="<%=R_maincopy%>"  maxlength="10" size="20"/>
					<br><font color="red"><strong>※ 최대 10자 제한 ※</strong></font>
					<br/><br/>
					상품명&nbsp; : <input type="text" name="R_itemname" value="<%=R_itemname%>" size="20" maxlength="8" />
					<br><font color="red"><strong>※ 최대 8자 제한 ※</strong></font>
					<br/><br/><span id="lyItemInfo2"></span>
				</td>
				<td align="right">
					상품코드 : <input type="text" name="R_itemid" value="<%=R_itemid%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'2')" title="상품코드" />
					<br/>
					<% If R_img <> "" Then %>
					<img src="<%=R_img%>" width="120" height="120"/>
					<% Else %>
					<img src="/images/admin_login_logo2.png" width="120" height="120" /></br>이미지를 등록 해주세요.
					<% End If %>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<%=R_img%><br/>
					이미지 등록 : <input type="file" name="R_img" class="file" title="이벤트 #1" require="N" style="width:80%;" />
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="R_newbest" value="0" checked/> 사용안함&nbsp;&nbsp;&nbsp;<input type="radio" name="R_newbest" value="1" <%=chkiif(R_newbest="1","checked","")%>/> NEW&nbsp;&nbsp;&nbsp; <input type="radio" name="R_newbest" value="2" <%=chkiif(R_newbest="2","checked","")%>/> BEST
				</td>
				<td align="right" style="padding-right:30px;">
					<input type="checkbox" name="R_delimg" value="Y" id="R_delimg"/> <label for="R_delimg">이미지 삭제</label>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->