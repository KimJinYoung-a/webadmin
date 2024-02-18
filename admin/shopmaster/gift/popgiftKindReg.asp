<%@ language=vbscript %>
<% option explicit
	Response.Expires = -1440
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
%>
<%
'####################################################
' Description :  사은품 종류 등록
' History : 2008.04.02 정윤정 생성
'			2020.03.27 한용민 수정(사은품구분 체크 추가)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<%
Dim clsGiftkind, sViewMode, sMode, strTxt,strImg,iitemid,igkCode, iprd_itemgubun, iprd_itemid, iprd_itemoption
Dim arrList, intLoop, giftkind_linkGbn, bcouponidx, listCount, gift_delivery, gift_code, tmpTitle, eFolder
dim clsGift, makerid
Dim isCouponType : isCouponType = FALSE
	gift_delivery = requestCheckVar(request("gift_delivery"),1)
	gift_code = requestCheckVar(getNumeric(request("gift_code")),10)
	strTxt = requestCheckVar(Request("sGKN"),32)
	sViewMode  = requestCheckVar(Request("sVM"),10)
	igkCode = requestCheckVar(Request("iGK"),10)
	makerid = requestCheckVar(Request("makerid"),32)

IF sViewMode = "" THEN sViewMode = -1
sMode = "KI"
listCount = 30

 ' 검색하려는 사은품 명이 있을 때 해당 리스트 보여준다.
IF sViewMode < 0 THEN
	set clsGiftkind = new CGift
		clsGiftkind.FSearchTxt = strTxt
		clsGiftkind.FPSize = listCount
		arrList = clsGiftkind.fnGetGiftKind
	set clsGiftkind = nothing
END IF

IF (sViewMode > 0) or (igkCode<>"") THEN
	set clsGift = new CGift
		clsGift.FGCode = gift_code

		'사은품관리코드의 배송방법을 가져온다.
		if gift_code<>"" then
			clsGift.fnGetGiftConts

			if clsGift.ftotalcount>0 then
				gift_delivery  = clsGift.FGDelivery
			end if
		end if
	set clsGift = nothing

	set clsGiftkind = new CGift
		sMode = "KU"

		if (igkCode="") then
			igkCode = sViewMode
		end if
		clsGiftkind.FGKindCode = igkCode
		clsGiftkind.fnGetGiftKindConts

		strTxt = clsGiftkind.FGKindName
		strImg = clsGiftkind.FGKindImg
		iitemid= clsGiftkind.FItemid
		iprd_itemgubun= clsGiftkind.Fprd_itemgubun
		iprd_itemid= clsGiftkind.Fprd_itemid
		iprd_itemoption= clsGiftkind.Fprd_itemoption

		''2011-10추가 :: 차후 사은품관리 - 사은품 상품 재고 연동시 dbo.tbl_giftkind_option 사용.. :: eastone
		giftkind_linkGbn= clsGiftkind.Fgiftkind_linkGbn
		bcouponidx= clsGiftkind.Fbcouponidx
		isCouponType = (NULL2Blank(giftkind_linkGbn)="B")
	set clsGiftkind = nothing
END IF

eFolder =   igkCode
if gift_delivery="" then gift_delivery="N"
%>
<script language="javascript">
<!--

// 검색
function jsSearch(){
	/*
	if(!document.frmSearch.sGKN.value){
		alert("사은품종류명을 입력해주세요");
		return;
	}
	*/

	frmSearch.action="/admin/shopmaster/gift/popgiftKindReg.asp";
	frmSearch.target="";
	document.frmSearch.submit();
}


// 등록 또는 검색 화면으로 변경
function jsChangeMode(sViewMode){
	if (sViewMode ==""){
	document.frmSearch.sGKN.value="";
	}
	document.frmSearch.sVM.value = sViewMode;
	frmSearch.action="/admin/shopmaster/gift/popgiftKindReg.asp";
	frmSearch.target="";
	document.frmSearch.submit();
}

// 사은품 종류등록
function jsSubmitGiftKind(){
	var frm = document.frmGift;
	if(!frm.sGKN.value){
		 alert("사은품종류명을 입력해주세요");
		 frm.sGKN.focus();
		 return false;
	}

	// 배송방법 체크
	if (frm.giftkind_linkGbn[0].checked){
		<% if gift_delivery="N" then %>
			if (frm.prd_itemgubun.value=="" || frm.prd_itemid.value=="" || frm.prd_itemoption.value==""){
				alert("사은품 구분을 상품을 선택 하셨습니다. 물류코드를 입력해 주세요.");
				frm.prd_itemid.focus();
				return;
			}
		<% end if %>
	}else if(frm.giftkind_linkGbn[1].checked){
		if (!confirm('현재 전체 사은품 증정 이벤트에만 사용 가능합니다. \n\n계속하시겠습니까?')){
			return false;
		}
	}

	if (confirm('사은품 종류를 <%= CHKIIF(sMode = "KU","수정","신규 등록") %> 하시겠습니까?')){
		frm.submit();
	}
	//return;
}

//검색된 사은품종류 적용
function jsSetGiftKind(igk, skn,strImg,iid,gKLGbn,bcouponidx){
	opener.document.all.iGK.value = igk;
	opener.document.all.sGKN.value= skn;
	if (opener.document.all.giftkind_linkGbn){
		opener.document.all.giftkind_linkGbn.value= gKLGbn;
	}
	if (gKLGbn=='B'){
		if (opener.document.all.bcouponidx){
			opener.document.all.bcouponidx.value= bcouponidx;
		}
	}

	//if(iid!=""){ //??
	//	opener.document.all.sGKN.value= opener.document.all.sGKN.value+'['+iid+']';
	//}
	if(strImg !=""){
	opener.document.all.spanImg.innerHTML = "<a href=javascript:jsImgView('"+strImg+"')><img src='"+strImg+"' border=0></a>";
	}
	window.close();
}

//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/lib/showimage.asp?img='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

function fnAddImage2(strImg,sName,sSpan){
	document.domain ="10x10.co.kr";
	eval("document.frmGift." + sName).value = strImg;
	eval("document.all." + sSpan ).innerHTML = "<img src='"+strImg+"' border=0 width='60' height='30'>";
}

function jsSetImg2(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;

	winImg = window.open('popgiftkindupload.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();

	//winImg = window.open('/admin/eventmanage/common/pop_event_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	//winImg.focus();
}

function jsSetImg(){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/shopmaster/gift/popgiftkindupload.asp','popImg','width=370,height=150');
	winImg.focus();
}

function fnAddImage(strImg){
	document.domain ="10x10.co.kr";
	document.frmGift.sGKImg.value = strImg;
	document.all.spanImg.innerHTML = "<img src='"+strImg+"' border=0 width='60' height='30'>";
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

function dispGKGbn(comp){
	if (comp.value=='B'){
		document.getElementById("dpGKGbn_I1").style.display = "none";
		document.getElementById("dpGKGbn_I2").style.display = "none";
		document.getElementById("dpGKGbn_B").style.display = "";
	}else{
		document.getElementById("dpGKGbn_I1").style.display = "";
		document.getElementById("dpGKGbn_I2").style.display = "";
		document.getElementById("dpGKGbn_B").style.display = "none";
	}
}

function jsPopSearchGiftItem() {
	var pop;

	winImg = window.open("/admin/shopmaster/gift/popgiftitemlist.asp?itemgubun=85",'jsPopSearchGiftItem','width=1280,height=960,scrollbars=yes');
	winImg.focus();
}

// 사은품(85코드)자동생성
function autoGiftItemreg(){
	if (frmGift.sGKN.value==""){
		alert("사은품을 자동등록 하실려면, 사은품 종류명을 먼저 입력해 주세요.");
		frmGift.sGKN.focus();
		return;
	}
	frmgiftreg.giftkind_name.value=frmGift.sGKN.value;
	frmgiftreg.makerid.value=frmGift.makerid.value;
	if (frmgiftreg.makerid.value==""){
		alert("사은품을 자동등록 하실려면, 브랜드ID를 입력해 주세요.");
		frmgiftreg.makerid.focus();
		return;
	}

	var ret = confirm('신규사은품[85코드]을 자동생성 하시겠습니까?');
	if (ret){
		frmgiftreg.sM.value="regautogiftitem";
		frmgiftreg.action="/admin/shopmaster/gift/giftproc.asp";
		frmgiftreg.target="framegift";
		frmgiftreg.submit();
	
	}
}

function ReActWithThis(itemgubun, itemid, itemoption) {
	var frm = document.frmGift;

	frm.prd_itemgubun.value = itemgubun;
	frm.prd_itemid.value = itemid;
	frm.prd_itemoption.value = itemoption;
}

//-->
</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 사은품종류 <%= CHKIIF(sMode = "KU","수정","신규등록") %></div>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="0" >
<form name="frmSearch" method="get" action="" style="margin:0px;" >
<input type="hidden" name="sVM" >
<input type="hidden" name="gift_delivery" value="<%= gift_delivery %>" >
<input type="hidden" name="gift_code" value="<%= gift_code %>" >
<input type="hidden" name="makerid" value="<%= makerid %>" >
<tr>
	<td height="30">
		<% if igkCode<>"" or isArray(arrList) then %>
			사은품종류명 : <input type="text" class="text" name="sGKN" size="40" maxlength="60" value="<%=strTxt%>">
		<% else %>
			<input type="hidden" name="sGKN" size="40" maxlength="60" value="">
		<% end if %>
		<input type="button" class="button" value="기존사은품검색" onClick="jsSearch();">
	</td>
	<td align="right">
		<% if igkCode<>"" or isArray(arrList) then %>
			<input type="button" class="button" value="새로등록" onClick="jsChangeMode('0');">	
		<% end if %>
	</td>
</tr>
</form>
<tr>
	<td colspan="2"><hr wudth="100%"></td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%IF isArray(arrList) THEN %>
		<tr bgcolor="<%= adminColor("tabletop") %>">
			<td height="25" align="center" width="60">사은품코드</td>
			<td align="left">사은품종류명</td>
			<td align="center" width="40">사은품<br>구분</td>
			<td align="center" width="60">상품<br>연결코드</td>
			<td align="center" width="65">이미지</td>
			<td align="center" width="80">등록일</td>
			<td align="center" width="50">수정</td>
			<!--
			<td align="center" width="60">사용여부</td>
			-->
			<td align="center" width="100">등록자</td>
			<td align="center" width="50">비고</td>
		</tr>
	<%
		For intLoop =0 To UBound(arrList,2)
			tmpTitle = arrList(1,intLoop)
			if (Len(tmpTitle) > 35) then
				tmpTitle = Left(tmpTitle, 35) + "..."
			end if
	%>
		<tr bgcolor="#FFFFFF">
			<td height="33" align="center"><a href="javascript:jsChangeMode('<%=arrList(0,intLoop)%>')" title="사은품종류 선택"><%=arrList(0,intLoop)%></a></td>
			<td align="left"><a href="javascript:jsChangeMode('<%=arrList(0,intLoop)%>')" title="<%= arrList(1,intLoop) %>"><%= tmpTitle %></a></td>
			<td align="center"><a href="javascript:jsChangeMode('<%=arrList(0,intLoop)%>')" title="사은품종류 선택">
			    <% if (arrList(5,intLoop)="B") then %>
			        <font color="#F08080">쿠폰</font>
			    <% else %>
			        상품
			    <% end if %>
			</a></td>
			<td align="center"><a href="javascript:jsChangeMode('<%=arrList(0,intLoop)%>')" title="사은품종류 선택">
			    <% if (arrList(5,intLoop)="B") then %>
			    <%=arrList(6,intLoop)%>
			    <% else %>
			    <%=arrList(3,intLoop)%>
			    <% end if %>
			</a></td>
			<td align="center"><%IF arrList(2,intLoop) <> "" THEN%><a href="javascript:jsImgView('<%=arrList(2,intLoop)%>')" title="이미지 확대보기"><img src="<%=arrList(2,intLoop)%>" width="60" height="30" border="0"></a><%END IF%></td>
			<td align="center"><a href="javascript:jsChangeMode('<%=arrList(0,intLoop)%>')" title="사은품종류 선택"><%=FormatDate(arrList(4,intLoop),"0000.00.00")%></a></td>
			<td align="center"><input type="button" value="수정" class="button" onClick="jsChangeMode('<%=arrList(0,intLoop)%>');"></td>
			<!--
			<td align="center">
				<% if Not IsNull(arrList(8,intLoop)) then %>Y<% end if %>
			</td>
			-->
			<td align="center"><%=arrList(7,intLoop)%></td>
			<td align="center"><input type="button" value="선택" class="button" onClick="jsSetGiftKind(<%=arrList(0,intLoop)%>,'<%=arrList(1,intLoop)%>','<%=arrList(2,intLoop)%>','<%=arrList(3,intLoop)%>','<%=arrList(5,intLoop)%>','<%=arrList(6,intLoop)%>');"></td>

		</tr>
	<% Next	%>
<%ELSE%>

	<%IF sViewMode = -1 AND strTxt <> "" THEN %>
		<tr><td colspan="2" height="50" bgcolor="#FFFFFF"><font color="#E08050"><%=strTxt%></font><br>에 해당하는 사은품 종류가 없습니다. 새로 등록해 주세요</td></tr>
	<%END IF%>
		<form name="frmGift" method="post" action="/admin/shopmaster/gift/giftProc.asp" >
		<input type="hidden" name="sM" value="<%=sMode%>">
		<input type="hidden" name="sGKImg" value="<%=strImg%>">
		<input type="hidden" name="iGK" value="<%=igkCode%>">
		<input type="hidden" name="gift_code" value="<%= gift_code %>" >
		<tr>
			<td align="center" width="100" height="30" bgcolor="<%= adminColor("tabletop") %>">사은품코드</td>
			<td bgcolor="#FFFFFF"><%=igkCode%></td>
		</tr>
		<tr>
			<td align="center" width="100" height="30" bgcolor="<%= adminColor("tabletop") %>">사은품종류명</td>
			<td bgcolor="#FFFFFF">
				<input type="text" class="text" name="sGKN" size="50" maxlength="60" value="<%=strTxt%>">
			</td>
		</tr>
		<tr>
		    <td align="center" height="45" bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
		    <td bgcolor="#FFFFFF">
				<% drawSelectBoxDesignerwithName "makerid", makerid %>
		    </td>
		</tr>
		<tr>
		    <td align="center" height="45" bgcolor="<%= adminColor("tabletop") %>">배송방법</td>
		    <td bgcolor="#FFFFFF">
				<select class="select" name="gift_delivery">
					<option value="N" <%IF gift_delivery = "N" THEN%>selected<%END IF%>>텐바이텐배송</option>
					<option value="Y" <%IF gift_delivery = "Y" THEN%>selected<%END IF%>>업체배송</option>
				</select>
		    </td>
		</tr>
		<tr>
		    <td align="center" height="45" bgcolor="<%= adminColor("tabletop") %>">사은품구분</td>
		    <td bgcolor="#FFFFFF">
		        <input type="radio" name="giftkind_linkGbn" value="I" <%= CHKIIF(Not isCouponType,"checked","") %> onClick="dispGKGbn(this);"> 상품
		        <input type="radio" name="giftkind_linkGbn" value="B" <%= CHKIIF(isCouponType,"checked","") %> onClick="dispGKGbn(this);"> 보너스<font color="#F08080">쿠폰</font>
		        <br>(현재 보너스 쿠폰은 전체 증정 이벤트만 가능)
		    </td>
		</tr>
		<tr id="dpGKGbn_I1" <%= CHKIIF(isCouponType,"style='display:none'","") %>>
			<td align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">당첨상품코드</td>
			<td bgcolor="#FFFFFF"><input type="text" class="text" name="itemid" size="10" value="<%=iitemid%>"></td>
		</tr>
		<tr id="dpGKGbn_I2" <%= CHKIIF(isCouponType,"style='display:none'","") %>>
			<td align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">물류코드</td>
			<td bgcolor="#FFFFFF">
				<input type="text" class="text_ro" name="prd_itemgubun" size="2" value="<%= iprd_itemgubun %>" readonly>
				<input type="text" class="text_ro" name="prd_itemid" size="8" value="<%= iprd_itemid %>" readonly>
				<input type="text" class="text_ro" name="prd_itemoption" size="4" value="<%= iprd_itemoption %>" readonly>
				<input type="button" class="button" value="검색" onClick="jsPopSearchGiftItem();" >
				<input type="button" class="button" value="신규사은품(85코드)자동생성" onClick="autoGiftItemreg();" >
				<br>(* 물류에서 사은품을 배송하는 경우 입력하세요)
			</td> <!-- 재고 디비와 연동 -->
		</tr>
		<tr id="dpGKGbn_B" <%= CHKIIF(isCouponType,"style='display:block'","style='display:none'") %>>
			<td align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">연결<font color="#F08080">쿠폰</font>코드</td>
			<td bgcolor="#FFFFFF"><input type="text" name="bcouponidx" class="text" size="10" value="<%=bcouponidx%>"></td>
		</tr>
		<tr>
			<td align="center" height="45" bgcolor="<%= adminColor("tabletop") %>">이미지<br>(이벤트내 사은품)</td>
			<td bgcolor="#FFFFFF">
			    <input type="button" class="button" value="이미지등록" onClick="jsSetImg2('<%=eFolder%>','<%=strImg%>','sGKImg','spanImg');" >
			    <div id="spanImg">
			    <%IF strImg <> "" THEN%>
			    <a href="javascript:jsImgView('<%=strImg%>');"><img src="<%=strImg%>" width="60" height="30" border="0"></a>
			    <a href="javascript:jsDelImg('sGKImg','spanImg');"><img src="/images/icon_delete2.gif" border="0"></a>
			    <%END IF%>
			    </div>

		    </td>
		</tr>
		<tr>
			<td colspan="2" bgcolor="#FFFFFF" align="center">
			    <input type="button" class="button" value="등록" onClick="jsSubmitGiftKind();">
			    <!--<input type="image" src="/images/icon_confirm.gif">-->
				<!--<a href="javascript:history.back(0);"><img src="/images/icon_cancel.gif" border="0"></a>-->
			</td>
		</tr>
		</form>
<%END IF%>
	</table>
</td>
</tr>
</table>

<form name="frmgiftreg" method="post" action="" >
<input type="hidden" name="sM" value="" >
<input type="hidden" name="giftkind_name" value="" >
<input type="hidden" name="gift_delivery" value="<%= gift_delivery %>" >
<input type="hidden" name="gift_code" value="<%= gift_code %>" >
<input type="hidden" name="makerid" value="<%= makerid %>" >
</form>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="framegift" name="framegift" src="" width="100%" height="300" frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="framegift" name="framegift" src="" width="0" height="0" frameborder="0" scrolling="no"></iframe>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
