<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script language="JavaScript" src="/js/common.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
	<!--[if IE]>
		<link rel="stylesheet" type="text/css" href="http://<%=CHKIIF(application("Svr_Info")="Dev","2013","")%>www.10x10.co.kr/lib/css/ie.css" />
	<![endif]-->

<link rel="stylesheet" type="text/css" href="http://www.10x10.co.kr/test/renewal/default.css" />
<link rel="stylesheet" type="text/css" href="http://www.10x10.co.kr/test/renewal/common.css" />
<script>
document.domain = "10x10.co.kr";

function popCate(c,d){
	var catepop = window.open("pop_cate.asp?catecode="+d+"&depth=3&inputname="+c+"","catepop","width=850,height=527, scrollbars=yes, resizable=yes");
	catepop.focus();
}

function jsItemReg(i){
	var itempop = window.open("pop_item.asp?catecode=108&itemid="+i+"","itempop","width=400,height=300, scrollbars=yes, resizable=yes");
	itempop.focus();
}

//브랜드 ID 검색 팝업창
function jsSearchBrandIDn(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/member/popBrandSearch.asp?isjsdomain=o&frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsRealServerReg(){
<%
	IF application("Svr_Info") = "Dev" THEN
		vWWW = "http://2013www.10x10.co.kr"
	Else
		vWWW = "http://www1.10x10.co.kr"
	End IF
%>
    var popCreateTemp = window.open("<%=vWWW%>/chtml/dispcate/menu_make_xml.asp?catecode=108&gb=temp","popCreateTemp","width=1200 height=930 scrollbars=yes resizable=yes");
	popCreateTemp.focus();
}

function jsSaveCateMenu(){
	for(var i=1; i<45; i++){
		if($("input[name=cate"+i+"]").val() == ""){
			alert("No."+i+" 선택하세요.");
			return;
		}
	}
	if($("input[name=itemid]").val() == ""){
		alert("Book을 등록해주세요.");
		return;
	}
	
	$("input[name=cate42code]").val($("input[name=cate42]").val());
	$("input[name=cate43code]").val($("input[name=cate43]").val());
	$("input[name=cate44code]").val($("input[name=cate44]").val());
	
	frmMenu.submit();
}
</script>
</head>
<body bgcolor="#F4F4F4">

<form name="frmMenu" method="post" action="templete_proc.asp" style="margin:0px;">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="cnt" value="44">
<div style="position:relative;height:370px;top:-20px;">
	<div class="gnbSubWrap col06" style="display:block;">
		<div class="gnbSub">
			<div class="ftLt fst" style="width:166px;">
				<dl>
					<dt><p>디자이너 브랜드</p></dt>
					<dd>
						<ul>
							<li>1. <input type="text" name="cate1" value="<%=cate(0)%>" size="17" style="cursor:pointer;" onClick="popCate('cate1','108101');" readonly></li>
							<li>2. <input type="text" name="cate2" value="<%=cate(1)%>" size="17" style="cursor:pointer;" onClick="popCate('cate2','108101');" readonly></li>
							<li>3. <input type="text" name="cate3" value="<%=cate(2)%>" size="17" style="cursor:pointer;" onClick="popCate('cate3','108101');" readonly></li>
							<li>4. <input type="text" name="cate4" value="<%=cate(3)%>" size="17" style="cursor:pointer;" onClick="popCate('cate4','108101');" readonly></li>
							<li>5. <input type="text" name="cate5" value="<%=cate(4)%>" size="17" style="cursor:pointer;" onClick="popCate('cate5','108101');" readonly></li>
							<li>6. <input type="text" name="cate6" value="<%=cate(5)%>" size="17" style="cursor:pointer;" onClick="popCate('cate6','108101');" readonly></li>
						</ul>
					</dd>
				</dl>
			</div>
			<div class="ftLt">
				<dl>
					<dt><p>Women</p></dt>
					<dd>
						<ul>
							<li>7. <input type="text" name="cate7" value="<%=cate(6)%>" size="17" style="cursor:pointer;" onClick="popCate('cate7','108102');" readonly></li>
							<li>8. <input type="text" name="cate8" value="<%=cate(7)%>" size="17" style="cursor:pointer;" onClick="popCate('cate8','108102');" readonly></li>
							<li>9. <input type="text" name="cate9" value="<%=cate(8)%>" size="17" style="cursor:pointer;" onClick="popCate('cate9','108102');" readonly></li>
						</ul>
					</dd>
				</dl>
				<dl>
					<dt><p>Men</p></dt>
					<dd>
						<ul>
							<li>10. <input type="text" name="cate10" value="<%=cate(9)%>" size="17" style="cursor:pointer;" onClick="popCate('cate10','108103');" readonly></li>
							<li>11. <input type="text" name="cate11" value="<%=cate(10)%>" size="17" style="cursor:pointer;" onClick="popCate('cate11','108103');" readonly></li>
							<li>12. <input type="text" name="cate12" value="<%=cate(11)%>" size="17" style="cursor:pointer;" onClick="popCate('cate12','108103');" readonly></li>
						</ul>
					</dd>
				</dl>
			</div>
			<div class="ftLt">
				<dl>
					<dt><p>주얼리</p></dt>
					<dd>
						<ul>
							<li>13. <input type="text" name="cate13" value="<%=cate(12)%>" size="17" style="cursor:pointer;" onClick="popCate('cate13','108104');" readonly></li>
							<li>14. <input type="text" name="cate14" value="<%=cate(13)%>" size="17" style="cursor:pointer;" onClick="popCate('cate14','108104');" readonly></li>
							<li>15. <input type="text" name="cate15" value="<%=cate(14)%>" size="17" style="cursor:pointer;" onClick="popCate('cate15','108104');" readonly></li>
							<li>16. <input type="text" name="cate16" value="<%=cate(15)%>" size="17" style="cursor:pointer;" onClick="popCate('cate16','108104');" readonly></li>
							<li>17. <input type="text" name="cate17" value="<%=cate(16)%>" size="17" style="cursor:pointer;" onClick="popCate('cate17','108104');" readonly></li>
						</ul>
					</dd>
				</dl>
				<dl>
					<dt><p>우산/레인코트</p></dt>
					<dd>
						<ul>
							<li>18. <input type="text" name="cate18" value="<%=cate(17)%>" size="17" style="cursor:pointer;" onClick="popCate('cate18','108106');" readonly></li>
							<li>19. <input type="text" name="cate19" value="<%=cate(18)%>" size="17" style="cursor:pointer;" onClick="popCate('cate19','108106');" readonly></li>
							<li>20. <input type="text" name="cate20" value="<%=cate(19)%>" size="17" style="cursor:pointer;" onClick="popCate('cate20','108106');" readonly></li>
							<li>21. <input type="text" name="cate21" value="<%=cate(20)%>" size="17" style="cursor:pointer;" onClick="popCate('cate21','108106');" readonly></li>
						</ul>
					</dd>
				</dl>
			</div>
			<div class="ftLt">
				<dl>
					<dt><p>패션 액세서리</p></dt>
					<dd>
						<ul>
							<li>22. <input type="text" name="cate22" value="<%=cate(21)%>" size="17" style="cursor:pointer;" onClick="popCate('cate22','108105');" readonly></li>
							<li>23. <input type="text" name="cate23" value="<%=cate(22)%>" size="17" style="cursor:pointer;" onClick="popCate('cate23','108105');" readonly></li>
							<li>24. <input type="text" name="cate24" value="<%=cate(23)%>" size="17" style="cursor:pointer;" onClick="popCate('cate24','108105');" readonly></li>
							<li>25. <input type="text" name="cate25" value="<%=cate(24)%>" size="17" style="cursor:pointer;" onClick="popCate('cate25','108105');" readonly></li>
							<li>26. <input type="text" name="cate26" value="<%=cate(25)%>" size="17" style="cursor:pointer;" onClick="popCate('cate26','108105');" readonly></li>
							<li>27. <input type="text" name="cate27" value="<%=cate(26)%>" size="17" style="cursor:pointer;" onClick="popCate('cate27','108105');" readonly></li>
							<li>28. <input type="text" name="cate28" value="<%=cate(27)%>" size="17" style="cursor:pointer;" onClick="popCate('cate28','108105');" readonly></li>
							<li>29. <input type="text" name="cate29" value="<%=cate(28)%>" size="17" style="cursor:pointer;" onClick="popCate('cate29','108105');" readonly></li>
							<li>30. <input type="text" name="cate30" value="<%=cate(29)%>" size="17" style="cursor:pointer;" onClick="popCate('cate30','108105');" readonly></li>
						</ul>
					</dd>
				</dl>
			</div>
			<div class="ftLt">
				<dl>
					<dt><p>뷰티/바디</p></dt>
					<dd>
						<ul>
							<li>31. <input type="text" name="cate31" value="<%=cate(30)%>" size="17" style="cursor:pointer;" onClick="popCate('cate31','108107');" readonly></li>
							<li>32. <input type="text" name="cate32" value="<%=cate(31)%>" size="17" style="cursor:pointer;" onClick="popCate('cate32','108107');" readonly></li>
							<li>33. <input type="text" name="cate33" value="<%=cate(32)%>" size="17" style="cursor:pointer;" onClick="popCate('cate33','108107');" readonly></li>
							<li>34. <input type="text" name="cate34" value="<%=cate(33)%>" size="17" style="cursor:pointer;" onClick="popCate('cate34','108107');" readonly></li>
							<li>35. <input type="text" name="cate35" value="<%=cate(34)%>" size="17" style="cursor:pointer;" onClick="popCate('cate35','108107');" readonly></li>
							<li>36. <input type="text" name="cate36" value="<%=cate(35)%>" size="17" style="cursor:pointer;" onClick="popCate('cate36','108107');" readonly></li>
							<li>37. <input type="text" name="cate37" value="<%=cate(36)%>" size="17" style="cursor:pointer;" onClick="popCate('cate37','108107');" readonly></li>
							<li>38. <input type="text" name="cate38" value="<%=cate(37)%>" size="17" style="cursor:pointer;" onClick="popCate('cate38','108107');" readonly></li>
							<li>39. <input type="text" name="cate39" value="<%=cate(38)%>" size="17" style="cursor:pointer;" onClick="popCate('cate39','108107');" readonly></li>
							<li>40. <input type="text" name="cate40" value="<%=cate(39)%>" size="17" style="cursor:pointer;" onClick="popCate('cate40','108107');" readonly></li>
							<li>41. <input type="text" name="cate41" value="<%=cate(40)%>" size="17" style="cursor:pointer;" onClick="popCate('cate41','108107');" readonly></li>
						</ul>
					</dd>
				</dl>
			</div>
			<div class="ftLt">
				<dl>
					<dt><p>BRANDS</p></dt>
					<dd>
						<ul>
							<li>42. <input type="text" name="cate42" value="<%=cate(41)%>" size="17" style="cursor:pointer;" onClick="jsSearchBrandIDn(this.form.name,'cate42');" readonly></li>
							<li>43. <input type="text" name="cate43" value="<%=cate(42)%>" size="17" style="cursor:pointer;" onClick="jsSearchBrandIDn(this.form.name,'cate43');" readonly></li>
							<li>44. <input type="text" name="cate44" value="<%=cate(43)%>" size="17" style="cursor:pointer;" onClick="jsSearchBrandIDn(this.form.name,'cate44');" readonly></li>
						</ul>
					</dd>
				</dl>
				<dl class="cBnrView">
					<dt><p><span><a href="">BOOK</a></span> <a href="javascript:jsItemReg('<%=vItemID%>');">[등록]</a></p></dt>
					<dd>
						<p class="cBnrImg">
						<input type="hidden" name="itemid" value="<%=vItemID%>">
						<input type="hidden" name="imglink" value="<%=vImgLink%>">
						<span id="itemidimg"><img src="<%=vImgLink%>"></span>
						</p>
					</dd>
				</dl>
			</div>
		</div>
	</div>
</div>
<input type="hidden" name="cate1code" value="<%=catecode(0)%>">
<input type="hidden" name="cate2code" value="<%=catecode(1)%>">
<input type="hidden" name="cate3code" value="<%=catecode(2)%>">
<input type="hidden" name="cate4code" value="<%=catecode(3)%>">
<input type="hidden" name="cate5code" value="<%=catecode(4)%>">
<input type="hidden" name="cate6code" value="<%=catecode(5)%>">
<input type="hidden" name="cate7code" value="<%=catecode(6)%>">
<input type="hidden" name="cate8code" value="<%=catecode(7)%>">
<input type="hidden" name="cate9code" value="<%=catecode(8)%>">
<input type="hidden" name="cate10code" value="<%=catecode(9)%>">
<input type="hidden" name="cate11code" value="<%=catecode(10)%>">
<input type="hidden" name="cate12code" value="<%=catecode(11)%>">
<input type="hidden" name="cate13code" value="<%=catecode(12)%>">
<input type="hidden" name="cate14code" value="<%=catecode(13)%>">
<input type="hidden" name="cate15code" value="<%=catecode(14)%>">
<input type="hidden" name="cate16code" value="<%=catecode(15)%>">
<input type="hidden" name="cate17code" value="<%=catecode(16)%>">
<input type="hidden" name="cate18code" value="<%=catecode(17)%>">
<input type="hidden" name="cate19code" value="<%=catecode(18)%>">
<input type="hidden" name="cate20code" value="<%=catecode(19)%>">
<input type="hidden" name="cate21code" value="<%=catecode(20)%>">
<input type="hidden" name="cate22code" value="<%=catecode(21)%>">
<input type="hidden" name="cate23code" value="<%=catecode(22)%>">
<input type="hidden" name="cate24code" value="<%=catecode(23)%>">
<input type="hidden" name="cate25code" value="<%=catecode(24)%>">
<input type="hidden" name="cate26code" value="<%=catecode(25)%>">
<input type="hidden" name="cate27code" value="<%=catecode(26)%>">
<input type="hidden" name="cate28code" value="<%=catecode(27)%>">
<input type="hidden" name="cate29code" value="<%=catecode(28)%>">
<input type="hidden" name="cate30code" value="<%=catecode(29)%>">
<input type="hidden" name="cate31code" value="<%=catecode(30)%>">
<input type="hidden" name="cate32code" value="<%=catecode(31)%>">
<input type="hidden" name="cate33code" value="<%=catecode(32)%>">
<input type="hidden" name="cate34code" value="<%=catecode(33)%>">
<input type="hidden" name="cate35code" value="<%=catecode(34)%>">
<input type="hidden" name="cate36code" value="<%=catecode(35)%>">
<input type="hidden" name="cate37code" value="<%=catecode(36)%>">
<input type="hidden" name="cate38code" value="<%=catecode(37)%>">
<input type="hidden" name="cate39code" value="<%=catecode(38)%>">
<input type="hidden" name="cate40code" value="<%=catecode(39)%>">
<input type="hidden" name="cate41code" value="<%=catecode(40)%>">
<input type="hidden" name="cate42code" value="<%=catecode(41)%>">
<input type="hidden" name="cate43code" value="<%=catecode(42)%>">
<input type="hidden" name="cate44code" value="<%=catecode(43)%>">
<input type="button" value=" 저장하기 " style="border:1px solid black;" onClick="jsSaveCateMenu()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<% If vItemID <> "" Then %><input type="button" value=" 미리보기 " style="border:1px solid black;" onClick="jsRealServerReg()"><% End If %>
</form>
<br>
</body>
</html>
