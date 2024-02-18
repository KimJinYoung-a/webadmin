<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// 즐겨찾기
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script type="text/javascript" src="/js/xl.js"></script>
<script type="text/javascript" src="/js/common.js"></script>
<script type="text/javascript" src="/js/report.js"></script>
<script type="text/javascript" src="/cscenter/js/cscenter.js"></script>
<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script type='text/javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;

	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "즐겨찾기에서 제외하시겠습니까?";
	} else {
		msg = "즐겨찾기에 추가하시겠습니까?";
	}

	ret = confirm(msg);

	if (ret) {
		frm.submit();
	}
}
</script>
</head>
<body>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

%>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function goPage(p){
	frm1.page.value = p;
	frm1.submit();
}

function popDetail(idx){	
	var popModi;
	popModi = window.open('appKeyView.asp?idx='+idx+'','popAppKeyView','width=1000,height=524,scrollbars=yes,resizable=yes');
	popModi.focus();
}

$(function(){
	$(".tbType1 .tbListRow").hover(function() {
		$(this).toggleClass('hover');
	});

    $("#userinputdata").keydown(function(key) {
        if (key.keyCode == 13) {
            doActionUseq();
            return false;
        }
    });
});

function doActionUseq() {
    if ($("#userinputdata").val()==""){
        alert("userid나 useq값을 입력해주세요.");
        return false;
    }

    $("#ttype").val($("#type").val());
    $("#tinputdata").val($("#userinputdata").val());

    var str = $.ajax({
        type: "post",
        url:"useqViewProc.asp",
        data: $("#frmuseq").serialize(),
        dataType: "text",
        async: false
    }).responseText;	

    if(!str){alert("시스템 오류입니다."); return false;}

    var resultData = JSON.parse(str);

    var reStr = resultData.data[0].result.split("|");
    var useq = resultData.data[0].useq;
    var userid = resultData.data[0].userid;		
    var userdiv = resultData.data[0].userdiv;		
    var lastlogin = resultData.data[0].lastlogin;		
    var counter = resultData.data[0].counter;		
    var userlevel = resultData.data[0].userlevel;		                

    if(reStr[0]=="OK"){		
        $("#useq").html(useq);
        $("#userid").html(userid);
        $("#userdiv").html(userdiv);
        $("#lastlogin").html(lastlogin);
        $("#counter").html(counter);
        $("#userlevel").html(userlevel);
    }else{
        var errorMsg = reStr[1].replace(">?n", "\n");
        alert(errorMsg);
    }
}
</script>

<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2><%=imenuposStr%></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="<%=menupos%>">
			</form>
			<a href="javascript:fnMenuFavoriteAct('addonefavorite')">즐겨찾기</a> l 
			<!-- 마스터이상 메뉴권한 설정 //-->
			<a href="Javascript:PopMenuEdit('<%=menupos%>');">권한변경</a> l 
			<!-- Help 설정 //-->
			<a href="Javascript:PopMenuHelp('<%=menupos%>');">HELP</a>
		</div>
	</div>

	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li>
					<label class="formTit" for="type">구분 :</label>
					<select name="type" class="formSlt" id="type">
						<option value="userid">userid</option>
						<option value="useq">useq</option>
					</select>
				</li>			
				<li>
                    <input type="text" name="userinputdata" id="userinputdata">
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<input type="submit" class="schBtn" onclick="doActionUseq();return false;" value="검색" />
	</div>
	<div class="pad20">
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>useq</div></th>
					<th><div>userid</div></th>
					<th><div>userdiv</div></th>
					<th><div>lastlogin</div></th>
					<th><div>counter</div></th>
					<th><div>userlevel</div></th>
				</tr>
				</thead>
				<tbody>
						<tr class="tbListRow">
                            <td><span id="useq"></span></td>
                            <td><span id="userid"></span></td>
                            <td><span id="userdiv"></span></td>
                            <td><span id="lastlogin"></span></td>
                            <td><span id="counter"></span></td>
                            <td><span id="userlevel"></span></td>
						</tr>
				</tbody>
			</table>
			<br />
		</div>
	</div>
</div>
<form name="frmuseq" id="frmuseq" method="post">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="ttype" id="ttype">
<input type="hidden" name="tinputdata" id="tinputdata">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
