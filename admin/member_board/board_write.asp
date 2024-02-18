<%@ language="VBScript" %>
<% option explicit %>
<%
'####################################################
' Description :  공지사항 등록
' History : 2011.02.23 김진영 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim g_MenuPos, writer
	dim yyyy1, mm1, dd1, yyyy2, mm2, dd2
	dim startDate, endDate, brd_type

	IF application("Svr_Info")="Dev" THEN '### 메뉴번호 지정.
		g_MenuPos   = "1288"
	Else
		g_MenuPos   = "1304"
	End If

	Dim cBoard
	Dim sBrd_Id, sBrd_Name, sBrd_Regdate
	Dim part_sn, job_sn, posit_sn
	Dim brd_content, arrFileList
	sBrd_Id 		= session("ssBctId")
	sBrd_Name		= session("ssBctCname")
	sBrd_Regdate	= Left(now(),10)
	brd_type		= requestCheckvar(request("brd_type"),3)

	yyyy1   = Year(Now())
	mm1     = Month(Now())
	dd1     = Day(Now())
	yyyy2   = yyyy1
	mm2     = mm1
	dd2     = dd1

	startDate = Left(Now(), 10)
	endDate = startDate

	set cBoard = new Board
		cBoard.fnBoardcontent
%>
<!-- daumeditor head ------------------------->
 <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
 <meta http-equiv="X-UA-Compatible" content="IE=10" />
 <link rel="stylesheet" href="/lib/util/daumeditor/css/editor.css" type="text/css" charset="euc-kr"/>
 <script src="/lib/util/daumeditor/js/editor_loader.js" type="text/javascript" charset="euc-kr"></script>
 <script src="/lib/util/daumeditor/js/editor_creator.js" type="text/javascript" charset="euc-kr"></script>
 <script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript">
    var config = {
        initializedId: "",
        wrapper: "tx_trex_container",
        form: 'frm',
        txIconPath: "/lib/util/daumeditor/images/icon/editor/",
        txDecoPath: "/lib/util/daumeditor/images/deco/contents/",

        events: {
            preventUnload: false
        },
        sidebar: {
            attachbox: {
                show: true
            },
            attacher: {
                 image: {
                    popPageUrl: "/lib/util/daumeditor/pages/trex/image.asp"
                }
            }

        }
    };
 </script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!-- //daumeditor head ------------------------->
<script language="javascript">

var blockChar=["&lt;script","<scrip","<form","&lt;form"];
 function chkContent(p) {
 for (var i=0; i<blockChar.length; i++) {
  if (p.indexOf(blockChar[i])>=0) {
   return blockChar[i];
  }
 }
 return null;
}

function notifyAbled(v){
	if(v == "17") {
		$("input[name='isNotify']").attr('disabled', false);
	}else{
		$("input[name='isNotify']").attr('disabled', true);
		$("input[name='isNotify']").removeAttr('checked');
	}
}

function fTeam(str){
	if(str == "all"){
		document.getElementById('brd_team').style.display = 'none';
		document.getElementById('brd_isNotify').style.display = '';
		for(var j=0; j<frm.part_sn.length; j++) {
			frm.part_sn[j].checked = false;
		}
	}else{
		document.getElementById('brd_team').style.display = '';
		document.getElementById('brd_isNotify').style.display = 'none';
		$("input[name='isNotify']").removeAttr('checked');
		//var sTeam = window.open("","","");
	}
}
function form_check(){
	var frm = document.frm;
//구분 선택 //
	if(frm.brd_type.value == "")
	{
		alert("공지 구분을 선택하세요");
		return false;
	}

//열람선택 여부//?/
	var chk = 0;
	for(var i=0; i<frm.open_team.length; i++) {
		if(frm.open_team[i].checked) chk++;
	}
	if(chk == "0"){
	  alert("열람선택을 선택하세요");
		return false;
	}
//팀공지 선택시 팀 선택 여부//
	if(frm.open_team[1].checked){
		var chk2 = 0;
	   if(document.all.did != undefined) {
	       chk2 =1;
	       if(document.all.did.length != undefined) {
	            for(i=0;i<document.all.did.length;i++){
	                if(i==0){
	                     frm.arrdid.value = document.all.did[i].value;
	                }else{
	                     frm.arrdid.value = frm.arrdid.value+","+document.all.did[i].value;
	                }
	            }
	        }else{
	            frm.arrdid.value = document.all.did.value;
	        }
	    }

		if(chk2 == "0"){
				alert("팀을 선택하세요");
			return false;
		}
	}

//팀공지 선택시 열람권한 선택 여부//
	if(frm.open_team[1].checked){
		var chk3 = 0;
		for(var k=0; k<frm.job_sn.length; k++) {
			if(frm.job_sn[k].selected){
				chk3++;
			}
		}
		if(chk3 == "0"){
		alert("직책을 선택하세요");
			return false;
		}
	}
//팀공지 선택시 직급 선택 여부//
	if(frm.open_team[1].checked){
		var chk4 = 0;
		for(var w=0; w<frm.posit_sn.length; w++) {
			if(frm.posit_sn[w].selected){
				chk4++;
			}
		}
		if(chk4 == "0"){
			alert("직급을 선택하세요");
			return false;
		}
	}
//제목 입력 여부//
	if(frm.brd_subject.value == ""){
			alert("제목을 입력하세요");
		frm.brd_subject.focus();
		return false;
	}

//고정선택 여부//
	var chk3 = 0;
	for(var k=0; k<frm.brd_fixed.length; k++) {
		if(frm.brd_fixed[k].checked) chk3++;
	}
	if(chk3 == "0"){
	alert("고정여부를 선택하세요");
		return false;
	}

	  var content = Editor.getContent();
	  	var str = chkContent(content);
		  if (str) {
		   alert("script태그 또는 form태그는 사용할 수 없는 문자열 입니다.");
		   return ;
		  }
      document.getElementById("brd_content").value = content;

  //파일처리
  $("input[name='doc_file']").each( function(index,elem) {
			     var a = $(elem).val();
			     if( document.frm.sFile.value==""){
			     	document.frm.sFile.value = a;
			    }else{
			     document.frm.sFile.value = document.frm.sFile.value + ","+a;
			   }
	 });

 $("input[name='doc_realfile']").each( function(index,elem) {
			     var a = $(elem).val();
			     if( document.frm.sRFile.value==""){
			     	document.frm.sRFile.value = a;
			    }else{
			     document.frm.sRFile.value = document.frm.sRFile.value + ","+a;
			   }
	 });

 	frm.action = "board_proc.asp";
	frm.submit();
}
function fileupload()
{
	window.open('board_popupload.asp','worker','width=420,height=200,scrollbars=yes');
}
function clearRow(tdObj) {
	if(confirm("선택하신 파일을 삭제하시겠습니까?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}


function jsOpenClose(cValue,iValue){
		if(eval("document.all.divB"+cValue+iValue).style.display=="none"){
			eval("document.all.Fimg"+iValue).src = "/images/dtree/openfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tminus.png";
			eval("document.all.divB"+cValue+iValue).style.display="";
		}else{
			eval("document.all.Fimg"+iValue).src = "/images/dtree/closedfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tplus.png";
			eval("document.all.divB"+cValue+iValue).style.display="none";
		}
	}


    function jsSelTeam(){
        var sdid="";
        if(document.all.did!=undefined){
            if(document.all.did.length!=undefined){
                for(i=0;i<document.all.did.length;i++){
                    if(i==0){
                        sdid =  document.all.did[i].value;
                    }else{
                        sdid = sdid +"," + document.all.did[i].value;
                     }
                }
            }else{
                sdid =  document.all.did.value;
            }
        }
        var winTeam = window.open("/admin/member_board/popselectteam.asp?did="+sdid,"popT","width=330,height=700,resizable=yes");
        winTeam.focus();
    }


    function jsTeamDel(i){
        eval("document.all.dvSTeam"+i).outerHTML = "";
    }

function jsSetEndDate(endDateGubun) {
	var frm = document.frm;
	var yyyy1, mm1, dd1, yyyy2, mm2, dd2;
	var startDate, endDate;

	startDate = new Date(frm.startDate.value);
	endDate = new Date(startDate);

	switch (endDateGubun) {
		case "1w":
			endDate.setDate(endDate.getDate() + 7);
			break;
		case "1m":
			endDate.setMonth(endDate.getMonth() + 1);
			break;
		case "3m":
			endDate.setMonth(endDate.getMonth() + 3);
			break;
		case "1y":
			endDate.setFullYear(endDate.getFullYear() + 1);
			break;
		default:
			alert('에러!!');
			return;
	}

	yyyy2 = endDate.getFullYear();
	mm2 = ("0" + (endDate.getMonth() + 1));
	dd2 = ("0" + endDate.getDate());

	frm.endDate.value = yyyy2 + "-" + mm2.substring(mm2.length - 2) + "-" + dd2.substring(dd2.length - 2);
}
</script>

<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td><img src="/images/icon_arrow_link.gif"></td>
	<td style="padding-top:3">&nbsp;<b>게시글 작성</b></td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<form name="frm"  method="post">
<input type = "hidden" name = "mode" value = "add">
<input type = "hidden" name = "menupos" value = "<%=menupos%>">
<input type = "hidden" name = "brd_sn" value = "<%=cBoard.Fbrd_sn + 1%>">
<input type = "hidden" name = "arrdid" value = "">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">번호</td>
			<td width="800" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No.<%= cBoard.Fbrd_sn + 1 %></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">등록자</td>
			<td width="800" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sBrd_Name%>(<%=sBrd_Id%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;등록일: <%=sBrd_Regdate%></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">공지구분</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td>
						<table class="a">
						<tr>
							<td><%=fnBrdType("w", "Y", brd_type, "")%>&nbsp;</td>
							<td><font color="RED"> ※ 인사관련공지 : 신규입사퇴사자공지,인사이동공지,기타인사관련
								<br>&nbsp;&nbsp;&nbsp;&nbsp;회사내규관련공지 : 입사,퇴사,급여 등 회사와 직원간의 규율 관련
								<br>&nbsp;&nbsp;&nbsp;&nbsp;일반 안내 공지 : 상기 항목 외 공지
								</font></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30" bgcolor="#FFFFFF">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">열람선택</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td>
						<label><input type="radio" name="open_team" value="Y" checked onclick="fTeam('all');">전체공지</label>&nbsp;&nbsp;&nbsp;
						<label><input type="radio" name="open_team" value="N" onclick="fTeam('team');">팀공지</label>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30" width="100%"   style="display:none;" id="brd_team">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">팀선택</td>
			<td width="800" bgcolor="#FFFFFF" style="padding: 0 0 0 5">

					<!--% call DrawPartinfoCombo("part_sn", part_sn,"") %-->
				     <div style="padding:5 5 5 5px;"><input type="button" class="button" value="팀선택" onClick="jsSelTeam();"></div>
				     <div id="dvTeam" style="padding:5 5 5 5px;"></div>
			 </td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">직책</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<% call DrawJobCombo("job_sn", job_sn) %>&nbsp;<font color="RED"> ※ 해당 직책 이상 열람 가능합니다. (일반을 선택하시면 모두 열람가능하게 됩니다)</font>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">직급</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<% ''call DrawPositCombo("posit_sn", posit_sn) %>
					<select name="posit_sn" onchange="notifyAbled(this.value);" class="select" >
					<option value='11' >정규직 이상</option>
					<option value='12' >월급계약직 이상</option>
					<option value='13' >시급계약직 이상</option>
					<option value='17' >전체</option>
					</select>
					&nbsp;&nbsp;&nbsp;<font color="RED">※ 해당 직급 이상 열람 가능합니다.</font>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">제 목</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="brd_subject" value="" size="95" maxlength="128">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내 용</td>
			<td bgcolor="#FFFFFF" style="padding: 5 4 2 5">
				<textarea name="brd_content" id="brd_content" style="width: 100%; height: 490px;"></textarea>
                <!-- daumeditor  -->
                <script type="text/javascript">
                    EditorCreator.convert(document.getElementById("brd_content"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
                        EditorJSLoader.ready(function (Editor) {
                            new Editor(config);
                            Editor.modify({
                                content: ''
                            });
                        });
                    });

                </script>
                <!-- daumeditor   -->
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">첨부파일</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td width="100" valign="top" style="padding:5 0 0 0">
						<input type="button" value="파일업로드" class="button" onclick="fileupload();">
					</td>
					<td width="100%" style="padding:3 0 3 10">
						<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
							<input type="hidden" name="sFile" value="">
							<input type="hidden" name="sRFile" value="">
						<%
						IF isArray(arrFileList) THEN
							For i =0 To UBound(arrFileList,2)
						%>
							<tr>
								<td>
									<input type='hidden' name='doc_file' value='<%=arrFileList(0,i)%>'>
									<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
									<a href='<%=arrFileList(0,i)%>' target='_blank'>
									<%=Split(Replace(arrFileList(0,i),"http://",""),"/")(4)%></a>
								</td>
							</tr>
						<%
							Next
							Response.Write "<input type='hidden' name='isfile' value='o'>"
						Else
						%>
							<tr>
								<td>
								</td>
							</tr>
						<% End If %>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">고정 여부</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label id="brd_fixed1"><input type="radio" name="brd_fixed" id="brd_fixed1" value="1">고정함</label>&nbsp;&nbsp;&nbsp;
				<label id="brd_fixed2"><input type="radio" name="brd_fixed" id="brd_fixed2" value="2" checked>고정안함</label>&nbsp;&nbsp;&nbsp;
				<font color = "RED"> ※고정 여부 Y를 선택하시면 게시글의 최상단에 위치하게 됩니다.</font>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">고정 기간</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input id="startDate" name="startDate" value="<%= startDate %>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startDate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
				<input id="endDate" name="endDate" value="<%= endDate %>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="endDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "startDate", trigger    : "startDate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "endDate", trigger    : "endDate_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
				&nbsp;
				<a href="javascript:jsSetEndDate('1w')">[일주일]</a>
				&nbsp;
				<a href="javascript:jsSetEndDate('1m')">[한달]</a>
				&nbsp;
				<a href="javascript:jsSetEndDate('3m')">[3개월]</a>
				&nbsp;
				<a href="javascript:jsSetEndDate('1y')">[일년]</a>
			</td>
		</tr>
		<tr height="30" id="brd_isNotify">
			<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">잔디 웹훅</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label id="isNotify1"><input type="radio" disabled name="isNotify" id="isNotify1" value="Y">Y</label>&nbsp;&nbsp;&nbsp;
				<label id="isNotify2"><input type="radio" disabled name="isNotify" id="isNotify2" value="N">N</label>&nbsp;&nbsp;&nbsp;
				<font color = "RED"> ※Y를 선택하시면 잔디메세지로 내용이 전송됩니다.(전체공지, 직급 전체일 때만 선택 가능) </font>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<table width="900" border="0" cellpadding="0" cellspacing="10" class="a">
<tr>
	<td width="50%" align="left"><a href="board_list.asp?menupos=<%=menupos%>&brd_type=<%=brd_type%>"><img src="/images/icon_list.gif" border="0"></a></td>
	<td width="50%" align="right">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><input type="image" src="/images/icon_confirm.gif" border="0" onclick="form_check();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
