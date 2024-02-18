<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%dim dispcate


Function fnDispCateSelectBox(depth, catecode, selname, selectedcode, onchange)
	Dim i, cDCS, vBody, vTempDepth

	SET cDCS = New cDispCate
	cDCS.FCurrPage = 1
	cDCS.FPageSize = 2000
	cDCS.FRectDepth = depth
	cDCS.FRectCateCode = catecode
	cDCS.GetDispCateList()

	For i=0 To cDCS.FResultCount-1

		If i = 0 Then
			vBody = vBody & "<select name="""&selname&""" class=""select"" "&onchange&">" & vbCrLf
			vBody = vBody & "	<option value=''>-선택-</option>" & vbCrLf
		End If
   if cDCS.FItemList(i).FCateCode <> "123" then
		vBody = vBody & "	<option value="""&cDCS.FItemList(i).FCateCode&""""
		If CStr(cDCS.FItemList(i).FCateCode) = CStr(selectedcode) Then
			vBody = vBody & " selected"
		End If
		vBody = vBody & ">"&cDCS.FItemList(i).FCateName&"</option>" & vbCrLf
	 end if
	Next
	vBody = vBody & "</select>" & vbCrLf

	SET cDCS = Nothing
	fnDispCateSelectBox = vBody
End Function

%>
		 <!-- daumeditor head --> 
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=10" /> 
<link rel="stylesheet" href="/lib/util/daumeditor/css/editor.css" type="text/css" charset="euc-kr"/>    
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script src="/lib/util/daumeditor/js/editor_loader.js" type="text/javascript" charset="euc-kr"></script> 
<script src="/lib/util/daumeditor/js/editor_creator.js" type="text/javascript" charset="euc-kr"></script> 
<!-- daumeditor  --> 
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script> 
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
    var config = {
        initializedId: "",
        wrapper: "tx_trex_container",
        form: 'boardform',
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
<!-- //daumeditor head -->  
<script type="text/javascript" > 
	var blockChar=["&lt;script","<scrip","<form","&lt;form","</form","&lt;/form"];  
 function chkContent(p) {
 for (var i=0; i<blockChar.length; i++) {
  if (p.indexOf(blockChar[i])>=0) {
   return blockChar[i];
  }
 }
 return null;
} 
 
 
	function checkform()
	{
		if (document.boardform.title.value == "") {
			alert("제목을 입력해 주십시요...");
			document.boardform.title.focus();
			return;
		}

		if (document.boardform.email.value == "") {
			alert("메일주소를 입력해 주세요...");
			document.boardform.email.focus();
			return;
		}

	//daum editor start---------
		var content = Editor.getContent(); 
		if(content==""||content=="<p>&nbsp;</p>"){
			alert("내용을 입력해주세요");
			return;
		}
		var str = chkContent(content); 
		  if (str) {
		   alert("script태그및 form 태그는 사용할 수 없는 문자열 입니다.\nHTML 버튼을 클릭하셔서 해당태그를 제거해주세요");
		   return ;  
		  } 
 
     document.getElementById("contents").value = content; 
 //daum editor end -----------
  
    if(document.boardform.fixnotics.checked ){
    	if(!document.boardform.sSD.value || !document.boardform.sSH.value) {
    		alert("상단고정 시작일을 입력해주세요");
    		return;
    	}
    	if(!document.boardform.sED.value || !document.boardform.sEH.value) {
    		alert("상단고정 종료일을 입력해주세요");
    		return;
    	}
    }	
    
     if(document.boardform.isPop.checked ){
    	if(!document.boardform.sPSD.value || !document.boardform.sPSH.value) {
    		alert("팝업공지 시작일을 입력해주세요");
    		return;
    	}
    	if(!document.boardform.sPED.value || !document.boardform.sPEH.value) {
    		alert("팝업공지 종료일을 입력해주세요");
    		return;
    	}
    }	
    	
		//폼실행
		document.boardform.submit();
	}
 
 	//파일첨부
function jsAttachFile(sP){
	var winAF = window.open('/admin/board/partnerRegFile.asp?sp='+sP,'popAF','width=400, height=300');
	winAF.focus();
}

//파일삭제
function jsFileDel(sName){
	$("#dF"+sName).remove(); 
}

//파일 다운로드
    function jsDownload(sDownURL, sRFN, sFN){
    var winFD = window.open(sDownURL+"/linkweb/board/procDownload.asp?sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
 }
</script>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
		<form method="POST" name="boardform" action="designer_notice_act.asp">
		<input type="hidden" name="writemode" value="write">
       <table width="850"  cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		  <tr>
            <td width="100" bgcolor="#eeeeee"  align="center">
                글쓴이  
            </td>
            <td width="407"  bgcolor="#FFFFFF">
			  <input type="text" name="name" maxlength="32" value='<%= session("ssBctCname") %>'>
			  (<%= session("ssBctId") %>)
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              상단고정 
            </td>
            <td    bgcolor="#FFFFFF">
			  <input type="checkbox" name="fixnotics" value="Y"> 상단 고정
			  <img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
			  <input type="text" id="sSD" name="sSD" size="10" />
			  <input type="text" id="sSH" name="sSH" size="2" value="00"/>: <input type="text" id="sSM" name="sSM" size="5" class="text_ro" readonly value="00:00" />
			  ~
			  <img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkEnd_trigger" onclick="return false;" />
			  <input type="text" id="sED" name="sED" size="10" />
			  <input type="text" id="sEH" name="sEH" size="2" value="23"/>:<input type="text" id="sEM" name="sEM" size="5" readonly class="text_ro" value="59:59"/>
			  <script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "sSD", trigger    : "ChkStart_trigger",
							onSelect: function() { 
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "sED", trigger    : "ChkEnd_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
             팝업공지 
            </td>
            <td    bgcolor="#FFFFFF">
			  <input type="checkbox" name="isPop" value="Y"> 사용
			   <img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger1" onclick="return false;" />
			  <input type="text" id="sPSD" name="sPSD" size="10" />
			  <input type="text" id="sPSH" name="sPSH" size="2" value="00"/>: <input type="text" id="sPSM" name="sPSM" size="5" class="text_ro" readonly value="00:00" />
			  ~
			  <img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkEnd_trigger1" onclick="return false;" />
			  <input type="text" id="sPED" name="sPED" size="10" />
			  <input type="text" id="sPEH" name="sPEH" size="2" value="23"/>:<input type="text" id="sPEM" name="sPEM" size="5" readonly class="text_ro" value="59:59"/>
			  <script type="text/javascript">
						var CAL_Start1 = new Calendar({
							inputField : "sPSD", trigger    : "ChkStart_trigger1",
							onSelect: function() { 
								var date = Calendar.intToDate(this.selection.get());
								CAL_End1.args.min = date;
								CAL_End1.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End1 = new Calendar({
							inputField : "sPED", trigger    : "ChkEnd_trigger1",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start1.args.max = date;
								CAL_Start1.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              메일발송유무 
            </td>
            <td width="407"   bgcolor="#FFFFFF">
            (이메일 발송은 공지사항 등록 후 공지사항 목록에서 가능합니다.)
<!--
			  <input type="checkbox" name="mailcheck" value="Y">메일발송하기
-->
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              메일 
            </td>
            <td width="407" height="6"  bgcolor="#FFFFFF">
              <input type="text" name="email" size="24" maxlength="128" value="<%= session("ssBctEmail") %>">
            </td>
          </tr>
           <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              업체 전시카테고리 
            </td>
            <td    bgcolor="#FFFFFF">
             <%=fnDispCateSelectBox(1,"","disp",dispCate,"") %>
             <font color="blue">선택시 해당 전시카테고리에 속한 상품이 있는  업체에게만 공지 됩니다.</font>
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              제목 
            </td>
            <td width="407"    bgcolor="#FFFFFF">
              <input type="text" name="title" size="54" maxlength="128">
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              공지사항 내용 
            </td>
            <td  bgcolor="#FFFFFF"> 
						<textarea name="contents" id="contents" style="width: 100%; height: 490px;" style="display:none;"></textarea>    
							 <script type="text/javascript">
							    EditorCreator.convert(document.getElementById("contents"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
		                            EditorJSLoader.ready(function (Editor) {
		                                new Editor(config);
		                                Editor.modify({ 
		                                    content: document.getElementById("contents") 
		                                });
		                            });
		                        });  
							    </script>
		        	<!-- daumeditor   -->
            </td>
          </tr>
          <tr>
					<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">파일첨부</td>
					<td bgcolor="#FFFFFF"><input type="button" value="파일첨부" class="button" onClick="jsAttachFile('');">
					<div id="dFile"></div> 
					</td>
				</tr>
        </table>
		<table border="0" align="center" cellpadding="0" cellspacing="5" width="800">
		<tr>
			<td>
				<a href="javascript:history.back()"><img src="/images/icon_cancel.gif" border="0" align="absmiddle"></a>
				&nbsp;
				<a href="javascript:checkform()"><img src="/images/icon_save.gif" border="0" align="absmiddle"></a> 
				
			</td>
		</tr>
		</table>
       </form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>
