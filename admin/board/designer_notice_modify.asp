<%@ language="VBScript" %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	Dim ix,i, page, pgsize
	Dim TotalPage, TotalCount
	Dim prepage, nextpage
	Dim mode
	Dim nIndent
	Dim name,email,title,contents
    Dim nboard,arrFile,intF
	Dim fixnotics, fixSdate, fixEdate, fixSH, fixSM, fixEH, fixEM
	dim dispCate
 dim fileName
 dim isPopup, popSdate, popEdate, popSH, popSM, popEH, popEM
 
	if Request("pgsize")="" then
		pgsize = 10
	else
		pgsize = Request("pgsize")
	end if

	if Request("page") = "" then
		page = 1
	else
		page = cInt(Request("page"))
	end if

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

set nboard = new CBoard
nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
nboard.design_notice_read request("idx")
nboard.FRectidx		 = request("idx")
arrFile   = nboard.fnGetAttachFile 

if Request.Form("modifymode") = "modify" then

name = Request("name")
email = Request("email")
title = Request("title")
contents = Request("contents")
title = replace(title,chr(34) , "")
title = replace(title,"'" , "&#8217;")
contents = replace(contents,"'","&#8217;")
fixnotics = Request("fixnotics")
dispCate	= requestCheckVar(Request("disp"),10) 
 
 fixSdate =  requestCheckVar(Request("sSD"),10) 
 fixEdate =  requestCheckVar(Request("sED"),10) 
 fixSH =  requestCheckVar(Request("sSH"),2) 
 fixSM =  requestCheckVar(Request("sSM"),5) 
 fixEH =  requestCheckVar(Request("sEH"),2) 
 fixEM =  requestCheckVar(Request("sEM"),5) 
 
 
 fixSdate = fixSdate&" "&format00(2,fixSH)&":"&fixSH 
 fixEdate = fixEdate&" "&format00(2,fixEH)&":"&fixEM 

 isPopup =  requestCheckVar(Request("isPop"),1) 
 

 popSdate =  requestCheckVar(Request("sPSD"),10) 
 popEdate =  requestCheckVar(Request("sPED"),10) 
 popSH =  requestCheckVar(Request("sPSH"),2) 
 popSM =  requestCheckVar(Request("sPSM"),5) 
 popEH =  requestCheckVar(Request("sPEH"),2) 
 popEM =  requestCheckVar(Request("sPEM"),5) 
 
 popSdate = popSdate&" "&format00(2,popSH)&":"&popSM 
 popEdate = popEdate&" "&format00(2,popEH)&":"&popEM 
 
   
 fileName 	= ReplaceRequestSpecialChar(Request("sFileP")) 
nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
nboard.FRectID = session("ssBctId")
nboard.FPageSize = pgsize
nboard.FRectName = name
nboard.FRectEmail = email
nboard.FRectTitle = title
nboard.FRectContents = contents
nboard.FCurrPage = page
nboard.FFixNotics = fixnotics
nboard.FRectDispCate = dispCate
nboard.FRectfileName	= fileName
nboard.FfixSdate = fixSdate
nboard.FfixEdate = fixEdate
nboard.FisPopup = ispopup
nboard.FpopSdate = popSdate
nboard.FpopEdate = popEdate
nboard.design_notice_modify request("idx")

response.redirect "designer_notice_read.asp?idx=" + request("idx") + "&menupos=79"

end if

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
			alert("제목을 입력해 주세요...");
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

<form method="POST" name="boardform" action="designer_notice_modify.asp?idx=<% =Request("idx")%>&pgsize=<% =Request("pgsize")%>&page=<% =Request("page")%>&menupos=79" >
<input type="hidden" name="modifymode" value="modify"> 
         <table width="850"  cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
          <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              글쓴이   
            </td>
            <td width="407" bgcolor="#ffffff">
              <input type=text name="name" value="<%= nboard.FRectName %>">
            </td>
          </tr> 
           <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              상단고정 
            </td>
            <td   bgcolor="#FFFFFF">
			  <input type="checkbox" name="fixnotics" value="Y" <% if nboard.FFixnotics="Y" then response.write "checked" %>> 상단 고정
			  <%dim sSD, sED, sSH, sEH
			  sSH = "00"
			  sEH = "23"
		 
			  if nboard.FFixSdate <> "" or not isNull(nboard.FfixSdate) then
			  	sSD = left(nboard.FFixSdate,10)
			  	sSH= mid(nboard.FFixSdate,12,2)
			 end if
			
			if nboard.FFixEdate <> "" or not isNull(nboard.FfixEdate) then
			  	sED = left(nboard.FFixEdate,10)
			  	sEH= mid(nboard.FFixEdate,12,2)
			 end if
			  %>
			  <img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
			  <input type="text" id="sSD" name="sSD" size="10" value="<%=sSD%>"/>
			  <input type="text" id="sSH" name="sSH" size="2" value="<%=sSH%>"/>: <input type="text" id="sSM" name="sSM" size="5" class="text_ro" readonly value="00:00" />
			  ~
			  <img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkEnd_trigger" onclick="return false;" />
			  <input type="text" id="sED" name="sED" size="10"  value="<%=sED%>" />
			  <input type="text" id="sEH" name="sEH" size="2" value="<%=sEH%>"/>:<input type="text" id="sEM" name="sEM" size="5" readonly class="text_ro" value="59:59"/>
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
			  <input type="checkbox" name="isPop" value="Y" <% if nboard.Fispopup="Y" then response.write "checked" %>> 사용
			  <%dim sPSD, sPED, sPSH, sPEH
			  sPSH = "00"
			  sPEH = "23"
		 
			  if nboard.FpopSdate <> "" or not isNull(nboard.FpopSdate) then
			  	sPSD = left(nboard.FpopSdate,10)
			  	sPSH= mid(nboard.FpopSdate,12,2)
			 end if
			
			if nboard.FpopEdate <> "" or not isNull(nboard.FpopEdate) then
			  	sPED = left(nboard.FpopEdate,10)
			  	sPEH= mid(nboard.FpopEdate,12,2)
			 end if
			  %>
			   <img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger1" onclick="return false;" />
			  <input type="text" id="sPSD" name="sPSD" size="10"  value="<%=sPSD%>"/>
			  <input type="text" id="sPSH" name="sPSH" size="2" value="<%=sPSH%>"/>: <input type="text" id="sPSM" name="sPSM" size="5" class="text_ro" readonly value="00:00" />
			  ~
			  <img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkEnd_trigger1" onclick="return false;" />
			  <input type="text" id="sPED" name="sPED" size="10" value="<%=sPED%>"/>
			  <input type="text" id="sPEH" name="sPEH" size="2" value="<%=sPEH%>"/>:<input type="text" id="sPEM" name="sPEM" size="5" readonly class="text_ro" value="59:59"/>
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
              메일   
            </td>
            <td width="407"  bgcolor="#ffffff">
              <input type="text" name="email" size="54" maxlength="128" value="<% =nboard.FRectEmail  %>">
            </td>
          </tr>
           <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              업체 전시카테고리 
            </td>
            <td width="407" height="6"  bgcolor="#FFFFFF">
             <%=fnDispCateSelectBox(1,"","disp",nboard.Fdispcate1,"") %>
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
              제목   
            </td>
            <td width="407"  bgcolor="#ffffff">
              <input type="text" name="title" size="54" maxlength="128" value="<% =nboard.FRectTitle  %>">
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" align="center">
               공지사항 내용   
            </td>
            <td  bgcolor="#ffffff"> 
						<textarea name="contents" id="contents" style="width: 100%; height: 490px;" style="display:none;"><%=nboard.FRectContents%></textarea>    
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
					<div id="dFile">
						<% Dim arrFName,arrF, sFName, intF2,intF3, iCount 
						IF isArray(arrFile) THEN
						For intF=0 To UBound(arrFile,2) 
					
								arrF = split(arrFile(2,intF),"/") 
							 	arrFName = arrF(ubound(arrF))
								sFName = split(arrFName,".")(0)  
						%>
						<div id="dF<%=sFName%>"><a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a>&nbsp;<input type="button" value="x" class="button" onclick="jsFileDel('<%=sFName%>')"> 
							<input type="hidden" name="sFileP"   value="<%= arrFile(2,intF)%>"></div>
					<%Next
						END IF
						%> 
						</div> 
					</td>
				</tr>
        </table>
        <p>
		<table border="0" width="850" align="center" cellpadding="0" cellspacing="5">
		<tr>
			<td>
				<a href="javascript:history.back()"><img src="/images/icon_cancel.gif" border="0" align="absmiddle"></a>&nbsp;
				<a href="javascript:checkform()"><img src="/images/icon_modify.gif" border="0" align="absmiddle"></a> 
			</td>
		</tr>
		</table>
       </form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>
