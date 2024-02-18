<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/partnerReferCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
 dim clsref, arrref, intLoop
 dim iCurrpage, iPageSize, iPercnt,iTotCnt,iTotalPage
 dim sTitle, tContents, sType, sregid, sregname, dregdate
 dim refidx
 Dim sMode
 dim arrFile ,intF
 dim strParm
 dim stType, selSearch,strSearch
 
 '--리스트 검색 파라미터================================================================
 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호 
 stType 		= requestCheckVar(Request("strefT"),4)
 selSearch = requestCheckVar(Request("selSearch"),10)
 strSearch = requestCheckVar(Request("strSearch"),200)
  
  strParm = "iC="&iCurrpage&"&strefT="&stType&"&selSearch="&selSearch&"&strSearch="&strSearch
 '--================================================================
  
 refidx =  requestCheckVar(Request("fidx"),10)
 sMode ="I"
 
 if refidx <> "" THEN	 
 		sMode ="U"
		set clsref = new CRefer
	 	clsref.FrefIdx = refidx
	 	clsref.FnGetReferConts
	 	sType 		= clsref.FrefType
	 	sTitle 		= clsref.FTitle
	 	tContents = clsref.FContents
	 	sregid 		= clsref.Fregid
	 	sregname 	= clsref.Fregname
	 	dregdate	= clsref.Fregdate
	 	
	 	arrFile   = clsref.fnGetAttachFile
		set clsref = nothing
	
	END IF	
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
 <!-- daumeditor head -------------------------> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=10" /> 
<link rel="stylesheet" href="/lib/util/daumeditor/css/editor.css" type="text/css" charset="utf-8"/>    
<script src="/lib/util/daumeditor/js/editor_loader.js" type="text/javascript" charset="utf-8"></script> 
<script src="/lib/util/daumeditor/js/editor_creator.js" type="text/javascript" charset="utf-8"></script>  
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
<!-- //daumeditor head ------------------------->
<script type="text/javascript"> 
	
	var blockChar=["&lt;script","<scrip","<form","&lt;form","</form","&lt;/form"];  
 function chkContent(p) {
 for (var i=0; i<blockChar.length; i++) {
  if (p.indexOf(blockChar[i])>=0) {
   return blockChar[i];
  }
 }
 return null;
} 
 
 
	function jsSubmit(){  
		var frm = document.frm; 
		if (!$("#selRefT").val()){
		alert("구분을 선택해주세요");
		 frm.selrefT.focus();
		 return
		}
	
		if(!frm.sT.value){
			alert("제목을 입력해주세요");
			frm.sT.focus();
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
 
     document.getElementById("tC").value = content; 
 		//daum editor end -----------
  

		//실행
		frm.submit();
	}
	
	function jsCancel(){
		location.href="/admin/board/partnerReferList.asp?menupos=<%=menupos%>&<%=strParm%>"
	}
	
	function jsDelete(){
		document.frm.hidM.value="D";
		if (confirm("삭제하시겠습니까?")){
			document.frm.submit();
		}
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
<form name="frm" method="post" action="partnerReferProc.asp?<%=strParm%>"> 
	<input type="hidden" name="hidM" value="<%=sMode%>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="850" border="0" class="a" cellpadding="3" cellspacing="0" > 
<tr>
	<td >
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<%if refidx <> "" THEN	 %>
			<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">idx</td>
		   		<td bgcolor="#FFFFFF"><%=refidx%><input type="hidden" name="fidx" value="<%=refidx%>"></td>
		   	</tr>
			 <tr>
			<%end if%> 	
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">구분</td>
		   		<td bgcolor="#FFFFFF"><%fnOptReferType sType%></td>
		   	</tr>
		   	 <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">제목</td>
		   		<td bgcolor="#FFFFFF"><input type="text" name="sT" size="60" maxlength="60" value="<%=sTitle%>"></td>
		   	</tr>
		   	 <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">내용</td>
		   		<td bgcolor="#FFFFFF">
		   			<textarea name="tC" id="tC" style="width: 100%; height: 490px;"><%=tContents%></textarea>    
					 	<script type="text/javascript"> 
					    EditorCreator.convert(document.getElementById("tC"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
                            EditorJSLoader.ready(function (Editor) {
                                new Editor(config);
                                Editor.modify({
                                    content: document.getElementById("tC")
                                });
                            });
                        });  
					    </script>
              <!-- daumeditor   -->
		   		</td>
		   	</tr>
		   	<tr>
					<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">첨부파일</td>
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
	</td>
</tr>	 
<tr>
	<td width="100%" align="center" style="padding-top:10px;">
		<input type="button" class="button" value="목록으로" style="width:80px;" onClick="jsCancel();"> &nbsp;
		<%IF sMode="U" THEN%>
		<input type="button" class="button" value="삭제" style="width:80px;" onClick="jsDelete();">
		<%END IF%> 
		<input type="button" class="button" value="등록" style="width:80px;color:red" onClick="jsSubmit();">
		
	</td>
</tr>
</table>
</form>  	
 
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>
