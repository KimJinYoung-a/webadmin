<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyheadUTF8.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/partnerFaqCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
 dim clsFaq, arrFaq, intLoop
 dim iCurrpage, iPageSize, iPercnt,iTotCnt,iTotalPage
 dim sTitle, tContents, sType, sregid, sregname, dregdate
 dim faqidx
 Dim sMode
 dim strParm
 dim stType, selSearch,strSearch
 
 '--리스트 검색 파라미터================================================================
 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호 
 stType 		= requestCheckVar(Request("stfaqT"),4)
 selSearch = requestCheckVar(Request("selSearch"),10)
 strSearch = requestCheckVar(Request("strSearch"),200)
  
  strParm = "iC="&iCurrpage&"&stfaqT="&stType&"&selSearch="&selSearch&"&strSearch="&strSearch
 '--================================================================

 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
 faqidx =  requestCheckVar(Request("fidx"),10)
 sMode ="I"
 
 if faqidx <> "" THEN	 
 		sMode ="U"
		set clsFaq = new CFaq
	 	clsFaq.FFaqIdx = faqidx
	 	clsFaq.FnGetFaqConts
	 	sType 		= clsFaq.FFaqType
	 	sTitle 		= clsFaq.FTitle
	 	tContents = clsFaq.FContents
	 	sregid 		= clsFaq.Fregid
	 	sregname 	= clsFaq.Fregname
	 	dregdate	= clsFaq.Fregdate
		set clsFaq = nothing
	
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
		if (!$("#selFaqT").val()){
		alert("FAQ 구분을 선택해주세요");
		 frm.selFaqT.focus();
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
		location.href="/admin/board/partnerfaqList.asp?menupos=<%=menupos%>&<%=strParm%>";
	}
	
	function jsDelete(){
		document.frm.hidM.value="D";
		if (confirm("삭제하시겠습니까?")){
			document.frm.submit();
		}
	} 
</script>
<form name="frm" method="post" action="partnerFaqProc.asp?<%=strParm%>"> 
	<input type="hidden" name="hidM" value="<%=sMode%>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="850" border="0" class="a" cellpadding="3" cellspacing="0" > 
<tr>
	<td >
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<%if faqidx <> "" THEN	 %>
			<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">FAQ idx</td>
		   		<td bgcolor="#FFFFFF"><%=faqidx%><input type="hidden" name="fidx" value="<%=faqidx%>"></td>
		   	</tr>
			 <tr>
			<%end if%> 	
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">FAQ 구분</td>
		   		<td bgcolor="#FFFFFF"><%fnOptFaqType sType%></td>
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
