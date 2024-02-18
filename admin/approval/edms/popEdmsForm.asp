<%@ language="VBScript" %>
<% option explicit %> 
<%
'###########################################################
' Description : 문서관리 문서내용
' History : 2011.02.24 정윤정  생성
'           2013.03.05 허진원 - 이노디터로 변경
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<%
Dim clsedms
Dim iedmsidx, tContents
iedmsidx = requestCheckvar(Request("ieidx"),10)
Set clsedms = new Cedms
	clsedms.Fedmsidx 	= iedmsidx
	clsedms.fnGetEdmsData
	tContents = clsedms.FedmsForm
Set clsedms = nothing
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
        form: 'frmReg',
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
 <script language="javascript" type="text/javascript">   
 var blockChar=["&lt;script","<scrip","<form","&lt;form","</form","&lt;/form"];  
 function chkContent(p) {
 for (var i=0; i<blockChar.length; i++) {
  if (p.indexOf(blockChar[i])>=0) {
   return blockChar[i];
  }
 }
 return null;
} 
   
	function chkForm() { 
		//daum editor start---------
 	var content = Editor.getContent(); 
//		if(content==""||content=="<p>&nbsp;</p>"){
//			alert("내용을 입력해주세요");
//			return ;
//		}
		var str = chkContent(content); 
		  if (str) {
		   alert("script태그및 form 태그는 사용할 수 없는 문자열 입니다.\nHTML 버튼을 클릭하셔서 해당태그를 제거해주세요");
		   return ;  
		  } 
 
     document.getElementById("editor").value = content; 
 //daum editor end -----------
 document.frmReg.submit();
	}
 </script>
 
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td><strong>문서폼 등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<form name="frmReg" method="post" action="procedms.asp">
		<input type="hidden" name="hidM" value="F">
		<input type="hidden" name="ieidx" value="<%=iedmsidx%>"> 
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50" align="center">문서폼</td>
			<td bgcolor="#FFFFFF"  align="center">
					<textarea name="editor" id="editor" style="width: 100%; height: 490px;"><%=tContents%></textarea>  
          <!-- daumeditor  --> 
          <script type="text/javascript">  
              EditorCreator.convert(document.getElementById("editor"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
                  EditorJSLoader.ready(function (Editor) {
                      new Editor(config);
                      Editor.modify({
                          content:  '<%=tContents%>'
                      });
                  });
              });
          
          </script> 
          <!-- daumeditor   -->				
			</td>
 		</table>
	</td>
</tr>
 <tr>
 	<td align="center"><input type="button" value="등록" class="button" onClick="chkForm()"></td>
 	</tr>
 </form>
</table>
<!-- 페이지 끝 -->
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>