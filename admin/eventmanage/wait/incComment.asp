<!-- #include virtual="/lib/classes/event/commentCls.asp"-->
<% 
'###########################################################
' Description :  코멘트
' History : 2016.08.18 생성
'################################################################## 
%>
<% 
dim arrcomm,clscomm,intC,sRectAuthId,sRegType
sRectAuthId =session("ssBctid")
sRegType ="A"
set clscomm = new CComment
	clscomm.FEvtCode = evtCode
	arrcomm = clscomm.fnGetCommentList
set clscomm = nothing

%>
<script type="text/javascript" src="/js/ajax.js"></script>
<script type="text/javascript">
  	//등록
			function jsSetCmt(){
				document.frmCI.target="ifmCmt";
				document.frmCI.action ="/admin/eventmanage/wait/procComment.asp";
				document.frmCI.submit();
			}
	 
	  //리스트 리로딩		
		// ajax =========================================================================================================
    initializeReturnFunction("processAjax()");
    initializeErrorFunction("onErrorAjax()");
     
    function processAjax(){
        var reTxt = xmlHttp.responseText;  
       	document.all.Cmtlist.innerHTML = reTxt; 
    }
    
    function onErrorAjax() {
            alert("ERROR : " + xmlHttp.status);
    }
    
    //선택한 카테고리에 대한 하위 카테고리 리스트 가져오기 Ajax
    function jsGetCmt(){  
    document.frmCI.tCmt.value=""; 
		var evtCode = document.frmCI.eC.value;	    
 
      initializeURL("/admin/eventmanage/wait/ajaxComment.asp?eC="+evtCode+"&sRAId=<%=sRectAuthId%>");
    	startRequest(); 
    }
    
    //삭제
	  function jsDelCmt(ivalue){
	  	if(confirm("삭제하시겠습니까?")){
	  	document.frmCD.iCidx.value =ivalue;
	  	document.frmCD.target="ifmCmt";
			document.frmCD.action ="/admin/eventmanage/wait/procComment.asp";
			document.frmCD.submit();
	 	 }
	  }
</script> 
<form name="frmCD" method="post" action="/admin/eventmanage/wait/procComment.asp">
 <input type="hidden" name="hidM" value="CD">
 <input type="hidden" name="iCidx" value="">
 <input type="hidden" name="eC" value="<%=evtCode%>">  
 </form>
<form name="frmCI" method="post" action="/admin/eventmanage/wait/procComment.asp">
	<input type="hidden" name="hidM" value="CI">
	 <input type="hidden" name="eC" value="<%=evtCode%>">  
	<input type="hidden" name="hidRT" value="<%=sRegType%>">
	<h3>COMMENT</h3>
	<p class="cmtInput">
		<span><textarea id="tCmt" name="tCmt"   rows="4" style="width:89%" class="formTxtA"></textarea></span>
		<span><input type="button" value="등록" class="btn fs12" id="btnSubmit"  onClick="jsSetCmt()"> </span>
	</p>
	<p style="width:100%; margin-top:10px">
		<div id="Cmtlist" style="padding-left:20px;padding-right:20px;"> 
					<%IF isArray(arrComm) THEN  
						For intC = 0 To UBound(arrComm,2)
						%> 
							<span style="font-size:11px;color:#696969"><%=arrComm(4,intC)%>(<%=arrComm(2,intC)%>)&nbsp;<%=formatdate(arrComm(3,intC),"0000.00.00")%></span>&nbsp;<%IF  sRectAuthId = arrComm(2,intC) THEN%><a href="javascript:jsDelCmt(<%=arrComm(0,intC)%>);"><img src="http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif" border="0"></a><%END IF%>
								 <br>
						  <div style="padding:5px;border-bottom:1px solid #BABABA;width=100%"><%=arrComm(1,intC)%></div><Br>
					<%	Next
					END IF%> 
		</div>
						<iframe name="ifmCmt" id="ifmCmt" src="about:blank" frameborder="0" height="0" width="0"></iframe>
	</p>
	</form> 
  
 