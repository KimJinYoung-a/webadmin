<!-- #include virtual="/lib/classes/board/commentCls.asp"-->
<%
dim arrcomm,clscomm,intC,sRectAuthId
sRectAuthId =session("ssBctid")
set clscomm = new CComment
	clscomm.FboardIdx = iboard_idx
	arrcomm = clscomm.fnGetCommentList
set clscomm = nothing

%>
<script type="text/javascript" src="/js/ajax.js"></script>
  <script language="javascript" >
  	//���
			function jsSetCmt(){
				document.frmCI.target="ifmCmt";
				document.frmCI.action ="/admin/board/procComment.asp";
				document.frmCI.submit();
			}
	 
	  //����Ʈ ���ε�		
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
    
    //������ ī�װ��� ���� ���� ī�װ� ����Ʈ �������� Ajax
    function jsGetCmt(){  
    document.frmCI.tCmt.value=""; 
		var iboardidx = document.frmCI.ibidx.value;	   		
      initializeURL("/admin/board/ajaxComment.asp?ibidx="+iboardidx+"&sRAId=<%=sRectAuthId%>"); 
    	startRequest(); 
    }
 
    
    //����
	  function jsDelCmt(ivalue){
	  	if(confirm("�����Ͻðڽ��ϱ�?")){
	  	document.frmCD.iCidx.value =ivalue;
	  	document.frmCD.target="ifmCmt";
			document.frmCD.action ="/admin/board/procComment.asp";
			document.frmCD.submit();
	 	 }
	  }
			</script> 
<form name="frmCD" method="post" action="/admin/board/procComment.asp">
 <input type="hidden" name="hidM" value="CD">
 <input type="hidden" name="iCidx" value="">
 <input type="hidden" name="ibidx" value="<%=iboard_idx%>">  
 </form>
<form name="frmCI" method="post" action="/admin/board/procComment.asp">
	<input type="hidden" name="hidM" value="CI">
	<input type="hidden" name="ibidx" value="<%=iboard_idx%>"> 
	<input type="hidden" name="hidRT" value="<%=sRegType%>">
<table width="100%"  cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
	<tr>
		<td bgcolor="#FFFFFF" >
			<table width="100%"  cellpadding="5" cellspacing="1" class="a" border="0" >
				<tr>
					<td colspan="2"> * COMMENT </td> 
				</tr>
				<!--
				<tr>
					<td align="center"><textarea id="tCmt" name="tCmt" rows="3" cols="90" ></textarea></td>
					<td><input type="button" value="���" class="button" style="height:50px;width:80px;vertical-align:top;"   id="btnSubmit" onClick="jsSetCmt()"> </td>
				</tr> 
				-->
				<tr>
					<td algin="Center"  colspan="2">	
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
					</td>
				</tr>
			</table> 
		</td>
	</tr>
</table>
</form> 
 