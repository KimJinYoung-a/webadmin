<html>
<body>
<%
dim swfName    : swfName = "DzEBankFlexAPI"
%>

<script src="AC_OETags.js" language="javascript"></script>
<script language="JavaScript" type="text/javascript">
<!--
	AC_FL_RunContent(
    	"src", "<%= swfName %>",
    	"width", "100%",
    	"height", "100",
    	"align", "middle",
    	"id", "<%= swfName %>",
    	"quality", "high",
    	"bgcolor", "#869ca7",
    	"name", "<%= swfName %>",
    	"allowScriptAccess","always", 
    	"type", "application/x-shockwave-flash",
    	"pluginspage", "http://www.adobe.com/go/getflashplayer"
    );
// -->
</script>


<script language='javascript'>
	 
function thisMovie(movieName){
    if(navigator.appName.indexOf("Microsoft") != -1){
        return window[movieName];
    }else {
        return document[movieName];
    }
}

function AddNew(key, value)
{
 var obj = new Object();
 obj.key = key;
 obj.value = value;
 return obj;
}
     

//01.로그인
function FxLogin(iid,ipwd){
    
    var obj = AddNew("ID", iid);     
    var obj1 = AddNew("PASSWD", ipwd); 
    var obj2 = AddNew("USER_IP", "<%= request.ServerVariables("REMOTE_ADDR") %>");    
    
    var arr = new Array(obj, obj1, obj2);
    
    thisMovie("<%= swfName %>").Login(arr);
    
}
     
//01.로그인 결과     
function FxLoginResult(retObj){   
    var ret = retObj.RESULT;
    alert(ret);
    if (ret=="00000"){
        //사업자번호 체크
        alert(retObj.NO_CUST);
    }else{
        alert(retObj.RESULT_MSG);
    }
    //alert("RESULT:" + retObj.RESULT); 
    //alert("NO_CUST:" + retObj.NO_CUST);
    //alert("RESULT:" + retObj.RESULT_MSG); 
}

function DzErrorEvent(faultEvent)
     {
        var errinfo = "";

        errinfo = "faultEvent.message:" + faultEvent.message + "\n";
        errinfo += "faultEvent.errorID:" + faultEvent.errorID + "\n";
        errinfo += "faultEvent.faultCode:" + faultEvent.faultCode + "\n";
        errinfo += "faultEvent.faultDetail:" + faultEvent.faultDetail + "\n";                
        errinfo += "faultEvent.faultString:" + faultEvent.faultString + "\n";  
        
        //form1.fxlog.value = errinfo;
        alert(errinfo);                    
     }    

</script>
<input type="button" value="login" onclick="FxLogin('tenbyten','20011010')">

</body>
</html>