//이중 옵션 인경우 필요
function CheckMultiOption2016(comp){
	setTimeout(function(){
		var frm = document.sbagfrm;
		var compid = comp;
		var compvalue = $("#"+comp).attr("value");
		var compname  = $("#"+comp).attr("name");
		var retname = $(".select"+(comp+1)).find('p').text();
		
		var optSelObj = $(".itemoption input[name='" + compname + "']");

		var PreSelObj = null;
		var NextSelObj = null;
		var ReDrawObj = null;

		if (!optSelObj.length){
			return;
		}

		if ((compid==0)&&(optSelObj.length>1)) {
			NextSelObj = optSelObj[1];
			if (optSelObj.length>2) {
				ReDrawObj = optSelObj[2];
			}else{
				ReDrawObj = optSelObj[1];
			}
		}

		if ((compid==1)&&(optSelObj.length>2)) {
			PreSelObj  = optSelObj[0];
			NextSelObj = optSelObj[2];
			ReDrawObj = optSelObj[2];
		}

		if (compid==2) {
			PreSelObj  = optSelObj[1];
		}

		if ((PreSelObj!=null)&&(PreSelObj.value.length<1)){
			alert('상위 옵션을 먼저 선택 하세요.');
			//console.log(PreSelObj);
			//console.log(comp);
			//comp.value = '';
			//$('.select'+comp).closest(".scrollArea").prev("p").text(retname);
			PreSelObj.focus();
			return;
		}

		// 최 하위만 품절 세팅
		var found = false;
		var issoldout = false;


		if ( (compvalue.length>0) && (( (ReDrawObj!=null)&&(optSelObj.length-compid==2) )||( (ReDrawObj!=null)&&(optSelObj.length-compid==3)&&(NextSelObj.value.length>0) ))) {
			for (var i=0; i<NextSelObj.length; i++){
				if (NextSelObj.options[i].value.length<1) continue;

				found = false;
				issoldout = false;
				for (var j=0;j<Mopt_Code.length;j++){
					// Box2Ea, Select1-Change
					if ((compid==0)&&(optSelObj.length==2)){
						if (Mopt_Code[j].substr(1,1)==compvalue.substr(1,1)&&(Mopt_Code[j].substr(2,1)==ReDrawObj.options[i].value.substr(1,1))){
							found = true;
							ReDrawObj.options[i].style.color= "#888888";
							break;
						}
					}

					// Box3Ea, Select2-Change
					else if ((compid==1)&&(optSelObj.length==3)) {
						if ((Mopt_Code[j].substr(1,1)==PreSelObj.value.substr(1,1))&&(Mopt_Code[j].substr(2,1)==comp.value.substr(1,1))&&(Mopt_Code[j].substr(3,1)==ReDrawObj.options[i].value.substr(1,1))){
							found = true;
							ReDrawObj.options[i].style.color= "#888888";
							break;
						}
					}

					// Box3Ea, Select2 Value Exists, Select1-Change
					else if ((compid==0)&&(optSelObj.length==3)&&(NextSelObj.value.length>0)){
						if ((Mopt_Code[j].substr(1,1)==compvalue.substr(1,1))&&(Mopt_Code[j].substr(2,1)==NextSelObj.value.substr(1,1))&&(Mopt_Code[j].substr(3,1)==ReDrawObj.options[i].value.substr(1,1))){
							found = true;
							ReDrawObj.options[i].style.color= "#888888";
							break;
						}
					}
				}


				if (!found){
					ReDrawObj.options[i].text = ReDrawObj.options[i].value.substr(2,255) + " (품절)";
					ReDrawObj.options[i].id = "S";
					ReDrawObj.options[i].style.color= "#DD8888";
				}else{
					if (Mopt_S[j]==true){
						ReDrawObj.options[i].text = ReDrawObj.options[i].value.substr(2,255) + " (품절)";
						ReDrawObj.options[i].id = "S";
						ReDrawObj.options[i].style.color= "#DD8888";
					}else{
						if ( Mopt_LimitEa[j].length>0){
							ReDrawObj.options[i].text = ReDrawObj.options[i].value.substr(2,255) + " (한정 " + Mopt_LimitEa[j] + " 개)";
						}else{
							ReDrawObj.options[i].text = ReDrawObj.options[i].value.substr(2,255);
						}
						ReDrawObj.options[i].style.color= "#888888";
						ReDrawObj.options[i].id = "";
					}
				}
			}
		}
	},0);
}

function plusComma(num){
	if (num < 0) { num *= -1; var minus = true}
	else var minus = false

	var dotPos = (num+"").split(".")
	var dotU = dotPos[0]
	var dotD = dotPos[1]
	var commaFlag = dotU.length%3

	if(commaFlag) {
		var out = dotU.substring(0, commaFlag)
		if (dotU.length > 3) out += ","
	}
	else var out = ""

	for (var i=commaFlag; i < dotU.length; i+=3) {
		out += dotU.substring(i, i+3)
		if( i < dotU.length-3) out += ","
	}

	if(minus) out = "-" + out
	if(dotD) return out + "." + dotD
	else return out
}

function goCategoryList(disp){
	top.location.href = "/diyshop/diyList.asp?dispCate="+disp+"";
}

function jsWishBtn(gb,itemid){
<% If IsUserLoginOK() Then %>
	$.ajax({
		url: "/search/act_wishproc.asp?gb="+gb+"&itemid="+itemid+"",
		cache: false,
		success: function(message) {
			if(message!="") {
				if(message == "I"){
					$("#favbtn").removeClass("favorOn");
					$("#favbtn").addClass("favorOn");
					$("#favbtnCnt").html(parseInt($("#favbtnCnt").html())+1);
				}else if(message = "D"){
					$("#favbtn").removeClass("favorOn");
					$("#favbtnCnt").html(parseInt($("#favbtnCnt").html())-1);
				}
			}
		}
		,error: function(err) { alert(err.responseText); }
	});
<% Else %>
	if(confirm("관심 등록은 로그인이 필요한 서비스입니다.\n로그인하시겠습니까?") == true) {
		location.href = "<%=SSLUrl%>/login/login.asp?backpath=<%=Server.URLEncode(CurrURLQ())%>";
		return true;
	} else {
		return false;
	}
<% End If %>
}