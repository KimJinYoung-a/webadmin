	if($("input[name=multiimg1]").val() == ""){
		alert("multi 이미지를 등록해주세요.");
		return;
	}
	if($("input[name=multiimg2]").val() == ""){
		alert("multi 이미지를 등록해주세요.");
		return;
	}
	if($("input[name=multiimg3]").val() == ""){
		alert("multi 이미지를 등록해주세요.");
		return;
	}
	
	for(var i=1; i<13; i++){
		if($("input[name=itemimg"+i+"]").val() == ""){
			alert("item"+i+" 이미지를 선택하세요.");
			return;
		}
	}
	
	for(var i=1; i<5; i++){
		if($("input[name=eventimg"+i+"]").val() == ""){
			alert("event"+i+" 이미지를 선택하세요.");
			return;
		}
	}
	
	if($("input[name=bookimg]").val() == ""){
		alert("bookimg 이미지를 등록해주세요.");
		return;
	}
	
	if($("input[name=recipeimg]").val() == ""){
		alert("recipeimg 이미지를 등록해주세요.");
		return;
	}