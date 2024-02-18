const decodeBase64 = function (str) {
  if (str == null) return null;
  return atob(str.replace(/_/g, "/").replace(/-/g, "+"));
};

// Call Java Api
const callApi = function (type, uri, data, success_callback, error_callback) {
  let api_url;
  if (location.hostname.startsWith("webadmin")) {
    api_url = "//fapi.10x10.co.kr/api/admin/v1";
  } else {
    api_url = "//testfapi.10x10.co.kr:8080/api/admin/v1";
  }

  if (error_callback === undefined) {
    error_callback = function (xhr) {
      console.log(xhr.responseText);
    };
  }

  $.ajax({
    type: type,
    url: api_url + uri,
    data: data,
    ContentType: "json",
    crossDomain: true,
    xhrFields: {
      withCredentials: true,
    },
    success: success_callback,
    error: error_callback,
  });
};

const callApiHttps = function (type, uri, data, success_callback, error_callback) {
  let api_url;
  if (location.hostname.startsWith("webadmin")) {
    api_url = "//fapi.10x10.co.kr/api/admin/v1";
  } else {
    api_url = "//testfapi.10x10.co.kr:8080/api/admin/v1";
    // api_url = "//localhost:8080/api/admin/v1";
  }

  if (error_callback === undefined) {
    error_callback = function (xhr) {
      console.log(xhr.responseText);
    };
  }

  $.ajax({
    type: type,
    url: api_url + uri,
    data: data,
    ContentType: "json",
    crossDomain: true,
    xhrFields: {
      withCredentials: true,
    },
    success: success_callback,
    error: error_callback,
  });
};

// url ���� ������
const callApiHttpsV2 = function (type, uri, data, success_callback, error_callback) {
  let api_url;
  if (location.hostname.startsWith("webadmin")) {
    api_url = "//fapi.10x10.co.kr/api/admin";
  } else {
    api_url = "//testfapi.10x10.co.kr:8080/api/admin";
    //api_url = "//localhost:8080/api/admin/v1";
  }

  if (error_callback === undefined) {
    error_callback = function (xhr) {
      console.log(xhr.responseText);
    };
  }

  $.ajax({
    type: type,
    url: api_url + uri,
    data: data,
    ContentType: "json",
    crossDomain: true,
    xhrFields: {
      withCredentials: true,
    },
    success: success_callback,
    error: error_callback,
  });
};

const getItemdivName = function(itemdiv){
    switch (itemdiv){
      case null : return "��ǰ ����"; break;
      case "08" :  return "Ƽ��(����)��ǰ"; break;
      case "09" :  return "Present��ǰ"; break;
      case "21" :  return "����ǰ"; break;
      default : return itemdiv;
    }
}

