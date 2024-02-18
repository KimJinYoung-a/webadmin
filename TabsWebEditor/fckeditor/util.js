function getCurrentPageFullURL(sub)
{
    var x = document.location.href;
    var y = x.substr(0, x.lastIndexOf('/'));
    return y + sub;
}

function getFileSize(fileSize) {
    var retFileSize;
    if (fileSize == 0) {
        retFileSize = '0KB';
    } else if (fileSize < 1024) {
        retFileSize = '1KB';
    } else if (fileSize < (1024 * 1024)) {
        retFileSize = formatNumber((fileSize / 1024), 1) + 'KB';
    } else {
        retFileSize = formatNumber((fileSize / (1024 * 1024)), 1) + 'MB'; 
    }
    return retFileSize; 
}

function formatNumber(number, digit) {
    var digits = 10 * digit;
    return Math.round(number * digits) / digits;
}