var querystring = require('querystring');
var https = require('https');
var schedule = require("node-schedule");
var gui = require('nw.gui');
var menu = new gui.Menu();
var tray = new gui.Tray({ icon: 'img/calendar.png' });
var fs = require('fs');
var path = require('path');

var retryCount = 0;
var retryLimit = 5;
var host = "outlook.office365.com";

menu.append(new gui.MenuItem({
    label: 'exit',
    click: function() {
        setTimeout('bye()', 100);
    }
}));
tray.menu = menu;

//    var conf = getConf();

var count = function(response) {
    response.on('data', function(chunk) {
        console.log(chunk.toString());
        var counting = chunk.toString();
        if (counting) {
            var i = 0;
            getDetail(i, Number(counting));
        } else {
            bye();
        }
    });
}

function getDetail(i, counting) {
    if(i === counting) return;

    httpGet(getDetailUrl(conf, i), detail);
    i++;
    setTimeout(function(){
        getDetail(i, counting);
    }, 10000);
}

var detail = function(response) {
    response.on('data', function(chunk) {
        console.log(chunk);
        console.log(chunk.toString());
        var value = chunk.toString().match(/\"value\":\[.*\]/);
        if (value) {
            var result = unescapeUnicode(value[0]);
            var event = JSON.parse('{' + result + '}').value[0];
            if(event) showEvent(event);
        }
    });
}

var job = schedule.scheduleJob('0,15,25,30,45,55 * * * *', function () {
    httpGet(getCountUrl(conf), count);
});

httpGet(getCountUrl(conf), count);


function getConf() {
    var confPath = path.dirname(process.execPath) + '\\config.json';
    var conf = require(confPath);

    if(conf == null || conf.id == null || conf.password == null) {
        setTimeout('fail()', 5000);
    }
    return conf;
}

function bye() {
    exit('END', 'goodbye.');
}

function fail() {
    exit('*** ERROR ***', 'config.json is invalid.');
}

function exit(title, message) {
    showNotification(title, message);
    tray.remove();
    gui.App.quit();
}

function getEventToken() {
    return getUrlInfo('/api/v1.0/me/calendarview/?startdatetime=' + getDate() + 'T00:00:00Z&enddatetime=' + getDate() + 'T23:59:59Z&$select=Subject,Start,End,Location&$orderby=Start&$filter=Start%20ge%20' + getStartTime() + '%20and%20Start%20le%20' + getEndTime() + '%20and%20IsCancelled%20eq%20false');
}

/*
function getCountToken() {
    return getUrlInfo('/api/v1.0/me/calendarview/$count?startdatetime=' + getDate() + 'T00:00:00Z&enddatetime=' + getDate() + 'T23:59:59Z&$select=Subject,Start,End,Location&$orderby=Start&$filter=Start%20ge%20' + getStartTime() + '%20and%20Start%20le%20' + getEndTime() + '%20and%20IsCancelled%20eq%20false');
}
*/

function getToken(path) {
    var conf = getConf();
    var auth = conf.id + ':' + conf.password;
    return  {
        host: host,
        port: 443,
        path: path,
        auth: auth
    };
}

//TODO çÌèúëŒè€
/*
function getCountUrl(conf) {
    var auth = conf.id + ':' + conf.password;
//    var path = '/api/v1.0/me/calendarview/$count?startdatetime=' + getStartTime() + '&enddatetime=' + getEndTime(30) + '&$select=Subject,Start,End,Location&$orderby=Start';
    var path = '/api/v1.0/me/calendarview/$count?startdatetime=' + getDate() + 'T00:00:00Z&enddatetime=' + getDate() + 'T23:59:59Z&$select=Subject,Start,End,Location&$orderby=Start&$filter=Start%20ge%20' + getStartTime() + '%20and%20Start%20le%20' + getEndTime() + '%20and%20IsCancelled%20eq%20false';

    return  {
        host: host,
        port: 443,
        path: path,
        auth: auth
    };
}
*/
/*
function getDetailUrl(conf, counting) {
    var auth = conf.id + ':' + conf.password;
//    var path = '/api/v1.0/me/calendarview?startdatetime=' + getStartTime() + '&enddatetime=' + getEndTime(30) + '&$select=Subject,Start,End,Location&$orderby=Start&$top=1&$skip=' + counting;
    var path = '/api/v1.0/me/calendarview/?startdatetime=' + getDate() + 'T00:00:00Z&enddatetime=' + getDate() + 'T23:59:59Z&$select=Subject,Start,End,Location&$orderby=Start&$filter=Start%20ge%20' + getStartTime() + '%20and%20Start%20le%20' + getEndTime() + '%20and%20IsCancelled%20eq%20false&$top=1&$skip=' + counting;

    return  {
        host: host,
        port: 443,
        path: path,
        auth: auth
    };
}
*/

/*
function httpGet(urlInfo, func) {
    var retry = function() {
        console.log('retry');
        httpGet(urlInfo, func);
    }
    var request = https.get(urlInfo,
        function(response) {
            console.log('Response: ' + response.statusCode);
            if (response.statusCode === 200) {
                func(response);
                retryCount = 0;
            } else {
                if(retryCount < retryLimit) {
                    retryCount++;
                    retry();
                } else {
                    showNotification('*** ERROR ***', 'fail to connect office365.');
                    setTimeout('exit()', 5000);
                }
            }
        }
    );
}
*/
function getEvent(token, dataCallback, endCallback) {
    var retryCount = 0;

    var retry = function() {
        console.log('retry');
        getEvent(token, dataCallback, endCallback);
    }
    var request = https.get(token,
        function(response) {
            console.log('Response: ' + response.statusCode);
            if (response.statusCode === 200) {
                response.setEncoding('utf-8');
                response.on('data', dataCallback);
                response.on('end',endCallback);
            } else {
                if(retryCount < retryLimit) {
                    retryCount++;
                    retry();
                } else {
                    setTimeout('fail()', 5000);
                }
            }
        }
    );
}

function showEvent(event) {
    var start = new Date(event.Start);
    var end = new Date(event.End);
    
    showNotification(event.Subject, 'Location: ' + event.Location.DisplayName + '\nStart: ' + start.toLocaleString() + '\nEnd: ' + end.toLocaleString());
}

function showNotification(title, body) {
    var notification = new Notification(title, {
        body: body
    });

    notification.onshow = function() {
        setTimeout(function() {
            notification.close();
        }, 5000000);
    }
}

// TODO Ç±ÇÍÇ‡Ç¢ÇÁÇ»Ç¢Ç©Ç‡
function unescapeUnicode(string) {
    return string.replace(/\\u([a-fA-F0-9]{4})/g, function(matchedString, group1) {
        return String.fromCharCode(parseInt(group1, 16));
    });
}

function getStartTime() {
    return getDateTime(new Date(), -15);
}

function getEndTime() {
    return getDateTime(new Date(), 15);
}

function getDate() {
    return getDateTime(new Date(), 0, 'YYYY-MM-DD');
}

function getDateTime(date, plusMin, format) {
    if(!format) format = 'YYYY-MM-DDThh:mm:00Z';
    if(plusMin) date.setUTCMinutes(date.getUTCMinutes() + plusMin);
    format = format.replace(/YYYY/g, date.getUTCFullYear());
    format = format.replace(/MM/g, ('0' + (date.getUTCMonth() + 1)).slice(-2));
    format = format.replace(/DD/g, ('0' + date.getUTCDate()).slice(-2));
    format = format.replace(/hh/g, ('0' + date.getUTCHours()).slice(-2));
    format = format.replace(/mm/g, ('0' + date.getUTCMinutes()).slice(-2));
    return format;
}

process.on('uncaughtException', function(err) {
    console.log('uncaughtException => ' + err);
    showNotification('*** ERROR ***', 'uncaughtException => ' + err);
});