onmessage = function(event) {
    var AA = event.data; // event.data
    //ここに処理を記述する
    var results = AA;
    postMessage(results);
}