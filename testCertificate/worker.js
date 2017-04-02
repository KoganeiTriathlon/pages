// キー入力のたびに呼ばれることになる
onmessage = function(message) {
    this.importScripts('./lib/rawdeflate.js');
    this.importScripts('./lib/rawinflate.js');
    this.importScripts('./lib/base64.js');

    // 開始を伝える
    this.postMessage({event: 'start'});

    // 圧縮結果をメッセージ送信(ついでにBase64変換)
    var result = result = Base64.toBase64(RawDeflate.deflate(str));
    this.postMessage({
        event: 'compressed',
        data: result
    });

    // 伸張結果をメッセージ送信
    result = RawDeflate.inflate(Base64.fromBase64(result));
    this.postMessage({
        event: 'complete',
        data: result
    });
}