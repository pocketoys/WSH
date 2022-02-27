// WSHでは、COM（ActiveX）などWindowsが持っている様々なオブジェクトを利用することができます。
// COMは、Windows上で再利用可能なオブジェクト（他のプログラムから利用できる機能をまとめたコンポーネント）を実現する仕組みです。
// COMオブジェクトとして作成されているオブジェクトは、各種プログラムやWSHなどから呼び出して利用することができます。

// Windowsには多様な機能を実現するCOMが標準搭載されているほか、
// 自分で独自のCOMを作成してインストールすることもできるため、
// WSHを使うとWSHの入出力やJavaScriptの処理能力とCOMの独自機能を組み合わせて
// 「Windowsのシステムにアクセスするスクリプト」を作成することができるわけですね。


var fso = new ActiveXObject('Scripting.FileSystemObject');
Editor.MessageBox("Hello World");
var stream = fso.CreateTextFile('G:\\プログラミング\\js\\ランダムデータ作成マクロ\\hoge.txt', true, false);

var allLow = Editor.GetLineCount(0);
var array = [];
for (i = 0 ; i < allLow; i++) {
    if (Editor.GetLineStr(i) === "") {
        continue;
    }
    array = Editor.GetLineStr(i).split("\t");
    Editor.MessageBox(array);

    if (array[1] === "G") {
        var mozi = "";
        for (i = 0 ; i < Number(array[3]); i++) {
            mozi = mozi + ggetSelect();
        }
        Editor.MessageBox(mozi);
    }
}
stream.Write(Editor.GetLineStr(i));
stream.Close();


//var text = Editor.GetSelectedString(0);
//Editor.InsText(text);
// console.log(zen[Number(ran_zen) - 1]);

//Editor.InsText(zen[ran]);

// G文字取得
function ggetSelect() {
    var zen = ["あ", "い", "う", "え", "お", "か", "き", "く", "け", "こ", "さ", "し", "す", "せ", "そ",
              "た", "ち", "つ", "て", "と", "な", "に", "ぬ", "ね", "の", "は", "ひ", "ふ", "へ", "ほ",
              "ま", "み", "む", "め", "も", "や", "ゆ", "よ", "ら", "り", "る", "れ", "ろ", "わ", "を",
              "ん"];

    var ran_zen10 = Math.floor(Math.random() * 5);
    var ran_zen = ""

    //console.log(ran_zen10);

    if (ran_zen10 === 0) {
        var ran_zen1 = Math.floor(Math.random() * 6) + 1;
        ran_zen = String(ran_zen1);
    } else {
        var ran_zen1 = Math.floor(Math.random() * 6) + 1;
        ran_zen = String(ran_zen10) + String(ran_zen1);
    }
    
    return zen[Number(ran_zen)];
}