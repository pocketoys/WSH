// WSHでは、COM（ActiveX）などWindowsが持っている様々なオブジェクトを利用することができる。
// COMは、Windows上で再利用可能なオブジェクト（他のプログラムから利用できる機能をまとめたコンポーネント）を実現する仕組み。
// COMオブジェクトとして作成されているオブジェクトは、各種プログラムやWSHなどから呼び出して利用することができる。

// Windowsには多様な機能を実現するCOMが標準搭載されているほか、
// 自分で独自のCOMを作成してインストールすることもできるため、
// WSHを使うとWSHの入出力やJavaScriptの処理能力とCOMの独自機能を組み合わせて
// 「Windowsのシステムにアクセスするスクリプト」を作成することができる。


var fso = new ActiveXObject('Scripting.FileSystemObject');
var stream = fso.CreateTextFile('G:\\プログラミング\\js\\ランダムデータ作成マクロ\\hoge.txt', true, false);

var allLow = Editor.GetLineCount(0);
var array = [];
var allarray = [];
for (var i = 0 ; i < allLow; i++) {
    if (Editor.GetLineStr(i + 1) === "") {
        continue;
    }
    array = Editor.GetLineStr(i + 1).split("\t");
    if (array[1] === "G") {
        var mozi = "";
        for (var count = 0 ; count < Number(array[3]); count++) {
            mozi = mozi + ggetSelect();
        }
        mozi = mozi + "\r\n";
        Editor.MessageBox("1" + mozi);
        allarray.push(mozi);
    }

    if (array[1] === "C") {
        var mozi = "";
        if (array[2] === "英") {
            for (var count = 0 ; count < Number(array[3]); count++) {
                mozi = mozi + c_eigetSelect();
            }
        } else if (array[2] === "数") {
            for (var count = 0 ; count < Number(array[3]); count++) {
                mozi = mozi + c_suugetSelect();
            }
        }
        mozi = mozi + "\r\n";
        Editor.MessageBox("2" + mozi);
        allarray.push(mozi);
    }
}

var allmozi = "";

for (var i = 0; i < allarray.length; i++) {
    allmozi += allarray[i];
}
Editor.MessageBox("3" + allmozi);
stream.Write(allmozi);
stream.Close();

// G文字取得
function ggetSelect() {
    var zen = ["あ", "い", "う", "え", "お", "か", "き", "く", "け", "こ", "さ", "し", "す", "せ", "そ",
              "た", "ち", "つ", "て", "と", "な", "に", "ぬ", "ね", "の", "は", "ひ", "ふ", "へ", "ほ",
              "ま", "み", "む", "め", "も", "や", "ゆ", "よ", "ら", "り", "る", "れ", "ろ", "わ", "を",
              "ん"];

    var ran_zen = Math.floor(Math.random() * 46);

    return zen[ran_zen];
}

// C文字_英取得
function c_eigetSelect() {
    var han = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o",
              "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"];

    var ran_han = Math.floor(Math.random() * 26);

    return han[ran_han];
}

// C文字_数取得
function c_suugetSelect() {
    var han = ["1", "2", "3", "4", "5", "6", "7", "8", "9"];

    var ran_han = Math.floor(Math.random() * 9);

    return han[ran_han];
}