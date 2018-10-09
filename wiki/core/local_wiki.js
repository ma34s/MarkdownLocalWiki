/*------------------------------------------------------------------**
	Definition Value of VBScript
**------------------------------------------------------------------*/
// 保存データの種類
		// StreamTypeEnum
		// http://msdn.microsoft.com/ja-jp/library/cc389884.aspx
var adTypeBinary = 1; // バイナリ
var adTypeText   = 2; // テキスト

// 読み込み方法
		// StreamReadEnum
		// http://msdn.microsoft.com/ja-jp/library/cc389881.aspx
var adReadAll  = -1; // 全行
var adReadLine = -2; // 一行ごと

// 書き込み方法
		// StreamWriteEnum
		// http://msdn.microsoft.com/ja-jp/library/cc389886.aspx
var adWriteChar = 0; // 改行なし
var adWriteLine = 1; // 改行あり

// ファイルの保存方法
		// SaveOptionsEnum 
		// http://msdn.microsoft.com/ja-jp/library/cc389870.aspx
var adSaveCreateNotExist  = 1; // ない場合は新規作成
var adSaveCreateOverWrite = 2; // ある場合は上書き

/*------------------------------------------------------------------**
	Local Wiki
**------------------------------------------------------------------*/
//グローバル変数を設定
var fso = new ActiveXObject("Scripting.FileSystemObject");
var oBaseFolder = getBaseFolder();
var pageNameStack = new Array();

//このHTMLファイルが置かれているフォルダのパス名を取得する
function getBaseFolder(){
    return fso.GetFolder(CONFIG.base_dirctory);
}


//エクスプローラでフォルダ・ファイルを開く
function OpenFolder(Path){
 //   alert(Path);
   var ss = Path.split(',');
   //  Shell関連の操作を提供するオブジェクトその2を取得
   var sh = new ActiveXObject( "Shell.Application" );
   var Path2 = "";

//     alert(ss.length);

   for(var i=0; i < ss.length; i++){
      if(!(ss[i] == "")){
        Path2 += ss[i] + '\\';
      };
   }
//  alert(Path2);

   //エクスプローラで開く
   sh.Open(Path2);
   //  オブジェクトを解放
   sh = null;
}

//ページ名を指定してページを開く
function open(pagename){
    var content = getContent(pagename);
    if(! content){
        return edit(pagename);
    }


//    var renderer = new marked.Renderer();
//    renderer.link = function(href, title, text) {
//     renderer.link = function (text, level) {
//    alert("text=" +  text + "level" + level );
//      alert(href  + "  " +  title + "  " + text );
//    };
//    renderer.heading = function (text, level) {
//	alert("text=" +  text + "level" + level );
//
//    var escapedText = text.toLowerCase().replace(/[^\w]+/g, '-');
//      return '<h' + level + '><a name="' +
//                escapedText +
//                 '" class="anchor" href="#' +
//                 escapedText +
//                 '"><span class="header-link"></span></a>' +
//                  text + '</h' + level + '>';
//    };


//    alert(marked('# heading+', { renderer: renderer }));
//    var html = marked(content, { renderer: renderer });


    //いわゆるmarked.jsを使用（一部改造）
    var html = marked(content);




    //ファイル名、更新日時等を付与
    if(CONFIG.showFileInfo == true)
    {
	    id('FileName').innerText = getFilePathAndName(pagename);
	    id('DateCreated').innerText = "作成日時:" + getDateCreated(pagename);
	    id('DateLastModified').innerText = "更新日時:" + getDateLastModified(pagename);
	}

    //表示履歴
    setHistory(pagename);
    getHistory();
    
    setPageName(pagename);
    setContent(html);
    showEditLink();
	createToc();
}

//履歴保存
function setHistory(pagename) {

    if(CONFIG.showHistory== false)
    {
        return;
    }
    //スタックに履歴を保存
    if (0 == pageNameStack.length) {
        //初回は無条件で保存
        pageNameStack.push(pagename);
    } else {
        //保存されている履歴外のみ保存
        var matched = false;
        for (var i = 0; i < pageNameStack.length; i++) {
            var tmp = pageNameStack[i];
            if (tmp == pagename) {
                matched = true;
            }
        }
        if (matched == false) {
            pageNameStack.push(pagename);
        }
    }
}

function getHistory() {
    if(CONFIG.showHistory== false)
    {
        return;
    }
    //スタック内の履歴を表示
    var html = "表示履歴："
    for (var i = 0; i < pageNameStack.length; i++) {
        var tmp = pageNameStack[i];
        html = html + "<a href='javascript:open(\"" + tmp + "\")'>" + tmp + "</a> ";
    }
    html = html + " ";
    id('history').innerHTML = html;
}

//マークダウン（md)ファイル名をフルパスで取得
function getFilePathAndName(pagename){
    var s = '';
    var filepath = getFilePath(pagename);
    var f = fso.GetFile(filepath);
    s = f.Path ;
    return(s);
}

//ファイルの作成日時を取得する
function getDateCreated(pagename){
   var s = '';
   var filepath = getFilePath(pagename);
   var f = fso.GetFile(filepath);
   var CreateDate = new Date(f.DateCreated);
   s += CreateDate.toLocaleDateString() + " " + CreateDate.toLocaleTimeString();  
   return(s);

}
//ファイルの更新日時を取得する
function getDateLastModified(pagename) {
    var s = '';
    var filepath = getFilePath(pagename);
    var f = fso.GetFile(filepath);
    var ModDate = new Date(f.DateLastModified);
    s += ModDate.toLocaleDateString() + " " + ModDate.toLocaleTimeString();
    return (s);
}

//ページ名を指定してページを編集する
function edit(pagename){
    var content = getContent(pagename);

//    var html =
//           "<form  onsubmit='save(\"" + pagename + "\"); return false;'>"
//         + "<textarea cols='120' rows='30' id='textarea' wrap='off'>" + content + "</textarea><br />"
//         + "<input type='submit' value='保存'><br />"
//         + "</form>";
//編集方法の修正 textareaへの格納はフォーム生成後のtextareaへテキスト入力
    var html =
           "<form  onsubmit='save(\"" + pagename + "\"); return false;'>"
         + "<textarea cols='120' rows='30' id='textarea' wrap='off'>" + "</textarea><br />"
         + "<input type='submit' value='保存'><br />"
         + "</form>";


    setPageName(pagename);
    setContent(html);
    id('textarea').innerText = content;
    hideEditLink();
}

//いま見ているページを編集する
function editCurrentPage(){
    pagename = id('headerH1').innerText;
    edit(pagename);
}

//フォームのテキストエリアの中身を、指定したページのコンテンツとして保存する
function save(pagename){

//編集方法の修正 textareaへの格納はフォーム生成後のtextareaへテキスト入力
//    var content = id('textarea').value;
    var content = id('textarea').innerText;

    var filepath = getFilePath(pagename);
    var file = fso.createTextFile(filepath);
    utf8_saveToFile(filepath, content);
    open(pagename);

}

//ページのリストを取得する
function getPageList(){
    var enuFiles = new Enumerator(oBaseFolder.Files);
   
    var myFiles = [];
    for (; !enuFiles.atEnd(); enuFiles.moveNext() ){ 
            var FilePath = enuFiles.item(); 
            var ExtensionName = fso.GetExtensionName(FilePath); // 拡張子を取得
            var BaseName = fso.GetBaseName(FilePath); // ベースネームを取得
            if(ExtensionName == "md"){ // 拡張子がmdだったら
                myFiles.push(BaseName);
            }
    }
    return myFiles;
}

//ページ一覧/検索画面を表示する
function viweCreatedList(myFiles,title){
    var list = [];
    list.push('|ファイル名|作成日|更新日|');
    list.push('|:-|:-:|:-:|');
    //ファイル保存用、表示用のフォーマット変換
    for(var i = 0; i < myFiles.length; i++){
//        list.push('<li><a href="javascript:open(\'' + myFiles[i] + '\');">' + myFiles[i] + '</a></li>');
//        list.push('<li>[' + myFiles[i] + '](' + myFiles[i] + ')</li>' );
        list.push('|[' + myFiles[i] + '](' + myFiles[i] + ')|' + getDateCreated(myFiles[i]) +'|' + getDateLastModified(myFiles[i]) +'|' );
    }

    var content = list.join("\r\n");

    if(CONFIG.showHistory==true)
    {//ファイル保存
        var filepath = getFilePath( title );
        utf8_saveToFile(filepath, content);
    }

    //いわゆるmarked.jsを使用（一部改造）
    var html = marked(content);

    //表示履歴
    setHistory( title );
    getHistory();

    //表示
    setPageName( title );
    setContent( html );
    showEditLink();
}

//ページ一覧画面を表示する
function openIndexPage(){
    var openIndexPageName = '【ページ一覧】';
    var myFiles = getPageList();
    viweCreatedList(myFiles,openIndexPageName);
}

//サイドのTOC作成
function createToc(){
    var myFiles = getPageList();
    var list = [];
    for(var i = 0; i < myFiles.length; i++){
        list.push(' - [' + myFiles[i] + '](' + myFiles[i] + ')' );
    }
    var content = list.join("\r\n");
    var html = marked(content);
    id('toc').innerHTML = html;
    fixed_side_toc();
}


//ページ名を指定して、該当するマークダウン（.md)のパス名を取得する
function getFilePath(pagename){
    var filename = pagename + ".md";
    return oBaseFolder.Path + "\\" +  filename;
}

//ページ名を指定して、該当するマークダウンファイルの中身を取得する
function getContent(pagename){
    var content = '';
    var filepath = getFilePath(pagename);

    if(fso.fileExists(filepath)){
        content = utf8_readAll(filepath);
    }
    return content;
}

//メニューの「編集」リンクを表示する
function showEditLink() {
    id('editLink').style.display = "inline";
}

//メニューの「編集」リンクを非表示にする
function hideEditLink() {
    id('editLink').style.display = "none";
}

//コンテンツを画面に表示する
function setContent(html) {
    id('content').innerHTML = html;
}

//ページ名を画面に表示する
function setPageName(pagename) {
    if(pagename){
        id('headerH1').innerText = pagename;
    }
}

//画面のページ名を取得する
function getPageName() {
    return id('headerH1').innerText;
        
}
//HTMLの要素を取得する
function id(s) {
    return document.getElementById(s);
}

//検索にヒットしたページの一覧画面を表示する
function FindIndexPage() {
    var enuFiles = new Enumerator(oBaseFolder.Files);
    var myFiles = [];
    var FindIndexPageName = '【検索結果】';

    for (; !enuFiles.atEnd(); enuFiles.moveNext()) {
        var FilePath = enuFiles.item();
        var ExtensionName = fso.GetExtensionName(FilePath); // 拡張子を取得
        var BaseName = fso.GetBaseName(FilePath); // ベースネームを取得

        if (ExtensionName == "md") { // 拡張子がmdだったら
            //ファイル名検索
            var database = FilePath;
            var sword = Text1.value;
            var check = BaseName.indexOf(sword, 0);
            if (0 <= check) {
                //ファイル名を格納
                myFiles.push(BaseName);
            }
            else {
                //ファイルの中身を検索
                if (0 <= openAndSerch(BaseName) )
                {
                    //ファイル名を格納
                    myFiles.push(BaseName);
                }
            }
        }
    }
    viweCreatedList(myFiles,FindIndexPageName  + Text1.value);
}

//ページ名を指定してページを開き、検索文字がヒットしたかチェックする
function openAndSerch(pagename) {
    var content = getContent(pagename);
    if (!content) {
        return edit(pagename);
    }
    //ファイル名検索
    var database = content;
    var sword = Text1.value;
    var check = database.indexOf(sword, 0);
    return check
}

//新規作成
function NewPage() {
    var pagename = Text2.value;
//    window.confirm(pagename);
    if (pagename) {
        open(pagename);
    }
    Text2.value = '';
}
//リネーム
function ReNamePage() {

    var srcPagename = getPageName();
    var dstPagename = Text2.value;

    // 「OK」時の処理開始 ＋ 確認ダイアログの表示
    if (window.confirm(dstPagename + ' にリネームしますか？'))
    {

        var srcFilepath = getFilePath(srcPagename);
        var dstFilepath = getFilePath(dstPagename);
    
        fso.MoveFile(srcFilepath, dstFilepath);
        window.alert(srcPagename + ' を ' + dstPagename + ' にリネームしました。');         
        open(dstPagename);
    }
    else {
        window.alert('キャンセルされました。'); // 警告ダイアログを表示
    }
    Text2.value = '';
}

//削除
function DeletePage() {

    var pagename = getPageName();

    // 「OK」時の処理開始 ＋ 確認ダイアログの表示
    if (window.confirm(pagename + ' を削除しますか？')) {

        var pagename = getPageName();
        var filepath = getFilePath(pagename);

        fso.DeleteFile(filepath);
        window.alert(pagename + ' を削除しました。'); 
        openTopPage();
    }
    else {
        window.alert('キャンセルされました。'); // 警告ダイアログを表示
    }
    Text2.value = '';
}

function utf8_readAll(filename){
    var sr = new ActiveXObject("ADODB.Stream");
    sr.Type = adTypeText;
    sr.charset = "utf-8";
    sr.Open();
    sr.LoadFromFile( filename );
    var temp = sr.ReadText( adReadAll );
    sr.Close();
    return temp;
}


function utf8_saveToFile(filename, text) {
	// ADODB.Streamのモード
	var adTypeBinary = 1;
	var adTypeText = 2;
	// ADODB.Streamを作成
	var pre = new ActiveXObject("ADODB.Stream");
	// 最初はテキストモードでUTF-8で書き込む
	pre.Type = adTypeText;
	pre.Charset = 'UTF-8';
	pre.Open();
	pre.WriteText(text);
	// バイナリモードにするためにPositionを一度0に戻す
	// Readするためにはバイナリタイプでないといけない
	pre.Position = 0;
	pre.Type = adTypeBinary;
	// Positionを3にしてから読み込むことで最初の3バイトをスキップする
	// つまりBOMをスキップします
	pre.Position = 3;
	var bin = pre.Read();
	pre.Close();

	// 読み込んだバイナリデータをバイナリデータとしてファイルに出力する
	// ここは一般的な書き方なので説明を省略
	var stm = new ActiveXObject("ADODB.Stream");
	stm.Type = adTypeBinary;
	stm.Open();
	stm.Write(bin);
	stm.SaveToFile(filename, 2); // force overwrite
	stm.Close();
};

//TopPageを表示する
function openTopPage(){
    open(CONFIG.topPage);
}

function initNavigation(){
    marked.setOptions({
        image_base: CONFIG.base_dirctory +"/"
    })


    id('topPageLinkA').innerText = CONFIG.topPage;
    id('editLinkA').innerText = VALUES.editTitle;
    id('PagelistLinkA').innerText = VALUES.pagelistTitle;
    id('SarchLinkA').innerText = VALUES.sarchTitle;

    id('newPageLinkA').innerText = VALUES.newPageTitle;
    id('renamePageLinkA').innerText = VALUES.reNamePageTitle;
    id('deletePageLinkA').innerText = VALUES.deletePageTitle;

    if(CONFIG.showHistory == false)
    {
        id('HRhistry').style.display = "none";
    }
    if(CONFIG.showFileInfo == false)
    {
        id('HRfileinfo').style.display = "none"
    }
    openTopPage();
}
