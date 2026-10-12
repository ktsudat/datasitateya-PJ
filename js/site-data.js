/*
  お知らせ追加: NEWS_ITEMS の先頭に {date, title, body} を複製して追加します。
  ホームには配列先頭の最新1件のタイトルだけが表示されます。

  プロジェクト追加: PROJECT_ITEMS に {title, description, image, url} を複製して追加します。
  画像は foto フォルダーに入れ、image に "foto/ファイル名" を指定します。
*/
window.NEWS_ITEMS = [
  { date:"2026-10-08", title:"アンケートデータのコーディング支援機能を更新しました", body:"質問票PDFの読込み、回答形式の修正、選択肢の並べ替え、「その他・自由記述」の設定、リアルタイムプレビューに対応しました。" },
  { date:"2026-10-06", title:"Excel・CSV出力と一時保存に対応しました", body:"加工結果の保存と、作業途中の一時保存に対応しました。" }
];
window.PROJECT_ITEMS = [
  { title:"おおいた地域交通見える化ボード", description:"大分県内の公共交通の利用状況を、独自調査のデータにもとづいて分かりやすくご覧いただけるポータルサイトです。", image:"foto/001.png", url:"https://ktsudat.github.io/mierukaPJ/" },
  { title:"ダイヤグラム作成ツール｜スジ屋", description:"ブラウザ上で、ダイヤグラム（スジ）を手軽に作成し、画像やPDFで保存できるツールです。", image:"foto/suziyaTOP.png", url:"https://ktsudat.github.io/daiyaguramupj/" },
  { title:"ホーバー運行記録帳", description:"大分空港と大分市西大分を結ぶホーバークラフトの、日々の運行・欠航情報を記録した非公式サイトです。", image:"foto/hoverTOP.png", url:"https://sites.google.com/view/oita-hover-log" }

];
window.RELATED_LINKS = [
  { label:"おおいた地域交通見える化ボード", url:"https://ktsudat.github.io/mierukaPJ/" },
  { label:"ダイヤグラム作成ツール｜スジ屋", url:"https://ktsudat.github.io/daiyaguramupj/" },
  { label:"ホーバー運行記録帳", url:"https://sites.google.com/view/oita-hover-log" },
  { label:"お問い合わせフォーム", url:"https://forms.gle/gLgvJF1ipYR7q5Ua7" }
];
