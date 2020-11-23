# fumen2xlsx

[テト譜](http://fumen.zui.jp) のフィールド・コメントを、Excelのセルに色をつけて転写するツールです。

コマンドラインから実行できます。

※ プログラムの実行には、[Java8](https://www.java.com/ja/download/) 以降を実行できる環境が必要です

### 出力サンプル

[default カラー](https://1drv.ms/x/s!AgZefBjz5OGajWpAuOWWg5yMuF_k): `mipi` をベースにライン消去・操作ミノが強調されるようなカラーリング

[mipi カラー](https://1drv.ms/x/s!AgZefBjz5OGajWlg-CtFxt7WnwLV): [@mipi_teto_puyo](https://twitter.com/mipi_teto_puyo) さんの資料をベースとしたカラーリング

[fumen カラー](https://1drv.ms/x/s!AgZefBjz5OGajWurEbKAn4QKM7ST): [テト譜](http://fumen.zui.jp) をベースとしたカラーリング

### 基本コマンド

```java -jar fumen2xlsx.jar -t v115@vhGSQYIAkeEfEXUb9AJHJWSJUIJXGJVBJTJYYAlvs2?A1sDfETo3ABlvs2A3HEfET4ZOB```

### オプション一覧

* `-t` `--tetfu`: テト譜を指定する
* `-s` `--size`: 出力されるセルのサイズを小数で指定する [default: `2.0`]
* `-l` `--line`: 出力されるフィールドの高さ。0以下の場合、最も高いページに合わせる [default: `-1`]
* `-n` `--num`: 出力されるフィールドの横に並ぶ数 [default: `5`]
* `-c` `--color`: 色の定義ファイル名。themeディレクトリ内のファイル名に対応 [default: `default`]
* `-C` `--comment`: テト譜のコメントを出力するかどうか [default: `yes`]
* `-o` `--output`: 出力先のファイルパス [default: `output/output.xlsx`]


### カラーテーマのカスタマイズ

themeディレクトリにプロパティファイルを追加すると、`--color` で指定できるようになります。

既にあるプロパティファイルをコピーして名前を変更したうえで、直接ファイルを更新してください。、

＊`T,I,O,S,Z,L,J,Gray`: 各ブロックの色
* `Gray`: せり上がりブロックの色
* `Empty`: 空白の色
* `Border`: ブロックまわりの線の色

* `.normal`: 通常時のブロックの色
* `.clear`: ラインが揃ったときに強調するための色
* `.piece`: 操作しているミノを強調するための色
