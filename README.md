# fumen2xlsx

[テト譜](http://fumen.zui.jp) のフィールド・コメントを、Excelのセルに色をつけて転写するツールです。

コマンドラインから実行できます。

※ プログラムの実行には、[Java8](https://www.java.com/ja/download/) 以降を実行できる環境が必要です

### 基本コマンド

```java -jar fumen2xlsx.jar -t v115@zgwhIewhH8AewhH8AewhH8AeI8KepIJvhA```

### オプション一覧

* `-t` `--tetfu`: テト譜を指定する
* `-s` `--size`: 出力されるセルのサイズを小数で指定する [default: `2.0`]
* `-l` `--line`: 出力されるフィールドの高さ。0以下の場合、最も高いページに合わせる [default: `-1`]
* `-n` `--num`: 出力されるフィールドの横に並ぶ数 [default: `5`]
* `-c` `--color`: 色の定義ファイル名。themeディレクトリ内のファイル名に対応 [default: `default`]
* `-C` `--comment`: テト譜のコメントを出力するかどうか [default: `yes`]
* `-o` `--output`: 出力先のファイルパス [default: `output/output.xlsx`]
