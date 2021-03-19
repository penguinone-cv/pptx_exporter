# ログから実験まとめスライドを生成するパッケージ

## インストール手順
予めインストールしたい環境に入っておく
1. このリポジトリをClone
2. `cd pptx-exporter`
3. `python setup.py develop`
依存パッケージは一緒に入るので大丈夫

## 使用方法
### テンプレートを作成
任意の場所にスライドのテンプレートを作成
(テンプレートが普段自分が使っているものになっていればタイトルスライドのみでOK)

### 設定ファイルを作る
任意の場所に設定ファイル(ymlファイル)を作成
```yml
name: log                       # 先頭スライドのタイトル
log_path: ./logs                # ログの保存先(logs/hoge/acc.pngのような保存の仕方ならlogsを指定)
template: template.pptx         # テンプレートの場所(自分のテンプレートを使ってタイトルスライドのみのpptxファイルを作っておく)
log_file: parameters.csv        # スライドにテキストボックスとして出しておきたい情報が書かれているファイルの名前(csvかyml(yaml)ファイルに対応)
acc_img_name: accuracy.png      # accuracyの曲線の画像名
loss_img_name: loss.png         # lossの曲線の画像
title_font_size: 28             # 各スライドのタイトルのフォントサイズ
txt_box:
  pos_top: 3.5                  # テキストボックスの位置(y座標)
  pos_left: 3.                  # テキストボックスの位置(x座標)
  width: 15                     # テキストボックスの幅
  height: 10                    # テキストボックスの高さ
  font: Verdana                 # テキストボックスのフォント
  font_size: 18                 # テキストボックスのフォントサイズ
img:
  base_pos_top: 3800000         # 画像の初期位置(y座標)※テキストボックスと画像では位置指定の座標がなぜか違う(python-pptxの仕様)
  base_pos_left: 8000000        # 画像の初期位置(x座標)
  width: 4000000                # 画像の幅※これも座標と同様
  height: auto                  # intを指定すると幅と高さが決めた値に強制される autoで幅に合わせてアスペクト比固定で拡縮を行う
  add_pos_top: -2800000         # 2枚目の画像の初期位置からの移動(y座標)※2枚目の座標は(base_pos_left+add_pos_left, base_pos_top+add_pos_top)
  add_pos_left: 0               # 2枚目の画像の初期位置からの移動(x座標)
```

### 使用方法
```python
import pptx_exporter

# クラス宣言
# テンプレートの呼び出しとタイトルスライドの書き換えを行う
# setting_fileには先ほど作成した設定ファイルのパスを指定
pptx_exporter = pptx_exporter.Makepptx(setting_file="your setting file path")

# ログの保存ディレクトリ名を指定してスライド作成を行う場合
# log_path/hogehoge/acc.pngのような保存の仕方ならlog_dir_name = "hogehoge"
# titleはスライドのタイトル
pptx_exporter.make_slide_from_name(log_dir_name, title)

# 保存されているログを全てスライドに書き出す場合
# log_path内の全ログを抽出(階層が深い場合はエラーが出るため注意)
# slide_titleは各スライドのタイトル
pptx_exporter.make_slides(slide_title)

# pptxファイルを保存
# pathは保存先，"log.pptx"など必ず拡張子を含めること
pptx_exporter.save(path)

# スライドをリセットする場合
pptx_exporter.reset()
```