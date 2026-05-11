# 追善法要のしおり 生成システム

## 必要なもの
- Python 3.9 以上（https://www.python.org/downloads/）
- このフォルダ内の 2 ファイル
  - `app.py`
  - `template.docx`

---

## 初回セットアップ（1回だけ）

コマンドプロンプト（Windows）または ターミナル（Mac）を開いて、
このフォルダに移動してから以下を実行：

```
pip install streamlit
```

---

## 起動方法

```
streamlit run app.py
```

自動でブラウザが開きます。

---

## クラウド公開（URL で共有したい場合）

1. https://github.com で無料アカウントを作成
2. 新しいリポジトリを作成（**Private** 推奨）
3. `app.py` `template.docx` `requirements.txt` の 3 ファイルをアップロード
4. https://share.streamlit.io にアクセス → GitHub ログイン
5. リポジトリを選択して Deploy
6. 生成された URL を上司・同僚に共有

---

## 注意事項
- `template.docx` は同フォルダから移動しないでください
- クラウド公開時はリポジトリを Private に設定することを推奨します
