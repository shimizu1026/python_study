# Pythonのコンテナ入門

## コンテナとは？

コンテナとは**データをまとめて入れておく入れ物**のことです。

コンテナがないと、こうなります。

```python
score1 = 80
score2 = 75
score3 = 92
score4 = 88
# 40人分書くの？😱
```

コンテナがあると、こうなります。

```python
scores = [80, 75, 92, 88, ...]
# すっきり！
```

Pythonの代表的なコンテナは以下の4種類です。

| 入れ物 | 現実のたとえ | 特徴 |
|---|---|---|
| リスト | 買い物リスト | 順番通りに並んでいる・変更可能 |
| タプル | 誕生日 (年,月,日) | 変更不可・固定データ向け |
| 辞書 | 電話帳 | キーと値の対応で管理 |
| セット | 会員リスト | 重複なし・順番なし |

---

## リスト（list）

変更可能（ミュータブル）な順序付きコレクションです。JavaScriptの配列に一番近い存在です。

### 基本的な使い方

```python
fruits = ["apple", "banana", "cherry"]

# インデックスでアクセス（0番から始まる）
print(fruits[0])   # "apple"
print(fruits[1])   # "banana"

# 追加
fruits.append("date")        # 末尾に追加
fruits.insert(1, "mango")    # 指定した位置に追加

# 上書き
fruits[1] = "blueberry"

# 削除
fruits.remove("apple")  # 値を指定して削除
fruits.pop(0)           # インデックスを指定して削除

# 並び替え
fruits.sort()     # アルファベット順
fruits.reverse()  # 逆順

# 含まれているか確認
print("apple" in fruits)      # True
print("grape" not in fruits)  # True

# 要素数
print(len(fruits))  # 数を返す
```

### 要素が1つのとき

```python
fruits = ["apple"]  # 普通に書けばOK
```

### リストの中にリストを入れる

```python
teams = [
    ["Alice", "Bob", "Charlie"],   # 1班
    ["Dave", "Eve", "Frank"],      # 2班
    ["Grace", "Hank", "Ivy"]       # 3班
]

print(teams[0])     # ["Alice", "Bob", "Charlie"]
print(teams[0][1])  # "Bob"
```

### どんなアプリに使われているか

- **SNS**：タイムラインの投稿一覧
- **ECサイト**：商品検索の結果一覧
- **音楽アプリ**：プレイリスト
- **ゲーム**：インベントリ（アイテム一覧）

---

## タプル（tuple）

変更不可（イミュータブル）な順序付きコレクションです。変わってほしくない固定データに使います。

### 基本的な使い方

```python
point = (10, 20)

print(point[0])  # 10
x, y = point     # アンパック

# 要素は変更できない
# point[0] = 5  → TypeError
```

### 要素が1つのとき（注意！）

```python
a = ("apple")   # これはタプルではなく文字列になる！
b = ("apple",)  # 末尾にカンマが必要

print(type(a))  # <class 'str'>
print(type(b))  # <class 'tuple'>
```

### どんなアプリに使われているか

- **地図アプリ**：緯度・経度の座標 `tokyo = (35.6895, 139.6917)`
- **画像・動画アプリ**：解像度 `size = (1920, 1080)`
- **カレンダーアプリ**：日付 `birthday = (2000, 3, 15)`

### リストとタプルの使い分け

| | 使う場面 |
|---|---|
| リスト | 追加・削除・変更がある（プレイリスト、商品一覧など） |
| タプル | 変わらない固定データ（座標、サイズ、日付など） |

「このデータは途中で変わることがあるか？」と考えて判断しましょう。

---

## 辞書（dict）

キーと値のペアで管理するコンテナです。名前をつけてデータを管理したいときに使います。Python 3.7以降は挿入順を保持します。

- **キー**：引き出しのラベル
- **バリュー**：引き出しの中身
- **マッピング**：キーとバリューを対応させること

### 基本的な使い方

```python
person = {"name": "Alice", "age": 30, "city": "Tokyo"}

# 値を取り出す
print(person["name"])                    # "Alice"
print(person.get("email", "なし"))       # キーがなくてもエラーにならない

# 追加・更新
person["email"] = "alice@test.com"  # 新しいキーを追加
person["age"] = 31                  # 既存の値を更新

# 削除
del person["city"]                  # 指定したキーを削除
age = person.pop("age")             # 削除しながら値も取り出す
person.clear()                      # 全部削除して空にする

# ループ
for key, value in person.items():
    print(f"{key}: {value}")
```

### 辞書の中にリストを入れる

```python
user = {
    "name": "Alice",
    "hobbies": ["reading", "cooking", "hiking"]
}

print(user["hobbies"][0])  # "reading"
```

### 辞書の中に辞書を入れる

```python
users = {
    "alice": {"age": 30, "city": "Tokyo"},
    "bob":   {"age": 25, "city": "Osaka"}
}

print(users["alice"]["city"])  # "Tokyo"
```

### 注意点

**キーは重複できません。**同じキーで登録すると上書きされます。

```python
d = {"name": "Alice", "name": "Bob"}
print(d)  # {"name": "Bob"}  Aliceが消える！
```

**リストはキーにできません。**タプルはキーにできます。

```python
d = {[1, 2]: "test"}   # エラー！
d = {(1, 2): "test"}   # OK
```

**存在しないキーにアクセスするとエラーになります。** `get()` を使うと安全です。

```python
print(person["email"])              # キーがないとエラー
print(person.get("email", "なし"))  # エラーにならず"なし"を返す
```

### どんなアプリに使われているか

- **SNS・ECサイト**：ユーザー情報の管理
- **アプリ設定**：テーマ、言語、フォントサイズなどの設定情報
- **WebアプリのAPI**：サーバーからのデータ受け取り（JSON形式）

---

## セット（set）

重複を許さないコンテナです。順番はありません。

### 基本的な使い方

```python
fruits = {"apple", "banana", "apple", "cherry"}
print(fruits)  # {"apple", "banana", "cherry"}  重複が消える！

# 追加
fruits.add("grape")

# 削除
fruits.remove("banana")

# 含まれているか確認
print("apple" in fruits)  # True
```

### セット同士の演算

```python
a = {"apple", "banana", "cherry"}
b = {"banana", "cherry", "grape"}

print(a | b)  # 和集合：どちらかに入っているもの全部
# {"apple", "banana", "cherry", "grape"}

print(a & b)  # 積集合：両方に入っているもの
# {"banana", "cherry"}

print(a - b)  # 差集合：aにあってbにないもの
# {"apple"}
```

### 注意点

順番が保証されないのでインデックスで取り出せません。

```python
fruits[0]  # エラー！
```

空のセットを作るときは `set()` を使います。

```python
a = {}      # これは空の辞書になる！
a = set()   # 空のセットはこう書く
```

### どんなアプリに使われているか

- **SNSのフォロー機能**：共通のフォロワーを調べる（積集合）
- **ECサイトのタグ管理**：同じタグが重複してつかないように管理
- **ログイン履歴**：アクセスしたユーザーを重複なく記録

---

## その他の豆知識

### クォーテーションについて

ダブルクォーテーション `"` とシングルクォーテーション `'` はどちらを使っても同じです。ただし文字列の中にクォーテーションを含めたいときは使い分けが便利です。

```python
text1 = "it's a pen"    # 中にシングルを入れたいときはダブルで囲う
text2 = 'He said "hello"'  # 中にダブルを入れたいときはシングルで囲う
```

### コンテナの中にコンテナを格納する

リスト・タプル・辞書・セットはお互いの中に入れ子にして格納できます。外側から順番に `[]` でたどっていくと読みやすいです。

```python
# 辞書の中にリスト
user = {"name": "Alice", "hobbies": ["reading", "cooking"]}
print(user["hobbies"][0])  # "reading"

# リストの中にリスト
matrix = [[1, 2, 3], [4, 5, 6]]
print(matrix[0][1])  # 2
```
