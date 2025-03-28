
# VBAでログ出力するクラスモジュール "VBALog"

簡単に処理ログをファイルに出力することができます。  
また、ファイルの出力場所や日付フォーマット・出力ログレベル等を `config.xml` で設定可能です。


## 使用方法

### **準備するもの**
1. **`config.xml` ファイル**（ログ設定用のファイル）  
   実行ファイルと同じディレクトリ階層に配置してください。


### **config.xml の内容**
```xml
<?xml version="1.0" encoding="UTF-8" ?>
<config>
    <logfile>C:\home\log\log.txt</logfile>
    <dateformat>yyyy/mm/dd hh:mm:ss</dateformat>
    <loglevel>DEBUG</loglevel>
    <debuglog>False</debuglog>
</config>
```


### **設定内容**
| 設定項目    | 説明 | 初期値 |
|------------|-----------------------------------|------------------------|
| `logfile`  | ログ出力パス | 実行ファイルと同じディレクトリ |
| `dateformat` | 出力日付フォーマット | `yyyy/mm/dd HH:MM:ss` |
| `loglevel` | ログ出力レベル (`DEBUG`, `INFO`, `ALEART`, `ERROR`) | `DEBUG` |
| `debuglog` | `DEBUG` ログの出力設定 (`True`: 出力, `False`: 出力しない) | `False` |


## **出力方法**
```vba
Call logger.DebugMsg("ログメッセージ", "メソッド名")
Call logger.InfoMsg("ログメッセージ", "メソッド名")
Call logger.AleartMsg("ログメッセージ", "メソッド名")
Call logger.ErrorMsg("ログメッセージ", "メソッド名")
```

### **出力例**
```
2024/06/18 11:30:00 [INFO] [メソッド名] ログメッセージ
```

---

## **エラー処理**
Excel の `Err` オブジェクトを使用すると、エラーの詳細メッセージを取得できます。

**例**
```vba
MsgBox Err.Number & " " & Err.Description
```

### **エラー時の挙動**
- 書き込みができない場合、ログは記録されません。

---
```