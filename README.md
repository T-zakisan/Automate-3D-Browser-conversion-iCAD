# Automate-3D-Browser-conversion-iCAD
3D Browser（iCAD）への自動変換

## 準備
1. 上記２つのファイルをDLし、共に
```
C:\Users\_YourUsername_\Documents\WindowsPowerShell
```
に保存する.

2. Windodws起動時にPowerShellを起動するようにする
```
 PowerShell > 設定 > スタートアップ時に起動 > on
```

3. iCAD_Conv.ps1　を環境に依存する部位を変更する
```
$PathWatch = "C:\path\to\watched\folder"  			# 監視するフォルダのパス
$PathExcel = "C:\path\to\your_vba_macro.xlsm"		# iCADをコントロールするExcel(VBA)
```

4. iCAD_conv.vbs　の内、`3D Browser`に変換するコマンドを環境に合わせて変更する
```
icadApp.ExecuteCommand ("3D Browser Export Command")
```

## このままでは、[ok] を押せないであろう…   
Excel VBA側で、`SendKeys`により、カーソルを`[ok]`ボタンに移動/実行する必要がありそうだ.
```
SendKeys "{TAB 2}"
SendKeys "{Enter}"
```

[Senkeys](https://learn.microsoft.com/ja-jp/office/vba/api/excel.application.sendkeys)

