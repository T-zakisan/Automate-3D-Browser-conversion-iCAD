#######################################################################################################
# 
# 内容：iCADファイル（*.icd）を3D Browserファイル（*.html）に変換するシステムの一部
# 
# 方法：
#  1.PowerShellをWindows起動時に実行する　※PowerShell > 設定 > スタートアップ時に起動 > on
#  2. $folderPath を使用条件に従って変更
#  3.このファイルを以下に保存
#      C:\Users\_YourUsername_\Documents\WindowsPowerShell
#     ※なければフォルダ作成のこと
#     ※_YourUsername_ は、各ユーザー名のため、使用環境によって異なる
#
#######################################################################################################
$PathWatch = "C:\path\to\watched\folder"  			# 監視するフォルダのパス
$PathExcel = "C:\path\to\your_vba_macro.xlsm"		# iCADをコントロールするExcel(VBA)
#######################################################################################################



# フォルダ監視設定
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $PathWatch
$watcher.Filter = "*.icd"
$watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName
$watcher.EnableRaisingEvents = $true

# イベントが発生したときのアクション
$action = {
	$filePath = $Event.SourceEventArgs.FullPath
	$destinationPath = Join-Path $PathWatch "convert.icd"  # ICADが認識できるファイル名にリネーム
	$newFileName = [System.IO.Path]::ChangeExtension($originalFile.FullName, ".html")

	# 既存のファイルがあれば削除してリネーム
	if (Test-Path $destinationPath) { Remove-Item $destinationPath }
	Rename-Item -Path $filePath -NewName "convert.icd"

	# ICADの起動とVBAマクロ実行を開始するプロセス（VBAからICADをコントロール）
	$excel = New-Object -ComObject Excel.Application	# Excelアプリケーションを起動
	$excel.Visible = $false	# Excelを非表示にする場合
	$workbook = $excel.Workbooks.Open( $PathExcel )	# 指定したExcelファイルを開く
	$excel.Application.Run( "Conv_iCAD_to_3D",  $destinationPath )  # 実行したいマクロ名を指定
	$workbook.Close( $false )  #	# Excelを閉じる（必要に応じて）	変更を保存しない場合
	$excel.Quit()  # Excelを終了
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)	# COMオブジェクトの解放

	#出力ファイル名を変換
	Rename-Item -Path $destinationPath -NewName $newFileName

}

# ファイル作成イベントにアクションをバインド
Register-ObjectEvent $watcher Created -Action $action


# スクリプトを終了させないために待機
while ($true) {
	Start-Sleep -Seconds 10
}
