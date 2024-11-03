Sub Conv_iCAD_to_3D( icadFile As String )

	' ICADを起動
	On Error Resume Next
	Dim icadApp As Object: Set icadApp = GetObject(, "ICAD.Application")
	If icadApp Is Nothing Then
			Set icadApp = CreateObject("ICAD.Application")
	End If
	On Error GoTo 0

	' ファイルを開く
	icadApp.Documents.Open icadFile

	' 変換コマンドの実行 (コマンドバーで変換ウィンドウを開く)
	icadApp.ExecuteCommand ("3D Browser Export Command")

	' SendKeysで[OK]を押す (ウィンドウがアクティブになるまで待機する時間を設定)
	' ここは、VBAに移動させるべき！
	Application.Wait Now + TimeValue("00:00:02")
	Application.SendKeys "{TAB 2}" ' TAB x 2
	Application.SendKeys "~" ' Enter


  ' ICADの終了処理（必要に応じて）
	Set icadApp = Nothing
End Sub
