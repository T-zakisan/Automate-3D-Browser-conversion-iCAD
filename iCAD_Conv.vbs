Sub Conv_iCAD_to_3D( icadFile As String )

	' ICAD���N��
	On Error Resume Next
	Dim icadApp As Object: Set icadApp = GetObject(, "ICAD.Application")
	If icadApp Is Nothing Then
			Set icadApp = CreateObject("ICAD.Application")
	End If
	On Error GoTo 0

	' �t�@�C�����J��
	icadApp.Documents.Open icadFile

	' �ϊ��R�}���h�̎��s (�R�}���h�o�[�ŕϊ��E�B���h�E���J��)
	icadApp.ExecuteCommand ("3D Browser Export Command")

	' SendKeys��[OK]������ (�E�B���h�E���A�N�e�B�u�ɂȂ�܂őҋ@���鎞�Ԃ�ݒ�)
	Application.Wait Now + TimeValue("00:00:02")
	Application.SendKeys "~" ' Enter


  ' ICAD�̏I�������i�K�v�ɉ����āj
	Set icadApp = Nothing
End Sub
