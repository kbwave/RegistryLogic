Imports Microsoft.Win32.Registry

''' <summary>
''' レジストリ読込み・書き込みを行うクラス
''' </summary>
''' <remarks></remarks>
Public Module commonRegistry

#Region "Private Variable"
	''' <summary>
	''' 事前準備を行ったか
	''' </summary>
	''' <remarks></remarks>
	Private _isInitialized As Boolean = False

	''' <summary>
	''' 使用するアプリ名を取得
	''' </summary>
	''' <remarks></remarks>
	Private _applicationName As String = String.Empty

	''' <summary>
	''' 書き込みを行うレジストリキー
	''' </summary>
	''' <remarks></remarks>
	Private _regKeyWrite As Microsoft.Win32.RegistryKey = Nothing

	''' <summary>
	''' 読み込みを行うレジストリキー
	''' </summary>
	''' <remarks></remarks>
	Private _regKeyRead As Microsoft.Win32.RegistryKey = Nothing

#End Region

#Region "Primary Function"

	''' <summary>
	''' Int型のデータを登録
	''' </summary>
	''' <param name="data"></param>
	''' <remarks></remarks>
	Public Sub RegistData(ByVal name As String, ByVal data As Integer)
		'SetValue(GetKey, "string", data)
		_regKeyWrite.SetValue(name, data, Microsoft.Win32.RegistryValueKind.DWord)
	End Sub

	''' <summary>
	''' Long型のデータを登録
	''' </summary>
	''' <param name="data"></param>
	''' <remarks></remarks>
	Public Sub RegistData(ByVal name As String, ByVal data As Long)
		_regKeyWrite.SetValue(name, data, Microsoft.Win32.RegistryValueKind.QWord)
	End Sub

	''' <summary>
	''' Single型を登録
	''' </summary>
	''' <param name="data"></param>
	''' <remarks></remarks>
	Public Sub RegistData(ByVal name As String, ByVal data As Single)
		'_regKeyWrite.SetValue(name, data, Microsoft.Win32.RegistryValueKind.DWord)
		_regKeyWrite.SetValue(name, data.ToString, Microsoft.Win32.RegistryValueKind.String)
	End Sub

	''' <summary>
	''' Double型を登録
	''' </summary>
	''' <param name="data"></param>
	''' <remarks></remarks>
	Public Sub RegistData(ByVal name As String, ByVal data As Double)
		'_regKeyWrite.SetValue(name, data, Microsoft.Win32.RegistryValueKind.QWord)
		_regKeyWrite.SetValue(name, data.ToString, Microsoft.Win32.RegistryValueKind.String)
	End Sub

	''' <summary>
	''' String型のデータを登録
	''' </summary>
	''' <param name="data"></param>
	''' <remarks></remarks>
	Public Sub RegistData(ByVal name As String, ByVal data As String, ByVal special As Boolean)
		If special Then
			_regKeyWrite.SetValue(name, data, Microsoft.Win32.RegistryValueKind.String)
		Else
			_regKeyWrite.SetValue(name, data, Microsoft.Win32.RegistryValueKind.ExpandString)
		End If
	End Sub

	''' <summary>
	''' 文字列配列を登録
	''' </summary>
	''' <param name="data"></param>
	''' <remarks></remarks>
	Public Sub RegistData(ByVal name As String, ByVal data As String())
		_regKeyWrite.SetValue(name, data, Microsoft.Win32.RegistryValueKind.MultiString)
	End Sub

	''' <summary>
	''' Int型で取得
	''' </summary>
	''' <param name="name"></param>
	''' <param name="noneValue"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetData(ByVal name As String, ByVal noneValue As Integer) As Integer
		Try
			If _regKeyRead.GetValueKind(name) = Microsoft.Win32.RegistryValueKind.DWord Then
				Return DirectCast(_regKeyRead.GetValue(name), Integer)
			Else
				Return noneValue
			End If

		Catch ex As IO.IOException
			Return noneValue

		Catch ex As Exception
			Return noneValue

		End Try
	End Function

	''' <summary>
	''' Long型で取得
	''' </summary>
	''' <param name="name"></param>
	''' <param name="noneValue"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetData(ByVal name As String, ByVal noneValue As Long) As Long
		Try
			If _regKeyRead.GetValueKind(name) = Microsoft.Win32.RegistryValueKind.QWord Then
				Return DirectCast(_regKeyRead.GetValue(name), Long)
			Else
				Return noneValue
			End If

		Catch ex As Exception
			Return noneValue

		End Try
	End Function

	''' <summary>
	''' Single型で取得
	''' </summary>
	''' <param name="name"></param>
	''' <param name="noneValue"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetData(ByVal name As String, ByVal noneValue As Single) As Single
		Try
			If _regKeyRead.GetValueKind(name) = Microsoft.Win32.RegistryValueKind.String Then
				'Return DirectCast(_regKeyRead.GetValue(name), Single)
				Return Single.Parse(_regKeyRead.GetValue(name).ToString)
			Else
				Return noneValue
			End If

		Catch ex As Exception
			Return noneValue

		End Try
	End Function

	''' <summary>
	''' Double型で取得
	''' </summary>
	''' <param name="name"></param>
	''' <param name="noneValue"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetData(ByVal name As String, ByVal noneValue As Double) As Double
		Try
			If _regKeyRead.GetValueKind(name) = Microsoft.Win32.RegistryValueKind.String Then
				'Return DirectCast(_regKeyRead.GetValue(name), Double)
				Return Double.Parse(_regKeyRead.GetValue(name).ToString)
			Else
				Return noneValue
			End If

		Catch ex As Exception
			Return noneValue

		End Try
	End Function

	''' <summary>
	''' String型で取得
	''' </summary>
	''' <param name="name"></param>
	''' <param name="noneValue"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetData(ByVal name As String, ByVal noneValue As String, ByVal special As Boolean) As String
		Try
			If _regKeyRead.GetValueKind(name) = Microsoft.Win32.RegistryValueKind.String AndAlso Not special Then
				'Return DirectCast(_regKeyRead.GetValue(name), String)
				Return _regKeyRead.GetValue(name).ToString

			ElseIf _regKeyRead.GetValueKind(name) = Microsoft.Win32.RegistryValueKind.ExpandString AndAlso special Then
				'Return DirectCast(_regKeyRead.GetValue(name), String)
				Return _regKeyRead.GetValue(name).ToString

			Else
				Return noneValue
			End If

		Catch ex As Exception
			Return noneValue

		End Try
	End Function

	''' <summary>
	''' String配列で取得
	''' </summary>
	''' <param name="name"></param>
	''' <param name="noneData"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function GetData(ByVal name As String, ByVal noneData As String()) As String()
		Try
			If _regKeyRead.GetValueKind(name) = Microsoft.Win32.RegistryValueKind.MultiString Then
				Return DirectCast(_regKeyRead.GetValue(name), String())
			Else
				Return noneData
			End If

		Catch ex As Exception
			Return noneData

		End Try
	End Function

#End Region

#Region "Another Function"
	''' <summary>
	''' このモジュールのスタートアップを掛ける
	''' </summary>
	''' <remarks></remarks>
	Public Sub StartUp()

		_isInitialized = True
		_applicationName = My.Application.Info.Title

		_regKeyWrite = CurrentUser.CreateSubKey(GetSubKeyName())
		_regKeyRead = CurrentUser.OpenSubKey(GetSubKeyName())
	End Sub

	''' <summary>
	''' スタートアップを行ったか
	''' </summary>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function IsInitialized() As Boolean
		Return IsInitialized
	End Function

	''' <summary>
	''' レジストリのサブキーを取得
	''' </summary>
	''' <returns></returns>
	''' <remarks></remarks>
	Private Function GetSubKeyName() As String
		'//kbwave\のサブ階層に設定値を登録
		Return "kbwave\" & _applicationName
	End Function

#End Region

End Module
