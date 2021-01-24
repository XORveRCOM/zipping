Option Explicit

	' 正規表現のブラックリスト
	Class BlackList
		Dim dict
		Function GetDic
			If IsEmpty(dict) Then
				Set dict = CreateObject("Scripting.Dictionary")
			End If
			Set GetDic = dict
		End Function
		Sub AddPattern(pat)
			If GetDic.Exists(pat) Then Exit Sub
			Dim regex
			Set regex = New RegExp
			regex.Pattern = pat
			regex.IgnoreCase = True
			GetDic.Add pat, regex
		End Sub
		' 登録
		Sub Init(arr)
			Dim item
			For Each item in arr
				AddPattern item
			Next
		End Sub
		' 正規表現リストと一致するかチェック
		Function IsMatch(check)
			Dim regex
			For Each regex In GetDic.Items
				If regex.Test(check) Then
					IsMatch = True
					Exit Function
				End If
			Next
			IsMatch = False
		End Function
	End Class
